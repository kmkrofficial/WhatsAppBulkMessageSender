const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const puppeteer = require('puppeteer');
const xlsx =require('xlsx');
const fs = require('fs');

if (require('electron-squirrel-startup')) {
  app.quit();
}

const DELAY_AFTER_LOGIN_CHECK = 2500;
const DELAY_BETWEEN_MESSAGES_MIN = 2000;
const DELAY_BETWEEN_MESSAGES_MAX = 3000;
const DELAY_PAGE_LOAD = 3500;
const DELAY_AFTER_SEND = 1500;
const DELAY_AFTER_TYPING_COMPLETES = 300;

const LAUNCH_OPTIONS = {
    headless: false,
    userDataDir: path.join(app.getPath('userData'), 'whatsapp_session_electron'),
    executablePath: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu']
};

const QR_CODE_SELECTOR = 'canvas[aria-label="Scan this QR code to link a device!"], div[data-testid="qrcode"]';
const SEARCH_INPUT_SELECTOR_AFTER_LOGIN = 'div[aria-label="Chat list"], div[data-testid="chat-list"]';
const MESSAGE_INPUT_SELECTOR = 'footer div[contenteditable="true"][data-tab="10"], footer div[contenteditable="true"][data-tab="9"]';
const SEND_BUTTON_SELECTOR = 'button[aria-label="Send"], span[data-icon="send"]';
const OK_BUTTON_SELECTOR_INVALID_NUMBER = 'div[data-testid="popup-controls-ok"]';

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 850,
        height: 780,
        icon: path.join(__dirname, 'assets/logo.png'),
        webPreferences: { nodeIntegration: true, contextIsolation: false, devTools: !app.isPackaged }
    });
    mainWindow.loadFile('index.html');
    mainWindow.on('closed', () => { mainWindow = null; });
}

app.whenReady().then(() => {
    createWindow();
    app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });

ipcMain.on('open-file-dialog', (event) => {
    dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }]
    }).then(result => {
        if (!result.canceled && result.filePaths.length > 0) event.sender.send('selected-file', result.filePaths[0]);
    }).catch(err => sendLog(event.sender, `File dialog error: ${err.message}`, 'error'));
});

ipcMain.on('get-excel-headers', async (event, filePath) => {
    try {
        if (!filePath) { event.sender.send('excel-headers-list', { error: "File path is missing." }); return; }
        const workbook = xlsx.readFile(filePath, { sheetRows: 1 });
        const sheetName = workbook.SheetNames[0];
        if (!sheetName) { event.sender.send('excel-headers-list', { error: "No sheets found." }); return; }
        const sheet = workbook.Sheets[sheetName];
        const jsonData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        if (jsonData.length > 0 && Array.isArray(jsonData[0])) {
            event.sender.send('excel-headers-list', jsonData[0].map(String));
        } else {
            event.sender.send('excel-headers-list', []);
        }
    } catch (error) { event.sender.send('excel-headers-list', { error: error.message }); }
});

ipcMain.on('open-screenshot-dir-dialog', (event) => {
    dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory', 'createDirectory']
    }).then(result => {
        if (!result.canceled && result.filePaths.length > 0) event.sender.send('selected-screenshot-dir', result.filePaths[0]);
    }).catch(err => sendLog(event.sender, `Screenshot directory dialog error: ${err.message}`, 'error'));
});

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)); }
function getRandomDelay(min, max) { return Math.floor(Math.random() * (max - min + 1) + min); }

function sanitizePhoneNumber(number, defaultCountryCode = '') {
    let phone = String(number).replace(/\D/g, '');
    if (defaultCountryCode && !phone.startsWith(defaultCountryCode) && phone.length === 10) phone = defaultCountryCode + phone;
    return phone;
}

async function readContacts(filePath, webContents) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        return xlsx.utils.sheet_to_json(sheet);
    } catch (error) {
        sendLog(webContents, `Error reading Excel file: ${error.message}`, 'error');
        return [];
    }
}

function sendLog(webContents, message, type = 'info') {
    if (webContents && !webContents.isDestroyed()) webContents.send('log-message', { message, type });
    console.log(`[${type.toUpperCase()}] ${message}`);
}

ipcMain.on('start-sending-messages', async (event, config) => {
    const { excelFilePath, messageTemplate, countryCode, phoneNumberColumn, nameColumn, screenshotDir } = config;
    const webContents = event.sender;

    let actualScreenshotDir = screenshotDir && screenshotDir.trim() !== '' ? screenshotDir : path.join(process.cwd(), 'failure_screenshots');
    try {
        if (!fs.existsSync(actualScreenshotDir)) {
            fs.mkdirSync(actualScreenshotDir, { recursive: true });
            sendLog(webContents, `Created screenshot directory: ${actualScreenshotDir}`, 'info');
        } else {
            sendLog(webContents, `Using screenshot directory: ${actualScreenshotDir}`, 'info');
        }
    } catch (dirError) { sendLog(webContents, `Error preparing screenshot directory: ${dirError.message}.`, 'error'); }

    if (!phoneNumberColumn || !nameColumn) {
        sendLog(webContents, 'Error: Column names not specified.', 'error');
        if (webContents && !webContents.isDestroyed()) {
             webContents.send('update-progress', { percentage: 100, etaFormatted: 'Error!', totalContacts: 0, processedCount: 0, successful: 0, failed: 0, invalid: 0 });
             webContents.send('process-finished', { totalContacts: 0, successful_messages_count: 0, failed_messages_count: 0, invalid_phone_numbers_count: 0, failed_ids: [], invalid_phone_numbers_ids: [] });
        }
        return;
    }

    sendLog(webContents, `Using Phone: "${phoneNumberColumn}", Name: "${nameColumn}"`, 'info');
    const contacts = await readContacts(excelFilePath, webContents);
    const totalContacts = contacts.length;

    if (totalContacts === 0) {
        sendLog(webContents, 'No contacts or error reading Excel.', 'error');
        if (webContents && !webContents.isDestroyed()) {
            webContents.send('update-progress', { percentage: 100, etaFormatted: 'N/A', totalContacts: 0, processedCount: 0, successful: 0, failed: 0, invalid: 0 });
            webContents.send('process-finished', { totalContacts: 0, successful_messages_count: 0, failed_messages_count: 0, invalid_phone_numbers_count: 0, failed_ids: [], invalid_phone_numbers_ids: [] });
        }
        return;
    }
    sendLog(webContents, `Found ${totalContacts} contacts. Processing...`, 'info');

    let browser, page;
    const startTime = Date.now(); 
    let s_count = 0, f_count = 0, inv_count = 0, processed_count = 0;
    let f_ids = [], inv_ids = [];

    if (webContents && !webContents.isDestroyed()) {
        webContents.send('update-progress', {
            percentage: 0, etaFormatted: 'Calculating...', totalContacts: totalContacts, 
            processedCount: 0, successful: 0, failed: 0, invalid: 0
        });
    }

    try {
        browser = await puppeteer.launch(LAUNCH_OPTIONS);
        page = await browser.newPage();
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.127 Safari/537.36'); // Updated User Agent
        await page.setViewport({ width: 1200, height: 800 });

        sendLog(webContents, 'Navigating to WhatsApp Web...', 'info');
        await page.goto('https://web.whatsapp.com/', { waitUntil: 'networkidle2', timeout: 60000 });

        sendLog(webContents, 'Waiting for login (max 60s)...', 'info');
        try {
            await page.waitForSelector(`${QR_CODE_SELECTOR}, ${SEARCH_INPUT_SELECTOR_AFTER_LOGIN}`, { timeout: 60000 });
            if (await page.$(QR_CODE_SELECTOR)) {
                sendLog(webContents, 'QR Code detected. Please scan.', 'info');
                await page.waitForSelector(SEARCH_INPUT_SELECTOR_AFTER_LOGIN, { timeout: 90000 });
                sendLog(webContents, 'Login successful!', 'info');
            } else {
                sendLog(webContents, 'Already logged in/session restored.', 'info');
            }
        } catch (err) {
            sendLog(webContents, `Login failed: ${err.message}`, 'error');
            if (browser) await browser.close();
            if (webContents && !webContents.isDestroyed()) {
                webContents.send('update-progress', { percentage: 0, etaFormatted: 'Login Error!', totalContacts: totalContacts, processedCount: 0, successful: 0, failed: 0, invalid: 0 });
                webContents.send('error-occurred', `Login failed: ${err.message}`);
            }
            return;
        }
        await sleep(DELAY_AFTER_LOGIN_CHECK);

        for (let i = 0; i < contacts.length; i++) {
            const contact = contacts[i];
            const name = contact[nameColumn] || 'Friend';
            let rawPhone = contact[phoneNumberColumn];
            
            let messageSentThisIteration = false;
            let invalidNumberThisIteration = false;
            let failedThisIteration = false;

            if (!rawPhone) {
                sendLog(webContents, `Skipping ${name} (Row ${i+2}): missing phone in "${phoneNumberColumn}".`, 'warn');
                invalidNumberThisIteration = true;
            }
            
            const sanNum = invalidNumberThisIteration ? "N/A" : sanitizePhoneNumber(String(rawPhone), countryCode);
            const finalMsg = messageTemplate.replace(/{{Name}}/g, name);

            if (!invalidNumberThisIteration) {
                sendLog(webContents, `[${i + 1}/${totalContacts}] To: ${name} (${sanNum})...`, 'info');
                const chatUrl = `https://web.whatsapp.com/send?phone=${sanNum}&text=&app_absent=0`;
                try {
                    sendLog(webContents, `Navigating to chat: ${sanNum}...`, 'info');
                    await page.goto(chatUrl, { waitUntil: 'domcontentloaded', timeout: 20000 });
                    await sleep(DELAY_PAGE_LOAD);

                    try {
                        await page.waitForFunction(
                            (okBtnSel) => { // Removed unused popupSel parameter
                                const el = document.querySelector('div[role="dialog"]');
                                if (el && el.innerText.toLowerCase().includes("phone number shared via url is invalid")) {
                                    const okBtn = el.querySelector(okBtnSel);
                                    if (okBtn) okBtn.click();
                                    return true;
                                } return false;
                            }, { timeout: 3000 }, OK_BUTTON_SELECTOR_INVALID_NUMBER // Pass only okBtnSel
                        );
                        sendLog(webContents, `Invalid number popup for ${sanNum} (${name}). Skipping.`, 'warn');
                        invalidNumberThisIteration = true;
                        await sleep(500);
                    } catch (e) {
                        sendLog(webContents, `No invalid number popup for ${sanNum}. Proceeding...`, 'info');
                    }

                    if (!invalidNumberThisIteration) {
                        await page.waitForSelector(MESSAGE_INPUT_SELECTOR, { visible: true, timeout: 15000 });
                        await page.focus(MESSAGE_INPUT_SELECTOR);
                        await page.evaluate((selector) => {
                            const el = document.querySelector(selector); if (el) el.innerHTML = '';
                        }, MESSAGE_INPUT_SELECTOR);
                        await sleep(50);

                        const messageLines = finalMsg.split('\n');
                        for (let j = 0; j < messageLines.length; j++) {
                            await page.type(MESSAGE_INPUT_SELECTOR, messageLines[j], { delay: 0 });
                            if (j < messageLines.length - 1) {
                                await page.keyboard.down('Shift');
                                await page.keyboard.press('Enter');
                                await page.keyboard.up('Shift');
                            }
                        }
                        await sleep(DELAY_AFTER_TYPING_COMPLETES);
                        await page.waitForSelector(SEND_BUTTON_SELECTOR, { visible: true, timeout: 5000 });
                        await page.click(SEND_BUTTON_SELECTOR);
                        sendLog(webContents, `Message sent to ${name} (${sanNum})!`, 'info');
                        messageSentThisIteration = true;
                        await sleep(DELAY_AFTER_SEND);
                    }
                } catch (err) {
                    sendLog(webContents, `Failed for ${name} (${sanNum}): ${err.message}`, 'error');
                    failedThisIteration = true;
                    if (err.message.includes('Target closed') || err.message.includes('Page crashed')) {
                        sendLog(webContents, "Browser/page crashed. Aborting.", 'error'); throw err;
                    }
                    try {
                        if (page && !page.isClosed()) {
                            const ssPath = path.join(actualScreenshotDir, `error_${sanNum}_${Date.now()}.png`);
                            await page.screenshot({ path: ssPath });
                            sendLog(webContents, `Screenshot: ${ssPath}`, 'warn');
                        }
                    } catch (scErr) { sendLog(webContents, `Screenshot failed: ${scErr.message}`, 'warn'); }
                }
            }

            processed_count++;
            if (invalidNumberThisIteration) {
                inv_count++;
                if (nameColumn && contact[nameColumn]) inv_ids.push(contact[nameColumn]); else inv_ids.push(sanNum === "N/A" ? "Missing Phone" : sanNum);
            } else if (failedThisIteration) {
                f_count++;
                if (nameColumn && contact[nameColumn]) f_ids.push(contact[nameColumn]); else f_ids.push(sanNum);
            } else if (messageSentThisIteration) {
                s_count++;
            }

            const percentage = (processed_count / totalContacts) * 100;
            const elapsedTime = (Date.now() - startTime) / 1000;
            const timePerMessage = processed_count > 0 ? elapsedTime / processed_count : 0;
            const remainingMessages = totalContacts - processed_count;
            const etaSeconds = remainingMessages * timePerMessage;
            let etaFormatted = 'Calculating...';

            if (processed_count > 0 && remainingMessages >= 0) {
                if (etaSeconds === Infinity || isNaN(etaSeconds) || etaSeconds < 0) { // Added check for etaSeconds < 0
                    etaFormatted = remainingMessages === 0 ? 'Done!' : 'Calculating...';
                } else if (remainingMessages === 0) {
                    etaFormatted = 'Finishing up...';
                } else {
                    const hours = Math.floor(etaSeconds / 3600);
                    const minutes = Math.floor((etaSeconds % 3600) / 60);
                    const seconds = Math.floor(etaSeconds % 60);
                    etaFormatted = '';
                    if (hours > 0) etaFormatted += `${hours}h `;
                    if (minutes > 0 || hours > 0) etaFormatted += `${minutes}m `;
                    etaFormatted += `${seconds}s`;
                    if (etaFormatted.trim() === '0s' && remainingMessages > 0) etaFormatted = "<1s ea, soon";
                    else if (etaFormatted.trim() === '0s' && remainingMessages === 0) etaFormatted = "Done!";
                }
            }
            
            if (webContents && !webContents.isDestroyed()) {
                webContents.send('update-progress', {
                    percentage: percentage, etaFormatted: etaFormatted, totalContacts: totalContacts,
                    processedCount: processed_count, successful: s_count, failed: f_count, invalid: inv_count
                });
            }
            
            if (i < totalContacts - 1) {
                const rndDelay = getRandomDelay(DELAY_BETWEEN_MESSAGES_MIN, DELAY_BETWEEN_MESSAGES_MAX);
                sendLog(webContents, `Waiting ${rndDelay / 1000}s...`, 'info');
                await sleep(rndDelay);
            }
        }

        const finalStats = {
            totalContacts: totalContacts, successful_messages_count: s_count, failed_messages_count: f_count,
            invalid_phone_numbers_count: inv_count, failed_ids: f_ids, invalid_phone_numbers_ids: inv_ids
        };
        if (webContents && !webContents.isDestroyed()) {
             webContents.send('update-progress', {
                percentage: 100, etaFormatted: 'Completed!', totalContacts: totalContacts,
                processedCount: totalContacts, successful: s_count, failed: f_count, invalid: inv_count
            });
            webContents.send('process-finished', finalStats);
        }

    } catch (error) {
        sendLog(webContents, `CRITICAL ERROR: ${error.message}`, 'error');
        if (webContents && !webContents.isDestroyed()) {
             webContents.send('update-progress', {
                percentage: (processed_count / totalContacts) * 100, etaFormatted: 'Error!',
                totalContacts: totalContacts, processedCount: processed_count, successful: s_count,
                failed: f_count, invalid: inv_count
            });
            webContents.send('error-occurred', error.message);
        }
    } finally {
        if (browser) {
            sendLog(webContents, 'Closing browser...', 'info');
            await browser.close();
        }
    }
});