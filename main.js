const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const puppeteer = require('puppeteer');
const xlsx =require('xlsx');
const fs = require('fs');

if (require('electron-squirrel-startup')) return;

// --- Configuration - Snappier Values ---
const DELAY_AFTER_LOGIN_CHECK = 2500; // Reduced
const DELAY_BETWEEN_MESSAGES_MIN = 2000; // Reduced
const DELAY_BETWEEN_MESSAGES_MAX = 4000; // Reduced
const DELAY_PAGE_LOAD = 3500; // Reduced (after navigating to chat URL)
const DELAY_AFTER_SEND = 1500; // Reduced
const DELAY_AFTER_TYPING_COMPLETES = 300; // Reduced (after all typing is done)

const LAUNCH_OPTIONS = {
    headless: false,
    userDataDir: path.join(app.getPath('userData'), 'whatsapp_session_electron'),
    args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage', '--disable-accelerated-2d-canvas', '--no-first-run', '--no-zygote', '--disable-gpu']
};

// Selectors (remain the same)
const QR_CODE_SELECTOR = 'canvas[aria-label="Scan this QR code to link a device!"], div[data-testid="qrcode"]';
const SEARCH_INPUT_SELECTOR_AFTER_LOGIN = 'div[aria-label="Chat list"], div[data-testid="chat-list"]';
const MESSAGE_INPUT_SELECTOR = 'footer div[contenteditable="true"][data-tab="10"], footer div[contenteditable="true"][data-tab="9"]';
const SEND_BUTTON_SELECTOR = 'button[aria-label="Send"], span[data-icon="send"]';
const INVALID_NUMBER_POPUP_TEXT_SELECTOR = 'div[role="button"]';
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
    }).catch(err => {
        console.error("File dialog error:", err);
        sendLog(event.sender, `File dialog error: ${err.message}`, 'error');
    });
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
    } catch (error) {
        console.error(`Error reading Excel headers: ${error.message}`);
        event.sender.send('excel-headers-list', { error: error.message });
    }
});

ipcMain.on('open-screenshot-dir-dialog', (event) => {
    dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory', 'createDirectory']
    }).then(result => {
        if (!result.canceled && result.filePaths.length > 0) event.sender.send('selected-screenshot-dir', result.filePaths[0]);
    }).catch(err => {
        console.error("Screenshot directory dialog error:", err);
        sendLog(event.sender, `Screenshot directory dialog error: ${err.message}`, 'error');
    });
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
    } catch (dirError) {
        sendLog(webContents, `Error preparing screenshot directory: ${dirError.message}.`, 'error');
    }

    if (!phoneNumberColumn || !nameColumn) {
        sendLog(webContents, 'Error: Column names not specified.', 'error');
        if (webContents && !webContents.isDestroyed()) webContents.send('process-finished', { successful_messages_count: 0, failed_messages_count: 0, invalid_phone_numbers_count: 0, failed_ids: [], invalid_phone_numbers_ids: [] });
        return;
    }

    sendLog(webContents, `Using Phone: "${phoneNumberColumn}", Name: "${nameColumn}"`, 'info');
    const contacts = await readContacts(excelFilePath, webContents);
    if (contacts.length === 0) {
        sendLog(webContents, 'No contacts or error reading Excel.', 'error');
        if (webContents && !webContents.isDestroyed()) webContents.send('process-finished', { successful_messages_count: 0, failed_messages_count: 0, invalid_phone_numbers_count: 0, failed_ids: [], invalid_phone_numbers_ids: [] });
        return;
    }
    sendLog(webContents, `Found ${contacts.length} contacts. Processing...`, 'info');

    let browser, page;
    try {
        browser = await puppeteer.launch(LAUNCH_OPTIONS);
        page = await browser.newPage();
        await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
        await page.setViewport({ width: 1200, height: 800 });

        sendLog(webContents, 'Navigating to WhatsApp Web...', 'info');
        await page.goto('https://web.whatsapp.com/', { waitUntil: 'networkidle2', timeout: 60000 }); // Reduced timeout

        sendLog(webContents, 'Waiting for login (max 60s)...', 'info'); // Reduced timeout
        try {
            await page.waitForSelector(`${QR_CODE_SELECTOR}, ${SEARCH_INPUT_SELECTOR_AFTER_LOGIN}`, { timeout: 60000 }); // Reduced
            if (await page.$(QR_CODE_SELECTOR)) {
                sendLog(webContents, 'QR Code detected. Please scan.', 'info');
                await page.waitForSelector(SEARCH_INPUT_SELECTOR_AFTER_LOGIN, { timeout: 90000 }); // Reduced (user scan time)
                sendLog(webContents, 'Login successful!', 'info');
            } else {
                sendLog(webContents, 'Already logged in/session restored.', 'info');
            }
        } catch (err) {
            sendLog(webContents, `Login failed: ${err.message}`, 'error');
            if (browser) await browser.close();
            if (webContents && !webContents.isDestroyed()) webContents.send('error-occurred', `Login failed: ${err.message}`);
            return;
        }
        await sleep(DELAY_AFTER_LOGIN_CHECK); // Uses reduced constant

        let s_count = 0, f_count = 0, inv_count = 0;
        let f_ids = [], inv_ids = [];

        for (let i = 0; i < contacts.length; i++) {
            const contact = contacts[i];
            const name = contact[nameColumn] || 'Friend';
            let rawPhone = contact[phoneNumberColumn];
            if (!rawPhone) {
                sendLog(webContents, `Skipping ${name} (Row ${i+2}): missing phone in "${phoneNumberColumn}".`, 'warn');
                continue;
            }
            const sanNum = sanitizePhoneNumber(String(rawPhone), countryCode);
            const finalMsg = messageTemplate.replace(/{{Name}}/g, name);

            sendLog(webContents, `[${i + 1}/${contacts.length}] To: ${name} (${sanNum})...`, 'info');
            const chatUrl = `https://web.whatsapp.com/send?phone=${sanNum}&text=&app_absent=0`;

            try {
                sendLog(webContents, `Navigating to chat: ${sanNum}...`, 'info');
                await page.goto(chatUrl, { waitUntil: 'domcontentloaded', timeout: 20000 }); // Reduced timeout
                await sleep(DELAY_PAGE_LOAD); // Uses reduced constant

                try {
                    await page.waitForFunction(
                        (popupSel, okBtnSel) => {
                            const el = document.querySelector('div[role="dialog"]');
                            if (el && el.innerText.toLowerCase().includes("phone number shared via url is invalid")) {
                                const okBtn = el.querySelector(okBtnSel);
                                if (okBtn) okBtn.click();
                                return true;
                            } return false;
                        }, { timeout: 3000 }, // Reduced timeout
                        INVALID_NUMBER_POPUP_TEXT_SELECTOR, OK_BUTTON_SELECTOR_INVALID_NUMBER
                    );
                    sendLog(webContents, `Invalid number popup for ${sanNum} (${name}). Skipping.`, 'warn');
                    inv_count++; inv_ids.push(name || sanNum);
                    await sleep(500); continue; // Reduced sleep
                } catch (e) {
                    sendLog(webContents, `No invalid number popup for ${sanNum}. Proceeding...`, 'info');
                }

                await page.waitForSelector(MESSAGE_INPUT_SELECTOR, { visible: true, timeout: 15000 }); // Reduced timeout
                await page.focus(MESSAGE_INPUT_SELECTOR);

                // Clear the input field first
                await page.evaluate((selector) => {
                    const el = document.querySelector(selector);
                    if (el) el.innerHTML = '';
                }, MESSAGE_INPUT_SELECTOR);
                await sleep(50); // Reduced sleep

                // Typing logic with multi-paragraph handling, delay: 0
                sendLog(webContents, `Quickly typing message for ${name}: "${finalMsg.substring(0,40).replace(/\n/g, '\\n')}..."`, 'info');
                const messageLines = finalMsg.split('\n');
                for (let j = 0; j < messageLines.length; j++) {
                    await page.type(MESSAGE_INPUT_SELECTOR, messageLines[j], { delay: 0 }); // delay: 0
                    if (j < messageLines.length - 1) { // If not the last line
                        await page.keyboard.down('Shift');
                        await page.keyboard.press('Enter');
                        await page.keyboard.up('Shift');
                        // await sleep(20); // Optional very small delay if issues with Shift+Enter recognition
                    }
                }
                
                await sleep(DELAY_AFTER_TYPING_COMPLETES); // Uses reduced constant

                await page.waitForSelector(SEND_BUTTON_SELECTOR, { visible: true, timeout: 5000 }); // Reduced timeout
                await page.click(SEND_BUTTON_SELECTOR);
                sendLog(webContents, `Message sent to ${name} (${sanNum})!`, 'info');
                s_count++;
                await sleep(DELAY_AFTER_SEND); // Uses reduced constant

            } catch (err) {
                sendLog(webContents, `Failed for ${name} (${sanNum}): ${err.message}`, 'error');
                f_count++; f_ids.push(name || sanNum);
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
            const rndDelay = getRandomDelay(DELAY_BETWEEN_MESSAGES_MIN, DELAY_BETWEEN_MESSAGES_MAX); // Uses reduced constants
            sendLog(webContents, `Waiting ${rndDelay / 1000}s...`, 'info');
            await sleep(rndDelay);
        }
        const stats = { successful_messages_count: s_count, failed_messages_count: f_count, invalid_phone_numbers_count: inv_count, failed_ids: f_ids, invalid_phone_numbers_ids: inv_ids };
        if (webContents && !webContents.isDestroyed()) webContents.send('process-finished', stats);
    } catch (error) {
        sendLog(webContents, `CRITICAL ERROR: ${error.message}`, 'error');
        if (webContents && !webContents.isDestroyed()) webContents.send('error-occurred', error.message);
        if (browser && page && !page.isClosed()) {
           try {
             const critSsPath = path.join(actualScreenshotDir, `critical_error_${Date.now()}.png`);
             await page.screenshot({ path: critSsPath });
             sendLog(webContents, `Crit Screenshot: ${critSsPath}`, 'warn');
           } catch (scErr) { sendLog(webContents, `Crit Screenshot failed: ${scErr.message}`, 'warn'); }
        }
    } finally {
        if (browser) {
            sendLog(webContents, 'Closing browser...', 'info');
            await browser.close();
        }
    }
});