const puppeteer = require('puppeteer');
const xlsx = require('xlsx');
const fs = require('fs');

// --- Configuration ---
const EXCEL_FILE_PATH = './Book1.xlsx'; // Path to your Excel file
const PHONE_NUMBER_COLUMN = 'Phone Number'; // Column name for phone numbers
const NAME_COLUMN = 'Name';                // Column name for names
const MESSAGE_TEMPLATE = `Hello {{Name}}, This is a sample message!`; // Personalize with {{Name}}
const COUNTRY_CODE = '91'; // Optional: Default country code if not in Excel (e.g., '91' for India, '1' for US). Leave empty if numbers are full E.164.

// Delays (in milliseconds) - VERY IMPORTANT to avoid being banned
const DELAY_AFTER_LOGIN_CHECK = 5000; // Time to wait after checking login status
const DELAY_BETWEEN_MESSAGES_MIN = 5000; // Minimum delay between sending messages
const DELAY_BETWEEN_MESSAGES_MAX = 8000; // Maximum delay between sending messages
const DELAY_PAGE_LOAD = 6000;          // Wait for chat page to load
const DELAY_AFTER_SEND = 3000;            // Short delay after clicking send

// Puppeteer launch options
const LAUNCH_OPTIONS = {
    headless: false, // Set to true for background, false to see the browser
    userDataDir: './whatsapp_session', // Saves session data (cookies, etc.) to avoid frequent QR scans
    args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote',
        // '--single-process', // Could cause issues
        '--disable-gpu'
    ]
};

// Selectors (these might change if WhatsApp Web updates its UI)
const QR_CODE_SELECTOR = 'canvas[aria-label="Scan this QR code to link a device!"]';
const SEARCH_INPUT_SELECTOR_AFTER_LOGIN = 'div[aria-label="Chat list"]'; // Try both
const MESSAGE_INPUT_SELECTOR = 'footer div[role="textbox"][contenteditable="true"]'; // Common selector for the message box
const SEND_BUTTON_SELECTOR = 'button[aria-label="Send"], span[data-icon="send"]'; // Common selector for the send button

// --- Helper Functions ---
function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function getRandomDelay(min, max) {
    return Math.floor(Math.random() * (max - min + 1) + min);
}

function sanitizePhoneNumber(number, defaultCountryCode = '') {
    let phone = String(number).replace(/\D/g, ''); // Remove non-digits
    if (defaultCountryCode && !phone.startsWith(defaultCountryCode) && phone.length < 11) { // Basic check, adjust as needed
        phone = defaultCountryCode + phone;
    }
    return phone;
}

async function readContacts(filePath) {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        return xlsx.utils.sheet_to_json(sheet);
    } catch (error) {
        console.error(`Error reading Excel file: ${error.message}`);
        return [];
    }
}

// --- Main Automation Logic ---
async function sendWhatsAppMessages() {
    const contacts = await readContacts(EXCEL_FILE_PATH);
    if (contacts.length === 0) {
        console.log('No contacts found in the Excel file or error reading it.');
        return;
    }

    console.log(`Found ${contacts.length} contacts. Preparing to send messages...`);

    const browser = await puppeteer.launch(LAUNCH_OPTIONS);
    const page = await browser.newPage();
    await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');
    await page.setViewport({ width: 1200, height: 800 });

    

    try {
        console.log('Navigating to WhatsApp Web...');
        await page.goto('https://web.whatsapp.com/', { waitUntil: 'networkidle2', timeout: 60000 });

        console.log('Please scan the QR code if prompted. Waiting for login...');
        try {
            // Wait for either QR code or the main chat interface (if already logged in)
            await page.waitForSelector(`${QR_CODE_SELECTOR}, ${SEARCH_INPUT_SELECTOR_AFTER_LOGIN}`, { timeout: 90000 }); // 90 seconds to scan

            // Check if QR code is visible, meaning we need to scan
            const isQrVisible = await page.$(QR_CODE_SELECTOR);
            if (isQrVisible) {
                console.log('QR Code detected. Please scan with your phone.');
                // Wait until QR code disappears (i.e., login is successful)
                // Or wait for search input to appear
                await page.waitForSelector(SEARCH_INPUT_SELECTOR_AFTER_LOGIN, { timeout: 0 }); // Wait indefinitely for login
                console.log('Login successful (QR code scanned or session restored)!');
            } else {
                console.log('Already logged in or session restored.');
            }
        } catch (err) {
            console.error('Login timeout or QR code not found. Please ensure WhatsApp Web loads correctly.', err);
            await browser.close();
            return;
        }

        await sleep(DELAY_AFTER_LOGIN_CHECK);
        
        let successful_messages_count = 0, failed_messages_count = 0, invalid_phone_numbers_count = 0;
        let failed_ids = [], invalid_phone_numbers_ids = [];

        for (let i = 0; i < contacts.length; i++) {

            const contact = contacts[i];
            const name = contact[NAME_COLUMN] || 'Friend'; // Fallback name
            let rawPhoneNumber = contact[PHONE_NUMBER_COLUMN];

            if (!rawPhoneNumber) {
                console.warn(`Skipping contact ${name} due to missing phone number.`);
                continue;
            }

            const phoneNumber = sanitizePhoneNumber(rawPhoneNumber, COUNTRY_CODE);
            const message = MESSAGE_TEMPLATE.replace(/{{Name}}/g, name);

            console.log(`\n[${i + 1}/${contacts.length}] Preparing to message ${name} (${phoneNumber})...`);

            const chatUrl = `https://web.whatsapp.com/send?phone=${phoneNumber}&text=${encodeURIComponent('')}&app_absent=0`;

            try {
                console.log(`Navigating to chat with ${phoneNumber}...`);
                await page.goto(chatUrl, { waitUntil: 'domcontentloaded', timeout: 30000 }); // networkidle0 can be slow
                await sleep(DELAY_PAGE_LOAD); // Extra wait for chat to settle

                // Wait for the message input box to be available
                try {
                    await page.waitForSelector(MESSAGE_INPUT_SELECTOR, { visible: true, timeout: 15000 });
                } catch (e) {

                     // Check for "Phone number shared via url is invalid."
                    const invalidNumberError = await page.evaluate(() => {
                        const el = document.querySelector('div[aria-label="Phone number shared via url is invalid."]'); // This selector might need adjustment
                        return el && el.innerText.toLowerCase().includes("phone number shared via url is invalid");
                    });

                    if (invalidNumberError) {
                        invalid_phone_numbers_count++;
                        invalid_phone_numbers_ids.push(contact[NAME_COLUMN])
                        console.warn(`Failed to open chat with ${phoneNumber}: Number might be invalid or not on WhatsApp. Skipping.`);
                        const okButtonSelector = 'div[data-testid="popup-controls-ok"]';
                        if (await page.$(okButtonSelector)) {
                            await page.click(okButtonSelector);
                            await sleep(1000); // give it a moment
                        }
                        continue;
                    } else {
                        failed_messages_count++;
                        failed_ids.push(NAME_COLUMN)
                        console.error(`Message input box not found for ${phoneNumber} after extended wait. It might be a non-WhatsApp number or a page load issue. Skipping.`);
                        // Take a screenshot for debugging
                        await page.screenshot({ path: `error_contact_${phoneNumber}.png` });
                        console.log(`Screenshot saved to error_contact_${phoneNumber}.png`);
                        continue;
                    }
                }


                console.log(`Typing message...`);
                await page.type(MESSAGE_INPUT_SELECTOR, message); // Type slowly

                // Wait for the send button to be available and click it
                await page.waitForSelector(SEND_BUTTON_SELECTOR, { visible: true, timeout: 10000 });
                await page.click(SEND_BUTTON_SELECTOR);
                console.log(`Message sent to ${name} (${phoneNumber})!`);

                await sleep(DELAY_AFTER_SEND);

            } catch (err) {
                console.error(`Failed to send message to ${name} (${phoneNumber}): ${err.message}`);
                if (err.message.includes('Target closed') || err.message.includes('Page crashed')) {
                    console.error("Browser or page crashed. Exiting.");
                    throw err; // Rethrow to stop the process if critical
                }
                await page.screenshot({ path: `error_send_${phoneNumber}.png` });
                console.log(`Screenshot saved to error_send_${phoneNumber}.png`);
            }

            const randomDelay = getRandomDelay(DELAY_BETWEEN_MESSAGES_MIN, DELAY_BETWEEN_MESSAGES_MAX);
            console.log(`Waiting for ${randomDelay / 1000} seconds before next message...`);
            await sleep(randomDelay);
        }

        console.log('\nAll messages processed.');
        console.log(`Stats: \nSuccessful - ${successful_messages_count} \nFailed - ${failed_messages_count} \nInvalid Phone Numbers - ${invalid_phone_numbers_count} \nFailed Phone Numbers - ${failed_ids} \nInvalid Phone Numbers - ${invalid_phone_numbers_ids}`)

    } catch (error) {
        console.error(`An unexpected error occurred: ${error.message}`);
        if (page) {
           await page.screenshot({ path: 'critical_error.png' });
           console.log('Screenshot captured: critical_error.png');
        }
    } finally {
        if (browser) {
            console.log('Closing browser...');
            await browser.close();
        }
    }
}

// Run the automation
sendWhatsAppMessages();