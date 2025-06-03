const { ipcRenderer } = require('electron');

const selectFileBtn = document.getElementById('selectFileBtn');
const excelFileInput = document.getElementById('excelFileInput'); // Hidden input
const filePathDisplay = document.getElementById('filePath');

// Select elements for column names
const phoneNumberColumnSelect = document.getElementById('phoneNumberColumn');
const nameColumnSelect = document.getElementById('nameColumn');

// New UI elements for screenshot directory
const screenshotDirInput = document.getElementById('screenshotDirInput');
const selectScreenshotDirBtn = document.getElementById('selectScreenshotDirBtn');
const useDefaultScreenshotDirBtn = document.getElementById('useDefaultScreenshotDirBtn');

const messageTemplateInput = document.getElementById('messageTemplate');
const countryCodeInput = document.getElementById('countryCode');
const startBtn = document.getElementById('startBtn');
const logOutput = document.getElementById('logOutput');

let selectedExcelFilePath = '';
let availableHeaders = []; // To store headers from Excel
let selectedScreenshotDir = ''; // To store the user-selected path, empty means default

// --- Helper function to populate select dropdowns ---
function populateColumnDropdowns(headers) {
    availableHeaders = headers || []; // Store for later use

    // Clear existing options
    phoneNumberColumnSelect.innerHTML = '';
    nameColumnSelect.innerHTML = '';

    if (availableHeaders.length === 0) {
        const defaultOption = new Option('-- No headers found or file not selected --', '');
        phoneNumberColumnSelect.add(defaultOption.cloneNode(true));
        nameColumnSelect.add(defaultOption);
        return;
    }

    const placeholderOption = new Option('-- Select a Column --', '');
    phoneNumberColumnSelect.add(placeholderOption.cloneNode(true));
    nameColumnSelect.add(placeholderOption);


    availableHeaders.forEach(header => {
        const option = new Option(header, header);
        phoneNumberColumnSelect.add(option.cloneNode(true));
        nameColumnSelect.add(option);
    });

    // Try to re-select previously saved values
    const savedPhoneNumberColumn = localStorage.getItem('phoneNumberColumn');
    const savedNameColumn = localStorage.getItem('nameColumn');

    if (savedPhoneNumberColumn && availableHeaders.includes(savedPhoneNumberColumn)) {
        phoneNumberColumnSelect.value = savedPhoneNumberColumn;
    }
    if (savedNameColumn && availableHeaders.includes(savedNameColumn)) {
        nameColumnSelect.value = savedNameColumn;
    }
}


// --- Load initial settings ---
messageTemplateInput.value = localStorage.getItem('messageTemplate') || 'Hello {{Name}}, This is a sample message!';
countryCodeInput.value = localStorage.getItem('countryCode') || '91';

selectedScreenshotDir = localStorage.getItem('screenshotDir') || ''; // Load saved dir or empty for default
if (selectedScreenshotDir) {
    screenshotDirInput.value = selectedScreenshotDir;
    screenshotDirInput.placeholder = ''; // Clear placeholder if custom path is set
} else {
    screenshotDirInput.placeholder = 'Using default: ./failure_screenshots'; // Set default placeholder
    screenshotDirInput.value = ''; // Ensure input value is clear if using default placeholder
}


const lastExcelPath = localStorage.getItem('lastExcelPath');
if (lastExcelPath) {
    selectedExcelFilePath = lastExcelPath;
    const fileName = lastExcelPath.split(/[\\/]/).pop();
    filePathDisplay.textContent = `Last used: ${fileName}`;
    // If a previous file path exists, request its headers
    ipcRenderer.send('get-excel-headers', selectedExcelFilePath);
} else {
    populateColumnDropdowns([]); // Initialize with empty/placeholder
}


// --- Event Listeners ---
selectFileBtn.addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog');
});

ipcRenderer.on('selected-file', (event, filePath) => {
    if (filePath) {
        selectedExcelFilePath = filePath;
        const fileName = filePath.split(/[\\/]/).pop();
        filePathDisplay.textContent = fileName;
        localStorage.setItem('lastExcelPath', filePath);

        // Request headers for the newly selected file
        logMessage('Reading Excel headers...', 'info');
        phoneNumberColumnSelect.innerHTML = '<option value="">-- Loading Headers... --</option>';
        nameColumnSelect.innerHTML = '<option value="">-- Loading Headers... --</option>';
        ipcRenderer.send('get-excel-headers', filePath);
    }
});

// Listen for headers from the main process
ipcRenderer.on('excel-headers-list', (event, headers) => {
    if (headers && headers.error) {
        logMessage(`Error reading headers: ${headers.error}`, 'error');
        populateColumnDropdowns([]);
    } else {
        logMessage('Excel headers received.', 'info');
        populateColumnDropdowns(headers);
    }
});

// Screenshot directory selection
selectScreenshotDirBtn.addEventListener('click', () => {
    ipcRenderer.send('open-screenshot-dir-dialog');
});

ipcRenderer.on('selected-screenshot-dir', (event, dirPath) => {
    if (dirPath) {
        selectedScreenshotDir = dirPath;
        screenshotDirInput.value = dirPath;
        screenshotDirInput.placeholder = ''; // Clear placeholder
        localStorage.setItem('screenshotDir', dirPath);
        logMessage(`Screenshot directory set to: ${dirPath}`, 'info');
    }
});

useDefaultScreenshotDirBtn.addEventListener('click', () => {
    selectedScreenshotDir = ''; // Empty string signifies default
    screenshotDirInput.value = ''; // Clear the input field
    screenshotDirInput.placeholder = 'Using default: ./failure_screenshots';
    localStorage.removeItem('screenshotDir'); // Or set to empty string: localStorage.setItem('screenshotDir', '');
    logMessage('Screenshot directory reset to default.', 'info');
});

startBtn.addEventListener('click', () => {
    if (!selectedExcelFilePath) {
        logMessage('Error: Please select an Excel file first.', 'error');
        alert('Please select an Excel file first.');
        return;
    }

    const phoneNumberColumn = phoneNumberColumnSelect.value;
    const nameColumn = nameColumnSelect.value;
    const messageTemplate = messageTemplateInput.value;
    const countryCode = countryCodeInput.value;

    if (!phoneNumberColumn) {
        logMessage('Error: Please select the Phone Number Column from the dropdown.', 'error');
        alert('Please select the Phone Number Column.');
        return;
    }
    if (!nameColumn) {
        logMessage('Error: Please select the Name Column from the dropdown.', 'error');
        alert('Please select the Name Column.');
        return;
    }
    if (!messageTemplate.trim()) {
        logMessage('Error: Message template cannot be empty.', 'error');
        alert('Message template cannot be empty.');
        return;
    }

    // Save settings for next time
    localStorage.setItem('phoneNumberColumn', phoneNumberColumn); // Save selected dropdown value
    localStorage.setItem('nameColumn', nameColumn);               // Save selected dropdown value
    localStorage.setItem('messageTemplate', messageTemplate);
    localStorage.setItem('countryCode', countryCode);
    // Screenshot dir is already saved on selection/reset

    logOutput.textContent = ''; // Clear previous logs
    logMessage('Starting process...', 'info');
    startBtn.disabled = true;
    selectFileBtn.disabled = true;
    phoneNumberColumnSelect.disabled = true;
    nameColumnSelect.disabled = true;
    selectScreenshotDirBtn.disabled = true;
    useDefaultScreenshotDirBtn.disabled = true;

    ipcRenderer.send('start-sending-messages', {
        excelFilePath: selectedExcelFilePath,
        phoneNumberColumn: phoneNumberColumn,
        nameColumn: nameColumn,
        messageTemplate: messageTemplate,
        countryCode: countryCode,
        screenshotDir: selectedScreenshotDir // Send empty if default, otherwise the path
    });
});

function reEnableControls() {
    startBtn.disabled = false;
    selectFileBtn.disabled = false;
    phoneNumberColumnSelect.disabled = false;
    nameColumnSelect.disabled = false;
    selectScreenshotDirBtn.disabled = false;
    useDefaultScreenshotDirBtn.disabled = false;
}

ipcRenderer.on('log-message', (event, { message, type }) => {
    logMessage(message, type);
});

ipcRenderer.on('process-finished', (event, stats) => {
    logMessage('------------------------------------', 'info');
    logMessage('Process Finished!', 'info');
    logMessage(`Stats: \nSuccessful - ${stats.successful_messages_count} \nFailed - ${stats.failed_messages_count} \nInvalid Phone Numbers - ${stats.invalid_phone_numbers_count}`, 'info');
    if (stats.failed_ids && stats.failed_ids.length > 0) {
        logMessage(`Failed for Names/IDs: ${stats.failed_ids.join(', ')}`, 'warn');
    }
    if (stats.invalid_phone_numbers_ids && stats.invalid_phone_numbers_ids.length > 0) {
        logMessage(`Invalid Numbers for Names/IDs: ${stats.invalid_phone_numbers_ids.join(', ')}`, 'warn');
    }
    logMessage('------------------------------------', 'info');
    reEnableControls();
});

ipcRenderer.on('error-occurred', (event, errorMessage) => {
    logMessage(`CRITICAL ERROR: ${errorMessage}`, 'error');
    logMessage('Process aborted due to a critical error.', 'error');
    reEnableControls();
});


function logMessage(message, type = 'info') {
    const now = new Date().toLocaleTimeString();
    const logEntry = document.createElement('div');
    logEntry.textContent = `[${now}] ${message}`;
    if (type === 'error') {
        logEntry.style.color = 'red';
    } else if (type === 'warn') {
        logEntry.style.color = 'orange';
    } else { // info
        logEntry.style.color = 'green';
    }
    logOutput.appendChild(logEntry);
    logOutput.scrollTop = logOutput.scrollHeight; // Auto-scroll
}