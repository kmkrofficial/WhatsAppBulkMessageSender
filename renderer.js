const { ipcRenderer } = require('electron');

const selectFileBtn = document.getElementById('selectFileBtn');
const excelFileInput = document.getElementById('excelFileInput');
const filePathDisplay = document.getElementById('filePath');
const phoneNumberColumnSelect = document.getElementById('phoneNumberColumn');
const nameColumnSelect = document.getElementById('nameColumn');
const screenshotDirInput = document.getElementById('screenshotDirInput');
const selectScreenshotDirBtn = document.getElementById('selectScreenshotDirBtn');
const useDefaultScreenshotDirBtn = document.getElementById('useDefaultScreenshotDirBtn');
const messageTemplateInput = document.getElementById('messageTemplate');
const countryCodeInput = document.getElementById('countryCode');
const startBtn = document.getElementById('startBtn');
const logOutput = document.getElementById('logOutput');

const progressSection = document.getElementById('progressSection');
const progressBar = document.getElementById('progressBar');
const progressText = document.getElementById('progressText');
const etaText = document.getElementById('etaText');
const statTotalContacts = document.getElementById('statTotalContacts');
const statProcessed = document.getElementById('statProcessed');
const statSuccessful = document.getElementById('statSuccessful');
const statFailed = document.getElementById('statFailed');
const statInvalid = document.getElementById('statInvalid');
const statRemaining = document.getElementById('statRemaining');

let selectedExcelFilePath = '';
let availableHeaders = [];
let selectedScreenshotDir = '';

function populateColumnDropdowns(headers) {
    availableHeaders = headers || [];
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

    const savedPhoneNumberColumn = localStorage.getItem('phoneNumberColumn');
    const savedNameColumn = localStorage.getItem('nameColumn');
    if (savedPhoneNumberColumn && availableHeaders.includes(savedPhoneNumberColumn)) phoneNumberColumnSelect.value = savedPhoneNumberColumn;
    if (savedNameColumn && availableHeaders.includes(savedNameColumn)) nameColumnSelect.value = savedNameColumn;
}

messageTemplateInput.value = localStorage.getItem('messageTemplate') || 'Hello {{Name}}, This is a sample message!';
countryCodeInput.value = localStorage.getItem('countryCode') || '91';
selectedScreenshotDir = localStorage.getItem('screenshotDir') || '';
if (selectedScreenshotDir) {
    screenshotDirInput.value = selectedScreenshotDir;
    screenshotDirInput.placeholder = '';
} else {
    screenshotDirInput.placeholder = 'Using default: ./failure_screenshots';
    screenshotDirInput.value = '';
}

const lastExcelPath = localStorage.getItem('lastExcelPath');
if (lastExcelPath) {
    selectedExcelFilePath = lastExcelPath;
    filePathDisplay.textContent = `Last used: ${lastExcelPath.split(/[\\/]/).pop()}`;
    ipcRenderer.send('get-excel-headers', selectedExcelFilePath);
} else {
    populateColumnDropdowns([]);
}

selectFileBtn.addEventListener('click', () => ipcRenderer.send('open-file-dialog'));

ipcRenderer.on('selected-file', (event, filePath) => {
    if (filePath) {
        selectedExcelFilePath = filePath;
        filePathDisplay.textContent = filePath.split(/[\\/]/).pop();
        localStorage.setItem('lastExcelPath', filePath);
        logMessage('Reading Excel headers...', 'info');
        phoneNumberColumnSelect.innerHTML = '<option value="">-- Loading Headers... --</option>';
        nameColumnSelect.innerHTML = '<option value="">-- Loading Headers... --</option>';
        ipcRenderer.send('get-excel-headers', filePath);
    }
});

ipcRenderer.on('excel-headers-list', (event, headers) => {
    if (headers && headers.error) {
        logMessage(`Error reading headers: ${headers.error}`, 'error');
        populateColumnDropdowns([]);
    } else {
        logMessage('Excel headers received.', 'info');
        populateColumnDropdowns(headers);
    }
});

selectScreenshotDirBtn.addEventListener('click', () => ipcRenderer.send('open-screenshot-dir-dialog'));

ipcRenderer.on('selected-screenshot-dir', (event, dirPath) => {
    if (dirPath) {
        selectedScreenshotDir = dirPath;
        screenshotDirInput.value = dirPath;
        screenshotDirInput.placeholder = '';
        localStorage.setItem('screenshotDir', dirPath);
        logMessage(`Screenshot directory set to: ${dirPath}`, 'info');
    }
});

useDefaultScreenshotDirBtn.addEventListener('click', () => {
    selectedScreenshotDir = '';
    screenshotDirInput.value = '';
    screenshotDirInput.placeholder = 'Using default: ./failure_screenshots';
    localStorage.removeItem('screenshotDir');
    logMessage('Screenshot directory reset to default.', 'info');
});

function resetProgressUI() {
    progressSection.style.display = 'none';
    progressBar.style.width = '0%';
    progressText.textContent = '0%';
    etaText.textContent = 'Calculating...';
    statTotalContacts.textContent = '0';
    statProcessed.textContent = '0';
    statSuccessful.textContent = '0';
    statFailed.textContent = '0';
    statInvalid.textContent = '0';
    statRemaining.textContent = '0';
}

startBtn.addEventListener('click', () => {
    if (!selectedExcelFilePath) { alert('Please select an Excel file first.'); return; }
    const phoneNumberColumn = phoneNumberColumnSelect.value;
    const nameColumn = nameColumnSelect.value;
    const messageTemplate = messageTemplateInput.value;
    if (!phoneNumberColumn) { alert('Please select the Phone Number Column.'); return; }
    if (!nameColumn) { alert('Please select the Name Column.'); return; }
    if (!messageTemplate.trim()) { alert('Message template cannot be empty.'); return; }

    localStorage.setItem('phoneNumberColumn', phoneNumberColumn);
    localStorage.setItem('nameColumn', nameColumn);
    localStorage.setItem('messageTemplate', messageTemplate);
    localStorage.setItem('countryCode', countryCodeInput.value);

    logOutput.textContent = '';
    logMessage('Starting process...', 'info');
    
    progressSection.style.display = 'block';
    resetProgressUI(); 
    progressSection.style.display = 'block'; // Ensure it's visible after reset
    statTotalContacts.textContent = 'Loading...';

    startBtn.disabled = true;
    selectFileBtn.disabled = true;
    phoneNumberColumnSelect.disabled = true;
    nameColumnSelect.disabled = true;
    selectScreenshotDirBtn.disabled = true;
    useDefaultScreenshotDirBtn.disabled = true;

    ipcRenderer.send('start-sending-messages', {
        excelFilePath: selectedExcelFilePath,
        phoneNumberColumn: phoneNumberColumn, nameColumn: nameColumn,
        messageTemplate: messageTemplate, countryCode: countryCodeInput.value,
        screenshotDir: selectedScreenshotDir
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

ipcRenderer.on('log-message', (event, { message, type }) => logMessage(message, type));

ipcRenderer.on('update-progress', (event, data) => {
    if (!progressSection || progressSection.style.display === 'none') {
        progressSection.style.display = 'block';
    }
    progressBar.style.width = `${data.percentage}%`;
    progressText.textContent = `${data.percentage.toFixed(1)}%`;
    etaText.textContent = data.etaFormatted ? data.etaFormatted : (data.percentage === 0 ? 'Calculating...' : etaText.textContent);
    statTotalContacts.textContent = data.totalContacts !== undefined ? data.totalContacts : 'N/A';
    statProcessed.textContent = data.processedCount !== undefined ? data.processedCount : '0';
    statSuccessful.textContent = data.successful !== undefined ? data.successful : '0';
    statFailed.textContent = data.failed !== undefined ? data.failed : '0';
    statInvalid.textContent = data.invalid !== undefined ? data.invalid : '0';
    const remaining = (data.totalContacts || 0) - (data.processedCount || 0);
    statRemaining.textContent = remaining >= 0 ? remaining : '0';
});

ipcRenderer.on('process-finished', (event, stats) => {
    progressBar.style.width = '100%';
    progressText.textContent = '100% Complete';
    etaText.textContent = 'Completed!';
    statTotalContacts.textContent = stats.totalContacts !== undefined ? stats.totalContacts : (stats.successful_messages_count + stats.failed_messages_count + stats.invalid_phone_numbers_count);
    statProcessed.textContent = statTotalContacts.textContent;
    statSuccessful.textContent = stats.successful_messages_count;
    statFailed.textContent = stats.failed_messages_count;
    statInvalid.textContent = stats.invalid_phone_numbers_count;
    statRemaining.textContent = '0';

    logMessage('------------------------------------', 'info');
    logMessage('Process Finished!', 'info');
    if (stats.failed_ids && stats.failed_ids.length > 0) logMessage(`Failed for Names/IDs: ${stats.failed_ids.join(', ')}`, 'warn');
    if (stats.invalid_phone_numbers_ids && stats.invalid_phone_numbers_ids.length > 0) logMessage(`Invalid Numbers for Names/IDs: ${stats.invalid_phone_numbers_ids.join(', ')}`, 'warn');
    logMessage('------------------------------------', 'info');
    reEnableControls();
});

ipcRenderer.on('error-occurred', (event, errorMessage) => {
    logMessage(`CRITICAL ERROR: ${errorMessage}`, 'error');
    logMessage('Process aborted due to a critical error.', 'error');
    etaText.textContent = 'Error!';
    reEnableControls();
});

function logMessage(message, type = 'info') {
    const now = new Date().toLocaleTimeString();
    const logEntry = document.createElement('div');
    logEntry.textContent = `[${now}] ${message}`;
    if (type === 'error') logEntry.style.color = 'red';
    else if (type === 'warn') logEntry.style.color = 'orange';
    else logEntry.style.color = 'green';
    logOutput.appendChild(logEntry);
    logOutput.scrollTop = logOutput.scrollHeight;
}
resetProgressUI();