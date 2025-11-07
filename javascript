// ====== CHART GENERATION WITH IMPROVEMENTS ======
/**
 * Generates a chart based on the specified type
 * @param {string} chartType - Must be one of: 
 *   'monthly-sales', 'monthly-report', 'channel-analysis',
 *   'product-performance', 'customer-segment'
 * @throws {Error} If chartType is invalid or creation fails
 */
function generateChart(chartType) {
  const container = document.getElementById('chart-container');
  if (!container) {
    console.error('Chart container not found');
    return;
  }

  // Chart handlers mapping (replaces switch-case)
  const chartHandlers = {
    'monthly-sales': createMonthlySalesChart,
    'monthly-report': createMonthlyReportChart,
    'channel-analysis': createChannelComparisonChart,
    'product-performance': createProductPerformanceChart,
    'customer-segment': createCustomerSegmentChart
  };

  // Validate chart type
  if (!chartHandlers[chartType]) {
    const error = new Error(`Invalid chart type: ${chartType}`);
    console.error(error.message);
    showNotification('error', 'Invalid Chart Type', `Unsupported chart type: ${chartType}`);
    return;
  }

  try {
    // Verify handler function exists
    if (typeof chartHandlers[chartType] !== 'function') {
      throw new Error(`Chart handler for '${chartType}' is not a function`);
    }

    // Execute chart creation
    chartHandlers[chartType]();
    
    // Update timestamp with enhanced function
    updateLastUpdateTime('last-update-time');
    
  } catch (error) {
    console.error(`Chart generation failed: ${error.message}`);
    showNotification('error', 'Chart Error', `Failed to generate chart: ${error.message}`);
  }
}

/**
 * Updates the last update time display
 * @param {string} elementId - The ID of the element to update. Defaults to 'last-update-time'
 */
function updateLastUpdateTime(elementId = 'last-update-time') {
  const element = document.getElementById(elementId);
  if (element) {
    element.textContent = new Date().toLocaleString();
  } else {
    console.warn(`Element with ID '${elementId}' not found for time update.`);
  }
}

// ====== DATA MANAGEMENT ======
// Fixed addData function with enhanced validation
function addData() {
    console.log('addData function called');
    
    const form = document.getElementById('salesForm');
    if (!form) {
        console.error('Form not found');
        return;
    }
    
    // Get form values
    const date = document.getElementById('entryDate').value;
    const branch = document.getElementById('entryBranch').value;
    const channel = document.getElementById('entryChannel').value;
    const transactions = document.getElementById('entryTransactions').value;
    const sales = document.getElementById('entrySales').value;
    const product = document.getElementById('entryProduct')?.value || 'General'; // New field
    const salesRep = document.getElementById('entrySalesRep')?.value || 'Unassigned'; // New field
    
    console.log('Form values:', { date, branch, channel, transactions, sales, product, salesRep });
    
    // Enhanced validation
    const validationErrors = [];
    
    if (!date) validationErrors.push('Please select a date');
    if (!branch) validationErrors.push('Please select a branch');
    if (!channel) validationErrors.push('Please select a channel');
    if (!transactions || transactions <= 0) validationErrors.push('Please enter a valid transaction count');
    if (!sales || sales <= 0) validationErrors.push('Please enter a valid sales amount');
    
    if (validationErrors.length > 0) {
        showNotification('error', 'Validation Error', validationErrors.join('<br>'));
        return;
    }
    
    // Convert to numbers
    const transactionsNum = parseInt(transactions);
    const salesNum = parseFloat(sales);
    const avgSale = transactionsNum > 0 ? salesNum / transactionsNum : 0;
    
    // Create enhanced data object
    const newData = {
        date,
        branch,
        channel,
        count: transactionsNum,
        sales: salesNum,
        avr: parseFloat(avgSale.toFixed(2)),
        product, // New field
        salesRep, // New field
        id: Date.now()
    };
    
    console.log('New data to add:', newData);
    
    // Add to array
    if (editingIndex >= 0) {
        // Update existing entry
        allData[editingIndex] = newData;
        editingIndex = -1;
        document.getElementById('addDataBtn').innerHTML = '<i class="bi bi-plus-circle me-2"></i> Add Data';
        document.getElementById('cancelBtn').style.display = 'none';
        showNotification('success', 'Success', 'Data updated successfully!');
    } else {
        // Add new entry
        allData.push(newData);
        showNotification('success', 'Success', 'Data added successfully!');
    }
    
    // Reset form
    form.reset();
    form.classList.remove('was-validated');
    
    // Update UI
    initializeApp();
}

// Fixed processFile function with enhanced error handling
function processFile(file) {
    console.log('Processing file:', file.name, file.type, file.size);
    
    // Validate file type
    if (!file.name.match(/\.(xlsx|xls|csv|json)$/i)) {
        showNotification('error', 'Invalid File', 'Please upload an Excel, CSV or JSON file');
        return;
    }
    
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            console.log('File read successfully, size:', e.target.result.byteLength);
            
            let processedData = [];
            
            if (file.name.match(/\.(xlsx|xls)$/i)) {
                // Process Excel file
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    raw: false,
                    codepage: 65001 // UTF-8
                });
                
                if (workbook.SheetNames.length === 0) {
                    throw new Error('No sheets found in the Excel file');
                }
                
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON with header
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    header: 1,
                    raw: false,
                    defval: ''
                });
                
                processedData = processExcelData(jsonData);
                
            } else if (file.name.match(/\.csv$/i)) {
                // Process CSV file
                const csv = e.target.result;
                const parsedData = Papa.parse(csv, { 
                    header: true,
                    skipEmptyLines: true,
                    transform: (value) => value.trim()
                });
                
                processedData = processCSVData(parsedData.data);
                
            } else if (file.name.match(/\.json$/i)) {
                // Process JSON file
                const jsonData = JSON.parse(e.target.result);
                processedData = processJSONData(jsonData);
            }
            
            if (processedData.length === 0) {
                throw new Error('No valid data found in the file');
            }
            
            // Add to existing data
            allData = [...allData, ...processedData];
            
            // Update UI
            initializeApp();
            
            showNotification('success', 'Success', `Successfully imported ${processedData.length} records from ${file.name}`);
            
        } catch (error) {
            console.error('Error processing file:', error);
            showNotification('error', 'Error', `Error processing file: ${error.message}`);
        }
    };
    
    reader.onerror = function(error) {
        console.error('File reading error:', error);
        showNotification('error', 'Error', 'Error reading file');
    };
    
    if (file.name.match(/\.(xlsx|xls)$/i)) {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsText(file);
    }
}

// New function to process CSV data
function processCSVData(csvData) {
    console.log('Processing CSV data, rows:', csvData.length);
    
    const salesData = [];
    
    // Expected headers (case insensitive)
    const expectedHeaders = ['date', 'branch', 'channel', 'transactions', 'sales'];
    const headerMap = {};
    
    // Create header mapping
    Object.keys(csvData[0]).forEach((header, index) => {
        const normalizedHeader = header.toLowerCase().trim();
        if (expectedHeaders.includes(normalizedHeader)) {
            headerMap[normalizedHeader] = header;
        }
    });
    
    console.log('Header mapping:', headerMap);
    
    // Validate required headers
    const missingHeaders = expectedHeaders.filter(h => !(h in headerMap));
    if (missingHeaders.length > 0) {
        throw new Error(`Missing required columns: ${missingHeaders.join(', ')}`);
    }
    
    // Process data rows
    for (let i = 0; i < csvData.length; i++) {
        const row = csvData[i];
        
        try {
            // Get values using header mapping
            const dateValue = row[headerMap.date];
            const branchValue = row[headerMap.branch];
            const channelValue = row[headerMap.channel] || 'Delivery';
            const transactionsValue = row[headerMap.transactions];
            const salesValue = row[headerMap.sales];
            
            // Skip empty rows
            if (!dateValue || !branchValue) {
                console.warn(`Skipping row ${i}: Missing date or branch`);
                continue;
            }
            
            // Convert and validate data
            let date = dateValue;
            
            // Validate date format
            if (!isValidDate(date)) {
                console.warn(`Invalid date format in row ${i}: ${dateValue}`);
                continue;
            }
            
            // Convert transactions and sales to numbers
            const transactions = parseInt(transactionsValue) || 0;
            const sales = parseFloat(salesValue) || 0;
            
            if (transactions <= 0 || sales <= 0) {
                console.warn(`Invalid transactions or sales in row ${i}: ${transactionsValue}, ${salesValue}`);
                continue;
            }
            
            // Calculate average sale
            const avgSale = transactions > 0 ? sales / transactions : 0;
            
            // Create data object
            const dataEntry = {
                date,
                branch: branchValue.trim(),
                channel: channelValue.trim(),
                count: transactions,
                sales: sales,
                avr: parseFloat(avgSale.toFixed(2)),
                id: Date.now() + i
            };
            
            salesData.push(dataEntry);
            
        } catch (error) {
            console.error(`Error processing row ${i}:`, error);
        }
    }
    
    console.log(`Processed ${salesData.length} valid entries`);
    return salesData;
}

// New function to process JSON data
function processJSONData(jsonData) {
    console.log('Processing JSON data, entries:', jsonData.length);
    
    const salesData = [];
    
    // Validate that it's an array
    if (!Array.isArray(jsonData)) {
        throw new Error('JSON data must be an array of objects');
    }
    
    // Expected fields (case insensitive)
    const expectedFields = ['date', 'branch', 'channel', 'transactions', 'sales'];
    
    for (let i = 0; i < jsonData.length; i++) {
        const entry = jsonData[i];
        
        try {
            // Normalize keys
            const normalizedEntry = {};
            Object.keys(entry).forEach(key => {
                normalizedEntry[key.toLowerCase().trim()] = entry[key];
            });
            
            // Check required fields
            const missingFields = expectedFields.filter(field => !normalizedEntry[field]);
            if (missingFields.length > 0) {
                console.warn(`Skipping entry ${i}: Missing fields: ${missingFields.join(', ')}`);
                continue;
            }
            
            // Extract values
            const dateValue = normalizedEntry.date;
            const branchValue = normalizedEntry.branch;
            const channelValue = normalizedEntry.channel || 'Delivery';
            const transactionsValue = normalizedEntry.transactions;
            const salesValue = normalizedEntry.sales;
            
            // Validate date
            if (!isValidDate(dateValue)) {
                console.warn(`Invalid date format in entry ${i}: ${dateValue}`);
                continue;
            }
            
            // Convert numbers
            const transactions = parseInt(transactionsValue) || 0;
            const sales = parseFloat(salesValue) || 0;
            
            if (transactions <= 0 || sales <= 0) {
                console.warn(`Invalid transactions or sales in entry ${i}: ${transactionsValue}, ${salesValue}`);
                continue;
            }
            
            // Calculate average
            const avgSale = transactions > 0 ? sales / transactions : 0;
            
            // Create data object
            const dataEntry = {
                date: dateValue,
                branch: branchValue.trim(),
                channel: channelValue.trim(),
                count: transactions,
                sales: sales,
                avr: parseFloat(avgSale.toFixed(2)),
                id: Date.now() + i
            };
            
            salesData.push(dataEntry);
            
        } catch (error) {
            console.error(`Error processing entry ${i}:`, error);
        }
    }
    
    console.log(`Processed ${salesData.length} valid entries`);
    return salesData;
}

// ====== EVENT LISTENERS ======
// Fixed event listeners setup with enhanced error handling
function setupEventListeners() {
    console.log('Setting up event listeners');
    
    try {
        // Language toggle
        const langToggle = document.getElementById('langToggle');
        if (langToggle) {
            langToggle.addEventListener('click', toggleLanguage);
        }
        
        // Dark mode toggle
        const darkModeToggle = document.getElementById('darkModeToggle');
        if (darkModeToggle) {
            darkModeToggle.addEventListener('click', toggleDarkMode);
        }
        
        // Add data button
        const addDataBtn = document.getElementById('addDataBtn');
        if (addDataBtn) {
            addDataBtn.addEventListener('click', addData);
        }
        
        // Cancel button
        const cancelBtn = document.getElementById('cancelBtn');
        if (cancelBtn) {
            cancelBtn.addEventListener('click', cancelEdit);
        }
        
        // Form submission
        const salesForm = document.getElementById('salesForm');
        if (salesForm) {
            salesForm.addEventListener('submit', function(e) {
                e.preventDefault();
                addData();
            });
        }
        
        // Update overall target
        const updateOverallTargetBtn = document.getElementById('updateOverallTargetBtn');
        if (updateOverallTargetBtn) {
            updateOverallTargetBtn.addEventListener('click', updateOverallTarget);
        }
        
        // Generate report button
        const generateReportBtn = document.getElementById('generateReportBtn');
        if (generateReportBtn) {
            generateReportBtn.addEventListener('click', updateBranchPerformanceReport);
        }
        
        // Export buttons
        const exportCSV = document.getElementById('exportCSV');
        if (exportCSV) {
            exportCSV.addEventListener('click', exportToCSV);
        }
        
        const exportJSON = document.getElementById('exportJSON');
        if (exportJSON) {
            exportJSON.addEventListener('click', exportToJSON);
        }
        
        const exportDaily = document.getElementById('exportDaily');
        if (exportDaily) {
            exportDaily.addEventListener('click', () => generatePDFReport('daily'));
        }
        
        const exportWeekly = document.getElementById('exportWeekly');
        if (exportWeekly) {
            exportWeekly.addEventListener('click', () => generatePDFReport('weekly'));
        }
        
        const exportMonthly = document.getElementById('exportMonthly');
        if (exportMonthly) {
            exportMonthly.addEventListener('click', () => generatePDFReport('monthly'));
        }
        
        // Save and clear data buttons
        const saveDataBtn = document.getElementById('saveDataBtn');
        if (saveDataBtn) {
            saveDataBtn.addEventListener('click', saveData);
        }
        
        const clearDataBtn = document.getElementById('clearDataBtn');
        if (clearDataBtn) {
            clearDataBtn.addEventListener('click', clearData);
        }
        
        // Backup buttons
        const createBackupBtn = document.getElementById('createBackupBtn');
        if (createBackupBtn) {
            createBackupBtn.addEventListener('click', createBackup);
        }
        
        const restoreBackupBtn = document.getElementById('restoreBackupBtn');
        if (restoreBackupBtn) {
            restoreBackupBtn.addEventListener('click', restoreBackup);
        }
        
        // File upload
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        
        if (uploadArea && fileInput) {
            uploadArea.addEventListener('click', () => fileInput.click());
            uploadArea.addEventListener('dragover', handleDragOver);
            uploadArea.addEventListener('drop', handleDrop);
            fileInput.addEventListener('change', handleFileSelect);
        }
        
        // Refresh button
        const refreshBtn = document.getElementById('refreshBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', function() {
                initializeApp();
            });
        }
        
        // Tab change event
        const tabButtons = document.querySelectorAll('#mainTabs button[data-bs-toggle="tab"]');
        tabButtons.forEach(tab => {
            tab.addEventListener('shown.bs.tab', function (event) {
                if (event.target.id === 'branch-performance-tab') {
                    updateBranchPerformanceReport();
                } else if (event.target.id === 'report-arrangement-tab') {
                    updateReportArrangement();
                }
            });
        });
        
        // Authentication events
        const closeAuthModal = document.getElementById('closeAuthModal');
        const cancelLogin = document.getElementById('cancelLogin');
        const loginBtn = document.getElementById('loginBtn');
        const userProfileLink = document.getElementById('userProfileLink');
        
        if (closeAuthModal) closeAuthModal.addEventListener('click', hideAuthModal);
        if (cancelLogin) cancelLogin.addEventListener('click', hideAuthModal);
        if (loginBtn) loginBtn.addEventListener('click', login);
        if (userProfileLink) {
            userProfileLink.addEventListener('click', function(e) {
                e.preventDefault();
                if (isAuthenticated) {
                    logout();
                } else {
                    showAuthModal();
                }
            });
        }
        
        // Advanced filters
        const applyFiltersBtn = document.getElementById('applyFiltersBtn');
        const clearFiltersBtn = document.getElementById('clearFiltersBtn');
        
        if (applyFiltersBtn) applyFiltersBtn.addEventListener('click', applyAdvancedFilters);
        if (clearFiltersBtn) clearFiltersBtn.addEventListener('click', clearAdvancedFilters);
        
        // Report arrangement events
        const arrangeReportsBtn = document.getElementById('arrangeReportsBtn');
        const resetArrangementBtn = document.getElementById('resetArrangementBtn');
        
        if (arrangeReportsBtn) arrangeReportsBtn.addEventListener('click', arrangeReports);
        if (resetArrangementBtn) resetArrangementBtn.addEventListener('click', resetReportArrangement);
        
        // Handle Enter key in login form
        const password = document.getElementById('password');
        if (password) {
            password.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    login();
                }
            });
        }
        
        console.log('Event listeners setup complete');
    } catch (error) {
        console.error('Error setting up event listeners:', error);
        showNotification('error', 'Setup Error', 'Failed to set up event listeners');
    }
}

// Fixed handleFileSelect and handleDrop functions
function handleFileSelect(e) {
    console.log('File selected');
    const files = e.target.files;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
}

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    const uploadArea = document.getElementById('uploadArea');
    if (uploadArea) {
        uploadArea.classList.add('drag-over');
    }
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    
    const uploadArea = document.getElementById('uploadArea');
    if (uploadArea) {
        uploadArea.classList.remove('drag-over');
    }
    
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
        processFile(files[0]);
    }
}

// ====== HELPER FUNCTIONS ======
// Helper function to validate date format
function isValidDate(dateString) {
    if (!dateString) return false;
    
    // Check if it's in MM/DD/YYYY format
    const dateRegex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
    if (!dateRegex.test(dateString)) return false;
    
    // Try to parse the date
    const parts = dateString.split('/');
    const month = parseInt(parts[0], 10);
    const day = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    
    // Check if it's a valid date
    const date = new Date(year, month - 1, day);
    return (
        date.getFullYear() === year &&
        date.getMonth() === month - 1 &&
        date.getDate() === day
    );
}

// Fixed convertExcelDate function
function convertExcelDate(dateNum) {
    if (typeof dateNum !== 'number') {
        return dateNum;
    }
    
    // Excel date is days since 1900-01-01, with 1900 incorrectly treated as leap year
    const excelEpoch = new Date(1900, 0, 1);
    const isLeapYearBug = dateNum > 59;
    const adjustedDateNum = isLeapYearBug ? dateNum + 1 : dateNum;
    
    const date = new Date(excelEpoch.getTime() + (adjustedDateNum - 1) * 24 * 60 * 60 * 1000);
    
    // Format as MM/DD/YYYY
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${month}/${day}/${year}`;
}

// ====== INITIALIZATION ======
// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    console.log('Initializing application');
    
    try {
        // Check required libraries
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS library not loaded');
        }
        
        if (typeof Papa === 'undefined') {
            throw new Error('Papa Parse library not loaded');
        }
        
        // Setup event listeners
        setupEventListeners();
        
        // Initialize app
        initializeApp();
        
        console.log('Application initialized successfully');
    } catch (error) {
        console.error('Initialization error:', error);
        showNotification('error', 'Initialization Error', `Failed to initialize: ${error.message}`);
    }
});
