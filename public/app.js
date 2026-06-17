// Global variables
let budgetData = {
    requestSummary: [],
    personnel: [],
    nonPersonnel: [],
    requestQA: [],
    budgetSummary: []
};
let filteredData = [];
let currentBudgetData = []; // ← ADD THIS LINE
// Program-level structured PBB attributes (Mandate, Cost Recovery, Final Score) sourced
// from the Programs Inventory's "Details" sheet. Falls back into the scoring engine when
// the Q&A narrative is empty or import boilerplate. Keyed by `${DEPT}::${PROGRAM}` and
// `*::${PROGRAM}` (uppercased) for tolerant lookup.
let programAttributesMap = {};

// Per-client toggle for the Access/Equity criterion. Some clients explicitly do not want
// equity to factor into PBB scoring or appear in the report. Default ON; persisted in
// localStorage so the choice survives reloads. Toggling off:
//   - excludes Access from totalScore / weightedScore
//   - hides the Access/Equity panels in the UI and exports
//   - suppresses the equity ask in the narrative
let includeAccessEquity = (typeof localStorage !== 'undefined')
    ? localStorage.getItem('pbb_include_access_equity') !== 'false'
    : true;

function setIncludeAccessEquity(on) {
    includeAccessEquity = !!on;
    if (typeof localStorage !== 'undefined') {
        localStorage.setItem('pbb_include_access_equity', includeAccessEquity ? 'true' : 'false');
    }
    if (typeof updateDecisionDashboard === 'function') updateDecisionDashboard();
}


// DOM elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadStatus = document.getElementById('uploadStatus');
const filtersSection = document.getElementById('filtersSection');
const generateBtn = document.getElementById('generateBtn');
const reportSection = document.getElementById('reportSection');
const reportContent = document.getElementById('reportContent');
const progressBar = document.getElementById('progressBar');
const progressFill = document.getElementById('progressFill');

// File upload handling
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', handleDragOver);
uploadArea.addEventListener('drop', handleDrop);
fileInput.addEventListener('change', handleFileSelect);

// ===== CURRENT BUDGET FILE UPLOAD HANDLING =====
const currentBudgetUploadArea = document.getElementById('currentBudgetUploadArea');
const currentBudgetFileInput = document.getElementById('currentBudgetFileInput');
const currentBudgetStatus = document.getElementById('currentBudgetStatus');

if (currentBudgetUploadArea) {
    currentBudgetUploadArea.addEventListener('click', () => currentBudgetFileInput.click());
    currentBudgetUploadArea.addEventListener('dragover', handleCurrentBudgetDragOver);
    currentBudgetUploadArea.addEventListener('drop', handleCurrentBudgetDrop);
}

if (currentBudgetFileInput) {
    currentBudgetFileInput.addEventListener('change', handleCurrentBudgetFileSelect);
}

// Access/Equity toggle — initialize from persisted state and wire change handler
const accessEquityToggle = document.getElementById('includeAccessEquityToggle');
if (accessEquityToggle) {
    accessEquityToggle.checked = includeAccessEquity;
    accessEquityToggle.addEventListener('change', (e) => setIncludeAccessEquity(e.target.checked));
}

function handleCurrentBudgetDragOver(e) {
    e.preventDefault();
    currentBudgetUploadArea.classList.add('dragover');
}

function handleCurrentBudgetDrop(e) {
    e.preventDefault();
    currentBudgetUploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processCurrentBudgetFile(files[0]);
    }
}

function handleCurrentBudgetFileSelect(e) {
    if (e.target.files.length > 0) {
        processCurrentBudgetFile(e.target.files[0]);
    }
}

function processCurrentBudgetFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showCurrentBudgetMessage('Please select a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }

    showCurrentBudgetMessage('Processing current budget file...', 'loading');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            console.log('Current Budget - Available sheets:', workbook.SheetNames);
            
            // Parse the Programs sheet (or first sheet if "Programs" doesn't exist)
            const programsSheet = workbook.Sheets['Programs'] || workbook.Sheets[workbook.SheetNames[0]];
            currentBudgetData = XLSX.utils.sheet_to_json(programsSheet, { defval: '' });

            console.log(`Loaded ${currentBudgetData.length} programs from current budget`);
            console.log('Sample program:', currentBudgetData[0]);

            // If the workbook also has a "Details" sheet (e.g. ResourceX Program Summary
            // export), build a program-attributes map from it. This gives the PBB scoring
            // engine access to Mandate / Cost Recovery / Final Score even when the request
            // Q&A is missing or boilerplate.
            programAttributesMap = {};
            const detailsSheet = workbook.Sheets['Details'] || workbook.Sheets['details'];
            if (detailsSheet) {
                const detailsRows = XLSX.utils.sheet_to_json(detailsSheet, { defval: '' });
                buildProgramAttributesMap(detailsRows);
                console.log(`Built program attributes for ${Object.keys(programAttributesMap).length} program keys from Details sheet`);
            }
            
            if (currentBudgetData.length > 0) {
                showCurrentBudgetMessage(`✅ Successfully loaded ${currentBudgetData.length} programs with current budget data!`, 'success');
                
                // If budget requests are already loaded, update the stats
                if (budgetData.requestSummary.length > 0) {
                    updateStats();
                }
            } else {
                showCurrentBudgetMessage('No data found in the current budget file', 'error');
            }
        } catch (error) {
            console.error('Error processing current budget file:', error);
            showCurrentBudgetMessage('Error processing file: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function showCurrentBudgetMessage(message, type) {
    const className = type === 'error' ? 'error-message' : 
                     type === 'success' ? 'success-message' : 
                     'loading';
    
    currentBudgetStatus.innerHTML = `<div class="${className}">${message}</div>`;
    
    if (type === 'success') {
        setTimeout(() => {
            currentBudgetStatus.innerHTML = '';
        }, 3000);
    }
}

function handleDragOver(e) {
    e.preventDefault();
    uploadArea.classList.add('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    if (e.target.files.length > 0) {
        processFile(e.target.files[0]);
    }
}

function processFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showMessage('Please select a valid Excel file (.xlsx or .xls)', 'error');
        return;
    }

    showMessage('Processing file...', 'loading');
    
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            console.log('Available sheets:', workbook.SheetNames);
            
            // Parse each sheet with enhanced debugging
            budgetData.requestSummary = parseSheetWithDebug(workbook, 'Request Summary');
            budgetData.personnel = parseSheetWithDebug(workbook, 'Personnel');
            budgetData.nonPersonnel = parseSheetWithDebug(workbook, 'NonPersonnel');
            budgetData.requestQA = parseSheetWithDebug(workbook, 'Request Q&A');
            budgetData.budgetSummary = parseSheetWithDebug(workbook, 'Budget Summary');
            
            console.log('All parsed data:', budgetData);
            
            if (budgetData.requestSummary.length > 0) {
                showMessage(`Successfully loaded ${budgetData.requestSummary.length} budget requests`, 'success');
                setupFilters();
                updateStats();
                filtersSection.style.display = 'block';
                generateBtn.disabled = false;
                
                // NEW: Show the current budget upload section
                const currentBudgetSection = document.getElementById('currentBudgetSection');
                if (currentBudgetSection) {
                    currentBudgetSection.style.display = 'block';
                }
            } else {
                showMessage('No data found in the Request Summary sheet', 'error');
            }
            
        } catch (error) {
            console.error('Error processing file:', error);
            showMessage('Error processing file: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseSheetWithDebug(workbook, sheetName) {
    console.log(`\n=== Parsing ${sheetName} ===`);
    
    if (!workbook.Sheets[sheetName]) {
        console.warn(`Sheet ${sheetName} not found`);
        return [];
    }
    
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    console.log(`Sheet range: ${sheet['!ref']}`);
    
    // Try to find headers by examining the first few rows
    let headerRow = -1;
    let headers = [];
    
    // Look for headers in the first 10 rows
    for (let r = 0; r < Math.min(10, range.e.r + 1); r++) {
        const rowData = [];
        let hasContent = false;
        
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellAddr = XLSX.utils.encode_cell({r: r, c: c});
            const cell = sheet[cellAddr];
            const value = cell ? (cell.v || '').toString().trim() : '';
            rowData.push(value);
            if (value) hasContent = true;
        }
        
        if (hasContent) {
            console.log(`Row ${r}:`, rowData.slice(0, 10)); // First 10 columns
            
            // Check if this row looks like headers
            const joinedRow = rowData.join(' ').toLowerCase();
            
            // Different header detection for different sheets
            let isHeaderRow = false;
            if (sheetName === 'Personnel' || sheetName === 'NonPersonnel') {
                isHeaderRow = joinedRow.includes('request') || 
                             joinedRow.includes('department') || 
                             joinedRow.includes('program') ||
                             joinedRow.includes('position') ||
                             joinedRow.includes('account');
            } else if (sheetName === 'Request Summary') {
                isHeaderRow = joinedRow.includes('request') || 
                             joinedRow.includes('description') ||
                             joinedRow.includes('status');
            } else if (sheetName === 'Request Q&A') {
                isHeaderRow = joinedRow.includes('question') || 
                             joinedRow.includes('answer');
            } else if (sheetName === 'Budget Summary') {
                isHeaderRow = joinedRow.includes('item') || 
                             joinedRow.includes('budget') ||
                             joinedRow.includes('fund');
            }
            
            if (isHeaderRow && headerRow === -1) {
                headerRow = r;
                headers = rowData.map(h => h.trim()).filter(h => h);
                console.log(`Found headers at row ${r}:`, headers);
                break;
            }
        }
    }
    
    if (headerRow === -1) {
        console.warn(`No headers found for ${sheetName}`);
        // Try standard parsing as fallback
        const data = XLSX.utils.sheet_to_json(sheet);
        console.log(`Fallback parsing got ${data.length} rows`);
        return data;
    }
    
    // Parse data rows starting after header row
    const data = [];
    for (let r = headerRow + 1; r <= range.e.r; r++) {
        const row = {};
        let hasData = false;
        
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellAddr = XLSX.utils.encode_cell({r: r, c: c});
            const cell = sheet[cellAddr];
            const value = cell ? cell.v : null;
            
            // Use position-based header mapping if we have headers
            if (c - range.s.c < headers.length && headers[c - range.s.c]) {
                const header = headers[c - range.s.c];
                row[header] = value;
                if (value !== null && value !== undefined && value.toString().trim() !== '') {
                    hasData = true;
                }
            } else if (value !== null && value !== undefined && value.toString().trim() !== '') {
                // Store by column index if no header
                row[`Col_${c}`] = value;
                hasData = true;
            }
        }
        
        if (hasData) {
            data.push(row);
        }
    }
    
    console.log(`Parsed ${data.length} data rows`);
    if (data.length > 0) {
        console.log('Sample row:', data[0]);
        console.log('All keys in first row:', Object.keys(data[0]));
    }
    
    return data;
}

// ADDITIONAL FIX: Enhanced setup filters to use proper field names
function setupFilters() {
    console.log('\n=== Setting up filters (DUAL SOURCE) ===');
    
    const filters = {
        fund: new Set(['all']),
        department: new Set(['all']),
        division: new Set(['all']),
        program: new Set(['all']),
        requestType: new Set(['all']),
        status: new Set(['all'])
    };

    // SOURCE 1: Collect from Budget Request Data
    console.log('Collecting filter values from Budget Request data...');
    
    // Collect unique values from all line items (Personnel + NonPersonnel)
    const allLineItems = [...budgetData.personnel, ...budgetData.nonPersonnel];
    
    allLineItems.forEach((item, idx) => {
        if (idx < 5) console.log(`Line item ${idx}:`, item);
        
        // Use explicit field names instead of fuzzy matching
        if (item.Fund) filters.fund.add(item.Fund);
        if (item.Department) filters.department.add(item.Department);
        if (item['Cost Center']) filters.department.add(item['Cost Center']);
        if (item.Division) filters.division.add(item.Division);
        if (item.Program) filters.program.add(item.Program);
        if (item.Status) filters.status.add(item.Status);
    });
    
    // From Request Summary (for request-level filters)
    budgetData.requestSummary.forEach(item => {
        if (item['Request Type']) filters.requestType.add(item['Request Type']);
        if (item.Status) filters.status.add(item.Status);
    });

    // SOURCE 2: Collect from Current Budget (Program Inventory) if available
    if (currentBudgetData.length > 0) {
        console.log('Collecting filter values from Current Budget (Program Inventory)...');
        
        currentBudgetData.forEach((item, idx) => {
            if (idx < 3) console.log(`Current Budget item ${idx}:`, item);
            
            // User Group is the department
            if (item['User Group']) filters.department.add(item['User Group']);
            // Program Name
            if (item['Program']) filters.program.add(item['Program']);
            // Division if available
            if (item['Division']) filters.division.add(item['Division']);
        });
    }

    console.log('Filter values found (from BOTH sources):', {
        fund: Array.from(filters.fund),
        department: Array.from(filters.department),
        division: Array.from(filters.division),
        program: Array.from(filters.program),
        requestType: Array.from(filters.requestType),
        status: Array.from(filters.status)
    });

    // Populate filter dropdowns
    populateSelect('fundFilter', filters.fund);
    populateSelect('departmentFilter', filters.department);
    populateSelect('divisionFilter', filters.division);
    populateSelect('programFilter', filters.program);
    populateSelect('requestTypeFilter', filters.requestType);
    populateSelect('statusFilter', filters.status);

    // Add event listeners
    document.querySelectorAll('select').forEach(select => {
        select.addEventListener('change', updateStats);
    });
}



function populateSelect(selectId, values) {
    const select = document.getElementById(selectId);
    select.innerHTML = '<option value="all">All</option>';
    
    Array.from(values).sort().forEach(value => {
        if (value !== 'all' && value) {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = value;
            select.appendChild(option);
        }
    });
}

// ISSUE 2 FIX: Corrected line item retrieval to ensure RequestID matches
function getLineItemsForRequest(requestId) {
    console.log(`Getting line items for Request ID: ${requestId}`);
    
    // FIXED: Only return items where the RequestID field explicitly matches
    const personnel = budgetData.personnel.filter(item => {
        return item.RequestID && item.RequestID.toString().trim() === requestId.toString().trim();
    });
    
    const nonPersonnel = budgetData.nonPersonnel.filter(item => {
        return item.RequestID && item.RequestID.toString().trim() === requestId.toString().trim();
    });
    
    console.log(`Request ${requestId}: Found ${personnel.length} personnel + ${nonPersonnel.length} non-personnel items`);
    
    return [...personnel, ...nonPersonnel];
}

// ISSUE 1 FIX: Corrected department filtering logic
function getFilteredData() {
    const filters = {
        fund: document.getElementById('fundFilter').value,
        department: document.getElementById('departmentFilter').value,
        division: document.getElementById('divisionFilter').value,
        program: document.getElementById('programFilter').value,
        requestType: document.getElementById('requestTypeFilter').value,
        status: document.getElementById('statusFilter').value
    };

    console.log('Applying filters:', filters);

    return budgetData.requestSummary.filter(request => {
        // Get the Request ID for this request
        const requestId = getRequestId(request);
        if (!requestId) return false;
        
        // FIXED: Get related personnel and non-personnel data FOR THIS SPECIFIC REQUEST ONLY
        const lineItems = getLineItemsForRequest(requestId);
        
        // IMPORTANT: If no line items found, exclude the request
        if (lineItems.length === 0) return false;

        // Check filters against line items that belong to THIS request
        if (filters.fund !== 'all') {
            const hasMatchingFund = lineItems.some(item => 
                item.Fund && item.Fund.toString() === filters.fund
            );
            if (!hasMatchingFund) return false;
        }

        if (filters.department !== 'all') {
            // FIXED: Check both Department and Cost Center fields properly
            const hasMatchingDept = lineItems.some(item => {
                const dept = item.Department || item['Cost Center'] || '';
                return dept.toString() === filters.department;
            });
            if (!hasMatchingDept) return false;
        }

        if (filters.division !== 'all') {
            const hasMatchingDiv = lineItems.some(item => 
                item.Division && item.Division.toString() === filters.division
            );
            if (!hasMatchingDiv) return false;
        }

        if (filters.program !== 'all') {
            const hasMatchingProgram = lineItems.some(item => 
                item.Program && item.Program.toString() === filters.program
            );
            if (!hasMatchingProgram) return false;
        }

        // Check request-level filters
        if (filters.requestType !== 'all') {
            const hasMatchingType = Object.keys(request).some(key => {
                const lowerKey = key.toLowerCase();
                return lowerKey.includes('type') &&
                       request[key] && request[key].toString() === filters.requestType;
            });
            if (!hasMatchingType) return false;
        }

        if (filters.status !== 'all') {
            const hasMatchingStatus = Object.keys(request).some(key => {
                const lowerKey = key.toLowerCase();
                return lowerKey.includes('status') &&
                       request[key] && request[key].toString() === filters.status;
            });
            if (!hasMatchingStatus) return false;
        }

        return true;
    });
}


function getRequestId(request) {
    // Look for Request ID in various field names
    const possibleFields = Object.keys(request).filter(key => {
        const lowerKey = key.toLowerCase();
        return lowerKey.includes('request') && lowerKey.includes('id');
    });
    
    for (const field of possibleFields) {
        if (request[field]) return request[field];
    }
    
    // Fallback: look for any field with 'id'
    const idFields = Object.keys(request).filter(key => 
        key.toLowerCase().includes('id')
    );
    
    for (const field of idFields) {
        if (request[field]) return request[field];
    }
    
    return null;
}

// Pick exactly one ongoing field and one onetime field from a record, avoiding
// double-counting when both "Requested" and "Approved" variants exist on the same row.
// Prefers a "Requested" column; otherwise picks any non-"Approved" match; falls back
// to the only remaining match. mustIncludeCost=true restricts to columns containing "cost".
function pickAmountFields(record, mustIncludeCost) {
    const keys = Object.keys(record);
    const isOnetime = (k) => {
        const lk = k.toLowerCase();
        if (mustIncludeCost && !lk.includes('cost')) return false;
        return lk.includes('onetime') || lk.includes('one-time');
    };
    const isOngoing = (k) => {
        const lk = k.toLowerCase();
        if (mustIncludeCost && !lk.includes('cost')) return false;
        // Exclude onetime to avoid a column matching both buckets
        if (lk.includes('onetime') || lk.includes('one-time')) return false;
        return lk.includes('ongoing');
    };
    const score = (k) => {
        const lk = k.toLowerCase();
        if (lk.includes('request')) return 2; // prefer "Requested ..."
        if (lk.includes('approve')) return 0; // deprioritize "Approved ..." to avoid using it as a duplicate
        return 1;
    };
    const pickBest = (predicate) => {
        const matches = keys.filter(predicate);
        if (matches.length === 0) return null;
        return matches.reduce((best, k) => (score(k) > score(best) ? k : best));
    };

    const ongoingKey = pickBest(isOngoing);
    const onetimeKey = pickBest(isOnetime);
    return {
        ongoing: ongoingKey ? (parseFloat(record[ongoingKey]) || 0) : 0,
        onetime: onetimeKey ? (parseFloat(record[onetimeKey]) || 0) : 0
    };
}

function getRequestAmount(request) {
    const { ongoing, onetime } = pickAmountFields(request, false);
    return { ongoing, onetime, total: ongoing + onetime };
}

// NonPersonnel line items carry an AcctType column that can be 'Expense' or 'Revenue'.
// Revenue rows are the request's funding OFFSET (e.g. Tourist Development Tax, Water
// Revenue, Ambulance Fees), not a cost. Summing them into cost totals double-counts the
// request — expense PLUS the revenue that offsets it — which is why department/program
// rollups were showing roughly 2x the true ask.
function isRevenueLineItem(item) {
    const t = item && (item.AcctType ?? item['Acct Type'] ?? item.AccountType);
    return t != null && t.toString().trim().toLowerCase() === 'revenue';
}

// NEW: Get the actual cost from a single line item (Personnel or NonPersonnel)
function getLineItemAmount(item) {
    // Revenue/offset rows are not costs — exclude them from every total.
    if (isRevenueLineItem(item)) return { ongoing: 0, onetime: 0, total: 0 };
    const { ongoing, onetime } = pickAmountFields(item, true);
    return { ongoing, onetime, total: ongoing + onetime };
}

function updateStats() {
    filteredData = getFilteredData();
    
    console.log(`Showing ${filteredData.length} filtered requests`);
    
    const totalRequests = filteredData.length;
    let totalOngoing = 0;
    let totalOnetime = 0;
    
    // Calculate quartile distribution
    const quartileStats = {
        'Most Aligned': 0,
        'More Aligned': 0,
        'Less Aligned': 0,
        'Least Aligned': 0
    };
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        
        // Add quartile amounts
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartileStats.hasOwnProperty(quartile)) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartileStats[quartile] += lineItemAmount.total;
            }
        });
    });
    
    const totalAmount = totalOngoing + totalOnetime;

    const statsCards = document.getElementById('statsCards');
    statsCards.innerHTML = `
        <div class="stat-card">
            <h3>${totalRequests}</h3>
            <p>Total Requests</p>
        </div>
        <div class="stat-card">
            <h3>$${formatCurrency(totalOngoing)}</h3>
            <p>Ongoing Requests</p>
        </div>
        <div class="stat-card">
            <h3>$${formatCurrency(totalOnetime)}</h3>
            <p>One-time Requests</p>
        </div>
        <div class="stat-card">
            <h3>$${formatCurrency(totalAmount)}</h3>
            <p>Total Amount</p>
        </div>
        <div class="stat-card quartile-most">
            <h3>$${formatCurrency(quartileStats['Most Aligned'])}</h3>
            <p>Most Aligned</p>
        </div>
        <div class="stat-card quartile-more">
            <h3>$${formatCurrency(quartileStats['More Aligned'])}</h3>
            <p>More Aligned</p>
        </div>
        <div class="stat-card quartile-less">
            <h3>$${formatCurrency(quartileStats['Less Aligned'])}</h3>
            <p>Less Aligned</p>
        </div>
        <div class="stat-card quartile-least">
            <h3>$${formatCurrency(quartileStats['Least Aligned'])}</h3>
            <p>Least Aligned</p>
        </div>
    `;
    
    // Update UI navigation after data loads
    const uniqueDepts = new Set();
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department');
            if (dept) uniqueDepts.add(dept);
        });
    });
    
    if (typeof updateUIAfterDataLoad === 'function') {
        updateUIAfterDataLoad(totalRequests, totalAmount, uniqueDepts.size);
    }
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('en-US').format(amount);
}

function showMessage(message, type) {
    const className = type === 'error' ? 'error-message' : 
                     type === 'success' ? 'success-message' : 
                     'loading';
    
    uploadStatus.innerHTML = `<div class="${className}">${message}</div>`;
    
    if (type === 'success') {
        setTimeout(() => {
            uploadStatus.innerHTML = '';
        }, 3000);
    }
}

// Generate report
generateBtn.addEventListener('click', generateReport);

function generateReport() {
    console.log('\n=== Generating Report ===');
    
    if (filteredData.length === 0) {
        showMessage('No data matches the current filters', 'error');
        return;
    }

    progressBar.style.display = 'block';
    let progress = 0;

    const progressInterval = setInterval(() => {
        progress += 10;
        progressFill.style.width = progress + '%';
        
        if (progress >= 100) {
            clearInterval(progressInterval);
            setTimeout(() => {
                progressBar.style.display = 'none';
                displayReport();
            }, 500);
        }
    }, 100);
}

// Download functionality
document.addEventListener('DOMContentLoaded', function() {
    const downloadBtn = document.getElementById('downloadBtn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', downloadReport);
    }
});

// Also add event listener when report is generated
function displayReport() {
    console.log('Displaying reports...');
    
    const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    document.getElementById('reportDate').textContent = `Generated on ${reportDate}`;
    document.getElementById('analyticalReportDate').textContent = `Generated on ${reportDate}`;

    const totalAmount = filteredData.reduce((sum, request) => {
        const amounts = getRequestAmount(request);
        return sum + amounts.total;
    }, 0);

    // ===== GENERATE STANDARD REPORT (WITHOUT ANALYSIS) =====
    let standardHtml = `
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #333; margin-bottom: 10px;">Priority Based Budgeting Report</h1>
            <p style="color: #666; font-size: 1.1rem;">Budget Request Analysis</p>
            <p style="color: #888;">Generated on ${reportDate}</p>
        </div>

        <div class="section-header">Executive Summary</div>
        <p>This report analyzes ${filteredData.length} budget requests totaling ${formatCurrency(totalAmount)} in requested funding. The requests span multiple departments and programs, with varying levels of alignment to organizational priorities.</p>
    `;

    standardHtml += generateFilterSummary();
    standardHtml += generateActualTableOfContents();
    standardHtml += generateDepartmentSummary();
    standardHtml += generateProgramSummary();
    standardHtml += generateQuartileAnalysis();
    standardHtml += generateCharts();
    standardHtml += generateRequestSummaryTable();
    standardHtml += generateDetailedRequestReportStandard(); // Standard version without analysis

    reportContent.innerHTML = standardHtml;
    reportSection.style.display = 'block';

    // ===== GENERATE ANALYTICAL REPORT (WITH SCORING AND RECOMMENDATIONS) =====
    let analyticalHtml = `
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #333; margin-bottom: 10px;">PBB Analysis & Recommendations</h1>
            <p style="color: #666; font-size: 1.1rem;">Textbook PBB Framework - Advisory Analysis Only</p>
            <p style="color: #888;">Generated on ${reportDate}</p>
            <p style="background: #fff3cd; border: 2px solid #ffc107; padding: 10px; border-radius: 8px; margin: 15px auto; max-width: 800px; font-size: 0.95rem; color: #856404;">
                <strong>⚠️ Advisory Report:</strong> This analysis represents what a textbook Priority Based Budgeting framework would suggest. 
                These are recommendations to inform decision-making, not actual funding decisions. Final decisions rest with leadership and governing bodies.
            </p>
        </div>

        <div class="section-header">Analysis Overview</div>
        <p>This analytical report provides Priority Based Budgeting (PBB) framework scoring and advisory recommendations for ${filteredData.length} budget requests totaling <strong class="amount">$${formatCurrency(totalAmount)}</strong>. Each request is evaluated across six criteria following standard PBB methodology. <strong>These are suggested considerations, not binding decisions.</strong></p>
    `;

    analyticalHtml += generateAnalyticalSummary();
    analyticalHtml += generateAnalyticalTableOfContents();
    analyticalHtml += generateDetailedRequestReportAnalytical(); // Analytical version with full scoring

    document.getElementById('analyticalReportContent').innerHTML = analyticalHtml;
    document.getElementById('analyticalReportSection').style.display = 'block';

    // Add download event listeners
    const downloadWordBtn = document.getElementById('downloadWordBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');
    const downloadAnalyticalWordBtn = document.getElementById('downloadAnalyticalWordBtn');
    const downloadAnalyticalPdfBtn = document.getElementById('downloadAnalyticalPdfBtn');

    if (downloadWordBtn) {
        downloadWordBtn.removeEventListener('click', downloadWordReport);
        downloadWordBtn.addEventListener('click', downloadWordReport);
    }

    if (downloadPdfBtn) {
        downloadPdfBtn.removeEventListener('click', downloadPdfReport);
        downloadPdfBtn.addEventListener('click', downloadPdfReport);
    }

    if (downloadAnalyticalWordBtn) {
        downloadAnalyticalWordBtn.removeEventListener('click', downloadAnalyticalWordReport);
        downloadAnalyticalWordBtn.addEventListener('click', downloadAnalyticalWordReport);
    }

    if (downloadAnalyticalPdfBtn) {
        downloadAnalyticalPdfBtn.removeEventListener('click', downloadAnalyticalPdfReport);
        downloadAnalyticalPdfBtn.addEventListener('click', downloadAnalyticalPdfReport);
    }
    
    // Wire up Excel Export button
    const exportPBBExcelBtn = document.getElementById('exportPBBExcelBtn');
    if (exportPBBExcelBtn) {
        exportPBBExcelBtn.removeEventListener('click', exportPBBAnalysisToExcel);
        exportPBBExcelBtn.addEventListener('click', exportPBBAnalysisToExcel);
    }

    // Render charts after HTML is added to DOM
    setTimeout(renderCharts, 100);
    
    // Enable navigation and update dashboard
    if (typeof enableReportNav === 'function') enableReportNav();
    if (typeof updateDecisionDashboard === 'function') updateDecisionDashboard();
}




function generateFilterSummary() {
    // Get current filter values
    const filters = {
        fund: document.getElementById('fundFilter').value,
        department: document.getElementById('departmentFilter').value,
        division: document.getElementById('divisionFilter').value,
        program: document.getElementById('programFilter').value,
        requestType: document.getElementById('requestTypeFilter').value,
        status: document.getElementById('statusFilter').value
    };

    // Calculate quartile distribution
    const quartileStats = {
        'Most Aligned': 0,
        'More Aligned': 0,
        'Less Aligned': 0,
        'Least Aligned': 0
    };
    
    let totalOngoing = 0;
    let totalOnetime = 0;
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartileStats.hasOwnProperty(quartile)) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartileStats[quartile] += lineItemAmount.total;
            }
        });
    });

    let html = `
       <div class="toc-section" id="toc-section-filters">
           <div class="toc-section-header" onclick="toggleTOCSection('toc-section-filters')">
               <h2>1. Report Filters & Summary</h2>
               <span class="toc-toggle-icon">▼</span>
           </div>
           <div class="toc-section-content">
        <div class="request-card">
            <div class="request-header">
                <div class="request-title">Applied Filters</div>
            </div>
            <div class="request-details">
                <div class="detail-grid">
                    <div class="detail-item">
                        <div class="detail-label">Fund</div>
                        <div class="detail-value">${filters.fund}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Department</div>
                        <div class="detail-value">${filters.department}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Division</div>
                        <div class="detail-value">${filters.division}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Program</div>
                        <div class="detail-value">${filters.program}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Request Type</div>
                        <div class="detail-value">${filters.requestType}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Status</div>
                        <div class="detail-value">${filters.status}</div>
                    </div>
                </div>
            </div>
        </div>

        <div class="request-card">
            <div class="request-header">
                <div class="request-title">Financial Summary</div>
            </div>
            <div class="request-details">
                <div class="detail-grid">
                    <div class="detail-item">
                        <div class="detail-label">Total Requests</div>
                        <div class="detail-value">${filteredData.length}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Ongoing Requests</div>
                        <div class="detail-value amount">$${formatCurrency(totalOngoing)}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">One-time Requests</div>
                        <div class="detail-value amount">$${formatCurrency(totalOnetime)}</div>
                    </div>
                    <div class="detail-item">
                        <div class="detail-label">Total Amount</div>
                        <div class="detail-value amount">$${formatCurrency(totalOngoing + totalOnetime)}</div>
                    </div>
                </div>
                
                <div style="margin-top: 20px;">
                    <h4 style="color: #667eea; margin-bottom: 10px;">Quartile Distribution</h4>
                    <div class="detail-grid">
                        <div class="detail-item">
                            <div class="detail-label">Most Aligned</div>
                            <div class="detail-value amount">$${formatCurrency(quartileStats['Most Aligned'])}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">More Aligned</div>
                            <div class="detail-value amount">$${formatCurrency(quartileStats['More Aligned'])}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Less Aligned</div>
                            <div class="detail-value amount">$${formatCurrency(quartileStats['Less Aligned'])}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Least Aligned</div>
                            <div class="detail-value amount">$${formatCurrency(quartileStats['Least Aligned'])}</div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
            </div>
        </div>
    `;

    return html;
}

function generateActualTableOfContents() {
    let html = `
        <div class="section-header">Table of Contents</div>
        <div class="request-card">
            <div class="request-details">
                <ol style="line-height: 2; font-size: 1.1rem;">
                    <li><a href="#report-filters" style="color: #667eea; text-decoration: none;">Report Filters & Summary</a></li>
                    <li><a href="#request-summary-table" style="color: #667eea; text-decoration: none;">Request Summary Table</a></li>
                    <li><a href="#department-summary" style="color: #667eea; text-decoration: none;">Department Summary</a></li>
                    <li><a href="#quartile-analysis" style="color: #667eea; text-decoration: none;">Program Alignment Analysis</a></li>
                    <li><a href="#individual-requests" style="color: #667eea; text-decoration: none;">Individual Budget Requests</a>
                        <ol style="margin-top: 10px; font-size: 1rem;">
    `;

    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        html += `<li><a href="#request-${requestId}" style="color: #667eea; text-decoration: none;">Request ${requestId}: ${description || 'N/A'}</a></li>`;
    });

    html += `
                        </ol>
                    </li>
                    <li><a href="#visual-analysis" style="color: #667eea; text-decoration: none;">Visual Analysis</a></li>
                </ol>
            </div>
        </div>
    `;

    return html;
}

function generateRequestSummaryTable() {
    console.log('Generating request summary table...');
    
    let html = `
        <div class="section-header" id="request-summary-table">Request Summary Table</div>
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
            <thead>
                <tr style="background: #f8f9ff;">
                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #667eea;">Request ID</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #667eea;">Description</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #667eea;">Department</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #667eea;">Primary Program</th>
                    <th style="padding: 12px; text-align: left; border-bottom: 2px solid #667eea;">Quartile</th>
                    <th style="padding: 12px; text-align: right; border-bottom: 2px solid #667eea;">Total Amount</th>
                </tr>
            </thead>
            <tbody>
    `;

    filteredData.forEach((request, idx) => {
        console.log(`Request summary row ${idx}:`, request);
        
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        
        console.log(`Request ${requestId}: ${lineItems.length} line items`);
        
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryProgram = getPrimaryValue(lineItems, 'program') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);

        console.log(`Request ${requestId}: Dept=${primaryDept}, Program=${primaryProgram}, Amount=${amounts.total}`);

        const quartileBadge = primaryQuartile !== 'N/A' ? 
            `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}">${primaryQuartile}</span>` : 
            'N/A';

        html += `
            <tr style="border-bottom: 1px solid #e0e0e0;">
                <td style="padding: 10px; font-weight: 600;"><a href="#request-${requestId}" style="color: #667eea; text-decoration: none;">${requestId || 'N/A'}</a></td>
                <td style="padding: 10px;">${description || 'N/A'}</td>
                <td style="padding: 10px;">${primaryDept}</td>
                <td style="padding: 10px;">${primaryProgram}</td>
                <td style="padding: 10px;">${quartileBadge}</td>
                <td style="padding: 10px; text-align: right; font-weight: 600; color: #28a745;">$${formatCurrency(amounts.total)}</td>
            </tr>
        `;
    });

    html += '</tbody></table>';
    return html;
}

function getRequestDescription(request) {
    // Look for description field
    const possibleFields = Object.keys(request).filter(key => {
        const lowerKey = key.toLowerCase();
        return lowerKey.includes('description') || lowerKey.includes('desc');
    });
    
    for (const field of possibleFields) {
        if (request[field]) return request[field];
    }
    
    return 'N/A';
}

// Normalize any quartile representation (1/2/3/4, Q1-Q4, "Most Aligned", etc.)
// to one of the four canonical bucket names, or null if unrecognized.
function normalizeQuartile(q) {
    if (q === null || q === undefined || q === '') return null;
    const s = q.toString().trim();
    const lower = s.toLowerCase();
    if (s === '1' || s === 'Q1' || lower === 'most aligned') return 'Most Aligned';
    if (s === '2' || s === 'Q2' || lower === 'more aligned') return 'More Aligned';
    if (s === '3' || s === 'Q3' || lower === 'less aligned') return 'Less Aligned';
    if (s === '4' || s === 'Q4' || lower === 'least aligned') return 'Least Aligned';
    return null;
}

// Resolve the quartile for a single line item. Tries the line item's own Quartile
// column first; if blank, falls back to looking up the program in the uploaded
// Current Budget (Program Inventory) by Program (+ Department when both have one).
// Build the program-attributes map from a Details-sheet row dump. Each program has
// many account-level rows but the program-level attributes are constant per program,
// so we record the first non-empty value seen per (User Group, Program) key.
function buildProgramAttributesMap(detailsRows) {
    if (!Array.isArray(detailsRows) || detailsRows.length === 0) return;
    const norm = v => (v == null ? '' : v.toString().trim().toUpperCase());
    for (const row of detailsRows) {
        const program = norm(row['Program']);
        if (!program) continue;
        const dept = norm(row['User Group(prgs)'] || row['User Group(accts)'] || row['Department'] || row['User Group']);
        const attrs = {
            mandate: classifyMandate(row['Mandate']),
            costRecovery: classifyCostRecovery(row['Cost Recovery']),
            finalScore: typeof row['Final Score'] === 'number' ? row['Final Score'] : parseFloat(row['Final Score']) || null,
            rawMandate: row['Mandate'] || null,
            rawCostRecovery: row['Cost Recovery'] || null
        };
        const keys = [`${dept}::${program}`, `*::${program}`];
        for (const k of keys) {
            const existing = programAttributesMap[k];
            if (!existing) {
                programAttributesMap[k] = attrs;
            } else {
                // Upgrade in place if the new row has stronger / non-empty values
                if (!existing.mandate && attrs.mandate) existing.mandate = attrs.mandate;
                if (existing.costRecovery == null && attrs.costRecovery != null) existing.costRecovery = attrs.costRecovery;
                if (!existing.finalScore && attrs.finalScore) existing.finalScore = attrs.finalScore;
            }
        }
    }
}

function classifyMandate(val) {
    if (!val) return null;
    const s = val.toString().toLowerCase();
    if (s.includes('state') || s.includes('federal')) return 'state_federal';
    if (s.includes('self') || s.includes('ordinance') || s.includes('charter') || s.includes('commission')) return 'self';
    if (s.includes('no mandate') || /^no\b/.test(s) || s.includes('(0)')) return 'none';
    return null;
}

function classifyCostRecovery(val) {
    if (val === '' || val == null) return null;
    const s = val.toString().toLowerCase();
    if (s.startsWith('yes') || /\(4\)|\(3\)|\(2\)/.test(s)) return true;
    if (s.startsWith('no') || s.includes('(0)')) return false;
    return null;
}

// Aggregate the strongest program-level signals across all line items for a request.
function getProgramAttributesForLineItems(lineItems) {
    if (!lineItems || lineItems.length === 0) return null;
    if (!programAttributesMap || Object.keys(programAttributesMap).length === 0) return null;

    const mandateRank = { state_federal: 3, self: 2, none: 1 };
    const inverseMandate = { 3: 'state_federal', 2: 'self', 1: 'none' };
    let bestMandate = 0;
    let anyCostRecovery = false;
    let sawCostRecovery = false;
    let bestFinalScore = null;
    let matched = false;

    for (const item of lineItems) {
        const program = (item.Program || '').toString().trim().toUpperCase();
        if (!program) continue;
        const dept = (item.Department || item['Cost Center'] || item['User Group'] || '').toString().trim().toUpperCase();
        const attrs = programAttributesMap[`${dept}::${program}`] || programAttributesMap[`*::${program}`];
        if (!attrs) continue;
        matched = true;
        if (attrs.mandate && mandateRank[attrs.mandate] > bestMandate) {
            bestMandate = mandateRank[attrs.mandate];
        }
        if (attrs.costRecovery != null) {
            sawCostRecovery = true;
            if (attrs.costRecovery) anyCostRecovery = true;
        }
        if (attrs.finalScore != null && (bestFinalScore == null || attrs.finalScore > bestFinalScore)) {
            bestFinalScore = attrs.finalScore;
        }
    }

    if (!matched) return null;
    return {
        mandate: bestMandate ? inverseMandate[bestMandate] : null,
        costRecovery: sawCostRecovery ? anyCostRecovery : null,
        finalScore: bestFinalScore
    };
}

function getQuartileForLineItem(item) {
    const direct = normalizeQuartile(item.Quartile);
    if (direct) return direct;

    if (currentBudgetData.length === 0) return null;

    const program = (item.Program || '').toString().trim().toUpperCase();
    if (!program) return null;
    const dept = (item.Department || item['Cost Center'] || item['User Group'] || '')
        .toString().trim().toUpperCase();

    const match = currentBudgetData.find(p => {
        const pName = (p.Program || '').toString().trim().toUpperCase();
        if (pName !== program) return false;
        const pDept = (p['User Group'] || p.Department || '').toString().trim().toUpperCase();
        // If both sides have a department, require a match; otherwise accept the program-only match.
        if (dept && pDept && dept !== pDept) return false;
        return true;
    });
    if (!match) return null;
    return normalizeQuartile(match.Quartile);
}

// HELPER FUNCTION: Get primary value with improved logic
function getPrimaryValue(lineItems, fieldType) {
    // Look for the specific field in line items
    for (const item of lineItems) {
        if (fieldType === 'department') {
            // Check both Department and Cost Center fields
            if (item.Department) return item.Department;
            if (item['Cost Center']) return item['Cost Center'];
        } else if (fieldType === 'program') {
            if (item.Program) return item.Program;
        } else if (fieldType === 'quartile') {
            const q = getQuartileForLineItem(item);
            if (q) return q;
        } else if (fieldType === 'fund') {
            if (item.Fund) return item.Fund;
        } else if (fieldType === 'division') {
            if (item.Division) return item.Division;
        }
    }
    return null;
}

// ===== MATCH PROGRAM WITH CURRENT BUDGET =====
function getCurrentBudgetForProgram(department, programName) {
    if (currentBudgetData.length === 0) {
        return null; // No current budget data loaded
    }
    
    console.log(`Looking for match: Dept="${department}", Program="${programName}"`);
    
    // Try to find exact match by User Group (Department) and Program Name
    const match = currentBudgetData.find(prog => {
        const userGroup = (prog['User Group'] || '').toString().trim().toUpperCase();
        const progName = (prog['Program'] || '').toString().trim().toUpperCase();
        const deptUpper = (department || '').toString().trim().toUpperCase();
        const progNameUpper = (programName || '').toString().trim().toUpperCase();
        
        return userGroup === deptUpper && progName === progNameUpper;
    });
    
    if (match) {
        console.log(`✅ Match found! Current budget: $${match['Total Program Cost']}`);
        return {
            totalCost: match['Total Program Cost'] || 0,
            personnel: match['Personnel'] || 0,
            nonPersonnel: match['NonPersonnel'] || 0,
            revenue: match['Revenue'] || 0,
            fte: match['FTE'] || 0,
            description: match['Description'] || ''
        };
    }
    
    console.log(`❌ No match found for "${department}" - "${programName}"`);
    return null; // No match found
}

// Helper function to convert Markdown-style formatting to HTML
function markdownToHtml(text) {
    if (!text) return text;
    
    // Convert **text** to <strong>text</strong>
    text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    
    // Convert *text* to <em>text</em>
    text = text.replace(/\*(.+?)\*/g, '<em>$1</em>');
    
    // Convert newlines to <br> tags
    text = text.replace(/\n/g, '<br>');
    
    return text;
}

// ===== ENHANCED PBB SCORING ENGINE WITH EXPLICIT REASONING =====

// Helper function to get best quartile - returns full name and handles BOTH formats
function getBestQuartile(quartiles) {
    // Normalize all quartiles to text format for comparison
    const normalizedQuartiles = quartiles.map(q => {
        if (!q) return null;
        const qStr = q.toString().trim();
        if (qStr === '1' || qStr === 'Q1' || qStr === 'Most Aligned') return 'Most Aligned';
        if (qStr === '2' || qStr === 'Q2' || qStr === 'More Aligned') return 'More Aligned';
        if (qStr === '3' || qStr === 'Q3' || qStr === 'Less Aligned') return 'Less Aligned';
        if (qStr === '4' || qStr === 'Q4' || qStr === 'Least Aligned') return 'Least Aligned';
        return null;
    }).filter(q => q !== null);
    
    // Return the best (most aligned) quartile found
    if (normalizedQuartiles.includes('Most Aligned')) return 'Most Aligned';
    if (normalizedQuartiles.includes('More Aligned')) return 'More Aligned';
    if (normalizedQuartiles.includes('Less Aligned')) return 'Less Aligned';
    if (normalizedQuartiles.includes('Least Aligned')) return 'Least Aligned';
    return null;
}


function getQuartileScore(quartile) {
    if (!quartile) return { score: 0, reason: "No quartile alignment data found in line items" };
    
    // Convert quartile to string and normalize
    const quartileStr = quartile.toString().trim();
    
    // Handle BOTH number format (1,2,3,4) AND text format ("Most Aligned", etc.)
    if (quartileStr === 'Most Aligned' || quartileStr === 'Q1' || quartileStr === '1' || quartile === 1) {
        return { score: 2, reason: `Program quartile is "Most Aligned" (Q1) - highest priority alignment with organizational strategic goals and community priorities` };
    }
    if (quartileStr === 'More Aligned' || quartileStr === 'Q2' || quartileStr === '2' || quartile === 2) {
        return { score: 2, reason: `Program quartile is "More Aligned" (Q2) - strong alignment with organizational strategic goals and community priorities` };
    }
    if (quartileStr === 'Less Aligned' || quartileStr === 'Q3' || quartileStr === '3' || quartile === 3) {
        return { score: 1, reason: `Program quartile is "Less Aligned" (Q3) - moderate alignment with organizational strategic goals` };
    }
    if (quartileStr === 'Least Aligned' || quartileStr === 'Q4' || quartileStr === '4' || quartile === 4) {
        return { score: 0, reason: `Program quartile is "Least Aligned" (Q4) - lower priority alignment with current strategic goals` };
    }
    
    return { score: 0, reason: `Quartile "${quartile}" not recognized - unable to score alignment` };
}




function getOutcomeScore(qa, qaText) {
    const hasMetrics = /kpi|target|baseline|metric|goal|measur/i.test(qaText);
    const hasData = /data|trend|statistics|baseline/i.test(qaText);
    
    if (hasMetrics && hasData) {
        return { score: 2, reason: "Request includes specific KPIs/metrics AND baseline data or trends showing measurable outcomes" };
    }
    if (hasMetrics) {
        return { score: 1, reason: "Request mentions performance targets or metrics, but lacks supporting baseline data or outcome trends" };
    }
    if (hasData && qa.length > 0 && !/n\/a|unknown|none/i.test(qaText)) {
        return { score: 1, reason: "Request includes some data or information, but lacks specific measurable performance targets" };
    }
    return { score: 0, reason: "No measurable outcomes, KPIs, targets, or performance data provided in request documentation" };
}

function getFundingScore(qa, qaText, progAttrs, fundProfile) {
    // Authoritative: the structured Fund column already names a non-General-Fund source.
    if (fundProfile && fundProfile.hasFundData && fundProfile.fundingClass === 'NonGF') {
        const names = fundProfile.funds.join(', ');
        return { score: fundProfile.gf > 0 ? 1 : 2, reason: `Funded by non-General Fund source(s) per the Fund column: ${names}` };
    }
    const hasGrant = /grant|outside funding.*yes/i.test(qaText);
    const hasFee = /fee|cost recovery|charge|revenue/i.test(qaText);
    const hasPartner = /partner|partnership|contribution|match/i.test(qaText);

    if ((hasGrant || hasFee || hasPartner) && qaText.match(/grant|fee|partner/gi)?.length >= 2) {
        return { score: 2, reason: "Request identifies MULTIPLE non-General Fund sources (grants, fees, cost recovery, or partnership funding)" };
    }
    if (hasGrant) {
        return { score: 1, reason: "Request mentions grant funding or outside funding sources, reducing General Fund dependency" };
    }
    if (hasFee || hasPartner) {
        return { score: 1, reason: "Request includes cost recovery mechanisms (fees/charges) or partnership contributions" };
    }
    if (/potential|exploring|seeking/i.test(qaText) && /grant|partner|fee/i.test(qaText)) {
        return { score: 1, reason: "Request mentions exploring or seeking non-General Fund sources, though not yet secured" };
    }
    // Fallback to program-level structured data when narrative is silent
    if (progAttrs && progAttrs.costRecovery === true) {
        return { score: 1, reason: "Program inventory marks this program as having cost recovery — non-General Fund offset is established at the program level" };
    }
    return { score: 0, reason: "No non-General Fund sources identified - request is 100% dependent on General Fund appropriation" };
}

function getMandateScore(qa, qaText, progAttrs) {
    const hasMandate = /board motion|consent decree|doj|mandate|statute|ordinance|charter/i.test(qaText);
    const hasCompliance = /audit|liability|compliance|risk|safety|violation|penalty/i.test(qaText);

    if (hasMandate && hasCompliance) {
        return { score: 2, reason: "Request cites specific legal/regulatory mandate (board motion, statute, consent decree) AND identifies compliance risks or penalties" };
    }
    if (hasMandate) {
        return { score: 1, reason: "Request references legal or regulatory mandate, board motion, or statutory requirement" };
    }
    if (hasCompliance) {
        return { score: 1, reason: "Request addresses compliance obligations, audit findings, liability mitigation, or safety risks" };
    }
    // Fallback to program-level structured data when narrative is silent
    if (progAttrs && progAttrs.mandate === 'state_federal') {
        return { score: 2, reason: "Program inventory classifies this program as a State or Federal Mandate — legal/regulatory obligation established at the program level" };
    }
    if (progAttrs && progAttrs.mandate === 'self') {
        return { score: 1, reason: "Program inventory classifies this program as a Self Mandate (commission, ordinance, or charter) — local compliance obligation at the program level" };
    }
    return { score: 0, reason: "No legal mandates, compliance obligations, or significant regulatory risks identified in request" };
}

function getEfficiencyScore(qa, qaText) {
    const hasROI = /roi|return on investment|payback|cost avoidance|cost savings/i.test(qaText);
    const hasEfficiency = /productivity|efficiency|streamline|reduce cost|automate/i.test(qaText);
    const hasQuantification = /\$\d+|save.*\d+|\d+%|\d+ hours|\d+ fte/i.test(qaText);
    
    if ((hasROI || hasEfficiency) && hasQuantification) {
        return { score: 2, reason: "Request demonstrates efficiency gains or ROI with QUANTIFIED savings, cost avoidance, or productivity improvements (includes dollar amounts, percentages, or time savings)" };
    }
    if (hasROI || (hasEfficiency && hasQuantification)) {
        return { score: 1, reason: "Request mentions efficiency improvements, cost savings, or ROI, with some quantification or specific metrics" };
    }
    if (hasEfficiency) {
        return { score: 1, reason: "Request describes efficiency improvements or process streamlining, but lacks quantified ROI or savings calculations" };
    }
    return { score: 0, reason: "No efficiency improvements, cost savings, ROI, or productivity gains identified in the request" };
}

function getAccessScore(qa, qaText) {
    const hasEquity = /equity|underserved|priority population|disparit|vulnerable|disadvantaged/i.test(qaText);
    const hasAccess = /access|barrier|inclusive|reach|serve/i.test(qaText);
    const hasPopData = /\d+%|portion|community|residents|population|demographic/i.test(qaText);
    
    if ((hasEquity || hasAccess) && hasPopData) {
        return { score: 2, reason: "Request explicitly addresses access or equity issues with SPECIFIC population data (percentages, demographics, or community impact metrics)" };
    }
    if (hasEquity) {
        return { score: 1, reason: "Request mentions equity, underserved populations, or vulnerable communities, but lacks specific demographic data" };
    }
    if (hasAccess || (/community|service|outreach/i.test(qaText) && hasPopData)) {
        return { score: 1, reason: "Request addresses community access or service delivery with some population information" };
    }
    return { score: 0, reason: "No specific attention to access, equity considerations, or underserved population impacts identified" };
}

// Classify a single Fund column value as drawing on the General Fund ('GF'),
// a dedicated/enterprise/restricted source ('NonGF'), or unrecognized ('Unknown').
// This is the AUTHORITATIVE funding signal — it reads the actual fund on the line
// item rather than guessing from narrative keywords.
function classifyFundName(fund) {
    if (fund === null || fund === undefined) return null;
    const s = fund.toString().trim().toLowerCase();
    // Blank or placeholder values carry no fund signal.
    if (!s || /^(n\/?a|n\.a\.?|none|tbd|unknown|null|0|-+)$/.test(s)) return null;
    // Recognizable General Fund: draws on countywide ad valorem / the General Fund.
    if (/general fund|ad valorem|countywide general|\bgf\b|^0*1$|^0*1\b|^00?1\s*-/.test(s)) {
        return 'GF';
    }
    // Any other NAMED fund is a distinct fiscal entity with its own dedicated revenue —
    // enterprise/utility (paid from user rates), special revenue, MSTU, grant, tourism,
    // road & bridge, solid waste, etc. None of these draw on the General Fund.
    return 'NonGF';
}

// Aggregate the structured Fund column across a request's line items, weighted by
// dollar amount. Returns the dominant funding class plus the distinct fund names so
// the UI can show the real source (e.g. "Enterprise/Utility Fund") instead of "GF Only".
function getFundProfile(lineItems) {
    let gf = 0, nonGf = 0, unknown = 0;
    const funds = new Set();
    let sawAnyFund = false;
    for (const item of lineItems) {
        const cls = classifyFundName(item.Fund);
        if (cls === null) continue; // blank / placeholder fund — no signal
        sawAnyFund = true;
        funds.add(item.Fund.toString().trim());
        const amt = getLineItemAmount(item).total || 0;
        const w = amt > 0 ? amt : 1; // fall back to a per-line vote when amounts are absent
        if (cls === 'GF') gf += w;
        else if (cls === 'NonGF') nonGf += w;
        else unknown += w;
    }
    if (!sawAnyFund) {
        return { hasFundData: false, funds: [], fundingClass: null, gf: 0, nonGf: 0, unknown: 0 };
    }
    // Any non-GF dollars => the request has a non-GF offset. Only fall to GFonly when
    // the recognizable funding is entirely General Fund. Funds we can't recognize at
    // all stay Unknown rather than silently defaulting to "GF Only".
    let fundingClass;
    if (nonGf > 0) fundingClass = 'NonGF';
    else if (gf > 0) fundingClass = 'GFonly';
    else fundingClass = 'Unknown';
    return { hasFundData: true, funds: [...funds], fundingClass, gf, nonGf, unknown };
}

// Display attributes for the Funding decision factor, including the real fund name(s).
function getFundingDisplay(analysis) {
    const fp = analysis.fundProfile;
    const sub = fp && fp.funds && fp.funds.length ? fp.funds.join(', ') : '';
    if (analysis.fundingType === 'NonGF') {
        return { label: '💚 Non-GF', sub, text: '#059669', grad: 'linear-gradient(135deg, #d1fae5, #a7f3d0)', plain: '#d1fae5', border: '#10b981' };
    }
    if (analysis.fundingType === 'Unknown') {
        return { label: '⚪ Unknown', sub, text: '#475569', grad: 'linear-gradient(135deg, #f1f5f9, #e2e8f0)', plain: '#f1f5f9', border: '#94a3b8' };
    }
    return { label: '🔴 GF Only', sub, text: '#dc2626', grad: 'linear-gradient(135deg, #fee2e2, #fecaca)', plain: '#fee2e2', border: '#ef4444' };
}

function scoreRequest(request) {
    const requestId = getRequestId(request);
    const lineItems = getLineItemsForRequest(requestId);
    const qa = getRequestQA(requestId);
    const amounts = getRequestAmount(request);
    const fundProfile = getFundProfile(lineItems);
    
    const quartiles = lineItems.map(li => getPrimaryValue([li], 'quartile')).filter(q => q);
    const bestQuartile = getBestQuartile(quartiles);
    const qaText = qa.map(q => Object.values(q).join(' ')).join(' ').toLowerCase();

    // Pull program-level structured PBB attributes (Mandate / Cost Recovery / Final Score)
    // from the uploaded Programs Inventory. These act as a fallback when the request Q&A
    // is missing or boilerplate (common with imported workbooks that have no narrative).
    const progAttrs = getProgramAttributesForLineItems(lineItems);

    // Score each criterion with explicit reasoning
    const quartileAnalysis = getQuartileScore(bestQuartile);
    const outcomeAnalysis = getOutcomeScore(qa, qaText);
    const fundingAnalysis = getFundingScore(qa, qaText, progAttrs, fundProfile);
    const mandateAnalysis = getMandateScore(qa, qaText, progAttrs);
    const efficiencyAnalysis = getEfficiencyScore(qa, qaText);
    const accessAnalysis = getAccessScore(qa, qaText);

    const analysis = {
        // Scores with explicit reasons
        quartileScore: quartileAnalysis.score,
        quartileReason: quartileAnalysis.reason,

        outcomeScore: outcomeAnalysis.score,
        outcomeReason: outcomeAnalysis.reason,

        fundingScore: fundingAnalysis.score,
        fundingReason: fundingAnalysis.reason,

        mandateScore: mandateAnalysis.score,
        mandateReason: mandateAnalysis.reason,

        efficiencyScore: efficiencyAnalysis.score,
        efficiencyReason: efficiencyAnalysis.reason,

        accessScore: accessAnalysis.score,
        accessReason: accessAnalysis.reason,

        // Surface the structured fallback so downstream UI / debug can show what fed the grid
        programAttributes: progAttrs,
        fundProfile: fundProfile,

        // Legacy fields used to derive grid axes — OR in structured signals so the grid
        // routes correctly when the narrative is silent.
        bestQuartile: bestQuartile,
        hasOutsideFunding:
            /outside funding.*yes|grant|fee|partner|cost recovery/i.test(qaText) ||
            (progAttrs && progAttrs.costRecovery === true),
        isMandated:
            /board motion|consent decree|doj|mandate|statute/i.test(qaText) ||
            (progAttrs && progAttrs.mandate === 'state_federal'),
        isCompliance:
            /audit|liability|compliance|risk|safety/i.test(qaText) ||
            (progAttrs && progAttrs.mandate === 'self')
    };
    
    // Calculate total score (Access/Equity excluded when client toggle is off)
    const totalScore = quartileAnalysis.score + outcomeAnalysis.score + fundingAnalysis.score +
                      mandateAnalysis.score + efficiencyAnalysis.score +
                      (includeAccessEquity ? accessAnalysis.score : 0);

    // Weighted scoring — Quartile is MOST important. Access weight is dropped when off so
    // the percentage isn't depressed by a criterion the client opted out of.
    const weightedScore = {
        quartile: quartileAnalysis.score * 2.0,      // Double weight (4 points max)
        outcomes: outcomeAnalysis.score * 1.5,       // 50% bonus (3 points max)
        funding: fundingAnalysis.score * 1.5,        // 50% bonus (3 points max)
        mandate: mandateAnalysis.score * 1.0,        // Standard weight (2 points max)
        efficiency: efficiencyAnalysis.score * 0.75, // Reduced weight (1.5 points max)
        access: includeAccessEquity ? accessAnalysis.score * 0.75 : 0
    };

    const maxWeightedScore = includeAccessEquity ? 16.5 : 15.0;
    const totalWeightedScore = Object.values(weightedScore).reduce((a, b) => a + b, 0);
    const weightedPercentage = Math.round((totalWeightedScore / maxWeightedScore) * 100);
    
    // Determine quartile band (High = Q1/Q2/1/2/Most/More, Low = Q3/Q4/3/4/Less/Least)
    // Unknown = no quartile data on line items or in Program Inventory — must NOT silently
    // collapse to "Low" (which previously routed every unknown to REJECT via archetype 24).
    const quartileStr = bestQuartile ? bestQuartile.toString().trim() : '';
    if (!quartileStr) {
        analysis.quartileBand = 'Unknown';
    } else if (
        quartileStr === 'Q1' || quartileStr === 'Q2' ||
        quartileStr === '1' || quartileStr === '2' ||
        quartileStr === 'Most Aligned' || quartileStr === 'More Aligned'
    ) {
        analysis.quartileBand = 'High';
    } else {
        analysis.quartileBand = 'Low';
    }
    
    
    // Determine mandate level
    if (analysis.isMandated) {
        analysis.mandateLevel = 'Mandated';
    } else if (analysis.isCompliance) {
        analysis.mandateLevel = 'Compliance';
    } else {
        analysis.mandateLevel = 'None';
    }
    
    // Determine funding type. Prefer the authoritative structured Fund column on the
    // line items; only fall back to narrative keywords / program-level cost recovery
    // when no fund is supplied. When nothing identifies the source, surface 'Unknown'
    // rather than silently asserting "GF Only" (the bug this replaces: utility requests
    // paid from user rates were being labeled General Fund).
    if (fundProfile.hasFundData && fundProfile.fundingClass !== 'Unknown') {
        analysis.fundingType = fundProfile.fundingClass;
    } else if (analysis.hasOutsideFunding) {
        analysis.fundingType = 'NonGF';
    } else {
        analysis.fundingType = 'Unknown';
    }
    // Keep the narrative's non-GF flag consistent with the structured signal.
    analysis.hasOutsideFunding = analysis.hasOutsideFunding ||
        (fundProfile.hasFundData && fundProfile.fundingClass === 'NonGF');
    
    // Determine outcomes strength
    analysis.outcomesStrength = outcomeAnalysis.score >= 2 ? 'Strong' : 'Weak';
    
    analysis.totalScore = totalScore;
    analysis.weightedScore = totalWeightedScore;
    analysis.weightedPercentage = weightedPercentage;

    // Apply the decision grid
    const gridDecision = applyDecisionGrid(analysis);
    
    analysis.disposition = gridDecision.disposition;
    analysis.dispositionColor = gridDecision.color;
    analysis.verifyNow = gridDecision.verifyNow;
    analysis.strengthenWith = gridDecision.strengthenWith;
    analysis.gridKey = gridDecision.gridKey;
    analysis.archetypeNumber = gridDecision.archetypeNumber;
    analysis.keyConsideration = gridDecision.keyConsideration;
    
    // Generate enhanced narrative
    analysis.narrative = generateEnhancedNarrative(request, lineItems, qa, analysis);
    
    return analysis;
}

// ===== REVISED DECISION GRID BASED ON GF FUNDING PHILOSOPHY =====
function applyDecisionGrid(analysis) {
    const { quartileBand, mandateLevel, fundingType, outcomesStrength } = analysis;

    // Create lookup key
    const gridKey = `${quartileBand}-${mandateLevel}-${fundingType}-${outcomesStrength}`;

    // Without quartile data, the entire strategic-priority axis collapses and the grid
    // can't produce a defensible disposition. Surface a REVIEW state instead of forcing
    // the request through a half-blind APPROVE/MODIFY/DEFER/REJECT decision.
    if (quartileBand === 'Unknown') {
        return {
            archetypeNumber: 0,
            disposition: 'REVIEW',
            color: '#64748b',
            keyConsideration: 'Insufficient quartile data — manual review required before applying the PBB framework',
            verifyNow: ['Add a Quartile value to the line items, or upload a Program Inventory that includes a Quartile column'],
            strengthenWith: ['Provide program-level alignment data so the framework can score this request'],
            gridKey: gridKey
        };
    }
    
    // Without a determinable funding source, the GFonly/NonGF axis collapses. Don't
    // force the request onto either rail (which previously meant a false "GF Only").
    // Surface a REVIEW state asking for the fund to be identified.
    if (fundingType === 'Unknown') {
        return {
            archetypeNumber: 0,
            disposition: 'REVIEW',
            color: '#64748b',
            keyConsideration: 'Funding source could not be determined — confirm whether this draws on the General Fund or a dedicated/enterprise (rate-, grant-, or fee-funded) source before applying a disposition.',
            verifyNow: ['Add a Fund value to the line items (e.g. General Fund, Enterprise/Utility Fund), or state the funding source in the request Q&A'],
            strengthenWith: ['Tag each line item with its fund so the framework can separate General Fund impact from rate- or grant-funded requests'],
            gridKey: gridKey
        };
    }

    // Decision grid mapping with archetype numbers matching the 24 Archetypes table
    const grid = {
        // HIGH RELEVANCE (Q1-Q2) - Strategic Priority Programs (Archetypes 1-12)
        'High-Mandated-NonGF-Strong': {
            archetypeNumber: 1,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Perfect alignment - strategic priority + mandate + no GF impact',
            verifyNow: ['Statute/board reference', 'Allowability of non-GF sources'],
            strengthenWith: ['Final KPI list', 'Compliance milestones', 'Data source & cadence']
        },
        'High-Mandated-GFonly-Strong': {
            archetypeNumber: 2,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Mission-critical with legal mandate backing strategic goals',
            verifyNow: ['Confirm mandate scope & minimums'],
            strengthenWith: ['Cost offsets (phase-down plan, reallocation)', 'Sunset/true-up triggers']
        },
        'High-Mandated-NonGF-Weak': {
            archetypeNumber: 3,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Mandate + external funding covers weak evidence',
            verifyNow: ['That mandate truly requires this spend'],
            strengthenWith: ['Baseline→target KPIs', '90-day evaluation plan', 'Interim check-in']
        },
        'High-Mandated-GFonly-Weak': {
            archetypeNumber: 4,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Mandate requires compliance regardless of evidence gaps',
            verifyNow: ['Minimum-viable compliance level'],
            strengthenWith: ['Add fee/grant search', 'Partner MOUs', 'Phased start', 'Sunset clause']
        },
        'High-Compliance-NonGF-Strong': {
            archetypeNumber: 5,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Strategic priority with strong case and low GF risk',
            verifyNow: ['Risk register link', 'Risk reduction metric'],
            strengthenWith: ['Cost avoidance calc', 'SLA updates', 'Internal control changes']
        },
        'High-Compliance-GFonly-Strong': {
            archetypeNumber: 6,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Strong case but push for cost recovery',
            verifyNow: ['Materiality of risk', 'Alternatives'],
            strengthenWith: ['Add partial cost recovery', 'Internal reallocation', 'Pilot scope']
        },
        'High-Compliance-NonGF-Weak': {
            archetypeNumber: 7,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Strengthen outcomes/evidence first',
            verifyNow: ['That non-GF is real & timely'],
            strengthenWith: ['KPIs', '6-mo pilot with go/no-go', 'Light-weight evaluation plan']
        },
        'High-Compliance-GFonly-Weak': {
            archetypeNumber: 8,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Require evidence plan before approval',
            verifyNow: ['Is this truly critical for safety/liability?'],
            strengthenWith: ['Strengthen outcomes evidence', 'Identify cost recovery', 'Narrow scope significantly', 'Stage gates with evaluation']
        },
        'High-None-NonGF-Strong': {
            archetypeNumber: 9,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Strategic with self-sustaining funding',
            verifyNow: ['No hidden GF backfill'],
            strengthenWith: ['Pay-for-itself math', 'Fee elasticity/grant terms', 'Partner commitments']
        },
        'High-None-GFonly-Strong': {
            archetypeNumber: 10,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Good case but seek fee/grant offset opportunities first',
            verifyNow: ['Alignment with strategic plan goals', 'Expected impact on outcomes'],
            strengthenWith: ['Explore cost recovery options', 'Unit-cost reduction opportunities', 'Potential partnerships']
        },
        'High-None-NonGF-Weak': {
            archetypeNumber: 11,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Strengthen business case or run pilot to prove value',
            verifyNow: ['Outcome plausibility'],
            strengthenWith: ['KPIs & evaluation', 'Start as pilot', 'Tighten deliverables']
        },
        'High-None-GFonly-Weak': {
            archetypeNumber: 12,
            disposition: 'DEFER',
            color: '#dc3545',
            keyConsideration: 'Build evidence and outcomes data before using limited GF',
            verifyNow: ['Why should GF be used for this lower-evidence request?'],
            strengthenWith: ['Tie to priority KPIs with clear metrics', 'Find non-GF sources', 'Reduce scope or integrate with higher-priority work']
        },
        
        // LOW RELEVANCE (Q3-Q4) - Lower Strategic Priority (Archetypes 13-24)
        'Low-Mandated-NonGF-Strong': {
            archetypeNumber: 13,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Legal mandate with no GF impact - proceed with compliance',
            verifyNow: ['Minimum compliance scope'],
            strengthenWith: ['Keep GF minimal', 'Escrow/offsets', 'Time-bound sunset']
        },
        'Low-Mandated-GFonly-Strong': {
            archetypeNumber: 14,
            disposition: 'VERIFY',
            color: '#6366f1',
            keyConsideration: 'Low-priority program with 100% GF reliance — mandate is the only thing keeping this out of REJECT. Validate before funding.',
            verifyNow: [
                'What does the mandate actually require — operate the program at all, or this specific incremental spend?',
                'Who is mandating it? (statute, regulation, court order, board motion) — get a citation',
                'Is there pass-through funding from the mandating authority to comply?',
                'Is the Q3/Q4 quartile mapping correct, or is this program being undervalued?',
                'What is the absolute minimum compliance level, and can timing be deferred?'
            ],
            strengthenWith: ['Document mandate citation in writing', 'Pursue mandating-authority funding aggressively', 'Swap lower-impact spend', 'Phase to minimum viable scope', 'Sunset provision tied to mandate review']
        },
        'Low-Mandated-NonGF-Weak': {
            archetypeNumber: 15,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Mandate + non-GF justifies proceeding despite weak evidence',
            verifyNow: ['That mandate truly applies to this program'],
            strengthenWith: ['KPI baseline→target', '90-day review', 'Non-GF documentation']
        },
        'Low-Mandated-GFonly-Weak': {
            archetypeNumber: 16,
            disposition: 'VERIFY',
            color: '#6366f1',
            keyConsideration: 'Low-priority + 100% GF + weak outcomes evidence — mandate is the sole basis to fund. The bar for verification is highest here.',
            verifyNow: [
                'What does the mandate actually require — operate the program at all, or this specific incremental spend?',
                'Who is mandating it? (statute, regulation, court order, board motion) — get a citation',
                'Is there pass-through funding from the mandating authority to comply?',
                'Without strong outcomes evidence, can we afford the lowest-impact compliance path?',
                'Can the spend be delayed to a future cycle without violating the mandate?'
            ],
            strengthenWith: ['Document mandate citation in writing', 'Pursue mandating-authority funding aggressively', 'Tight scope — minimum viable compliance only', 'Hard sunset / mandate-review trigger', 'Required path to add non-GF funding within 6–12 months']
        },
        'Low-Compliance-NonGF-Strong': {
            archetypeNumber: 17,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Risk mitigation but ensure no GF creep',
            verifyNow: ['Non-GF terms & durability'],
            strengthenWith: ['No-GF pledge', 'Measurable risk reduction', 'Pilot + review']
        },
        'Low-Compliance-GFonly-Strong': {
            archetypeNumber: 18,
            disposition: 'MODIFY',
            color: '#ffc107',
            keyConsideration: 'Compliance need but require cost offsets for low-priority program',
            verifyNow: ['Why is this low-priority program requiring GF for compliance?'],
            strengthenWith: ['Require cost recovery mechanism', 'Internal reallocation from Q3/Q4 programs', 'Consider program redesign or elimination']
        },
        'Low-Compliance-NonGF-Weak': {
            archetypeNumber: 19,
            disposition: 'DEFER',
            color: '#dc3545',
            keyConsideration: 'Weak case even with non-GF - prove value first',
            verifyNow: ['Realism of benefits'],
            strengthenWith: ['Basic KPI set', 'Partner LOIs', 'Phase to prove value']
        },
        'Low-Compliance-GFonly-Weak': {
            archetypeNumber: 20,
            disposition: 'DEFER',
            color: '#dc3545',
            keyConsideration: 'Low priority + GF only + weak evidence = defer',
            verifyNow: ['If imminent risk, treat as mandate'],
            strengthenWith: ['Pilot w/ non-GF', 'Quantify liability avoided', 'Combine with Q1/Q2 work or eliminate']
        },
        'Low-None-NonGF-Strong': {
            archetypeNumber: 21,
            disposition: 'APPROVE',
            color: '#28a745',
            keyConsideration: 'Self-sustaining with strong outcomes - proceed if no GF needed',
            verifyNow: ['No GF drift', 'Sustainability of non-GF sources'],
            strengthenWith: ['Full cost recovery plan', 'Service redesign to increase relevance', 'Path to Q1/Q2 alignment']
        },
        'Low-None-GFonly-Strong': {
            archetypeNumber: 22,
            disposition: 'DEFER',
            color: '#dc3545',
            keyConsideration: 'Competes with higher-priority needs - phase behind Q1/Q2',
            verifyNow: ['Why use limited GF on low-priority program?'],
            strengthenWith: ['Add fee/grant/partner funding', 'ROI calculation showing strategic value', 'Phase behind Q1/Q2 priorities', 'Consider program elimination']
        },
        'Low-None-NonGF-Weak': {
            archetypeNumber: 23,
            disposition: 'DEFER',
            color: '#dc3545',
            keyConsideration: 'Prove demand and willingness-to-pay before proceeding',
            verifyNow: ['Is there any compelling reason to continue this program?'],
            strengthenWith: ['Strong KPIs showing strategic value', 'Tighten scope drastically', 'Prove demand/willingness-to-pay', 'Consider elimination']
        },
        'Low-None-GFonly-Weak': {
            archetypeNumber: 24,
            disposition: 'REJECT',
            color: '#dc3545',
            keyConsideration: 'No compelling case for limited GF resources - reframe or eliminate',
            verifyNow: ['N/A - does not meet funding criteria'],
            strengthenWith: ['Reframe to demonstrate higher-Q outcome alignment', 'Secure 100% non-GF funding', 'Consolidate with higher-priority programs', 'Recommend program elimination']
        }
    };
    
    const decision = grid[gridKey] || {
        archetypeNumber: 0,
        disposition: 'MODIFY',
        color: '#ffc107',
        keyConsideration: 'Unable to categorize - manual review needed',
        verifyNow: ['Unable to categorize - manual review needed'],
        strengthenWith: ['Provide complete information on mandate, funding, and outcomes']
    };
    
    decision.gridKey = gridKey;
    return decision;
}

// ===== SIMPLIFIED DECISION TREE EXPLANATION =====
function explainDecisionLogic(analysis) {
    const { quartileBand, mandateLevel, fundingType, outcomesStrength } = analysis;
    
    let logic = "\n**Decision Logic Applied:**\n\n";
    
    // Step 1: Check if asking for GF
    if (fundingType === 'GFonly') {
        logic += "➡️ This request asks for **General Fund money**\n\n";
        
        // Step 2: Check priority
        if (quartileBand === 'High') {
            logic += "✅ **High Priority** (Q1/Q2) - Advances strategic plan goals\n";
            
            if (outcomesStrength === 'Strong') {
                logic += "✅ **Strong Outcomes** - Clear metrics and evidence\n";
                logic += "**→ Strong candidate for APPROVE** - Using GF for strategic priorities with proven outcomes\n";
            } else {
                if (mandateLevel === 'Mandated') {
                    logic += "⚖️ **Mandated** - Legal/regulatory requirement\n";
                    logic += "**→ APPROVE with conditions** - Mandate requires it, but strengthen outcomes\n";
                } else if (mandateLevel === 'Compliance') {
                    logic += "⚠️ **Compliance/Risk** - Addresses safety or liability\n";
                    logic += "**→ MODIFY/DEFER** - High priority but weak case; compliance alone insufficient\n";
                } else {
                    logic += "❌ **Weak Outcomes** - Insufficient evidence\n";
                    logic += "**→ DEFER** - High priority isn't enough without strong evidence\n";
                }
            }
        } else {
            logic += "⚠️ **Low Priority** (Q3/Q4) - Lower strategic alignment\n";
            
            if (mandateLevel === 'Mandated') {
                logic += "⚖️ **Mandated** - Legal requirement forces consideration\n";
                logic += "**→ APPROVE minimum scope** - But aggressively pursue cost offsets\n";
            } else if (mandateLevel === 'Compliance') {
                logic += "⚠️ **Compliance/Risk** - Addresses safety or liability\n";
                logic += "**→ DEFER** - Low priority + GF request needs compelling risk justification\n";
            } else {
                logic += "❌ **No compelling justification** for GF use\n";
                logic += "**→ DEFER/REJECT** - Limited GF should prioritize Q1/Q2 programs\n";
            }
        }
    } else {
        logic += "💰 This request has **non-GF funding** (grants, fees, partnerships)\n\n";
        
        if (quartileBand === 'High') {
            logic += "✅ **High Priority** (Q1/Q2) + External funding\n";
            logic += "**→ APPROVE** - Leveraging external resources for strategic priorities\n";
        } else {
            logic += "⚠️ **Low Priority** (Q3/Q4) but external funding available\n";
            
            if (outcomesStrength === 'Strong') {
                logic += "✅ **Strong Outcomes** - Could increase program relevance\n";
                logic += "**→ APPROVE** - Investment could move program toward Q1/Q2 alignment\n";
            } else {
                logic += "❌ **Weak Outcomes** - Unclear strategic value\n";
                logic += "**→ MODIFY/DEFER** - Even with external funding, needs clearer strategic alignment\n";
            }
        }
    }
    
    return logic;
}

// ===== ENHANCED NARRATIVE GENERATOR =====
function generateEnhancedNarrative(request, lineItems, qa, analysis) {
    const requestId = getRequestId(request);
    const amounts = getRequestAmount(request);
    const dept = getPrimaryValue(lineItems, 'department') || 'Unknown';
    const program = getPrimaryValue(lineItems, 'program') || 'Unknown';
    
    let narrative = `**Program:** ${program} (${dept})\n`;
    narrative += `**Quartile:** ${analysis.bestQuartile} (${analysis.quartileBand} Relevance)\n`;
    narrative += `**Total Amount:** $${formatCurrency(amounts.total)}\n`;
    narrative += `**Decision Profile:** ${analysis.gridKey}\n\n`;
    
    narrative += `---\n\n`;
    
    // Context flags
    if (analysis.mandateLevel === 'Mandated') {
        narrative += `⚖️ **MANDATED**: This request is legally mandated or tied to a Board Motion/consent decree.\n\n`;
    } else if (analysis.mandateLevel === 'Compliance') {
        narrative += `**COMPLIANCE/RISK**: This request addresses compliance obligations or risk mitigation.\n\n`;
    }
    
    if (analysis.hasOutsideFunding) {
        narrative += `**NON-GF FUNDING**: Includes non-General Fund sources (grants, fees, or partnerships).\n\n`;
    } else if (analysis.quartileBand === 'Low') {
        narrative += `🚨 **FUNDING CONCERN**: 100% General Fund requested for a lower-relevance (Q3/Q4) program.\n\n`;
    }
    
    if (analysis.outcomesStrength === 'Strong') {
        narrative += `**STRONG EVIDENCE**: Clear performance metrics and outcome targets provided.\n\n`;
    } else {
        narrative += `📋 **WEAK EVIDENCE**: Insufficient outcome data, KPIs, or evaluation plan.\n\n`;
    }
    
    narrative += `---\n\n`;
    
    // Add decision tree explanation
    narrative += explainDecisionLogic(analysis);
    narrative += `\n---\n\n`;
    
    // Disposition and recommendation with PBB suggests language
    narrative += `**PBB FRAMEWORK SUGGESTS: ${analysis.disposition}** (Score: ${analysis.totalScore}/12)\n\n`;
    narrative += `*Note: This is an advisory recommendation based on textbook PBB methodology, not a final decision.*\n\n`;
    
    // Main recommendation based on disposition
    if (analysis.disposition === 'APPROVE') {
        narrative += `*PBB suggests APPROVE means this request meets the framework's funding criteria — it does NOT mean "fund regardless of cost." Even strong cases compete for finite General Fund resources, so all approvals are subject to overall budget capacity.*\n\n`;
        if (analysis.mandateLevel === 'Mandated') {
            narrative += `**PBB Framework Advisory:** PBB suggests APPROVE. This is a mandated program with ${analysis.outcomesStrength.toLowerCase()} outcomes evidence. `;
            if (analysis.fundingType === 'GFonly' && analysis.quartileBand === 'Low') {
                narrative += `Given the lower quartile, PBB suggests requiring offsetting reductions or pursuing non-GF sources. `;
            }
            if (analysis.outcomesStrength === 'Weak') {
                narrative += `PBB suggests requiring metrics and evaluation plan as condition of approval.\n\n`;
            } else {
                narrative += `General Fund support appears justified based on mandate requirements.\n\n`;
            }
        } else if (analysis.fundingType === 'NonGF') {
            narrative += `**PBB Framework Advisory:** PBB suggests APPROVE with non-GF priority. Strong proposal with external funding sources. `;
            if (analysis.quartileBand === 'Low') {
                narrative += `For Q3/Q4 programs, PBB suggests ensuring minimal or no GF backfill. `;
            }
            narrative += `PBB recommends proceeding with clear cost recovery and sustainability plan.\n\n`;
        } else {
            narrative += `**PBB Framework Advisory:** PBB suggests APPROVE but strengthen funding strategy. While outcomes are strong, PBB recommends adding cost recovery or partnership elements to reduce General Fund reliance.\n\n`;
        }
    } else if (analysis.disposition === 'VERIFY') {
        narrative += `**PBB Framework Advisory:** PBB suggests APPROVE — but only AFTER mandate verification. This is a low-priority (Q3/Q4) program with 100% General Fund reliance; the mandate is the only thing keeping it from REJECT. Before approving, document:\n\n`;
        narrative += `1. **What** does the mandate actually require — operate the program at all, or this specific incremental spend?\n`;
        narrative += `2. **Who** is mandating it (statute, regulation, court order, board motion) — get the citation in writing.\n`;
        narrative += `3. **Is there pass-through funding** from the mandating authority to help comply, or is the local government bearing the full cost?\n`;
        narrative += `4. **Minimum compliance path** — can scope, timing, or service level be reduced and still satisfy the mandate?\n\n`;
        narrative += `If the mandate is verified and binding, fund the minimum-viable compliance path with a sunset / mandate-review trigger. If the mandate is loose, stale, or doesn't actually require this incremental spend, this request should drop to DEFER or REJECT.\n\n`;
    } else if (analysis.disposition === 'MODIFY') {
        narrative += `**PBB Framework Advisory:** PBB suggests MODIFY before approval. This request shows merit but PBB recommends adjustments before proceeding:\n\n`;
    } else if (analysis.disposition === 'DEFER') {
        narrative += `**PBB Framework Advisory:** PBB suggests DEFER. Insufficient business case for current approval based on PBB criteria. `;
        if (analysis.mandateLevel === 'Mandated') {
            narrative += `PBB recommends monitoring mandate requirements. `;
        }
        narrative += `See PBB-recommended strengthening actions below.\n\n`;
    } else if (analysis.disposition === 'REJECT') {
        narrative += `**PBB Framework Advisory:** PBB suggests REJECT OR SIGNIFICANT REDESIGN. `;
        narrative += `This low-relevance, GF-only request with weak outcomes does not meet PBB funding criteria. PBB recommends fundamental changes before reconsideration.\n\n`;
    } else if (analysis.disposition === 'REVIEW') {
        narrative += `**PBB Framework Advisory:** PBB suggests MANUAL REVIEW. `;
        narrative += `This request is missing the quartile alignment data required to apply the PBB framework. Add a Quartile value to the line items, or upload a Program Inventory that includes a Quartile column, then re-run the analysis. Until then, no APPROVE/VERIFY/MODIFY/DEFER/REJECT recommendation should be inferred.\n\n`;
    }
    
    // Verification requirements
    if (analysis.verifyNow && analysis.verifyNow.length > 0 && analysis.verifyNow[0] !== 'N/A') {
        narrative += `**VERIFY NOW:**\n\n`;
        analysis.verifyNow.forEach(item => {
            narrative += `- ${item}\n`;
        });
        narrative += `\n`;
    }
    
    // Strengthening actions
    if (analysis.strengthenWith && analysis.strengthenWith.length > 0) {
        narrative += `**TO STRENGTHEN THIS REQUEST:**\n\n`;
        analysis.strengthenWith.forEach(item => {
            narrative += `- ${item}\n`;
        });
        narrative += `\n`;
    }
    
    // Specific follow-up prompts based on weaknesses
    narrative += `**SPECIFIC FOLLOW-UP ACTIONS:**\n\n`;
    
    if (analysis.outcomeScore < 2) {
        narrative += `**KPIs & Evaluation:** Please add baseline→target values for 2–3 KPIs, the data source, and review cadence (e.g., monthly). We'll approve as a 90-day pilot pending KPI progress.\n\n`;
    }
    
    if (analysis.fundingScore === 0 && (analysis.quartileBand === 'Low' || analysis.disposition !== 'APPROVE')) {
        narrative += `**Funding/Offsets:** Identify at least one non-GF source (fee, grant, partner, restricted fund) covering ≥30% of the request, or propose an internal reallocation/offset equal to ≥20%.\n\n`;
    }
    
    if (analysis.mandateLevel === 'Mandated' && analysis.outcomeScore < 2) {
        narrative += `**Mandate Evidence:** Attach the statute/board motion/consent decree citation and define the minimum compliance scope. Include milestones and success criteria.\n\n`;
    }
    
    if (analysis.mandateLevel === 'Compliance') {
        narrative += `**Risk Reduction:** Link this request to a specific risk register item and quantify the expected reduction (e.g., 'reduce audit findings by 50% in 12 months').\n\n`;
    }
    
    if (analysis.efficiencyScore < 2 && analysis.disposition !== 'REJECT') {
        narrative += `**ROI/Efficiency:** Provide a cost-avoidance or productivity calculation (unit cost, throughput, payback). If uncertain, start with a 6-month pilot and measure.\n\n`;
    }
    
    if (includeAccessEquity && analysis.equityScore < 2 && analysis.quartileBand === 'High') {
        narrative += `**Equity:** Name the priority population and specify a measurable access/outcome improvement (e.g., 'decrease wait time for X group from 12 to 6 weeks').\n\n`;
    }
    
    if (analysis.quartileBand === 'Low' && analysis.fundingType === 'GFonly') {
        narrative += `**Scope/Phasing:** Consider a phased approach (Phase 1 core features, Phase 2 optional enhancements) to reduce near-term GF use.\n\n`;
    }
    
    if (analysis.fundingScore === 1) {
        narrative += `**Partnership:** Add letters of intent (LOIs) or MOUs for partner contributions (space, staff time, cash match).\n\n`;
    }
    
    if (analysis.mandateLevel === 'Mandated' && analysis.quartileBand === 'Low') {
        narrative += `**Sunset/True-up:** Add a 12-month sunset and a true-up clause to right-size funding based on measured demand and KPI performance.\n\n`;
    }
    
    return narrative;
}

// ===== END OF SCORING ENGINE =====

function getRequestQA(requestId) {
    // Find Q&A entries for this request
    return budgetData.requestQA.filter(qa => {
        // Look for RequestID match in any field
        return Object.values(qa).some(value => 
            value && value.toString().trim() === requestId.toString().trim()
        );
    });
}

function generateProgramSummary() {
    console.log('Generating program summary...');
    
    const programData = {};
    
    // STEP 1: Aggregate data from budget REQUESTS (Personnel + NonPersonnel line items)
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department') || 'Unknown Department';
            const program = getPrimaryValue([item], 'program') || 'Unknown Program';
            const quartile = getPrimaryValue([item], 'quartile') || 'N/A';
            
            // Create department key if it doesn't exist
            if (!programData[dept]) {
                programData[dept] = {};
            }
            
            // Create program key if it doesn't exist
            if (!programData[dept][program]) {
                programData[dept][program] = {
                    quartile: quartile,
                    totalCost: 0,
                    requestedAmount: 0,
                    proposedTotalCost: 0,
                    requestCount: 0,
                    isNewProgram: false,
                    hasRequests: true // This program has requests
                };
            }
            
            // Add to requested amount using ACTUAL line item cost (not divided total)
            const lineItemAmount = getLineItemAmount(item);
            programData[dept][program].requestedAmount += lineItemAmount.total;
            programData[dept][program].requestCount++;
            
            // Get current budget from uploaded data or use $0 for new programs
            if (programData[dept][program].totalCost === 0) {
                const currentBudget = getCurrentBudgetForProgram(dept, program);
                if (currentBudget) {
                    programData[dept][program].totalCost = currentBudget.totalCost;
                    programData[dept][program].isNewProgram = false;
                    console.log(`Using real budget for ${program}: $${currentBudget.totalCost}`);
                } else {
                    // New program - no current budget
                    programData[dept][program].totalCost = 0;
                    programData[dept][program].isNewProgram = true;
                    console.log(`New program (no current budget): ${program}`);
                }
            }
            
            // Calculate proposed total
            programData[dept][program].proposedTotalCost = 
                programData[dept][program].totalCost + programData[dept][program].requestedAmount;
        });
    });
    
    // STEP 2: Add ALL programs from current budget that DON'T have requests
    if (currentBudgetData.length > 0) {
        console.log('Adding programs from current budget that have no requests...');
        
        currentBudgetData.forEach(budgetProg => {
            const dept = budgetProg['User Group'] || 'Unknown Department';
            const program = budgetProg['Program'] || 'Unknown Program';
            const totalCost = budgetProg['Total Program Cost'] || 0;
            
            // Skip if this program already has requests (already in programData)
            if (programData[dept] && programData[dept][program]) {
                return; // Already processed in Step 1
            }
            
            // This program exists in current budget but has NO requests
            if (!programData[dept]) {
                programData[dept] = {};
            }
            
            programData[dept][program] = {
                quartile: 'N/A', // No quartile since no request
                totalCost: totalCost,
                requestedAmount: 0, // No new requests
                proposedTotalCost: totalCost, // Same as current
                requestCount: 0,
                isNewProgram: false,
                hasRequests: false // This program has NO requests
            };
            
            console.log(`Added program with no requests: ${dept} - ${program} ($${totalCost})`);
        });
    }

    // STEP 3: Generate HTML output
    let html = `<div class="section-header" id="program-summary">Program Summary</div>
                <p>Below is a summary of programs showing current budget, requested amounts, and proposed totals, organized by department and quartile alignment.</p>`;
    
    // Generate table for each department
    Object.entries(programData).forEach(([dept, programs]) => {
        let departmentTotal = {
            totalCost: 0,
            requestedAmount: 0,
            proposedTotalCost: 0
        };
        
        html += `
            <div class="request-card">
                <div class="request-header">
                    <div class="request-title">${dept}</div>
                </div>
                <div class="request-details">
                    <table style="width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 0.9rem;">
                        <thead>
                            <tr style="background: #667eea; color: white;">
                                <th style="padding: 12px 8px; text-align: center; width: 80px;">Quartile</th>
                                <th style="padding: 12px 8px; text-align: left;">Program</th>
                                <th style="padding: 12px 8px; text-align: right; width: 120px;">Current Budget</th>
                                <th style="padding: 12px 8px; text-align: right; width: 120px;">Requested Amount</th>
                                <th style="padding: 12px 8px; text-align: right; width: 140px;">Proposed Total</th>
                            </tr>
                        </thead>
                        <tbody>
        `;
        
        // Sort programs: 
        // 1. Programs with requests come first (sorted by quartile)
        // 2. Programs without requests come last (alphabetically)
        const sortedPrograms = Object.entries(programs).sort((a, b) => {
            const aHasRequests = a[1].hasRequests;
            const bHasRequests = b[1].hasRequests;
            
            // Programs with requests come first
            if (aHasRequests && !bHasRequests) return -1;
            if (!aHasRequests && bHasRequests) return 1;
            
            // Within programs with requests, sort by quartile
            if (aHasRequests && bHasRequests) {
                const quartileOrder = {'Most Aligned': 1, 'More Aligned': 2, 'Less Aligned': 3, 'Least Aligned': 4, 'N/A': 5};
                const aOrder = quartileOrder[a[1].quartile] || 5;
                const bOrder = quartileOrder[b[1].quartile] || 5;
                return aOrder - bOrder;
            }
            
            // Within programs without requests, sort alphabetically
            return a[0].localeCompare(b[0]);
        });
        
        sortedPrograms.forEach(([program, data]) => {
            departmentTotal.totalCost += data.totalCost;
            departmentTotal.requestedAmount += data.requestedAmount;
            departmentTotal.proposedTotalCost += data.proposedTotalCost;
            
            // Quartile badge or "No Request" indicator
            let quartileBadge;
            if (data.hasRequests) {
                quartileBadge = data.quartile !== 'N/A' ? 
                    `<span class="quartile-badge quartile-${data.quartile.toLowerCase().replace(' ', '-')}" style="font-size: 0.8rem; padding: 4px 8px;">${data.quartile.replace(' Aligned', '')}</span>` : 
                    '<span style="color: #666;">N/A</span>';
            } else {
                quartileBadge = '<span style="color: #999; font-size: 0.7rem; font-style: italic;">No Request</span>';
            }
            
            // New program badge
            const newProgramBadge = data.isNewProgram ? 
                '<span style="background: #17a2b8; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.7rem; margin-left: 5px;">NEW PROGRAM</span>' : 
                '';
            
            // Calculate percentage increase
            let percentIncrease = '';
            if (data.totalCost > 0 && data.requestedAmount > 0) {
                const pct = ((data.requestedAmount / data.totalCost) * 100).toFixed(1);
                percentIncrease = ` (+${pct}%)`;
            } else if (data.totalCost === 0 && data.requestedAmount > 0) {
                percentIncrease = ' (New)';
            }
            
            // Row styling based on whether it has requests
            const rowStyle = data.hasRequests ? '' : 'background: #f9f9f9; opacity: 0.8;';
            
            html += `
                <tr style="border-bottom: 1px solid #e0e0e0; ${rowStyle}">
                    <td style="padding: 10px 8px; text-align: center;">${quartileBadge}</td>
                    <td style="padding: 10px 8px;">${program}${newProgramBadge}</td>
                    <td style="padding: 10px 8px; text-align: right; color: #333;">$${formatCurrency(Math.round(data.totalCost))}</td>
                    <td style="padding: 10px 8px; text-align: right; color: ${data.requestedAmount > 0 ? '#ffc107' : '#999'}; font-weight: ${data.requestedAmount > 0 ? '600' : 'normal'};">$${formatCurrency(Math.round(data.requestedAmount))}</td>
                    <td style="padding: 10px 8px; text-align: right; color: #28a745; font-weight: 600;">$${formatCurrency(Math.round(data.proposedTotalCost))}${percentIncrease}</td>
                </tr>
            `;
        });
        
        // Add department total row
        const deptPctIncrease = departmentTotal.totalCost > 0 ? 
            ((departmentTotal.requestedAmount / departmentTotal.totalCost) * 100).toFixed(1) : 
            'N/A';
        
        html += `
                <tr style="background: #f8f9ff; border-top: 2px solid #667eea; font-weight: 600;">
                    <td style="padding: 12px 8px; text-align: center; color: #667eea;">TOTAL</td>
                    <td style="padding: 12px 8px; color: #667eea;">${dept} Department Total</td>
                    <td style="padding: 12px 8px; text-align: right; color: #333;">$${formatCurrency(Math.round(departmentTotal.totalCost))}</td>
                    <td style="padding: 12px 8px; text-align: right; color: #ffc107;">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</td>
                    <td style="padding: 12px 8px; text-align: right; color: #28a745;">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</td>
                </tr>
            </tbody>
        </table>
        
        <div style="margin-top: 15px; padding: 10px; background: #f0f8ff; border-radius: 5px; border-left: 4px solid #667eea;">
            <strong>Department Impact Summary:</strong> ${dept} has ${Object.keys(programs).length} total programs 
            (${Object.values(programs).filter(p => p.hasRequests).length} with requests, 
            ${Object.values(programs).filter(p => !p.hasRequests).length} without requests).
            ${departmentTotal.requestedAmount > 0 ? `Requesting 
            <span style="color: #ffc107; font-weight: 600;">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</span> 
            in additional funding, which would increase the department's total budget from 
            <span style="color: #333;">$${formatCurrency(Math.round(departmentTotal.totalCost))}</span> to 
            <span style="color: #28a745; font-weight: 600;">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</span> 
            (${deptPctIncrease}% increase).` : 'No new funding requests for this department.'}
        </div>
        
        </div>
    </div>
        `;
    });

    return html;
}



function generateDepartmentSummary() {
    const departments = {};
    
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department');
            if (dept) {
                if (!departments[dept]) {
                    departments[dept] = { 
                        requests: new Set(), 
                        amount: 0,
                        programs: new Set(),
                        quartiles: {
                            'Most Aligned': 0,
                            'More Aligned': 0,
                            'Less Aligned': 0,
                            'Least Aligned': 0
                        }
                    };
                }
                departments[dept].requests.add(requestId);
                departments[dept].amount += amounts.total;
                
                const program = getPrimaryValue([item], 'program');
                if (program) departments[dept].programs.add(program);
                
                // Add quartile tracking
                const quartile = getPrimaryValue([item], 'quartile');
                if (quartile && departments[dept].quartiles.hasOwnProperty(quartile)) {
                    // Use ACTUAL line item cost
                    const lineItemAmount = getLineItemAmount(item);
                    departments[dept].quartiles[quartile] += lineItemAmount.total;
                }
            }
        });
    });

    let html = `<div class="section-header" id="department-summary">Department Summary</div>`;
    
    Object.entries(departments).forEach(([dept, data]) => {
        html += `
            <div class="request-card">
                <div class="request-header">
                    <div class="request-title">${dept}</div>
                </div>
                <div class="request-details">
                    <div class="detail-grid">
                        <div class="detail-item">
                            <div class="detail-label">Total Requests</div>
                            <div class="detail-value">${data.requests.size}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Programs Impacted</div>
                            <div class="detail-value">${data.programs.size}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Total Amount</div>
                            <div class="detail-value amount">$${formatCurrency(data.amount)}</div>
                        </div>
                    </div>
                    
                    <div style="margin-top: 20px;">
                        <h4 style="color: #667eea; margin-bottom: 10px;">Quartile Alignment Distribution</h4>
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-label">Most Aligned</div>
                                <div class="detail-value amount">$${formatCurrency(data.quartiles['Most Aligned'])}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">More Aligned</div>
                                <div class="detail-value amount">$${formatCurrency(data.quartiles['More Aligned'])}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Less Aligned</div>
                                <div class="detail-value amount">$${formatCurrency(data.quartiles['Less Aligned'])}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Least Aligned</div>
                                <div class="detail-value amount">$${formatCurrency(data.quartiles['Least Aligned'])}</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });

    return html;
}

function generateQuartileAnalysis() {
    const quartiles = {
        'Most Aligned': { count: 0, amount: 0 },
        'More Aligned': { count: 0, amount: 0 },
        'Less Aligned': { count: 0, amount: 0 },
        'Least Aligned': { count: 0, amount: 0 }
    };

    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartiles[quartile]) {
                quartiles[quartile].count++;
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartiles[quartile].amount += lineItemAmount.total;
            }
        });
    });

    let html = `<div class="section-header" id="quartile-analysis">Program Alignment Analysis</div>
               <p>Budget requests are categorized by their alignment to organizational priorities. Most Aligned programs receive the highest priority for funding consideration.</p>`;
    
    Object.entries(quartiles).forEach(([quartile, data]) => {
        const badgeClass = quartile.toLowerCase().replace(' ', '-');
        html += `
            <div class="request-card">
                <div class="request-header">
                    <div class="request-title">
                        <span class="quartile-badge quartile-${badgeClass}">${quartile}</span>
                    </div>
                </div>
                <div class="request-details">
                    <div class="detail-grid">
                        <div class="detail-item">
                            <div class="detail-label">Line Items</div>
                            <div class="detail-value">${data.count}</div>
                        </div>
                        <div class="detail-item">
                            <div class="detail-label">Total Amount</div>
                            <div class="detail-value amount">$${formatCurrency(data.amount)}</div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });

    return html;
}

function generateDetailedRequestReport() {
    let html = `<div class="section-header" id="individual-requests">Individual Budget Requests</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
        console.log(`Request ${requestId}: Found ${qa.length} Q&A items`);
        
        // Add page break style for each request (except the first)
        const pageBreakStyle = index > 0 ? 'page-break-before: always;' : '';

        html += `
            <div class="request-card" id="request-${requestId}" style="${pageBreakStyle} margin-top: 40px;">
                <div class="request-header" style="background: linear-gradient(135deg, #667eea, #764ba2); color: white;">
                    <div class="request-title" style="color: white; font-size: 1.4rem;">Request ID: ${requestId} - ${description}</div>
                </div>
                <div class="request-details">
                    <div style="margin-bottom: 25px;">
                        <h3 style="color: #667eea; margin-bottom: 15px; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Request Summary</h3>
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-label">Request ID</div>
                                <div class="detail-value">${requestId}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Description</div>
                                <div class="detail-value">${description}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Total Amount</div>
                                <div class="detail-value amount">$${formatCurrency(amounts.total)}</div>
                            </div>
                        </div>
                    </div>
        `;

        // Add Request Q&A section FIRST (most important context)
        if (qa.length > 0) {
            html += generateRequestQASection(qa);
        }

        // Add line item details
        if (lineItems.length > 0) {
            html += generateLineItemSection(lineItems);
        }

        html += `</div></div>`;
    });

    return html;
}

// ===== STANDARD REPORT (NO ANALYSIS) WITH COLLAPSIBLE REQUESTS =====
function generateDetailedRequestReportStandard() {
    let html = `
        <div class="section-header" id="individual-requests">Individual Budget Requests</div>
        
        <!-- Expand/Collapse All Controls -->
        <div class="collapse-controls">
            <button class="collapse-btn" onclick="expandAllStandardRequests()">📂 Expand All Requests</button>
            <button class="collapse-btn" onclick="collapseAllStandardRequests()">📁 Collapse All Requests</button>
        </div>
    `;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
        const uniqueId = `standard-request-accordion-${requestId}`;
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryProgram = getPrimaryValue(lineItems, 'program') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';

        html += `
            <!-- Request Accordion -->
            <div class="request-accordion" id="request-${requestId}">
                <div class="request-accordion-header" onclick="toggleRequestAccordion('${uniqueId}')">
                    <div class="request-accordion-title">
                        <strong>Request ${requestId}:</strong> ${description}
                    </div>
                    <span class="request-accordion-badge" style="background: #28a745;">
                        $${formatCurrency(amounts.total)}
                    </span>
                    ${primaryQuartile !== 'N/A' ? 
                        `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}" style="margin: 0 10px;">${primaryQuartile}</span>` 
                        : ''}
                    <span class="request-accordion-arrow" id="${uniqueId}-arrow">▼</span>
                </div>
                
                <div class="request-accordion-content" id="${uniqueId}">
                    <div class="request-accordion-body">
                        
                        <!-- Quick Summary Card -->
                        <div class="summary-card-compact">
                            <h4 style="color: #667eea; margin-bottom: 10px;">📊 Request Summary</h4>
                            <div class="summary-grid">
                                <div class="summary-item">
                                    <div class="summary-label">Request ID</div>
                                    <div class="summary-value">${requestId}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Total Amount</div>
                                    <div class="summary-value" style="color: #28a745;">$${formatCurrency(amounts.total)}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Department</div>
                                    <div class="summary-value" style="font-size: 1rem;">${primaryDept}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Program</div>
                                    <div class="summary-value" style="font-size: 1rem;">${primaryProgram}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Quartile</div>
                                    <div class="summary-value">
                                        ${primaryQuartile !== 'N/A' ? 
                                            `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}">${primaryQuartile}</span>` 
                                            : 'N/A'}
                                    </div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Line Items</div>
                                    <div class="summary-value">${lineItems.length}</div>
                                </div>
                            </div>
                        </div>
                        
                        <!-- REQUEST CONTEXT & DETAILS (Collapsible) -->
                        ${qa.length > 0 ? `
                        <div class="collapsible-header" onclick="toggleCollapsible('standard-qa-${requestId}')">
                            <h3>📋 Request Context & Details</h3>
                            <span class="collapsible-toggle" id="standard-qa-${requestId}-toggle">▼</span>
                        </div>
                        <div class="collapsible-content" id="standard-qa-${requestId}">
                            ${generateRequestQASection(qa)}
                        </div>
                        ` : ''}
                        
                        <!-- LINE ITEMS (Collapsible) -->
                        ${lineItems.length > 0 ? `
                        <div class="collapsible-header" onclick="toggleCollapsible('standard-line-items-${requestId}')">
                            <h3>💼 Line Item Details (${lineItems.length} items)</h3>
                            <span class="collapsible-toggle" id="standard-line-items-${requestId}-toggle">▼</span>
                        </div>
                        <div class="collapsible-content" id="standard-line-items-${requestId}">
                            ${generateLineItemSection(lineItems)}
                        </div>
                        ` : ''}
                        
                    </div>
                </div>
            </div>
        `;
    });

    return html;
}



// ===== ANALYTICAL REPORT (WITH SCORING) =====
function generateDetailedRequestReportAnalytical() {
    let html = `
        <div class="section-header" id="analytical-requests">Detailed Request Analysis</div>
        
        <!-- Expand/Collapse All Controls -->
        <div class="collapse-controls">
            <button class="collapse-btn" onclick="expandAllRequests()">📂 Expand All Requests</button>
            <button class="collapse-btn" onclick="collapseAllRequests()">📁 Collapse All Requests</button>
        </div>
    `;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
        // SCORE THE REQUEST
        const analysis = scoreRequest(request);
        
        const uniqueId = `request-accordion-${requestId}`;

        html += `
            <!-- Request Accordion -->
            <div class="request-accordion" id="analytical-request-${requestId}">
                <div class="request-accordion-header" onclick="toggleRequestAccordion('${uniqueId}')">
                    <div class="request-accordion-title">
                        <strong>Request ${requestId}:</strong> ${description}
                    </div>
                    <span class="request-accordion-badge" style="background: ${analysis.dispositionColor};">
                        ${analysis.disposition}
                    </span>
                    <span class="request-accordion-badge" style="background: #667eea;">
                        Archetype #${analysis.archetypeNumber}
                    </span>
                    <span class="request-accordion-arrow" id="${uniqueId}-arrow">▼</span>
                </div>
                
                <div class="request-accordion-content" id="${uniqueId}">
                    <div class="request-accordion-body">
                        
                        <!-- Quick Summary Card (Always Visible When Expanded) -->
                        <div class="summary-card-compact">
                            <h4 style="color: #667eea; margin-bottom: 10px;">📊 Quick Summary</h4>
                            <div class="summary-grid">
                                <div class="summary-item">
                                    <div class="summary-label">Request ID</div>
                                    <div class="summary-value">${requestId}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Total Amount</div>
                                    <div class="summary-value" style="color: #28a745;">$${formatCurrency(amounts.total)}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Department</div>
                                    <div class="summary-value" style="font-size: 1rem;">${getPrimaryValue(lineItems, 'department') || 'N/A'}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Program</div>
                                    <div class="summary-value" style="font-size: 1rem;">${getPrimaryValue(lineItems, 'program') || 'N/A'}</div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">Quartile</div>
                                    <div class="summary-value">
                                        ${analysis.bestQuartile
                                            ? `<span class="quartile-badge quartile-${analysis.bestQuartile.toString().toLowerCase().replace(/ /g, '-')}">${analysis.bestQuartile}</span>`
                                            : `<span class="quartile-badge">N/A</span>`}
                                    </div>
                                </div>
                                <div class="summary-item">
                                    <div class="summary-label">PBB Recommendation</div>
                                    <div class="summary-value" style="color: ${analysis.dispositionColor};">${analysis.disposition}</div>
                                </div>
                            </div>
                        </div>
                        
                        <!-- PBB ARCHETYPE & RECOMMENDATION (Collapsible) -->
                        <div class="collapsible-header" onclick="toggleCollapsible('pbb-score-${requestId}')">
                            <h3>PBB Archetype Analysis</h3>
                            <span class="collapsible-toggle" id="pbb-score-${requestId}-toggle">▼</span>
                        </div>
                        <div class="collapsible-content" id="pbb-score-${requestId}">
                            <div style="background: linear-gradient(135deg, #f8f9ff, #ffffff); padding: 25px; margin-bottom: 25px; border-radius: 8px; border: 2px solid ${analysis.dispositionColor};">
                                
                                <!-- PRIMARY: Archetype Determination -->
                                <div style="text-align: center; margin-bottom: 25px; padding: 25px; background: linear-gradient(135deg, ${analysis.dispositionColor}15, ${analysis.dispositionColor}05); border-radius: 12px; border: 2px solid ${analysis.dispositionColor};">
                                    <div style="font-size: 0.9rem; color: #666; text-transform: uppercase; letter-spacing: 0.1em; margin-bottom: 8px;">Budget Request Archetype</div>
                                    <div style="font-size: 3rem; font-weight: 800; color: ${analysis.dispositionColor}; margin-bottom: 5px;">
                                        #${analysis.archetypeNumber}
                                    </div>
                                    <div style="font-size: 1.8rem; font-weight: 700; color: ${analysis.dispositionColor}; margin-bottom: 15px;">
                                        ${analysis.disposition}
                                    </div>
                                    <div style="font-size: 1.1rem; color: #444; font-style: italic; max-width: 600px; margin: 0 auto; line-height: 1.5;">
                                        "${analysis.keyConsideration}"
                                    </div>
                                </div>
                                
                                <!-- 4 DECISION FACTORS (What determines the archetype) -->
                                <h4 style="color: #1e3a5f; margin: 25px 0 15px; font-size: 1.2rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px;">
                                    📊 Decision Factors (4 inputs that determine archetype)
                                </h4>
                                <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 25px;">
                                    <!-- Factor 1: Quartile -->
                                    <div style="background: ${analysis.quartileBand === 'High' ? 'linear-gradient(135deg, #d1fae5, #a7f3d0)' : 'linear-gradient(135deg, #fee2e2, #fecaca)'}; padding: 15px; border-radius: 10px; text-align: center; border: 2px solid ${analysis.quartileBand === 'High' ? '#10b981' : '#ef4444'};">
                                        <div style="font-size: 0.75rem; color: #666; text-transform: uppercase; margin-bottom: 5px;">Quartile</div>
                                        <div style="font-size: 1.3rem; font-weight: 700; color: ${analysis.quartileBand === 'High' ? '#059669' : '#dc2626'};">
                                            ${analysis.quartileBand === 'High' ? '🟢 High' : '🔴 Low'}
                                        </div>
                                        <div style="font-size: 0.85rem; color: #555;">${analysis.bestQuartile}</div>
                                    </div>
                                    <!-- Factor 2: Mandate -->
                                    <div style="background: ${analysis.mandateLevel === 'Mandated' ? 'linear-gradient(135deg, #dbeafe, #bfdbfe)' : analysis.mandateLevel === 'Compliance' ? 'linear-gradient(135deg, #fef3c7, #fde68a)' : 'linear-gradient(135deg, #f1f5f9, #e2e8f0)'}; padding: 15px; border-radius: 10px; text-align: center; border: 2px solid ${analysis.mandateLevel === 'Mandated' ? '#3b82f6' : analysis.mandateLevel === 'Compliance' ? '#f59e0b' : '#94a3b8'};">
                                        <div style="font-size: 0.75rem; color: #666; text-transform: uppercase; margin-bottom: 5px;">Mandate</div>
                                        <div style="font-size: 1.3rem; font-weight: 700; color: ${analysis.mandateLevel === 'Mandated' ? '#2563eb' : analysis.mandateLevel === 'Compliance' ? '#d97706' : '#64748b'};">
                                            ${analysis.mandateLevel === 'Mandated' ? '⚖️ Mandated' : analysis.mandateLevel === 'Compliance' ? '⚠️ Compliance' : '➖ None'}
                                        </div>
                                    </div>
                                    <!-- Factor 3: Funding -->
                                    ${(() => { const fd = getFundingDisplay(analysis); return `
                                    <div style="background: ${fd.grad}; padding: 15px; border-radius: 10px; text-align: center; border: 2px solid ${fd.border};">
                                        <div style="font-size: 0.75rem; color: #666; text-transform: uppercase; margin-bottom: 5px;">Funding</div>
                                        <div style="font-size: 1.3rem; font-weight: 700; color: ${fd.text};">
                                            ${fd.label}
                                        </div>
                                        ${fd.sub ? `<div style="font-size: 0.7rem; color: #555; margin-top: 4px;">${fd.sub}</div>` : ''}
                                    </div>`; })()}
                                    <!-- Factor 4: Evidence -->
                                    <div style="background: ${analysis.outcomesStrength === 'Strong' ? 'linear-gradient(135deg, #d1fae5, #a7f3d0)' : 'linear-gradient(135deg, #fee2e2, #fecaca)'}; padding: 15px; border-radius: 10px; text-align: center; border: 2px solid ${analysis.outcomesStrength === 'Strong' ? '#10b981' : '#ef4444'};">
                                        <div style="font-size: 0.75rem; color: #666; text-transform: uppercase; margin-bottom: 5px;">Evidence</div>
                                        <div style="font-size: 1.3rem; font-weight: 700; color: ${analysis.outcomesStrength === 'Strong' ? '#059669' : '#dc2626'};">
                                            ${analysis.outcomesStrength === 'Strong' ? '📊 Strong' : '📋 Weak'}
                                        </div>
                                    </div>
                                </div>
                                
                                <!-- ADDITIONAL CONSIDERATIONS (Efficiency & Access - informational only) -->
                                <h4 style="color: #64748b; margin: 25px 0 15px; font-size: 1rem;">
                                    📝 Additional Considerations (informational - do not affect archetype)
                                </h4>
                                <div style="display: grid; grid-template-columns: repeat(${includeAccessEquity ? 2 : 1}, 1fr); gap: 12px; margin-bottom: 25px; opacity: 0.85;">
                                    <div style="background: #f8fafc; padding: 12px 15px; border-radius: 8px; border-left: 3px solid #6f42c1;">
                                        <div style="display: flex; justify-content: space-between; align-items: center;">
                                            <span style="color: #6f42c1; font-weight: 600;">Efficiency/ROI</span>
                                            <span style="background: #6f42c1; color: white; padding: 3px 10px; border-radius: 12px; font-size: 0.85rem;">${analysis.efficiencyScore}/2</span>
                                        </div>
                                        <p style="margin: 5px 0 0; font-size: 0.85rem; color: #666;">${analysis.efficiencyReason}</p>
                                    </div>
                                    ${includeAccessEquity ? `<div style="background: #f8fafc; padding: 12px 15px; border-radius: 8px; border-left: 3px solid #e83e8c;">
                                        <div style="display: flex; justify-content: space-between; align-items: center;">
                                            <span style="color: #e83e8c; font-weight: 600;">Access/Equity</span>
                                            <span style="background: #e83e8c; color: white; padding: 3px 10px; border-radius: 12px; font-size: 0.85rem;">${analysis.accessScore}/2</span>
                                        </div>
                                        <p style="margin: 5px 0 0; font-size: 0.85rem; color: #666;">${analysis.accessReason}</p>
                                    </div>` : ''}
                                </div>
                                
                                <!-- Strategic Recommendation -->
                                <div class="narrative-box" style="background: #f8f9ff; padding: 20px; border-radius: 8px; border-left: 4px solid ${analysis.dispositionColor};">
                                    <h4 style="color: #667eea; margin-bottom: 15px; font-size: 1.2rem;">🔍 Strategic Recommendation</h4>
                                    <div style="white-space: pre-wrap; font-size: 1.05rem; line-height: 1.8;">${markdownToHtml(analysis.narrative)}</div>
                                </div>
                            </div>
                        </div>
                        
                        <!-- REQUEST DETAILS (Collapsible) -->
                        <div class="collapsible-header" onclick="toggleCollapsible('request-qa-${requestId}')">
                            <h3>📋 Request Context & Details</h3>
                            <span class="collapsible-toggle" id="request-qa-${requestId}-toggle">▼</span>
                        </div>
                        <div class="collapsible-content" id="request-qa-${requestId}">
        `;

        if (qa.length > 0) {
            html += generateRequestQASection(qa);
        }

        html += `</div>`;
        
        // LINE ITEMS (Collapsible)
        html += `
                        <div class="collapsible-header" onclick="toggleCollapsible('line-items-${requestId}')">
                            <h3>💼 Line Item Details (${lineItems.length} items)</h3>
                            <span class="collapsible-toggle" id="line-items-${requestId}-toggle">▼</span>
                        </div>
                        <div class="collapsible-content" id="line-items-${requestId}">
        `;
        
        if (lineItems.length > 0) {
            html += generateLineItemSection(lineItems);
        }

        html += `
                        </div>
                    </div>
                </div>
            </div>
        `;
    });

    return html;
}



// Summary for Analytical Report
function generateAnalyticalSummary() {
    const scores = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };
    const amounts = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };

    filteredData.forEach(request => {
        const analysis = scoreRequest(request);
        const requestAmounts = getRequestAmount(request);
        const key = (analysis.disposition || '').toLowerCase();
        if (scores[key] !== undefined) {
            scores[key]++;
            amounts[key] += requestAmounts.total;
        }
    });

    const tile = (label, count, amt, bg, border, textColor) => `
        <div class="detail-item" style="background: ${bg}; border: 2px solid ${border};">
            <div class="detail-label">${label}</div>
            <div class="detail-value" style="font-size: 1.5rem; color: ${textColor};">${count} Requests</div>
            <div class="amount" style="font-size: 1.2rem; color: ${textColor};">$${formatCurrency(amt)}</div>
        </div>
    `;

    // Only show REVIEW tile when there are review-state requests, to avoid clutter when
    // quartile data is fully populated.
    const reviewTile = scores.review > 0
        ? tile('PBB Framework Needs Review', scores.review, amounts.review,
               'linear-gradient(135deg, #e2e8f0, #cbd5e1)', '#64748b', '#475569')
        : '';

    return `
        <div class="section-header">Recommendation Summary</div>
        <div class="request-card">
            <div class="request-details">
                <div class="detail-grid">
                    ${tile('PBB Framework Suggests Approve', scores.approve, amounts.approve,
                           'linear-gradient(135deg, #d4edda, #c3e6cb)', '#28a745', '#28a745')}
                    ${tile('PBB Suggests Verify Mandate', scores.verify, amounts.verify,
                           'linear-gradient(135deg, #e0e7ff, #c7d2fe)', '#6366f1', '#4338ca')}
                    ${tile('PBB Framework Suggests Modify', scores.modify, amounts.modify,
                           'linear-gradient(135deg, #fff3cd, #ffeeba)', '#ffc107', '#856404')}
                    ${tile('PBB Framework Suggests Defer', scores.defer, amounts.defer,
                           'linear-gradient(135deg, #ffe5d0, #ffd1a8)', '#fd7e14', '#9a4a09')}
                    ${tile('PBB Framework Suggests Reject', scores.reject, amounts.reject,
                           'linear-gradient(135deg, #f8d7da, #f5c6cb)', '#dc3545', '#dc3545')}
                    ${reviewTile}
                </div>
            </div>
        </div>
    `;
}

// Table of Contents for Analytical Report
function generateAnalyticalTableOfContents() {
    let html = `
        <div class="section-header">Table of Contents</div>
        <div class="request-card">
            <div class="request-details">
                <ol style="line-height: 2; font-size: 1.1rem;">
                    <li><a href="#analytical-requests" style="color: #667eea; text-decoration: none;">Detailed Request Analysis</a>
                        <ol style="margin-top: 10px; font-size: 1rem;">
    `;

    filteredData.forEach((request) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const analysis = scoreRequest(request);
        const badgeColor = analysis.disposition === 'APPROVE' ? '#28a745' :
                          analysis.disposition === 'VERIFY' ? '#6366f1' :
                          analysis.disposition === 'MODIFY' ? '#ffc107' :
                          analysis.disposition === 'DEFER'  ? '#fd7e14' :
                          analysis.disposition === 'REJECT' ? '#dc3545' :
                          '#64748b';
        
        html += `<li>
            <a href="#analytical-request-${requestId}" style="color: #667eea; text-decoration: none;">
                Request ${requestId}: ${description || 'N/A'}
            </a>
            <span style="background: ${badgeColor}; color: white; padding: 2px 8px; border-radius: 10px; font-size: 0.8rem; margin-left: 10px;">
                ${analysis.disposition} (${analysis.totalScore}/12)
            </span>
        </li>`;
    });

    html += `
                        </ol>
                    </li>
                </ol>
            </div>
        </div>
    `;

    return html;
}

// Download functions for analytical report
function downloadAnalyticalWordReport() {
    if (filteredData.length === 0) {
        alert('Please generate a report first.');
        return;
    }
    
    const reportDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    let totalAmount = 0, totalOngoing = 0, totalOnetime = 0;
    const dStats = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };
    const dAmounts = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };

    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalAmount += amounts.total;
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        const analysis = scoreRequest(request);
        const disp = (analysis.disposition || '').toLowerCase();
        if (dStats[disp] !== undefined) { dStats[disp]++; dAmounts[disp] += amounts.total; }
    });
    
    // Build summary table rows
    let summaryRows = '';
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);
        const analysis = scoreRequest(request);
        const shortDesc = description && description.length > 30 ? description.substring(0, 30) + '...' : (description || 'N/A');
        
        const dispColor = analysis.disposition === 'APPROVE' ? '#10b981' :
                         analysis.disposition === 'VERIFY'  ? '#6366f1' :
                         analysis.disposition === 'MODIFY'  ? '#f59e0b' :
                         analysis.disposition === 'DEFER'   ? '#fd7e14' :
                         analysis.disposition === 'REJECT'  ? '#ef4444' :
                         '#64748b';

        summaryRows += `
            <tr>
                <td style="padding: 8px; border: 1px solid #e2e8f0;">${requestId}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0;">${shortDesc}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0;">${primaryDept}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0;">${primaryQuartile}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0; text-align: right;">$${formatCurrency(amounts.total)}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0; text-align: center; font-weight: bold;">#${analysis.archetypeNumber}</td>
                <td style="padding: 8px; border: 1px solid #e2e8f0; background: ${dispColor}; color: white; font-weight: bold; text-align: center;">${analysis.disposition}</td>
            </tr>
        `;
    });
    
    // Build detailed analysis for each request
    let detailedAnalysis = '';
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);
        const analysis = scoreRequest(request);
        
        const dispColor = analysis.disposition === 'APPROVE' ? '#10b981' :
                         analysis.disposition === 'VERIFY'  ? '#6366f1' :
                         analysis.disposition === 'MODIFY'  ? '#f59e0b' :
                         analysis.disposition === 'DEFER'   ? '#fd7e14' :
                         analysis.disposition === 'REJECT'  ? '#ef4444' :
                         '#64748b';

        // Q&A section
        let qaHtml = '';
        if (qa.length > 0) {
            qa.forEach(qItem => {
                let question = '', answer = '';
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && !lowerKey.includes('type') && qItem[key]) question = qItem[key];
                    if (lowerKey.includes('answer') && qItem[key]) answer = qItem[key];
                });
                if (question && answer && answer.trim()) {
                    qaHtml += `<div style="background: #fffbeb; border-left: 4px solid #f59e0b; padding: 10px; margin: 8px 0;"><strong style="color: #1e3a5f;">${question}</strong><br/>${answer}</div>`;
                }
            });
        }
        
        detailedAnalysis += `
            <div style="page-break-inside: avoid; margin-bottom: 30px; border: 2px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
                <div style="background: linear-gradient(135deg, #1e3a5f, #2a4a73); color: white; padding: 15px;">
                    <strong style="font-size: 16px;">Request ${requestId}: ${description || 'No Description'}</strong>
                </div>
                <div style="padding: 15px;">
                    <table style="width: 100%; margin-bottom: 15px;">
                        <tr>
                            <td style="width: 16%; padding: 8px; background: #f8fafc; border-radius: 4px; text-align: center;"><div style="font-size: 10px; color: #64748b;">Department</div><div style="font-weight: 600;">${primaryDept}</div></td>
                            <td style="width: 16%; padding: 8px; background: #f8fafc; border-radius: 4px; text-align: center;"><div style="font-size: 10px; color: #64748b;">Quartile</div><div style="font-weight: 600;">${primaryQuartile}</div></td>
                            <td style="width: 16%; padding: 8px; background: #f8fafc; border-radius: 4px; text-align: center;"><div style="font-size: 10px; color: #64748b;">Total Amount</div><div style="font-weight: 600; color: #10b981;">$${formatCurrency(amounts.total)}</div></td>
                            <td style="width: 16%; padding: 8px; background: #f8fafc; border-radius: 4px; text-align: center;"><div style="font-size: 10px; color: #64748b;">Ongoing</div><div style="font-weight: 600;">$${formatCurrency(amounts.ongoing)}</div></td>
                            <td style="width: 16%; padding: 8px; background: #f8fafc; border-radius: 4px; text-align: center;"><div style="font-size: 10px; color: #64748b;">One-time</div><div style="font-weight: 600;">$${formatCurrency(amounts.onetime)}</div></td>
                            <td style="width: 16%; padding: 8px; background: ${dispColor}; border-radius: 4px; text-align: center; color: white;"><div style="font-size: 10px;">Recommendation</div><div style="font-weight: 700;">${analysis.disposition}</div></td>
                        </tr>
                    </table>
                    
                    <h4 style="color: #1e3a5f; margin: 15px 0 10px; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">🎯 Archetype #${analysis.archetypeNumber}: ${analysis.disposition}</h4>
                    <p style="font-style: italic; color: #555; margin-bottom: 15px;">"${analysis.keyConsideration}"</p>
                    
                    <h4 style="color: #64748b; margin: 15px 0 10px;">Decision Factors (4 inputs)</h4>
                    <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                        <tr style="background: #f8fafc;">
                            <td style="padding: 8px; width: 25%;"><strong>Quartile</strong></td>
                            <td style="padding: 8px; width: 15%; text-align: center; font-weight: 700; color: ${analysis.quartileBand === 'High' ? '#059669' : '#dc2626'};">${analysis.quartileBand === 'High' ? '🟢 High' : '🔴 Low'}</td>
                            <td style="padding: 8px;">${analysis.quartileReason}</td>
                        </tr>
                        <tr>
                            <td style="padding: 8px;"><strong>Mandate</strong></td>
                            <td style="padding: 8px; text-align: center; font-weight: 700;">${analysis.mandateLevel === 'Mandated' ? '⚖️ Mandated' : analysis.mandateLevel === 'Compliance' ? '⚠️ Compliance' : '➖ None'}</td>
                            <td style="padding: 8px;">${analysis.mandateReason}</td>
                        </tr>
                        <tr style="background: #f8fafc;">
                            <td style="padding: 8px;"><strong>Funding</strong></td>
                            <td style="padding: 8px; text-align: center; font-weight: 700; color: ${getFundingDisplay(analysis).text};">${getFundingDisplay(analysis).label}</td>
                            <td style="padding: 8px;">${analysis.fundingReason}</td>
                        </tr>
                        <tr>
                            <td style="padding: 8px;"><strong>Evidence</strong></td>
                            <td style="padding: 8px; text-align: center; font-weight: 700; color: ${analysis.outcomesStrength === 'Strong' ? '#059669' : '#dc2626'};">${analysis.outcomesStrength === 'Strong' ? '📊 Strong' : '📋 Weak'}</td>
                            <td style="padding: 8px;">${analysis.outcomeReason}</td>
                        </tr>
                    </table>
                    
                    <h4 style="color: #64748b; margin: 15px 0 10px;">Additional Considerations (informational)</h4>
                    <table style="width: 100%; border-collapse: collapse; font-size: 11px; opacity: 0.85;">
                        <tr style="background: #f8fafc;">
                            <td style="padding: 8px; width: 25%;"><strong>Efficiency/ROI</strong></td>
                            <td style="padding: 8px; width: 15%; text-align: center;">${analysis.efficiencyScore}/2</td>
                            <td style="padding: 8px;">${analysis.efficiencyReason}</td>
                        </tr>
                        ${includeAccessEquity ? `<tr>
                            <td style="padding: 8px;"><strong>Access/Equity</strong></td>
                            <td style="padding: 8px; text-align: center;">${analysis.accessScore}/2</td>
                            <td style="padding: 8px;">${analysis.accessReason}</td>
                        </tr>` : ''}
                    </table>
                    
                    <div style="margin-top: 15px; padding: 12px; background: #f0f9ff; border-radius: 6px; border-left: 4px solid #0ea5e9;">
                        <strong style="color: #0369a1;">Overall Rationale:</strong><br/>
                        ${analysis.narrative}
                    </div>
                    
                    ${qaHtml ? `<h4 style="color: #1e3a5f; margin: 20px 0 10px; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px;">Request Context & Details</h4>${qaHtml}` : ''}
                </div>
            </div>
        `;
    });
    
    const wordHtml = `
        <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word">
        <head><meta charset="UTF-8"><title>PBB Analysis Report</title></head>
        <body style="font-family: Arial, sans-serif; margin: 40px; line-height: 1.5;">
            <div style="text-align: center; margin-bottom: 40px; padding-bottom: 20px; border-bottom: 3px solid #10b981;">
                <h1 style="color: #1e3a5f; font-size: 28px; margin-bottom: 10px;">🎯 PBB Analysis & Recommendations Report</h1>
                <p style="color: #64748b; font-size: 14px;">Priority Based Budgeting Framework Analysis</p>
                <p style="color: #64748b; font-size: 12px;">Generated on ${reportDate}</p>
            </div>
            
            <div style="background: #fff7ed; border: 2px solid #f59e0b; padding: 15px; border-radius: 8px; margin-bottom: 30px;">
                <p style="margin: 0; color: #92400e;"><strong>⚠️ Advisory Report:</strong> This analysis represents what a textbook PBB framework would suggest. These are recommendations to inform decision-making, not actual funding decisions.</p>
            </div>
            
            <h2 style="color: #1e3a5f; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px;">Executive Summary</h2>
            <p>This report analyzes <strong>${filteredData.length} budget requests</strong> totaling <strong style="color: #10b981;">$${formatCurrency(totalAmount)}</strong>.</p>
            
            <table style="width: 100%; border-collapse: separate; border-spacing: 8px; margin: 20px 0;">
                <tr>
                    <td style="width: 20%; padding: 18px; text-align: center; background: linear-gradient(135deg, #d1fae5, #a7f3d0); border-radius: 8px;">
                        <div style="font-size: 30px; font-weight: bold; color: #059669;">${dStats.approve}</div>
                        <div style="color: #065f46; font-weight: 600;">✓ APPROVE</div>
                        <div style="font-size: 12px; color: #065f46;">$${formatCurrency(dAmounts.approve)}</div>
                    </td>
                    <td style="width: 20%; padding: 18px; text-align: center; background: linear-gradient(135deg, #e0e7ff, #c7d2fe); border-radius: 8px;">
                        <div style="font-size: 30px; font-weight: bold; color: #4f46e5;">${dStats.verify}</div>
                        <div style="color: #3730a3; font-weight: 600;">🔍 VERIFY</div>
                        <div style="font-size: 12px; color: #3730a3;">$${formatCurrency(dAmounts.verify)}</div>
                    </td>
                    <td style="width: 20%; padding: 18px; text-align: center; background: linear-gradient(135deg, #fef3c7, #fde68a); border-radius: 8px;">
                        <div style="font-size: 30px; font-weight: bold; color: #d97706;">${dStats.modify}</div>
                        <div style="color: #92400e; font-weight: 600;">⚠ MODIFY</div>
                        <div style="font-size: 12px; color: #92400e;">$${formatCurrency(dAmounts.modify)}</div>
                    </td>
                    <td style="width: 20%; padding: 18px; text-align: center; background: linear-gradient(135deg, #e2e8f0, #cbd5e1); border-radius: 8px;">
                        <div style="font-size: 30px; font-weight: bold; color: #475569;">${dStats.defer}</div>
                        <div style="color: #334155; font-weight: 600;">⏸ DEFER</div>
                        <div style="font-size: 12px; color: #334155;">$${formatCurrency(dAmounts.defer)}</div>
                    </td>
                    <td style="width: 20%; padding: 18px; text-align: center; background: linear-gradient(135deg, #fee2e2, #fecaca); border-radius: 8px;">
                        <div style="font-size: 30px; font-weight: bold; color: #dc2626;">${dStats.reject}</div>
                        <div style="color: #991b1b; font-weight: 600;">✗ REJECT</div>
                        <div style="font-size: 12px; color: #991b1b;">$${formatCurrency(dAmounts.reject)}</div>
                    </td>
                </tr>
            </table>
            
            <h2 style="color: #1e3a5f; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; margin-top: 40px;">Summary Table</h2>
            <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                <thead>
                    <tr style="background: #1e3a5f; color: white;">
                        <th style="padding: 10px; text-align: left;">ID</th>
                        <th style="padding: 10px; text-align: left;">Description</th>
                        <th style="padding: 10px; text-align: left;">Department</th>
                        <th style="padding: 10px; text-align: left;">Quartile</th>
                        <th style="padding: 10px; text-align: right;">Amount</th>
                        <th style="padding: 10px; text-align: center;">Archetype</th>
                        <th style="padding: 10px; text-align: center;">Recommendation</th>
                    </tr>
                </thead>
                <tbody>${summaryRows}</tbody>
            </table>
            
            <h2 style="color: #1e3a5f; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; margin-top: 40px;">Detailed Request Analysis</h2>
            ${detailedAnalysis}
            
            <div style="margin-top: 40px; padding: 20px; background: #f1f5f9; border-radius: 8px; text-align: center;">
                <p style="margin: 0; color: #64748b; font-size: 12px;">Generated by PBB Budget Request Analyzer • Tyler Technologies Budget Intelligence</p>
            </div>
        </body>
        </html>
    `;
    
    const blob = new Blob([wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `PBB_Analysis_Report_${new Date().toISOString().split('T')[0]}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function downloadAnalyticalPdfReport() {
    if (filteredData.length === 0) {
        alert('Please generate a report first.');
        return;
    }
    
    const reportDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    let totalAmount = 0, totalOngoing = 0, totalOnetime = 0;
    const dStats = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };
    const dAmounts = { approve: 0, verify: 0, modify: 0, defer: 0, reject: 0, review: 0 };
    const deptStats = {};
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalAmount += amounts.total;
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        const analysis = scoreRequest(request);
        const disp = (analysis.disposition || '').toLowerCase();
        if (dStats[disp] !== undefined) { dStats[disp]++; dAmounts[disp] += amounts.total; }
        
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const dept = getPrimaryValue(lineItems, 'department') || 'Unknown';
        if (!deptStats[dept]) deptStats[dept] = { count: 0, amount: 0 };
        deptStats[dept].count++;
        deptStats[dept].amount += amounts.total;
    });
    
    // Build summary table rows
    let tableRows = '';
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);
        const analysis = scoreRequest(request);
        const shortDesc = description && description.length > 35 ? description.substring(0, 35) + '...' : (description || 'N/A');
        
        const qBadge = primaryQuartile.includes('Most') || primaryQuartile.includes('More') ? 'badge-high' : 'badge-low';
        const dispBadge = analysis.disposition === 'APPROVE' ? 'badge-approve' :
                         analysis.disposition === 'VERIFY'  ? 'badge-verify' :
                         analysis.disposition === 'MODIFY'  ? 'badge-modify' :
                         analysis.disposition === 'DEFER'   ? 'badge-defer' :
                         analysis.disposition === 'REJECT'  ? 'badge-reject' :
                         'badge-review';
        
        tableRows += `<tr>
            <td>${requestId}</td>
            <td>${shortDesc}</td>
            <td>${primaryDept}</td>
            <td><span class="badge ${qBadge}">${primaryQuartile}</span></td>
            <td class="amount">$${formatCurrency(amounts.total)}</td>
            <td style="text-align: center; font-weight: bold;">#${analysis.archetypeNumber}</td>
            <td><span class="badge ${dispBadge}">${analysis.disposition}</span></td>
        </tr>`;
    });
    
    // Build detailed analysis pages for ALL requests
    let detailedPagesHtml = '';
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);
        const analysis = scoreRequest(request);
        
        const dispColor = analysis.disposition === 'APPROVE' ? '#10b981' :
                         analysis.disposition === 'VERIFY'  ? '#6366f1' :
                         analysis.disposition === 'MODIFY'  ? '#f59e0b' :
                         analysis.disposition === 'DEFER'   ? '#fd7e14' :
                         analysis.disposition === 'REJECT'  ? '#ef4444' :
                         '#64748b';

        // Q&A Section
        let qaHtml = '';
        if (qa.length > 0) {
            qa.forEach(qItem => {
                let question = '', answer = '';
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && !lowerKey.includes('type') && qItem[key]) question = qItem[key];
                    if (lowerKey.includes('answer') && qItem[key]) answer = qItem[key];
                });
                if (!question) {
                    const questionKeys = ['Question', 'C', 'Col_2', 'Col_C'];
                    for (const key of questionKeys) {
                        if (qItem[key] && qItem[key].toString().trim()) { question = qItem[key]; break; }
                    }
                }
                if (question && answer && answer.trim()) {
                    qaHtml += `<div class="qa-item"><div class="qa-question">${question}</div><div class="qa-answer">${answer}</div></div>`;
                }
            });
        }
        
        // Line Items with all fields
        let lineItemsHtml = '';
        lineItems.forEach((item, idx) => {
            const itemQuartile = getPrimaryValue([item], 'quartile');
            const qClass = itemQuartile && (itemQuartile.includes('Most') || itemQuartile.includes('More')) ? 'badge-high' : 'badge-low';
            
            let fieldsHtml = '<div class="field-grid">';
            const allFields = ['REQUESTID', 'REQUEST TYPE', 'STATUS', 'ONGOING COST', 'ONETIME COST', 'FUND', 'DEPARTMENT', 'PROGRAM', 'PROGRAMID', 'QUARTILE'];
            allFields.forEach(field => {
                const value = findFieldValue(item, field);
                if (value !== null) {
                    const displayValue = formatFieldValue(field, value);
                    fieldsHtml += `<div class="field-item"><div class="field-label">${field}</div><div class="field-value">${displayValue}</div></div>`;
                }
            });
            fieldsHtml += '</div>';
            
            lineItemsHtml += `<div class="line-item-card"><div class="line-item-header">Line Item ${idx + 1} ${itemQuartile ? `<span class="badge ${qClass}">${itemQuartile}</span>` : ''}</div>${fieldsHtml}</div>`;
        });
        
        detailedPagesHtml += `
            <div class="request-page page-break">
                <div class="page-header"><span class="page-title">Request Analysis</span><span class="page-number">${index + 1} of ${filteredData.length}</span></div>
                <div class="request-card">
                    <div class="request-header" style="background: linear-gradient(135deg, ${dispColor}, ${dispColor}dd);">Request ${requestId}: ${description || 'No Description'}</div>
                    <div class="request-body">
                        <!-- Archetype Badge -->
                        <div style="text-align: center; padding: 15px; margin-bottom: 15px; background: linear-gradient(135deg, ${dispColor}15, ${dispColor}05); border-radius: 8px; border: 2px solid ${dispColor};">
                            <div style="font-size: 8px; color: #666; text-transform: uppercase;">Archetype</div>
                            <div style="font-size: 24px; font-weight: 800; color: ${dispColor};">#${analysis.archetypeNumber} ${analysis.disposition}</div>
                            <div style="font-size: 9px; color: #444; font-style: italic;">"${analysis.keyConsideration}"</div>
                        </div>
                        
                        <div class="meta-grid">
                            <div class="meta-item"><div class="meta-label">Department</div><div class="meta-value">${primaryDept}</div></div>
                            <div class="meta-item"><div class="meta-label">Total Amount</div><div class="meta-value amount">$${formatCurrency(amounts.total)}</div></div>
                            <div class="meta-item"><div class="meta-label">Ongoing</div><div class="meta-value">$${formatCurrency(amounts.ongoing)}</div></div>
                            <div class="meta-item"><div class="meta-label">One-time</div><div class="meta-value">$${formatCurrency(amounts.onetime)}</div></div>
                        </div>
                        
                        <h4 class="section-header">Decision Factors (4 inputs)</h4>
                        <div class="scoring-grid" style="grid-template-columns: repeat(4, 1fr);">
                            <div class="score-card" style="background: ${analysis.quartileBand === 'High' ? '#d1fae5' : '#fee2e2'};"><div class="score-name">Quartile</div><div class="score-value" style="color: ${analysis.quartileBand === 'High' ? '#059669' : '#dc2626'};">${analysis.quartileBand === 'High' ? '🟢 High' : '🔴 Low'}</div><div class="score-reason">${analysis.bestQuartile}</div></div>
                            <div class="score-card"><div class="score-name">Mandate</div><div class="score-value">${analysis.mandateLevel === 'Mandated' ? '⚖️' : analysis.mandateLevel === 'Compliance' ? '⚠️' : '➖'}</div><div class="score-reason">${analysis.mandateLevel}</div></div>
                            ${(() => { const fd = getFundingDisplay(analysis); return `<div class="score-card" style="background: ${fd.plain};"><div class="score-name">Funding</div><div class="score-value" style="color: ${fd.text};">${fd.label}</div>${fd.sub ? `<div class="score-reason">${fd.sub}</div>` : ''}</div>`; })()}
                            <div class="score-card" style="background: ${analysis.outcomesStrength === 'Strong' ? '#d1fae5' : '#fee2e2'};"><div class="score-name">Evidence</div><div class="score-value" style="color: ${analysis.outcomesStrength === 'Strong' ? '#059669' : '#dc2626'};">${analysis.outcomesStrength === 'Strong' ? '📊 Strong' : '📋 Weak'}</div></div>
                        </div>
                        
                        <h4 class="section-header" style="color: #64748b;">Additional Considerations</h4>
                        <div class="scoring-grid" style="grid-template-columns: repeat(${includeAccessEquity ? 2 : 1}, 1fr); opacity: 0.85;">
                            <div class="score-card"><div class="score-name">Efficiency/ROI</div><div class="score-value">${analysis.efficiencyScore}/2</div><div class="score-reason">${analysis.efficiencyReason}</div></div>
                            ${includeAccessEquity ? `<div class="score-card"><div class="score-name">Access/Equity</div><div class="score-value">${analysis.accessScore}/2</div><div class="score-reason">${analysis.accessReason}</div></div>` : ''}
                        </div>
                        
                        <div class="rationale-box"><strong>Overall Rationale:</strong> ${analysis.narrative}</div>
                        
                        ${qaHtml ? `<h4 class="section-header">Request Context & Details</h4><div class="qa-section">${qaHtml}</div>` : ''}
                        
                        <h4 class="section-header">Line Item Details</h4>
                        ${lineItemsHtml}
                    </div>
                </div>
            </div>
        `;
    });
    
    const pdfHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>PBB Analysis Report</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
@page { size: A4; margin: 0.5in; }
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Inter', sans-serif; line-height: 1.5; color: #1e293b; background: white; font-size: 10px; }

/* Cover Page - Green Theme */
.cover-page { height: 100vh; display: flex; flex-direction: column; background: linear-gradient(135deg, #059669 0%, #10b981 50%, #059669 100%); color: white; page-break-after: always; }
.cover-header { padding: 40px 60px; }
.cover-brand { font-size: 14px; font-weight: 500; letter-spacing: 0.05em; opacity: 0.9; }
.cover-main { flex: 1; display: flex; flex-direction: column; justify-content: center; padding: 0 60px; }
.cover-title { font-size: 42px; font-weight: 700; line-height: 1.1; margin-bottom: 20px; }
.cover-subtitle { font-size: 20px; font-weight: 300; opacity: 0.9; margin-bottom: 40px; }
.cover-stats { display: flex; gap: 40px; margin-top: 30px; }
.cover-stat-value { font-size: 32px; font-weight: 700; }
.cover-stat-label { font-size: 12px; opacity: 0.8; text-transform: uppercase; letter-spacing: 0.05em; }
.cover-footer { padding: 40px 60px; border-top: 1px solid rgba(255,255,255,0.2); display: flex; justify-content: space-between; font-size: 12px; opacity: 0.7; }

/* Content Pages */
.content-page, .request-page { padding: 35px 45px; }
.page-header { display: flex; justify-content: space-between; padding-bottom: 10px; border-bottom: 2px solid #e2e8f0; margin-bottom: 20px; }
.page-title { font-size: 10px; text-transform: uppercase; letter-spacing: 0.1em; color: #64748b; font-weight: 600; }
.page-number { font-size: 10px; color: #64748b; }
.section-title { font-size: 22px; font-weight: 700; color: #059669; margin-bottom: 6px; }
.section-subtitle { font-size: 13px; color: #64748b; margin-bottom: 20px; }
.section-header { font-size: 13px; font-weight: 700; color: #059669; margin: 20px 0 10px; border-bottom: 1px solid #e2e8f0; padding-bottom: 5px; }
.page-break { page-break-before: always; }

/* Summary Cards */
.summary-cards { display: grid; grid-template-columns: repeat(5, 1fr); gap: 10px; margin-bottom: 25px; }
.summary-card { padding: 14px 12px; border-radius: 10px; text-align: center; }
.summary-card.approve { background: linear-gradient(135deg, #d1fae5, #a7f3d0); border: 2px solid #10b981; }
.summary-card.verify { background: linear-gradient(135deg, #e0e7ff, #c7d2fe); border: 2px solid #6366f1; }
.summary-card.modify { background: linear-gradient(135deg, #fef3c7, #fde68a); border: 2px solid #f59e0b; }
.summary-card.defer { background: linear-gradient(135deg, #e2e8f0, #cbd5e1); border: 2px solid #64748b; }
.summary-card.reject { background: linear-gradient(135deg, #fee2e2, #fecaca); border: 2px solid #ef4444; }
.summary-card-value { font-size: 26px; font-weight: 700; }
.summary-card.approve .summary-card-value { color: #059669; }
.summary-card.verify .summary-card-value { color: #4f46e5; }
.summary-card.modify .summary-card-value { color: #d97706; }
.summary-card.defer .summary-card-value { color: #475569; }
.summary-card.reject .summary-card-value { color: #dc2626; }
.summary-card-label { font-size: 11px; font-weight: 600; margin-top: 3px; }
.summary-card-amount { font-size: 12px; margin-top: 6px; font-weight: 500; }

/* Data Table */
.data-table { width: 100%; border-collapse: collapse; font-size: 9px; }
.data-table thead { background: #059669; color: white; }
.data-table th { padding: 8px 6px; text-align: left; font-weight: 600; text-transform: uppercase; font-size: 8px; }
.data-table td { padding: 8px 6px; border-bottom: 1px solid #e2e8f0; }
.data-table tr:nth-child(even) { background: #f8fafc; }

/* Badges */
.badge { display: inline-block; padding: 2px 6px; border-radius: 10px; font-size: 8px; font-weight: 600; }
.badge-approve { background: #10b981; color: white; }
.badge-verify { background: #6366f1; color: white; }
.badge-modify { background: #f59e0b; color: white; }
.badge-defer { background: #fd7e14; color: white; }
.badge-reject { background: #ef4444; color: white; }
.badge-review { background: #64748b; color: white; }
.badge-high { background: #10b981; color: white; }
.badge-low { background: #64748b; color: white; }
.amount { color: #059669; font-weight: 600; }

/* Request Cards */
.request-card { border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden; }
.request-header { background: linear-gradient(135deg, #059669, #10b981); color: white; padding: 12px 16px; font-weight: 600; font-size: 13px; }
.request-body { padding: 16px; }
.meta-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 8px; margin-bottom: 15px; }
.meta-item { background: #f8fafc; padding: 8px; border-radius: 5px; text-align: center; border: 1px solid #e2e8f0; }
.meta-item.highlight { border: none; }
.meta-label { font-size: 8px; color: #64748b; font-weight: 600; margin-bottom: 2px; }
.meta-value { font-size: 11px; color: #1e293b; font-weight: 500; }

/* Scoring Grid */
.scoring-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 15px; }
.score-card { background: #f8fafc; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0; }
.score-name { font-size: 9px; font-weight: 600; color: #059669; margin-bottom: 3px; }
.score-value { font-size: 16px; font-weight: 700; color: #1e3a5f; margin-bottom: 4px; }
.score-reason { font-size: 8px; color: #64748b; line-height: 1.4; }

/* Rationale Box */
.rationale-box { margin: 15px 0; padding: 12px; background: #f0fdf4; border-radius: 6px; border-left: 4px solid #10b981; font-size: 10px; line-height: 1.5; }

/* Q&A Section */
.qa-section { margin-bottom: 15px; }
.qa-item { background: #fffbeb; border-left: 3px solid #f59e0b; padding: 10px; margin-bottom: 8px; border-radius: 0 5px 5px 0; }
.qa-question { font-weight: 600; color: #1e3a5f; font-size: 10px; margin-bottom: 4px; }
.qa-answer { color: #475569; font-size: 9px; line-height: 1.4; }

/* Line Items */
.line-item-card { margin-bottom: 10px; padding: 10px; background: #f8fafc; border-radius: 6px; border-left: 3px solid #667eea; }
.line-item-header { font-weight: 600; font-size: 10px; color: #1e3a5f; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; }
.field-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; }
.field-item { background: white; padding: 5px; border-radius: 3px; border: 1px solid #e2e8f0; text-align: center; }
.field-label { font-size: 6px; color: #64748b; font-weight: 600; text-transform: uppercase; }
.field-value { font-size: 8px; color: #1e293b; font-weight: 500; word-break: break-word; }

@media print { 
    .page-break { page-break-before: always; } 
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .request-card { break-inside: avoid; }
    .score-card { break-inside: avoid; }
}
</style></head><body>

<!-- COVER PAGE -->
<div class="cover-page">
    <div class="cover-header"><span class="cover-brand">🎯 PBB FRAMEWORK ANALYSIS • TYLER TECHNOLOGIES</span></div>
    <div class="cover-main">
        <h1 class="cover-title">PBB Analysis &<br>Recommendations Report</h1>
        <p class="cover-subtitle">Comprehensive Priority Based Budgeting Framework Scoring</p>
        <div class="cover-stats">
            <div class="cover-stat"><div class="cover-stat-value">${filteredData.length}</div><div class="cover-stat-label">Requests Analyzed</div></div>
            <div class="cover-stat"><div class="cover-stat-value">$${formatCurrency(totalAmount)}</div><div class="cover-stat-label">Total Amount</div></div>
            <div class="cover-stat"><div class="cover-stat-value">${dStats.approve}</div><div class="cover-stat-label">Recommended Approvals</div></div>
        </div>
    </div>
    <div class="cover-footer"><span>Generated on ${reportDate}</span><span>Advisory Analysis • Not Binding Decisions</span></div>
</div>

<!-- EXECUTIVE SUMMARY -->
<div class="content-page page-break">
    <div class="page-header"><span class="page-title">Executive Summary</span></div>
    <h2 class="section-title">PBB Framework Recommendations</h2>
    <p class="section-subtitle">Analysis of ${filteredData.length} budget requests totaling $${formatCurrency(totalAmount)}</p>
    
    <div class="summary-cards">
        <div class="summary-card approve"><div class="summary-card-value">${dStats.approve}</div><div class="summary-card-label">✓ Approve</div><div class="summary-card-amount">$${formatCurrency(dAmounts.approve)}</div></div>
        <div class="summary-card verify"><div class="summary-card-value">${dStats.verify}</div><div class="summary-card-label">🔍 Verify Mandate</div><div class="summary-card-amount">$${formatCurrency(dAmounts.verify)}</div></div>
        <div class="summary-card modify"><div class="summary-card-value">${dStats.modify}</div><div class="summary-card-label">⚠ Modify</div><div class="summary-card-amount">$${formatCurrency(dAmounts.modify)}</div></div>
        <div class="summary-card defer"><div class="summary-card-value">${dStats.defer}</div><div class="summary-card-label">⏸ Defer</div><div class="summary-card-amount">$${formatCurrency(dAmounts.defer)}</div></div>
        <div class="summary-card reject"><div class="summary-card-value">${dStats.reject}</div><div class="summary-card-label">✗ Reject</div><div class="summary-card-amount">$${formatCurrency(dAmounts.reject)}</div></div>
    </div>
    
    <h3 style="font-size: 16px; color: #1e3a5f; margin: 25px 0 12px;">All Requests Summary</h3>
    <table class="data-table">
        <thead><tr><th>ID</th><th>Description</th><th>Department</th><th>Quartile</th><th>Amount</th><th>Archetype</th><th>Recommendation</th></tr></thead>
        <tbody>${tableRows}</tbody>
    </table>
</div>

<!-- DETAILED REQUEST ANALYSIS PAGES -->
${detailedPagesHtml}

</body></html>`;
    
    const newWindow = window.open('', '_blank');
    newWindow.document.write(pdfHtml);
    newWindow.document.close();
    newWindow.focus();
    setTimeout(() => alert('Comprehensive PBB Analysis Report opened!\\n\\nTo save as PDF:\\n1. Press Ctrl+P (Cmd+P on Mac)\\n2. Select "Save as PDF"\\n3. Click Save'), 500);
}

function generateRequestQASection(qa) {
    if (qa.length === 0) return '';
    
    let html = `
        <div style="margin-bottom: 25px;">
            <h3 style="color: #667eea; margin-bottom: 15px; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Request Context & Details</h3>
    `;
    
    qa.forEach(qItem => {
        // Find question and answer fields - UPDATED LOGIC
        let question = '';
        let answer = '';
        
        Object.keys(qItem).forEach(key => {
            const lowerKey = key.toLowerCase();
            // Look for Column C (Question) instead of Column F (Question Type)
            if (lowerKey.includes('question') && !lowerKey.includes('type') && qItem[key]) {
                question = qItem[key];
            }
            if (lowerKey.includes('answer') && qItem[key]) {
                answer = qItem[key];
            }
        });
        
        // If no question found with above logic, try direct column references
        if (!question) {
            // Try common column names for the actual question text
            const questionKeys = ['Question', 'C', 'Col_2', 'Col_C'];
            for (const key of questionKeys) {
                if (qItem[key] && qItem[key].toString().trim()) {
                    question = qItem[key];
                    break;
                }
            }
        }
        
        if (question && answer && answer.trim()) {
            html += `
                <div style="margin: 15px 0; padding: 20px; background: #fff8f0; border-radius: 8px; border-left: 4px solid #ffc107;">
                    <div style="font-weight: 600; color: #667eea; margin-bottom: 12px; font-size: 1.1rem;">${question}</div>
                    <div style="line-height: 1.6; font-size: 1rem; color: #333;">${answer}</div>
                </div>
            `;
        }
    });
    
    html += '</div>';
    return html;
}

function generateLineItemSection(lineItems) {
    let html = `
        <div style="margin-bottom: 25px;">
            <h3 style="color: #667eea; margin-bottom: 15px; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Line Item Details</h3>
    `;
    
    lineItems.forEach((item, idx) => {
        // Get quartile for badge
        const quartile = getPrimaryValue([item], 'quartile');
        const quartileBadge = quartile ? 
            `<span class="quartile-badge quartile-${quartile.toLowerCase().replace(' ', '-')}" style="margin-left: 10px;">${quartile}</span>` : 
            '';

        html += `
            <div style="margin: 15px 0; padding: 15px; background: #f8f9ff; border-radius: 5px; border-left: 4px solid #667eea;">
                <div style="font-weight: 600; margin-bottom: 10px;">Line Item ${idx + 1} ${quartileBadge}</div>
                <div class="detail-grid">
        `;
        
        // Show all fields from this line item
        Object.entries(item).forEach(([key, value]) => {
            if (value !== null && value !== undefined && value.toString().trim() !== '') {
                // Use the centralized formatting function
                const displayValue = formatFieldValue(key, value);

                html += `
                    <div class="detail-item">
                        <div class="detail-label">${key}</div>
                        <div class="detail-value">${displayValue}</div>
                    </div>
                `;
            }
        });
        
        
        html += `
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    return html;
}

// Replace this function in your app.js file for the WEB UI charts

function generateCharts() {
    return `
        <div class="section-header" id="visual-analysis">Visual Analysis</div>
        <div class="charts-section">
            <div class="chart-container">
                <canvas id="departmentChart" width="400" height="200"></canvas>
            </div>
            <div class="chart-container">
                <canvas id="quartileChart" width="400" height="200"></canvas>
            </div>
        </div>
    `;
}

function renderCharts() {
    // Department chart
    const departments = {};
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department');
            if (dept) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                departments[dept] = (departments[dept] || 0) + lineItemAmount.total;
            }
        });
    });

    if (Object.keys(departments).length > 0) {
        new Chart(document.getElementById('departmentChart'), {
            type: 'bar',
            data: {
                labels: Object.keys(departments),
                datasets: [{
                    label: 'Total Requested Amount',
                    data: Object.values(departments),
                    backgroundColor: ['#667eea', '#764ba2', '#f093fb', '#f5576c', '#4facfe']
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Budget Requests by Department'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return '$' + value.toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    }

    // CHANGED: Quartile chart from pie to bar chart
    const quartiles = {
        'Most Aligned': 0,
        'More Aligned': 0,
        'Less Aligned': 0,
        'Least Aligned': 0
    };

    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartiles.hasOwnProperty(quartile)) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartiles[quartile] += lineItemAmount.total;
            }
        });
    });

    if (Object.values(quartiles).some(val => val > 0)) {
        new Chart(document.getElementById('quartileChart'), {
            type: 'bar',
            data: {
                labels: Object.keys(quartiles),
                datasets: [{
                    label: 'Total Budget Amount',
                    data: Object.values(quartiles),
                    backgroundColor: ['#28a745', '#17a2b8', '#ffc107', '#dc3545']
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    title: {
                        display: true,
                        text: 'Budget Requests by Quartile Alignment'
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return '$' + value.toLocaleString();
                            }
                        }
                    }
                }
            }
        });
    }
}

function generateWordProgramSummary() {
    // Reuse the same program aggregation logic
    const programData = {};
    
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department') || 'Unknown Department';
            const program = getPrimaryValue([item], 'program') || 'Unknown Program';
            const quartile = getPrimaryValue([item], 'quartile') || 'N/A';
            
            if (!programData[dept]) {
                programData[dept] = {};
            }
            
            if (!programData[dept][program]) {
                programData[dept][program] = {
                    quartile: quartile,
                    totalCost: 0,
                    requestedAmount: 0,
                    proposedTotalCost: 0,
                    requestCount: 0
                };
            }
            
            // Use ACTUAL line item cost (not divided total)
            const lineItemAmount = getLineItemAmount(item);
            programData[dept][program].requestedAmount += lineItemAmount.total;
            programData[dept][program].requestCount++;
            
            // Get current budget from uploaded Program Inventory data
            if (programData[dept][program].totalCost === 0) {
                const currentBudget = getCurrentBudgetForProgram(dept, program);
                if (currentBudget) {
                    programData[dept][program].totalCost = currentBudget.totalCost;
                }
                // If no current budget found, leave as 0 (new program)
            }
            
            programData[dept][program].proposedTotalCost = 
                programData[dept][program].totalCost + programData[dept][program].requestedAmount;
        });
    });

    let html = `
        <div class="section-header" id="program-summary">Program Summary</div>
        <p>Below is a summary of programs and their total requested amount and potential new total cost, organized by department and quartile alignment.</p>
    `;
    
    Object.entries(programData).forEach(([dept, programs]) => {
        let departmentTotal = {
            totalCost: 0,
            requestedAmount: 0,
            proposedTotalCost: 0
        };
        
        html += `
            <div class="card">
                <div class="card-header">${dept}</div>
                <div class="card-body">
                    <table style="width: 100%; font-size: 11px;">
                        <thead>
                            <tr style="background: #667eea; color: white;">
                                <th style="padding: 8px 6px; text-align: center;">Quartile</th>
                                <th style="padding: 8px 6px; text-align: left;">Program</th>
                                <th style="padding: 8px 6px; text-align: right;">Total Cost</th>
                                <th style="padding: 8px 6px; text-align: right;">Requested</th>
                                <th style="padding: 8px 6px; text-align: right;">Proposed Total</th>
                            </tr>
                        </thead>
                        <tbody>
        `;
        
        const sortedPrograms = Object.entries(programs).sort((a, b) => {
            const quartileOrder = {'Most Aligned': 1, 'More Aligned': 2, 'Less Aligned': 3, 'Least Aligned': 4};
            const aOrder = quartileOrder[a[1].quartile] || 5;
            const bOrder = quartileOrder[b[1].quartile] || 5;
            return aOrder - bOrder;
        });
        
        sortedPrograms.forEach(([program, data]) => {
            departmentTotal.totalCost += data.totalCost;
            departmentTotal.requestedAmount += data.requestedAmount;
            departmentTotal.proposedTotalCost += data.proposedTotalCost;
            
            const quartileBadge = data.quartile !== 'N/A' ? 
                `<span class="quartile-badge quartile-${data.quartile.toLowerCase().replace(' ', '-')}" style="font-size: 8px; padding: 2px 6px;">${data.quartile.replace(' Aligned', '')}</span>` : 
                'N/A';
            
            html += `
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 6px 4px; text-align: center;">${quartileBadge}</td>
                    <td style="padding: 6px 4px; font-size: 10px;">${program}</td>
                    <td style="padding: 6px 4px; text-align: right;">$${formatCurrency(Math.round(data.totalCost))}</td>
                    <td style="padding: 6px 4px; text-align: right; color: #ffc107;" class="amount">$${formatCurrency(Math.round(data.requestedAmount))}</td>
                    <td style="padding: 6px 4px; text-align: right; color: #28a745;" class="amount">$${formatCurrency(Math.round(data.proposedTotalCost))}</td>
                </tr>
            `;
        });
        
        html += `
                <tr style="background: #f8f9ff; border-top: 2px solid #667eea; font-weight: 600;">
                    <td style="padding: 8px 4px; text-align: center; color: #667eea;">TOTAL</td>
                    <td style="padding: 8px 4px; color: #667eea; font-size: 10px;">${dept} Total</td>
                    <td style="padding: 8px 4px; text-align: right;">$${formatCurrency(Math.round(departmentTotal.totalCost))}</td>
                    <td style="padding: 8px 4px; text-align: right; color: #ffc107;">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</td>
                    <td style="padding: 8px 4px; text-align: right; color: #28a745;">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</td>
                </tr>
            </tbody>
        </table>
        
        <div style="margin-top: 10px; padding: 8px; background: #f0f8ff; border-radius: 5px; font-size: 10px;">
            <strong>Impact:</strong> ${Object.keys(programs).length} programs requesting 
            <span class="amount">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</span>, 
            increasing budget from $${formatCurrency(Math.round(departmentTotal.totalCost))} to 
            <span class="amount">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</span> 
            (${((departmentTotal.requestedAmount / departmentTotal.totalCost) * 100).toFixed(1)}% increase).
        </div>
        
        </div>
    </div>
        `;
    });

    return html;
}

function downloadWordReport() {
    // Generate the report content fresh for Word format
    const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    const totalAmount = filteredData.reduce((sum, request) => {
        const amounts = getRequestAmount(request);
        return sum + amounts.total;
    }, 0);

    // Create comprehensive Word document with enhanced formatting
    let wordHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Priority Based Budgeting Report</title>
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    margin: 40px; 
                    line-height: 1.6; 
                    color: #333;
                }
                .header { 
                    text-align: center; 
                    margin-bottom: 40px; 
                    padding-bottom: 20px;
                    border-bottom: 3px solid #667eea;
                }
                .header h1 { 
                    color: #667eea; 
                    font-size: 2.5rem; 
                    margin-bottom: 10px; 
                }
                .header p { 
                    color: #666; 
                    font-size: 1.1rem; 
                    margin: 5px 0;
                }
                .section-header { 
                    color: #667eea; 
                    font-size: 1.3rem; 
                    font-weight: 600; 
                    margin: 40px 0 20px 0; 
                    border-bottom: 2px solid #e0e0e0; 
                    padding-bottom: 10px; 
                    page-break-after: avoid;
                }
                .card { 
                    border: 2px solid #e0e0e0; 
                    margin: 20px 0; 
                    border-radius: 8px; 
                    page-break-inside: avoid;
                    background: #fafafa;
                }
                .card-header { 
                    background: #667eea; 
                    color: white; 
                    padding: 15px 20px; 
                    font-size: 1.3rem; 
                    font-weight: 600; 
                }
                .card-body { 
                    padding: 20px; 
                }
                .detail-grid { 
                    display: table; 
                    width: 100%; 
                    border-collapse: collapse;
                }
                .detail-row { 
                    display: table-row; 
                }
                .detail-cell { 
                    display: table-cell; 
                    padding: 8px 12px; 
                    border-bottom: 1px solid #eee;
                    vertical-align: top;
                    width: 50%;
                }
                .detail-label { 
                    font-weight: 600; 
                    color: #555; 
                }
                .detail-value { 
                    color: #333; 
                }
                .amount { 
                    font-weight: 600; 
                    color: #28a745; 
                    font-size: 1.1rem;
                }
                .quartile-badge { 
                    display: inline-block; 
                    padding: 6px 16px; 
                    border-radius: 20px; 
                    font-size: 0.9rem; 
                    font-weight: 600; 
                    color: white;
                    margin: 2px;
                }
                .quartile-most, .quartile-most-aligned { background: #28a745; }
                .quartile-more, .quartile-more-aligned { background: #17a2b8; }
                .quartile-less, .quartile-less-aligned { background: #ffc107; color: black; }
                .quartile-least, .quartile-least-aligned { background: #dc3545; }
                table { 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 20px 0; 
                    font-size: 0.95rem;
                }
                th { 
                    background: #667eea; 
                    color: white; 
                    padding: 12px 8px; 
                    text-align: left; 
                    font-weight: 600; 
                }
                td { 
                    padding: 10px 8px; 
                    border-bottom: 1px solid #ddd; 
                    vertical-align: top;
                }
                tr:nth-child(even) { 
                    background: #f8f9ff; 
                }
                .toc { 
                    background: #f8f9ff; 
                    padding: 20px; 
                    border-radius: 8px; 
                    margin: 20px 0;
                }
                .toc ol { 
                    line-height: 1.8; 
                    font-size: 1.1rem; 
                }
                .toc li { 
                    margin: 8px 0; 
                }
                .toc a {
                    color: #667eea;
                    text-decoration: none;
                }
                .toc a:hover {
                    text-decoration: underline;
                }
                .qa-section {
                    background: #fff8f0;
                    border-left: 4px solid #ffc107;
                    padding: 15px 20px;
                    margin: 15px 0;
                    border-radius: 0 8px 8px 0;
                }
                .qa-question {
                    font-weight: 600;
                    color: #667eea;
                    font-size: 1.1rem;
                    margin-bottom: 8px;
                }
                .qa-answer {
                    line-height: 1.6;
                    color: #333;
                }
                .line-item {
                    background: #f8f9ff;
                    border-left: 4px solid #667eea;
                    padding: 15px;
                    margin: 15px 0;
                    border-radius: 0 8px 8px 0;
                }
                .line-item-header {
                    font-weight: 600;
                    margin-bottom: 10px;
                    color: #333;
                }
                .page-break { 
                    page-break-before: always; 
                }
                .section-break {
                    page-break-before: always;
                }
                .summary-stats {
                    background: #e8f4fd;
                    border: 2px solid #667eea;
                    border-radius: 8px;
                    padding: 20px;
                    margin: 20px 0;
                }
                .stats-grid {
                    display: table;
                    width: 100%;
                }
                .stats-row {
                    display: table-row;
                }
                .stats-cell {
                    display: table-cell;
                    text-align: center;
                    padding: 15px;
                    border-right: 1px solid #ccc;
                }
                .stats-cell:last-child {
                    border-right: none;
                }
                .stats-value {
                    font-size: 1.5rem;
                    font-weight: bold;
                    color: #667eea;
                    display: block;
                }
                .stats-label {
                    color: #666;
                    font-size: 0.9rem;
                    margin-top: 5px;
                }
                .chart-placeholder {
                    background: #f8f9ff;
                    border: 2px dashed #667eea;
                    border-radius: 8px;
                    padding: 40px 20px;
                    text-align: center;
                    margin: 20px 0;
                    color: #667eea;
                    font-size: 1.1rem;
                    font-weight: 600;
                }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Priority Based Budgeting Report</h1>
                <p>Budget Request Analysis and Recommendations</p>
                <p>Generated on ${reportDate}</p>
            </div>
    `;

    // Executive Summary
    wordHtml += `
        <div class="section-header">Executive Summary</div>
        <p>This comprehensive report analyzes <strong>${filteredData.length} budget requests</strong> totaling <strong class="amount">$${formatCurrency(totalAmount)}</strong> in requested funding. The requests span multiple departments and programs, with varying levels of alignment to organizational priorities.</p>
    `;

    // Filter Summary with page break
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordFilterSummary();

    // Visual Analysis Section - ADDED
    wordHtml += generateWordVisualAnalysis();

    // Table of Contents with clickable links
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordTableOfContents();

    // Request Summary Table
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordRequestTable();

    // Department Summary  
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordDepartmentSummary();

    // Program Summary
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordProgramSummary();

    // Individual Requests
    wordHtml += `<div class="section-break"></div>`;
    wordHtml += generateWordDetailedRequests();

    wordHtml += `
        </body>
        </html>
    `;

    const blob = new Blob([wordHtml], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Priority_Based_Budgeting_Report_${new Date().toISOString().split('T')[0]}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// NEW: Generate Visual Analysis section for Word document

function generateWordDetailedRequests() {
    let html = `<div class="section-header" id="individual-requests">Individual Budget Requests</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
        const pageBreak = index > 0 ? 'page-break' : '';

        html += `
            <div class="card ${pageBreak}" id="request-${requestId}">
                <div class="card-header">Request ${requestId}: ${description}</div>
                <div class="card-body">
                    <div class="detail-grid">
                        <div class="detail-row">
                            <div class="detail-cell detail-label">Request ID:</div>
                            <div class="detail-cell detail-value">${requestId}</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-cell detail-label">Description:</div>
                            <div class="detail-cell detail-value">${description}</div>
                        </div>
                        <div class="detail-row">
                            <div class="detail-cell detail-label">Total Amount:</div>
                            <div class="detail-cell detail-value amount">$${formatCurrency(amounts.total)}</div>
                        </div>
                    </div>
        `;

        // Add Q&A
        if (qa.length > 0) {
            html += `<h4 style="color: #667eea; margin: 20px 0 15px 0;">Request Context & Details</h4>`;
            qa.forEach(qItem => {
                let question = '';
                let answer = '';
                
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && qItem[key]) {
                        question = qItem[key];
                    }
                    if (lowerKey.includes('answer') && qItem[key]) {
                        answer = qItem[key];
                    }
                });
                
                if (question && answer && answer.trim()) {
                    html += `
                        <div class="qa-section">
                            <div class="qa-question">${question}</div>
                            <div class="qa-answer">${answer}</div>
                        </div>
                    `;
                }
            });
        }

        // Add line items
        if (lineItems.length > 0) {
            html += `<h4 style="color: #667eea; margin: 20px 0 15px 0;">Line Item Details</h4>`;
            lineItems.forEach((item, idx) => {
                const quartile = getPrimaryValue([item], 'quartile');
                const quartileBadge = quartile ? 
                    `<span class="quartile-badge quartile-${quartile.toLowerCase().replace(' ', '-')}">${quartile}</span>` : 
                    '';

                html += `
                    <div class="line-item">
                        <div class="line-item-header">Line Item ${idx + 1} ${quartileBadge}</div>
                        <div class="detail-grid">
                `;
                
                Object.entries(item).forEach(([key, value]) => {
                    if (value !== null && value !== undefined && value.toString().trim() !== '') {
                        html += `
                            <div class="detail-row">
                                <div class="detail-cell detail-label">${key}:</div>
                                <div class="detail-cell detail-value">${value}</div>
                            </div>
                        `;
                    }
                });
                
                html += `</div></div>`;
            });
        }

        html += `</div></div>`;
    });

    return html;
}

// Add these functions to the end of your app.js file

function generateWordFilterSummary() {
    const filters = {
        fund: document.getElementById('fundFilter').value,
        department: document.getElementById('departmentFilter').value,
        division: document.getElementById('divisionFilter').value,
        program: document.getElementById('programFilter').value,
        requestType: document.getElementById('requestTypeFilter').value,
        status: document.getElementById('statusFilter').value
    };

    const quartileStats = {
        'Most Aligned': 0,
        'More Aligned': 0,
        'Less Aligned': 0,
        'Least Aligned': 0
    };
    
    let totalOngoing = 0;
    let totalOnetime = 0;
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartileStats.hasOwnProperty(quartile)) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartileStats[quartile] += lineItemAmount.total;
            }
        });
    });

    return `
        <div class="section-header" id="report-summary">Report Summary</div>
        
        <div class="card">
            <div class="card-header">Applied Filters</div>
            <div class="card-body">
                <div class="detail-grid">
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Fund:</div>
                        <div class="detail-cell detail-value">${filters.fund}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Department:</div>
                        <div class="detail-cell detail-value">${filters.department}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Division:</div>
                        <div class="detail-cell detail-value">${filters.division}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Program:</div>
                        <div class="detail-cell detail-value">${filters.program}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Request Type:</div>
                        <div class="detail-cell detail-value">${filters.requestType}</div>
                    </div>
                    <div class="detail-row">
                        <div class="detail-cell detail-label">Status:</div>
                        <div class="detail-cell detail-value">${filters.status}</div>
                    </div>
                </div>
            </div>
        </div>

        <div class="summary-stats">
            <div class="stats-grid">
                <div class="stats-row">
                    <div class="stats-cell">
                        <span class="stats-value">${filteredData.length}</span>
                        <span class="stats-label">Total Requests</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(totalOngoing)}</span>
                        <span class="stats-label">Ongoing</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(totalOnetime)}</span>
                        <span class="stats-label">One-time</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(totalOngoing + totalOnetime)}</span>
                        <span class="stats-label">Total Amount</span>
                    </div>
                </div>
            </div>
            
            <h4 style="color: #667eea; margin: 20px 0 10px 0;">Quartile Distribution</h4>
            <div class="stats-grid">
                <div class="stats-row">
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(quartileStats['Most Aligned'])}</span>
                        <span class="stats-label">Most Aligned</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(quartileStats['More Aligned'])}</span>
                        <span class="stats-label">More Aligned</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(quartileStats['Less Aligned'])}</span>
                        <span class="stats-label">Less Aligned</span>
                    </div>
                    <div class="stats-cell">
                        <span class="stats-value amount">${formatCurrency(quartileStats['Least Aligned'])}</span>
                        <span class="stats-label">Least Aligned</span>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// Enhanced Word document functions - replace these in your app.js

function generateWordVisualAnalysis() {
    // Calculate quartile distribution and departments
    const quartiles = {
        'Most Aligned': 0,
        'More Aligned': 0,
        'Less Aligned': 0,
        'Least Aligned': 0
    };
    const departments = {};

    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            // Use ACTUAL line item cost
            const lineItemAmount = getLineItemAmount(item);
            
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartiles.hasOwnProperty(quartile)) {
                quartiles[quartile] += lineItemAmount.total;
            }

            const dept = getPrimaryValue([item], 'department');
            if (dept) {
                departments[dept] = (departments[dept] || 0) + lineItemAmount.total;
            }
        });
    });

    // Create ASCII bar charts for Word
    const maxDeptAmount = Math.max(...Object.values(departments));
    const maxQuartileAmount = Math.max(...Object.values(quartiles));

    let html = `
        <div class="section-break"></div>
        <div class="section-header" id="visual-analysis">Visual Analysis</div>
        
        <div class="card">
            <div class="card-header">Budget Requests by Department</div>
            <div class="card-body">
                <table style="width: 100%; margin: 20px 0;">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Department</th>
                            <th style="width: 50%;">Visual Distribution</th>
                            <th style="width: 20%; text-align: right;">Amount</th>
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    Object.entries(departments).forEach(([dept, amount]) => {
        const percentage = (amount / maxDeptAmount) * 100;
        const barLength = Math.round(percentage / 5); // Scale to reasonable length
        const bar = '█'.repeat(barLength) + '░'.repeat(20 - barLength);
        
        html += `
            <tr>
                <td>${dept}</td>
                <td style="font-family: monospace; font-size: 14px; color: #667eea;">${bar} ${Math.round(percentage)}%</td>
                <td style="text-align: right;" class="amount">$${formatCurrency(amount)}</td>
            </tr>
        `;
    });
    
    html += `
                    </tbody>
                </table>
            </div>
        </div>

        <div class="card">
            <div class="card-header">Budget Requests by Quartile Alignment</div>
            <div class="card-body">
                <table style="width: 100%; margin: 20px 0;">
                    <thead>
                        <tr>
                            <th style="width: 30%;">Quartile</th>
                            <th style="width: 50%;">Visual Distribution</th>
                            <th style="width: 20%; text-align: right;">Amount</th>
                        </tr>
                    </thead>
                    <tbody>
    `;
    
    const quartileColors = {
        'Most Aligned': '#28a745',
        'More Aligned': '#17a2b8', 
        'Less Aligned': '#ffc107',
        'Least Aligned': '#dc3545'
    };

    Object.entries(quartiles).forEach(([quartile, amount]) => {
        const percentage = maxQuartileAmount > 0 ? (amount / maxQuartileAmount) * 100 : 0;
        const barLength = Math.round(percentage / 5);
        const bar = '█'.repeat(barLength) + '░'.repeat(20 - barLength);
        const badgeClass = quartile.toLowerCase().replace(' ', '-');
        
        html += `
            <tr>
                <td><span class="quartile-badge quartile-${badgeClass}">${quartile}</span></td>
                <td style="font-family: monospace; font-size: 14px; color: ${quartileColors[quartile]};">${bar} ${Math.round(percentage)}%</td>
                <td style="text-align: right;" class="amount">$${formatCurrency(amount)}</td>
            </tr>
        `;
    });
    
    html += `
                    </tbody>
                </table>
            </div>
        </div>
    `;

    return html;
}

function generateWordTableOfContents() {
    let html = `
        <div class="section-header" id="table-of-contents">Table of Contents</div>
        <div class="toc">
            <ol>
                <li><a href="#report-summary">Report Summary</a></li>
                <li><a href="#visual-analysis">Visual Analysis</a></li>
                <li><a href="#request-summary-table">Request Summary Table</a></li>
                <li><a href="#department-analysis">Department Analysis</a></li>
                <li><a href="#program-summary">Program Summary</a></li>
                <li><a href="#individual-requests">Individual Budget Requests</a>
                    <ol>
    `;

    filteredData.forEach((request) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        html += `<li><a href="#request-${requestId}">Request ${requestId}: ${description || 'N/A'}</a></li>`;
    });

    html += `
                    </ol>
                </li>
            </ol>
        </div>
    `;

    return html;
}

function generateWordRequestTable() {
    let html = `
        <div class="section-break"></div>
        <div class="section-header" id="request-summary-table">Request Summary Table</div>
        <table style="width: 100%; font-size: 0.85rem; margin: 10px 0;">
            <thead>
                <tr style="background: #667eea; color: white;">
                    <th style="padding: 8px 6px;">ID</th>
                    <th style="padding: 8px 6px;">Description</th>
                    <th style="padding: 8px 6px;">Dept</th>
                    <th style="padding: 8px 6px;">Program</th>
                    <th style="padding: 8px 6px;">Quartile</th>
                    <th style="padding: 8px 6px; text-align: right;">Amount</th>
                </tr>
            </thead>
            <tbody>
    `;

    filteredData.forEach((request, idx) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryProgram = getPrimaryValue(lineItems, 'program') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);

        // Truncate long descriptions for table
        const shortDesc = description && description.length > 25 ? 
            description.substring(0, 25) + '...' : (description || 'N/A');
        const shortProgram = primaryProgram.length > 20 ? 
            primaryProgram.substring(0, 20) + '...' : primaryProgram;

        const quartileBadge = primaryQuartile !== 'N/A' ? 
            `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}" style="font-size: 0.7rem; padding: 2px 8px;">${primaryQuartile.replace(' Aligned', '')}</span>` : 
            'N/A';

        const rowStyle = idx % 2 === 0 ? 'background: #f8f9ff;' : '';

        html += `
            <tr style="${rowStyle}">
                <td style="padding: 6px 4px;"><strong><a href="#request-${requestId}" style="color: #667eea; text-decoration: none;">${requestId}</a></strong></td>
                <td style="padding: 6px 4px; font-size: 0.8rem;">${shortDesc}</td>
                <td style="padding: 6px 4px;">${primaryDept}</td>
                <td style="padding: 6px 4px; font-size: 0.8rem;">${shortProgram}</td>
                <td style="padding: 6px 4px; text-align: center;">${quartileBadge}</td>
                <td style="padding: 6px 4px; text-align: right; font-weight: 600;" class="amount">$${formatCurrency(amounts.total)}</td>
            </tr>
        `;
    });

    html += '</tbody></table>';
    return html;
}

function generateWordDepartmentSummary() {
    const departments = {};
    
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department');
            if (dept) {
                if (!departments[dept]) {
                    departments[dept] = { 
                        requests: new Set(), 
                        amount: 0,
                        programs: new Set(),
                        quartiles: {
                            'Most Aligned': 0,
                            'More Aligned': 0,
                            'Less Aligned': 0,
                            'Least Aligned': 0
                        }
                    };
                }
                departments[dept].requests.add(requestId);
                departments[dept].amount += amounts.total;
                
                const program = getPrimaryValue([item], 'program');
                if (program) departments[dept].programs.add(program);
                
                const quartile = getPrimaryValue([item], 'quartile');
                if (quartile && departments[dept].quartiles.hasOwnProperty(quartile)) {
                    // Use ACTUAL line item cost
                    const lineItemAmount = getLineItemAmount(item);
                    departments[dept].quartiles[quartile] += lineItemAmount.total;
                }
            }
        });
    });

    let html = `
        <div class="section-break"></div>
        <div class="section-header" id="department-analysis">Department Analysis</div>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(400px, 1fr)); gap: 15px;">`;
    
    Object.entries(departments).forEach(([dept, data]) => {
        html += `
            <div class="card" style="margin: 10px 0; break-inside: avoid;">
                <div class="card-header" style="background: #667eea; color: white; padding: 12px 15px; font-size: 1.1rem;">${dept}</div>
                <div class="card-body" style="padding: 15px;">
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; text-align: center; margin-bottom: 15px;">
                        <div style="background: #f8f9ff; padding: 10px; border-radius: 5px;">
                            <div style="font-size: 1.2rem; font-weight: bold; color: #667eea;">${data.requests.size}</div>
                            <div style="font-size: 0.8rem; color: #666;">Requests</div>
                        </div>
                        <div style="background: #f8f9ff; padding: 10px; border-radius: 5px;">
                            <div style="font-size: 1.2rem; font-weight: bold; color: #667eea;">${data.programs.size}</div>
                            <div style="font-size: 0.8rem; color: #666;">Programs</div>
                        </div>
                        <div style="background: #f8f9ff; padding: 10px; border-radius: 5px;">
                            <div style="font-size: 1.1rem; font-weight: bold; color: #28a745;">$${formatCurrency(data.amount)}</div>
                            <div style="font-size: 0.8rem; color: #666;">Total</div>
                        </div>
                    </div>
                    
                    <h4 style="color: #667eea; margin: 15px 0 8px 0; font-size: 0.9rem;">Quartile Distribution</h4>
                    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 5px; font-size: 0.8rem;">
                        <div style="display: flex; justify-content: space-between; padding: 4px 8px; background: #f0f8f0; border-radius: 3px;">
                            <span>Most Aligned:</span>
                            <span class="amount">$${formatCurrency(data.quartiles['Most Aligned'])}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; padding: 4px 8px; background: #f0f8ff; border-radius: 3px;">
                            <span>More Aligned:</span>
                            <span class="amount">$${formatCurrency(data.quartiles['More Aligned'])}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; padding: 4px 8px; background: #fff8f0; border-radius: 3px;">
                            <span>Less Aligned:</span>
                            <span class="amount">$${formatCurrency(data.quartiles['Less Aligned'])}</span>
                        </div>
                        <div style="display: flex; justify-content: space-between; padding: 4px 8px; background: #fff0f0; border-radius: 3px;">
                            <span>Least Aligned:</span>
                            <span class="amount">$${formatCurrency(data.quartiles['Least Aligned'])}</span>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });

    html += '</div>';
    return html;
}

function generateWordDetailedRequests() {
    let html = `
        <div class="section-break"></div>
        <div class="section-header" id="individual-requests">Individual Budget Requests</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);

        html += `
            <div class="card page-break" id="request-${requestId}" style="margin: 15px 0;">
                <div class="card-header" style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 15px 20px;">
                    <div style="font-size: 1.2rem; font-weight: 600;">Request ${requestId}: ${description}</div>
                </div>
                <div class="card-body" style="padding: 20px;">
                    <!-- Quick Summary Section -->
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px; text-align: center;">
                        <div style="background: #f8f9ff; padding: 12px; border-radius: 8px; border-left: 4px solid #667eea;">
                            <div style="font-size: 0.8rem; color: #666; margin-bottom: 5px;">Request ID</div>
                            <div style="font-size: 1.1rem; font-weight: 600; color: #667eea;">${requestId}</div>
                        </div>
                        <div style="background: #f0f8f0; padding: 12px; border-radius: 8px; border-left: 4px solid #28a745;">
                            <div style="font-size: 0.8rem; color: #666; margin-bottom: 5px;">Total Amount</div>
                            <div style="font-size: 1.1rem; font-weight: 600; color: #28a745;">$${formatCurrency(amounts.total)}</div>
                        </div>
                        <div style="background: #fff8f0; padding: 12px; border-radius: 8px; border-left: 4px solid #ffc107;">
                            <div style="font-size: 0.8rem; color: #666; margin-bottom: 5px;">Line Items</div>
                            <div style="font-size: 1.1rem; font-weight: 600; color: #ffc107;">${lineItems.length}</div>
                        </div>
                    </div>
        `;

        // Add Q&A section - more compact
        if (qa.length > 0) {
            html += `<div style="margin-bottom: 20px;">
                        <h4 style="color: #667eea; margin-bottom: 10px; font-size: 1rem; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Request Details</h4>`;
            
            qa.forEach((qItem, idx) => {
                let question = '';
                let answer = '';
                
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && qItem[key]) {
                        question = qItem[key];
                    }
                    if (lowerKey.includes('answer') && qItem[key]) {
                        answer = qItem[key];
                    }
                });
                
                if (question && answer && answer.trim()) {
                    html += `
                        <div style="margin: 10px 0; padding: 12px 15px; background: #fff8f0; border-radius: 5px; border-left: 3px solid #ffc107;">
                            <div style="font-weight: 600; color: #667eea; font-size: 0.9rem; margin-bottom: 6px;">${question}</div>
                            <div style="line-height: 1.4; font-size: 0.85rem; color: #333;">${answer}</div>
                        </div>
                    `;
                }
            });
            html += '</div>';
        }

        // Add line items - more compact grid layout
        if (lineItems.length > 0) {
            html += `<div style="margin-bottom: 15px;">
                        <h4 style="color: #667eea; margin-bottom: 10px; font-size: 1rem; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Line Items</h4>
                        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 10px;">`;
            
            lineItems.forEach((item, idx) => {
                const quartile = getPrimaryValue([item], 'quartile');
                const quartileBadge = quartile ? 
                    `<span class="quartile-badge quartile-${quartile.toLowerCase().replace(' ', '-')}" style="font-size: 0.7rem; padding: 2px 8px; margin-left: 8px;">${quartile.replace(' Aligned', '')}</span>` : 
                    '';

                html += `
                    <div style="background: #f8f9ff; padding: 12px; border-radius: 5px; border-left: 3px solid #667eea;">
                        <div style="font-weight: 600; font-size: 0.9rem; margin-bottom: 8px; color: #333;">
                            Line Item ${idx + 1}${quartileBadge}
                        </div>
                `;
                
                // Show key fields only
                const keyFields = ['Department', 'Program', 'Position Title', 'Account', 'Description'];
                let shownFields = 0;
                
                Object.entries(item).forEach(([key, value]) => {
                    if (value !== null && value !== undefined && value.toString().trim() !== '' && shownFields < 4) {
                        const isKeyField = keyFields.some(kf => key.toLowerCase().includes(kf.toLowerCase()));
                        if (isKeyField || shownFields < 2) {
                            // Add dollar signs to cost fields
                            let displayValue = value;
                            const lowerKey = key.toLowerCase();
                            if ((lowerKey.includes('onetime') && lowerKey.includes('cost')) ||
                                (lowerKey.includes('ongoing') && lowerKey.includes('cost'))) {
                                // Check if the value is numeric
                                const numValue = parseFloat(value);
                                if (!isNaN(numValue)) {
                                    displayValue = `$${formatCurrency(numValue)}`;
                                }
                            }

                            html += `
                <div style="display: flex; justify-content: space-between; margin: 3px 0; font-size: 0.8rem;">
                    <span style="color: #666; font-weight: 500;">${key}:</span>
                    <span style="color: #333; text-align: right;">${displayValue}</span>
                </div>
            `;
                            shownFields++;
                        }
                    }
                });
                
                html += `</div>`;
            });
            
            html += '</div></div>';
        }

        html += `</div></div>`;
    });

    return html;
}

function downloadPdfReport() {
    if (filteredData.length === 0) {
        alert('Please generate a report first.');
        return;
    }
    
    const reportDate = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    let totalAmount = 0, totalOngoing = 0, totalOnetime = 0;
    const quartileStats = { 'Most Aligned': 0, 'More Aligned': 0, 'Less Aligned': 0, 'Least Aligned': 0 };
    const deptStats = {};
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalAmount += amounts.total;
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
        
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const dept = getPrimaryValue(lineItems, 'department') || 'Unknown';
        
        if (!deptStats[dept]) deptStats[dept] = { count: 0, amount: 0 };
        deptStats[dept].count++;
        deptStats[dept].amount += amounts.total;
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartileStats.hasOwnProperty(quartile)) {
                // Use ACTUAL line item cost
                const lineItemAmount = getLineItemAmount(item);
                quartileStats[quartile] += lineItemAmount.total;
            }
        });
    });
    
    // Build request table rows
    let tableRows = '';
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);
        const shortDesc = description && description.length > 40 ? description.substring(0, 40) + '...' : (description || 'N/A');
        
        const qBadge = primaryQuartile.includes('Most') || primaryQuartile.includes('More') ? 'badge-high' : 'badge-low';
        
        tableRows += `<tr>
            <td>${requestId}</td>
            <td>${shortDesc}</td>
            <td>${primaryDept}</td>
            <td><span class="badge ${qBadge}">${primaryQuartile}</span></td>
            <td class="amount">$${formatCurrency(amounts.total)}</td>
        </tr>`;
    });
    
    // Build department summary
    let deptRows = '';
    Object.entries(deptStats).sort((a, b) => b[1].amount - a[1].amount).forEach(([dept, stats]) => {
        deptRows += `<tr>
            <td>${dept}</td>
            <td style="text-align: center;">${stats.count}</td>
            <td class="amount">$${formatCurrency(stats.amount)}</td>
            <td style="text-align: center;">${((stats.amount / totalAmount) * 100).toFixed(1)}%</td>
        </tr>`;
    });
    
    // ===== BUILD PROGRAM SUMMARY DATA =====
    const programData = {};
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        const amounts = getRequestAmount(request);
        
        lineItems.forEach(item => {
            const dept = getPrimaryValue([item], 'department') || 'Unknown Department';
            const program = getPrimaryValue([item], 'program') || 'Unknown Program';
            const quartile = getPrimaryValue([item], 'quartile') || 'N/A';
            
            if (!programData[dept]) programData[dept] = {};
            if (!programData[dept][program]) {
                programData[dept][program] = { quartile: quartile, totalCost: 0, requestedAmount: 0, proposedTotalCost: 0 };
            }
            // Use ACTUAL line item cost (not divided total)
            const lineItemAmount = getLineItemAmount(item);
            programData[dept][program].requestedAmount += lineItemAmount.total;
            // Get current budget from uploaded Program Inventory data
            if (programData[dept][program].totalCost === 0) {
                const currentBudget = getCurrentBudgetForProgram(dept, program);
                if (currentBudget) {
                    programData[dept][program].totalCost = currentBudget.totalCost;
                }
                // If no current budget found, leave as 0 (new program)
            }
            programData[dept][program].proposedTotalCost = programData[dept][program].totalCost + programData[dept][program].requestedAmount;
        });
    });
    
    // Build Program Summary HTML
    let programSummaryHtml = '';
    Object.entries(programData).forEach(([dept, programs]) => {
        let deptTotal = { totalCost: 0, requestedAmount: 0, proposedTotalCost: 0 };
        let programRows = '';
        
        const sortedPrograms = Object.entries(programs).sort((a, b) => {
            const quartileOrder = { 'Most Aligned': 1, 'More Aligned': 2, 'Less Aligned': 3, 'Least Aligned': 4 };
            return (quartileOrder[a[1].quartile] || 5) - (quartileOrder[b[1].quartile] || 5);
        });
        
        sortedPrograms.forEach(([program, data]) => {
            deptTotal.totalCost += data.totalCost;
            deptTotal.requestedAmount += data.requestedAmount;
            deptTotal.proposedTotalCost += data.proposedTotalCost;
            
            const qClass = data.quartile.includes('Most') ? 'badge-q1' : data.quartile.includes('More') ? 'badge-q2' : data.quartile.includes('Less') ? 'badge-q3' : 'badge-q4';
            programRows += `<tr>
                <td><span class="badge ${qClass}">${data.quartile}</span></td>
                <td>${program}</td>
                <td style="text-align: right;">$${formatCurrency(Math.round(data.totalCost))}</td>
                <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(data.requestedAmount))}</td>
                <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(data.proposedTotalCost))}</td>
            </tr>`;
        });
        
        programSummaryHtml += `
            <div class="dept-card">
                <div class="dept-header">${dept}</div>
                <div class="dept-body">
                    <table class="data-table">
                        <thead><tr><th>Quartile</th><th>Program</th><th style="text-align: right;">Current Cost</th><th style="text-align: right;">Requested</th><th style="text-align: right;">Proposed Total</th></tr></thead>
                        <tbody>
                            ${programRows}
                            <tr class="total-row">
                                <td colspan="2"><strong>${dept} Total</strong></td>
                                <td style="text-align: right;"><strong>$${formatCurrency(Math.round(deptTotal.totalCost))}</strong></td>
                                <td style="text-align: right;" class="amount"><strong>$${formatCurrency(Math.round(deptTotal.requestedAmount))}</strong></td>
                                <td style="text-align: right;" class="amount"><strong>$${formatCurrency(Math.round(deptTotal.proposedTotalCost))}</strong></td>
                            </tr>
                        </tbody>
                    </table>
                    <div class="impact-note">
                        <strong>Impact:</strong> ${Object.keys(programs).length} programs requesting $${formatCurrency(Math.round(deptTotal.requestedAmount))}, 
                        increasing budget from $${formatCurrency(Math.round(deptTotal.totalCost))} to $${formatCurrency(Math.round(deptTotal.proposedTotalCost))} 
                        (${((deptTotal.requestedAmount / deptTotal.totalCost) * 100).toFixed(1)}% increase)
                    </div>
                </div>
            </div>
        `;
    });
    
    // ===== BUILD DETAILED REQUEST PAGES =====
    let detailedRequestsHtml = '';
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        
        // Q&A Section
        let qaHtml = '';
        if (qa.length > 0) {
            qa.forEach(qItem => {
                let question = '', answer = '';
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && !lowerKey.includes('type') && qItem[key]) question = qItem[key];
                    if (lowerKey.includes('answer') && qItem[key]) answer = qItem[key];
                });
                if (!question) {
                    const questionKeys = ['Question', 'C', 'Col_2', 'Col_C'];
                    for (const key of questionKeys) {
                        if (qItem[key] && qItem[key].toString().trim()) { question = qItem[key]; break; }
                    }
                }
                if (question && answer && answer.trim()) {
                    qaHtml += `<div class="qa-item"><div class="qa-question">${question}</div><div class="qa-answer">${answer}</div></div>`;
                }
            });
        }
        
        // Line Items Section
        let lineItemsHtml = '';
        lineItems.forEach((item, idx) => {
            const itemQuartile = getPrimaryValue([item], 'quartile');
            const qClass = itemQuartile && (itemQuartile.includes('Most') || itemQuartile.includes('More')) ? 'badge-high' : 'badge-low';
            
            // Build field grids
            let fieldsHtml = '<div class="field-grid">';
            
            // Row 1 - Basic Info
            const basicFields = ['REQUESTID', 'REQUEST DESCRIPTION', 'REQUEST TYPE', 'STATUS', 'ONGOING COST'];
            basicFields.forEach(field => {
                const value = findFieldValue(item, field);
                if (value !== null) {
                    const displayValue = formatFieldValue(field, value);
                    fieldsHtml += `<div class="field-item"><div class="field-label">${field}</div><div class="field-value">${displayValue}</div></div>`;
                }
            });
            fieldsHtml += '</div><div class="field-grid">';
            
            // Row 2 - Financial
            const financialFields = ['ONETIME COST', 'NUMBEROFITEMS', 'COST CENTER', 'ACCTTYPE', 'ACCTCODE'];
            financialFields.forEach(field => {
                const value = findFieldValue(item, field);
                if (value !== null) {
                    const displayValue = formatFieldValue(field, value);
                    fieldsHtml += `<div class="field-item"><div class="field-label">${field}</div><div class="field-value">${displayValue}</div></div>`;
                }
            });
            fieldsHtml += '</div><div class="field-grid">';
            
            // Row 3 - Organizational
            const orgFields = ['FUND', 'DEPARTMENT', 'ACCOUNT CATEGORY', 'PROGRAM', 'PROGRAMID'];
            orgFields.forEach(field => {
                const value = findFieldValue(item, field);
                if (value !== null) {
                    fieldsHtml += `<div class="field-item"><div class="field-label">${field}</div><div class="field-value">${value}</div></div>`;
                }
            });
            fieldsHtml += '</div><div class="field-grid scoring">';
            
            // Row 4 - Scoring Criteria
            const scoringFields = ['CHANGE IN DEMAND FOR THE PROGRAM', 'MANDATED TO PROVIDE PROGRAM', 'RELIANCE ON ORGANIZATION TO PROVIDE PROGRAM', 'PORTION OF THE COMMUNITY SERVED'];
            scoringFields.forEach(field => {
                const value = findFieldValue(item, field);
                if (value !== null) {
                    fieldsHtml += `<div class="field-item scoring"><div class="field-label">${field}</div><div class="field-value">${value}</div></div>`;
                }
            });
            fieldsHtml += '</div>';
            
            // Row 5 - Additional
            const additionalFields = ['QUARTILE', 'COST RECOVERY OF PROGRAM'];
            const foundAdditional = additionalFields.filter(f => findFieldValue(item, f) !== null);
            if (foundAdditional.length > 0) {
                fieldsHtml += '<div class="field-grid additional">';
                foundAdditional.forEach(field => {
                    const value = findFieldValue(item, field);
                    fieldsHtml += `<div class="field-item additional"><div class="field-label">${field}</div><div class="field-value">${value}</div></div>`;
                });
                fieldsHtml += '</div>';
            }
            
            lineItemsHtml += `
                <div class="line-item-card">
                    <div class="line-item-header">Line Item ${idx + 1} ${itemQuartile ? `<span class="badge ${qClass}">${itemQuartile}</span>` : ''}</div>
                    ${fieldsHtml}
                </div>
            `;
        });
        
        detailedRequestsHtml += `
            <div class="request-detail-page page-break">
                <div class="page-header"><span class="page-title">Request Detail</span><span class="page-number">Request ${index + 1} of ${filteredData.length}</span></div>
                <div class="request-card">
                    <div class="request-header">Request ${requestId}: ${description || 'No Description'}</div>
                    <div class="request-body">
                        <div class="request-meta-grid">
                            <div class="meta-item"><div class="meta-label">Request ID</div><div class="meta-value">${requestId}</div></div>
                            <div class="meta-item"><div class="meta-label">Total Amount</div><div class="meta-value amount">$${formatCurrency(amounts.total)}</div></div>
                            <div class="meta-item"><div class="meta-label">Department</div><div class="meta-value">${primaryDept}</div></div>
                            <div class="meta-item"><div class="meta-label">Quartile</div><div class="meta-value">${primaryQuartile}</div></div>
                            <div class="meta-item"><div class="meta-label">Line Items</div><div class="meta-value">${lineItems.length}</div></div>
                            <div class="meta-item"><div class="meta-label">Ongoing</div><div class="meta-value">$${formatCurrency(amounts.ongoing)}</div></div>
                        </div>
                        ${qaHtml ? `<div class="qa-section"><h4>Request Context & Details</h4>${qaHtml}</div>` : ''}
                        <div class="line-items-section"><h4>Line Item Details</h4>${lineItemsHtml}</div>
                    </div>
                </div>
            </div>
        `;
    });
    
    const pdfHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Priority Based Budgeting Report</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
@page { size: A4; margin: 0.5in; }
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Inter', sans-serif; line-height: 1.5; color: #1e293b; background: white; font-size: 11px; }

/* Cover Page */
.cover-page { height: 100vh; display: flex; flex-direction: column; background: linear-gradient(135deg, #1e3a5f 0%, #2a4a73 50%, #1e3a5f 100%); color: white; page-break-after: always; }
.cover-header { padding: 40px 60px; }
.cover-brand { font-size: 14px; font-weight: 500; letter-spacing: 0.05em; opacity: 0.9; }
.cover-main { flex: 1; display: flex; flex-direction: column; justify-content: center; padding: 0 60px; }
.cover-title { font-size: 44px; font-weight: 700; line-height: 1.1; margin-bottom: 20px; }
.cover-subtitle { font-size: 22px; font-weight: 300; opacity: 0.9; margin-bottom: 40px; }
.cover-stats { display: flex; gap: 40px; margin-top: 30px; }
.cover-stat-value { font-size: 32px; font-weight: 700; color: #10b981; }
.cover-stat-label { font-size: 13px; opacity: 0.8; text-transform: uppercase; letter-spacing: 0.05em; }
.cover-footer { padding: 40px 60px; border-top: 1px solid rgba(255,255,255,0.1); display: flex; justify-content: space-between; font-size: 12px; opacity: 0.7; }

/* Content Pages */
.content-page { padding: 40px 50px; }
.page-header { display: flex; justify-content: space-between; padding-bottom: 12px; border-bottom: 2px solid #e2e8f0; margin-bottom: 25px; }
.page-title { font-size: 11px; text-transform: uppercase; letter-spacing: 0.1em; color: #64748b; font-weight: 600; }
.page-number { font-size: 11px; color: #64748b; }
.section-title { font-size: 24px; font-weight: 700; color: #1e3a5f; margin-bottom: 6px; }
.section-subtitle { font-size: 14px; color: #64748b; margin-bottom: 20px; }
.page-break { page-break-before: always; }

/* Stats Grid */
.stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 15px; margin-bottom: 30px; }
.stat-card { padding: 20px; border-radius: 10px; text-align: center; background: linear-gradient(135deg, #667eea, #764ba2); color: white; }
.stat-value { font-size: 26px; font-weight: 700; display: block; }
.stat-label { font-size: 11px; opacity: 0.9; margin-top: 4px; }

/* Quartile Grid */
.quartile-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 30px; }
.quartile-card { padding: 16px; border-radius: 8px; text-align: center; }
.quartile-card.q1 { background: linear-gradient(135deg, #d1fae5, #a7f3d0); border: 2px solid #10b981; }
.quartile-card.q2 { background: linear-gradient(135deg, #dbeafe, #bfdbfe); border: 2px solid #3b82f6; }
.quartile-card.q3 { background: linear-gradient(135deg, #fef3c7, #fde68a); border: 2px solid #f59e0b; }
.quartile-card.q4 { background: linear-gradient(135deg, #fee2e2, #fecaca); border: 2px solid #ef4444; }
.quartile-value { font-size: 20px; font-weight: 700; }
.quartile-card.q1 .quartile-value { color: #059669; }
.quartile-card.q2 .quartile-value { color: #2563eb; }
.quartile-card.q3 .quartile-value { color: #d97706; }
.quartile-card.q4 .quartile-value { color: #dc2626; }
.quartile-label { font-size: 10px; font-weight: 600; margin-top: 4px; color: #475569; }

/* Findings Grid */
.findings-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 20px; }
.finding-card { padding: 15px; background: #f8fafc; border-radius: 10px; border-left: 4px solid #1e3a5f; }
.finding-title { font-size: 12px; font-weight: 600; color: #1e3a5f; margin-bottom: 6px; }
.finding-text { font-size: 11px; color: #475569; }

/* Data Tables */
.data-table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 10px; }
.data-table thead { background: #1e3a5f; color: white; }
.data-table th { padding: 10px 8px; text-align: left; font-weight: 600; text-transform: uppercase; font-size: 9px; }
.data-table td { padding: 10px 8px; border-bottom: 1px solid #e2e8f0; }
.data-table tr:nth-child(even) { background: #f8fafc; }
.total-row { background: #e2e8f0 !important; border-top: 2px solid #1e3a5f; }

/* Badges */
.badge { display: inline-block; padding: 3px 8px; border-radius: 12px; font-size: 9px; font-weight: 600; }
.badge-high { background: #10b981; color: white; }
.badge-low { background: #64748b; color: white; }
.badge-q1 { background: #10b981; color: white; }
.badge-q2 { background: #3b82f6; color: white; }
.badge-q3 { background: #f59e0b; color: white; }
.badge-q4 { background: #ef4444; color: white; }
.amount { color: #10b981; font-weight: 600; }

/* Department Cards */
.dept-card { margin-bottom: 25px; border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; page-break-inside: avoid; }
.dept-header { background: linear-gradient(135deg, #1e3a5f, #2a4a73); color: white; padding: 12px 16px; font-weight: 600; font-size: 13px; }
.dept-body { padding: 15px; }
.impact-note { margin-top: 12px; padding: 10px; background: #f0f9ff; border-radius: 6px; font-size: 11px; color: #0369a1; border-left: 3px solid #0ea5e9; }

/* Request Detail Pages */
.request-detail-page { padding: 30px 40px; }
.request-card { border: 1px solid #e2e8f0; border-radius: 10px; overflow: hidden; }
.request-header { background: linear-gradient(135deg, #1e3a5f, #2a4a73); color: white; padding: 15px 20px; font-weight: 600; font-size: 14px; }
.request-body { padding: 20px; }
.request-meta-grid { display: grid; grid-template-columns: repeat(6, 1fr); gap: 10px; margin-bottom: 20px; }
.meta-item { background: #f8fafc; padding: 10px; border-radius: 6px; text-align: center; border: 1px solid #e2e8f0; }
.meta-label { font-size: 9px; color: #64748b; font-weight: 600; margin-bottom: 3px; }
.meta-value { font-size: 12px; color: #1e293b; font-weight: 500; }

/* Q&A Section */
.qa-section { margin: 20px 0; }
.qa-section h4 { color: #1e3a5f; font-size: 13px; margin-bottom: 12px; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }
.qa-item { background: #fffbeb; border-left: 4px solid #f59e0b; padding: 12px; margin-bottom: 10px; border-radius: 0 6px 6px 0; }
.qa-question { font-weight: 600; color: #1e3a5f; font-size: 11px; margin-bottom: 6px; }
.qa-answer { color: #475569; font-size: 11px; line-height: 1.5; }

/* Line Items Section */
.line-items-section { margin-top: 20px; }
.line-items-section h4 { color: #1e3a5f; font-size: 13px; margin-bottom: 12px; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }
.line-item-card { margin-bottom: 15px; padding: 12px; background: #f8fafc; border-radius: 8px; border-left: 4px solid #667eea; page-break-inside: avoid; }
.line-item-header { font-weight: 600; font-size: 12px; color: #1e3a5f; margin-bottom: 10px; display: flex; align-items: center; gap: 10px; }
.field-grid { display: grid; grid-template-columns: repeat(5, 1fr); gap: 6px; margin-bottom: 8px; }
.field-item { background: white; padding: 6px; border-radius: 4px; border: 1px solid #e2e8f0; text-align: center; }
.field-item.scoring { background: #fffbeb; border-color: #f59e0b; }
.field-item.additional { background: #f0fdf4; border-color: #10b981; }
.field-label { font-size: 7px; color: #64748b; font-weight: 600; margin-bottom: 2px; text-transform: uppercase; }
.field-value { font-size: 9px; color: #1e293b; font-weight: 500; word-break: break-word; }

@media print { 
    .page-break { page-break-before: always; } 
    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    .dept-card { break-inside: avoid; }
    .line-item-card { break-inside: avoid; }
    .qa-item { break-inside: avoid; }
}
</style></head><body>

<!-- COVER PAGE -->
<div class="cover-page">
    <div class="cover-header"><span class="cover-brand">TYLER TECHNOLOGIES • BUDGET INTELLIGENCE</span></div>
    <div class="cover-main">
        <h1 class="cover-title">Priority Based Budgeting<br>Report</h1>
        <p class="cover-subtitle">Comprehensive Budget Request Analysis</p>
        <div class="cover-stats">
            <div class="cover-stat"><div class="cover-stat-value">${filteredData.length}</div><div class="cover-stat-label">Budget Requests</div></div>
            <div class="cover-stat"><div class="cover-stat-value">$${formatCurrency(totalAmount)}</div><div class="cover-stat-label">Total Amount</div></div>
            <div class="cover-stat"><div class="cover-stat-value">${Object.keys(deptStats).length}</div><div class="cover-stat-label">Departments</div></div>
        </div>
    </div>
    <div class="cover-footer"><span>Generated on ${reportDate}</span><span>Confidential • For Internal Use Only</span></div>
</div>

<!-- EXECUTIVE SUMMARY -->
<div class="content-page page-break">
    <div class="page-header"><span class="page-title">Executive Summary</span></div>
    <h2 class="section-title">Budget Request Overview</h2>
    <p class="section-subtitle">Analysis of ${filteredData.length} budget requests totaling $${formatCurrency(totalAmount)}</p>
    
    <div class="stats-grid">
        <div class="stat-card"><span class="stat-value">${filteredData.length}</span><div class="stat-label">Total Requests</div></div>
        <div class="stat-card"><span class="stat-value">$${formatCurrency(totalOngoing)}</span><div class="stat-label">Ongoing</div></div>
        <div class="stat-card"><span class="stat-value">$${formatCurrency(totalOnetime)}</span><div class="stat-label">One-time</div></div>
        <div class="stat-card"><span class="stat-value">$${formatCurrency(totalAmount)}</span><div class="stat-label">Total Amount</div></div>
    </div>
    
    <h3 style="font-size: 16px; color: #1e3a5f; margin: 25px 0 15px;">Quartile Distribution</h3>
    <div class="quartile-grid">
        <div class="quartile-card q1"><div class="quartile-value">$${formatCurrency(quartileStats['Most Aligned'])}</div><div class="quartile-label">Most Aligned (Q1)</div></div>
        <div class="quartile-card q2"><div class="quartile-value">$${formatCurrency(quartileStats['More Aligned'])}</div><div class="quartile-label">More Aligned (Q2)</div></div>
        <div class="quartile-card q3"><div class="quartile-value">$${formatCurrency(quartileStats['Less Aligned'])}</div><div class="quartile-label">Less Aligned (Q3)</div></div>
        <div class="quartile-card q4"><div class="quartile-value">$${formatCurrency(quartileStats['Least Aligned'])}</div><div class="quartile-label">Least Aligned (Q4)</div></div>
    </div>
    
    <h3 style="font-size: 16px; color: #1e3a5f; margin: 25px 0 15px;">Key Findings</h3>
    <div class="findings-grid">
        <div class="finding-card"><div class="finding-title">High Priority Requests</div><div class="finding-text">$${formatCurrency(quartileStats['Most Aligned'] + quartileStats['More Aligned'])} (${totalAmount > 0 ? Math.round((quartileStats['Most Aligned'] + quartileStats['More Aligned']) / totalAmount * 100) : 0}%) in Q1/Q2 aligned programs</div></div>
        <div class="finding-card"><div class="finding-title">Funding Mix</div><div class="finding-text">Ongoing: $${formatCurrency(totalOngoing)} | One-time: $${formatCurrency(totalOnetime)}</div></div>
        <div class="finding-card"><div class="finding-title">Department Coverage</div><div class="finding-text">${Object.keys(deptStats).length} departments with budget requests submitted</div></div>
        <div class="finding-card"><div class="finding-title">Average Request</div><div class="finding-text">$${formatCurrency(Math.round(totalAmount / filteredData.length))} per request</div></div>
    </div>
</div>

<!-- DEPARTMENT SUMMARY -->
<div class="content-page page-break">
    <div class="page-header"><span class="page-title">Department Summary</span></div>
    <h2 class="section-title">Requests by Department</h2>
    <p class="section-subtitle">Budget request distribution across organizational units</p>
    <table class="data-table">
        <thead><tr><th>Department</th><th style="text-align: center;">Requests</th><th>Total Amount</th><th style="text-align: center;">% of Total</th></tr></thead>
        <tbody>${deptRows}</tbody>
    </table>
</div>

<!-- REQUEST SUMMARY TABLE -->
<div class="content-page page-break">
    <div class="page-header"><span class="page-title">Request Summary</span></div>
    <h2 class="section-title">All Budget Requests</h2>
    <p class="section-subtitle">Complete listing of all ${filteredData.length} budget requests</p>
    <table class="data-table">
        <thead><tr><th>ID</th><th>Description</th><th>Department</th><th>Quartile</th><th style="text-align: right;">Amount</th></tr></thead>
        <tbody>${tableRows}</tbody>
    </table>
</div>

<!-- PROGRAM SUMMARY -->
<div class="content-page page-break">
    <div class="page-header"><span class="page-title">Program Summary</span></div>
    <h2 class="section-title">Programs by Department</h2>
    <p class="section-subtitle">Program-level budget analysis showing current costs, requested amounts, and proposed totals</p>
    ${programSummaryHtml}
</div>

<!-- DETAILED REQUEST PAGES -->
${detailedRequestsHtml}

</body></html>`;
    
    const newWindow = window.open('', '_blank');
    newWindow.document.write(pdfHtml);
    newWindow.document.close();
    newWindow.focus();
    setTimeout(() => alert('Comprehensive PDF Report opened!\\n\\nTo save as PDF:\\n1. Press Ctrl+P (Cmd+P on Mac)\\n2. Select "Save as PDF"\\n3. Click Save'), 500);
}

function captureChartAsImage(chartId) {
    try {
        const canvas = document.getElementById(chartId);
        if (canvas && canvas.getContext) {
            return canvas.toDataURL('image/png', 1.0);
        }
    } catch (error) {
        console.error(`Error capturing chart ${chartId}:`, error);
    }
    return 'data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iNDAwIiBoZWlnaHQ9IjIwMCIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj48cmVjdCB3aWR0aD0iMTAwJSIgaGVpZ2h0PSIxMDAlIiBmaWxsPSIjZjhmOWZmIiBzdHJva2U9IiNlMGUwZTAiLz48dGV4dCB4PSI1MCUiIHk9IjUwJSIgZG9taW5hbnQtYmFzZWxpbmU9Im1pZGRsZSIgdGV4dC1hbmNob3I9Im1pZGRsZSIgZmlsbD0iIzY2N2VlYSI+Q2hhcnQgUGxhY2Vob2xkZXI8L3RleHQ+PC9zdmc+';
}

function generatePDFSummaryStats() {
    let totalOngoing = 0;
    let totalOnetime = 0;
    
    filteredData.forEach(request => {
        const amounts = getRequestAmount(request);
        totalOngoing += amounts.ongoing;
        totalOnetime += amounts.onetime;
    });

    return `
        <div class="stats-grid">
            <div class="stats-card">
                <span class="stats-value">${filteredData.length}</span>
                <div class="stats-label">Total Requests</div>
            </div>
            <div class="stats-card">
                <span class="stats-value">$${formatCurrency(totalOngoing)}</span>
                <div class="stats-label">Ongoing</div>
            </div>
            <div class="stats-card">
                <span class="stats-value">$${formatCurrency(totalOnetime)}</span>
                <div class="stats-label">One-time</div>
            </div>
            <div class="stats-card">
                <span class="stats-value">$${formatCurrency(totalOngoing + totalOnetime)}</span>
                <div class="stats-label">Total Amount</div>
            </div>
        </div>
    `;
}

function generatePDFRequestTable() {
    let html = `
        <div class="section-header">Request Summary</div>
        <table>
            <thead>
                <tr>
                    <th>ID</th><th>Description</th><th>Department</th><th>Quartile</th><th>Amount</th>
                </tr>
            </thead>
            <tbody>
    `;

    filteredData.forEach((request) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);

        const shortDesc = description && description.length > 25 ? 
            description.substring(0, 25) + '...' : (description || 'N/A');

        const quartileBadge = primaryQuartile !== 'N/A' ? 
            `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}">${primaryQuartile.replace(' Aligned', '')}</span>` : 'N/A';

        html += `
            <tr>
                <td><strong>${requestId}</strong></td>
                <td>${shortDesc}</td>
                <td>${primaryDept}</td>
                <td>${quartileBadge}</td>
                <td class="amount">$${formatCurrency(amounts.total)}</td>
            </tr>
        `;
    });

    html += '</tbody></table>';
    return html;
}

function generatePDFDetailedRequests() {
    let html = `<div class="section-header">Individual Budget Requests</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);

        html += `
            <div class="page-break" style="margin: 15px 0;">
                <div style="background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 12px 15px; border-radius: 8px 8px 0 0;">
                    <h3 style="margin: 0; font-size: 14px;">Request ID: ${requestId} - ${description}</h3>
                </div>
                <div style="border: 1px solid #e0e0e0; border-top: none; padding: 15px; background: #fafafa;">
                    
                    <!-- Quick Summary Section -->
                    <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 15px;">
                        <div style="background: #f8f9ff; padding: 8px; border-radius: 5px; text-align: center; border-left: 3px solid #667eea;">
                            <div style="font-size: 8px; color: #666; margin-bottom: 3px;">Request ID</div>
                            <div style="font-size: 11px; font-weight: 600; color: #667eea;">${requestId}</div>
                        </div>
                        <div style="background: #f0f8f0; padding: 8px; border-radius: 5px; text-align: center; border-left: 3px solid #28a745;">
                            <div style="font-size: 8px; color: #666; margin-bottom: 3px;">Total Amount</div>
                            <div style="font-size: 11px; font-weight: 600; color: #28a745;">$${formatCurrency(amounts.total)}</div>
                        </div>
                        <div style="background: #fff8f0; padding: 8px; border-radius: 5px; text-align: center; border-left: 3px solid #ffc107;">
                            <div style="font-size: 8px; color: #666; margin-bottom: 3px;">Line Items</div>
                            <div style="font-size: 11px; font-weight: 600; color: #ffc107;">${lineItems.length}</div>
                        </div>
                    </div>
        `;

        // Add Q&A section - Complete details
        if (qa.length > 0) {
            html += `<div style="margin-bottom: 15px;">
                        <h4 style="color: #667eea; margin-bottom: 8px; font-size: 11px; border-bottom: 1px solid #e0e0e0; padding-bottom: 3px;">Request Context & Details</h4>`;
            
            qa.forEach((qItem, idx) => {
                let question = '';
                let answer = '';
                
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    if (lowerKey.includes('question') && qItem[key]) {
                        question = qItem[key];
                    }
                    if (lowerKey.includes('answer') && qItem[key]) {
                        answer = qItem[key];
                    }
                });
                
                if (question && answer && answer.trim()) {
                    html += `
                        <div style="margin: 8px 0; padding: 8px 10px; background: #fff8f0; border-radius: 4px; border-left: 3px solid #ffc107;">
                            <div style="font-weight: 600; color: #667eea; font-size: 9px; margin-bottom: 4px;">${question}</div>
                            <div style="line-height: 1.3; font-size: 8px; color: #333;">${answer}</div>
                        </div>
                    `;
                }
            });
            html += '</div>';
        }

        // Add line items - Complete scoring details with FIXED dollar formatting
        if (lineItems.length > 0) {
            html += `<div style="margin-bottom: 15px;">
                        <h4 style="color: #667eea; margin-bottom: 8px; font-size: 11px; border-bottom: 1px solid #e0e0e0; padding-bottom: 3px;">Line Item Details</h4>`;
            
            lineItems.forEach((item, idx) => {
                const quartile = getPrimaryValue([item], 'quartile');
                const quartileBadge = quartile ? 
                    `<span class="quartile-badge quartile-${quartile.toLowerCase().replace(' ', '-')}" style="font-size: 7px; padding: 2px 6px; margin-left: 6px;">${quartile}</span>` : 
                    '';

                html += `
                    <div style="background: #f8f9ff; padding: 10px; border-radius: 4px; border-left: 3px solid #667eea; margin: 8px 0; page-break-inside: avoid;">
                        <div style="font-weight: 600; font-size: 9px; margin-bottom: 6px; color: #333;">
                            Line Item ${idx + 1}${quartileBadge}
                        </div>
                        
                        <!-- Basic Info Grid -->
                        <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; margin-bottom: 8px;">
                `;
                
                // First row - Basic Info
                const basicFields = ['REQUESTID', 'REQUEST DESCRIPTION', 'REQUEST TYPE', 'STATUS', 'ONGOING COST'];
                basicFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        const displayValue = formatFieldValue(field, value); // ADD THIS LINE
                        html += `
                            <div style="background: white; padding: 4px; border-radius: 3px; text-align: center;">
                                <div style="font-size: 6px; color: #666; font-weight: 600;">${field}</div>
                                <div style="font-size: 8px; color: #333; margin-top: 2px;">${displayValue}</div> <!-- CHANGE FROM ${value} TO ${displayValue} -->
                            </div>
                        `;
                    }
                });

                // Second row - Financial
                const financialFields = ['ONETIME COST', 'NUMBEROFITEMS', 'COST CENTER', 'ACCTTYPE', 'ACCTCODE'];
                financialFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        const displayValue = formatFieldValue(field, value); // ADD THIS LINE
                        html += `
                            <div style="background: white; padding: 4px; border-radius: 3px; text-align: center;">
                                <div style="font-size: 6px; color: #666; font-weight: 600;">${field}</div>
                                <div style="font-size: 8px; color: #333; margin-top: 2px;">${displayValue}</div> <!-- CHANGE FROM ${value} TO ${displayValue} -->
                            </div>
                        `;
                    }
                });
                
                
                html += `</div><div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; margin-bottom: 8px;">`;
                
                // Third row - Organizational details (no formatting needed)
                const orgFields = ['FUND', 'DEPARTMENT', 'ACCOUNT CATEGORY', 'PROGRAM', 'PROGRAMID'];
                orgFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        html += `
                            <div style="background: white; padding: 4px; border-radius: 3px; text-align: center;">
                                <div style="font-size: 6px; color: #666; font-weight: 600;">${field}</div>
                                <div style="font-size: 8px; color: #333; margin-top: 2px;">${value}</div>
                            </div>
                        `;
                    }
                });
                
                html += `</div><div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 5px;">`;
                
                // Fourth row - Scoring details (no formatting needed)
                const scoringFields = ['CHANGE IN DEMAND FOR THE PROGRAM', 'MANDATED TO PROVIDE PROGRAM', 'RELIANCE ON ORGANIZATION TO PROVIDE PROGRAM', 'PORTION OF THE COMMUNITY SERVED'];
                scoringFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        html += `
                            <div style="background: #fff8f0; padding: 4px; border-radius: 3px; text-align: center;">
                                <div style="font-size: 6px; color: #666; font-weight: 600;">${field}</div>
                                <div style="font-size: 8px; color: #333; margin-top: 2px;">${value}</div>
                            </div>
                        `;
                    }
                });
                
                html += `</div>`;
                
                // Cost recovery if available
                const costRecovery = findFieldValue(item, 'COST RECOVERY OF PROGRAM');
                if (costRecovery) {
                    html += `
                        <div style="margin-top: 6px; padding: 4px 8px; background: #f0f0f0; border-radius: 3px;">
                            <span style="font-size: 6px; color: #666; font-weight: 600;">COST RECOVERY: </span>
                            <span style="font-size: 8px; color: #333;">${costRecovery}</span>
                        </div>
                    `;
                }
                
                html += `</div>`;
            });
            
            html += '</div>';
        }

        html += `</div></div>`;
    });

    return html;
}

// Helper function to find field values flexibly
function findFieldValue(item, targetField) {
    // Direct match
    if (item[targetField] !== undefined && item[targetField] !== null && item[targetField].toString().trim() !== '') {
        return item[targetField];
    }
    
    // Flexible matching - check if any key contains the target field name
    for (const [key, value] of Object.entries(item)) {
        if (key.toUpperCase().includes(targetField.toUpperCase()) && value !== null && value !== undefined && value.toString().trim() !== '') {
            return value;
        }
    }
    
    return null;
}

function formatFieldValue(field, value) {
    console.log('formatFieldValue called with:', field, value);
    
    const lowerKey = field.toLowerCase();
    // Check for any variation of "ongoing cost" or "onetime cost"
    if (lowerKey.includes('ongoing') && lowerKey.includes('cost')) {
        console.log('Found ongoing cost field:', field);
        const numValue = parseFloat(value);
        if (!isNaN(numValue)) {
            const formatted = `$${formatCurrency(numValue)}`;
            console.log('Formatting', value, 'to', formatted);
            return formatted;
        }
    } else if (lowerKey.includes('onetime') && lowerKey.includes('cost')) {
        console.log('Found onetime cost field:', field);
        const numValue = parseFloat(value);
        if (!isNaN(numValue)) {
            const formatted = `$${formatCurrency(numValue)}`;
            console.log('Formatting', value, 'to', formatted);
            return formatted;
        }
    }
    return value;
}

// ===== COLLAPSIBLE SECTION CONTROLS =====

function toggleCollapsible(sectionId) {
    const content = document.getElementById(sectionId);
    const toggle = document.getElementById(sectionId + '-toggle');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        toggle.classList.remove('expanded');
    } else {
        content.classList.add('expanded');
        toggle.classList.add('expanded');
    }
}

function toggleRequestAccordion(accordionId) {
    const content = document.getElementById(accordionId);
    const arrow = document.getElementById(accordionId + '-arrow');
    
    if (content.classList.contains('expanded')) {
        content.classList.remove('expanded');
        arrow.classList.remove('expanded');
    } else {
        content.classList.add('expanded');
        arrow.classList.add('expanded');
    }
}

function expandAllRequests() {
    document.querySelectorAll('.request-accordion-content').forEach(content => {
        content.classList.add('expanded');
    });
    document.querySelectorAll('.request-accordion-arrow').forEach(arrow => {
        arrow.classList.add('expanded');
    });
}

function collapseAllRequests() {
    document.querySelectorAll('.request-accordion-content').forEach(content => {
        content.classList.remove('expanded');
    });
    document.querySelectorAll('.request-accordion-arrow').forEach(arrow => {
        arrow.classList.remove('expanded');
    });
}

// ===== STANDARD REPORT COLLAPSIBLE CONTROLS =====

function expandAllStandardRequests() {
    // Expand all request accordions
    document.querySelectorAll('[id^="standard-request-accordion-"]').forEach(content => {
        content.classList.add('expanded');
    });
    document.querySelectorAll('[id^="standard-request-accordion-"][id$="-arrow"]').forEach(arrow => {
        arrow.classList.add('expanded');
    });
    
    // Also expand all internal collapsible sections
    document.querySelectorAll('[id^="standard-qa-"], [id^="standard-line-items-"]').forEach(content => {
        content.classList.add('expanded');
    });
    document.querySelectorAll('[id^="standard-qa-"][id$="-toggle"], [id^="standard-line-items-"][id$="-toggle"]').forEach(toggle => {
        toggle.classList.add('expanded');
    });
}

function collapseAllStandardRequests() {
    // Collapse all request accordions
    document.querySelectorAll('[id^="standard-request-accordion-"]').forEach(content => {
        content.classList.remove('expanded');
    });
    document.querySelectorAll('[id^="standard-request-accordion-"][id$="-arrow"]').forEach(arrow => {
        arrow.classList.remove('expanded');
    });
    
    // Also collapse all internal collapsible sections
    document.querySelectorAll('[id^="standard-qa-"], [id^="standard-line-items-"]').forEach(content => {
        content.classList.remove('expanded');
    });
    document.querySelectorAll('[id^="standard-qa-"][id$="-toggle"], [id^="standard-line-items-"][id$="-toggle"]').forEach(toggle => {
        toggle.classList.remove('expanded');
    });
}

// ===== COLLAPSIBLE TABLE OF CONTENTS SECTION CONTROLS =====

// Toggle individual TOC section
function toggleTOCSection(sectionId) {
    const section = document.getElementById(sectionId);
    if (section) {
        section.classList.toggle('collapsed');
    }
}

// Expand all TOC sections
function expandAllTOCSections() {
    document.querySelectorAll('.toc-section').forEach(section => {
        section.classList.remove('collapsed');
    });
}

// Collapse all TOC sections
function collapseAllTOCSections() {
    document.querySelectorAll('.toc-section').forEach(section => {
        section.classList.add('collapsed');
    });
}

// ===== EXCEL EXPORT OF PBB ANALYSIS =====
function exportPBBAnalysisToExcel() {
    console.log('Starting PBB Analysis Excel export...');
    
    if (filteredData.length === 0) {
        alert('No data to export. Please generate a report first.');
        return;
    }

    // Create a new workbook
    const wb = XLSX.utils.book_new();
    
    // Create header row with all columns
    const pbbData = [[
        'Request ID',
        'Description',
        'Department',
        'Program',
        'Quartile',
        'Total Amount',
        'Archetype #',
        'Decision Profile',
        'PBB Total Score (0-12)',
        'PBB Recommendation',
        '1. Program Alignment Score (0-2)',
        '1. Program Alignment Notes',
        '2. Outcome Evidence Score (0-2)',
        '2. Outcome Evidence Notes',
        '3. Funding Strategy Score (0-2)',
        '3. Funding Strategy Notes',
        '4. Mandate/Risk Score (0-2)',
        '4. Mandate/Risk Notes',
        '5. Efficiency/ROI Score (0-2)',
        '5. Efficiency/ROI Notes',
        ...(includeAccessEquity ? ['6. Access Score (0-2)', '6. Access Notes'] : []),
        'Overall Rationale'
    ]];
    
    // Process each request
    filteredData.forEach(request => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryProgram = getPrimaryValue(lineItems, 'program') || 'N/A';
        // Get and normalize quartile for display
        let primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        if (primaryQuartile !== 'N/A') {
            const qStr = primaryQuartile.toString().trim();
            if (qStr === '1' || qStr === 'Q1') primaryQuartile = 'Most Aligned (Q1)';
            else if (qStr === '2' || qStr === 'Q2') primaryQuartile = 'More Aligned (Q2)';
            else if (qStr === '3' || qStr === 'Q3') primaryQuartile = 'Less Aligned (Q3)';
            else if (qStr === '4' || qStr === 'Q4') primaryQuartile = 'Least Aligned (Q4)';
        }
        const amounts = getRequestAmount(request);
        
        // Calculate PBB scores using your existing scoreRequest function
        const analysis = scoreRequest(request);
        
        pbbData.push([
            requestId,
            description,
            primaryDept,
            primaryProgram,
            primaryQuartile,
            amounts.total,
            analysis.archetypeNumber,
            analysis.gridKey,
            analysis.totalScore,
            analysis.disposition,
            analysis.quartileScore,
            analysis.quartileReason,
            analysis.outcomeScore,
            analysis.outcomeReason,
            analysis.fundingScore,
            analysis.fundingReason,
            analysis.mandateScore,
            analysis.mandateReason,
            analysis.efficiencyScore,
            analysis.efficiencyReason,
            ...(includeAccessEquity ? [analysis.accessScore, analysis.accessReason] : []),
            analysis.narrative
        ]);
    });
    
    // Create worksheet from data
    const ws = XLSX.utils.aoa_to_sheet(pbbData);
    
    // Set column widths for readability
    ws['!cols'] = [
        { wch: 12 },  // Request ID
        { wch: 40 },  // Description
        { wch: 20 },  // Department
        { wch: 25 },  // Program
        { wch: 15 },  // Quartile
        { wch: 15 },  // Total Amount
        { wch: 12 },  // Archetype #
        { wch: 28 },  // Decision Profile
        { wch: 12 },  // PBB Score
        { wch: 15 },  // Recommendation
        { wch: 10 },  // Program Alignment Score
        { wch: 60 },  // Program Alignment Notes
        { wch: 10 },  // Outcome Evidence Score
        { wch: 60 },  // Outcome Evidence Notes
        { wch: 10 },  // Funding Strategy Score
        { wch: 60 },  // Funding Strategy Notes
        { wch: 10 },  // Mandate/Risk Score
        { wch: 60 },  // Mandate/Risk Notes
        { wch: 10 },  // Efficiency/ROI Score
        { wch: 60 },  // Efficiency/ROI Notes
        ...(includeAccessEquity ? [{ wch: 10 }, { wch: 60 }] : []),
        { wch: 80 }   // Overall Rationale
    ];
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'PBB Analysis');
    
    // Generate filename with timestamp
    const timestamp = new Date().toISOString().split('T')[0];
    const filename = `PBB_Analysis_${timestamp}.xlsx`;
    
    // Download the file
    XLSX.writeFile(wb, filename);
    
    console.log('Excel export complete!');
    alert(`PBB Analysis exported successfully!\n\nFile: ${filename}\n\nThe Excel file contains detailed scoring for all ${filteredData.length} requests with explicit explanations for each score.`);
}