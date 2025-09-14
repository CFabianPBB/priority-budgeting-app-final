// Global variables
let budgetData = {
    requestSummary: [],
    personnel: [],
    nonPersonnel: [],
    requestQA: [],
    budgetSummary: []
};
let filteredData = [];

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

function setupFilters() {
    console.log('\n=== Setting up filters ===');
    
    const filters = {
        fund: new Set(['all']),
        department: new Set(['all']),
        division: new Set(['all']),
        program: new Set(['all']),
        requestType: new Set(['all']),
        status: new Set(['all'])
    };

    // Collect unique values from all data sources
    console.log('Collecting filter values...');
    
    // From Personnel data
    budgetData.personnel.forEach((item, idx) => {
        if (idx < 3) console.log(`Personnel item ${idx}:`, item);
        
        Object.keys(item).forEach(key => {
            const value = item[key];
            if (value && typeof value === 'string') {
                const lowerKey = key.toLowerCase();
                if (lowerKey.includes('fund')) filters.fund.add(value);
                if (lowerKey.includes('department') || lowerKey.includes('dept')) filters.department.add(value);
                if (lowerKey.includes('division') || lowerKey.includes('div')) filters.division.add(value);
                if (lowerKey.includes('program')) filters.program.add(value);
                if (lowerKey.includes('status')) filters.status.add(value);
            }
        });
    });
    
    // From NonPersonnel data
    budgetData.nonPersonnel.forEach((item, idx) => {
        if (idx < 3) console.log(`NonPersonnel item ${idx}:`, item);
        
        Object.keys(item).forEach(key => {
            const value = item[key];
            if (value && typeof value === 'string') {
                const lowerKey = key.toLowerCase();
                if (lowerKey.includes('fund')) filters.fund.add(value);
                if (lowerKey.includes('department') || lowerKey.includes('dept')) filters.department.add(value);
                if (lowerKey.includes('division') || lowerKey.includes('div')) filters.division.add(value);
                if (lowerKey.includes('program')) filters.program.add(value);
                if (lowerKey.includes('status')) filters.status.add(value);
            }
        });
    });
    
    // From Request Summary
    budgetData.requestSummary.forEach(item => {
        Object.keys(item).forEach(key => {
            const value = item[key];
            if (value && typeof value === 'string') {
                const lowerKey = key.toLowerCase();
                if (lowerKey.includes('type')) filters.requestType.add(value);
                if (lowerKey.includes('status')) filters.status.add(value);
            }
        });
    });

    console.log('Filter values found:', {
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

function getLineItemsForRequest(requestId) {
    // Find Personnel and NonPersonnel items for this request
    const personnel = budgetData.personnel.filter(item => {
        // Look for RequestID in any field that might contain it
        return Object.values(item).some(value => 
            value && value.toString().trim() === requestId.toString().trim()
        );
    });
    
    const nonPersonnel = budgetData.nonPersonnel.filter(item => {
        // Look for RequestID in any field that might contain it
        return Object.values(item).some(value => 
            value && value.toString().trim() === requestId.toString().trim()
        );
    });
    
    console.log(`Request ${requestId}: Found ${personnel.length} personnel + ${nonPersonnel.length} non-personnel items`);
    
    return [...personnel, ...nonPersonnel];
}

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
        
        // Get related personnel and non-personnel data
        const lineItems = getLineItemsForRequest(requestId);

        // Check filters against line items
        if (filters.fund !== 'all') {
            const hasMatchingFund = lineItems.some(item => 
                Object.values(item).some(value => 
                    value && value.toString() === filters.fund
                )
            );
            if (!hasMatchingFund) return false;
        }

        if (filters.department !== 'all') {
            const hasMatchingDept = lineItems.some(item => 
                Object.keys(item).some(key => {
                    const lowerKey = key.toLowerCase();
                    return (lowerKey.includes('department') || lowerKey.includes('dept')) &&
                           item[key] && item[key].toString() === filters.department;
                })
            );
            if (!hasMatchingDept) return false;
        }

        if (filters.division !== 'all') {
            const hasMatchingDiv = lineItems.some(item => 
                Object.keys(item).some(key => {
                    const lowerKey = key.toLowerCase();
                    return (lowerKey.includes('division') || lowerKey.includes('div')) &&
                           item[key] && item[key].toString() === filters.division;
                })
            );
            if (!hasMatchingDiv) return false;
        }

        if (filters.program !== 'all') {
            const hasMatchingProgram = lineItems.some(item => 
                Object.keys(item).some(key => {
                    const lowerKey = key.toLowerCase();
                    return lowerKey.includes('program') &&
                           item[key] && item[key].toString() === filters.program;
                })
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

function getRequestAmount(request) {
    let ongoing = 0;
    let onetime = 0;
    
    Object.keys(request).forEach(key => {
        const lowerKey = key.toLowerCase();
        const value = parseFloat(request[key]) || 0;
        
        if (lowerKey.includes('ongoing')) ongoing += value;
        if (lowerKey.includes('onetime') || lowerKey.includes('one-time')) onetime += value;
    });
    
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
                quartileStats[quartile] += amounts.total / lineItems.length;
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
    console.log('Displaying report...');
    
    const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    document.getElementById('reportDate').textContent = `Generated on ${reportDate}`;

    const totalAmount = filteredData.reduce((sum, request) => {
        const amounts = getRequestAmount(request);
        return sum + amounts.total;
    }, 0);

    let html = `
        <div style="text-align: center; margin-bottom: 30px;">
            <h1 style="color: #333; margin-bottom: 10px;">Priority Based Budgeting Report</h1>
            <p style="color: #666; font-size: 1.1rem;">Budget Request Analysis and Recommendations</p>
            <p style="color: #888;">Generated on ${reportDate}</p>
        </div>

        <div class="section-header">Executive Summary</div>
        <p>This report analyzes ${filteredData.length} budget requests totaling ${formatCurrency(totalAmount)} in requested funding. The requests span multiple departments and programs, with varying levels of alignment to organizational priorities.</p>
    `;

    // Add filter summary
    html += generateFilterSummary();

    // Add actual table of contents with links
    console.log('Generating table of contents...');
    html += generateActualTableOfContents();

    // Add request summary table
    html += generateRequestSummaryTable();

    // Add department summary
    html += generateDepartmentSummary();

    // Add quartile analysis
    html += generateQuartileAnalysis();

    // Add individual request details
    html += generateDetailedRequestReport();

    // Add charts
    html += generateCharts();

    reportContent.innerHTML = html;
    reportSection.style.display = 'block';

    // Add download event listener after report is displayed
    
    // Add BOTH download event listeners
    const downloadWordBtn = document.getElementById('downloadWordBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');

    if (downloadWordBtn) {
        downloadWordBtn.removeEventListener('click', downloadWordReport);
        downloadWordBtn.addEventListener('click', downloadWordReport);
    }

    if (downloadPdfBtn) {
        downloadPdfBtn.removeEventListener('click', downloadPdfReport);
        downloadPdfBtn.addEventListener('click', downloadPdfReport);
    }
    
    // Render charts after HTML is added to DOM
    setTimeout(renderCharts, 100);
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
                quartileStats[quartile] += amounts.total / lineItems.length;
            }
        });
    });

    let html = `
        <div class="section-header" id="report-filters">Report Filters & Summary</div>
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

function getPrimaryValue(lineItems, fieldType) {
    // Look for the field type in line items
    for (const item of lineItems) {
        for (const key of Object.keys(item)) {
            const lowerKey = key.toLowerCase();
            if (lowerKey.includes(fieldType) && item[key]) {
                return item[key];
            }
        }
    }
    return null;
}

function getRequestQA(requestId) {
    // Find Q&A entries for this request
    return budgetData.requestQA.filter(qa => {
        // Look for RequestID match in any field
        return Object.values(qa).some(value => 
            value && value.toString().trim() === requestId.toString().trim()
        );
    });
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
                    departments[dept].quartiles[quartile] += amounts.total / lineItems.length;
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
                quartiles[quartile].amount += amounts.total / lineItems.length; // Distribute amount
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

function generateRequestQASection(qa) {
    if (qa.length === 0) return '';
    
    let html = `
        <div style="margin-bottom: 25px;">
            <h3 style="color: #667eea; margin-bottom: 15px; border-bottom: 1px solid #e0e0e0; padding-bottom: 5px;">Request Context & Details</h3>
    `;
    
    qa.forEach(qItem => {
        // Find question and answer fields
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
                html += `
                    <div class="detail-item">
                        <div class="detail-label">${key}</div>
                        <div class="detail-value">${value}</div>
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
                departments[dept] = (departments[dept] || 0) + (amounts.total / lineItems.length);
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
                quartiles[quartile] += amounts.total / lineItems.length;
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
                quartileStats[quartile] += amounts.total / lineItems.length;
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
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartiles.hasOwnProperty(quartile)) {
                quartiles[quartile] += amounts.total / lineItems.length;
            }

            const dept = getPrimaryValue([item], 'department');
            if (dept) {
                departments[dept] = (departments[dept] || 0) + (amounts.total / lineItems.length);
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
                    departments[dept].quartiles[quartile] += amounts.total / lineItems.length;
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
                            html += `
                                <div style="display: flex; justify-content: space-between; margin: 3px 0; font-size: 0.8rem;">
                                    <span style="color: #666; font-weight: 500;">${key}:</span>
                                    <span style="color: #333; text-align: right;">${value}</span>
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

// PDF DOCUMENT DOWNLOAD - With actual chart images
function downloadPdfReport() {
    // Capture chart images
    const departmentChartImg = captureChartAsImage('departmentChart');
    const quartileChartImg = captureChartAsImage('quartileChart');
    
    const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    const totalAmount = filteredData.reduce((sum, request) => {
        const amounts = getRequestAmount(request);
        return sum + amounts.total;
    }, 0);

    let pdfHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Priority Based Budgeting Report - PDF</title>
            <style>
                @page { size: A4; margin: 0.75in; }
                body { font-family: Arial, sans-serif; line-height: 1.4; color: #333; font-size: 11px; }
                .page-break { page-break-before: always; }
                .header { text-align: center; margin-bottom: 30px; padding-bottom: 15px; border-bottom: 3px solid #667eea; }
                .header h1 { color: #667eea; font-size: 24px; margin-bottom: 8px; }
                .section-header { color: #667eea; font-size: 16px; font-weight: 600; margin: 25px 0 15px 0; border-bottom: 2px solid #e0e0e0; padding-bottom: 8px; }
                .stats-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin: 15px 0; }
                .stats-card { text-align: center; padding: 12px; background: linear-gradient(135deg, #667eea, #764ba2); color: white; border-radius: 8px; }
                .stats-value { font-size: 16px; font-weight: bold; display: block; margin-bottom: 4px; }
                .stats-label { font-size: 9px; }
                .charts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0; }
                .chart-container { text-align: center; margin: 20px 0; page-break-inside: avoid; }
                .chart-image { max-width: 100%; height: 250px; border: 1px solid #e0e0e0; border-radius: 8px; margin: 10px 0; }
                .amount { font-weight: 600; color: #28a745; }
                table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 10px; }
                th { background: #667eea; color: white; padding: 8px 6px; text-align: left; font-weight: 600; }
                td { padding: 6px; border-bottom: 1px solid #ddd; }
                tr:nth-child(even) { background: #f8f9ff; }
                .quartile-badge { display: inline-block; padding: 3px 8px; border-radius: 10px; font-size: 8px; font-weight: 600; color: white; }
                .quartile-most-aligned { background: #28a745; }
                .quartile-more-aligned { background: #17a2b8; }
                .quartile-less-aligned { background: #ffc107; color: black; }
                .quartile-least-aligned { background: #dc3545; }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>Priority Based Budgeting Report</h1>
                <p>Budget Request Analysis and Recommendations</p>
                <p>Generated on ${reportDate}</p>
            </div>

            <div class="section-header">Executive Summary</div>
            <p>This report analyzes <strong>${filteredData.length} budget requests</strong> totaling <strong>$${formatCurrency(totalAmount)}</strong> in requested funding.</p>
            
            ${generatePDFSummaryStats()}
            
            <div class="page-break"></div>
            <div class="section-header">Visual Analysis</div>
            <div class="charts-grid">
                <div class="chart-container">
                    <h4>Budget Requests by Department</h4>
                    <img src="${departmentChartImg}" class="chart-image" alt="Department Chart">
                </div>
                <div class="chart-container">
                    <h4>Budget Requests by Quartile</h4>
                    <img src="${quartileChartImg}" class="chart-image" alt="Quartile Chart">
                </div>
            </div>
            
            <div class="page-break"></div>
            ${generatePDFRequestTable()}

            <div class="page-break"></div>
            ${generatePDFDetailedRequests()}
            
        </body>
        </html>
    `;

    const blob = new Blob([pdfHtml], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Priority_Budgeting_Report_PDF_${new Date().toISOString().split('T')[0]}.html`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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

        // Add line items - Complete scoring details
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
                
                // Show key basic fields
                const basicFields = ['REQUESTID', 'REQUEST DESCRIPTION', 'REQUEST TYPE', 'STATUS', 'ONGOING COST'];
                basicFields.forEach(field => {
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
                
                html += `</div>`;
                
                // Second row - Financial details
                html += `<div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; margin-bottom: 8px;">`;
                
                const financialFields = ['ONETIME COST', 'NUMBEROFITEMS', 'COST CENTER', 'ACCTTYPE', 'ACCTCODE'];
                financialFields.forEach(field => {
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
                
                html += `</div>`;
                
                // Third row - Organizational details
                html += `<div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 5px; margin-bottom: 8px;">`;
                
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
                
                html += `</div>`;
                
                // Fourth row - Scoring details
                html += `<div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 5px;">`;
                
                const scoringFields = ['CHANGE IN DEMAND FOR THE PROGRAM', 'MANDATED TO PROVIDE PROGRAM', 'RELIANCE ON CITY TO PROVIDE PROGRAM', 'PORTION OF THE COMMUNITY SERVED'];
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
    
    // Flexible matching
    for (const [key, value] of Object.entries(item)) {
        if (key.toUpperCase().includes(targetField.toUpperCase()) && value !== null && value !== undefined && value.toString().trim() !== '') {
            return value;
        }
    }
    
    return null;
}