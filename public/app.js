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

// ADDITIONAL FIX: Enhanced setup filters to use proper field names
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

    console.log('Collecting filter values...');
    
    // Collect unique values from all line items (Personnel + NonPersonnel)
    const allLineItems = [...budgetData.personnel, ...budgetData.nonPersonnel];
    
    allLineItems.forEach((item, idx) => {
        if (idx < 5) console.log(`Line item ${idx}:`, item);
        
        // FIXED: Use explicit field names instead of fuzzy matching
        if (item.Fund) filters.fund.add(item.Fund);
        if (item.Department) filters.department.add(item.Department);
        if (item['Cost Center']) filters.department.add(item['Cost Center']); // Cost Center as department
        if (item.Division) filters.division.add(item.Division);
        if (item.Program) filters.program.add(item.Program);
        if (item.Status) filters.status.add(item.Status);
    });
    
    // From Request Summary (for request-level filters)
    budgetData.requestSummary.forEach(item => {
        if (item['Request Type']) filters.requestType.add(item['Request Type']);
        if (item.Status) filters.status.add(item.Status);
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
            <h1 style="color: #333; margin-bottom: 10px;">🎯 PBB Analysis & Recommendations</h1>
            <p style="color: #666; font-size: 1.1rem;">Detailed Scoring and Strategic Recommendations</p>
            <p style="color: #888;">Generated on ${reportDate}</p>
        </div>

        <div class="section-header">Analysis Overview</div>
        <p>This analytical report provides detailed Priority Based Budgeting (PBB) scoring and strategic recommendations for ${filteredData.length} budget requests totaling <strong class="amount">$${formatCurrency(totalAmount)}</strong>. Each request is evaluated across six criteria with actionable recommendations.</p>
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
            if (item.Quartile) return item.Quartile;
        } else if (fieldType === 'fund') {
            if (item.Fund) return item.Fund;
        } else if (fieldType === 'division') {
            if (item.Division) return item.Division;
        }
    }
    return null;
}

// ===== ENHANCED PBB SCORING ENGINE WITH EXPLICIT REASONING =====

function getQuartileScore(quartile) {
    if (!quartile) return { score: 0, reason: "No quartile alignment data found in line items" };
    if (quartile === 'Q1') return { score: 2, reason: `Program quartile is Q1 (Most Aligned) - highest priority alignment with city strategic goals and community priorities` };
    if (quartile === 'Q2') return { score: 2, reason: `Program quartile is Q2 (More Aligned) - strong alignment with city strategic goals and community priorities` };
    if (quartile === 'Q3') return { score: 1, reason: `Program quartile is Q3 (Less Aligned) - moderate alignment with city strategic goals` };
    return { score: 0, reason: `Program quartile is Q4 (Least Aligned) - lower priority alignment with current strategic goals` };
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

function getFundingScore(qa, qaText) {
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
    return { score: 0, reason: "No non-General Fund sources identified - request is 100% dependent on General Fund appropriation" };
}

function getMandateScore(qa, qaText) {
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

function scoreRequest(request) {
    const requestId = getRequestId(request);
    const lineItems = getLineItemsForRequest(requestId);
    const qa = getRequestQA(requestId);
    const amounts = getRequestAmount(request);
    
    const quartiles = lineItems.map(li => getPrimaryValue([li], 'quartile')).filter(q => q);
    const bestQuartile = getBestQuartile(quartiles);
    const qaText = qa.map(q => Object.values(q).join(' ')).join(' ').toLowerCase();
    
    // Score each criterion with explicit reasoning
    const quartileAnalysis = getQuartileScore(bestQuartile);
    const outcomeAnalysis = getOutcomeScore(qa, qaText);
    const fundingAnalysis = getFundingScore(qa, qaText);
    const mandateAnalysis = getMandateScore(qa, qaText);
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
        
        // Legacy fields for backwards compatibility
        bestQuartile: bestQuartile,
        hasOutsideFunding: /outside funding.*yes|grant|fee|partner|cost recovery/i.test(qaText),
        isMandated: /board motion|consent decree|doj|mandate|statute/i.test(qaText),
        isCompliance: /audit|liability|compliance|risk|safety/i.test(qaText)
    };
    
    // Calculate total score
    const totalScore = quartileAnalysis.score + outcomeAnalysis.score + fundingAnalysis.score + 
                      mandateAnalysis.score + efficiencyAnalysis.score + accessAnalysis.score;
    
    // Determine quartile band (High = Q1/Q2, Low = Q3/Q4)
    analysis.quartileBand = (bestQuartile === 'Q1' || bestQuartile === 'Q2') ? 'High' : 'Low';
    
    // Determine mandate level
    if (analysis.isMandated) {
        analysis.mandateLevel = 'Mandated';
    } else if (analysis.isCompliance) {
        analysis.mandateLevel = 'Compliance';
    } else {
        analysis.mandateLevel = 'None';
    }
    
    // Determine funding type
    analysis.fundingType = analysis.hasOutsideFunding ? 'NonGF' : 'GFonly';
    
    // Determine outcomes strength
    analysis.outcomesStrength = outcomeAnalysis.score >= 2 ? 'Strong' : 'Weak';
    
    analysis.totalScore = totalScore;
    
    // Apply the decision grid
    const gridDecision = applyDecisionGrid(analysis);
    
    analysis.disposition = gridDecision.disposition;
    analysis.dispositionColor = gridDecision.color;
    analysis.verifyNow = gridDecision.verifyNow;
    analysis.strengthenWith = gridDecision.strengthenWith;
    analysis.gridKey = gridDecision.gridKey;
    
    // Generate enhanced narrative
    analysis.narrative = generateEnhancedNarrative(request, lineItems, qa, analysis);
    
    return analysis;
}

// ===== DECISION GRID LOGIC =====
function applyDecisionGrid(analysis) {
    const { quartileBand, mandateLevel, fundingType, outcomesStrength } = analysis;
    
    // Create lookup key
    const gridKey = `${quartileBand}-${mandateLevel}-${fundingType}-${outcomesStrength}`;
    
    // Decision grid mapping
    const grid = {
        // HIGH RELEVANCE (Q1-Q2)
        'High-Mandated-NonGF-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['Statute/board reference', 'Allowability of non-GF sources'],
            strengthenWith: ['Final KPI list', 'Compliance milestones', 'Data source & cadence']
        },
        'High-Mandated-GFonly-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['Confirm mandate scope & minimums'],
            strengthenWith: ['Cost offsets (phase-down plan, reallocation)', 'Sunset/true-up triggers']
        },
        'High-Mandated-NonGF-Weak': {
            disposition: 'APPROVE',
            color: '#ffc107',
            verifyNow: ['That mandate truly requires this spend'],
            strengthenWith: ['Baseline→target KPIs', '90-day evaluation plan', 'Interim check-in']
        },
        'High-Mandated-GFonly-Weak': {
            disposition: 'APPROVE',
            color: '#ffc107',
            verifyNow: ['Minimum-viable compliance level'],
            strengthenWith: ['Add fee/grant search', 'Partner MOUs', 'Phased start', 'Sunset clause']
        },
        'High-Compliance-NonGF-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['Risk register link', 'Risk reduction metric'],
            strengthenWith: ['Cost avoidance calc', 'SLA updates', 'Internal control changes']
        },
        'High-Compliance-GFonly-Strong': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Materiality of risk', 'Alternatives'],
            strengthenWith: ['Add partial cost recovery', 'Internal reallocation', 'Pilot scope']
        },
        'High-Compliance-NonGF-Weak': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['That non-GF is real & timely'],
            strengthenWith: ['KPIs', '6-mo pilot with go/no-go', 'Light-weight evaluation plan']
        },
        'High-Compliance-GFonly-Weak': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Criticality (safety/liability)?'],
            strengthenWith: ['Narrow scope', 'Stage gates', 'Non-GF plan within 60–90 days']
        },
        'High-None-NonGF-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['No hidden GF backfill'],
            strengthenWith: ['Pay-for-itself math', 'Fee elasticity/grant terms', 'Partner commitments']
        },
        'High-None-GFonly-Strong': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Alternatives considered'],
            strengthenWith: ['Add cost recovery/partners', 'Unit-cost reduction', 'Partial reallocation']
        },
        'High-None-NonGF-Weak': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Outcome plausibility'],
            strengthenWith: ['KPIs & evaluation', 'Start as pilot', 'Tighten deliverables']
        },
        'High-None-GFonly-Weak': {
            disposition: 'DEFER',
            color: '#dc3545',
            verifyNow: ['N/A'],
            strengthenWith: ['Tie to priority KPIs', 'Find non-GF', 'Reduce scope or integrate with Q1/Q2 work']
        },
        
        // LOW RELEVANCE (Q3-Q4)
        'Low-Mandated-NonGF-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['Minimum compliance scope'],
            strengthenWith: ['Keep GF minimal', 'Escrow/offsets', 'Time-bound sunset']
        },
        'Low-Mandated-GFonly-Strong': {
            disposition: 'APPROVE',
            color: '#ffc107',
            verifyNow: ['Is Q3/Q4 mapping correct?'],
            strengthenWith: ['Identify fees/grants', 'Swap lower-impact spend', 'Phase', 'Sunset']
        },
        'Low-Mandated-NonGF-Weak': {
            disposition: 'APPROVE',
            color: '#ffc107',
            verifyNow: ['That mandate truly applies to this program'],
            strengthenWith: ['KPI baseline→target', '90-day review', 'Non-GF documentation']
        },
        'Low-Mandated-GFonly-Weak': {
            disposition: 'APPROVE',
            color: '#ffc107',
            verifyNow: ['Cheapest compliance path'],
            strengthenWith: ['Tight scope', 'Offsets', 'Timeline to add non-GF', 'Exit criteria']
        },
        'Low-Compliance-NonGF-Strong': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Non-GF terms & durability'],
            strengthenWith: ['No-GF pledge', 'Measurable risk reduction', 'Pilot + review']
        },
        'Low-Compliance-GFonly-Strong': {
            disposition: 'MODIFY',
            color: '#ffc107',
            verifyNow: ['Impact scale vs. alternatives'],
            strengthenWith: ['Require cost recovery', 'Internal reallocation', 'Narrower scope']
        },
        'Low-Compliance-NonGF-Weak': {
            disposition: 'DEFER',
            color: '#dc3545',
            verifyNow: ['Realism of benefits'],
            strengthenWith: ['Basic KPI set', 'Partner LOIs', 'Phase to prove value']
        },
        'Low-Compliance-GFonly-Weak': {
            disposition: 'DEFER',
            color: '#dc3545',
            verifyNow: ['If imminent, treat as mandate'],
            strengthenWith: ['Pilot w/ non-GF', 'Quantify liability avoided', 'Combine with Q1/Q2']
        },
        'Low-None-NonGF-Strong': {
            disposition: 'APPROVE',
            color: '#28a745',
            verifyNow: ['No GF drift'],
            strengthenWith: ['Full cost recovery', 'Service redesign', 'Contribution margin']
        },
        'Low-None-GFonly-Strong': {
            disposition: 'DEFER',
            color: '#dc3545',
            verifyNow: ['Competes with higher-Q needs'],
            strengthenWith: ['Add fee/grant/partner', 'ROI calc', 'Phase behind Q1/Q2']
        },
        'Low-None-NonGF-Weak': {
            disposition: 'DEFER',
            color: '#dc3545',
            verifyNow: ['N/A'],
            strengthenWith: ['KPIs', 'Tighten scope', 'Prove demand/willingness-to-pay']
        },
        'Low-None-GFonly-Weak': {
            disposition: 'REJECT',
            color: '#dc3545',
            verifyNow: ['N/A'],
            strengthenWith: ['Reframe to higher-Q outcome', 'Non-GF plan', 'Consolidate/streamline']
        }
    };
    
    const decision = grid[gridKey] || {
        disposition: 'MODIFY',
        color: '#ffc107',
        verifyNow: ['Unable to categorize - manual review needed'],
        strengthenWith: ['Provide complete information on mandate, funding, and outcomes']
    };
    
    decision.gridKey = gridKey;
    return decision;
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
        narrative += `⚠️ **COMPLIANCE/RISK**: This request addresses compliance obligations or risk mitigation.\n\n`;
    }
    
    if (analysis.hasOutsideFunding) {
        narrative += `✅ **NON-GF FUNDING**: Includes non-General Fund sources (grants, fees, or partnerships).\n\n`;
    } else if (analysis.quartileBand === 'Low') {
        narrative += `🚨 **FUNDING CONCERN**: 100% General Fund requested for a lower-relevance (Q3/Q4) program.\n\n`;
    }
    
    if (analysis.outcomesStrength === 'Strong') {
        narrative += `📊 **STRONG EVIDENCE**: Clear performance metrics and outcome targets provided.\n\n`;
    } else {
        narrative += `📋 **WEAK EVIDENCE**: Insufficient outcome data, KPIs, or evaluation plan.\n\n`;
    }
    
    narrative += `---\n\n`;
    
    // Disposition and recommendation with PBB suggests language
    narrative += `## 🎯 PBB SUGGESTS: **${analysis.disposition}** (Score: ${analysis.totalScore}/12)\n\n`;
    
    // Main recommendation based on disposition
    if (analysis.disposition === 'APPROVE') {
        if (analysis.mandateLevel === 'Mandated') {
            narrative += `**PBB Recommendation:** PBB suggests APPROVE. This is a mandated program with ${analysis.outcomesStrength.toLowerCase()} outcomes evidence. `;
            if (analysis.fundingType === 'GFonly' && analysis.quartileBand === 'Low') {
                narrative += `Given the lower quartile, PBB suggests requiring offsetting reductions or pursuing non-GF sources. `;
            }
            if (analysis.outcomesStrength === 'Weak') {
                narrative += `PBB suggests requiring metrics and evaluation plan as condition of approval.\n\n`;
            } else {
                narrative += `General Fund support appears justified based on mandate requirements.\n\n`;
            }
        } else if (analysis.fundingType === 'NonGF') {
            narrative += `**PBB Recommendation:** PBB suggests APPROVE with non-GF priority. Strong proposal with external funding sources. `;
            if (analysis.quartileBand === 'Low') {
                narrative += `For Q3/Q4 programs, PBB suggests ensuring minimal or no GF backfill. `;
            }
            narrative += `PBB recommends proceeding with clear cost recovery and sustainability plan.\n\n`;
        } else {
            narrative += `**PBB Recommendation:** PBB suggests APPROVE but strengthen funding strategy. While outcomes are strong, PBB recommends adding cost recovery or partnership elements to reduce General Fund reliance.\n\n`;
        }
    } else if (analysis.disposition === 'MODIFY') {
        narrative += `**PBB Recommendation:** PBB suggests MODIFY before approval. This request shows merit but PBB recommends adjustments before proceeding:\n\n`;
    } else if (analysis.disposition === 'DEFER') {
        narrative += `**PBB Recommendation:** PBB suggests DEFER. Insufficient business case for current approval based on PBB criteria. `;
        if (analysis.mandateLevel === 'Mandated') {
            narrative += `PBB recommends monitoring mandate requirements. `;
        }
        narrative += `See PBB-recommended strengthening actions below.\n\n`;
    } else if (analysis.disposition === 'REJECT') {
        narrative += `**PBB Recommendation:** PBB suggests REJECT OR SIGNIFICANT REDESIGN. `;
        narrative += `This low-relevance, GF-only request with weak outcomes does not meet PBB funding criteria. PBB recommends fundamental changes before reconsideration.\n\n`;
    }
    
    // Verification requirements
    if (analysis.verifyNow && analysis.verifyNow.length > 0 && analysis.verifyNow[0] !== 'N/A') {
        narrative += `### ✅ VERIFY NOW:\n\n`;
        analysis.verifyNow.forEach(item => {
            narrative += `- ${item}\n`;
        });
        narrative += `\n`;
    }
    
    // Strengthening actions
    if (analysis.strengthenWith && analysis.strengthenWith.length > 0) {
        narrative += `### 💪 TO STRENGTHEN THIS REQUEST:\n\n`;
        analysis.strengthenWith.forEach(item => {
            narrative += `- ${item}\n`;
        });
        narrative += `\n`;
    }
    
    // Specific follow-up prompts based on weaknesses
    narrative += `### 📝 SPECIFIC FOLLOW-UP ACTIONS:\n\n`;
    
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
    
    if (analysis.equityScore < 2 && analysis.quartileBand === 'High') {
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
    
    // Aggregate data by program within each department
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
                    totalCost: 0, // This would come from existing budget data
                    requestedAmount: 0,
                    proposedTotalCost: 0,
                    requestCount: 0
                };
            }
            
            // Add to requested amount (distribute across line items for this request)
            programData[dept][program].requestedAmount += amounts.total / lineItems.length;
            programData[dept][program].requestCount++;
            
            // For demo purposes, we'll estimate total cost as 8x the requested amount
            // In a real implementation, this would come from existing budget data
            if (programData[dept][program].totalCost === 0) {
                programData[dept][program].totalCost = amounts.total * 8; // Rough estimate
            }
            
            // Calculate proposed total
            programData[dept][program].proposedTotalCost = 
                programData[dept][program].totalCost + programData[dept][program].requestedAmount;
        });
    });

    let html = `<div class="section-header" id="program-summary">Program Summary</div>
                <p>Below is a summary of programs and their total requested amount and potential new total cost, organized by department and quartile alignment.</p>`;
    
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
                                <th style="padding: 12px 8px; text-align: right; width: 120px;">Total Cost</th>
                                <th style="padding: 12px 8px; text-align: right; width: 120px;">Requested Amount</th>
                                <th style="padding: 12px 8px; text-align: right; width: 140px;">Proposed Total Cost</th>
                            </tr>
                        </thead>
                        <tbody>
        `;
        
        // Sort programs by quartile (1=Most Aligned first)
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
                `<span class="quartile-badge quartile-${data.quartile.toLowerCase().replace(' ', '-')}" style="font-size: 0.8rem; padding: 4px 8px;">${data.quartile.replace(' Aligned', '')}</span>` : 
                '<span style="color: #666;">N/A</span>';
            
            html += `
                <tr style="border-bottom: 1px solid #e0e0e0;">
                    <td style="padding: 10px 8px; text-align: center;">${quartileBadge}</td>
                    <td style="padding: 10px 8px;">${program}</td>
                    <td style="padding: 10px 8px; text-align: right; color: #333;">$${formatCurrency(Math.round(data.totalCost))}</td>
                    <td style="padding: 10px 8px; text-align: right; color: #ffc107; font-weight: 600;">$${formatCurrency(Math.round(data.requestedAmount))}</td>
                    <td style="padding: 10px 8px; text-align: right; color: #28a745; font-weight: 600;">$${formatCurrency(Math.round(data.proposedTotalCost))}</td>
                </tr>
            `;
        });
        
        // Add department total row
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
            <strong>Department Impact Summary:</strong> ${dept} has ${Object.keys(programs).length} programs requesting 
            <span style="color: #ffc107; font-weight: 600;">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</span> 
            in additional funding, which would increase the department's total budget from 
            <span style="color: #333;">$${formatCurrency(Math.round(departmentTotal.totalCost))}</span> to 
            <span style="color: #28a745; font-weight: 600;">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</span> 
            (${((departmentTotal.requestedAmount / departmentTotal.totalCost) * 100).toFixed(1)}% increase).
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

// ===== STANDARD REPORT (NO ANALYSIS) =====
function generateDetailedRequestReportStandard() {
    let html = `<div class="section-header" id="individual-requests">Individual Budget Requests</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
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

        if (qa.length > 0) {
            html += generateRequestQASection(qa);
        }

        if (lineItems.length > 0) {
            html += generateLineItemSection(lineItems);
        }

        html += `</div></div>`;
    });

    return html;
}

// ===== ANALYTICAL REPORT (WITH SCORING) =====
function generateDetailedRequestReportAnalytical() {
    let html = `<div class="section-header" id="analytical-requests">Detailed Request Analysis</div>`;
    
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);
        
        // SCORE THE REQUEST
        const analysis = scoreRequest(request);
        
        const pageBreakStyle = index > 0 ? 'page-break-before: always;' : '';

        html += `
            <div class="request-card" id="analytical-request-${requestId}" style="${pageBreakStyle} margin-top: 40px; border-left: 5px solid ${analysis.dispositionColor};">
                <div class="request-header" style="background: ${analysis.dispositionColor}; color: white;">
                    <div class="request-title" style="color: white; font-size: 1.4rem;">
                        Request ${requestId} - ${description}
                        <span class="analysis-badge badge-${analysis.disposition.toLowerCase()}" style="float: right; margin-left: 15px;">
                            ${analysis.disposition}
                        </span>
                    </div>
                </div>
                <div class="request-details">
                    
                    <!-- PBB SCORING SECTION -->
                    <div style="background: linear-gradient(135deg, #f8f9ff, #ffffff); padding: 25px; margin-bottom: 25px; border-radius: 8px; border: 2px solid ${analysis.dispositionColor};">
                        <h3 style="color: ${analysis.dispositionColor}; margin-bottom: 20px; font-size: 1.5rem;">
                            📊 PBB Analysis Score: ${analysis.totalScore}/12
                        </h3>
                        
                        <!-- Score Breakdown with Explicit Reasons -->
                        <div style="margin: 20px 0;">
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #667eea;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #667eea; font-size: 1.1rem;">1. Program Alignment (Quartile)</strong>
                                    <span style="background: #667eea; color: white; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.quartileScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.quartileReason}</em></p>
                            </div>
                            
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #17a2b8;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #17a2b8; font-size: 1.1rem;">2. Outcome Evidence</strong>
                                    <span style="background: #17a2b8; color: white; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.outcomeScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.outcomeReason}</em></p>
                            </div>
                            
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #28a745;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #28a745; font-size: 1.1rem;">3. Funding Strategy</strong>
                                    <span style="background: #28a745; color: white; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.fundingScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.fundingReason}</em></p>
                            </div>
                            
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #ffc107;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #856404; font-size: 1.1rem;">4. Mandate/Risk</strong>
                                    <span style="background: #ffc107; color: #333; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.mandateScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.mandateReason}</em></p>
                            </div>
                            
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #6f42c1;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #6f42c1; font-size: 1.1rem;">5. Efficiency/ROI</strong>
                                    <span style="background: #6f42c1; color: white; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.efficiencyScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.efficiencyReason}</em></p>
                            </div>
                            
                            <div style="margin: 15px 0; padding: 15px; background: white; border-radius: 8px; border-left: 4px solid #e83e8c;">
                                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                    <strong style="color: #e83e8c; font-size: 1.1rem;">6. Access</strong>
                                    <span style="background: #e83e8c; color: white; padding: 6px 16px; border-radius: 20px; font-weight: bold; font-size: 1.1rem;">${analysis.accessScore}/2</span>
                                </div>
                                <p style="margin: 0; color: #555; font-size: 1rem; line-height: 1.6;"><em>${analysis.accessReason}</em></p>
                            </div>
                        </div>
                        
                        
                        
                        <!-- Strategic Recommendation -->
                        <div class="narrative-box">
                            <h4 style="color: #667eea; margin-bottom: 15px; font-size: 1.2rem;">📝 Strategic Recommendation</h4>
                            <div style="white-space: pre-wrap; font-size: 1.05rem; line-height: 1.8;">${analysis.narrative}</div>
                        </div>
                    </div>
                    
                    <!-- Request Details -->
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
                            <div class="detail-item">
                                <div class="detail-label">Department</div>
                                <div class="detail-value">${getPrimaryValue(lineItems, 'department') || 'N/A'}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Program</div>
                                <div class="detail-value">${getPrimaryValue(lineItems, 'program') || 'N/A'}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Quartile</div>
                                <div class="detail-value">
                                    <span class="quartile-badge quartile-${analysis.bestQuartile.toLowerCase()}">${analysis.bestQuartile}</span>
                                </div>
                            </div>
                        </div>
                    </div>
        `;

        if (qa.length > 0) {
            html += generateRequestQASection(qa);
        }

        if (lineItems.length > 0) {
            html += generateLineItemSection(lineItems);
        }

        html += `</div></div>`;
    });

    return html;
}

// Summary for Analytical Report
function generateAnalyticalSummary() {
    const scores = { approve: 0, modify: 0, defer: 0 };
    const amounts = { approve: 0, modify: 0, defer: 0 };
    
    filteredData.forEach(request => {
        const analysis = scoreRequest(request);
        const requestAmounts = getRequestAmount(request);
        
        if (analysis.disposition === 'APPROVE') {
            scores.approve++;
            amounts.approve += requestAmounts.total;
        } else if (analysis.disposition === 'MODIFY') {
            scores.modify++;
            amounts.modify += requestAmounts.total;
        } else {
            scores.defer++;
            amounts.defer += requestAmounts.total;
        }
    });

    return `
        <div class="section-header">Recommendation Summary</div>
        <div class="request-card">
            <div class="request-details">
                <div class="detail-grid">
                    <div class="detail-item" style="background: linear-gradient(135deg, #d4edda, #c3e6cb); border: 2px solid #28a745;">
                        <div class="detail-label">✅ Approve</div>
                        <div class="detail-value" style="font-size: 1.5rem; color: #28a745;">${scores.approve} Requests</div>
                        <div class="amount" style="font-size: 1.2rem;">$${formatCurrency(amounts.approve)}</div>
                    </div>
                    <div class="detail-item" style="background: linear-gradient(135deg, #fff3cd, #ffeeba); border: 2px solid #ffc107;">
                        <div class="detail-label">⚠️ Modify</div>
                        <div class="detail-value" style="font-size: 1.5rem; color: #856404;">${scores.modify} Requests</div>
                        <div class="amount" style="font-size: 1.2rem; color: #856404;">$${formatCurrency(amounts.modify)}</div>
                    </div>
                    <div class="detail-item" style="background: linear-gradient(135deg, #f8d7da, #f5c6cb); border: 2px solid #dc3545;">
                        <div class="detail-label">❌ Defer</div>
                        <div class="detail-value" style="font-size: 1.5rem; color: #dc3545;">${scores.defer} Requests</div>
                        <div class="amount" style="font-size: 1.2rem; color: #dc3545;">$${formatCurrency(amounts.defer)}</div>
                    </div>
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
                          analysis.disposition === 'MODIFY' ? '#ffc107' : '#dc3545';
        
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
    alert('Analytical Word Report download coming soon! Use the PDF option for now.');
    // You can implement this similar to downloadWordReport but using the analytical content
}

function downloadAnalyticalPdfReport() {
    alert('Analytical PDF Report download coming soon! You can print the analytical report using your browser\'s print function (Ctrl+P).');
    // You can implement this similar to downloadPdfReport but using the analytical content
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
            
            programData[dept][program].requestedAmount += amounts.total / lineItems.length;
            programData[dept][program].requestCount++;
            
            if (programData[dept][program].totalCost === 0) {
                programData[dept][program].totalCost = amounts.total * 8;
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
    const reportDate = new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });

    const totalAmount = filteredData.reduce((sum, request) => {
        const amounts = getRequestAmount(request);
        return sum + amounts.total;
    }, 0);

    // Calculate summary stats for the report
    let totalOngoing = 0;
    let totalOnetime = 0;
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
        
        const requestId = getRequestId(request);
        const lineItems = getLineItemsForRequest(requestId);
        
        lineItems.forEach(item => {
            const quartile = getPrimaryValue([item], 'quartile');
            if (quartile && quartileStats.hasOwnProperty(quartile)) {
                quartileStats[quartile] += amounts.total / lineItems.length;
            }
        });
    });

    // Create a print-optimized HTML document
    let printHtml = `
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Priority Based Budgeting Report</title>
            <style>
                @page { 
                    size: A4; 
                    margin: 0.75in; 
                    @top-center {
                        content: "Priority Based Budgeting Report - Page " counter(page);
                    }
                }
                
                body { 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    line-height: 1.5; 
                    color: #333; 
                    font-size: 12px;
                    margin: 0;
                    padding: 0;
                }
                
                .print-instruction {
                    background: #fff3cd;
                    border: 2px solid #ffc107;
                    padding: 15px;
                    margin: 20px 0;
                    border-radius: 8px;
                    text-align: center;
                    font-weight: bold;
                    color: #856404;
                }
                
                .header { 
                    text-align: center; 
                    margin-bottom: 40px; 
                    padding-bottom: 20px;
                    border-bottom: 3px solid #667eea;
                }
                
                .header h1 { 
                    color: #667eea; 
                    font-size: 28px; 
                    margin-bottom: 10px; 
                }
                
                .header p { 
                    color: #666; 
                    font-size: 14px; 
                    margin: 5px 0;
                }
                
                .section-header { 
                    color: #667eea; 
                    font-size: 18px; 
                    font-weight: 600; 
                    margin: 30px 0 20px 0; 
                    border-bottom: 2px solid #e0e0e0; 
                    padding-bottom: 10px; 
                    page-break-after: avoid;
                }
                
                .stats-container {
                    display: grid;
                    grid-template-columns: repeat(4, 1fr);
                    gap: 15px;
                    margin: 20px 0;
                    page-break-inside: avoid;
                }
                
                .stat-card {
                    background: linear-gradient(135deg, #667eea, #764ba2);
                    color: white;
                    padding: 20px;
                    border-radius: 10px;
                    text-align: center;
                    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                }
                
                .stat-value {
                    font-size: 24px;
                    font-weight: bold;
                    display: block;
                    margin-bottom: 8px;
                }
                
                .stat-label {
                    font-size: 11px;
                    opacity: 0.9;
                }
                
                .card { 
                    border: 2px solid #e0e0e0; 
                    margin: 20px 0; 
                    border-radius: 8px; 
                    page-break-inside: avoid;
                    background: white;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }
                
                .card-header { 
                    background: linear-gradient(135deg, #667eea, #764ba2); 
                    color: white; 
                    padding: 15px 20px; 
                    font-size: 16px; 
                    font-weight: 600; 
                    border-radius: 6px 6px 0 0;
                }
                
                .card-body { 
                    padding: 20px; 
                }
                
                table { 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 15px 0; 
                    font-size: 11px;
                    page-break-inside: auto;
                }
                
                th { 
                    background: #667eea; 
                    color: white; 
                    padding: 12px 8px; 
                    text-align: left; 
                    font-weight: 600; 
                    font-size: 12px;
                }
                
                td { 
                    padding: 10px 8px; 
                    border-bottom: 1px solid #ddd; 
                    vertical-align: top;
                }
                
                tr:nth-child(even) { 
                    background: #f8f9ff; 
                }
                
                .quartile-badge { 
                    display: inline-block; 
                    padding: 4px 12px; 
                    border-radius: 15px; 
                    font-size: 10px; 
                    font-weight: 600; 
                    color: white;
                }
                
                .quartile-most-aligned { background: #28a745; }
                .quartile-more-aligned { background: #17a2b8; }
                .quartile-less-aligned { background: #ffc107; color: black; }
                .quartile-least-aligned { background: #dc3545; }
                
                .amount { 
                    font-weight: 600; 
                    color: #28a745; 
                    font-size: 13px;
                }
                
                .page-break { 
                    page-break-before: always; 
                }
                
                .detail-section {
                    margin: 20px 0;
                    padding: 15px;
                    background: #f8f9ff;
                    border-radius: 8px;
                    border-left: 4px solid #667eea;
                }
                
                .detail-grid {
                    display: grid;
                    grid-template-columns: repeat(2, 1fr);
                    gap: 10px;
                    margin: 10px 0;
                }
                
                .detail-item {
                    padding: 8px;
                    background: white;
                    border-radius: 5px;
                    border: 1px solid #e0e0e0;
                }
                
                .detail-label {
                    font-size: 10px;
                    color: #666;
                    font-weight: 600;
                    margin-bottom: 3px;
                }
                
                .detail-value {
                    font-size: 12px;
                    color: #333;
                    font-weight: 500;
                }
                
                .qa-section {
                    background: #fff8f0;
                    border-left: 4px solid #ffc107;
                    padding: 15px;
                    margin: 15px 0;
                    border-radius: 0 8px 8px 0;
                    page-break-inside: avoid;
                }
                
                .qa-question {
                    font-weight: 600;
                    color: #667eea;
                    font-size: 13px;
                    margin-bottom: 8px;
                }
                
                .qa-answer {
                    line-height: 1.6;
                    color: #333;
                    font-size: 12px;
                }
                
                .chart-placeholder {
                    background: linear-gradient(45deg, #667eea, #764ba2);
                    color: white;
                    padding: 40px 20px;
                    text-align: center;
                    margin: 20px 0;
                    border-radius: 8px;
                    font-size: 16px;
                    font-weight: 600;
                }
                
                @media print {
                    .print-instruction { display: none; }
                    body { font-size: 11px; }
                    .page-break { page-break-before: always; }
                    .card { break-inside: avoid; }
                    .detail-section { break-inside: avoid; }
                    .qa-section { break-inside: avoid; }
                }
            </style>
        </head>
        <body>
            <div class="print-instruction">
                📄 To save as PDF: Press Ctrl+P (or Cmd+P on Mac), then select "Save as PDF" as your destination
            </div>
            
            <div class="header">
                <h1>Priority Based Budgeting Report</h1>
                <p>Budget Request Analysis and Recommendations</p>
                <p>Generated on ${reportDate}</p>
            </div>

            <div class="section-header">Executive Summary</div>
            <p>This comprehensive report analyzes <strong>${filteredData.length} budget requests</strong> totaling <strong class="amount">$${formatCurrency(totalAmount)}</strong> in requested funding. The requests span multiple departments and programs, with varying levels of alignment to organizational priorities.</p>
            
            <div class="stats-container">
                <div class="stat-card">
                    <span class="stat-value">${filteredData.length}</span>
                    <div class="stat-label">Total Requests</div>
                </div>
                <div class="stat-card">
                    <span class="stat-value">$${formatCurrency(totalOngoing)}</span>
                    <div class="stat-label">Ongoing Requests</div>
                </div>
                <div class="stat-card">
                    <span class="stat-value">$${formatCurrency(totalOnetime)}</span>
                    <div class="stat-label">One-time Requests</div>
                </div>
                <div class="stat-card">
                    <span class="stat-value">$${formatCurrency(totalAmount)}</span>
                    <div class="stat-label">Total Amount</div>
                </div>
            </div>

            <div class="section-header">Request Summary</div>
            <table>
                <thead>
                    <tr>
                        <th>Request ID</th>
                        <th>Description</th>
                        <th>Department</th>
                        <th>Quartile</th>
                        <th style="text-align: right;">Amount</th>
                    </tr>
                </thead>
                <tbody>
    `;

    // Add request summary table
    filteredData.forEach((request) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const primaryDept = getPrimaryValue(lineItems, 'department') || 'N/A';
        const primaryQuartile = getPrimaryValue(lineItems, 'quartile') || 'N/A';
        const amounts = getRequestAmount(request);

        const shortDesc = description && description.length > 40 ? 
            description.substring(0, 40) + '...' : (description || 'N/A');

        const quartileBadge = primaryQuartile !== 'N/A' ? 
            `<span class="quartile-badge quartile-${primaryQuartile.toLowerCase().replace(' ', '-')}">${primaryQuartile.replace(' Aligned', '')}</span>` : 'N/A';

        printHtml += `
            <tr>
                <td><strong>${requestId}</strong></td>
                <td>${shortDesc}</td>
                <td>${primaryDept}</td>
                <td>${quartileBadge}</td>
                <td style="text-align: right;" class="amount">$${formatCurrency(amounts.total)}</td>
            </tr>
        `;
    });

    printHtml += `
                </tbody>
            </table>
    `;

    // Add Program Summary to PDF
    printHtml += `
        <div class="section-header">Program Summary</div>
        <p>Below is a summary of programs and their total requested amount and potential new total cost, organized by department and quartile alignment.</p>
    `;

    // Generate program data for PDF
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
                    proposedTotalCost: 0
                };
            }

            programData[dept][program].requestedAmount += amounts.total / lineItems.length;

            if (programData[dept][program].totalCost === 0) {
                programData[dept][program].totalCost = amounts.total * 8;
            }

            programData[dept][program].proposedTotalCost =
                programData[dept][program].totalCost + programData[dept][program].requestedAmount;
        });
    });

    // Generate Program Summary tables for PDF
    Object.entries(programData).forEach(([dept, programs]) => {
        let departmentTotal = { totalCost: 0, requestedAmount: 0, proposedTotalCost: 0 };

        printHtml += `
            <div class="card">
                <div class="card-header">${dept}</div>
                <div class="card-body">
                    <table>
                        <thead>
                            <tr>
                                <th>Quartile</th>
                                <th>Program</th>
                                <th style="text-align: right;">Total Cost</th>
                                <th style="text-align: right;">Requested</th>
                                <th style="text-align: right;">Proposed Total</th>
                            </tr>
                        </thead>
                        <tbody>
        `;

        const sortedPrograms = Object.entries(programs).sort((a, b) => {
            const quartileOrder = { 'Most Aligned': 1, 'More Aligned': 2, 'Less Aligned': 3, 'Least Aligned': 4 };
            const aOrder = quartileOrder[a[1].quartile] || 5;
            const bOrder = quartileOrder[b[1].quartile] || 5;
            return aOrder - bOrder;
        });

        sortedPrograms.forEach(([program, data]) => {
            departmentTotal.totalCost += data.totalCost;
            departmentTotal.requestedAmount += data.requestedAmount;
            departmentTotal.proposedTotalCost += data.proposedTotalCost;

            const quartileBadge = data.quartile !== 'N/A' ?
                `<span class="quartile-badge quartile-${data.quartile.toLowerCase().replace(' ', '-')}">${data.quartile.replace(' Aligned', '')}</span>` : 'N/A';

            printHtml += `
                <tr>
                    <td>${quartileBadge}</td>
                    <td>${program}</td>
                    <td style="text-align: right;">$${formatCurrency(Math.round(data.totalCost))}</td>
                    <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(data.requestedAmount))}</td>
                    <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(data.proposedTotalCost))}</td>
                </tr>
            `;
        });

        printHtml += `
                <tr style="background: #f8f9ff; border-top: 2px solid #667eea; font-weight: 600;">
                    <td>TOTAL</td>
                    <td>${dept} Total</td>
                    <td style="text-align: right;">$${formatCurrency(Math.round(departmentTotal.totalCost))}</td>
                    <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</td>
                    <td style="text-align: right;" class="amount">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</td>
                </tr>
            </tbody>
        </table>
    
        <p style="margin-top: 15px; padding: 10px; background: #f0f8ff; border-radius: 5px; font-size: 12px;">
            <strong>Impact:</strong> ${Object.keys(programs).length} programs requesting 
            <span class="amount">$${formatCurrency(Math.round(departmentTotal.requestedAmount))}</span>, 
            increasing budget from $${formatCurrency(Math.round(departmentTotal.totalCost))} to 
            <span class="amount">$${formatCurrency(Math.round(departmentTotal.proposedTotalCost))}</span> 
            (${((departmentTotal.requestedAmount / departmentTotal.totalCost) * 100).toFixed(1)}% increase).
        </p>
    
        </div>
    </div>
        `;
    });

    // Add detailed request sections
    filteredData.forEach((request, index) => {
        const requestId = getRequestId(request);
        const description = getRequestDescription(request);
        const lineItems = getLineItemsForRequest(requestId);
        const qa = getRequestQA(requestId);
        const amounts = getRequestAmount(request);

        printHtml += `
            <div class="page-break">
                <div class="card">
                    <div class="card-header">Request ${requestId}: ${description}</div>
                    <div class="card-body">
                        <div class="detail-grid">
                            <div class="detail-item">
                                <div class="detail-label">Request ID</div>
                                <div class="detail-value">${requestId}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Total Amount</div>
                                <div class="detail-value amount">$${formatCurrency(amounts.total)}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Line Items</div>
                                <div class="detail-value">${lineItems.length}</div>
                            </div>
                            <div class="detail-item">
                                <div class="detail-label">Department</div>
                                <div class="detail-value">${getPrimaryValue(lineItems, 'department') || 'N/A'}</div>
                            </div>
                        </div>
        `;

        // Add Q&A sections with corrected question detection
        if (qa.length > 0) {
            qa.forEach(qItem => {
                let question = '';
                let answer = '';
                
                Object.keys(qItem).forEach(key => {
                    const lowerKey = key.toLowerCase();
                    // Look for Column C (Question) instead of Column F (Question Type) - SAME FIX AS UI
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
                    printHtml += `
                        <div class="qa-section">
                            <div class="qa-question">${question}</div>
                            <div class="qa-answer">${answer}</div>
                        </div>
                    `;
                }
            });
        }

        // Add detailed line items (replace the existing line items summary section)
        if (lineItems.length > 0) {
            printHtml += `
                <div class="detail-section" style="page-break-inside: avoid;">
                    <h4 style="margin-bottom: 15px; color: #667eea;">Line Item Details</h4>
            `;
            
            lineItems.forEach((item, idx) => {
                const quartile = getPrimaryValue([item], 'quartile');
                const quartileBadge = quartile ? 
                    `<span class="quartile-badge quartile-${quartile.toLowerCase().replace(' ', '-')}" style="margin-left: 10px;">${quartile}</span>` : '';

                printHtml += `
                    <div style="margin: 15px 0; padding: 15px; background: #f8f9ff; border-radius: 5px; border-left: 4px solid #667eea; page-break-inside: avoid;">
                        <div style="font-weight: 600; margin-bottom: 10px; font-size: 14px;">Line Item ${idx + 1} ${quartileBadge}</div>
                        
                        <!-- Comprehensive field display matching UI layout -->
                        <div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 8px; margin-bottom: 12px;">
                `;
                
                // First row - Basic Info
                const basicFields = ['REQUESTID', 'REQUEST DESCRIPTION', 'REQUEST TYPE', 'STATUS', 'ONGOING COST'];
                basicFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        const displayValue = formatFieldValue(field, value);
                        printHtml += `
                            <div style="background: white; padding: 8px; border-radius: 4px; border: 1px solid #e0e0e0; text-align: center;">
                                <div style="font-size: 9px; color: #666; font-weight: 600; margin-bottom: 3px;">${field}</div>
                                <div style="font-size: 11px; color: #333; font-weight: 500;">${displayValue}</div>
                            </div>
                        `;
                    }
                });

                printHtml += `</div><div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 8px; margin-bottom: 12px;">`;

                // Second row - Financial
                const financialFields = ['ONETIME COST', 'NUMBEROFITEMS', 'COST CENTER', 'ACCTTYPE', 'ACCTCODE'];
                financialFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        const displayValue = formatFieldValue(field, value);
                        printHtml += `
                            <div style="background: white; padding: 8px; border-radius: 4px; border: 1px solid #e0e0e0; text-align: center;">
                                <div style="font-size: 9px; color: #666; font-weight: 600; margin-bottom: 3px;">${field}</div>
                                <div style="font-size: 11px; color: #333; font-weight: 500;">${displayValue}</div>
                            </div>
                        `;
                    }
                });
                
                printHtml += `</div><div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 8px; margin-bottom: 12px;">`;
                
                // Third row - Organizational
                const orgFields = ['FUND', 'DEPARTMENT', 'ACCOUNT CATEGORY', 'PROGRAM', 'PROGRAMID'];
                orgFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        printHtml += `
                            <div style="background: white; padding: 8px; border-radius: 4px; border: 1px solid #e0e0e0; text-align: center;">
                                <div style="font-size: 9px; color: #666; font-weight: 600; margin-bottom: 3px;">${field}</div>
                                <div style="font-size: 11px; color: #333; font-weight: 500;">${value}</div>
                            </div>
                        `;
                    }
                });
                
                printHtml += `</div><div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 8px; margin-bottom: 8px;">`;
                
                // Fourth row - Scoring criteria
                const scoringFields = ['CHANGE IN DEMAND FOR THE PROGRAM', 'MANDATED TO PROVIDE PROGRAM', 'RELIANCE ON CITY TO PROVIDE PROGRAM', 'PORTION OF THE COMMUNITY SERVED'];
                scoringFields.forEach(field => {
                    const value = findFieldValue(item, field);
                    if (value !== null) {
                        printHtml += `
                            <div style="background: #fff8f0; padding: 8px; border-radius: 4px; border: 1px solid #ffc107; text-align: center;">
                                <div style="font-size: 8px; color: #666; font-weight: 600; margin-bottom: 3px;">${field}</div>
                                <div style="font-size: 10px; color: #333; font-weight: 500;">${value}</div>
                            </div>
                        `;
                    }
                });
                
                printHtml += `</div>`;
                
                // Fifth row - Additional fields
                const additionalFields = ['QUARTILE', 'COST RECOVERY OF PROGRAM'];
                const foundAdditional = additionalFields.filter(field => findFieldValue(item, field) !== null);
                
                if (foundAdditional.length > 0) {
                    printHtml += `<div style="display: grid; grid-template-columns: repeat(${foundAdditional.length}, 1fr); gap: 8px;">`;
                    foundAdditional.forEach(field => {
                        const value = findFieldValue(item, field);
                        printHtml += `
                            <div style="background: #f0f8f0; padding: 8px; border-radius: 4px; border: 1px solid #28a745; text-align: center;">
                                <div style="font-size: 9px; color: #666; font-weight: 600; margin-bottom: 3px;">${field}</div>
                                <div style="font-size: 11px; color: #333; font-weight: 500;">${value}</div>
                            </div>
                        `;
                    });
                    printHtml += `</div>`;
                }
                
                printHtml += `</div>`;
            });
            
            printHtml += `</div>`;
        }
        
        printHtml += `</div></div>`;
    });

    printHtml += `
        </body>
        </html>
    `;

    // Open in new window for printing
    const newWindow = window.open('', '_blank');
    newWindow.document.write(printHtml);
    newWindow.document.close();
    
    // Focus the new window
    newWindow.focus();
    
    // Show instruction alert
    setTimeout(() => {
        alert('Your report has opened in a new window. To save as PDF:\n\n1. Press Ctrl+P (or Cmd+P on Mac)\n2. Select "Save as PDF" as destination\n3. Click Save\n\nThe yellow instruction bar will not appear in the printed PDF.');
    }, 500);
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
