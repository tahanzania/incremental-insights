/**
 * Incremental Insights - Core Logic
 * Handles file parsing, data processing, and UI interactions locally.
 */

// --- State Management ---
const AppState = {
    rawData: [],
    processedData: [], // Data with normalized keys
    fieldMap: {}, // Maps internal keys to actual CSV headers
    filters: {
        partner: 'all',
        advertiser: 'all',
        campaign: 'all',
        kpiTypes: new Set(), // Changed to Set for multi-select
        beatingKpi: false
    },
    sortConfig: {
        key: 'calculatedOpportunity',
        direction: 'desc'
    },
    meta: {
        totalOpportunity: 0,
        qualifyingCount: 0
    }
};

// --- Configuration ---
// Possible variations of column headers to detect
const FIELD_MAPPING_CONFIG = {
    partner: ['Partner'],
    advertiser: ['Advertiser'],
    campaign: ['Campaign'],
    // User data shows "Market Type" holds "Decisioned"
    decisioned: ['Market Type', 'Decisioned', 'Decisioned/Non-Decisioned'],
    score: ['Avg. Campaign Decision Power Score', 'Decision Power Score', 'Score'],
    // User data specific header
    incrementalBudget: ['Avg. Campaign Daily Incremental Budget', 'Avg. Campaign Daily Incremental Budget Ideal', 'Campaign Incremental Budget', 'Daily Budget'],
    daysRemaining: ['Flight Days Remaining', 'Days Remaining', 'Remaining Days'],
    pacing: ['Pacing', 'Pacing Percentage', 'Campaign Pacing'],

    // New Fields
    beatingGoal: ['Beating KPI Goal'],
    kpiType: ['Goal Type', 'KPI Type'],
    goalValue: ['Goal Value'],
    avgKpiValue: ['Average KPI Value']
};

// --- DOM References ---
const UI = {
    dropArea: document.getElementById('drop-area'),
    fileInput: document.getElementById('file-input'),
    uploadSection: document.getElementById('upload-section'),
    dashboardSection: document.getElementById('dashboard-section'),

    // Filters
    filterPartner: document.getElementById('filter-partner'),
    filterAdvertiser: document.getElementById('filter-advertiser'),
    filterCampaign: document.getElementById('filter-campaign'),
    filterScore: document.getElementById('filter-score'),
    scoreVal: document.getElementById('score-val'),
    filterPacing: document.getElementById('filter-pacing'), // New
    pacingVal: document.getElementById('pacing-val'), // New
    filterKpiContainer: document.getElementById('filter-kpi-container'), // Checklist container
    filterBeatingKpi: document.getElementById('filter-beating-kpi'),

    // Help View
    btnHelp: document.getElementById('btn-help'),
    btnBackHome: document.getElementById('btn-back-home'),
    howItWorksSection: document.getElementById('how-it-works-section'),

    // Stats
    statTotal: document.getElementById('stat-total-campaigns'),
    statQualifying: document.getElementById('stat-qualifying'),
    statOpportunity: document.getElementById('stat-opportunity'),

    // Table
    tableTitle: document.getElementById('table-title'),
    recordCount: document.getElementById('record-count'),
    tableHead: document.querySelector('#data-table thead'),
    tableBody: document.querySelector('#data-table tbody'),

    // Views
    viewData: document.getElementById('view-data'),
    viewEmail: document.getElementById('view-email'),
    btnViewData: document.getElementById('btn-view-data'),
    btnViewEmail: document.getElementById('btn-email-view'),
    viewLevelSelect: document.getElementById('view-level-select'), // New aggregation switch

    // Actions
    btnCalculate: document.getElementById('btn-calculate'),
    btnReset: document.getElementById('btn-reset-file'),

    // Email
    emailTemplateType: document.getElementById('email-template-type'),
    emailStyle: document.getElementById('email-style'),
    emailTargetSelect: document.getElementById('email-target-select'),
    emailTargetLabel: document.getElementById('email-target-label'),
    btnGenerateEmail: document.getElementById('btn-generate-email'),
    emailOutput: document.getElementById('email-output'),
    btnCopyEmail: document.getElementById('btn-copy-email')
};

// --- Initialization ---
function init() {
    setupUploadListeners();
    setupFilterListeners();
    setupFilterListeners();
    setupNavigation();
    setupViewControls(); // New Listener
    setupEmailBuilder();
    setupHelpListeners(); // New listener logic
}

// --- File Upload & Parsing ---
function setupUploadListeners() {
    // Drag & Drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        UI.dropArea.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        UI.dropArea.addEventListener(eventName, () => UI.dropArea.classList.add('drag-over'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        UI.dropArea.addEventListener(eventName, () => UI.dropArea.classList.remove('drag-over'), false);
    });

    UI.dropArea.addEventListener('drop', handleDrop, false);
    UI.fileInput.addEventListener('change', handleFiles, false);

    // Click on drop area triggers input
    UI.dropArea.addEventListener('click', () => UI.fileInput.click());
    UI.fileInput.addEventListener('click', (e) => e.stopPropagation()); // Prevent bubbling
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles({ target: { files: files } });
}

function handleFiles(e) {
    const files = e.target.files;
    if (files.length === 0) return;

    const file = files[0];
    processFile(file);
}

function processFile(file) {
    const reader = new FileReader();

    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        try {
            const workbook = XLSX.read(data, { type: 'array' });

            // Assume first sheet
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (json.length === 0) {
                alert("File appears to be empty.");
                return;
            }

            normalizeData(json);
            transitionToDashboard();

        } catch (error) {
            console.error(error);
            alert("Error parsing file. Please ensure it is a valid Excel or CSV file.");
        }
    };

    reader.readAsArrayBuffer(file);
}

// --- Data Normalization ---
function normalizeData(json) {
    // 1. Identify Columns
    const headers = Object.keys(json[0]);
    AppState.fieldMap = {};

    // Map internal keys to found headers
    for (const [key, searchTerms] of Object.entries(FIELD_MAPPING_CONFIG)) {
        // Find a header that includes one of the search terms (case-insensitive)
        const match = headers.find(h =>
            searchTerms.some(term => h.trim().toLowerCase() === term.toLowerCase())
        );
        if (match) {
            AppState.fieldMap[key] = match;
        }
    }

    // 2. Process Rows
    AppState.rawData = json.map((row, index) => {
        const normalized = { _id: index, _row: row }; // Keep original

        // Extract known fields
        for (const [key, mapping] of Object.entries(AppState.fieldMap)) {
            let val = row[mapping];

            // Clean value (logic for comma removal etc)
            if (val !== undefined && val !== null) {
                // Clean value: remove commas and currency symbols ($)
                const cleanedStr = String(val).replace(/[$,]/g, '');

                if (key === 'pacing') {
                    if (String(val).includes('%')) {
                        // "98%" -> 98
                        val = parseFloat(cleanedStr.replace('%', ''));
                    } else {
                        val = parseFloat(cleanedStr);
                        // Heuristic: If value is small (<= 2.0 like 1.05), assume ratio (1.05 -> 105%). 
                        // If value is > 2.0 (like 98), assume percentage (98 -> 98%).
                        if (!isNaN(val) && val <= 2.0) {
                            val = val * 100;
                        }
                    }

                    // Cap at 100%
                    if (val > 100) val = 100;
                    // Ensure decimals are handled if needed, but 100 is max.
                } else if (['score', 'incrementalBudget', 'daysRemaining', 'avgKpiValue', 'goalValue'].includes(key)) {
                    val = parseFloat(cleanedStr) || 0;
                } else {
                    val = String(val).trim(); // Text fields
                }
            } else {
                val = ''; // Default empty
                if (['score', 'incrementalBudget', 'daysRemaining', 'pacing'].includes(key)) val = 0;
            }

            normalized[key] = val;
        }

        // Calculate KPI Performance Ratio
        // (Average KPI Value / Goal Value) as requested
        const avg = parseFloat(normalized.avgKpiValue) || 0;
        const goal = parseFloat(normalized.goalValue) || 0;

        if (goal !== 0) {
            normalized.kpiPerfRatio = avg / goal;
        } else {
            normalized.kpiPerfRatio = 0;
        }

        // Normalize beatingGoal to boolean
        // RECALCULATE based on user rule:
        // CPA/Cost: Beats if under goal.
        // Others: Beats if over goal.

        let calculatedBeating = false;
        const type = String(normalized.kpiType).toLowerCase();

        if (type.includes('cpa') || type.includes('cost')) {
            // Lower is better
            // Beating if Avg <= Goal (and Goal is not 0)
            if (goal > 0 && avg <= goal) calculatedBeating = true;
        } else {
            // Higher is better (CTR, VCR, ROAS)
            if (avg >= goal) calculatedBeating = true;
        }

        // Use calculated if goal exists, otherwise fallback to CSV
        if (goal > 0) {
            normalized.beatingGoalBool = calculatedBeating;
        } else {
            const bg = String(normalized.beatingGoal).toLowerCase();
            normalized.beatingGoalBool = bg === 'true' || bg === 'yes' || bg === '1';
        }

        // Default 'calculatedOpportunity'
        normalized.calculatedOpportunity = 0;

        return normalized;
    });

    AppState.processedData = [...AppState.rawData];

    // Initial UI Population
    populateFilters(AppState.processedData);

    // Auto-Run Calculation
    runCalculation();
}

// --- UI Logic ---
function transitionToDashboard() {
    UI.uploadSection.classList.add('hidden');
    UI.uploadSection.classList.remove('active-view');
    UI.dashboardSection.classList.remove('hidden');
}

UI.btnReset.addEventListener('click', () => {
    location.reload(); // Simple reset
});

// --- Filtering ---
function populateFilters(data) {
    // Get unique values
    const partners = [...new Set(data.map(d => d.partner).filter(Boolean))].sort();
    const advertisers = [...new Set(data.map(d => d.advertiser).filter(Boolean))].sort();
    const campaigns = [...new Set(data.map(d => d.campaign).filter(Boolean))].sort();
    const kpiTypes = [...new Set(data.map(d => d.kpiType).filter(Boolean))].sort();
    fillSelect(UI.filterPartner, partners);
    fillSelect(UI.filterAdvertiser, advertisers);
    fillSelect(UI.filterCampaign, campaigns);

    // KPI Types - Checkbox List
    UI.filterKpiContainer.innerHTML = '';

    if (kpiTypes.length === 0) {
        UI.filterKpiContainer.innerHTML = '<span style="color:var(--text-muted); font-size:0.85rem;">No KPI Types found</span>';
    } else {
        kpiTypes.forEach(type => {
            const label = document.createElement('label');
            label.style.display = 'flex';
            label.style.alignItems = 'center';
            label.style.gap = '0.5rem';
            label.style.fontSize = '0.9rem';
            label.style.color = 'var(--text-main)';
            label.style.cursor = 'pointer';

            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.value = type;
            checkbox.classList.add('kpi-checkbox');

            // Default: Select all EXCEPT "No Goal"
            const lower = type.toLowerCase();
            if (lower !== 'no goal' && lower !== 'none' && lower !== 'n/a') {
                checkbox.checked = true;
            }

            label.appendChild(checkbox);
            label.appendChild(document.createTextNode(type));

            UI.filterKpiContainer.appendChild(label);
        });
    }
}

function fillSelect(select, values) {
    const current = select.value;
    // Keep first option
    select.innerHTML = select.options[0].outerHTML;

    values.forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        select.appendChild(opt);
    });

    if (values.includes(current)) select.value = current;
}

function setupFilterListeners() {
    const inputs = [UI.filterPartner, UI.filterAdvertiser, UI.filterCampaign, UI.filterBeatingKpi];
    inputs.forEach(input => {
        input.addEventListener('change', applyFilters);
    });

    // Delegate changes from checklist
    if (UI.filterKpiContainer) {
        UI.filterKpiContainer.addEventListener('change', applyFilters);
    }

    // Slider listener
    if (UI.filterScore) {
        UI.filterScore.addEventListener('input', (e) => {
            UI.scoreVal.textContent = e.target.value;
            // Debounce or just run? It's client side, just run.
            runCalculation();
        });
    }

    if (UI.filterPacing) {
        UI.filterPacing.addEventListener('input', (e) => {
            UI.pacingVal.textContent = e.target.value + '%';
            runCalculation();
        });
    }

    UI.btnCalculate.addEventListener('click', () => {
        runCalculation();
        // Visual feedback
        const originalText = UI.btnCalculate.innerHTML;
        UI.btnCalculate.innerHTML = `<i data-lucide="check"></i> Calculated`;
        if (window.lucide) lucide.createIcons();
        setTimeout(() => UI.btnCalculate.innerHTML = originalText, 2000);
    });
}

function setupViewControls() {
    if (UI.viewLevelSelect) {
        UI.viewLevelSelect.addEventListener('change', () => {
            // Reset sort when changing view? Or keep?
            // Resetting might be safer as keys change.
            AppState.sortConfig.key = 'calculatedOpportunity';
            renderTable(AppState.processedData);
        });
    }
}

// Global function for pivot toggle
window.togglePivotRow = function (id) {
    const row = document.getElementById(id);
    if (!row) return;

    const isExpanded = row.getAttribute('data-expanded') === 'true';
    row.setAttribute('data-expanded', !isExpanded);

    // Toggle icon
    const icon = row.querySelector('.pivot-icon');
    if (icon) {
        icon.innerHTML = !isExpanded ?
            `<i data-lucide="chevron-down" width="16" height="16"></i>` :
            `<i data-lucide="chevron-right" width="16" height="16"></i>`;
        if (window.lucide) lucide.createIcons();
    }

    // Find children
    const level = parseInt(row.getAttribute('data-level'));
    let nextRow = row.nextElementSibling;

    while (nextRow) {
        const nextLevel = parseInt(nextRow.getAttribute('data-level'));
        if (nextLevel <= level) break; // End of this block

        // Logic:
        // If expanding parent: show immediate children (level + 1).
        // If collapsing parent: hide ALL descendants.

        if (!isExpanded) {
            // EXPANDING
            // Only show immediate children
            if (nextLevel === level + 1) {
                nextRow.classList.remove('hidden');
                // Ensure their state is reflected (e.g. if they were expanded, should their children show?
                // For simplicity, let's keep their children hidden unless they assume 'expanded' state logic which is complex.
                // Let's just show immediate children as collapsed unless we track state deeply.

                // Better approach: When expanding, we show direct children.
                // If a direct child was previously expanded, we might want to show its children too, but standard behavior is fine to keep them hidden or restore state.
                // Simple: Show direct children.
                nextRow.setAttribute('data-expanded', 'false'); // Reset child expansion? or keep?

                // Let's go with: Show direct children. For grandchildren, check if the direct child is expanded.
            }
        } else {
            // COLLAPSING
            // Hide everything deeper until sibling
            nextRow.classList.add('hidden');
            // We should also visually reset the expansion of children if we want a clean state next time?
            // Optional.
        }

        nextRow = nextRow.nextElementSibling;
    }

    // Re-run expansion check if we want to restore deep state?
    // Actually, "Toggle" logic usually hides all descendants.
    // "Expand" usually only shows direct children OR restores previous state.
    // Let's implement robust "Hide All Descendants" on collapse.
    // On Expand: Show direct children. If a direct child is marked expanded, show its children?
    // Let's do simple: Collapse hides all. Expand shows direct children.
};

function applyFilters() {
    const fPartner = UI.filterPartner.value;
    const fAdvertiser = UI.filterAdvertiser.value;
    const fCampaign = UI.filterCampaign.value;

    // Get all checked boxes
    const checkedBoxes = UI.filterKpiContainer.querySelectorAll('.kpi-checkbox:checked');
    const selectedKpiTypes = Array.from(checkedBoxes).map(cb => cb.value);

    const fBeatingKpi = UI.filterBeatingKpi.checked;

    AppState.processedData = AppState.rawData.filter(item => {
        if (fPartner !== 'all' && item.partner !== fPartner) return false;
        if (fAdvertiser !== 'all' && item.advertiser !== fAdvertiser) return false;
        if (fCampaign !== 'all' && item.campaign !== fCampaign) return false;

        // Multi-select KPI check
        if (!selectedKpiTypes.includes(item.kpiType)) return false;

        if (fBeatingKpi && !item.beatingGoalBool) return false;

        return true;
    });

    // Re-run calculation only, which updates stats and table
    runCalculation();
}

// --- Calculation Logic ---
function runCalculation() {
    let grandTotal = 0;
    let qualifyingCount = 0;

    // Get dynamic threshold
    const scoreThreshold = UI.filterScore ? parseInt(UI.filterScore.value) : 100;
    const pacingThreshold = UI.filterPacing ? parseInt(UI.filterPacing.value) : 99;

    AppState.processedData.forEach(item => {
        // Logic:
        // Pacing > Threshold
        // AND Power Score > Threshold

        let isPacingHigh = false;
        // Or "1" if formatted number.
        // We look for "At or near 100" or exceeding 100

        // Normalized to 0-100 scale.
        if (item.pacing >= pacingThreshold) isPacingHigh = true;

        const isScoreHigh = item.score > scoreThreshold;

        if (isPacingHigh && isScoreHigh) {
            // Compute: Incremental Budget * Days Remaining
            const opp = item.incrementalBudget * item.daysRemaining;
            item.calculatedOpportunity = opp;
            grandTotal += opp;
            qualifyingCount++;
        } else {
            item.calculatedOpportunity = 0;
        }
    });

    AppState.meta.totalOpportunity = grandTotal;
    AppState.meta.qualifyingCount = qualifyingCount;

    // Refresh table and stats
    renderTable(AppState.processedData);
    updateStats(AppState.processedData);
}


// --- Rendering ---
function updateStats(data) {
    UI.statTotal.textContent = data.length;
    UI.statQualifying.textContent = AppState.meta.qualifyingCount;
    UI.statOpportunity.textContent = formatCurrency(AppState.meta.totalOpportunity);
    UI.recordCount.textContent = `${data.length} records`;
}

function handleTableSort(key) {
    if (AppState.sortConfig.key === key) {
        // Toggle direction
        AppState.sortConfig.direction = AppState.sortConfig.direction === 'asc' ? 'desc' : 'asc';
    } else {
        // New key, default to desc for numbers? or asc for text?
        // Let's standard default to desc for easier view of 'top' items
        AppState.sortConfig.key = key;
        AppState.sortConfig.direction = 'desc';
    }
    // Re-render
    renderTable(AppState.processedData);
}

function renderTable(data) {
    const viewLevel = UI.viewLevelSelect ? UI.viewLevelSelect.value : 'campaign';

    if (viewLevel === 'pivot') {
        renderPivotView(data);
        return;
    }

    // --- AGGREGATION LOGIC ---
    let tableData = data;

    if (viewLevel !== 'campaign') {
        // Aggregation needed
        const groups = {};

        data.forEach(item => {
            const key = item[viewLevel] || 'Unknown';
            if (!groups[key]) {
                groups[key] = {
                    name: key,
                    partner: item.partner, // fallback for advertiser view
                    count: 0,
                    calculatedOpportunity: 0,
                    incrementalBudget: 0,
                    totalScore: 0,
                    // We can't easily sum avgKpi without weight, omit for now or show sample
                };
            }

            groups[key].count++;
            groups[key].calculatedOpportunity += (item.calculatedOpportunity || 0);
            groups[key].incrementalBudget += (item.incrementalBudget || 0);
            groups[key].totalScore += (item.score || 0);
        });

        // Convert back to array
        tableData = Object.values(groups).map(g => ({
            name: g.name,
            partner: g.partner,
            count: g.count,
            calculatedOpportunity: g.calculatedOpportunity,
            incrementalBudget: g.incrementalBudget,
            avgScore: Math.round(g.totalScore / g.count)
        }));
    }


    // --- SORTING LOGIC ---
    const { key: sortKey, direction } = AppState.sortConfig;

    // Safety check: if sort key doesn't exist in aggregated data (e.g. 'pacing'), fallback
    // But 'calculatedOpportunity' exists in both.

    tableData.sort((a, b) => {
        let valA = a[sortKey];
        let valB = b[sortKey];

        // Handle fallback logic for missing keys in aggregate view
        if (valA === undefined) valA = 0;
        if (valB === undefined) valB = 0;

        const isString = typeof valA === 'string' || typeof valB === 'string';
        if (isString) {
            valA = String(valA).toLowerCase();
            valB = String(valB).toLowerCase();
        }

        if (valA < valB) return direction === 'asc' ? -1 : 1;
        if (valA > valB) return direction === 'asc' ? 1 : -1;
        return 0;
    });


    // --- RENDERING ---
    let headers = [];
    let keys = [];

    if (viewLevel === 'campaign') {
        // Detailed View
        keys = ['partner', 'advertiser', 'campaign', 'kpiType', 'avgKpiValue', 'goalValue', 'kpiPerfRatio', 'score', 'daysRemaining', 'pacing', 'incrementalBudget', 'calculatedOpportunity'];
        headers = ['Partner', 'Advertiser', 'Campaign', 'Goal Type', 'Avg. KPI', 'Goal', 'KPI Perf.', 'Score', 'Days', 'Pacing', 'Inc. Budget', 'Total Inc. Opp.'];
    } else {
        // Summary Views
        if (viewLevel === 'advertiser') {
            keys = ['name', 'partner', 'count', 'avgScore', 'incrementalBudget', 'calculatedOpportunity'];
            headers = ['Advertiser', 'Partner', 'Campaigns', 'Avg. Score', 'Total Inc. Budget', 'Total Inc. Opp.'];
        } else {
            // Partner
            keys = ['name', 'count', 'avgScore', 'incrementalBudget', 'calculatedOpportunity'];
            headers = ['Partner', 'Campaigns', 'Avg. Score', 'Total Inc. Budget', 'Total Inc. Opp.'];
        }
    }

    // Header Generation
    let htmlHead = '<tr>';
    headers.forEach((h, index) => {
        const key = keys[index];
        const isSorted = sortKey === key;
        const arrow = isSorted ? (direction === 'asc' ? '▲' : '▼') : '';
        htmlHead += `<th style="cursor:pointer;" data-sort-key="${key}">${h} <span style="font-size:0.8rem; margin-left:4px;">${arrow}</span></th>`;
    });
    htmlHead += '</tr>';
    UI.tableHead.innerHTML = htmlHead;

    // Attach listeners
    const ths = UI.tableHead.querySelectorAll('th');
    ths.forEach(th => {
        th.addEventListener('click', () => {
            const key = th.dataset.sortKey;
            handleTableSort(key);
        });
    });

    // Body Render
    const subset = tableData.slice(0, 100);

    UI.tableBody.innerHTML = subset.map(item => {
        if (viewLevel === 'campaign') {
            // ... existing campaign row logic ...
            const perf = item.kpiPerfRatio ? formatRatio(item.kpiPerfRatio) : '-';
            const color = item.beatingGoalBool ? 'var(--success)' : '#ef4444';
            const formattedKpi = formatKpi(item.avgKpiValue, item.kpiType);
            const formattedGoal = formatKpi(item.goalValue, item.kpiType);

            return `<tr>
                <td>${item.partner || '-'}</td>
                <td>${item.advertiser || '-'}</td>
                <td><div style="max-width:200px; overflow:hidden; text-overflow:ellipsis;" title="${item.campaign}">${item.campaign || '-'}</div></td>
                <td>${item.kpiType || item.decisioned || '-'}</td>
                <td>${formattedKpi}</td>
                <td>${formattedGoal}</td>
                <td style="color:${color}; font-weight:500;">${perf}</td>
                <td>${item.score || 0}</td>
                <td>${item.daysRemaining || 0}</td>
                <td>${formatPercent(item.pacing)}</td>
                <td>${formatCurrency(item.incrementalBudget)}</td>
                <td style="color: var(--success); font-weight:600;">${item.calculatedOpportunity > 0 ? formatCurrency(item.calculatedOpportunity) : '-'}</td>
            </tr>`;
        } else {
            // Summary Row
            // Re-use logic for Advertiser/Partner
            let row = '<tr>';
            row += `<td><span style="font-weight:600;">${item.name || 'Unknown'}</span></td>`;
            if (viewLevel === 'advertiser') {
                row += `<td>${item.partner || '-'}</td>`;
            }
            row += `<td>${item.count}</td>`;
            row += `<td>${item.avgScore}</td>`;
            row += `<td>${formatCurrency(item.incrementalBudget)}</td>`;
            row += `<td style="color: var(--success); font-weight:600;">${formatCurrency(item.calculatedOpportunity)}</td>`;
            row += '</tr>';
            return row;
        }
    }).join('');

    if (tableData.length > 100) {
        UI.tableBody.innerHTML += `<tr><td colspan="${headers.length}" style="text-align:center; opacity:0.5;">...and ${tableData.length - 100} more rows</td></tr>`;
    }
}

function renderPivotView(data) {
    // 1. Build Hierarchy
    // Partner -> Advertiser -> Campaign
    const tree = {};

    data.forEach(item => {
        const pName = item.partner || 'Unknown Partner';
        const aName = item.advertiser || 'Unknown Advertiser';

        if (!tree[pName]) {
            tree[pName] = {
                id: `p-${sanitizeId(pName)}`,
                name: pName,
                metrics: createMetrics(),
                children: {}
            };
        }

        if (!tree[pName].children[aName]) {
            tree[pName].children[aName] = {
                id: `p-${sanitizeId(pName)}-a-${sanitizeId(aName)}`,
                name: aName,
                metrics: createMetrics(),
                children: [] // campaigns list
            };
        }

        // Add campaign
        tree[pName].children[aName].children.push(item);

        // Aggregate Metrics
        accumulateMetrics(tree[pName].metrics, item);
        accumulateMetrics(tree[pName].children[aName].metrics, item);
    });

    // 2. Sort Hierarchy
    const sortConfig = AppState.sortConfig; // e.g. calculatedOpportunity desc
    const sortFn = (a, b) => {
        let valA, valB;
        // Check if sorting by metric or name
        if (sortConfig.key === 'partner' || sortConfig.key === 'advertiser' || sortConfig.key === 'campaign') {
            valA = (a.name || a.campaign || '').toLowerCase();
            valB = (b.name || b.campaign || '').toLowerCase();
        } else {
            // Metrics (calculatedOpportunity, incrementalBudget, score(avg?))
            valA = a.metrics ? a.metrics[sortConfig.key] : a[sortConfig.key];
            valB = b.metrics ? b.metrics[sortConfig.key] : b[sortConfig.key];

            // Special case for score in pivot nodes (we store sum and count, need avg for sort?)
            if (sortConfig.key === 'score' && a.metrics) {
                valA = a.metrics.score / a.metrics.count;
                valB = b.metrics.score / b.metrics.count;
            }
        }

        valA = valA || 0;
        valB = valB || 0;

        if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
    };

    const partners = Object.values(tree).sort(sortFn);
    partners.forEach(p => {
        // Sort Advertisers
        const advertisers = Object.values(p.children).sort(sortFn);
        p.sortedChildren = advertisers;

        advertisers.forEach(a => {
            // Sort Campaigns
            a.children.sort(sortFn);
        });
    });

    // 3. Render
    // Headers
    const headers = [
        { key: 'partner', label: 'Hierarchy' },
        { key: 'score', label: 'Avg Score' },
        { key: 'incrementalBudget', label: 'Inc. Budget' },
        { key: 'calculatedOpportunity', label: 'Inc. Opportunity' }
    ];

    let htmlHead = '<tr>';
    headers.forEach(h => {
        const isSorted = sortConfig.key === h.key;
        const arrow = isSorted ? (sortConfig.direction === 'asc' ? '▲' : '▼') : '';
        htmlHead += `<th style="cursor:pointer;" data-sort-key="${h.key}">${h.label} <span style="font-size:0.8rem; margin-left:4px;">${arrow}</span></th>`;
    });
    htmlHead += '</tr>';
    UI.tableHead.innerHTML = htmlHead;

    // Attach sort listeners
    UI.tableHead.querySelectorAll('th').forEach(th => {
        th.addEventListener('click', () => {
            handleTableSort(th.dataset.sortKey);
        });
    });

    // Body logic
    let htmlBody = '';

    partners.forEach(p => {
        // Partner Row (Level 1)
        const pOpp = formatCurrency(p.metrics.calculatedOpportunity);
        const pBud = formatCurrency(p.metrics.incrementalBudget);
        const pScore = Math.round(p.metrics.score / p.metrics.count);

        htmlBody += `<tr id="${p.id}" class="pivot-row level-1" data-level="1" data-expanded="false" onclick="togglePivotRow('${p.id}')" style="cursor:pointer; background:rgba(0,0,0,0.02);">
            <td style="font-weight:700; color:var(--primary);">
                <div style="display:flex; align-items:center; gap:0.5rem;">
                    <span class="pivot-icon"><i data-lucide="chevron-right" width="16" height="16"></i></span>
                    ${p.name} <span class="badge" style="font-size:0.75rem;">${p.metrics.count}</span>
                </div>
            </td>
            <td>${pScore}</td>
            <td>${pBud}</td>
            <td style="font-weight:600; color:var(--success);">${pOpp}</td>
        </tr>`;

        p.sortedChildren.forEach(a => {
            // Advertiser Row (Level 2) - Hidden by default
            const aOpp = formatCurrency(a.metrics.calculatedOpportunity);
            const aBud = formatCurrency(a.metrics.incrementalBudget);
            const aScore = Math.round(a.metrics.score / a.metrics.count);

            htmlBody += `<tr id="${a.id}" class="pivot-row level-2 hidden" data-level="2" data-expanded="false" onclick="togglePivotRow('${a.id}')" style="cursor:pointer;">
                <td style="padding-left:2.5rem; font-weight:600;">
                    <div style="display:flex; align-items:center; gap:0.5rem;">
                        <span class="pivot-icon"><i data-lucide="chevron-right" width="16" height="16"></i></span>
                        ${a.name} <span class="badge" style="font-size:0.75rem; background:rgba(0,0,0,0.05);">${a.metrics.count}</span>
                    </div>
                </td>
                <td>${aScore}</td>
                <td>${aBud}</td>
                <td style="font-weight:600; color:var(--success);">${aOpp}</td>
            </tr>`;

            a.children.forEach(c => {
                // Campaign Row (Level 3) - Hidden by default
                const cOpp = formatCurrency(c.calculatedOpportunity);
                const cBud = formatCurrency(c.incrementalBudget);
                // Extra context for campaigns

                htmlBody += `<tr class="pivot-row level-3 hidden" data-level="3" style="background:rgba(255,255,255,0.5);">
                    <td style="padding-left:5rem; font-size:0.9rem;">
                        ${c.campaign}
                    </td>
                    <td>${c.score}</td>
                    <td>${cBud}</td>
                    <td style="color:var(--success);">${cOpp}</td>
                </tr>`;
            });
        });
    });

    UI.tableBody.innerHTML = htmlBody;
    if (window.lucide) lucide.createIcons();
}

// Helpers
function createMetrics() {
    return { count: 0, score: 0, incrementalBudget: 0, calculatedOpportunity: 0 };
}

function accumulateMetrics(target, item) {
    target.count++;
    target.score += (item.score || 0);
    target.incrementalBudget += (item.incrementalBudget || 0);
    target.calculatedOpportunity += (item.calculatedOpportunity || 0);
}

function sanitizeId(str) {
    return String(str).replace(/[^a-zA-Z0-9-_]/g, '-').toLowerCase();
}

function formatCurrency(val) {
    return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val);
}

function formatPercent(val) {
    if (val === undefined || val === null || isNaN(val)) return '0%';
    // Cap visual display at 100%
    let num = Math.round(val);
    if (num > 100) num = 100;
    return num + '%';
}


// --- Email Builder ---
function setupNavigation() {
    UI.btnViewData.addEventListener('click', () => switchView('data'));
    UI.btnViewEmail.addEventListener('click', () => switchView('email'));
    document.getElementById('btn-back-data')?.addEventListener('click', () => switchView('data'));
}

function switchView(view) {
    if (view === 'data') {
        UI.viewData.classList.remove('hidden');
        UI.viewData.classList.add('active');
        UI.viewEmail.classList.add('hidden');
        UI.viewEmail.classList.remove('active');
        UI.btnViewData.classList.add('active');
        UI.btnViewEmail.classList.remove('active');
    } else {
        UI.viewData.classList.add('hidden');
        UI.viewData.classList.remove('active');
        UI.viewEmail.classList.remove('hidden');
        UI.viewEmail.classList.add('active');
        UI.btnViewData.classList.remove('active');
        UI.btnViewEmail.classList.add('active');

        // Update targets on switch
        updateEmailTargets();
    }
}

function setupEmailBuilder() {
    UI.emailTemplateType.addEventListener('change', updateEmailTargets);
    UI.btnGenerateEmail.addEventListener('click', generateEmail);
    UI.btnCopyEmail.addEventListener('click', copyEmail);
}

function setupHelpListeners() {
    UI.btnHelp.addEventListener('click', () => toggleHelp(true));
    UI.btnBackHome.addEventListener('click', () => toggleHelp(false));
}

function toggleHelp(show) {
    if (show) {
        // Save current view state implicitly by verifying which one is not hidden later
        // Just hide both main sections and show help
        UI.uploadSection.classList.add('hidden');
        UI.dashboardSection.classList.add('hidden');
        UI.howItWorksSection.classList.remove('hidden');
    } else {
        // Return. If data processed, go to dashboard. Else upload.
        UI.howItWorksSection.classList.add('hidden');

        if (AppState.processedData.length > 0) {
            UI.dashboardSection.classList.remove('hidden');
        } else {
            UI.uploadSection.classList.remove('hidden');
        }
    }
}

function updateEmailTargets() {
    const type = UI.emailTemplateType.value;
    let options = [];

    // Based on filtered data
    const data = AppState.processedData;

    if (type === 'partner') {
        UI.emailTargetLabel.textContent = "Select Partner";
        options = [...new Set(data.map(d => d.partner).filter(Boolean))].sort();
    } else if (type === 'advertiser') {
        UI.emailTargetLabel.textContent = "Select Advertiser";
        options = [...new Set(data.map(d => d.advertiser).filter(Boolean))].sort();
    } else {
        UI.emailTargetLabel.textContent = "Select Campaign";
        options = data.map(d => d.campaign).filter(Boolean).sort();
    }

    fillSelect(UI.emailTargetSelect, options);
}

function generateEmail() {
    const type = UI.emailTemplateType.value;
    const target = UI.emailTargetSelect.value;
    const style = UI.emailStyle.value; // Get style

    if (!target) {
        UI.emailOutput.value = "Please select a target first.";
        return;
    }

    // Filter data for this email
    let scopeData = AppState.processedData;
    if (type === 'partner') scopeData = scopeData.filter(d => d.partner === target);
    if (type === 'advertiser') scopeData = scopeData.filter(d => d.advertiser === target);
    if (type === 'campaign') scopeData = scopeData.filter(d => d.campaign === target);

    // Only include qualifying opportunities
    const opportunities = scopeData.filter(d => d.calculatedOpportunity > 0);

    if (opportunities.length === 0) {
        const scoreThreshold = UI.filterScore ? UI.filterScore.value : 100;
        UI.emailOutput.value = `No qualifying incremental opportunities found for '${target}' based on current logic (Score > ${scoreThreshold}, Pacing ~100%).`;
        return;
    }

    const totalOpp = opportunities.reduce((acc, curr) => acc + curr.calculatedOpportunity, 0);

    let subject = "";
    let body = "";

    // --- TEMPLATE LOGIC ---
    if (style === 'executive') {
        // EXECUTIVE SUMMARY
        subject = `Executive Summary: Incremental Growth Opportunity - ${target}`;
        body = `Hi Team,\n\nWe have identified a significant incremental budget opportunity of **${formatCurrency(totalOpp)}** across high-performing campaigns for ${target}.\n\n`;
        body += `**Key Highlights:**\n`;
        body += `• Total Opportunity: ${formatCurrency(totalOpp)}\n`;
        body += `• Campaigns Qualifying: ${opportunities.length}\n\n`;
        body += `These campaigns are currently pacing at 100% capacity with high performance scores. Unlocking this budget will directly maximize flight delivery.\n\n`;
        body += `Shall we proceed with this allocation?\n\nBest,\n[Your Name]`;

    } else if (style === 'action') {
        // ACTION / URGENT
        subject = `ACTION REQUIRED: Unlock ${formatCurrency(totalOpp)} for ${target}`;
        body = `Hi everyone,\n\nPerformance alert for ${target}: We are capped on high-value inventory.\n\n`;
        body += `We are leaving **${formatCurrency(totalOpp)}** on the table for campaigns pacing at 100%.\n\n`;
        body += `**Recommended Action:**\n`;
        body += `approve incremental budget for the following ${opportunities.length} campaigns immediately to capture this demand.\n\n`;

        // Brief list
        opportunities.slice(0, 5).forEach(opp => {
            body += `• ${opp.campaign}: +${formatCurrency(opp.calculatedOpportunity)}\n`;
        });
        if (opportunities.length > 5) body += `...and ${opportunities.length - 5} others.\n`;

        body += `\nPlease confirm approval by EOD.\n\nThanks,\n[Your Name]`;

    } else {
        // STANDARD (Detailed)
        subject = `Incremental Opportunity: ${target}`;
        body = `Hi Team,\n\nWe analyzed the current campaign performance for ${target} and identified meaningful incremental opportunities.\n\n`;

        body += ` SUMMARY\n`;
        body += `--------------------------------------------------\n`;
        body += `Total Incremental Opportunity: ${formatCurrency(totalOpp)}\n`;
        body += `Qualifying Campaigns: ${opportunities.length}\n`;
        body += `--------------------------------------------------\n\n`;

        body += `Below are the top campaigns pacing at 100% with high decision power that could utilize additional budget:\n\n`;

        // List top 10
        opportunities.slice(0, 10).forEach(opp => {
            body += `• ${opp.campaign} (${opp.advertiser})\n`;
            body += `  - Opportunity: ${formatCurrency(opp.calculatedOpportunity)}\n`;
            body += `  - Days Remaining: ${opp.daysRemaining}\n`;
            body += `  - Current Pacing: Pacing Well\n`;
            body += `  - Avg. KPI: ${formatKpi(opp.avgKpiValue, opp.kpiType)}\n`;
            body += `  - KPI Goal: ${formatKpi(opp.goalValue, opp.kpiType)}\n`;
            body += `  - KPI Performance: ${formatRatio(opp.kpiPerfRatio)} (Goal Beaten: ${opp.beatingGoalBool ? 'Yes' : 'No'})\n`;
            body += `  - Decision Power Score: ${opp.score}\n\n`;
        });

        if (opportunities.length > 10) {
            body += `...and ${opportunities.length - 10} more.\n\n`;
        }

        body += `We recommend unlocking this incremental budget to maximize performance for the remainder of the flight.\n\n`;
        body += `Please let us know if you'd like to proceed.\n\nBest,\n[Your Name]`;
    }

    UI.emailOutput.value = `Subject: ${subject}\n\n${body}`;
}

function copyEmail() {
    UI.emailOutput.select();
    document.execCommand('copy');

    const original = UI.btnCopyEmail.innerHTML;
    UI.btnCopyEmail.innerHTML = `<i data-lucide="check"></i> Copied`;
    if (window.lucide) lucide.createIcons();
    setTimeout(() => UI.btnCopyEmail.innerHTML = original, 2000);
}


function formatRatio(val) {
    if (val === undefined || val === null || isNaN(val)) return '-';
    // Always convert ratio to percentage: 1.2 -> 120%, 0.8 -> 80%
    return Math.round(val * 100) + '%';
}

function formatNumber(val) {
    if (val === undefined || val === null || isNaN(val)) return '0';
    return new Intl.NumberFormat('en-US').format(val);
}

function formatKpi(val, type) {
    if (!type || val === undefined || val === null || isNaN(val)) return val || '-';

    const t = type.toLowerCase();

    // Currency Types
    if (t.includes('cpa') || t.includes('revenue') || t.includes('cost')) {
        return formatCurrency(val);
    }

    // Percentage Types
    if (t.includes('ctr') || t.includes('vcr') || t.includes('rate')) {
        // Assume raw values are already in correct scale, just append %
        // User example: 0.27 -> 0.27%, 92.7 -> 92.7%
        // We can limit decimals to 2?
        // If integer, don't add decimals.
        return parseFloat(val) + '%';
    }

    // Human Counts
    if (t === 'incremental reach') {
        return '+' + formatNumber(val) + ' people';
    }

    if (t.includes('reach') || t.includes('users') || t.includes('visitors')) {
        return formatNumber(val);
    }

    if (t === 'no goal' || t === 'none' || t === 'n/a') {
        return 'N/A';
    }

    // Default
    return formatNumber(val);
}

// Start
document.addEventListener('DOMContentLoaded', init);
