// --- SETUP ---
let hot = null;
let chartInstance = null;
let confirmCallback = null;
const confirmationModal = document.getElementById('confirmationModal');
const modalMessage = document.getElementById('modalMessage');
const modalConfirmBtn = document.getElementById('modalConfirmBtn');
const modalCancelBtn = document.getElementById('modalCancelBtn');
const spreadsheetContainer = document.getElementById('spreadsheet');

let spreadsheetSaveDebounceTimer = null;
let selectionStatsDebounceTimer = null;
let chartRefreshAfterEditTimer = null;
let handsontableSizingRaf = null;

const SPREADSHEET_SAVE_DEBOUNCE_MS = 420;
const SELECTION_STATS_DEBOUNCE_MS = 90;
const CHART_REFRESH_AFTER_EDIT_MS = 160;

// --- MOTYWY HANDSONTABLE (oficjalne: main, horizon, classic × jasny / ciemny / auto) ---
const HANDSONTABLE_THEME_CLASSES = [
    'ht-theme-main', 'ht-theme-main-dark', 'ht-theme-main-dark-auto',
    'ht-theme-horizon', 'ht-theme-horizon-dark', 'ht-theme-horizon-dark-auto',
    'ht-theme-classic', 'ht-theme-classic-dark', 'ht-theme-classic-dark-auto',
    'ht-theme-custom', 'ht-theme-dark'
];

const HANDSONTABLE_THEMES_META = [
    { id: 'ht-theme-main', label: 'HT Main — jasny' },
    { id: 'ht-theme-main-dark', label: 'HT Main — ciemny' },
    { id: 'ht-theme-main-dark-auto', label: 'HT Main — auto (system)' },
    { id: 'ht-theme-horizon', label: 'HT Horizon — jasny' },
    { id: 'ht-theme-horizon-dark', label: 'HT Horizon — ciemny' },
    { id: 'ht-theme-horizon-dark-auto', label: 'HT Horizon — auto (system)' },
    { id: 'ht-theme-classic', label: 'HT Classic — jasny' },
    { id: 'ht-theme-classic-dark', label: 'HT Classic — ciemny' },
    { id: 'ht-theme-classic-dark-auto', label: 'HT Classic — auto (system)' }
];

const VALID_HANDSONTABLE_THEME_IDS = new Set(HANDSONTABLE_THEMES_META.map((t) => t.id));

let prefersColorSchemeMql = null;

function getShellDataThemeForHandsontable(htThemeClass) {
    if (htThemeClass.endsWith('-dark-auto')) {
        return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }
    if (htThemeClass.endsWith('-dark')) {
        return 'dark';
    }
    return 'light';
}

function onShellThemeFromSystemChange() {
    const ht = localStorage.getItem('handsontableTheme');
    if (!ht || !ht.endsWith('-dark-auto')) return;
    document.documentElement.setAttribute('data-theme', getShellDataThemeForHandsontable(ht));
    updateChart();
}

function bindHandsontableAutoThemeListener(htThemeClass) {
    if (prefersColorSchemeMql) {
        prefersColorSchemeMql.removeEventListener('change', onShellThemeFromSystemChange);
        prefersColorSchemeMql = null;
    }
    if (htThemeClass && htThemeClass.endsWith('-dark-auto')) {
        prefersColorSchemeMql = window.matchMedia('(prefers-color-scheme: dark)');
        prefersColorSchemeMql.addEventListener('change', onShellThemeFromSystemChange);
    }
}

function getStoredHandsontableTheme() {
    const stored = localStorage.getItem('handsontableTheme');
    if (stored && VALID_HANDSONTABLE_THEME_IDS.has(stored)) {
        return stored;
    }
    const legacy = localStorage.getItem('theme');
    const migrated = {
        dark: 'ht-theme-main-dark',
        light: 'ht-theme-main',
        cyberpunk: 'ht-theme-horizon-dark',
        terminal: 'ht-theme-classic-dark'
    }[legacy];
    const next = migrated || 'ht-theme-main-dark';
    localStorage.setItem('handsontableTheme', next);
    return next;
}

function stripHandsontableThemeClasses(el) {
    if (!el) return;
    HANDSONTABLE_THEME_CLASSES.forEach((c) => el.classList.remove(c));
}

function syncThemeSelectorUi(htThemeClass) {
    const themeTrigger = document.getElementById('themeTrigger');
    const meta = HANDSONTABLE_THEMES_META.find((t) => t.id === htThemeClass);
    if (themeTrigger && meta) {
        themeTrigger.querySelector('.select-text').textContent = meta.label;
    }
    document.querySelectorAll('#themeOptions .option').forEach((opt) => {
        opt.classList.toggle('selected', opt.getAttribute('data-value') === htThemeClass);
    });
}

function refreshHotAfterThemeChange() {
    requestAnimationFrame(() => {
        requestAnimationFrame(() => {
            if (!hot) return;
            hot.refreshDimensions();
            hot.render();
            try {
                hot.getPlugin('autoColumnSize')?.recalculateAllColumnsWidth?.();
                hot.getPlugin('autoRowSize')?.recalculateAllRowsHeight?.();
            } catch (_) { /* ignore */ }
            hot.render();
        });
    });
}

function syncShellForHandsontableTheme(htThemeClass) {
    if (!VALID_HANDSONTABLE_THEME_IDS.has(htThemeClass)) return;
    localStorage.setItem('handsontableTheme', htThemeClass);
    document.documentElement.setAttribute('data-theme', getShellDataThemeForHandsontable(htThemeClass));
    stripHandsontableThemeClasses(spreadsheetContainer);
    spreadsheetContainer.classList.add(htThemeClass);
    bindHandsontableAutoThemeListener(htThemeClass);
    syncThemeSelectorUi(htThemeClass);
}

/** Zmiana motywu arkusza (klasy CSS + opcja `theme` w Handsontable). */
const applyHandsontableTheme = (htThemeClass) => {
    if (!VALID_HANDSONTABLE_THEME_IDS.has(htThemeClass)) return;
    syncShellForHandsontableTheme(htThemeClass);
    updateChart();
    if (hot) {
        try {
            hot.updateSettings({ theme: htThemeClass });
        } catch (_) { /* ignore */ }
        refreshHotAfterThemeChange();
    }
};

// Theme selector handling moved to custom select

// --- POLISH TRANSLATIONS ---
// Translations are loaded from trans-pl.js file

// --- CORE FUNCTIONS ---
function cancelSpreadsheetSideTimers() {
    if (spreadsheetSaveDebounceTimer) {
        clearTimeout(spreadsheetSaveDebounceTimer);
        spreadsheetSaveDebounceTimer = null;
    }
    if (chartRefreshAfterEditTimer) {
        clearTimeout(chartRefreshAfterEditTimer);
        chartRefreshAfterEditTimer = null;
    }
    if (selectionStatsDebounceTimer) {
        clearTimeout(selectionStatsDebounceTimer);
        selectionStatsDebounceTimer = null;
    }
    if (handsontableSizingRaf !== null) {
        cancelAnimationFrame(handsontableSizingRaf);
        handsontableSizingRaf = null;
    }
}

/** Przelicza szerokości kolumn i wysokości wierszy (pluginy AutoColumnSize / AutoRowSize). */
function scheduleHandsontableAutoSizingRecalc(instance) {
    if (!instance) return;
    if (handsontableSizingRaf !== null) {
        cancelAnimationFrame(handsontableSizingRaf);
    }
    const target = instance;
    handsontableSizingRaf = requestAnimationFrame(() => {
        handsontableSizingRaf = null;
        if (!hot || target !== hot) return;
        try {
            target.getPlugin('autoColumnSize')?.recalculateAllColumnsWidth?.();
        } catch (_) { /* ignore */ }
        try {
            target.getPlugin('autoRowSize')?.recalculateAllRowsHeight?.();
        } catch (_) { /* ignore */ }
        try {
            target.render();
        } catch (_) { /* ignore */ }
    });
}

function initSpreadsheet(data) {
    cancelSpreadsheetSideTimers();
    if (hot) hot.destroy();
    if (chartInstance) chartInstance.destroy();

    const htTheme = getStoredHandsontableTheme();
    syncShellForHandsontableTheme(htTheme);

    hot = new Handsontable(spreadsheetContainer, {
        data: data || loadFromLocalStorage(),
        rowHeaders: true, colHeaders: true, height: '100%', width: '100%',
        licenseKey: 'non-commercial-and-evaluation', language: 'pl-PL',
        theme: htTheme,
        contextMenu: true, dropdownMenu: true, filters: true, multiColumnSorting: true,
        manualColumnResize: true, manualRowResize: true, manualColumnMove: true, manualRowMove: true,
        mergeCells: true, comments: true, undo: true, redo: true, copyPaste: true,
        fillHandle: { autoInsertRow: true }, search: { searchResultClass: 'hot-search-result' },
        formulas: { engine: HyperFormula },
        autoRowSize: true,
        autoColumnSize: true,
        renderAllRows: false,
        currentRowClassName: 'current-row', currentColClassName: 'current-col', stretchH: 'all',
        afterInit: function () {
            scheduleHandsontableAutoSizingRecalc(this);
        },
        afterLoadData: function () {
            scheduleHandsontableAutoSizingRecalc(this);
        },
        afterChange: (changes, source) => {
            if (source !== 'loadData') {
                scheduleSaveToLocalStorage();
                updateStatusBar();
                scheduleChartRefreshAfterEdit();
                if (typeof updateQuickActionButtons === 'function') updateQuickActionButtons();
            }
        },
        afterSelection: () => { scheduleDebouncedSelectionStats(); },
        afterSelectionEnd: () => { flushSelectionStatsAndChart(); },
        afterCreateRow: () => {
            updateStatusBar();
            scheduleChartRefreshAfterEdit();
            if (typeof updateQuickActionButtons === 'function') updateQuickActionButtons();
        },
        afterRemoveRow: () => {
            updateStatusBar();
            scheduleChartRefreshAfterEdit();
            if (typeof updateQuickActionButtons === 'function') updateQuickActionButtons();
        },
        afterUndo: () => {
            if (typeof updateQuickActionButtons === 'function') updateQuickActionButtons();
            scheduleChartRefreshAfterEdit();
        },
        afterRedo: () => {
            if (typeof updateQuickActionButtons === 'function') updateQuickActionButtons();
            scheduleChartRefreshAfterEdit();
        },
    });
    refreshHotAfterThemeChange();
    updateChart();
    // Update quick action buttons state
    setTimeout(() => {
        if (typeof updateQuickActionButtons === 'function') {
            updateQuickActionButtons();
        }
    }, 100);
}

function saveToLocalStorage() {
    if (!hot) return;
    try {
        localStorage.setItem('spreadsheetData', JSON.stringify(hot.getData()));
    } catch (error) {
        console.error('Błąd podczas zapisywania:', error);
        showNotification('Błąd podczas zapisu lokalnego', true);
    }
}

function scheduleSaveToLocalStorage() {
    if (spreadsheetSaveDebounceTimer) clearTimeout(spreadsheetSaveDebounceTimer);
    spreadsheetSaveDebounceTimer = setTimeout(() => {
        spreadsheetSaveDebounceTimer = null;
        saveToLocalStorage();
    }, SPREADSHEET_SAVE_DEBOUNCE_MS);
}

function flushPendingSpreadsheetSave() {
    if (spreadsheetSaveDebounceTimer) {
        clearTimeout(spreadsheetSaveDebounceTimer);
        spreadsheetSaveDebounceTimer = null;
    }
    saveToLocalStorage();
}

function scheduleDebouncedSelectionStats() {
    if (selectionStatsDebounceTimer) clearTimeout(selectionStatsDebounceTimer);
    selectionStatsDebounceTimer = setTimeout(() => {
        selectionStatsDebounceTimer = null;
        updateStatusBar();
    }, SELECTION_STATS_DEBOUNCE_MS);
}

function flushSelectionStatsAndChart() {
    if (selectionStatsDebounceTimer) {
        clearTimeout(selectionStatsDebounceTimer);
        selectionStatsDebounceTimer = null;
    }
    updateStatusBar();
    updateChart();
}

function scheduleChartRefreshAfterEdit() {
    if (chartRefreshAfterEditTimer) clearTimeout(chartRefreshAfterEditTimer);
    chartRefreshAfterEditTimer = setTimeout(() => {
        chartRefreshAfterEditTimer = null;
        updateChart();
    }, CHART_REFRESH_AFTER_EDIT_MS);
}

function loadFromLocalStorage() {
    const savedData = localStorage.getItem('spreadsheetData');
    if (savedData) {
        try {
            const parsedData = JSON.parse(savedData);
            if (Array.isArray(parsedData) && parsedData.length > 0) return parsedData;
        } catch (error) { console.error('Błąd podczas wczytywania danych:', error); }
    }
    return Handsontable.helper.createEmptySpreadsheetData(50, 20);
}

// --- UI HELPER FUNCTIONS ---
function showLoader() { document.getElementById('loadingOverlay').style.display = 'flex'; }
function hideLoader() { document.getElementById('loadingOverlay').style.display = 'none'; }

function showNotification(message, isError = false) {
    const container = document.getElementById('notification-container');
    const notification = document.createElement('div');
    notification.textContent = message;
    notification.style.cssText = `padding: 0.75rem 1.25rem; border: 1px solid ${isError ? 'var(--red-color)' : 'var(--status-active-bg)'}; background-color: var(--card-bg); color: ${isError ? 'var(--red-color)' : 'var(--status-active-bg)'}; font-size: 0.8rem; font-weight: 600; opacity: 0; transition: all 0.3s ease; transform: translateX(20px);`;
    container.appendChild(notification);
    setTimeout(() => {
        notification.style.opacity = '1';
        notification.style.transform = 'translateX(0)';
    }, 10);
    setTimeout(() => {
        notification.style.opacity = '0';
        notification.addEventListener('transitionend', () => notification.remove());
    }, 3000);
}

function showConfirmationModal(message, callback) {
    modalMessage.textContent = message;
    confirmCallback = callback;
    confirmationModal.style.display = 'flex';
}

function hideConfirmationModal() {
    confirmationModal.style.display = 'none';
    confirmCallback = null;
}

function updateStatusBar() {
    if (!hot) return;
    const numbers = [];
    let selectedCellCount = 0;
    const selected = hot.getSelected();

    if (selected) {
        selected.forEach(range => {
            const [fromRow, fromCol, toRow, toCol] = [Math.min(range[0], range[2]), Math.min(range[1], range[3]), Math.max(range[0], range[2]), Math.max(range[1], range[3])];
            for (let r = fromRow; r <= toRow; r++) {
                for (let c = fromCol; c <= toCol; c++) {
                    selectedCellCount++;
                    const value = hot.getDataAtCell(r, c);
                    const num = parseFloat(String(value).replace(',', '.'));
                    if (!isNaN(num) && value !== null && String(value).trim() !== '') {
                        numbers.push(num);
                    }
                }
            }
        });
    }

    document.getElementById('stat-rowCount').textContent = hot.countRows() - hot.countEmptyRows();
    document.getElementById('stat-selectedCount').textContent = selectedCellCount;

    const locale = 'pl-PL';
    const options = { maximumFractionDigits: 2 };
    if (numbers.length > 0) {
        const sum = numbers.reduce((a, b) => a + b, 0);
        document.getElementById('stat-sum').textContent = sum.toLocaleString(locale, options);
        document.getElementById('stat-avg').textContent = (sum / numbers.length).toLocaleString(locale, options);
        document.getElementById('stat-min').textContent = Math.min(...numbers).toLocaleString(locale, options);
        document.getElementById('stat-max').textContent = Math.max(...numbers).toLocaleString(locale, options);
    } else {
        ['sum', 'avg', 'min', 'max'].forEach(id => {
            document.getElementById(`stat-${id}`).textContent = '-';
        });
    }
}

function buildChartDataFromSelection() {
    if (!hot) return null;
    const selected = hot.getSelected();
    if (!selected || selected.length === 0) return { empty: true, labels: [], dataPoints: [] };

    const range = selected[0];
    const [fromRow, fromCol, toRow, toCol] = [Math.min(range[0], range[2]), Math.min(range[1], range[3]), Math.max(range[0], range[2]), Math.max(range[1], range[3])];

    const labels = [];
    const dataPoints = [];
    const colCount = toCol - fromCol + 1;

    if (colCount >= 2) {
        for (let r = fromRow; r <= toRow; r++) {
            const label = hot.getDataAtCell(r, fromCol);
            const val = parseFloat(String(hot.getDataAtCell(r, fromCol + 1)).replace(',', '.'));
            if (!isNaN(val)) { dataPoints.push(val); labels.push(label); }
        }
    } else if (colCount === 1) {
        for (let r = fromRow; r <= toRow; r++) {
            const val = parseFloat(String(hot.getDataAtCell(r, fromCol)).replace(',', '.'));
            if (!isNaN(val)) { dataPoints.push(val); labels.push(`Wiersz ${hot.getRowHeader(r)}`); }
        }
    }

    if (dataPoints.length === 0) return { empty: true, labels, dataPoints };
    return { empty: false, labels, dataPoints };
}

function updateChart() {
    if (!hot) return;
    const chartPlaceholder = document.getElementById('chart-placeholder');
    const ctx = document.getElementById('dataChart').getContext('2d');
    const chartType = getSelectedChartType();
    const built = buildChartDataFromSelection();

    if (built.empty) {
        if (chartInstance) {
            chartInstance.destroy();
            chartInstance = null;
        }
        chartPlaceholder.style.display = 'flex';
        return;
    }

    const { labels, dataPoints } = built;
    chartPlaceholder.style.display = 'none';

    const style = getComputedStyle(document.documentElement);
    const textColor = style.getPropertyValue('--text-muted').trim();
    const gridColor = style.getPropertyValue('--border-color').trim();
    const highlightColor = style.getPropertyValue('--highlight-color').trim();
    const panelBgColor = style.getPropertyValue('--panel-bg').trim();
    const statusActiveColor = style.getPropertyValue('--status-active-bg').trim();
    const chartColors = [highlightColor, statusActiveColor, textColor, panelBgColor];
    const bgFill = (chartType === 'pie' || chartType === 'doughnut') ? chartColors : `${highlightColor}99`;

    const cartesianScales = {
        y: { beginAtZero: true, ticks: { color: textColor }, grid: { color: gridColor } },
        x: { ticks: { color: textColor }, grid: { color: gridColor } }
    };

    const canPatch = chartInstance
        && chartInstance.config
        && chartInstance.config.type === chartType
        && chartInstance.data.datasets[0];

    if (canPatch) {
        try {
            chartInstance.data.labels = labels;
            chartInstance.data.datasets[0].data = dataPoints;
            chartInstance.data.datasets[0].backgroundColor = bgFill;
            chartInstance.data.datasets[0].borderColor = highlightColor;
            chartInstance.data.datasets[0].pointBackgroundColor = highlightColor;
            if (chartInstance.options.plugins && chartInstance.options.plugins.legend) {
                chartInstance.options.plugins.legend.labels = chartInstance.options.plugins.legend.labels || {};
                chartInstance.options.plugins.legend.labels.color = textColor;
            }
            if (chartType === 'bar' || chartType === 'line') {
                chartInstance.options.scales = { ...cartesianScales };
            } else {
                chartInstance.options.scales = {};
            }
            chartInstance.update('none');
            return;
        } catch (_) {
            chartInstance.destroy();
            chartInstance = null;
        }
    }

    if (chartInstance) chartInstance.destroy();

    chartInstance = new Chart(ctx, {
        type: chartType,
        data: {
            labels,
            datasets: [{
                label: 'Zaznaczone dane',
                data: dataPoints,
                backgroundColor: bgFill,
                borderColor: highlightColor, borderWidth: 2, pointBackgroundColor: highlightColor, tension: 0.2
            }]
        },
        options: {
            animation: false,
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { labels: { color: textColor } }, title: { display: false } },
            scales: (chartType === 'bar' || chartType === 'line') ? cartesianScales : {}
        }
    });
}

function performSearch() {
    if (!hot) return;
    hot.getPlugin('search').query(document.getElementById('searchInput').value);
    hot.render();
}

// --- QUICK ACTIONS FUNCTIONS ---
function getSelectedRange() {
    if (!hot) return null;
    const selected = hot.getSelected();
    return selected && selected.length > 0 ? selected[0] : null;
}

// Removed unused functions - keeping only essential ones

function insertRow() {
    if (!hot) return;
    const range = getSelectedRange();
    const rowIndex = range ? Math.min(range[0], range[2]) : 0;
    hot.alter('insert_row_above', rowIndex);
    showNotification('Dodano wiersz');
}

function insertColumn() {
    if (!hot) return;
    const range = getSelectedRange();
    const colIndex = range ? Math.min(range[1], range[3]) : 0;
    hot.alter('insert_col_start', colIndex);
    showNotification('Dodano kolumnę');
}

function performUndo() {
    if (!hot) return;
    try {
        // Use Handsontable's built-in undo method
        if (hot.undo) {
            hot.undo();
            showNotification('Cofnięto ostatnią akcję');
        } else {
            // Try to trigger undo via keyboard shortcut simulation
            const event = new KeyboardEvent('keydown', {
                key: 'z',
                ctrlKey: true,
                bubbles: true
            });
            hot.rootElement.dispatchEvent(event);
            showNotification('Cofnięto ostatnią akcję');
        }
    } catch (error) {
        console.error('Undo error:', error);
        showNotification('Błąd podczas cofania', true);
    }
}

function performRedo() {
    if (!hot) return;
    try {
        // Use Handsontable's built-in redo method
        if (hot.redo) {
            hot.redo();
            showNotification('Ponowiono akcję');
        } else {
            // Try to trigger redo via keyboard shortcut simulation
            const event = new KeyboardEvent('keydown', {
                key: 'y',
                ctrlKey: true,
                bubbles: true
            });
            hot.rootElement.dispatchEvent(event);
            showNotification('Ponowiono akcję');
        }
    } catch (error) {
        console.error('Redo error:', error);
        showNotification('Błąd podczas ponowienia', true);
    }
}

// Removed copy/paste functions - keeping only essential ones

function updateQuickActionButtons() {
    if (!hot) return;

    try {
        // Enable all buttons - they will handle their own validation
        const insertRowBtn = document.getElementById('insertRowBtn');
        const insertColBtn = document.getElementById('insertColBtn');
        const undoBtn = document.getElementById('undoBtn');
        const redoBtn = document.getElementById('redoBtn');

        if (insertRowBtn) insertRowBtn.disabled = false;
        if (insertColBtn) insertColBtn.disabled = false;
        if (undoBtn) undoBtn.disabled = false;
        if (redoBtn) redoBtn.disabled = false;
    } catch (error) {
        console.warn('Error updating quick action buttons:', error);
    }
}

// --- EVENT LISTENERS ---
document.getElementById('importBtn').addEventListener('click', () => document.getElementById('fileInput').click());
document.getElementById('fileInput').addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (!file) return;
    showLoader();
    Papa.parse(file, {
        complete: (results) => {
            if (results.data?.length) initSpreadsheet(results.data);
            showNotification('Plik CSV zaimportowany pomyślnie.');
            hideLoader();
        },
        error: () => { showNotification('Błąd podczas wczytywania pliku CSV', true); hideLoader(); }
    });
    event.target.value = '';
});

document.getElementById('exportBtn').addEventListener('click', () => hot?.getPlugin('exportFile').downloadFile('csv', { filename: `Arkusz_${new Date().toISOString().slice(0, 10)}` }));
document.getElementById('saveBtn').addEventListener('click', () => {
    flushPendingSpreadsheetSave();
    showNotification('Dane zapisane w przeglądarce.');
});

document.getElementById('clearBtn').addEventListener('click', () => {
    showConfirmationModal('Czy na pewno chcesz wyczyścić cały arkusz?', () => {
        initSpreadsheet(Handsontable.helper.createEmptySpreadsheetData(50, 20));
        localStorage.removeItem('spreadsheetData');
        showNotification('Arkusz został wyczyszczony.');
    });
});

document.getElementById('searchBtn').addEventListener('click', performSearch);
document.getElementById('searchInput').addEventListener('keydown', (e) => e.key === 'Enter' && performSearch());
// Chart type handling moved to custom select
modalConfirmBtn.addEventListener('click', () => { if (confirmCallback) confirmCallback(); hideConfirmationModal(); });
modalCancelBtn.addEventListener('click', hideConfirmationModal);

// --- QUICK ACTIONS EVENT LISTENERS ---
// Safely add event listeners only if elements exist
const addQuickActionListener = (id, handler) => {
    const element = document.getElementById(id);
    if (element) {
        element.addEventListener('click', handler);
    } else {
        console.warn(`Element with id '${id}' not found`);
    }
};

// Essential functions
addQuickActionListener('insertRowBtn', insertRow);
addQuickActionListener('insertColBtn', insertColumn);
addQuickActionListener('undoBtn', performUndo);
addQuickActionListener('redoBtn', performRedo);

// --- CUSTOM SELECT HANDLERS ---
function initCustomSelects() {
    // Initialize default selections
    document.querySelector('#chartTypeOptions .option[data-value="bar"]')?.classList.add('selected');

    // Chart Type Select
    const chartTypeSelect = document.getElementById('chartTypeSelect');
    const chartTypeTrigger = document.getElementById('chartTypeTrigger');
    const chartTypeOptions = document.getElementById('chartTypeOptions');

    if (chartTypeTrigger && chartTypeOptions) {
        chartTypeTrigger.addEventListener('click', (e) => {
            e.stopPropagation();
            chartTypeTrigger.classList.toggle('active');
            chartTypeOptions.classList.toggle('show');

            // Close other dropdowns
            document.getElementById('themeTrigger')?.classList.remove('active');
            document.getElementById('themeOptions')?.classList.remove('show');
        });

        chartTypeOptions.addEventListener('click', (e) => {
            if (e.target.classList.contains('option')) {
                const value = e.target.getAttribute('data-value');
                const text = e.target.textContent.trim();

                // Update trigger text
                chartTypeTrigger.querySelector('.select-text').textContent = text;

                // Remove previous selected
                chartTypeOptions.querySelectorAll('.option').forEach(opt => opt.classList.remove('selected'));
                e.target.classList.add('selected');

                // Close dropdown
                chartTypeTrigger.classList.remove('active');
                chartTypeOptions.classList.remove('show');

                // Update chart
                updateChart();
            }
        });
    }

    // Theme Select
    const themeSelect = document.getElementById('themeSelect');
    const themeTrigger = document.getElementById('themeTrigger');
    const themeOptions = document.getElementById('themeOptions');

    if (themeTrigger && themeOptions) {
        themeTrigger.addEventListener('click', (e) => {
            e.stopPropagation();
            themeTrigger.classList.toggle('active');
            themeOptions.classList.toggle('show');

            // Close other dropdowns
            document.getElementById('chartTypeTrigger')?.classList.remove('active');
            document.getElementById('chartTypeOptions')?.classList.remove('show');
        });

        themeOptions.addEventListener('click', (e) => {
            if (e.target.classList.contains('option')) {
                const value = e.target.getAttribute('data-value');
                const text = e.target.textContent.trim();

                // Update trigger text
                themeTrigger.querySelector('.select-text').textContent = text;

                // Remove previous selected
                themeOptions.querySelectorAll('.option').forEach(opt => opt.classList.remove('selected'));
                e.target.classList.add('selected');

                // Close dropdown
                themeTrigger.classList.remove('active');
                themeOptions.classList.remove('show');

                // Apply theme
                applyHandsontableTheme(value);
            }
        });
    }

    // Close dropdowns when clicking outside
    document.addEventListener('click', () => {
        document.querySelectorAll('.select-trigger').forEach(trigger => {
            trigger.classList.remove('active');
        });
        document.querySelectorAll('.select-options').forEach(options => {
            options.classList.remove('show');
        });
    });

    syncThemeSelectorUi(getStoredHandsontableTheme());
}

function getSelectedChartType() {
    const selected = document.querySelector('#chartTypeOptions .option.selected');
    return selected ? selected.getAttribute('data-value') : 'bar';
}

// --- DYNAMIC FORM SYSTEM ---
class DynamicForm {
    constructor(containerId) {
        this.container = document.getElementById(containerId) || document.body;
        this.components = [];
        this.panels = [];
    }

    // Create panel with header and content
    createPanel(config) {
        const panel = document.createElement('div');
        panel.className = 'panel';
        if (config.gridStyle) {
            panel.style.cssText = config.gridStyle;
        }

        // Panel header
        if (config.header) {
            const header = document.createElement('div');
            header.className = `panel-header ${config.headerClass || ''}`;
            header.innerHTML = config.header.icon ?
                `<i class="${config.header.icon}"></i> ${config.header.title}` :
                config.header.title;
            panel.appendChild(header);
        }

        // Panel content
        const content = document.createElement('div');
        content.className = 'panel-content';
        if (config.contentClass) {
            content.classList.add(config.contentClass);
        }
        if (config.contentStyle) {
            content.style.cssText = config.contentStyle;
        }

        panel.appendChild(content);
        this.container.appendChild(panel);

        const panelObj = { element: panel, content, config, id: config.id };
        this.panels.push(panelObj);
        return panelObj;
    }

    // Create toolbar with buttons
    createToolbar(panel, buttons) {
        const toolbar = document.createElement('div');
        toolbar.className = 'toolbar';

        buttons.forEach(btn => {
            if (btn.type === 'file') {
                const input = document.createElement('input');
                input.type = 'file';
                input.id = btn.id;
                input.accept = btn.accept || '';
                input.style.display = 'none';
                toolbar.appendChild(input);
            } else if (btn.type === 'search') {
                const input = document.createElement('input');
                input.type = 'search';
                input.id = btn.id;
                input.placeholder = btn.placeholder || '';
                input.className = btn.className || '';
                toolbar.appendChild(input);
            } else {
                const button = document.createElement('button');
                button.id = btn.id;
                button.className = btn.className || 'btn';
                button.innerHTML = btn.icon ?
                    `<i class="${btn.icon}"></i> ${btn.text}` :
                    btn.text;
                if (btn.title) button.title = btn.title;
                if (btn.onClick) button.addEventListener('click', btn.onClick);
                toolbar.appendChild(button);
            }
        });

        panel.content.appendChild(toolbar);
        return toolbar;
    }

    // Create quick actions bar
    createQuickActions(panel, groups) {
        const quickBar = document.createElement('div');
        quickBar.className = 'quick-actions-bar';

        groups.forEach(group => {
            const actionGroup = document.createElement('div');
            actionGroup.className = 'action-group';

            if (group.label) {
                const label = document.createElement('span');
                label.className = 'group-label';
                label.textContent = group.label;
                actionGroup.appendChild(label);
            }

            group.actions.forEach(action => {
                const button = document.createElement('button');
                button.id = action.id;
                button.className = action.className || 'quick-btn';
                button.title = action.title || '';

                if (action.icons && Array.isArray(action.icons)) {
                    button.innerHTML = action.icons.map(icon => `<i class="${icon}"></i>`).join(' ');
                } else if (action.icon) {
                    button.innerHTML = `<i class="${action.icon}"></i>`;
                }

                if (action.onClick) button.addEventListener('click', action.onClick);
                actionGroup.appendChild(button);
            });

            quickBar.appendChild(actionGroup);
        });

        panel.content.appendChild(quickBar);
        return quickBar;
    }

    // Create custom select dropdown
    createCustomSelect(panel, config) {
        const selectWrapper = document.createElement('div');
        selectWrapper.className = 'custom-select';
        selectWrapper.id = config.id;

        const trigger = document.createElement('div');
        trigger.className = 'select-trigger';
        trigger.id = config.triggerId;
        trigger.innerHTML = `
            <span class="select-text">${config.defaultText}</span>
            <i class="fas fa-chevron-down"></i>
        `;

        const options = document.createElement('div');
        options.className = 'select-options';
        options.id = config.optionsId;

        config.options.forEach(opt => {
            const option = document.createElement('div');
            option.className = 'option';
            option.setAttribute('data-value', opt.value);
            option.innerHTML = opt.icon ?
                `<i class="${opt.icon}"></i> ${opt.text}` :
                opt.text;
            options.appendChild(option);
        });

        selectWrapper.appendChild(trigger);
        selectWrapper.appendChild(options);
        panel.content.appendChild(selectWrapper);

        // Add event handlers
        this.initSelectHandlers(trigger, options, config.onChange);

        return selectWrapper;
    }

    // Create stats grid
    createStatsGrid(panel, stats) {
        const statsGrid = document.createElement('div');
        statsGrid.className = 'stats-grid';

        stats.forEach(stat => {
            const card = document.createElement('div');
            card.className = 'info-card';

            const header = document.createElement('div');
            header.className = `card-header ${stat.headerClass || 'info-badge'}`;
            header.innerHTML = stat.icon ?
                `<i class="${stat.icon}"></i> ${stat.label}` :
                stat.label;

            const title = document.createElement('div');
            title.className = `card-title ${stat.titleClass || 'light-badge'}`;
            title.id = stat.id;
            title.textContent = stat.defaultValue || '-';

            card.appendChild(header);
            card.appendChild(title);
            statsGrid.appendChild(card);
        });

        panel.content.appendChild(statsGrid);
        return statsGrid;
    }

    // Initialize select handlers
    initSelectHandlers(trigger, options, onChange) {
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            trigger.classList.toggle('active');
            options.classList.toggle('show');

            // Close other dropdowns
            document.querySelectorAll('.select-trigger').forEach(t => {
                if (t !== trigger) t.classList.remove('active');
            });
            document.querySelectorAll('.select-options').forEach(o => {
                if (o !== options) o.classList.remove('show');
            });
        });

        options.addEventListener('click', (e) => {
            if (e.target.classList.contains('option')) {
                const value = e.target.getAttribute('data-value');
                const text = e.target.textContent.trim();

                trigger.querySelector('.select-text').textContent = text;
                options.querySelectorAll('.option').forEach(opt => opt.classList.remove('selected'));
                e.target.classList.add('selected');

                trigger.classList.remove('active');
                options.classList.remove('show');

                if (onChange) onChange(value, text);
            }
        });
    }

    // Create complete spreadsheet layout
    createSpreadsheetLayout() {
        // Main grid container
        const grid = document.createElement('div');
        grid.className = 'command-center-grid';
        this.container.appendChild(grid);

        // Main spreadsheet panel
        const mainPanel = this.createPanel({
            id: 'main-panel',
            gridStyle: 'grid-column: 1 / 3; grid-row: 1 / 2;',
            header: { icon: 'fas fa-table', title: 'Arkusz Danych' }
        });
        grid.appendChild(mainPanel.element);

        // Stats panel
        const statsPanel = this.createPanel({
            id: 'stats-panel',
            gridStyle: 'grid-column: 1 / 2; grid-row: 2 / 3;',
            header: { icon: 'fas fa-calculator', title: 'Statystyki Zaznaczenia' },
            headerClass: 'panel-header-secondary',
            contentClass: 'stats-grid'
        });
        grid.appendChild(statsPanel.element);

        // Visualization panel
        const vizPanel = this.createPanel({
            id: 'viz-panel',
            gridStyle: 'grid-column: 2 / 3; grid-row: 2 / 3;',
            header: { icon: 'fas fa-chart-bar', title: 'Wizualizacja Danych' },
            headerClass: 'panel-header-secondary',
            contentStyle: 'display: flex; flex-direction: column; gap: 0.5rem;'
        });
        grid.appendChild(vizPanel.element);

        return { mainPanel, statsPanel, vizPanel, grid };
    }

    // Get panel by ID
    getPanel(id) {
        return this.panels.find(p => p.config.id === id);
    }

    // Remove panel
    removePanel(id) {
        const panelIndex = this.panels.findIndex(p => p.config.id === id);
        if (panelIndex > -1) {
            this.panels[panelIndex].element.remove();
            this.panels.splice(panelIndex, 1);
        }
    }

    // Update panel content
    updatePanel(id, newContent) {
        const panel = this.getPanel(id);
        if (panel) {
            panel.content.innerHTML = '';
            panel.content.appendChild(newContent);
        }
    }
}

// Initialize dynamic form system
function initDynamicSystem() {
    const form = new DynamicForm('app-container');

    // Create complete layout
    const layout = form.createSpreadsheetLayout();

    // Add toolbar to main panel
    form.createToolbar(layout.mainPanel, [
        { type: 'file', id: 'fileInput', accept: '.csv' },
        { id: 'importBtn', className: 'btn', icon: 'fas fa-file-import', text: 'Importuj CSV', onClick: () => document.getElementById('fileInput').click() },
        { id: 'exportBtn', className: 'btn', icon: 'fas fa-file-export', text: 'Eksportuj CSV' },
        { id: 'saveBtn', className: 'btn', icon: 'fas fa-save', text: 'Zapisz lokalnie' },
        { id: 'clearBtn', className: 'btn danger', icon: 'fas fa-trash', text: 'Wyczyść arkusz' },
        { type: 'search', id: 'searchInput', placeholder: 'Szukaj...' },
        { id: 'searchBtn', className: 'btn', icon: 'fas fa-search', text: 'Szukaj' }
    ]);

    // Add quick actions
    form.createQuickActions(layout.mainPanel, [
        {
            label: 'Dodaj:',
            actions: [
                { id: 'insertRowBtn', title: 'Wstaw wiersz', icons: ['fas fa-plus', 'fas fa-table-rows'], onClick: insertRow },
                { id: 'insertColBtn', title: 'Wstaw kolumnę', icons: ['fas fa-plus', 'fas fa-table-columns'], onClick: insertColumn }
            ]
        },
        {
            label: 'Akcje:',
            actions: [
                { id: 'undoBtn', title: 'Cofnij', icon: 'fas fa-undo', onClick: performUndo },
                { id: 'redoBtn', title: 'Ponów', icon: 'fas fa-redo', onClick: performRedo }
            ]
        }
    ]);

    // Add spreadsheet container
    const spreadsheetDiv = document.createElement('div');
    spreadsheetDiv.id = 'spreadsheet';
    layout.mainPanel.content.appendChild(spreadsheetDiv);

    // Add stats to stats panel
    form.createStatsGrid(layout.statsPanel, [
        { id: 'stat-sum', label: 'Suma', icon: 'fas fa-plus' },
        { id: 'stat-avg', label: 'Średnia', icon: 'fas fa-divide' },
        { id: 'stat-min', label: 'Minimum', icon: 'fas fa-arrow-down' },
        { id: 'stat-max', label: 'Maksimum', icon: 'fas fa-arrow-up' },
        { id: 'stat-selectedCount', label: 'Zaznaczone Komórki', icon: 'fas fa-mouse-pointer', defaultValue: '0' },
        { id: 'stat-rowCount', label: 'Wiersze w Użyciu', icon: 'fas fa-list', defaultValue: '0' }
    ]);

    // Add chart selector to viz panel
    form.createCustomSelect(layout.vizPanel, {
        id: 'chartTypeSelect',
        triggerId: 'chartTypeTrigger',
        optionsId: 'chartTypeOptions',
        defaultText: 'Słupkowy',
        options: [
            { value: 'bar', text: 'Słupkowy', icon: 'fas fa-chart-column' },
            { value: 'line', text: 'Liniowy', icon: 'fas fa-chart-line' },
            { value: 'pie', text: 'Kołowy', icon: 'fas fa-chart-pie' },
            { value: 'doughnut', text: 'Pierścieniowy', icon: 'fas fa-circle-notch' }
        ],
        onChange: (value) => updateChart()
    });

    // Add chart container
    const chartContainer = document.createElement('div');
    chartContainer.id = 'chart-container';
    chartContainer.style.cssText = 'flex-grow: 1; position: relative;';
    chartContainer.innerHTML = `
        <canvas id="dataChart"></canvas>
        <div id="chart-placeholder" style="position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; color: var(--text-muted); font-size: 0.8rem; text-align: center; padding: 1rem;">
            Zaznacz dane numeryczne, aby wygenerować wykres.
        </div>
    `;
    layout.vizPanel.content.appendChild(chartContainer);

    return form;
}

// Example: Create additional panels dynamically
function createAdditionalPanels() {
    const form = new DynamicForm('body');

    // Create settings panel
    const settingsPanel = form.createPanel({
        id: 'settings-panel',
        gridStyle: 'position: fixed; top: 100px; right: 20px; width: 300px; z-index: 1000;',
        header: { icon: 'fas fa-cog', title: 'Ustawienia' },
        headerClass: 'panel-header-secondary'
    });

    // Add settings controls
    form.createCustomSelect(settingsPanel, {
        id: 'languageSelect',
        triggerId: 'languageTrigger',
        optionsId: 'languageOptions',
        defaultText: 'Polski',
        options: [
            { value: 'pl', text: 'Polski', icon: 'fas fa-flag' },
            { value: 'en', text: 'English', icon: 'fas fa-flag' },
            { value: 'de', text: 'Deutsch', icon: 'fas fa-flag' }
        ],
        onChange: (value) => console.log('Language changed:', value)
    });

    // Create notification panel
    const notificationPanel = form.createPanel({
        id: 'notification-panel',
        gridStyle: 'position: fixed; bottom: 20px; left: 20px; width: 350px; max-height: 200px; overflow-y: auto; z-index: 1001;',
        header: { icon: 'fas fa-bell', title: 'Powiadomienia' },
        headerClass: 'panel-header-secondary'
    });

    // Add status badges to notification panel
    const notificationContent = document.createElement('div');
    notificationContent.innerHTML = `
        <div class="status-badge success" style="margin: 0.25rem; width: calc(100% - 0.5rem);">
            <i class="fas fa-check"></i> Dane zapisane pomyślnie
        </div>
        <div class="status-badge info" style="margin: 0.25rem; width: calc(100% - 0.5rem);">
            <i class="fas fa-info"></i> Nowy wiersz dodany
        </div>
        <div class="status-badge warning" style="margin: 0.25rem; width: calc(100% - 0.5rem);">
            <i class="fas fa-exclamation"></i> Sprawdź format danych
        </div>
    `;
    notificationPanel.content.appendChild(notificationContent);

    return form;
}

// Example: Dynamic form builder
function createCustomForm(containerId, formConfig) {
    const form = new DynamicForm(containerId);

    formConfig.panels.forEach(panelConfig => {
        const panel = form.createPanel(panelConfig);

        // Add components based on config
        if (panelConfig.toolbar) {
            form.createToolbar(panel, panelConfig.toolbar);
        }

        if (panelConfig.quickActions) {
            form.createQuickActions(panel, panelConfig.quickActions);
        }

        if (panelConfig.customSelects) {
            panelConfig.customSelects.forEach(selectConfig => {
                form.createCustomSelect(panel, selectConfig);
            });
        }

        if (panelConfig.statsGrid) {
            form.createStatsGrid(panel, panelConfig.statsGrid);
        }

        // Add custom content
        if (panelConfig.customContent) {
            panel.content.appendChild(panelConfig.customContent);
        }
    });

    return form;
}

// Global dynamic form instance
let globalDynamicForm = null;

const WORKSPACE_CHART_WIDTH_KEY = 'workspaceChartWidthPx';

/** Ustawiane w initWorkspaceSplitter — odświeża szerokość po zmianie okna. */
let workspaceSplitRefit = () => { };

function initWorkspaceSplitter() {
    const grid = document.querySelector('.command-center-grid');
    const splitter = document.getElementById('workspaceSplitter');
    const panelChart = document.querySelector('.panel-chart');
    if (!grid || !splitter || !panelChart) return;

    const mq = window.matchMedia('(max-width: 1024px)');
    const MIN_MAIN = 260;
    const MIN_CHART = 220;
    const MAX_CHART_FR = 0.78;

    let dragRaf = null;
    let pendingDragPx = null;
    let keyCommitTimer = null;

    function readStoredOrDefault() {
        const raw = localStorage.getItem(WORKSPACE_CHART_WIDTH_KEY);
        if (raw) {
            const n = parseInt(raw, 10);
            if (!Number.isNaN(n) && n > 0) return n;
        }
        return Math.round(grid.clientWidth * 0.28);
    }

    function clampChartWidth(px) {
        const maxW = Math.max(MIN_CHART, Math.floor(grid.clientWidth * MAX_CHART_FR));
        const maxByMain = Math.max(MIN_CHART, grid.clientWidth - MIN_MAIN - splitter.offsetWidth - 8);
        const cap = Math.min(maxW, maxByMain);
        return Math.max(MIN_CHART, Math.min(Math.round(px), cap));
    }

    /** Tylko układ DOM — tanie, można wołać w pętli przy przeciąganiu. */
    function paintChartWidth(px) {
        if (mq.matches) {
            panelChart.style.flex = '';
            panelChart.style.maxWidth = '';
            return null;
        }
        const w = clampChartWidth(px);
        panelChart.style.flex = `0 0 ${w}px`;
        panelChart.style.maxWidth = `${Math.round(grid.clientWidth * MAX_CHART_FR)}px`;
        return w;
    }

    /** Po zakończeniu przeciągania / starcie: zapis + Handsontable + wykres. */
    function commitWorkspaceWidth(px) {
        if (mq.matches) {
            panelChart.style.flex = '';
            panelChart.style.maxWidth = '';
            return null;
        }
        const w = paintChartWidth(px);
        if (w === null) return null;
        localStorage.setItem(WORKSPACE_CHART_WIDTH_KEY, String(w));
        hot?.render();
        chartInstance?.resize?.();
        return w;
    }

    workspaceSplitRefit = () => {
        if (mq.matches) {
            panelChart.style.flex = '';
            panelChart.style.maxWidth = '';
            return;
        }
        const rectW = panelChart.getBoundingClientRect().width;
        const w = paintChartWidth(rectW > MIN_CHART ? rectW : readStoredOrDefault());
        if (w !== null) {
            localStorage.setItem(WORKSPACE_CHART_WIDTH_KEY, String(w));
        }
    };

    function applyInitialDesktopWidth() {
        if (mq.matches) return;
        commitWorkspaceWidth(readStoredOrDefault());
    }

    applyInitialDesktopWidth();

    mq.addEventListener('change', () => {
        if (mq.matches) {
            panelChart.style.flex = '';
            panelChart.style.maxWidth = '';
        } else {
            commitWorkspaceWidth(readStoredOrDefault());
        }
        hot?.render();
        chartInstance?.resize?.();
    });

    let dragging = false;
    let startX = 0;
    let startWidth = 0;

    function flushDragPaint() {
        if (dragRaf !== null) {
            cancelAnimationFrame(dragRaf);
            dragRaf = null;
        }
        if (pendingDragPx !== null) {
            paintChartWidth(pendingDragPx);
            pendingDragPx = null;
        }
    }

    function stopDrag(e) {
        if (!dragging) return;
        dragging = false;
        splitter.classList.remove('is-dragging');
        document.body.classList.remove('is-resizing-workspace');
        flushDragPaint();
        if (e && typeof e.pointerId === 'number') {
            try {
                splitter.releasePointerCapture(e.pointerId);
            } catch (_) { /* ignore */ }
        }
    }

    function endDragCommit() {
        if (mq.matches) return;
        const w = clampChartWidth(panelChart.getBoundingClientRect().width);
        localStorage.setItem(WORKSPACE_CHART_WIDTH_KEY, String(w));
        hot?.render();
        chartInstance?.resize?.();
    }

    splitter.addEventListener('pointerdown', (e) => {
        if (e.button !== 0 || mq.matches) return;
        dragging = true;
        pendingDragPx = null;
        if (dragRaf !== null) {
            cancelAnimationFrame(dragRaf);
            dragRaf = null;
        }
        startX = e.clientX;
        startWidth = panelChart.getBoundingClientRect().width;
        splitter.classList.add('is-dragging');
        document.body.classList.add('is-resizing-workspace');
        splitter.setPointerCapture(e.pointerId);
        e.preventDefault();
    });

    splitter.addEventListener('pointermove', (e) => {
        if (!dragging) return;
        pendingDragPx = startWidth + (e.clientX - startX);
        if (dragRaf !== null) return;
        dragRaf = requestAnimationFrame(() => {
            dragRaf = null;
            if (pendingDragPx !== null) {
                paintChartWidth(pendingDragPx);
            }
        });
    });

    splitter.addEventListener('pointerup', (e) => {
        const was = dragging;
        stopDrag(e);
        if (was) endDragCommit();
    });

    splitter.addEventListener('pointercancel', (e) => {
        const was = dragging;
        stopDrag(e);
        if (was) endDragCommit();
    });

    splitter.addEventListener('keydown', (e) => {
        if (mq.matches) return;
        const step = e.shiftKey ? 40 : 16;
        let next = null;
        if (e.key === 'ArrowLeft') next = panelChart.getBoundingClientRect().width - step;
        else if (e.key === 'ArrowRight') next = panelChart.getBoundingClientRect().width + step;
        if (next === null) return;
        e.preventDefault();
        paintChartWidth(next);
        clearTimeout(keyCommitTimer);
        keyCommitTimer = setTimeout(() => {
            keyCommitTimer = null;
            endDragCommit();
        }, 120);
    });
}

window.addEventListener('load', () => {
    initSpreadsheet();
    initWorkspaceSplitter();
    initCustomSelects();
    initContextMenu(); // Initialize the context menu system
    initRippleEffects(); // Initialize ripple effects

    // Example usage - uncomment to test:
    // globalDynamicForm = createAdditionalPanels();

    // Or use complete dynamic system instead of static HTML:
    // const dynamicForm = initDynamicSystem();
});

let workspaceWindowResizeRaf = null;
window.addEventListener('resize', () => {
    if (workspaceWindowResizeRaf !== null) return;
    workspaceWindowResizeRaf = requestAnimationFrame(() => {
        workspaceWindowResizeRaf = null;
        hot?.render();
        chartInstance?.resize?.();
        workspaceSplitRefit();
    });
});

window.addEventListener('beforeunload', () => {
    flushPendingSpreadsheetSave();
});
window.addEventListener('pagehide', () => {
    flushPendingSpreadsheetSave();
});

// Export for global use
window.DynamicForm = DynamicForm;
window.createCustomForm = createCustomForm;

// --- RIPPLE EFFECT SYSTEM ---
function createRipple(event) {
    const button = event.currentTarget;
    const ripple = document.createElement('span');
    const rect = button.getBoundingClientRect();
    const size = Math.max(rect.width, rect.height);
    const x = event.clientX - rect.left - size / 2;
    const y = event.clientY - rect.top - size / 2;

    ripple.className = 'ripple';
    ripple.style.width = ripple.style.height = `${size}px`;
    ripple.style.left = `${x}px`;
    ripple.style.top = `${y}px`;

    button.appendChild(ripple);

    ripple.addEventListener('animationend', () => {
        ripple.remove();
    });
}

// Initialize ripple effects
function initRippleEffects() {
    // Add ripple to buttons using event delegation
    const buttonSelectors = [
        '.btn',
        '.quick-btn',
        '.about-modal-close',
        '.context-menu-item'
    ];

    // Use single event listener with delegation for better performance
    document.addEventListener('click', (e) => {
        buttonSelectors.forEach(selector => {
            const button = e.target.closest(selector);
            if (button) {
                // Only add ripple if it doesn't already have one in progress
                if (!button.querySelector('.ripple')) {
                    // Create a new event object with the button as currentTarget
                    const rippleEvent = {
                        currentTarget: button,
                        clientX: e.clientX,
                        clientY: e.clientY
                    };
                    createRipple(rippleEvent);
                }
            }
        });
    });
}

// --- CONTEXT MENU SYSTEM ---
// Enhanced notification system
function showContextNotification(message, type = 'info', duration = 3000) {
    const notification = document.createElement('div');
    notification.classList.add('notification');
    notification.textContent = message;

    // Different colors for different notification types
    switch (type) {
        case 'success':
            notification.style.backgroundColor = 'rgba(46, 204, 113, 0.9)';
            break;
        case 'error':
            notification.style.backgroundColor = 'rgba(231, 76, 60, 0.9)';
            break;
        case 'warning':
            notification.style.backgroundColor = 'rgba(241, 196, 15, 0.9)';
            break;
        default:
            notification.style.backgroundColor = 'rgba(0,0,0,0.8)';
    }

    document.body.appendChild(notification);

    setTimeout(() => {
        notification.style.animation = 'slideOut 0.3s ease-out';
        setTimeout(() => {
            if (document.body.contains(notification)) {
                document.body.removeChild(notification);
            }
        }, 300);
    }, duration);
}

// Dynamic menu positioning function
function calculateContextMenuPosition(e, menuWidth, menuHeight) {
    const screenWidth = window.innerWidth;
    const screenHeight = window.innerHeight;

    let left = e.clientX;
    let top = e.clientY;

    // Adjust horizontal position
    left = Math.min(left, screenWidth - menuWidth - 10);
    left = Math.max(10, left);

    // Adjust vertical position
    top = Math.min(top, screenHeight - menuHeight - 10);
    top = Math.max(10, top);

    return { left, top };
}

// Initialize context menu system
function initContextMenu() {
    document.addEventListener('contextmenu', function (e) {
        e.preventDefault();

        // Remove previous menu
        const oldMenu = document.getElementById('custom-context-menu');
        if (oldMenu) oldMenu.remove();

        // Create new menu
        const menu = document.createElement('div');
        menu.id = 'custom-context-menu';
        menu.classList.add('context-menu');

        // Render menu to calculate dimensions
        document.body.appendChild(menu);

        const currentHtTheme = getStoredHandsontableTheme();
        const handsonTableThemeMenuItems = HANDSONTABLE_THEMES_META.map((meta) => ({
            text: `${meta.label}${currentHtTheme === meta.id ? ' ✓' : ''}`,
            icon: 'fas fa-palette',
            isActive: currentHtTheme === meta.id,
            action: () => {
                applyHandsontableTheme(meta.id);
                showNotification(`Motyw arkusza: ${meta.label}`);
            }
        }));

        // Menu options
        const options = [
            {
                group: 'Extensions',
                icon: 'fas fa-puzzle-piece',
                items: [
                    {
                        text: 'uBlock Origin',
                        icon: 'mdi mdi-shield-check',
                        action: () => {
                            window.open('https://addons.mozilla.org/pl/firefox/addon/ublock-origin/?utm_source=addons.mozilla.org&utm_medium=referral&utm_content=search', '_blank');
                            showNotification('Otwarto uBlock Origin');
                        }
                    },
                    {
                        text: 'Dark Reader',
                        icon: 'mdi mdi-weather-night',
                        action: () => {
                            window.open('https://addons.mozilla.org/pl/firefox/addon/darkreader/?utm_source=addons.mozilla.org&utm_medium=referral&utm_content=search', '_blank');
                            showNotification('Otwarto Dark Reader');
                        }
                    },
                    {
                        text: 'Stylus',
                        icon: 'mdi mdi-brush',
                        action: () => {
                            window.open('https://addons.mozilla.org/pl/firefox/addon/styl-us/?utm_source=addons.mozilla.org&utm_medium=referral&utm_content=search', '_blank');
                            showNotification('Otwarto Stylus');
                        }
                    }
                ]
            },
            {
                group: 'Tools Event',
                icon: 'fas fa-calendar-check',
                items: [
                    {
                        text: 'INCSV',
                        icon: 'mdi mdi-file-import',
                        action: () => {
                            window.open('https://incsv.carrd.co/', '_blank');
                            showNotification('Otwarto INCSV');
                        }
                    }
                ]
            },
            {
                group: 'Fast List Print',
                icon: 'fas fa-print',
                items: [
                    {
                        text: 'W5E',
                        icon: 'mdi mdi-web',
                        action: () => {
                            window.open('https://w5e.carrd.co/', '_blank');
                            showNotification('Otwarto W5E');
                        }
                    }
                ]
            },
            {
                group: 'Spreadsheet Tools',
                icon: 'fas fa-table',
                items: [
                    {
                        text: 'Export CSV',
                        icon: 'mdi mdi-file-export',
                        action: () => {
                            if (hot) {
                                hot.getPlugin('exportFile').downloadFile('csv', { filename: `Arkusz_${new Date().toISOString().slice(0, 10)}` });
                                showNotification('Eksportowano CSV');
                            } else {
                                showNotification('Arkusz nie jest załadowany', true);
                            }
                        }
                    },
                    {
                        text: 'Save Locally',
                        icon: 'mdi mdi-content-save',
                        action: () => {
                            flushPendingSpreadsheetSave();
                            showNotification('Zapisano lokalnie');
                        }
                    },
                    {
                        text: 'Insert Row',
                        icon: 'mdi mdi-table-row-plus-after',
                        action: () => {
                            insertRow();
                        }
                    },
                    {
                        text: 'Insert Column',
                        icon: 'mdi mdi-table-column-plus-after',
                        action: () => {
                            insertColumn();
                        }
                    }
                ]
            },
            {
                group: 'Motyw arkusza (Handsontable)',
                icon: 'fas fa-palette',
                items: handsonTableThemeMenuItems
            },
            {
                group: 'System Tools',
                icon: 'fas fa-cog',
                items: [
                    {
                        text: 'Refresh Page',
                        icon: 'mdi mdi-refresh',
                        action: () => {
                            location.reload();
                            showNotification('Odświeżono stronę');
                        }
                    }
                ]
            },
            {
                group: 'Information',
                icon: 'fas fa-info-circle',
                items: [
                    {
                        text: 'About',
                        icon: 'mdi mdi-information',
                        action: () => {
                            showAboutModal();
                        }
                    }
                ]
            }
        ];

        // Generate menu
        options.forEach((group, groupIndex) => {
            // Group header with icon
            const groupHeader = document.createElement('div');
            groupHeader.classList.add('context-menu-group-header');

            // Add icon to group header
            const groupIcon = document.createElement('i');
            groupIcon.className = group.icon || 'fas fa-folder';
            groupIcon.style.marginRight = '10px';
            groupIcon.style.opacity = '0.7';

            const groupText = document.createElement('span');
            groupText.textContent = group.group;

            groupHeader.appendChild(groupIcon);
            groupHeader.appendChild(groupText);

            menu.appendChild(groupHeader);

            // Group items
            group.items.forEach(option => {
                const item = document.createElement('div');
                item.classList.add('context-menu-item');

                // Add active class for current theme
                if (option.isActive) {
                    item.classList.add('context-menu-item-active');
                }

                // Icon and text
                const icon = document.createElement('i');
                icon.className = option.icon;

                const text = document.createElement('span');
                text.textContent = option.text;

                item.appendChild(icon);
                item.appendChild(text);

                // Event handling
                item.addEventListener('click', () => {
                    option.action();
                    menu.remove();
                });

                menu.appendChild(item);
            });

            // Separator between groups
            if (groupIndex < options.length - 1) {
                const divider = document.createElement('div');
                divider.classList.add('context-menu-divider');
                menu.appendChild(divider);
            }
        });

        // Calculate dimensions and position
        const menuWidth = menu.offsetWidth;
        const menuHeight = menu.offsetHeight;
        const { left, top } = calculateContextMenuPosition(e, menuWidth, menuHeight);

        // Set position
        menu.style.left = `${left}px`;
        menu.style.top = `${top}px`;
    });

    // Close menu when clicking outside
    document.addEventListener('click', (e) => {
        const menu = document.getElementById('custom-context-menu');
        if (menu && !menu.contains(e.target)) {
            menu.remove();
        }
    });
}

// About modal function
function showAboutModal() {
    // Check if modal already exists
    if (document.getElementById('about-modal')) return;

    // Create modal with information
    const modal = document.createElement('div');
    modal.id = 'about-modal';

    modal.innerHTML = `
        <div class="about-modal-header">
            <div class="about-modal-title-section">
                <i class="mdi mdi-information"></i>
                <h2 class="about-modal-title">About Spreadsheet App</h2>
            </div>
            <button id="close-modal" class="about-modal-close">×</button>
        </div>
        
        <div class="about-modal-content">
            <p><strong>Version:</strong> 2.1.0</p>
            <p><strong>Last Updated:</strong> ${new Date().toLocaleDateString('pl-PL')}</p>
            
            <h3 class="about-modal-section-title">Key Features</h3>
            <ul>
                <li>Advanced Spreadsheet with Handsontable</li>
                <li>CSV Import/Export</li>
                <li>Dynamic Charts & Visualizations</li>
                <li>Custom Context Menu</li>
                <li>Multiple Theme Support</li>
                <li>Local Storage Integration</li>
                <li>Statistics & Data Analysis</li>
            </ul>
            
            <h3 class="about-modal-section-title">System Info</h3>
            <div class="about-system-info">
                <p><strong>Browser:</strong> ${navigator.userAgent.split(' ').slice(-1)[0]}</p>
                <p><strong>Screen Resolution:</strong> ${window.screen.width}x${window.screen.height}</p>
                <p><strong>Viewport:</strong> ${window.innerWidth}x${window.innerHeight}</p>
                <p><strong>Theme:</strong> ${document.documentElement.getAttribute('data-theme') || 'dark'}</p>
            </div>
        </div>
    `;

    // Overlay
    const overlay = document.createElement('div');
    overlay.id = 'about-modal-overlay';

    // Add to document
    document.body.appendChild(overlay);
    document.body.appendChild(modal);

    // Appearance animation
    requestAnimationFrame(() => {
        overlay.classList.add('show');
        modal.classList.add('show');
    });

    // Close function
    function closeModal() {
        modal.classList.remove('show');
        overlay.classList.remove('show');

        setTimeout(() => {
            if (document.body.contains(modal)) document.body.removeChild(modal);
            if (document.body.contains(overlay)) document.body.removeChild(overlay);
        }, 300);
    }

    // Close handlers
    const closeButton = modal.querySelector('#close-modal');
    closeButton.addEventListener('click', closeModal);

    // Close when clicking outside modal
    overlay.addEventListener('click', closeModal);

    // Close on Escape key
    document.addEventListener('keydown', function escHandler(e) {
        if (e.key === 'Escape') {
            closeModal();
            document.removeEventListener('keydown', escHandler);
        }
    });
}

// Add these to your existing script section
document.addEventListener('DOMContentLoaded', function () {
    // Remove draggable attribute from all elements
    document.querySelectorAll('[draggable="true"]').forEach(el => {
        el.removeAttribute('draggable');
    });

    // Prevent dragstart event
    document.addEventListener('dragstart', function (e) {
        e.preventDefault();
        return false;
    });

    // Prevent drop event
    document.addEventListener('drop', function (e) {
        e.preventDefault();
        return false;
    });

    // Prevent dragover event
    document.addEventListener('dragover', function (e) {
        e.preventDefault();
        return false;
    });
});