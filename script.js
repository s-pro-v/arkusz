// Register custom cell renderer for formatting globally
Handsontable.renderers.registerRenderer('customRenderer', function (instance, td, row, col, prop, value, cellProperties) {
    // Use the base text renderer first
    Handsontable.renderers.TextRenderer.apply(this, arguments);

    // Apply background color if defined
    if (cellProperties.backgroundColor) {
        td.style.backgroundColor = cellProperties.backgroundColor;
    }

    // Apply text color if defined
    if (cellProperties.color) {
        td.style.color = cellProperties.color;
    }
});

document.addEventListener('DOMContentLoaded', function () {
    // Register Polish language dictionary for Handsontable
    const plPL = {
        languageCode: 'pl-PL',
        dict: {
            'contextMenu': {
                'items': {
                    'row_above': 'Wstaw wiersz powyżej',
                    'row_below': 'Wstaw wiersz poniżej',
                    'col_left': 'Wstaw kolumnę z lewej',
                    'col_right': 'Wstaw kolumnę z prawej',
                    'remove_row': 'Usuń wiersz',
                    'remove_col': 'Usuń kolumnę',
                    'copy': 'Kopiuj',
                    'cut': 'Wytnij',
                    'paste': 'Wklej',
                    'alignment': 'Wyrównanie',
                    'merge_cells': 'Scal komórki',
                    'unmerge_cells': 'Rozdziel komórki',
                    'borders': 'Obramowanie',
                    'clear_custom': 'Wyczyść zawartość'
                }
            }
        }
    };

    Handsontable.languages.registerLanguageDictionary(plPL);

    let hot = null;
    let isDestroying = false; // Flag to track if table is in destroying state
    let currentTheme = localStorage.getItem('theme') || 'ht-theme-main';

    // Add console logging functionality
    function logToConsole(message, type = 'info') {
        const consoleContent = document.querySelector('.console-content');
        const entry = document.createElement('div');
        entry.className = `console-entry console-${type}`;

        // Add timestamp
        const timestamp = new Date().toLocaleTimeString();
        entry.textContent = `[${timestamp}] ${message}`;

        consoleContent.appendChild(entry);

        // Auto-scroll to bottom
        const consoleArea = document.getElementById('consoleOutput');
        consoleArea.scrollTop = consoleArea.scrollHeight;

        // Limit number of entries (keep last 50)
        const entries = consoleContent.getElementsByClassName('console-entry');
        if (entries.length > 50) {
            consoleContent.removeChild(entries[0]);
        }
    }

    // Make console draggable
    function makeDraggable(element) {
        let pos1 = 0, pos2 = 0, pos3 = 0, pos4 = 0;
        const header = element.querySelector('.console-header');

        if (header) {
            header.onmousedown = dragMouseDown;
        }

        function dragMouseDown(e) {
            e = e || window.event;
            e.preventDefault();
            // Get initial mouse cursor position
            pos3 = e.clientX;
            pos4 = e.clientY;
            document.onmouseup = closeDragElement;
            // Call function whenever cursor moves
            document.onmousemove = elementDrag;
        }

        function elementDrag(e) {
            e = e || window.event;
            e.preventDefault();
            // Calculate new cursor position
            pos1 = pos3 - e.clientX;
            pos2 = pos4 - e.clientY;
            pos3 = e.clientX;
            pos4 = e.clientY;
            // Set element's new position
            element.style.top = (element.offsetTop - pos2) + "px";
            element.style.left = (element.offsetLeft - pos1) + "px";
            element.style.right = "auto";
            element.style.bottom = "auto";
        }

        function closeDragElement() {
            // Stop moving when mouse button is released
            document.onmouseup = null;
            document.onmousemove = null;
        }
    }

    // Theme name mapping for display
    const themeDisplayNames = {
        'ht-theme-main': 'Main Light',
        'ht-theme-horizon': 'Horizon Light',
        'ht-theme-main-dark': 'Main Dark',
        'ht-theme-horizon-dark': 'Horizon Dark',
        'ht-no-theme': 'No theme'
    };

    function saveToLocalStorage() {
        if (!hot || isDestroying) return;
        try {
            const data = hot.getData();
            localStorage.setItem('spreadsheetData', JSON.stringify(data));

            // Save cell formatting
            const cellMeta = [];
            for (let row = 0; row < hot.countRows(); row++) {
                for (let col = 0; col < hot.countCols(); col++) {
                    const meta = hot.getCellMeta(row, col);
                    if (meta.backgroundColor || meta.color) {
                        cellMeta.push({
                            row,
                            col,
                            backgroundColor: meta.backgroundColor,
                            color: meta.color
                        });
                    }
                }
            }
            localStorage.setItem('spreadsheetCellMeta', JSON.stringify(cellMeta));

            logToConsole('Dane zapisane lokalnie', 'success');
        } catch (error) {
            console.error('Błąd podczas zapisywania:', error);
            logToConsole('Błąd podczas zapisywania: ' + error.message, 'error');
        }
    }

    function showNotification(message, isError = false) {
        // Function disabled - using console instead
        // Display console if it's hidden and this is an error
        if (isError) {
            const consoleArea = document.getElementById('consoleOutput');
            if (consoleArea.style.display === 'none' || consoleArea.style.display === '') {
                consoleArea.style.display = 'block';
            }
        }
        // Log to console instead
        logToConsole(message, isError ? 'error' : 'success');
    }

    function loadFromLocalStorage() {
        const savedData = localStorage.getItem('spreadsheetData');
        if (savedData) {
            try {
                const parsedData = JSON.parse(savedData);
                if (Array.isArray(parsedData) && parsedData.length > 0) {
                    return parsedData;
                }
            } catch (error) {
                console.error('Błąd podczas wczytywania danych:', error);
            }
        }
        return [['']];
    }

    function applyStoredCellFormatting() {
        if (!hot || isDestroying) return;

        try {
            const savedCellMeta = localStorage.getItem('spreadsheetCellMeta');
            if (!savedCellMeta) return;

            const parsedCellMeta = JSON.parse(savedCellMeta);
            if (!Array.isArray(parsedCellMeta)) return;

            parsedCellMeta.forEach(meta => {
                if (meta.row < hot.countRows() && meta.col < hot.countCols()) {
                    hot.setCellMeta(meta.row, meta.col, 'backgroundColor', meta.backgroundColor);
                    hot.setCellMeta(meta.row, meta.col, 'color', meta.color);
                    hot.setCellMeta(meta.row, meta.col, 'renderer', 'customRenderer');
                }
            });

            hot.render();
            logToConsole('Zastosowano zapisane formatowanie komórek', 'info');
        } catch (error) {
            console.error('Błąd podczas wczytywania formatowania komórek:', error);
            logToConsole('Błąd podczas wczytywania formatowania: ' + error.message, 'error');
        }
    }

    function changeTheme() {
        const selectedTheme = document.getElementById('themeSelect').value;
        currentTheme = selectedTheme;

        // Update theme display in status bar
        document.getElementById('currentTheme').textContent = 'Motyw: ' + themeDisplayNames[selectedTheme];
        logToConsole(`Zmieniono motyw na: ${themeDisplayNames[selectedTheme]}`, 'theme');

        // Save theme preference
        localStorage.setItem('theme', selectedTheme);

        // Set UI mode (light/dark) based on theme
        const isDarkTheme = selectedTheme.includes('dark');
        document.documentElement.setAttribute('data-theme', isDarkTheme ? 'dark' : 'light');

        // Apply theme to console elements
        const consoleArea = document.getElementById('consoleOutput');
        if (consoleArea) {
            // The CSS variables will be automatically applied through the data-theme attribute
            logToConsole('Zastosowano nowy motyw do konsoli', 'theme');
        }

        // Get data before destroying
        let data = [['']];
        if (hot && !isDestroying) {
            try {
                data = hot.getData();
            } catch (error) {
                console.error('Błąd podczas pobierania danych:', error);
                data = loadFromLocalStorage();
            }
        }

        // Set destroying flag to prevent hooks from firing
        isDestroying = true;

        // Destroy the old instance
        if (hot) {
            try {
                hot.destroy();
            } catch (error) {
                console.error('Błąd podczas niszczenia tabeli:', error);
            }
        }

        // Reset the isDestroying flag after a small delay
        setTimeout(() => {
            isDestroying = false;

            // Update the container class
            document.getElementById('spreadsheet').className = selectedTheme;

            // Reinitialize the table with the new theme
            initSpreadsheet(data);
        }, 50);
    }

    function showLoader() {
        document.getElementById('loadingOverlay').style.display = 'flex';
    }

    function hideLoader() {
        document.getElementById('loadingOverlay').style.display = 'none';
    }

    function updateStatusBar() {
        if (!hot || isDestroying) return;
        try {
            const totalRows = hot.countRows() - (hot.countEmptyRows ? hot.countEmptyRows() : 0);
            const selected = hot.getSelected() || [];
            const selectedCount = selected.reduce((acc, curr) => {
                if (!curr || curr.length < 3) return acc;
                const [row1, , row2] = curr;
                return acc + Math.abs(row2 - row1) + 1;
            }, 0);

            document.getElementById('rowCount').textContent = `Wiersze: ${totalRows}`;
            document.getElementById('selectedCount').textContent = `Zaznaczone: ${selectedCount}`;

            // Update current cell info if a selection exists
            if (selected && selected.length > 0) {
                const [row, col] = selected[0];
                if (typeof row === 'number' && typeof col === 'number') {
                    const colLetter = String.fromCharCode(65 + col);
                    document.getElementById('currentCell').textContent = `Komórka: ${colLetter}${row + 1}`;
                }
            } else {
                document.getElementById('currentCell').textContent = 'Komórka: -';
            }
        } catch (error) {
            console.error('Błąd podczas aktualizacji paska statusu:', error);
        }
    }

    // Removing CSS based highlighting function
    function highlightRowAndColumn(row, col) {
        // Function disabled
        if (typeof row === 'number' && typeof col === 'number') {
            // Only update current cell display in status bar
            const colLetter = String.fromCharCode(65 + col);
            document.getElementById('currentCell').textContent = `Komórka: ${colLetter}${row + 1}`;
        }
    }

    function clearHighlights() {
        // Function disabled
    }

    function initSpreadsheet(data = loadFromLocalStorage()) {
        // Clear previous instance fully
        const container = document.getElementById('spreadsheet');
        container.innerHTML = '';

        // Reset the flags
        isDestroying = false;

        hot = new Handsontable(container, {
            data: data,
            rowHeaders: true,
            colHeaders: true,
            height: '100%',
            width: '100%',
            licenseKey: 'non-commercial-and-evaluation',
            language: 'pl-PL',
            contextMenu: {
                items: {
                    'row_above': {
                        name: 'Wstaw wiersz powyżej'
                    },
                    'row_below': {
                        name: 'Wstaw wiersz poniżej'
                    },
                    'col_left': {
                        name: 'Wstaw kolumnę z lewej'
                    },
                    'col_right': {
                        name: 'Wstaw kolumnę z prawej'
                    },
                    'remove_row': {
                        name: 'Usuń wiersz'
                    },
                    'remove_col': {
                        name: 'Usuń kolumnę'
                    },
                    'separator': Handsontable.plugins.ContextMenu.SEPARATOR,
                    'copy': {
                        name: 'Kopiuj'
                    },
                    'cut': {
                        name: 'Wytnij'
                    },
                    'paste': {
                        name: 'Wklej'
                    },
                    'separator2': Handsontable.plugins.ContextMenu.SEPARATOR,
                    'alignment': {
                        name: 'Wyrównanie'
                    },
                    'merge_cells': {
                        name: 'Scal komórki'
                    },
                    'unmerge_cells': {
                        name: 'Rozdziel komórki'
                    },
                    'borders': {
                        name: 'Obramowanie'
                    },
                    'clear_custom': {
                        name: 'Wyczyść zawartość',
                        callback: function (key, selection) {
                            this.clear(selection);
                        }
                    },
                    'separator3': Handsontable.plugins.ContextMenu.SEPARATOR,
                    'remove_format': {
                        name: 'Usuń formatowanie',
                        callback: function (key, selection) {
                            removeFormatting();
                        }
                    }
                }
            },
            dropdownMenu: true,
            filters: true,
            columnSorting: true,
            manualColumnResize: true,
            manualRowResize: true,
            manualColumnMove: true,
            manualRowMove: true,
            mergeCells: true,
            comments: true,
            undo: true,
            redo: true,
            fillHandle: {
                autoInsertRow: true,
                direction: 'vertical'
            },
            allowInvalid: false,
            wordWrap: true,
            fixedRowsTop: 0,
            fixedColumnsLeft: 0,
            search: {
                searchResultClass: 'hot-search-result',
                queryMethod: 'contains'
            },
            customBorders: {
                borderWidth: 1,
                className: 'custom-border'
            },
            autoRowSize: {
                syncLimit: 500
            },
            autoColumnSize: {
                syncLimit: 500
            },
            selectionMode: 'multiple',
            observeChanges: true,
            beforeDestroy: function () {
                // Mark as destroying to prevent hook issues
                isDestroying = true;
            },
            afterChange: function (changes, source) {
                if (isDestroying) return;
                updateStatusBar();
                if (changes) {
                    saveToLocalStorage();
                }
            },
            afterSelection: function (row, col, row2, col2) {
                if (isDestroying) return;
                // Update only the cell position indicator
                try {
                    if (typeof row === 'number' && typeof col === 'number') {
                        const colLetter = String.fromCharCode(65 + col);
                        document.getElementById('currentCell').textContent = `Komórka: ${colLetter}${row + 1}`;
                    }
                } catch (error) {
                    console.error('Błąd podczas aktualizacji komórki:', error);
                }

                // Removed highlight function call
            },
            afterDeselect: function () {
                if (isDestroying) return;
                // Removed clearHighlights call
                document.getElementById('currentCell').textContent = 'Komórka: -';
            },
            afterRender: function () {
                if (isDestroying) return;
                // Removed re-apply highlighting code
            },
            outsideClickDeselects: false,
            enterMoves: {
                row: 1,
                col: 0
            },
            autoWrapRow: true,
            autoWrapCol: true,
            stretchH: 'all',
            trimRows: true,
            hiddenRows: {
                indicators: true
            },
            hiddenColumns: {
                indicators: true
            },
            renderAllRows: false,
            viewportRowRenderingOffset: 70,
            viewportColumnRenderingOffset: 30
        });

        // Update status bar safely with a small delay
        setTimeout(() => {
            if (!isDestroying) {
                updateStatusBar();

                // Apply stored cell formatting
                applyStoredCellFormatting();
            }
        }, 50);
    }

    // Apply theme from localStorage on page load
    if (currentTheme) {
        document.getElementById('spreadsheet').className = currentTheme;
        document.getElementById('themeSelect').value = currentTheme;

        // Set UI mode (light/dark) based on theme
        const isDarkTheme = currentTheme.includes('dark');
        document.documentElement.setAttribute('data-theme', isDarkTheme ? 'dark' : 'light');

        // Update theme display name
        document.getElementById('currentTheme').textContent = 'Motyw: ' + themeDisplayNames[currentTheme];
    }

    // Event Listeners
    document.getElementById('importBtn').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });

    document.getElementById('fileInput').addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        showLoader();
        logToConsole(`Importowanie pliku: ${file.name}`, 'import');

        Papa.parse(file, {
            complete: (results) => {
                if (results.data && results.data.length > 0) {
                    initSpreadsheet(results.data);
                    logToConsole(`Zaimportowano ${results.data.length} wierszy`, 'success');
                }
                hideLoader();
            },
            error: (error) => {
                console.error('Błąd podczas wczytywania pliku CSV:', error);
                alert('Błąd podczas wczytywania pliku CSV');
                logToConsole('Błąd podczas wczytywania pliku CSV: ' + error, 'error');
                hideLoader();
            }
        });
    });

    document.getElementById('exportBtn').addEventListener('click', () => {
        if (!hot || isDestroying) return;

        try {
            const exportPlugin = hot.getPlugin('exportFile');
            const filename = `eksport_${new Date().toISOString()}`;
            exportPlugin.downloadFile('csv', {
                filename: filename,
                columnHeaders: true,
                rowHeaders: true
            });
            logToConsole(`Wyeksportowano dane do pliku ${filename}.csv`, 'export');
        } catch (error) {
            console.error('Błąd podczas eksportu CSV:', error);
            logToConsole('Błąd podczas eksportu CSV: ' + error.message, 'error');
            // Auto-show console on error
            document.getElementById('consoleOutput').style.display = 'block';
        }
    });

    document.getElementById('saveBtn').addEventListener('click', saveToLocalStorage);

    // Theme change listener
    document.getElementById('themeSelect').addEventListener('change', changeTheme);

    // Initialize spreadsheet on load
    initSpreadsheet();

    // Initialize draggable console
    makeDraggable(document.getElementById('consoleOutput'));

    // Console toggle functionality
    document.getElementById('consoleBtn').addEventListener('click', () => {
        const consoleArea = document.getElementById('consoleOutput');
        if (consoleArea.style.display === 'none' || consoleArea.style.display === '') {
            consoleArea.style.display = 'flex';
            logToConsole('Konsola otwarta', 'success');
        } else {
            consoleArea.style.display = 'none';
        }
    });

    // Close console button
    document.getElementById('closeConsoleBtn').addEventListener('click', () => {
        document.getElementById('consoleOutput').style.display = 'none';
    });

    // Clear console button
    document.getElementById('clearConsoleBtn').addEventListener('click', () => {
        const consoleContent = document.querySelector('.console-content');
        consoleContent.innerHTML = '';
        const entry = document.createElement('div');
        entry.className = 'console-entry console-info';
        entry.textContent = '[' + new Date().toLocaleTimeString() + '] Konsola wyczyszczona';
        consoleContent.appendChild(entry);
    });

    // Log initial application start
    logToConsole('Aplikacja uruchomiona', 'success');
    logToConsole(`Bieżący motyw: ${themeDisplayNames[currentTheme]}`, 'theme');
    logToConsole('Arkusz gotowy do użycia', 'info');
    logToConsole('Skróty klawiszowe: Ctrl+Shift+F (formatowanie), Ctrl+Shift+R (usuń format)', 'info');

    // Format button functionality
    document.getElementById('formatBtn').addEventListener('click', () => {
        const formatPanel = document.getElementById('formatPanel');
        if (formatPanel.style.display === 'none' || formatPanel.style.display === '') {
            formatPanel.style.display = 'block';

            // Reset color previews
            document.getElementById('bgColorPreview').style.backgroundColor = '#ffffff';
            document.getElementById('textColorPreview').style.backgroundColor = '#000000';

            // Log opening of format panel
            logToConsole('Otworzono panel formatowania', 'info');
        } else {
            formatPanel.style.display = 'none';
        }
    });

    // Close format panel button
    document.getElementById('closeFormatPanelBtn').addEventListener('click', () => {
        document.getElementById('formatPanel').style.display = 'none';
    });

    // Initialize format panel variables
    let selectedBgColor = '#ffffff';
    let selectedTextColor = '#000000';

    // Background color picker functionality
    const bgColorOptions = document.querySelectorAll('#bgColorPicker .color-option');
    bgColorOptions.forEach(option => {
        option.addEventListener('click', () => {
            selectedBgColor = option.getAttribute('data-color');
            document.getElementById('bgColorPreview').style.backgroundColor = selectedBgColor;

            // Update active state
            bgColorOptions.forEach(o => o.classList.remove('active'));
            option.classList.add('active');

            logToConsole(`Wybrano kolor tła: ${selectedBgColor}`, 'info');
        });
    });

    // Text color picker functionality
    const textColorOptions = document.querySelectorAll('#textColorPicker .color-option');
    textColorOptions.forEach(option => {
        option.addEventListener('click', () => {
            selectedTextColor = option.getAttribute('data-color');
            document.getElementById('textColorPreview').style.backgroundColor = selectedTextColor;

            // Update active state
            textColorOptions.forEach(o => o.classList.remove('active'));
            option.classList.add('active');

            logToConsole(`Wybrano kolor tekstu: ${selectedTextColor}`, 'info');
        });
    });

    // Reset formatting button
    document.getElementById('resetFormatBtn').addEventListener('click', () => {
        // Reset selected colors
        selectedBgColor = '#ffffff';
        selectedTextColor = '#000000';

        // Reset previews
        document.getElementById('bgColorPreview').style.backgroundColor = selectedBgColor;
        document.getElementById('textColorPreview').style.backgroundColor = selectedTextColor;

        // Reset active states
        bgColorOptions.forEach(o => o.classList.remove('active'));
        textColorOptions.forEach(o => o.classList.remove('active'));

        // Find and make the default colors active
        document.querySelector('#bgColorPicker [data-color="#ffffff"]').classList.add('active');
        document.querySelector('#textColorPicker [data-color="#000000"]').classList.add('active');

        logToConsole('Zresetowano wybór kolorów', 'warning');
    });

    // Function to remove all formatting from selected cells
    function removeFormatting() {
        if (!hot) return;

        const selected = hot.getSelected();
        if (!selected || selected.length === 0) {
            logToConsole('Nie wybrano żadnej komórki do usunięcia formatowania', 'error');
            return;
        }

        try {
            let cellCount = 0;

            // For each selection range
            selected.forEach(range => {
                const [startRow, startCol, endRow, endCol] = range;

                // Remove formatting from each cell in the range
                for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row++) {
                    for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col++) {
                        // Get cell meta
                        hot.removeCellMeta(row, col, 'backgroundColor');
                        hot.removeCellMeta(row, col, 'color');
                        hot.removeCellMeta(row, col, 'renderer');

                        cellCount++;
                    }
                }
            });

            // Render the changes
            hot.render();

            logToConsole(`Usunięto formatowanie z ${cellCount} komórek`, 'warning');
        } catch (error) {
            console.error('Błąd podczas usuwania formatowania:', error);
            logToConsole(`Błąd podczas usuwania formatowania: ${error.message}`, 'error');
        }
    }

    // Remove formatting button 
    document.getElementById('removeFormatBtn').addEventListener('click', () => {
        removeFormatting();
        document.getElementById('formatPanel').style.display = 'none';
    });

    // Apply formatting button
    document.getElementById('applyFormatBtn').addEventListener('click', () => {
        if (!hot) return;

        const selected = hot.getSelected();
        if (!selected || selected.length === 0) {
            logToConsole('Nie wybrano żadnej komórki', 'error');
            return;
        }

        try {
            // For each selection range
            selected.forEach(range => {
                const [startRow, startCol, endRow, endCol] = range;

                // Apply formatting to each cell in the range
                for (let row = Math.min(startRow, endRow); row <= Math.max(startRow, endRow); row++) {
                    for (let col = Math.min(startCol, endCol); col <= Math.max(startCol, endCol); col++) {
                        // Get cell meta
                        const cellMeta = hot.getCellMeta(row, col);

                        // Set the renderer to our custom renderer
                        cellMeta.renderer = 'customRenderer';

                        // Set the background and text colors
                        cellMeta.backgroundColor = selectedBgColor;
                        cellMeta.color = selectedTextColor;

                        // Update cell meta
                        hot.setCellMeta(row, col, 'backgroundColor', selectedBgColor);
                        hot.setCellMeta(row, col, 'color', selectedTextColor);
                        hot.setCellMeta(row, col, 'renderer', 'customRenderer');
                    }
                }
            });

            // Render the changes
            hot.render();

            // Close the format panel
            document.getElementById('formatPanel').style.display = 'none';

            // Log success
            const cellCount = selected.reduce((count, range) => {
                const [startRow, startCol, endRow, endCol] = range;
                return count + ((Math.abs(endRow - startRow) + 1) * (Math.abs(endCol - startCol) + 1));
            }, 0);

            logToConsole(`Sformatowano ${cellCount} komórek`, 'success');
        } catch (error) {
            console.error('Błąd podczas formatowania komórek:', error);
            logToConsole(`Błąd podczas formatowania: ${error.message}`, 'error');
        }
    });

    // Remove formatting button in toolbar
    document.getElementById('clearFormatBtn').addEventListener('click', () => {
        removeFormatting();
    });

    // Add keyboard shortcuts for formatting 
    document.addEventListener('keydown', (event) => {
        // Only process if hot exists and is initialized
        if (!hot) return;

        // Ctrl+Shift+F - Open format panel
        if (event.ctrlKey && event.shiftKey && event.key === 'F') {
            event.preventDefault();
            const formatPanel = document.getElementById('formatPanel');
            formatPanel.style.display = formatPanel.style.display === 'none' || formatPanel.style.display === '' ? 'block' : 'none';
            if (formatPanel.style.display === 'block') {
                document.getElementById('bgColorPreview').style.backgroundColor = '#ffffff';
                document.getElementById('textColorPreview').style.backgroundColor = '#000000';
            }
        }

        // Ctrl+Shift+R - Remove formatting from selected cells
        if (event.ctrlKey && event.shiftKey && event.key === 'R') {
            event.preventDefault();
            removeFormatting();
        }
    });

    // Info button functionality
    document.getElementById('infoBtn').addEventListener('click', () => {
        const infoPopup = document.getElementById('infoPopup');
        if (infoPopup.style.display === 'none' || infoPopup.style.display === '') {
            infoPopup.style.display = 'block';
        } else {
            infoPopup.style.display = 'none';
        }
    });

    // Close info popup button
    document.getElementById('closeInfoPopupBtn').addEventListener('click', () => {
        document.getElementById('infoPopup').style.display = 'none';
    });

    // Close info popup when clicking outside
    document.addEventListener('mousedown', (event) => {
        const infoPopup = document.getElementById('infoPopup');
        const infoBtn = document.getElementById('infoBtn');

        if (infoPopup.style.display === 'block') {
            // Check if the click is outside the popup and not on the info button
            if (!infoPopup.contains(event.target) && !infoBtn.contains(event.target)) {
                infoPopup.style.display = 'none';
            }
        }
    });

    // Escape key closes the popups
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') {
            // Close info popup if open
            const infoPopup = document.getElementById('infoPopup');
            if (infoPopup.style.display === 'block') {
                infoPopup.style.display = 'none';
            }

            // Close format panel if open
            const formatPanel = document.getElementById('formatPanel');
            if (formatPanel.style.display === 'block') {
                formatPanel.style.display = 'none';
            }

            // Close console if open
            const consoleOutput = document.getElementById('consoleOutput');
            if (consoleOutput.style.display === 'flex') {
                consoleOutput.style.display = 'none';
            }
        }
    });
});