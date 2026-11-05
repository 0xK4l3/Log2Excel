// Log2Excel - Main JavaScript Application Logic (Individual & Massive Conversion)

document.addEventListener('DOMContentLoaded', () => {
    // DOM Elements
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const btnSelectFiles = document.getElementById('btn-select-files');
    const btnSelectFolder = document.getElementById('btn-select-folder');

    const delimiterSelect = document.getElementById('delimiter-select');
    const customDelimiterInput = document.getElementById('custom-delimiter');
    const hasHeadersCheckbox = document.getElementById('has-headers');
    const directBatchModeCheckbox = document.getElementById('direct-batch-mode');

    // Batch View Elements
    const batchCard = document.getElementById('batch-card');
    const batchTitle = document.getElementById('batch-title');
    const batchSubtitle = document.getElementById('batch-subtitle');
    const batchProgressText = document.getElementById('batch-progress-text');
    const batchProgressPercent = document.getElementById('batch-progress-percent');
    const batchProgressBar = document.getElementById('batch-progress-bar');
    const batchFileList = document.getElementById('batch-file-list');
    const btnOpenFolder = document.getElementById('btn-open-folder');
    const btnReconvertBatch = document.getElementById('btn-reconvert-batch');
    const btnClearBatch = document.getElementById('btn-clear-batch');

    // Single Preview Elements
    const previewCard = document.getElementById('preview-card');
    const fileNameEl = document.getElementById('file-name');
    const fileDetailsEl = document.getElementById('file-details');
    const previewTableHeaders = document.getElementById('table-headers');
    const previewTableBody = document.getElementById('table-body');
    const btnClear = document.getElementById('btn-clear');
    const btnSaveSameFolder = document.getElementById('btn-save-same-folder');
    const btnExport = document.getElementById('btn-export');

    // Application State (Single File Preview)
    let loadedFileContent = null;
    let loadedFileName = '';
    let loadedFileSize = 0;
    let loadedFilePath = '';
    let parsedHeaders = [];
    let parsedRows = [];

    // Application State (Batch / Massive Conversion)
    let currentBatchItems = []; // [{ name, path, size, browserFile }]
    let lastOutputFolder = '';
    let isBatchProcessing = false;

    // Initialize Lucide Icons
    lucide.createIcons();

    function hasPythonApi() {
        return Boolean(window.pywebview && window.pywebview.api);
    }

    // Event Listeners for Configuration changes
    delimiterSelect.addEventListener('change', (e) => {
        if (e.target.value === 'custom') {
            customDelimiterInput.classList.remove('hidden');
            customDelimiterInput.focus();
        } else {
            customDelimiterInput.classList.add('hidden');
            customDelimiterInput.value = '';
            onSettingsChanged();
        }
    });

    customDelimiterInput.addEventListener('input', () => {
        onSettingsChanged();
    });

    hasHeadersCheckbox.addEventListener('change', () => {
        onSettingsChanged();
    });

    function onSettingsChanged() {
        if (!previewCard.classList.contains('hidden') && loadedFileContent) {
            processAndPreview();
        } else if (!batchCard.classList.contains('hidden') && currentBatchItems.length > 0 && !isBatchProcessing) {
            runBatchConversion(currentBatchItems);
        }
    }

    // Drag and Drop Events
    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
        }, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
        }, false);
    });

    dropZone.addEventListener('drop', async (e) => {
        const dt = e.dataTransfer;
        const files = dt ? dt.files : null;
        if (!files || files.length === 0) return;

        // If inside WebView2 + pywebview, capture full filesystem paths from dropped files
        if (hasPythonApi() && window.chrome && window.chrome.webview && window.chrome.webview.postMessageWithAdditionalObjects) {
            try {
                await window.pywebview.api.clear_dropped_paths();
                window.chrome.webview.postMessageWithAdditionalObjects('FilesDropped', files);
                const droppedItems = await window.pywebview.api.get_dropped_paths();
                if (droppedItems && droppedItems.length > 0) {
                    handleIncomingItems(droppedItems);
                    return;
                }
            } catch (err) {
                console.warn('No se pudieron obtener rutas nativas del drag & drop, usando lector estándar:', err);
            }
        }

        // Fallback to browser File objects
        const items = Array.from(files).map(f => ({
            name: f.name,
            path: f.path || '',
            size: f.size,
            browserFile: f
        }));
        handleIncomingItems(items);
    });

    // Buttons & Click to Select Files / Folder
    btnSelectFiles.addEventListener('click', async (e) => {
        e.stopPropagation();
        await triggerFileSelection();
    });

    btnSelectFolder.addEventListener('click', async (e) => {
        e.stopPropagation();
        if (hasPythonApi()) {
            const items = await window.pywebview.api.select_folder();
            if (items && items.length > 0) {
                handleIncomingItems(items);
            } else if (items && items.length === 0) {
                // User cancelled or folder had no .log/.txt/.csv files
            }
        } else {
            fileInput.click();
        }
    });

    dropZone.addEventListener('click', async () => {
        await triggerFileSelection();
    });

    async function triggerFileSelection() {
        if (hasPythonApi()) {
            const items = await window.pywebview.api.select_files();
            if (items && items.length > 0) {
                handleIncomingItems(items);
            }
        } else {
            fileInput.click();
        }
    }

    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            const items = Array.from(e.target.files).map(f => ({
                name: f.name,
                path: f.path || '',
                size: f.size,
                browserFile: f
            }));
            handleIncomingItems(items);
        }
    });

    // Clear, Reconvert & Export Events
    btnClear.addEventListener('click', clearState);
    btnClearBatch.addEventListener('click', clearState);
    btnReconvertBatch.addEventListener('click', () => {
        if (currentBatchItems.length > 0 && !isBatchProcessing) {
            runBatchConversion(currentBatchItems);
        }
    });
    btnOpenFolder.addEventListener('click', () => {
        if (hasPythonApi() && lastOutputFolder) {
            window.pywebview.api.open_folder(lastOutputFolder);
        }
    });
    btnExport.addEventListener('click', () => exportSingleToExcel(false));
    btnSaveSameFolder.addEventListener('click', () => exportSingleToExcel(true));

    // Route incoming files to Batch Mode (no preview table) or Single Preview Mode
    function handleIncomingItems(items) {
        if (!items || items.length === 0) return;

        const useBatchMode = items.length > 1 || directBatchModeCheckbox.checked;
        if (useBatchMode) {
            currentBatchItems = items;
            runBatchConversion(items);
        } else {
            handleSingleFile(items[0]);
        }
    }

    // Read file text either from Python path or browser File object
    async function readItemText(item) {
        if (item.path && hasPythonApi()) {
            const res = await window.pywebview.api.read_file_content(item.path);
            if (res && res.success) {
                return res.content;
            }
            throw new Error((res && res.error) || 'Error al leer el archivo desde disco');
        }
        if (item.browserFile) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => resolve(e.target.result);
                reader.onerror = () => reject(new Error('Error al leer el archivo'));
                reader.readAsText(item.browserFile);
            });
        }
        throw new Error('No se pudo acceder al contenido del archivo');
    }

    // Parse log text into headers and rows (single-pass)
    function parseLogContent(content, delimiter, hasHeaders) {
        const rawLines = content.split(/\r?\n/);
        const rows = [];
        let headers = [];
        let maxCols = 0;
        let lineCount = 0;
        let validLinesCount = 0;

        for (let i = 0; i < rawLines.length; i++) {
            const trimmedLine = rawLines[i].trim();
            if (trimmedLine === '') continue;

            validLinesCount++;
            const row = trimmedLine.split(delimiter);
            if (row.length > maxCols) {
                maxCols = row.length;
            }

            if (hasHeaders && lineCount === 0) {
                headers = row;
            } else {
                rows.push(row);
            }
            lineCount++;
        }

        if (validLinesCount === 0) {
            return { headers: [], rows: [], validLinesCount: 0 };
        }

        if (!hasHeaders) {
            for (let i = 1; i <= maxCols; i++) {
                headers.push(`Columna ${i}`);
            }
        } else {
            while (headers.length < maxCols) {
                headers.push(`Columna ${headers.length + 1}`);
            }
        }

        return { headers, rows, validLinesCount };
    }

    // Helper: Build output Excel filename
    function getExportFilename(originalName) {
        const baseName = originalName.substring(0, originalName.lastIndexOf('.')) || originalName;
        return `${baseName}_convertido.xlsx`;
    }

    // Massive / Batch Conversion Logic (No spreadsheet preview, saves in same folder)
    async function runBatchConversion(items) {
        if (isBatchProcessing) return;
        isBatchProcessing = true;

        dropZone.classList.add('hidden');
        previewCard.classList.add('hidden');
        batchCard.classList.remove('hidden');

        btnReconvertBatch.disabled = true;
        btnClearBatch.disabled = true;
        btnOpenFolder.classList.add('hidden');

        const total = items.length;
        batchTitle.textContent = `Convirtiendo ${total} ${total === 1 ? 'archivo' : 'archivos'} a Excel...`;
        batchSubtitle.textContent = 'Los archivos convertidos se guardarán automáticamente en la misma carpeta.';
        updateBatchProgress(0, total);

        // Render initial batch list
        batchFileList.innerHTML = '';
        items.forEach((item, idx) => {
            const exportName = getExportFilename(item.name);
            const rowEl = document.createElement('div');
            rowEl.className = 'batch-item';
            rowEl.id = `batch-item-${idx}`;
            rowEl.innerHTML = `
                <div class="batch-item-info">
                    <div class="batch-item-name" title="${escapeHtml(item.path || item.name)}">${escapeHtml(item.name)} &rarr; ${escapeHtml(exportName)}</div>
                    <div class="batch-item-meta" id="batch-meta-${idx}">${formatBytes(item.size)}${item.path ? ' • ' + escapeHtml(getFolderPath(item.path)) : ''}</div>
                </div>
                <span class="batch-badge pending" id="batch-badge-${idx}">Pendiente</span>
            `;
            batchFileList.appendChild(rowEl);
        });

        const delimiter = getDelimiter();
        const hasHeaders = hasHeadersCheckbox.checked;

        // If running in pywebview but files lacked paths, ask once for destination folder
        let fallbackFolder = '';
        const needsFallbackFolder = hasPythonApi() && items.some(it => !it.path);
        if (needsFallbackFolder) {
            fallbackFolder = await window.pywebview.api.select_output_folder();
            if (!fallbackFolder) {
                batchTitle.textContent = 'Conversión cancelada';
                batchSubtitle.textContent = 'No se seleccionó una carpeta de destino.';
                btnReconvertBatch.disabled = false;
                btnClearBatch.disabled = false;
                isBatchProcessing = false;
                return;
            }
        }

        let completedCount = 0;
        let successCount = 0;
        lastOutputFolder = fallbackFolder || '';

        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const badgeEl = document.getElementById(`batch-badge-${i}`);
            const metaEl = document.getElementById(`batch-meta-${i}`);

            badgeEl.className = 'batch-badge processing';
            badgeEl.textContent = 'Convirtiendo...';

            // Yield briefly so UI updates smoothly
            await new Promise(r => setTimeout(r, 15));

            try {
                const content = await readItemText(item);
                const parsed = parseLogContent(content, delimiter, hasHeaders);

                if (parsed.validLinesCount === 0) {
                    throw new Error('Archivo vacío');
                }

                const sheetData = [parsed.headers, ...parsed.rows];
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(wb, ws, 'Datos Log');

                const exportName = getExportFilename(item.name);

                if (hasPythonApi()) {
                    const base64Content = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
                    let saveRes;
                    if (item.path) {
                        saveRes = await window.pywebview.api.save_excel_in_same_folder(base64Content, item.path, exportName);
                    } else {
                        saveRes = await window.pywebview.api.save_excel_to_folder(base64Content, fallbackFolder, exportName);
                    }

                    if (!saveRes || !saveRes.success) {
                        throw new Error((saveRes && saveRes.error) || 'No se pudo guardar el archivo');
                    }

                    lastOutputFolder = saveRes.folder || lastOutputFolder;
                    metaEl.textContent = `${formatBytes(item.size)} • ${parsed.validLinesCount.toLocaleString()} filas • Guardado en: ${saveRes.output_path}`;
                } else {
                    XLSX.writeFile(wb, exportName);
                    metaEl.textContent = `${formatBytes(item.size)} • ${parsed.validLinesCount.toLocaleString()} filas • Descargado`;
                }

                badgeEl.className = 'batch-badge done';
                badgeEl.textContent = '✓ Convertido';
                successCount++;
            } catch (err) {
                console.error(`Error convirtiendo ${item.name}:`, err);
                badgeEl.className = 'batch-badge error';
                badgeEl.textContent = 'Error';
                metaEl.textContent = `${formatBytes(item.size)} • ${err.message || 'Error al procesar'}`;
            }

            completedCount++;
            updateBatchProgress(completedCount, total);
        }

        batchTitle.textContent = `Conversión completada (${successCount} de ${total} convertidos)`;
        if (lastOutputFolder) {
            batchSubtitle.textContent = `Archivos Excel guardados en la misma carpeta: ${lastOutputFolder}`;
            btnOpenFolder.classList.remove('hidden');
        } else {
            batchSubtitle.textContent = 'Todos los archivos fueron procesados.';
        }

        btnReconvertBatch.disabled = false;
        btnClearBatch.disabled = false;
        isBatchProcessing = false;
        lucide.createIcons();
    }

    function updateBatchProgress(completed, total) {
        const pct = total > 0 ? Math.round((completed / total) * 100) : 0;
        batchProgressText.textContent = `${completed} de ${total} convertidos`;
        batchProgressPercent.textContent = `${pct}%`;
        batchProgressBar.style.width = `${pct}%`;
    }

    // Single File Handler (with preview table)
    async function handleSingleFile(item) {
        if (!item) return;

        loadedFileName = item.name;
        loadedFileSize = item.size;
        loadedFilePath = item.path || '';

        try {
            loadedFileContent = await readItemText(item);
            processAndPreview();

            dropZone.classList.add('hidden');
            batchCard.classList.add('hidden');
            previewCard.classList.remove('hidden');

            if (loadedFilePath && hasPythonApi()) {
                btnSaveSameFolder.classList.remove('hidden');
            } else {
                btnSaveSameFolder.classList.add('hidden');
            }
        } catch (err) {
            alert('Error al leer el archivo. Inténtalo de nuevo.');
        }
    }

    // Get current configured delimiter
    function getDelimiter() {
        const selectVal = delimiterSelect.value;
        if (selectVal === 'custom') {
            return customDelimiterInput.value || ';';
        }
        if (selectVal === '\\t') {
            return '\t';
        }
        return selectVal;
    }

    // Parse logic & UI preview builder for Single File
    function processAndPreview() {
        if (!loadedFileContent) return;

        const delimiter = getDelimiter();
        const hasHeaders = hasHeadersCheckbox.checked;

        const parsed = parseLogContent(loadedFileContent, delimiter, hasHeaders);
        if (parsed.validLinesCount === 0) {
            alert('El archivo seleccionado está vacío.');
            clearState();
            return;
        }

        parsedHeaders = parsed.headers;
        parsedRows = parsed.rows;

        fileNameEl.textContent = loadedFileName;
        const totalRowsStr = parsed.validLinesCount.toLocaleString();
        const sizeStr = formatBytes(loadedFileSize);
        fileDetailsEl.textContent = `${sizeStr} • ${totalRowsStr} filas detectadas`;

        renderTablePreview();
    }

    // Render preview table (first 15 rows)
    function renderTablePreview() {
        previewTableHeaders.innerHTML = '';
        previewTableBody.innerHTML = '';

        parsedHeaders.forEach(headerText => {
            const th = document.createElement('th');
            th.textContent = headerText;
            previewTableHeaders.appendChild(th);
        });

        const previewRows = parsedRows.slice(0, 15);
        previewRows.forEach(row => {
            const tr = document.createElement('tr');
            for (let i = 0; i < parsedHeaders.length; i++) {
                const td = document.createElement('td');
                const cellValue = row[i] !== undefined ? row[i].trim() : '';

                if (cellValue === '') {
                    td.textContent = '-';
                    td.classList.add('empty-cell');
                } else {
                    td.textContent = cellValue;
                }
                tr.appendChild(td);
            }
            previewTableBody.appendChild(tr);
        });

        lucide.createIcons();
    }

    // Export Single File to Excel
    function exportSingleToExcel(saveInSameFolder = false) {
        if (parsedRows.length === 0 && parsedHeaders.length === 0) {
            alert('No hay datos disponibles para exportar.');
            return;
        }

        const targetBtn = saveInSameFolder ? btnSaveSameFolder : btnExport;
        targetBtn.disabled = true;
        const originalText = targetBtn.innerHTML;
        targetBtn.innerHTML = `<i class="animate-spin" data-lucide="loader"></i> Procesando...`;
        lucide.createIcons();

        setTimeout(async () => {
            try {
                const sheetData = [parsedHeaders, ...parsedRows];
                const wb = XLSX.utils.book_new();
                const ws = XLSX.utils.aoa_to_sheet(sheetData);
                XLSX.utils.book_append_sheet(wb, ws, 'Datos Log');

                const exportName = getExportFilename(loadedFileName);

                if (hasPythonApi()) {
                    const base64Content = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' });
                    if (saveInSameFolder && loadedFilePath) {
                        const res = await window.pywebview.api.save_excel_in_same_folder(base64Content, loadedFilePath, exportName);
                        if (res && res.success) {
                            fileDetailsEl.textContent = `✓ Guardado en: ${res.output_path}`;
                        } else {
                            alert('Error al guardar en la misma carpeta.');
                        }
                    } else {
                        await window.pywebview.api.save_excel(base64Content, exportName);
                    }
                } else {
                    XLSX.writeFile(wb, exportName);
                }
            } catch (error) {
                console.error('Error generating Excel file:', error);
                alert('Ocurrió un error al generar el archivo Excel.');
            } finally {
                targetBtn.disabled = false;
                targetBtn.innerHTML = originalText;
                lucide.createIcons();
            }
        }, 80);
    }

    // Reset application state
    function clearState() {
        loadedFileContent = null;
        loadedFileName = '';
        loadedFileSize = 0;
        loadedFilePath = '';
        parsedHeaders = [];
        parsedRows = [];
        currentBatchItems = [];
        fileInput.value = '';

        previewCard.classList.add('hidden');
        batchCard.classList.add('hidden');
        dropZone.classList.remove('hidden');
    }

    function getFolderPath(fullPath) {
        if (!fullPath) return '';
        const lastSep = Math.max(fullPath.lastIndexOf('\\'), fullPath.lastIndexOf('/'));
        return lastSep > 0 ? fullPath.substring(0, lastSep) : fullPath;
    }

    function escapeHtml(str) {
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    // Helper: Format bytes to human readable sizes
    function formatBytes(bytes, decimals = 2) {
        if (!bytes || bytes === 0) return '0 Bytes';
        const k = 1024;
        const dm = decimals < 0 ? 0 : decimals;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
    }
});
