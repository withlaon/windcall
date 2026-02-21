document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const columnSelector = document.getElementById('column-selector');
    const columnList = document.getElementById('column-list');
    const exportSection = document.getElementById('export-section');
    const exportBtn = document.getElementById('export-btn');
    const previewTable = document.getElementById('preview-table');
    const dataPreview = document.getElementById('data-preview');
    const emptyState = document.getElementById('empty-state');
    const rowCountSpan = document.getElementById('row-count');
    const selectAllBtn = document.getElementById('select-all');
    const deselectAllBtn = document.getElementById('deselect-all');
    const summarySection = document.getElementById('summary-section');

    let workbookData = null;
    let headers = [];
    let fullData = [];
    let selectedColumns = new Set();

    // Drag and Drop Handlers
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles(files);
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    async function handleFiles(files) {
        if (files.length === 0) return;
        const file = files[0];
        const reader = new FileReader();

        reader.onload = async (e) => {
            const data = new Uint8Array(e.target.result);
            let workbookBuffer = data;
            
            try {
                // Strategy 1: Try reading without password first
                XLSX.read(data, { type: 'array' });
                console.log('File is not encrypted or successfully read without password.');
            } catch (noPassError) {
                console.log('File may be encrypted. Attempting decryption with XlsxPopulate...');
                try {
                    // Strategy 2: Use XlsxPopulate to decrypt (Supports Agile Encryption)
                    const workbook = await XlsxPopulate.fromDataAsync(data, { password: '3674601220' });
                    // Convert decrypted workbook back to buffer for SheetJS
                    workbookBuffer = await workbook.outputAsync();
                    console.log('Decryption successful using XlsxPopulate.');
                } catch (passError) {
                    console.error('Decryption failed with XlsxPopulate:', passError);
                    alert('파일을 읽는 중 오류가 발생했습니다.\n1. 암호가 맞지 않음 (3674601220)\n2. 지원하지 않는 파일 형식\n3. 파일 손상');
                    return;
                }
            }

            try {
                // Now read the (potentially decrypted) buffer with SheetJS
                const workbook = XLSX.read(workbookBuffer, { type: 'array' });
                
                // --- CUSTOM EXTRACTION LOGIC ---
                const targetSheetName = "갑지_협력사 전체 정산 확인용";
                let worksheet = workbook.Sheets[targetSheetName];
                
                if (worksheet) {
                    extractSummaryData(worksheet);
                } else {
                    console.warn(`Sheet "${targetSheetName}" not found. Falling back to first sheet.`);
                    worksheet = workbook.Sheets[workbook.SheetNames[0]];
                    summarySection.classList.add('hidden');
                }
                // -------------------------------

                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (json.length > 0) {
                    processData(json);
                } else {
                    alert('파일에 데이터가 없습니다.');
                }
            } catch (processError) {
                console.error('Data processing error:', processError);
                alert('데이터 처리 중 오류가 발생했습니다.');
            }
        };

        reader.readAsArrayBuffer(file);
    }

    /**
     * Extracts specific cells F24, F25, G24, G25, P24, P25
     * @param {Object} worksheet SheetJS Worksheet
     */
    function extractSummaryData(worksheet) {
        const getVal = (cell) => {
            const data = worksheet[cell];
            if (!data) return 0;
            return typeof data.v === 'number' ? data.v : 0;
        };

        const f24 = getVal('F24');
        const f25 = getVal('F25');
        const g24 = getVal('G24');
        const g25 = getVal('G25');
        const p24 = getVal('P24');
        const p25 = getVal('P25');

        const format = (num) => num.toLocaleString('ko-KR');

        document.getElementById('val-f24').textContent = format(f24);
        document.getElementById('val-f25').textContent = format(f25);
        document.getElementById('total-f').textContent = format(f24 + f25);

        document.getElementById('val-g24').textContent = format(g24);
        document.getElementById('val-g25').textContent = format(g25);
        document.getElementById('total-g').textContent = format(g24 + g25);

        document.getElementById('val-p24').textContent = format(p24);
        document.getElementById('val-p25').textContent = format(p25);
        document.getElementById('total-p').textContent = format(p24 + p25);

        summarySection.classList.remove('hidden');
    }

    function processData(data) {
        headers = data[0];
        fullData = data.slice(1);
        selectedColumns = new Set(headers); // Initially select all

        // Show UI elements
        emptyState.classList.add('hidden');
        dataPreview.classList.remove('hidden');
        columnSelector.classList.remove('hidden');
        exportSection.classList.remove('hidden');

        renderColumnList();
        renderPreview();
    }

    function renderColumnList() {
        columnList.innerHTML = '';
        if (!headers) return;

        headers.forEach(header => {
            const div = document.createElement('div');
            div.className = 'column-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = selectedColumns.has(header);
            checkbox.id = `col-${header}`;
            
            const label = document.createElement('span');
            label.textContent = header || '(Empty Header)';
            
            div.appendChild(checkbox);
            div.appendChild(label);
            
            div.addEventListener('click', (e) => {
                if (e.target !== checkbox) {
                    checkbox.checked = !checkbox.checked;
                }
                toggleColumn(header, checkbox.checked);
            });

            columnList.appendChild(div);
        });
    }

    function toggleColumn(header, isChecked) {
        if (isChecked) {
            selectedColumns.add(header);
        } else {
            selectedColumns.delete(header);
        }
        renderPreview();
    }

    function renderPreview() {
        const selectedHeaders = headers.filter(h => selectedColumns.has(h));
        rowCountSpan.textContent = `(${fullData.length} 행)`;

        // Clear table
        previewTable.innerHTML = '';

        // Header row
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        selectedHeaders.forEach(h => {
            const th = document.createElement('th');
            th.textContent = h || '';
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        previewTable.appendChild(thead);

        // Body rows (Preview limited to 20 for performance)
        const tbody = document.createElement('tbody');
        const previewLimit = Math.min(fullData.length, 20);
        
        for (let i = 0; i < previewLimit; i++) {
            const rowData = fullData[i];
            const tr = document.createElement('tr');
            
            headers.forEach((h, index) => {
                if (selectedColumns.has(h)) {
                    const td = document.createElement('td');
                    td.textContent = rowData[index] !== undefined ? rowData[index] : '';
                    tr.appendChild(td);
                }
            });
            tbody.appendChild(tr);
        }
        previewTable.appendChild(tbody);
    }

    // Select/Deselect All
    selectAllBtn.addEventListener('click', () => {
        selectedColumns = new Set(headers);
        renderColumnList();
        renderPreview();
    });

    deselectAllBtn.addEventListener('click', () => {
        selectedColumns.clear();
        renderColumnList();
        renderPreview();
    });

    // Export Logic
    exportBtn.addEventListener('click', () => {
        if (selectedColumns.size === 0) {
            alert('최소 하나 이상의 열을 선택해야 합니다.');
            return;
        }

        const selectedIndices = headers.reduce((acc, h, i) => {
            if (selectedColumns.has(h)) acc.push(i);
            return acc;
        }, []);

        const exportData = [
            headers.filter(h => selectedColumns.has(h)), // Headers
            ...fullData.map(row => selectedIndices.map(i => row[i])) // Filtered rows
        ];

        const ws = XLSX.utils.aoa_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Extracted Data");
        
        XLSX.writeFile(wb, "extracted_data.xlsx");
    });
});
