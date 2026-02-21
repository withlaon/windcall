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

    function handleFiles(files) {
        if (files.length === 0) return;
        const file = files[0];
        const reader = new FileReader();

        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            try {
                const workbook = XLSX.read(data, { 
                    type: 'array',
                    password: '3674601220'
                });
                
                // Assume first sheet
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                
                // Convert to JSON
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                if (json.length > 0) {
                    processData(json);
                } else {
                    alert('파일에 데이터가 없습니다.');
                }
            } catch (error) {
                console.error('Excel processing error:', error);
                alert('파일을 읽는 중 오류가 발생했습니다. 암호가 맞지 않거나 손상된 파일일 수 있습니다.');
            }
        };

        reader.readAsArrayBuffer(file);
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
        headers.forEach(header => {
            const div = document.createElement('div');
            div.className = 'column-item';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = selectedColumns.has(header);
            checkbox.id = `col-${header}`;
            
            const label = document.createElement('span');
            label.textContent = header;
            
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
            th.textContent = h;
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
