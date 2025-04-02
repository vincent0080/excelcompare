document.addEventListener('DOMContentLoaded', function() {
    // 獲取DOM元素
    const oldFileInput = document.getElementById('old-file');
    const newFileInput = document.getElementById('new-file');
    const oldFileName = document.getElementById('old-file-name');
    const newFileName = document.getElementById('new-file-name');
    const oldSheetSelect = document.getElementById('old-sheet-select');
    const newSheetSelect = document.getElementById('new-sheet-select');
    const compareBtn = document.getElementById('compare-btn');
    const resultsSection = document.getElementById('results-section');
    const summary = document.getElementById('summary');
    const resultsTable = document.getElementById('results-table');
    const resultsBody = document.getElementById('results-body');

    // 存儲上傳的檔案和工作表數據
    let oldWorkbook = null;
    let newWorkbook = null;

    // 檢查是否可以啟用比較按鈕
    function checkEnableCompareButton() {
        // 當兩個檔案都上傳後，啟用比較按鈕
        if (oldFileInput.files.length > 0 && newFileInput.files.length > 0) {
            compareBtn.disabled = false;
        } else {
            compareBtn.disabled = true;
        }
    }

    // 處理舊版檔案上傳
    oldFileInput.addEventListener('change', function(e) {
        if (this.files.length > 0) {
            const file = this.files[0];
            oldFileName.textContent = file.name;
            
            // 讀取Excel檔案
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    oldWorkbook = XLSX.read(data, {type: 'array'});
                    
                    // 更新工作表選擇下拉選單
                    oldSheetSelect.innerHTML = '';
                    oldWorkbook.SheetNames.forEach(function(sheetName) {
                        const option = document.createElement('option');
                        option.value = sheetName;
                        option.textContent = sheetName;
                        oldSheetSelect.appendChild(option);
                    });
                    
                    oldSheetSelect.disabled = false;
                    
                    // 檢查是否可以啟用比較按鈕
                    checkEnableCompareButton();
                } catch (error) {
                    console.error('讀取舊版Excel檔案時發生錯誤:', error);
                    alert('無法讀取Excel檔案，請確認檔案格式正確。');
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            oldFileName.textContent = '';
            oldSheetSelect.innerHTML = '<option value="">請先上傳檔案</option>';
            oldSheetSelect.disabled = true;
            oldWorkbook = null;
            checkEnableCompareButton();
        }
    });

    // 處理新版檔案上傳
    newFileInput.addEventListener('change', function(e) {
        if (this.files.length > 0) {
            const file = this.files[0];
            newFileName.textContent = file.name;
            
            // 讀取Excel檔案
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    newWorkbook = XLSX.read(data, {type: 'array'});
                    
                    // 更新工作表選擇下拉選單
                    newSheetSelect.innerHTML = '';
                    newWorkbook.SheetNames.forEach(function(sheetName) {
                        const option = document.createElement('option');
                        option.value = sheetName;
                        option.textContent = sheetName;
                        newSheetSelect.appendChild(option);
                    });
                    
                    newSheetSelect.disabled = false;
                    
                    // 檢查是否可以啟用比較按鈕
                    checkEnableCompareButton();
                } catch (error) {
                    console.error('讀取新版Excel檔案時發生錯誤:', error);
                    alert('無法讀取Excel檔案，請確認檔案格式正確。');
                }
            };
            reader.readAsArrayBuffer(file);
        } else {
            newFileName.textContent = '';
            newSheetSelect.innerHTML = '<option value="">請先上傳檔案</option>';
            newSheetSelect.disabled = true;
            newWorkbook = null;
            checkEnableCompareButton();
        }
    });

    // 比較按鈕點擊事件
    compareBtn.addEventListener('click', function() {
        if (!oldWorkbook || !newWorkbook) {
            alert('請先上傳兩個Excel檔案');
            return;
        }

        const oldSheetName = oldSheetSelect.value;
        const newSheetName = newSheetSelect.value;

        if (!oldSheetName || !newSheetName) {
            alert('請選擇要比較的工作表');
            return;
        }

        // 獲取工作表數據
        const oldSheet = oldWorkbook.Sheets[oldSheetName];
        const newSheet = newWorkbook.Sheets[newSheetName];

        // 將工作表轉換為JSON
        const oldData = XLSX.utils.sheet_to_json(oldSheet, {header: 1});
        const newData = XLSX.utils.sheet_to_json(newSheet, {header: 1});

        // 比較兩個工作表的差異
        compareSheets(oldData, newData);
    });

    // 比較兩個工作表的差異
    function compareSheets(oldData, newData) {
        // 清空結果表格
        resultsBody.innerHTML = '';

        let differences = 0;
        const maxRows = Math.max(oldData.length, newData.length);
        const maxCols = Math.max(
            ...oldData.map(row => row.length),
            ...newData.map(row => row.length)
        );

        // 比較每個儲存格
        for (let row = 0; row < maxRows; row++) {
            for (let col = 0; col < maxCols; col++) {
                // 獲取儲存格值，並確保undefined和null值被轉換為空字串
                let oldValue = row < oldData.length && col < oldData[row].length ? oldData[row][col] : '';
                let newValue = row < newData.length && col < newData[row].length ? newData[row][col] : '';
                
                // 處理undefined和null值
                oldValue = oldValue === undefined || oldValue === null ? '' : oldValue;
                newValue = newValue === undefined || newValue === null ? '' : newValue;

                // 如果值不同，添加到結果表格
                if (oldValue !== newValue) {
                    differences++;
                    const tr = document.createElement('tr');
                    
                    // 儲存格位置（Excel風格的A1, B2等）
                    const cellRef = XLSX.utils.encode_cell({r: row, c: col});
                    
                    // 確保顯示時不會出現undefined字樣
                    tr.innerHTML = `
                        <td>${cellRef}</td>
                        <td>${String(oldValue)}</td>
                        <td>${String(newValue)}</td>
                    `;
                    
                    resultsBody.appendChild(tr);
                }
            }
        }

        // 更新摘要信息
        if (differences > 0) {
            summary.innerHTML = `<p>找到 ${differences} 個差異</p>`;
            resultsTable.style.display = 'table';
        } else {
            summary.innerHTML = '<p>沒有發現差異</p>';
            resultsTable.style.display = 'none';
        }

        // 顯示結果區域
        resultsSection.style.display = 'block';
        resultsSection.scrollIntoView({behavior: 'smooth'});
    }

    // 初始化拖放功能
    ['old-file-area', 'new-file-area'].forEach(function(id) {
        const dropArea = document.getElementById(id);
        const fileInput = dropArea.querySelector('input[type="file"]');
        
        // 阻止默認拖放行為
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function(eventName) {
            dropArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        // 高亮拖放區域
        ['dragenter', 'dragover'].forEach(function(eventName) {
            dropArea.addEventListener(eventName, function() {
                dropArea.classList.add('highlight');
            }, false);
        });

        ['dragleave', 'drop'].forEach(function(eventName) {
            dropArea.addEventListener(eventName, function() {
                dropArea.classList.remove('highlight');
            }, false);
        });

        // 處理拖放的檔案
        dropArea.addEventListener('drop', function(e) {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if (files.length > 0) {
                fileInput.files = files;
                // 觸發change事件
                const event = new Event('change');
                fileInput.dispatchEvent(event);
            }
        }, false);
    });
});