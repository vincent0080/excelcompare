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

        // 尋找可能的主鍵列（通常是第一列或包含ID的列）
        const keyColumnIndex = findPossibleKeyColumn(oldData, newData);
        
        // 使用智能比對算法
        const differences = smartCompare(oldData, newData, keyColumnIndex);
        
        // 顯示結果
        displayResults(differences, oldData, newData);

        // 顯示結果區域
        resultsSection.style.display = 'block';
        resultsSection.scrollIntoView({behavior: 'smooth'});
    }
    
    // 尋找可能的主鍵列
    function findPossibleKeyColumn(oldData, newData) {
        // 如果第一行存在，檢查標題中是否有ID、編號等關鍵字
        if (oldData.length > 0 && newData.length > 0) {
            const oldHeaders = oldData[0];
            
            for (let i = 0; i < oldHeaders.length; i++) {
                const header = String(oldHeaders[i]).toLowerCase();
                if (header.includes('id') || 
                    header.includes('編號') || 
                    header.includes('代碼') || 
                    header.includes('序號')) {
                    return i;
                }
            }
        }
        
        // 默認使用第一列作為主鍵
        return 0;
    }
    
    // 智能比對算法
    function smartCompare(oldData, newData, keyColumnIndex) {
        const differences = [];
        const oldRowMap = new Map();
        const newRowMap = new Map();
        
        // 跳過標題行（如果存在）
        const startIndex = oldData.length > 0 && newData.length > 0 ? 1 : 0;
        
        // 建立舊數據的映射表，以主鍵為索引
        for (let rowIndex = startIndex; rowIndex < oldData.length; rowIndex++) {
            const row = oldData[rowIndex];
            if (row.length === 0) continue;
            
            const keyValue = row[keyColumnIndex];
            if (keyValue !== undefined && keyValue !== '') {
                oldRowMap.set(String(keyValue), { row, rowIndex });
            } else {
                // 對於沒有主鍵的行，嘗試使用整行內容作為標識
                const rowSignature = row.join('|');
                oldRowMap.set(rowSignature, { row, rowIndex });
            }
        }
        
        // 建立新數據的映射表，用於後續比較
        for (let rowIndex = startIndex; rowIndex < newData.length; rowIndex++) {
            const row = newData[rowIndex];
            if (row.length === 0) continue;
            
            const keyValue = row[keyColumnIndex];
            if (keyValue !== undefined && keyValue !== '') {
                newRowMap.set(String(keyValue), { row, rowIndex });
            } else {
                // 對於沒有主鍵的行，嘗試使用整行內容作為標識
                const rowSignature = row.join('|');
                newRowMap.set(rowSignature, { row, rowIndex });
            }
        }
        
        // 追蹤已匹配的舊行
        const matchedOldRows = new Set();
        const matchedNewRows = new Set();
        
        // 先處理有主鍵的行匹配
        for (let rowIndex = startIndex; rowIndex < newData.length; rowIndex++) {
            const newRow = newData[rowIndex];
            if (newRow.length === 0) continue;
            
            const newKeyValue = newRow[keyColumnIndex];
            let matchFound = false;
            
            if (newKeyValue !== undefined && newKeyValue !== '') {
                // 嘗試通過主鍵匹配
                const oldRowInfo = oldRowMap.get(String(newKeyValue));
                
                if (oldRowInfo) {
                    matchFound = true;
                    matchedOldRows.add(String(newKeyValue));
                    matchedNewRows.add(rowIndex);
                    
                    const { row: oldRow, rowIndex: oldRowIndex } = oldRowInfo;
                    
                    // 比較行內容
                    for (let colIndex = 0; colIndex < Math.max(oldRow.length, newRow.length); colIndex++) {
                        let oldValue = colIndex < oldRow.length ? oldRow[colIndex] : '';
                        let newValue = colIndex < newRow.length ? newRow[colIndex] : '';
                        
                        // 確保undefined和null值被轉換為空字串
                        oldValue = oldValue === undefined || oldValue === null ? '' : oldValue;
                        newValue = newValue === undefined || newValue === null ? '' : newValue;
                        
                        if (String(oldValue) !== String(newValue)) {
                            differences.push({
                                type: 'changed',
                                oldRowIndex,
                                newRowIndex: rowIndex,
                                colIndex,
                                oldValue,
                                newValue
                            });
                        }
                    }
                }
            }
        }
        
        // 處理沒有主鍵的行，嘗試通過內容匹配
        for (let rowIndex = startIndex; rowIndex < newData.length; rowIndex++) {
            if (matchedNewRows.has(rowIndex)) continue; // 跳過已匹配的行
            
            const newRow = newData[rowIndex];
            if (newRow.length === 0) continue;
            
            // 使用行內容作為標識
            const rowSignature = newRow.join('|');
            const oldRowInfo = oldRowMap.get(rowSignature);
            
            if (oldRowInfo && !matchedOldRows.has(rowSignature)) {
                matchedOldRows.add(rowSignature);
                matchedNewRows.add(rowIndex);
                // 內容完全相同，不需要記錄差異
            } else if (!matchedNewRows.has(rowIndex)) {
                // 如果沒有找到匹配，標記為新增行
                differences.push({
                    type: 'added_row',
                    newRowIndex: rowIndex,
                    row: newRow
                });
            }
        }
        
        // 查找刪除的行
        for (let rowIndex = startIndex; rowIndex < oldData.length; rowIndex++) {
            const oldRow = oldData[rowIndex];
            if (oldRow.length === 0) continue;
            
            const oldKeyValue = oldRow[keyColumnIndex];
            let key;
            
            if (oldKeyValue !== undefined && oldKeyValue !== '') {
                key = String(oldKeyValue);
            } else {
                // 使用行內容作為標識
                key = oldRow.join('|');
            }
            
            if (!matchedOldRows.has(key)) {
                differences.push({
                    type: 'deleted_row',
                    oldRowIndex: rowIndex,
                    row: oldRow
                });
            }
        });
        
        // 檢測整列的插入或刪除
        const oldColCount = oldData.length > 0 ? oldData[0].length : 0;
        const newColCount = newData.length > 0 ? newData[0].length : 0;
        
        if (oldColCount !== newColCount) {
            // 檢測新增的列
            if (newColCount > oldColCount) {
                for (let colIndex = oldColCount; colIndex < newColCount; colIndex++) {
                    differences.push({
                        type: 'added_column',
                        colIndex
                    });
                }
            }
            
            // 檢測刪除的列
            if (oldColCount > newColCount) {
                for (let colIndex = newColCount; colIndex < oldColCount; colIndex++) {
                    differences.push({
                        type: 'deleted_column',
                        colIndex
                    });
                }
            }
        }
        
        return differences;
    }
    
    // 顯示比較結果
    function displayResults(differences, oldData, newData) {
        if (differences.length === 0) {
            summary.innerHTML = '<p>兩個工作表完全相同，沒有發現差異。</p>';
            resultsTable.style.display = 'none';
            return;
        }
        
        // 獲取標題行
        const oldHeaders = oldData.length > 0 ? oldData[0] : [];
        const newHeaders = newData.length > 0 ? newData[0] : [];
        
        // 分類差異
        const changedCells = differences.filter(d => d.type === 'changed');
        const addedRows = differences.filter(d => d.type === 'added_row');
        const deletedRows = differences.filter(d => d.type === 'deleted_row');
        const addedColumns = differences.filter(d => d.type === 'added_column');
        const deletedColumns = differences.filter(d => d.type === 'deleted_column');
        
        // 更新摘要信息
        summary.innerHTML = `<p>發現 ${differences.length} 個差異</p>`;
        
        // 創建結果表格
        let html = '';
        
        // 顯示列變更
        if (addedColumns.length > 0 || deletedColumns.length > 0) {
            html += '<tr><td colspan="3" class="diff-section-header">列變更</td></tr>';
            
            addedColumns.forEach(diff => {
                const colName = newHeaders[diff.colIndex] || `列 ${diff.colIndex + 1}`;
                html += `<tr class="diff-added">
                    <td>新增列</td>
                    <td></td>
                    <td>${colName}</td>
                </tr>`;
            });
            
            deletedColumns.forEach(diff => {
                const colName = oldHeaders[diff.colIndex] || `列 ${diff.colIndex + 1}`;
                html += `<tr class="diff-removed">
                    <td>刪除列</td>
                    <td>${colName}</td>
                    <td></td>
                </tr>`;
            });
        }
        
        // 顯示行變更
        if (addedRows.length > 0 || deletedRows.length > 0) {
            html += '<tr><td colspan="3" class="diff-section-header">行變更</td></tr>';
            
            addedRows.forEach(diff => {
                const rowNum = diff.newRowIndex + 1;
                // 確保firstCell不是undefined或null
                const firstCell = diff.row[0] === undefined || diff.row[0] === null ? '' : diff.row[0];
                html += `<tr class="diff-added">
                    <td>新增行 ${rowNum}</td>
                    <td></td>
                    <td>${firstCell}...</td>
                </tr>`;
            });
            
            deletedRows.forEach(diff => {
                const rowNum = diff.oldRowIndex + 1;
                // 確保firstCell不是undefined或null
                const firstCell = diff.row[0] === undefined || diff.row[0] === null ? '' : diff.row[0];
                html += `<tr class="diff-removed">
                    <td>刪除行 ${rowNum}</td>
                    <td>${firstCell}...</td>
                    <td></td>
                </tr>`;
            });
        }
        
        // 顯示單元格變更
        if (changedCells.length > 0) {
            html += '<tr><td colspan="3" class="diff-section-header">單元格變更</td></tr>';
            
            changedCells.forEach(diff => {
                const colName = oldHeaders[diff.colIndex] || newHeaders[diff.colIndex] || `列 ${diff.colIndex + 1}`;
                const cellRef = XLSX.utils.encode_cell({r: diff.oldRowIndex, c: diff.colIndex});
                
                // 確保undefined和null值被轉換為空字串
                const oldValueDisplay = diff.oldValue === undefined || diff.oldValue === null ? '' : diff.oldValue;
                const newValueDisplay = diff.newValue === undefined || diff.newValue === null ? '' : diff.newValue;
                
                html += `<tr class="diff-changed">
                    <td>${cellRef} (${colName})</td>
                    <td>${oldValueDisplay}</td>
                    <td>${newValueDisplay}</td>
                </tr>`;
            });
        }
        
        resultsBody.innerHTML = html;
        resultsTable.style.display = 'table';
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
                // 無法直接設置fileInput.files，改為使用模擬點擊方式
                // 先清除現有的選擇
                fileInput.value = '';
                
                // 如果是舊版檔案上傳區域，則設置oldFileInput的檔案
                if (id === 'old-file-area') {
                    oldFileInput.files = files;
                    oldFileName.textContent = files[0].name;
                    
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
                    reader.readAsArrayBuffer(files[0]);
                } 
                // 如果是新版檔案上傳區域，則設置newFileInput的檔案
                else if (id === 'new-file-area') {
                    newFileInput.files = files;
                    newFileName.textContent = files[0].name;
                    
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
                    reader.readAsArrayBuffer(files[0]);
                }
            }
        }, false);
    });
});