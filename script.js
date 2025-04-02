document.addEventListener('DOMContentLoaded', function() {
    // 獲取DOM元素
    const oldFileInput = document.getElementById('old-file');
    const newFileInput = document.getElementById('new-file');
    const oldFileName = document.getElementById('old-file-name');
    const newFileName = document.getElementById('new-file-name');
    const oldSheetSelect = document.getElementById('old-sheet-select');
    const newSheetSelect = document.getElementById('new-sheet-select');
    const compareBtn = document.getElementById('compare-btn');
    const compareAllBtn = document.getElementById('compare-all-btn');
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
            compareAllBtn.disabled = false;
        } else {
            compareBtn.disabled = true;
            compareAllBtn.disabled = true;
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
        
        // 建立舊數據的映射表，以主鍵為索引
        oldData.forEach((row, rowIndex) => {
            if (row.length > 0) {
                const keyValue = row[keyColumnIndex];
                if (keyValue !== undefined && keyValue !== '') {
                    oldRowMap.set(String(keyValue), { row, rowIndex });
                } else {
                    // 對於沒有主鍵的行，使用行的內容作為標識而不是行號
                    // 這樣可以更準確地匹配行，避免誤報
                    const rowContent = JSON.stringify(row);
                    oldRowMap.set(rowContent, { row, rowIndex });
                }
            }
        });
        
        // 追蹤已匹配的舊行
        const matchedOldRows = new Set();
        
        // 比較新數據與舊數據
        newData.forEach((newRow, newRowIndex) => {
            if (newRow.length === 0) return;
            
            const newKeyValue = newRow[keyColumnIndex];
            let matchFound = false;
            
            if (newKeyValue !== undefined && newKeyValue !== '') {
                // 嘗試通過主鍵匹配
                const oldRowInfo = oldRowMap.get(String(newKeyValue));
                
                if (oldRowInfo) {
                    matchFound = true;
                    matchedOldRows.add(String(newKeyValue));
                    
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
                                newRowIndex,
                                colIndex,
                                oldValue,
                                newValue
                            });
                        }
                    }
                }
            }
            
            if (!matchFound && newKeyValue !== undefined && newKeyValue !== '') {
                // 如果通過主鍵沒有找到匹配，標記為新增行
                differences.push({
                    type: 'added_row',
                    newRowIndex,
                    row: newRow
                });
            } else if (!matchFound) {
                // 對於沒有主鍵的行，嘗試通過行內容匹配
                const rowContent = JSON.stringify(newRow);
                const oldRowInfo = oldRowMap.get(rowContent);
                
                if (oldRowInfo) {
                    matchFound = true;
                    matchedOldRows.add(rowContent);
                    
                    // 不需要比較行內容，因為它們已經完全匹配
                } else {
                    // 如果仍然沒有找到匹配，標記為新增行
                    differences.push({
                        type: 'added_row',
                        newRowIndex,
                        row: newRow
                    });
                }
            }
        });
        
        // 查找刪除的行
        oldData.forEach((oldRow, oldRowIndex) => {
            if (oldRow.length === 0) return;
            
            const oldKeyValue = oldRow[keyColumnIndex];
            let key;
            
            if (oldKeyValue !== undefined && oldKeyValue !== '') {
                key = String(oldKeyValue);
            } else {
                // 對於沒有主鍵的行，使用行的內容作為標識而不是行號
                // 這樣可以更準確地匹配行，避免誤報
                key = JSON.stringify(oldRow);
            }
            
            if (!matchedOldRows.has(key)) {
                differences.push({
                    type: 'deleted_row',
                    oldRowIndex,
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
                fileInput.files = files;
                // 觸發change事件
                const event = new Event('change');
                fileInput.dispatchEvent(event);
            }
        }, false);
    });
});

// 比對整本活頁簿按鈕點擊事件
compareAllBtn.addEventListener('click', function() {
    if (!oldWorkbook || !newWorkbook) {
        alert('請先上傳兩個Excel檔案');
        return;
    }

    // 清空結果表格
    resultsBody.innerHTML = '';
    let allDifferences = [];

    // 遍歷所有工作表進行比較
    const oldSheets = new Set(oldWorkbook.SheetNames);
    const newSheets = new Set(newWorkbook.SheetNames);
    const allSheets = new Set([...oldSheets, ...newSheets]);

    allSheets.forEach(sheetName => {
        const oldSheet = oldWorkbook.Sheets[sheetName];
        const newSheet = newWorkbook.Sheets[sheetName];

        if (!oldSheet) {
            // 新增的工作表
            const newData = XLSX.utils.sheet_to_json(newSheet, {header: 1});
            allDifferences.push({
                type: 'added_sheet',
                sheetName: sheetName,
                data: newData
            });
        } else if (!newSheet) {
            // 刪除的工作表
            const oldData = XLSX.utils.sheet_to_json(oldSheet, {header: 1});
            allDifferences.push({
                type: 'deleted_sheet',
                sheetName: sheetName,
                data: oldData
            });
        } else {
            // 比較工作表內容
            const oldData = XLSX.utils.sheet_to_json(oldSheet, {header: 1});
            const newData = XLSX.utils.sheet_to_json(newSheet, {header: 1});
            const differences = smartCompare(oldData, newData, findPossibleKeyColumn(oldData, newData));
            if (differences.length > 0) {
                allDifferences.push({
                    type: 'modified_sheet',
                    sheetName: sheetName,
                    differences: differences
                });
            }
        }
    });

    // 顯示所有差異
    displayAllResults(allDifferences);

    // 顯示結果區域
    resultsSection.style.display = 'block';
    resultsSection.scrollIntoView({behavior: 'smooth'});
});

// 顯示所有工作表的比較結果
function displayAllResults(allDifferences) {
    if (allDifferences.length === 0) {
        summary.innerHTML = '<p>兩個Excel檔案完全相同，沒有發現差異。</p>';
        resultsTable.style.display = 'none';
        return;
    }

    // 更新摘要信息
    summary.innerHTML = `<p>發現 ${allDifferences.length} 個工作表有差異</p>`;

    // 創建結果表格
    let html = '';

    allDifferences.forEach(diff => {
        html += `<tr><td colspan="3" class="diff-section-header">工作表：${diff.sheetName}</td></tr>`;

        if (diff.type === 'added_sheet') {
            html += `<tr class="diff-added">
                <td>新增工作表</td>
                <td></td>
                <td>包含 ${diff.data.length} 行資料</td>
            </tr>`;
        } else if (diff.type === 'deleted_sheet') {
            html += `<tr class="diff-removed">
                <td>刪除工作表</td>
                <td>包含 ${diff.data.length} 行資料</td>
                <td></td>
            </tr>`;
        } else if (diff.type === 'modified_sheet') {
            // 顯示工作表內容的差異
            const changes = diff.differences;
            
            // 分類差異
            const changedCells = changes.filter(d => d.type === 'changed');
            const addedRows = changes.filter(d => d.type === 'added_row');
            const deletedRows = changes.filter(d => d.type === 'deleted_row');
            const addedColumns = changes.filter(d => d.type === 'added_column');
            const deletedColumns = changes.filter(d => d.type === 'deleted_column');

            // 顯示列變更
            if (addedColumns.length > 0 || deletedColumns.length > 0) {
                addedColumns.forEach(change => {
                    html += `<tr class="diff-added">
                        <td>新增列</td>
                        <td></td>
                        <td>列 ${change.colIndex + 1}</td>
                    </tr>`;
                });

                deletedColumns.forEach(change => {
                    html += `<tr class="diff-removed">
                        <td>刪除列</td>
                        <td>列 ${change.colIndex + 1}</td>
                        <td></td>
                    </tr>`;
                });
            }

            // 顯示行變更
            if (addedRows.length > 0 || deletedRows.length > 0) {
                addedRows.forEach(change => {
                    const rowNum = change.newRowIndex + 1;
                    const firstCell = change.row[0] === undefined || change.row[0] === null ? '' : change.row[0];
                    html += `<tr class="diff-added">
                        <td>新增行 ${rowNum}</td>
                        <td></td>
                        <td>${firstCell}...</td>
                    </tr>`;
                });

                deletedRows.forEach(change => {
                    const rowNum = change.oldRowIndex + 1;
                    const firstCell = change.row[0] === undefined || change.row[0] === null ? '' : change.row[0];
                    html += `<tr class="diff-removed">
                        <td>刪除行 ${rowNum}</td>
                        <td>${firstCell}...</td>
                        <td></td>
                    </tr>`;
                });
            }

            // 顯示單元格變更
            if (changedCells.length > 0) {
                changedCells.forEach(change => {
                    const cellRef = XLSX.utils.encode_cell({r: change.oldRowIndex, c: change.colIndex});
                    const oldValueDisplay = change.oldValue === undefined || change.oldValue === null ? '' : change.oldValue;
                    const newValueDisplay = change.newValue === undefined || change.newValue === null ? '' : change.newValue;
                    html += `<tr class="diff-changed">
                        <td>${cellRef}</td>
                        <td>${oldValueDisplay}</td>
                        <td>${newValueDisplay}</td>
                    </tr>`;
                });
            }
        }
    });

    resultsBody.innerHTML = html;
    resultsTable.style.display = 'table';
}