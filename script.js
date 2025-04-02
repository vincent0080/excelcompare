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
    
    // 比對整本活頁簿按鈕點擊事件
    compareAllBtn.addEventListener('click', function() {
        if (!oldWorkbook || !newWorkbook) {
            alert('請先上傳兩個Excel檔案');
            return;
        }
        
        // 比較所有工作表
        compareAllSheets(oldWorkbook, newWorkbook);
    });
    
    // 比較所有工作表的差異
    function compareAllSheets(oldWorkbook, newWorkbook) {
        // 清空結果表格
        resultsBody.innerHTML = '';
        
        // 獲取所有工作表名稱
        const oldSheetNames = oldWorkbook.SheetNames;
        const newSheetNames = newWorkbook.SheetNames;
        
        // 合併所有工作表名稱（不重複）
        const allSheetNames = [...new Set([...oldSheetNames, ...newSheetNames])];
        
        // 追蹤差異總數
        let totalDifferences = 0;
        let sheetResults = [];
        
        // 比較每個工作表
        allSheetNames.forEach(sheetName => {
            const oldSheetExists = oldSheetNames.includes(sheetName);
            const newSheetExists = newSheetNames.includes(sheetName);
            
            // 如果工作表在兩個檔案中都存在，則比較內容
            if (oldSheetExists && newSheetExists) {
                const oldSheet = oldWorkbook.Sheets[sheetName];
                const newSheet = newWorkbook.Sheets[sheetName];
                
                // 將工作表轉換為JSON
                const oldData = XLSX.utils.sheet_to_json(oldSheet, {header: 1});
                const newData = XLSX.utils.sheet_to_json(newSheet, {header: 1});
                
                // 尋找可能的主鍵列
                const keyColumnIndex = findPossibleKeyColumn(oldData, newData);
                
                // 使用智能比對算法
                const differences = smartCompare(oldData, newData, keyColumnIndex);
                
                // 記錄結果
                sheetResults.push({
                    sheetName,
                    status: 'changed',
                    differences,
                    oldData,
                    newData
                });
                
                totalDifferences += differences.length;
            } else if (oldSheetExists) {
                // 工作表在舊檔案中存在，但在新檔案中不存在
                sheetResults.push({
                    sheetName,
                    status: 'deleted',
                    differences: [{ type: 'sheet_deleted' }]
                });
                
                totalDifferences += 1;
            } else if (newSheetExists) {
                // 工作表在新檔案中存在，但在舊檔案中不存在
                sheetResults.push({
                    sheetName,
                    status: 'added',
                    differences: [{ type: 'sheet_added' }]
                });
                
                totalDifferences += 1;
            }
        });
        
        // 顯示整體結果
        displayAllSheetsResults(sheetResults, totalDifferences);
        
        // 顯示結果區域
        resultsSection.style.display = 'block';
        resultsSection.scrollIntoView({behavior: 'smooth'});
    }
    
    // 顯示所有工作表的比較結果
    function displayAllSheetsResults(sheetResults, totalDifferences) {
        // 更新摘要信息
        if (totalDifferences === 0) {
            summary.innerHTML = '<p>所有工作表完全相同，沒有發現差異。</p>';
            resultsTable.style.display = 'none';
            return;
        }
        
        // 顯示總差異數
        summary.innerHTML = `<p>共發現 ${totalDifferences} 個差異，涉及 ${sheetResults.length} 個工作表</p>`;
        
        // 創建結果表格
        let html = '';
        
        // 遍歷每個工作表的結果
        sheetResults.forEach((result, index) => {
            // 添加工作表標題
            html += `<tr><td colspan="3" class="diff-sheet-header">${result.sheetName} 工作表</td></tr>`;
            
            // 處理工作表狀態
            if (result.status === 'added') {
                html += `<tr class="diff-added">
                    <td>工作表變更</td>
                    <td>不存在</td>
                    <td>新增工作表</td>
                </tr>`;
            } else if (result.status === 'deleted') {
                html += `<tr class="diff-removed">
                    <td>工作表變更</td>
                    <td>已刪除工作表</td>
                    <td>不存在</td>
                </tr>`;
            } else if (result.differences.length === 0) {
                html += `<tr>
                    <td colspan="3">此工作表完全相同，沒有發現差異</td>
                </tr>`;
            } else {
                // 獲取標題行
                const oldHeaders = result.oldData && result.oldData.length > 0 ? result.oldData[0] : [];
                const newHeaders = result.newData && result.newData.length > 0 ? result.newData[0] : [];
                
                // 分類差異
                const changedCells = result.differences.filter(d => d.type === 'changed');
                const addedRows = result.differences.filter(d => d.type === 'added_row');
                const deletedRows = result.differences.filter(d => d.type === 'deleted_row');
                const addedColumns = result.differences.filter(d => d.type === 'added_column');
                const deletedColumns = result.differences.filter(d => d.type === 'deleted_column');
                
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
            }
            
            // 添加分隔線（除了最後一個工作表）
            if (index < sheetResults.length - 1) {
                html += '<tr><td colspan="3" class="diff-separator"></td></tr>';
            }
        });
        
        resultsBody.innerHTML = html;
        resultsTable.style.display = 'table';
    }

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
        
        // 建立舊數據的映射表，以主鍵為索引
        oldData.forEach((row, rowIndex) => {
            if (row.length > 0) {
                const keyValue = row[keyColumnIndex];
                if (keyValue !== undefined && keyValue !== '') {
                    oldRowMap.set(String(keyValue), { row, rowIndex });
                } else {
                    // 對於沒有主鍵的行，我們將整行內容作為鍵值
                    const rowContent = row.join('|');
                    oldRowMap.set(rowContent, { row, rowIndex });
                }
            }
        });
        
        // 建立新數據的映射表，用於後續比較
        newData.forEach((row, rowIndex) => {
            if (row.length > 0) {
                const keyValue = row[keyColumnIndex];
                if (keyValue !== undefined && keyValue !== '') {
                    newRowMap.set(String(keyValue), { row, rowIndex });
                } else {
                    // 對於沒有主鍵的行，我們將整行內容作為鍵值
                    const rowContent = row.join('|');
                    newRowMap.set(rowContent, { row, rowIndex });
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
            } else {
                // 對於沒有主鍵的行，嘗試通過行內容匹配
                const rowContent = newRow.join('|');
                const oldRowInfo = oldRowMap.get(rowContent);
                
                if (oldRowInfo) {
                    matchFound = true;
                    matchedOldRows.add(rowContent);
                    // 如果內容完全相同，則不需要添加差異
                }
            }
            
            if (!matchFound) {
                // 如果沒有找到匹配，標記為新增行
                differences.push({
                    type: 'added_row',
                    newRowIndex,
                    row: newRow
                });
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
                // 對於沒有主鍵的行，使用行內容作為鍵值
                key = oldRow.join('|');
            }
            
            if (!matchedOldRows.has(key)) {
                // 檢查這一行是否真的被刪除，而不是僅僅因為行號變化
                const rowContent = oldRow.join('|');
                const newRowInfo = newRowMap.get(rowContent);
                
                if (!newRowInfo) {
                    // 如果在新數據中找不到相同內容的行，才標記為刪除
                    differences.push({
                        type: 'deleted_row',
                        oldRowIndex,
                        row: oldRow
                    });
                }
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