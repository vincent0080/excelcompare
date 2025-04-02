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
                
                // 添加data屬性以存儲單元格位置信息
                html += `<tr class="diff-changed" data-cell-ref="${cellRef}">
                    <td>${cellRef} (${colName})</td>
                    <td>${oldValueDisplay}</td>
                    <td>${newValueDisplay}</td>
                </tr>`;
            });
        }
        
        resultsBody.innerHTML = html;
        resultsTable.style.display = 'table';
        
        // 為單元格變更項目添加點擊事件
        addClickEventToChangedCells(oldFileInput.files[0], oldSheetSelect.value);
    }
    
    // 為單元格變更項目添加點擊事件
    function addClickEventToChangedCells(excelFile, sheetName) {
        const changedRows = document.querySelectorAll('#results-body tr.diff-changed');
        
        changedRows.forEach(row => {
            row.style.cursor = 'pointer'; // 改變滑鼠游標樣式，提示可點擊
            row.title = '點擊開啟Excel並跳轉到此儲存格'; // 添加提示文字
            
            // 添加點擊事件
            row.addEventListener('click', function() {
                const cellRef = this.getAttribute('data-cell-ref');
                if (!cellRef) return;
                
                // 如果沒有上傳檔案，提示使用者
                if (!excelFile) {
                    alert('請確保已上傳Excel檔案');
                    return;
                }
                
                // 獲取檔案名稱
                const fileName = excelFile.name;
                
                // 開啟Excel並跳轉到指定儲存格
                openExcelAndNavigateToCell(cellRef, sheetName, fileName);
            });
        });
    }
    
    // 開啟Excel並跳轉到指定儲存格
    function openExcelAndNavigateToCell(cellRef, sheetName, fileName) {
        // 嘗試多種方法開啟Excel並跳轉到指定儲存格
        
        // 創建一個模態對話框，提供使用者選項
        const modalContainer = document.createElement('div');
        modalContainer.className = 'excel-modal-container';
        modalContainer.style.position = 'fixed';
        modalContainer.style.top = '0';
        modalContainer.style.left = '0';
        modalContainer.style.width = '100%';
        modalContainer.style.height = '100%';
        modalContainer.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
        modalContainer.style.display = 'flex';
        modalContainer.style.justifyContent = 'center';
        modalContainer.style.alignItems = 'center';
        modalContainer.style.zIndex = '1000';
        
        const modalContent = document.createElement('div');
        modalContent.className = 'excel-modal-content';
        modalContent.style.backgroundColor = 'white';
        modalContent.style.padding = '20px';
        modalContent.style.borderRadius = '5px';
        modalContent.style.maxWidth = '500px';
        modalContent.style.width = '90%';
        modalContent.style.boxShadow = '0 4px 8px rgba(0, 0, 0, 0.2)';
        
        modalContent.innerHTML = `
            <h3 style="margin-top: 0;">開啟Excel檔案</h3>
            <p>您想要跳轉到工作表「<strong>${sheetName}</strong>」的儲存格「<strong>${cellRef}</strong>」</p>
            <p>請選擇開啟方式：</p>
            <div style="display: flex; flex-direction: column; gap: 10px;">
                <button id="open-excel-app" style="padding: 10px; cursor: pointer;">1. 開啟Excel應用程式</button>
                <button id="select-excel-file" style="padding: 10px; cursor: pointer;">2. 選擇Excel檔案</button>
                <button id="copy-cell-info" style="padding: 10px; cursor: pointer;">3. 複製儲存格資訊</button>
                <button id="close-modal" style="padding: 10px; margin-top: 10px; cursor: pointer;">取消</button>
            </div>
            <p style="margin-top: 15px; font-size: 0.9em; color: #666;">提示：由於瀏覽器安全限制，可能無法直接跳轉到特定儲存格。</p>
        `;
        
        modalContainer.appendChild(modalContent);
        document.body.appendChild(modalContainer);
        
        // 開啟Excel應用程式
        document.getElementById('open-excel-app').addEventListener('click', function() {
            try {
                const tempLink = document.createElement('a');
                tempLink.href = 'ms-excel:nft|u|';
                tempLink.style.display = 'none';
                document.body.appendChild(tempLink);
                tempLink.click();
                
                alert(`Excel應用程式已開啟。請手動開啟檔案「${fileName}」，然後前往工作表「${sheetName}」的儲存格「${cellRef}」`);
                
                setTimeout(() => {
                    document.body.removeChild(tempLink);
                }, 100);
                
                document.body.removeChild(modalContainer);
            } catch (error) {
                console.error('開啟Excel時發生錯誤:', error);
                alert('無法自動開啟Excel應用程式，請手動開啟。');
            }
        });
        
        // 選擇Excel檔案
        document.getElementById('select-excel-file').addEventListener('click', function() {
            const fileSelector = document.createElement('input');
            fileSelector.type = 'file';
            fileSelector.accept = '.xlsx, .xls';
            fileSelector.style.display = 'none';
            document.body.appendChild(fileSelector);
            
            fileSelector.addEventListener('change', function() {
                if (this.files.length > 0) {
                    const selectedFileName = this.files[0].name;
                    
                    // 嘗試使用Office URI Schemes
                    try {
                        // 嘗試使用ms-excel:ofv協議
                        const tempLink = document.createElement('a');
                        tempLink.href = 'ms-excel:nft|u|';
                        tempLink.click();
                        
                        alert(`Excel應用程式已開啟。請手動開啟檔案「${selectedFileName}」，然後前往工作表「${sheetName}」的儲存格「${cellRef}」`);
                        
                        document.body.removeChild(tempLink);
                    } catch (error) {
                        console.error('開啟Excel時發生錯誤:', error);
                        alert(`請手動開啟檔案「${selectedFileName}」，然後前往工作表「${sheetName}」的儲存格「${cellRef}」`);
                    }
                }
                document.body.removeChild(fileSelector);
            });
            
            fileSelector.click();
            document.body.removeChild(modalContainer);
        });
        
        // 複製儲存格資訊
        document.getElementById('copy-cell-info').addEventListener('click', function() {
            const cellInfo = `工作表: ${sheetName}, 儲存格: ${cellRef}`;
            
            // 使用Clipboard API複製文字
            navigator.clipboard.writeText(cellInfo).then(function() {
                alert(`儲存格資訊已複製到剪貼簿：${cellInfo}`);
            }, function() {
                // 如果Clipboard API失敗，使用傳統方法
                const textarea = document.createElement('textarea');
                textarea.value = cellInfo;
                textarea.style.position = 'fixed';
                document.body.appendChild(textarea);
                textarea.focus();
                textarea.select();
                
                try {
                    const successful = document.execCommand('copy');
                    if (successful) {
                        alert(`儲存格資訊已複製到剪貼簿：${cellInfo}`);
                    } else {
                        alert(`無法複製儲存格資訊：${cellInfo}`);
                    }
                } catch (err) {
                    alert(`無法複製儲存格資訊：${cellInfo}`);
                }
                
                document.body.removeChild(textarea);
            });
            
            document.body.removeChild(modalContainer);
        });
        
        // 關閉模態對話框
        document.getElementById('close-modal').addEventListener('click', function() {
            document.body.removeChild(modalContainer);
        });
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