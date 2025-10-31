// 全局变量
let workbook = null;
let currentSheet = null;
let selectedCell = null;
let jumpConfigs = new Map(); // 存储跳转配置

// 加载Excel文件
function loadExcel() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    
    if (!file) {
        showStatus('请选择Excel文件', 'error');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            
            // 显示工作表选择器
            const sheetSelect = document.getElementById('sheetSelect');
            sheetSelect.innerHTML = '<option disabled selected>请选择工作表</option>';
            
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            
            document.getElementById('sheetSelector').classList.remove('hidden');
            document.getElementById('actionButtons').classList.remove('hidden');
            document.getElementById('jumpConfig').classList.remove('hidden');
            
            showStatus('Excel文件加载成功！', 'success');
        } catch (error) {
            showStatus('加载Excel文件失败: ' + error.message, 'error');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 显示工作表内容
function displaySheet() {
    const sheetName = document.getElementById('sheetSelect').value;
    if (!sheetName || !workbook) return;
    
    currentSheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(currentSheet, { header: 1 });
    
    displayTable(data);
    document.getElementById('excelDisplay').classList.remove('hidden');
}

// 显示表格
function displayTable(data) {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');
    
    // 清空表格
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';
    
    if (data.length === 0) return;
    
    // 创建表头
    const headerRow = document.createElement('tr');
    const maxCols = Math.max(...data.map(row => row.length));
    
    // 添加列字母
    for (let i = 0; i < maxCols; i++) {
        const th = document.createElement('th');
        th.textContent = getColumnLetter(i);
        headerRow.appendChild(th);
    }
    tableHead.appendChild(headerRow);
    
    // 创建表格内容
    data.forEach((row, rowIndex) => {
        const tr = document.createElement('tr');
        
        for (let colIndex = 0; colIndex < maxCols; colIndex++) {
            const td = document.createElement('td');
            const cellAddress = getColumnLetter(colIndex) + (rowIndex + 1);
            
            // 检查是否有跳转配置
            if (jumpConfigs.has(cellAddress)) {
                const config = jumpConfigs.get(cellAddress);
                td.innerHTML = `<input type="checkbox" class="checkbox checkbox-primary" 
                    data-cell="${cellAddress}" onchange="handleCheckboxChange(this)">`;
                td.title = `跳转到: ${config.targetFile || '当前文件'} - ${config.targetSheet} - ${config.targetCell}`;
            } else {
                td.textContent = row[colIndex] || '';
                td.onclick = () => selectCell(cellAddress, td);
            }
            
            td.id = `cell-${cellAddress}`;
            tr.appendChild(td);
        }
        
        tableBody.appendChild(tr);
    });
}

// 获取列字母
function getColumnLetter(colIndex) {
    let letter = '';
    while (colIndex >= 0) {
        letter = String.fromCharCode(65 + (colIndex % 26)) + letter;
        colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
}

// 选择单元格
function selectCell(cellAddress, element) {
    // 移除之前的选择
    document.querySelectorAll('td').forEach(td => {
        td.classList.remove('bg-blue-200');
    });
    
    // 高亮当前选择
    element.classList.add('bg-blue-200');
    selectedCell = cellAddress;
    
    // 更新配置区域
    const config = jumpConfigs.get(cellAddress);
    if (config) {
        document.getElementById('targetFile').value = config.targetFile || '';
        document.getElementById('targetSheet').value = config.targetSheet || '';
        document.getElementById('targetCell').value = config.targetCell || '';
    } else {
        document.getElementById('targetFile').value = '';
        document.getElementById('targetSheet').value = '';
        document.getElementById('targetCell').value = '';
    }
}

// 保存跳转配置
function saveJumpConfig() {
    if (!selectedCell) {
        showStatus('请先选择单元格', 'error');
        return;
    }
    
    const targetFile = document.getElementById('targetFile').value;
    const targetSheet = document.getElementById('targetSheet').value;
    const targetCell = document.getElementById('targetCell').value;
    
    if (!targetSheet || !targetCell) {
        showStatus('请填写目标工作表和单元格', 'error');
        return;
    }
    
    jumpConfigs.set(selectedCell, {
        targetFile: targetFile,
        targetSheet: targetSheet,
        targetCell: targetCell
    });
    
    // 重新显示表格以更新复选框
    displaySheet();
    showStatus('跳转配置已保存', 'success');
}

// 添加复选框到选中的单元格
function addCheckboxToSelected() {
    if (!selectedCell) {
        showStatus('请先选择单元格', 'error');
        return;
    }
    
    const targetSheet = prompt('请输入目标工作表名称:');
    const targetCell = prompt('请输入目标单元格 (如 A1):');
    
    if (!targetSheet || !targetCell) {
        showStatus('必须填写目标工作表和单元格', 'error');
        return;
    }
    
    jumpConfigs.set(selectedCell, {
        targetFile: '',
        targetSheet: targetSheet,
        targetCell: targetCell
    });
    
    displaySheet();
    showStatus('复选框已添加', 'success');
}

// 处理复选框变化
function handleCheckboxChange(checkbox) {
    if (checkbox.checked) {
        const cellAddress = checkbox.getAttribute('data-cell');
        const config = jumpConfigs.get(cellAddress);
        
        if (config) {
            jumpToCell(config);
        }
    }
}

// 跳转到目标单元格
function jumpToCell(config) {
    if (config.targetFile) {
        // 跳转到外部文件
        alert(`将打开外部文件: ${config.targetFile}\n工作表: ${config.targetSheet}\n单元格: ${config.targetCell}`);
        // 在实际应用中，这里可以调用系统命令打开文件
    } else {
        // 跳转到当前文件的其他位置
        const sheetName = document.getElementById('sheetSelect').value;
        if (config.targetSheet !== sheetName) {
            // 切换到目标工作表
            document.getElementById('sheetSelect').value = config.targetSheet;
            displaySheet();
        }
        
        // 高亮目标单元格
        setTimeout(() => {
            const targetElement = document.getElementById(`cell-${config.targetCell}`);
            if (targetElement) {
                targetElement.scrollIntoView({ behavior: 'smooth', block: 'center' });
                targetElement.classList.add('bg-yellow-200');
                setTimeout(() => {
                    targetElement.classList.remove('bg-yellow-200');
                }, 2000);
            }
        }, 100);
    }
}

// 测试跳转
function testJump() {
    if (jumpConfigs.size === 0) {
        showStatus('请先配置跳转规则', 'error');
        return;
    }
    
    showStatus('测试模式已激活，请勾选复选框进行跳转测试', 'info');
}

// 导出Excel
function exportExcel() {
    if (!workbook) {
        showStatus('请先加载Excel文件', 'error');
        return;
    }
    
    // 创建新的工作簿
    const newWorkbook = XLSX.utils.book_new();
    
    // 复制所有工作表
    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        
        // 添加跳转配置到工作表
        jumpConfigs.forEach((config, cellAddress) => {
            if (config.targetSheet === sheetName) {
                // 在单元格中添加注释或特殊标记
                if (!worksheet[cellAddress]) {
                    worksheet[cellAddress] = { t: 's', v: '' };
                }
                worksheet[cellAddress].c = [{
                    t: `跳转至: ${config.targetFile || '当前文件'} - ${config.targetSheet} - ${config.targetCell}`
                }];
            }
        });
        
        XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
    });
    
    // 导出文件
    const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = '带跳转配置的Excel.xlsx';
    a.click();
    
    URL.revokeObjectURL(url);
    showStatus('Excel文件已导出', 'success');
}

// 显示状态信息
function showStatus(message, type) {
    const status = document.getElementById('status');
    const statusText = document.getElementById('statusText');
    
    statusText.textContent = message;
    status.className = `alert alert-${type}`;
    status.classList.remove('hidden');
    
    setTimeout(() => {
        status.classList.add('hidden');
    }, 3000);
}

// 页面加载完成后的初始化
document.addEventListener('DOMContentLoaded', function() {
    showStatus('欢迎使用Excel单元格跳转助手！请先上传Excel文件', 'info');
});
