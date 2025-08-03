const { ipcRenderer } = require('electron');
const XLSX = require('xlsx');

let currentData = [];

// DOM 元素
const uploadArea = document.getElementById('uploadArea');
const importBtn = document.getElementById('importBtn');
const exportBtn = document.getElementById('exportBtn');
const importLoading = document.getElementById('importLoading');
const alertContainer = document.getElementById('alertContainer');
const statsContainer = document.getElementById('statsContainer');
const dataTable = document.getElementById('dataTable');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const totalRecords = document.getElementById('totalRecords');
const validRecords = document.getElementById('validRecords');
const invalidRecords = document.getElementById('invalidRecords');

// 事件监听器
importBtn.addEventListener('click', handleImport);
exportBtn.addEventListener('click', handleExport);

// 拖拽上传功能
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        if (isExcelFile(file)) {
            handleFileUpload(file);
        } else {
            showAlert('请选择有效的Excel文件 (.xlsx 或 .xls)', 'error');
        }
    }
});

uploadArea.addEventListener('click', () => {
    importBtn.click();
});

// IPC 通信
ipcRenderer.on('excel-data', (event, data) => {
    currentData = data;
    displayData(data);
    updateStats(data);
    exportBtn.disabled = false;
    hideLoading();
    showAlert(`成功导入 ${data.length} 条记录`, 'success');
});

ipcRenderer.on('request-export-data', () => {
    ipcRenderer.send('export-data', currentData);
});

// 处理导入
function handleImport() {
    showLoading();
    // 触发主进程的文件选择对话框
    // 这里我们通过点击隐藏的文件输入来实现
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.onchange = (e) => {
        const file = e.target.files[0];
        if (file) {
            handleFileUpload(file);
        } else {
            hideLoading();
        }
    };
    input.click();
}

// 处理文件上传
function handleFileUpload(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            currentData = jsonData;
            displayData(jsonData);
            updateStats(jsonData);
            exportBtn.disabled = false;
            hideLoading();
            showAlert(`成功导入 ${jsonData.length} 条记录`, 'success');
        } catch (error) {
            hideLoading();
            showAlert(`文件读取失败: ${error.message}`, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// 处理导出
function handleExport() {
    if (currentData.length === 0) {
        showAlert('没有数据可导出', 'error');
        return;
    }
    
    try {
        const worksheet = XLSX.utils.json_to_sheet(currentData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '客户数据');
        
        // 创建下载链接
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = '客户数据导出.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showAlert('数据导出成功', 'success');
    } catch (error) {
        showAlert(`导出失败: ${error.message}`, 'error');
    }
}

// 显示数据表格
function displayData(data) {
    if (data.length === 0) {
        dataTable.classList.add('hidden');
        return;
    }
    
    // 生成表头
    const headers = Object.keys(data[0]);
    tableHead.innerHTML = `
        <tr>
            ${headers.map(header => `<th>${header}</th>`).join('')}
        </tr>
    `;
    
    // 生成表格内容（只显示前100行以提高性能）
    const displayData = data.slice(0, 100);
    tableBody.innerHTML = displayData.map(row => `
        <tr>
            ${headers.map(header => `<td>${row[header] || ''}</td>`).join('')}
        </tr>
    `).join('');
    
    dataTable.classList.remove('hidden');
    
    if (data.length > 100) {
        showAlert(`显示前100条记录，共${data.length}条记录`, 'success');
    }
}

// 更新统计信息
function updateStats(data) {
    const total = data.length;
    let valid = 0;
    let invalid = 0;
    
    // 简单的数据验证逻辑
    data.forEach(row => {
        const values = Object.values(row);
        const hasValidData = values.some(value => 
            value !== null && 
            value !== undefined && 
            value !== '' && 
            String(value).trim() !== ''
        );
        
        if (hasValidData) {
            valid++;
        } else {
            invalid++;
        }
    });
    
    totalRecords.textContent = total;
    validRecords.textContent = valid;
    invalidRecords.textContent = invalid;
    
    statsContainer.classList.remove('hidden');
}

// 显示提示信息
function showAlert(message, type = 'success') {
    const alertDiv = document.createElement('div');
    alertDiv.className = `alert alert-${type}`;
    alertDiv.textContent = message;
    
    alertContainer.innerHTML = '';
    alertContainer.appendChild(alertDiv);
    
    // 3秒后自动隐藏
    setTimeout(() => {
        if (alertContainer.contains(alertDiv)) {
            alertContainer.removeChild(alertDiv);
        }
    }, 3000);
}

// 显示加载状态
function showLoading() {
    importLoading.classList.remove('hidden');
    importBtn.disabled = true;
}

// 隐藏加载状态
function hideLoading() {
    importLoading.classList.add('hidden');
    importBtn.disabled = false;
}

// 检查是否为Excel文件
function isExcelFile(file) {
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const validExtensions = ['.xlsx', '.xls'];
    
    return validTypes.includes(file.type) || 
           validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));
}

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    console.log('POS CRM 导入工具已启动');
});