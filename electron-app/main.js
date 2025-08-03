const { app, BrowserWindow, Menu, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

let mainWindow;

function createWindow() {
  // 创建浏览器窗口
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets/icon.png')
  });

  // 加载应用的 index.html
  mainWindow.loadFile('index.html');

  // 打开开发者工具（开发时使用）
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }

  // 当窗口被关闭时触发
  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // 创建菜单
  createMenu();
}

function createMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        {
          label: '导入Excel',
          accelerator: 'CmdOrCtrl+O',
          click: () => {
            importExcel();
          }
        },
        {
          label: '导出Excel',
          accelerator: 'CmdOrCtrl+S',
          click: () => {
            exportExcel();
          }
        },
        { type: 'separator' },
        {
          label: '退出',
          accelerator: process.platform === 'darwin' ? 'Cmd+Q' : 'Ctrl+Q',
          click: () => {
            app.quit();
          }
        }
      ]
    },
    {
      label: '编辑',
      submenu: [
        { role: 'undo', label: '撤销' },
        { role: 'redo', label: '重做' },
        { type: 'separator' },
        { role: 'cut', label: '剪切' },
        { role: 'copy', label: '复制' },
        { role: 'paste', label: '粘贴' }
      ]
    },
    {
      label: '视图',
      submenu: [
        { role: 'reload', label: '重新加载' },
        { role: 'forceReload', label: '强制重新加载' },
        { role: 'toggleDevTools', label: '切换开发者工具' },
        { type: 'separator' },
        { role: 'resetZoom', label: '实际大小' },
        { role: 'zoomIn', label: '放大' },
        { role: 'zoomOut', label: '缩小' },
        { type: 'separator' },
        { role: 'togglefullscreen', label: '切换全屏' }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '关于',
              message: 'POS CRM 导入工具',
              detail: '版本 1.0.0\n用于导入和管理客户数据的工具'
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

async function importExcel() {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  if (!result.canceled && result.filePaths.length > 0) {
    const filePath = result.filePaths[0];
    try {
      const workbook = XLSX.readFile(filePath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet);
      
      // 发送数据到渲染进程
      mainWindow.webContents.send('excel-data', data);
      
      dialog.showMessageBox(mainWindow, {
        type: 'info',
        title: '导入成功',
        message: `成功导入 ${data.length} 条记录`
      });
    } catch (error) {
      dialog.showErrorBox('导入失败', `无法读取文件: ${error.message}`);
    }
  }
}

async function exportExcel() {
  const result = await dialog.showSaveDialog(mainWindow, {
    filters: [
      { name: 'Excel Files', extensions: ['xlsx'] }
    ],
    defaultPath: '客户数据导出.xlsx'
  });

  if (!result.canceled) {
    // 从渲染进程获取数据
    mainWindow.webContents.send('request-export-data');
  }
}

// 处理来自渲染进程的导出数据
ipcMain.on('export-data', (event, data) => {
  dialog.showSaveDialog(mainWindow, {
    filters: [
      { name: 'Excel Files', extensions: ['xlsx'] }
    ],
    defaultPath: '客户数据导出.xlsx'
  }).then(result => {
    if (!result.canceled) {
      try {
        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, '客户数据');
        XLSX.writeFile(workbook, result.filePath);
        
        dialog.showMessageBox(mainWindow, {
          type: 'info',
          title: '导出成功',
          message: `数据已导出到: ${result.filePath}`
        });
      } catch (error) {
        dialog.showErrorBox('导出失败', `无法保存文件: ${error.message}`);
      }
    }
  });
});

// 当 Electron 完成初始化并准备创建浏览器窗口时调用此方法
app.whenReady().then(createWindow);

// 当所有窗口都关闭时退出应用
app.on('window-all-closed', () => {
  // 在 macOS 上，应用和菜单栏通常会保持活跃状态
  // 直到用户明确地使用 Cmd + Q 退出
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  // 在 macOS 上，当点击 dock 图标并且没有其他窗口打开时
  // 通常会重新创建一个窗口
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// 在这个文件中你可以包含应用特定的主进程代码
// 你也可以将它们放在单独的文件中并在这里引入