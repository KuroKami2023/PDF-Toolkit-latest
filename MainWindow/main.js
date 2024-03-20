const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const fs = require('fs');
const path = require('path');

let mainWindow;
let hasConfirmedClose = false;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    autoHideMenuBar: true,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      preload: path.join(__dirname, 'preload.js')
    },
    visibleOnAllWorkspaces: true,
  });

  mainWindow.loadFile('./HTML/index.html');
  mainWindow.maximize();
  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  mainWindow.on('close', (event) => {
    if (!hasConfirmedClose) {
      event.preventDefault();
      confirmClose();
    }
  });
}

async function confirmClose() {
  const choice = await dialog.showMessageBox(mainWindow, {
    type: 'question',
    buttons: ['Yes', 'No'],
    title: 'Confirm',
    message: 'Are you sure you want to quit?'
  });

  if (choice.response === 0) {
    hasConfirmedClose = true;
    app.quit();
  } else {
    if (mainWindow) {
      mainWindow.show();
    }
  }
}


app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', (event) => {
  if (!hasConfirmedClose) {
    event.preventDefault();
    confirmClose();
  }
});

ipcMain.on('save-excel', async (event, excelBuffer) => {
  const { filePath } = await dialog.showSaveDialog({
      defaultPath: 'output.xlsx',
      filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
  });

  if (filePath) {
      fs.writeFileSync(filePath, excelBuffer);
      event.reply('excel-saved', filePath);
  }
});

ipcMain.on('open-save-dialog', async (event, args) => {
  const { filePaths } = await dialog.showOpenDialog({
      title: 'Save Split PDF',
      defaultPath: args.defaultPath,
      properties: ['openDirectory']
  });
  event.sender.send('selected-directory', filePaths[0]);
});