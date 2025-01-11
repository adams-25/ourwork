const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');
const { initializeExcelConnection } = require('./com.js'); // Import the function


let mainWindow;

app.whenReady().then(() => {
    mainWindow = new BrowserWindow({
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    mainWindow.loadFile('index.html'); // Load index.html for the first page


   

    mainWindow.webContents.on('did-finish-load', async () => {
        console.log('Page finished loading:', mainWindow.webContents.getURL());
        try {
          
            await initializeExcelConnection(mainWindow);
        } catch (error) {
            console.error('Error during initializeExcelConnection:', error.message);
        }
    });


    app.on('window-all-closed', () => {
        if (process.platform !== 'darwin') {
            app.quit();
        }
    });
});
