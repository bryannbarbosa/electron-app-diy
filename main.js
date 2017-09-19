const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;
const ipcMain = electron.ipcMain;

const dialog = electron.dialog;
const fs = require('fs');

let mainWindow = null;


app.on('ready', () => {
    mainWindow = new BrowserWindow({
        width: 400,
        height: 400,

    });
    mainWindow.loadURL(`file://${__dirname}/app/index.html`);
    dialog.showOpenDialog({filters: [
        {name: 'Arquivos do Excel', extensions: ['xlsx']}
      ]},(fileNames) => {
        if(fileNames === undefined) {
            console.log('No files executed');
            return;
        }
        console.log(fileNames[0]);
    });
});


