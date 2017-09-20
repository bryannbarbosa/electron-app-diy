const electron = require('electron');
const ipcRenderer = electron.ipcRenderer;

document.getElementById('startProcess').addEventListener('click', () => {
    ipcRenderer.send('runFile', {start: true});
});


