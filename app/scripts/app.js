const electron = require('electron');
const ipcRenderer = electron.ipcRenderer;

document.getElementById('startProcess').addEventListener('click', () => {
    ipcRenderer.send('runFile');
    ipcRenderer.on('getFormData', (event, args) => {
        let data = {
            ddi: document.getElementById('ddi').value.toString(),
            ddd: document.getElementById('ddd').value.toString()
        };
        event.sender.send('sendFormData', data);
    });
});


