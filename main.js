const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;
const ipcMain = electron.ipcMain;

const Excel = require('exceljs');

const dialog = electron.dialog;
const fs = require('fs');



let mainWindow = null;


app.on('ready', () => {
    mainWindow = new BrowserWindow({
        width: 400,
        height: 400,

    });
    mainWindow.loadURL(`file://${__dirname}/app/index.html`);
    
});

ipcMain.on('runFile', (event, args) => {
    dialog.showOpenDialog({filters: [
        {name: 'Planilhas do Excel', extensions: ['xlsx']}
      ]},(fileNames) => {
        if(fileNames === undefined) {
            console.log('Nenhum arquivo executado');
            return;
        }
        workbook = new Excel.Workbook();
        workbook.xlsx.readFile(fileNames[0])
        .then(function() {
            let worksheet = workbook.getWorksheet(1);
    
            let arr = [70, 77, 78, 79];

            for(let i = 0; i < 50; i++) {
              arr.push(i);
            }

            event.sender.send('getFormData');

            ipcMain.on('sendFormData', (event, args) => {
              let ddi = args.ddi;
              let ddd = args.ddd;

              for(let i = 1; i <= worksheet.rowCount; i++) {
              let row = worksheet.getRow(i);
              let value = row.getCell(1).value.toString();
              value = value.trim();
              value = value.replace(/\s/g, '');
              value = value.replace(/[`a-zA-Z~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, '');
              row.getCell(1).value = value;
              row.commit();
              let length = row.getCell(1).value.toString().length;
              
              if(length == 8 && arr.indexOf(Number(value.substr(0,2))) > -1) {
               row.getCell(1).value = ddi + ddd + value;
              }
              else if(length == 8 && !arr.indexOf(Number(value.substr(0,2))) > -1) {
                row.getCell(1).value = ddi + ddd + '9' + value;
              }
    
              if(length == 9) {
               row.getCell(1).value = ddi + ddd + value;
              }
    
              if(length == 10 && arr.indexOf(Number(value.substr(2,2))) > -1) {
                row.getCell(1).value = ddi + value;
              }
    
              else if(length == 10 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
                let sub = value.substr(0,2) + '9' + value.substr(2);
                  row.getCell(1).value = ddi + sub.toString();
              }
              if(length == 11) {
                row.getCell(1).value = ddi + row.getCell(1).value.toString();
              }
              if(length == 12 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
                let sub = ddi + value.slice(0, -1).toString();
                row.getCell(1).value = sub;
              }
              row.commit();
            }

            let arr_exclude = [];

            for(let i = 0; i < 50; i++) {
              arr_exclude.push(i);
            }
    
            for(let i = 1; i <= worksheet.rowCount; i++) {
              let row = worksheet.getRow(i);
              let value = row.getCell(1).value.toString();
              let length = row.getCell(1).value.toString().length;
              
              if(length <= 7) {
                worksheet.spliceRows(i, 1);
              }

              if(length == 12 && arr_exclude.indexOf(Number(value.substr(4,2))) > -1) {
                worksheet.spliceRows(i, 1);
              }
            }

            dialog.showSaveDialog({filters: [
                {name: 'Planilhas do Excel', extensions: ['*']}]},(fileName) => {
                if (fileName === undefined) return;
                dialog.showMessageBox({ message: "Planilha filtrada e salva com sucesso!",buttons: ["OK"] });
                return workbook.xlsx.writeFile(fileName+ '_quant' + '_' +  worksheet.rowCount + '.xlsx');
            });
          });
       })
   });
});




// backup




/*

const app = angular.module('excelApp', []);
const fs = require('fs')
const { dialog } = require('electron').remote
const Excel = require('exceljs')
let workbook = new Excel.Workbook()
document.getElementById('openFile').addEventListener('click', () => {
    dialog.showOpenDialog({filters: [
        {name: 'Arquivos do Excel', extensions: ['xlsx']}
      ]},(fileNames) => {
        if(fileNames === undefined) {
            console.log('No files executed')
            return
        }
        workbook.xlsx.readFile(fileNames[0])
        .then(function() {
            let worksheet = workbook.getWorksheet(1)
    
            let arr = [70, 77, 78, 79]

            for(let i = 0; i < 50; i++) {
              arr.push(i);
            }
            let ddi = document.getElementById('ddi').value.toString()
            let ddd = document.getElementById('ddd').value.toString()
            
            for(let i = 1; i <= worksheet.rowCount; i++) {
              let row = worksheet.getRow(i)
              let value = row.getCell(1).value.toString()
              value = value.trim()
              value = value.replace(/\s/g, '')
              value = value.replace(/[`a-zA-Z~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, '')
              row.getCell(1).value = value
              row.commit()
              let length = row.getCell(1).value.toString().length
              
              if(length == 8 && arr.indexOf(Number(value.substr(0,2))) > -1) {
               row.getCell(1).value = ddi + ddd + value
              }
              else if(length == 8 && !arr.indexOf(Number(value.substr(0,2))) > -1) {
                row.getCell(1).value = ddi + ddd + '9' + value
              }
    
              if(length == 9) {
               row.getCell(1).value = ddi + ddd + value
              }
    
              if(length == 10 && arr.indexOf(Number(value.substr(2,2))) > -1) {
                row.getCell(1).value = ddi + value
              }
    
              else if(length == 10 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
                let sub = value.substr(0,2) + '9' + value.substr(2)
                  row.getCell(1).value = ddi + sub.toString()
              }
              if(length == 11) {
                row.getCell(1).value = ddi + row.getCell(1).value.toString();
              }
              if(length == 12 && !arr.indexOf(Number(value.substr(2,2))) > -1) {
                let sub = ddi + value.slice(0, -1).toString()
                row.getCell(1).value = sub
              }
              row.commit()
            }

            let arr_exclude = []

            for(let i = 0; i < 50; i++) {
              arr_exclude.push(i)
            }
    
            for(let i = 1; i <= worksheet.rowCount; i++) {
              let row = worksheet.getRow(i)
              let value = row.getCell(1).value.toString()
              let length = row.getCell(1).value.toString().length
              
              if(length <= 7) {
                worksheet.spliceRows(i, 1)
              }

              if(length == 12 && arr_exclude.indexOf(Number(value.substr(4,2))) > -1) {
                worksheet.spliceRows(i, 1)
              }
            }

            dialog.showSaveDialog({filters: [
                {name: 'Arquivos do Excel', extensions: ['*']}]},(fileName) => {
                if (fileName === undefined) return;
                dialog.showMessageBox({ message: "Planilha filtrada e salva com sucesso!",buttons: ["OK"] })
                return workbook.xlsx.writeFile(fileName+ '_quant' + '_' +  worksheet.rowCount + '.xlsx')
                
            });
            
        })
    })
}, false)

*/
