const { ipcRenderer } = require('electron');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');

// Specify the path to pdf.worker.js here
PDFJS.workerSrc = '../JS/pdf.worker.js';

document.getElementById('convert-btn').addEventListener('click', async () => {
    const pdfFile = document.getElementById('pdf-file').files[0];
    if (!pdfFile) {
        alert('Please select a PDF file.');
        return;
    }

    const progressBar = document.getElementById('progress');
    const progressMessage = document.getElementById('conversion-msg');
    progressBar.style.width = '0%';

    try {
        progressMessage.innerText = 'Converting...';
        const pdfBuffer = await pdfFile.arrayBuffer();
        const textContent = await pdfParse(pdfBuffer);
        const text = textContent.text;

        const excelWorkbook = new ExcelJS.Workbook();
        const excelSheet = excelWorkbook.addWorksheet('Sheet 1');

        const lines = text.split('\n');
        lines.forEach((line, index) => {
            const cell = excelSheet.getCell(index + 1, 1); // Assuming one column
            cell.value = line;
        });

        const excelBuffer = await excelWorkbook.xlsx.writeBuffer();
        ipcRenderer.send('save-excel', excelBuffer);
        progressMessage.innerText = 'Conversion successful!';
    } catch (error) {
        console.error('Conversion error:', error);
        progressMessage.innerText = 'Error occurred during conversion.';
    }
});
