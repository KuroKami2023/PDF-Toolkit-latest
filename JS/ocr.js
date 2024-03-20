const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');

const apiKey = '39fbc901-2c3f-40ec-bee0-6096b60d75c6';

document.getElementById('convertBtn').addEventListener('click', async () => {
    const inputFiles = document.getElementById('pdfInput').files;
    if (inputFiles.length === 0) {
        alert('Please select at least one PDF file to convert.');
        return;
    }

    const inputPaths = [];
    for (let i = 0; i < inputFiles.length; i++) {
        inputPaths.push(inputFiles[i].path); // Assuming you're using Electron or similar to access file paths
    }

    const excelUrls = [];
    for (const inputPath of inputPaths) {
        const excelUrl = await convertToExcel(inputPath);
        if (excelUrl) {
            excelUrls.push(excelUrl);
        }
    }

    const downloadedExcelPaths = [];
    for (let i = 0; i < excelUrls.length; i++) {
        const outputPath = path.join(__dirname, `output_${i + 1}.xlsx`);
        await downloadFile(excelUrls[i], outputPath);
        downloadedExcelPaths.push(outputPath);
    }

    const mergedExcel = path.join(__dirname, 'Output_File.xlsx'); // This is the Output
    await mergeExcelFiles(downloadedExcelPaths, mergedExcel);
    console.log('Merged Excel file:', mergedExcel);
});

// Convert PDF (API)
async function convertToExcel(inputPath) {
    const fileStream = fs.createReadStream(inputPath);
    const data = new FormData();
    data.append('file', fileStream);

    const config = {
        method: 'post',
        maxBodyLength: Infinity,
        url: 'https://api.pdfrest.com/excel',
        headers: {
            'Api-Key': apiKey,
            ...data.getHeaders(),
        },
        data: data,
    };

    try {
        const response = await axios(config);
        return response.data.outputUrl;
    } catch (error) {
        if (error.response) {
            // The request was made and the server responded with a status code
            console.log('Error response status:', error.response.status);
            console.log('Error response data:', error.response.data);
        } else if (error.request) {
            // The request was made but no response was received
            console.log('Error request:', error.request);
        } else {
            // Something happened in setting up the request that triggered an Error
            console.log('Error:', error.message);
        }
        return null;
    }
}


// Download the converted files
async function downloadFile(url, outputPath) {
    const response = await axios({
        url: url,
        method: 'GET',
        responseType: 'stream',
    });

    const outputStream = fs.createWriteStream(outputPath);
    response.data.pipe(outputStream);

    return new Promise((resolve, reject) => {
        outputStream.on('finish', resolve);
        outputStream.on('error', reject);
    });
}

async function mergeExcelFiles(filePaths, outputFilePath) {
    const mergedWorkbook = new ExcelJS.Workbook();

    let worksheetCounter = 1;

    for (const filePath of filePaths) {
        const workbook = xlsx.readFile(filePath);

        workbook.SheetNames.forEach(sheetName => {
            let newSheetName = sheetName;
            let counter = 1;
            while (mergedWorkbook.getWorksheet(newSheetName)) {
                newSheetName = `${sheetName} (${counter})`;
                counter++;
            }

            const worksheet = workbook.Sheets[sheetName];
            const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
            const newWorksheet = mergedWorkbook.addWorksheet(newSheetName);

            for (let i = 0; i < data.length; i++) {
                newWorksheet.addRow(data[i]);
            }
        });
    }

    await mergedWorkbook.xlsx.writeFile(outputFilePath);
}
