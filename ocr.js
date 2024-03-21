const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');

const apiKey = 'b7f47eca-5b1d-442e-9398-8fa8f8544ee7';

// Split pdf
async function splitPDF(inputPath) {
    const pdfDoc = await PDFDocument.load(fs.readFileSync(inputPath));
    const totalPages = pdfDoc.getPageCount();
    const splitPaths = [];

    for (let i = 0; i < totalPages; i++) {
        const newPdf = await PDFDocument.create();
        const [copiedPage] = await newPdf.copyPages(pdfDoc, [i]);
        newPdf.addPage(copiedPage);

        const outputPath = path.join(__dirname, `page_${i + 1}.pdf`);
        fs.writeFileSync(outputPath, await newPdf.save());
        splitPaths.push(outputPath);
    }

    return splitPaths;
}

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
        console.log('Error converting PDF to Excel:', error);
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

// Merge Excel Files
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

// Main Function
async function main(inputPDFPath, outputDir) {
  const splitPaths = await splitPDF(inputPDFPath);
  const excelUrls = [];

  for (const splitPath of splitPaths) {
    const excelUrl = await convertToExcel(splitPath);
    if (excelUrl) {
      excelUrls.push(excelUrl);
    }
    fs.unlinkSync(splitPath); // Delete intermediate PDF files
  }

  const downloadedExcelPaths = [];
  for (let i = 0; i < excelUrls.length; i++) {
    const outputPath = path.join(outputDir, `output_${i + 1}.xlsx`);
    await downloadFile(excelUrls[i], outputPath);
    downloadedExcelPaths.push(outputPath);
  }

  const mergedExcel = path.join(outputDir, 'output.xlsx');
 await mergeExcelFiles(downloadedExcelPaths, mergedExcel);

  // Delete intermediate Excel files
  for (const excelPath of downloadedExcelPaths) {
    fs.unlinkSync(excelPath);
  }

  return mergedExcel; // Return the output path of the merged Excel file
}

// Export all functions
module.exports = {
    splitPDF,
    convertToExcel,
    downloadFile,
    mergeExcelFiles,
    main
};
