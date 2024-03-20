const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const ExcelJS = require('exceljs');
const xlsx = require('xlsx');

const apiKey = '39fbc901-2c3f-40ec-bee0-6096b60d75c6';

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


// main function, this is where i put the file and where we get the output
async function main() {
    const inputPDF = 'test2.pdf'; // PDF Sample
    const splitPaths = await splitPDF(inputPDF);

    const excelUrls = [];
    for (const splitPath of splitPaths) {
        const excelUrl = await convertToExcel(splitPath);
        if (excelUrl) {
            excelUrls.push(excelUrl);
        }
        fs.unlinkSync(splitPath);
    }

    const downloadedExcelPaths = [];
    for (let i = 0; i < excelUrls.length; i++) {
        const outputPath = path.join(__dirname, `output_${i + 1}.xlsx`);
        await downloadFile(excelUrls[i], outputPath);
        downloadedExcelPaths.push(outputPath);
    }

    const mergedExcel = path.join(__dirname, 'output.xlsx'); // This is the Output
    await mergeExcelFiles(downloadedExcelPaths, mergedExcel);
    console.log('Merged Excel file:', mergedExcel);
}

// call main function
main().catch(error => console.error('Error:', error));
