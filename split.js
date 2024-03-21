const fs = require('fs-extra');
const { PDFDocument } = require('pdf-lib');
const { ipcRenderer } = require('electron');
const Swal = require('sweetalert2');

async function splitPDF(inputPDFPath, outputFolder, inputFileName) {
    const inputFile = await fs.readFile(inputPDFPath);
    const pdfDoc = await PDFDocument.load(inputFile);
    
    for (let pageNumber = 0; pageNumber < pdfDoc.getPageCount(); pageNumber++) {
        const pdfWriter = await PDFDocument.create();
        const [copiedPage] = await pdfWriter.copyPages(pdfDoc, [pageNumber]);
        pdfWriter.addPage(copiedPage);

        const outputPagePath = `${outputFolder}/${inputFileName}_page_${pageNumber + 1}.pdf`;
        await fs.ensureDir(outputFolder);
        await fs.writeFile(outputPagePath, await pdfWriter.save());
        console.log(`Page ${pageNumber + 1} extracted and saved as ${outputPagePath}`);
    }

    Swal.fire({
        position: "center",
        icon: "success",
        title: "Split PDF Successful",
        showConfirmButton: false,
        timer: 1500
      });
}

document.getElementById('convert-btn').addEventListener('click', async () => {
    const inputPDF = document.getElementById('pdf-file').files[0];

    if (inputPDF) {
        const inputPDFPath = inputPDF.path;

        ipcRenderer.send('open-save-dialog', { defaultPath: 'output_folder' });
    } else {
        Swal.fire({
            position: "center",
            icon: "error",
            title: "No file chosen",
            showConfirmButton: false,
            timer: 1500
          });
    }
});

ipcRenderer.on('selected-directory', async (event, filePath) => {
    const outputFolder = filePath;

    try {
        const inputPDF = document.getElementById('pdf-file').files[0];
        const inputPDFPath = inputPDF.path;

        const inputFileName = inputPDF.name.replace(/\.[^/.]+$/, ''); 
        await splitPDF(inputPDFPath, outputFolder, inputFileName);
        console.log("Splitting completed successfully.");
    } catch (error) {
        console.error("Error:", error);
        Swal.fire({
            icon: 'error',
            title: 'Error',
            text: 'An error occurred while splitting the PDF.',
        });
    }
});
