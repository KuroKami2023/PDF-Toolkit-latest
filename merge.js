const fs = require('fs');
const { PDFDocument } = require('pdf-lib');

async function mergePdfPages(pdfPaths, outputPath) {
    const mergedPdf = await PDFDocument.create();

    for (const pdfPath of pdfPaths) {
        const pdfBytes = await fs.promises.readFile(pdfPath);
        const pdfDoc = await PDFDocument.load(pdfBytes);
        const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
        copiedPages.forEach((page) => {
            mergedPdf.addPage(page);
        });
    }

    const mergedPdfBytes = await mergedPdf.save();
    await fs.promises.writeFile(outputPath, mergedPdfBytes);

    console.log('PDF pages merged successfully!');
}

// Example usage:
const pdfPaths = ['../input_folder/page_1.pdf', '../input_folder/page_2.pdf', '../input_folder/page_3.pdf']; // Array of paths to the PDF files to merge
const outputPath = 'merged_pdf.pdf'; // Output path for the merged PDF
mergePdfPages(pdfPaths, outputPath);
