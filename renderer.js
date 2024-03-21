const { ipcRenderer } = require('electron');

// Listen for click event on the convert button
document.getElementById('convertBtn').addEventListener('click', async () => {
    const fileInput = document.getElementById('pdfInput');
    const file = fileInput.files[0]; // Get the selected file
    const outputDir = "./"; // Provide the default output directory
    if (file) {
      // Send the selected file path and output directory to the main process to start conversion
      ipcRenderer.send('start-conversion', { filePath: file.path, outputDir }); // Pass outputDir
    }
  });
  

  ipcRenderer.on('conversion-complete', async (event, convertedFilePath) => {
    // Prompt the user to choose the download location for the converted file
    const response = await ipcRenderer.invoke('save-excel', convertedFilePath); // Pass convertedFilePath
  
    if (response && response.filePath) {
      const { filePath } = response;
      // Move the converted file to the chosen location
      fs.rename(convertedFilePath, filePath, (err) => {
        if (err) {
          console.error('Error moving file:', err);
        } else {
          console.log('File moved successfully');
        }
      });
    } else {
      console.log('Save dialog canceled or encountered an error');
    }
});
