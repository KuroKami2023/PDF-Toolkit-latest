const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

const file = fs.readFileSync('moogootest.pdf');

const data = new FormData();
data.append('file', file, { filename: 'moogootest.pdf' }); 
data.append('output', 'OUTPUT2');

const config = {
  method: 'post',
  maxContentLength: Infinity, 
  url: 'https://api.pdfrest.com/excel',
  headers: {
    'Api-Key': '39fbc901-2c3f-40ec-bee0-6096b60d75c6',
    ...data.getHeaders(),
  },
  data: data,
};

axios(config)
  .then(function (response) {
    console.log(JSON.stringify(response.data));
  })
  .catch(function (error) {
    console.log(error);
  });
