const fs = require("fs");
const path = require('path')  
const process = require("process");
const axios = require('axios')

async function downloadImage (imageUrl, fileName) {  
  const url = imageUrl;
  const pathToSave = `${process.cwd()}/${fileName}.jpg`;
  const writer = fs.createWriteStream(pathToSave);

  const response = await axios({
    url,
    method: 'GET',
    responseType: 'stream'
  })

  response.data.pipe(writer)

  return new Promise((resolve, reject) => {
    writer.on('close',  () => {
        resolve(pathToSave);
    })
    writer.on('error', err => {
        error = err;
        writer.close();
        reject(err);
    })
  })
}

module.exports = {
  downloadImage
}