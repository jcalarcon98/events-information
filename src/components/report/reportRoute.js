const express = require('express');

const app = express();
const reportController = require('./reportController.js');

app.post('/', reportController.generateReport);

app.get('/download/:folder/:documentName', reportController.downloadReport);

module.exports = app;
