const express = require('express');

const app = express();
const reportController = require('./reportController.js');

app.post('/event', reportController.generateEventReport);

app.post('/events', reportController.generateEventsReport);

app.get('/download/:folder/:documentName', reportController.downloadReport);

module.exports = app;
