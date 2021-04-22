const express = require('express');
const bodyParser = require('body-parser');

const app = express();
const reportRoutes = require('./src/components/report/reportRoute');

app.use(bodyParser.json());

app.use('/api/docx', reportRoutes);

app.listen(3000, () => {
  console.log(`We are running node in port: 3000`);
});
