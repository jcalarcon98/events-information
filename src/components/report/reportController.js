const process = require('process');
const { generateEventReport, generateEventsReport } = require('./reportUtils');

exports.generateEventReport = async (req, res) => {
  const data = req.body;
  const documentPath = await generateEventReport(data);

  res.json(documentPath);
};

exports.generateEventsReport = async(req, res) => {
  const data = req.body;
  const documentPath = await generateEventsReport(data);

  res.json(documentPath)
}


exports.downloadReport = (req, res) => {
  const { folder, documentName } = req.params;

  if (!folder || !documentName) {
    res.status(400).send({
      message: 'Folder and Document name are required',
    });
  }
  const documentPath = `${process.cwd()}/reports/${folder}/${documentName}`;
  res.download(documentPath);
};
