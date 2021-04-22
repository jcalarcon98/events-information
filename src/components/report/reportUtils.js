const docx = require("docx");
const fs = require("fs");
const process = require("process");
const util = require("util");

const {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  Media,
  AlignmentType,
  VerticalAlign,
  TextRun,
  PageBreak,
} = docx;

/**
 * Generate a paragraph, it will be used as initial title after table.
 * @param  {string} text - The text that will be displayed
 * @param  {number} fontSize - The size of the text
 * @param  {number} alignment - Text alignment, if is 0 aligment will be LEFT else CENTER.
 * @returns {Paragraph} Paragraph object, it is part of the .docx library.
 */
function generateTitle(content, fontSize, alignment) {
  const currentAlignment =
    alignment === 0 ? AlignmentType.LEFT : AlignmentType.CENTER;

  const paragraph = new Paragraph({
    children: [
      new TextRun({
        text: content,
        bold: true,
        size: fontSize,
      }),
    ],
    alignment: currentAlignment,
  });

  return paragraph;
}

/**
 * Generates random name for the current .docx report.
 * @param  {string} degree - The name of the degree.
 * @param  {string} stage - MITAD DE CICLO || FINAL DE CICLO.
 * @param  {string} initDate - the period init date.
 */
function getRandomDocumentName(type, eventId) {
  const currentDate = new Date();
  const randomNumber = currentDate.getMilliseconds() + currentDate.getDate();
  const currentStringDate = currentDate.toDateString()
  const basePath = `${process.cwd()}/reports`;

  const reportPath = `${basePath}/${type}`;

  if (!fs.existsSync(reportPath)) {
    fs.mkdirSync(reportPath);
  }

  const documentName = `report-${eventId}-${currentStringDate}-${randomNumber}.docx`;

  return {
    documentName,
    folder: type,
  };
}


async function generateDocument(document, folder, documentName) {
  const buffer = await Packer.toBuffer(document);
  const pathToSave = `reports/${folder}/${documentName}`;

  fs.writeFileSync(pathToSave, buffer);
  return pathToSave;
}


async function generateEventReport({ evento: event }) {
  const {
    organizadores: organizers,
    lugares: places,
    valor: values,
    ponentes: speakers,
    evidencias: evidences,
  } = event;

  console.log(util.inspect(organizers, false, null, true), "Organizers");
  console.log(util.inspect(places, false, null, true), "Places");
  console.log(util.inspect(values, false, null, true), "Values");
  console.log(util.inspect(speakers, false, null, true), "Speakers");
  console.log(util.inspect(evidences, false, null, true), "Evidences");

  const document = new Document();

  const documentElements = [];

  const documentTitle = `INFORME DEL EVENTO "${event.nombre}"`;
  const paragraph = generateTitle(documentTitle, 30, 1);

  documentElements.push(paragraph);

  document.addSection({
    children: documentElements
  })

  const documentInformation = getRandomDocumentName('evento', event.id)

  documentPath = await generateDocument(document, 'evento', documentInformation.documentName)

  return documentInformation;
}



async function generateEventsReport(eventsData) {
  console.log(util.inspect(eventsData, false, null, true));
}

module.exports = {
  generateEventReport,
  generateEventsReport,
};
