const docx = require("docx");
const fs = require("fs");
const process = require("process");
const util = require("util");
const { downloadImage } = require("../images/imageUtils");

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
} = docx;

/**
 * Generate a paragraph, it will be used as initial title after table.
 * @param  {string} content - The text that will be displayed
 * @param  {number} fontSize - The size of the text
 * @param  {number} alignment - Text alignment, if is 0 aligment will be LEFT else CENTER.
 * @returns {Paragraph} Paragraph object, it is part of the .docx library.
 */
function generateText(
  content,
  fontSize,
  alignment = AlignmentType.LEFT,
  bold = true
) {
  const paragraph = new Paragraph({
    children: [
      new TextRun({
        text: content,
        bold: bold,
        size: fontSize,
      }),
    ],
    alignment,
  });

  return paragraph;
}

/**
 * Generate automatically widths of each cell, this approach is compatible with all formats
 * (Google Docs, Libre Office and Microsoft word)
 * @param  {number} syllabusesLenght amount of syllabuses.
 * @param  {number} alternativesLength amoount of alternatives.
 * @returns {number[]} array with all widths values
 */
function generateAutomaticallyWidths(firstRowDivider) {
  const columnWidths = [];
  let originalWidth = 9600;
  const leftRow = originalWidth / firstRowDivider;
  const rightRow = originalWidth - leftRow;
  columnWidths.push(leftRow, rightRow);
  return columnWidths;
}

function generateTableRow(rowElements) {
  const generatedChildren = [];

  rowElements.forEach((element) => {
    const textCell = generateText(
      element.content,
      element.fontSize,
      element.alignment,
      element.bold
    );

    const currentTableCell = new TableCell({
      children: [textCell],
      verticalAlign: VerticalAlign.CENTER,
    });

    generatedChildren.push(currentTableCell);
  });

  const tableRow = new TableRow({
    children: generatedChildren,
  });

  return tableRow;
}

function generateTable(tableRows, tableFirstColumnDivider) {
  const columnWidths = generateAutomaticallyWidths(tableFirstColumnDivider);

  const table = new Table({
    rows: tableRows,
    width: 0,
    columnWidths,
  });

  return table;
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
  const currentStringDate = currentDate.toDateString();
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

function customizeInfo(content, isBold) {
  return {
    content: content,
    fontSize: 22,
    alignment: AlignmentType.LEFT,
    bold: isBold,
  };
}

function getFirstTableElements(event) {
  const firstTableElements = [
    [
      customizeInfo("Fecha de inicio de inscripciones:", true),
      customizeInfo(event.fechaInscripcionInicio, false),
    ],
    [
      customizeInfo("Fecha de finalización de inscripciones:", true),
      customizeInfo(event.fechaInscripcionFin, false),
    ],
    [
      customizeInfo("Fecha de inicio del Evento:", true),
      customizeInfo(event.fechaEventoInicio, false),
    ],
    [
      customizeInfo("Fecha de finalización del Evento:", true),
      customizeInfo(event.fechaEventoFin, false),
    ],
    [
      customizeInfo("Tipo del Evento:", true),
      customizeInfo(event.tipoEvento, false),
    ],
    [
      customizeInfo("Categoría del Evento:", true),
      customizeInfo(event.categoriaEvento, false),
    ],
    [customizeInfo("Cupo:", true), customizeInfo(event.cupo, false)],
  ];

  return firstTableElements;
}

function getSimpleElements(simpleContent, simpleArray) {
  const simpleTableElements = [];

  simpleArray.forEach((element) => {
    const currentElement = [
      customizeInfo(simpleContent, true),
      customizeInfo(element, false),
    ];

    simpleTableElements.push(currentElement);
  });
  return simpleTableElements;
}

function emptyLine() {
  return generateText("", 20);
}

function getTableRows(elementsArray) {
  const tableRows = [];

  elementsArray.forEach((row) => {
    const currentTableRow = generateTableRow(row);
    tableRows.push(currentTableRow);
  });

  return tableRows;
}

function getSimpleContent(titleContent, defaultCellContent, arrayContent) {
  const title = titleContent;
  const paragraph = generateText(title, 25, AlignmentType.LEFT, true);
  const simpleTableElements = getSimpleElements(
    defaultCellContent,
    arrayContent
  );
  const simpleTablesRows = getTableRows(simpleTableElements);
  const simpleTable = generateTable(simpleTablesRows, 3);

  return {
    paragraph,
    simpleTable,
  };
}

function getValuesElements(values) {
  const valuesElements = [
    [customizeInfo("ROL", true), customizeInfo("VALOR", true)],
  ];

  values.forEach((currentPairValue) => {
    currentPair = [
      customizeInfo(currentPairValue.rol, false),
      customizeInfo(currentPairValue.valor, false),
    ];

    valuesElements.push(currentPair);
  });

  return valuesElements;
}

function getSpeakerElements(speaker) {
  const speakersElements = [];

  const speakerName = [
    customizeInfo("Nombres", true),
    customizeInfo(speaker.nombre, false),
  ];

  const speakerLastName = [
    customizeInfo("Apellidos", true),
    customizeInfo(speaker.apellido, false),
  ];

  const speakerEmail = [
    customizeInfo("Correo", true),
    customizeInfo(speaker.correo, false),
  ];

  const speakerSummary = [
    customizeInfo("Resumen", true),
    customizeInfo(speaker.resumen, false),
  ];

  speakersElements.push(
    speakerName,
    speakerLastName,
    speakerEmail,
    speakerSummary
  );

  return speakersElements;
}

async function generateEventReport({ evento: event }) {
  const {
    organizadores: organizers,
    lugares: places,
    valor: values,
    ponentes: speakers,
    evidencias: evidences,
  } = event;

  const document = new Document();

  const documentElements = [];

  const documentTitle = `INFORME DEL EVENTO "${event.nombre}"`;
  const paragraph = generateText(documentTitle, 30, AlignmentType.CENTER, true);

  documentElements.push(paragraph, emptyLine());

  const documentDescriptionTitle = generateText(
    "Descripción: ",
    22,
    AlignmentType.LEFT,
    true
  );

  const documentDescriptionContent = generateText(
    event.descripcion,
    22,
    AlignmentType.JUSTIFIED,
    false
  );

  documentElements.push(documentDescriptionTitle, documentDescriptionContent, emptyLine());

  const firstTableElements = getFirstTableElements(event);
  const firstTableRows = getTableRows(firstTableElements);
  const firstTable = generateTable(firstTableRows, 3);

  documentElements.push(firstTable, emptyLine());

  const {
    paragraph: paragrahOrganizers,
    simpleTable: organizersTable,
  } = getSimpleContent(
    "Organizador/es:",
    "Nombre del Organizador:",
    organizers
  );
  
  documentElements.push(paragrahOrganizers, emptyLine(), organizersTable, emptyLine());

  const {
    paragraph: paragraphPlaces,
    simpleTable: placesTable,
  } = getSimpleContent("Lugar/es:", "Nombre del Lugar:", places);

  documentElements.push(paragraphPlaces, emptyLine(), placesTable, emptyLine());

  const valuesTitle = `Valor/es`;
  const paragraphValues = generateText(
    valuesTitle,
    25,
    AlignmentType.LEFT,
    true
  );
  
  documentElements.push(paragraphValues, emptyLine());

  const valuesElements = getValuesElements(values);
  const valuesRows = getTableRows(valuesElements);
  const valuesTable = generateTable(valuesRows, 3);

  documentElements.push(valuesTable, emptyLine());

  const speakersTitle = "Ponente/s";
  const paragraphSpeaker = generateText(
    speakersTitle,
    25,
    AlignmentType.LEFT,
    true
  );

  documentElements.push(paragraphSpeaker, emptyLine());

  speakers.forEach(speaker => {
    const speakersElements = getSpeakerElements(speaker);
    const speakerRows = getTableRows(speakersElements);
    const currentSpeakerTable = generateTable(speakerRows, 4);

    documentElements.push(currentSpeakerTable, emptyLine());
  });

  const eventImagePath = await downloadImage(event.imagen, 'evento');
  const eventImageBuffer = fs.readFileSync(eventImagePath);
  const eventImage = Media.addImage(document, eventImageBuffer);

  const eventImageParagraph = new Paragraph({
    children: [eventImage],
    alignment: AlignmentType.CENTER,
  });

  documentElements.push(eventImageParagraph);

  document.addSection({
    children: documentElements,
  });

  const documentInformation = getRandomDocumentName("evento", event.id);

  documentPath = await generateDocument(
    document,
    "evento",
    documentInformation.documentName
  );

  return documentInformation;
}

async function generateEventsReport(eventsData) {
  console.log(util.inspect(eventsData, false, null, true));
}

module.exports = {
  generateEventReport,
  generateEventsReport,
};
