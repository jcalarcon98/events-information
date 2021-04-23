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
    ]
  ];

  if (event.tipoEvento) {
    firstTableElements.push(
      [
        customizeInfo("Tipo del Evento:", true),
        customizeInfo(event.tipoEvento, false),
      ]
    );
  } 

  if (event.categoriaEvento) {
    firstTableElements.push(
      [
        customizeInfo("Categoría del Evento:", true),
        customizeInfo(event.categoriaEvento, false),
      ]
    );
  } 

  firstTableElements.push(
    [
      customizeInfo("Cupo:", true), 
      customizeInfo(event.cupo, false)
    ],
  )
  

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

function getEventDescriptionSection(description) {
  const documentDescriptionTitle = generateText("Descripción: ", 22, AlignmentType.LEFT, true);
  const documentDescriptionContent = generateText(description, 22, AlignmentType.JUSTIFIED, false);
  return { documentDescriptionTitle, documentDescriptionContent };
}

function getEventDescriptionTable(event) {
  const firstTableElements = getFirstTableElements(event);
  const firstTableRows = getTableRows(firstTableElements);
  const firstTable = generateTable(firstTableRows, 3);
  return firstTable;
}

function getValuesSection(values) {
  const paragraphValues = generateText('Valor/es', 25, AlignmentType.LEFT, true);
  const valuesElements = getValuesElements(values);
  const valuesRows = getTableRows(valuesElements);
  const valuesTable = generateTable(valuesRows, 3);

  return  {paragraphValues, valuesTable}
}

function getSpeakerTable(speaker) {
  const speakersElements = getSpeakerElements(speaker);
  const speakerRows = getTableRows(speakersElements);
  const currentSpeakerTable = generateTable(speakerRows, 4);

  return currentSpeakerTable;
}

async function getImageSection(titleSection, document, image) {
  const paragraphImage = generateText(titleSection, 25, AlignmentType.LEFT, true);
  const eventImage = await generateImage(document,image, 'evento');
  return [ paragraphImage, eventImage ]
}

function getActivitiesElements(activity) {

  /**
   *  "nombre": "Programación con Python",
      "tipoActividad": "Técnica",
      "categoriaActividad": "Categoria Actividad",
      "organizador": "Jean Carlos Alarcón",
      "ponente": "Juan Francisco Tenorio",
      "lugar": "Universidad Nacional de Loja",
      "fechaActividad": "24 de Mayo de 2020"
   */

  const activitiesElements = [];

  const activityName = [
    customizeInfo("Nombre:", true),
    customizeInfo(activity.nombre, false),
  ];

  const activityType = [
    customizeInfo("Tipo de Actividad:", true),
    customizeInfo(activity.tipoActividad, false),
  ];

  const activityCategory = [
    customizeInfo("Categoría de Actividad:", true),
    customizeInfo(activity.categoriaActividad, false),
  ];

  const activityOrganizer = [
    customizeInfo("Organizador:", true),
    customizeInfo(activity.organizador, false),
  ];

  const activitySpeaker = [
    customizeInfo("Ponente:", true),
    customizeInfo(activity.ponente, false),
  ];
  
  const activityPlace = [
    customizeInfo("Lugar:", true),
    customizeInfo(activity.lugar, false),
  ];

  const activityDate = [
    customizeInfo("Fecha de desarrollo de la Actividad:", true),
    customizeInfo(activity.fechaActividad, false),
  ];

  activitiesElements.push(
    activityName,
    activityType,
    activityCategory,
    activityOrganizer,
    activitySpeaker,
    activityPlace,
    activityDate
  );

  return activitiesElements;
}

function getActivityTable(activity) {
  const activityElements = getActivitiesElements(activity);
  const activityRows = getTableRows(activityElements);
  const currentActivityTable = generateTable(activityRows, 3);

  return currentActivityTable;
}

async function generateImage(document, imageUrl, imageName) {
  const imagePath = await downloadImage(imageUrl, imageName);
  const imageBuffer = fs.readFileSync(imagePath);
  const image = Media.addImage(document, imageBuffer, 530, 330);

  const imageParagraph = new Paragraph({
    children: [image],
    alignment: AlignmentType.CENTER,
  });

  return imageParagraph;
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

  const paragraph = generateText(`INFORME DEL EVENTO "${event.nombre}"`, 30, AlignmentType.CENTER, true);
  documentElements.push(paragraph, emptyLine());

  const {documentDescriptionTitle, documentDescriptionContent} = getEventDescriptionSection(event.descripcion);
  documentElements.push(documentDescriptionTitle, documentDescriptionContent, emptyLine());

  const eventTableDescription = getEventDescriptionTable(event);
  documentElements.push(eventTableDescription, emptyLine());

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

  const { paragraphValues, valuesTable } = getValuesSection(values);
  documentElements.push(paragraphValues, emptyLine(), valuesTable, emptyLine());

  const paragraphSpeaker = generateText('Ponente/s', 25, AlignmentType.LEFT, true);
  documentElements.push(paragraphSpeaker, emptyLine());

  speakers.forEach(speaker => {
    documentElements.push(getSpeakerTable(speaker), emptyLine());
  });

  const [paragraphImage, eventImage]  = await getImageSection('Imagen del evento:', document, event.imagen);
  documentElements.push(paragraphImage, emptyLine(), eventImage, emptyLine());

  const [paragraphImageEventSchedule, eventScheduleImage]  = await getImageSection('Imagen del cronograma del evento:', document, event.cronograma);
  documentElements.push(paragraphImageEventSchedule, emptyLine(), eventScheduleImage, emptyLine());

  const paragraphEvidencesTitle = generateText('Evidencias del evento desarrollado', 25, AlignmentType.LEFT, true);
  documentElements.push(paragraphEvidencesTitle, emptyLine());
  
  for (const [index, evidence] of evidences.entries()) {
    const currentImageEvidence = await generateImage(document, evidence, `evidencia_${index + 1}`);
    documentElements.push(currentImageEvidence, emptyLine());
  }

  document.addSection({
    children: documentElements,
  });

  const documentInformation = getRandomDocumentName("evento", event.id);
  documentPath = await generateDocument(document, "evento", documentInformation.documentName);

  return documentInformation;
}

async function generateEventsReport({ eventos:  events}) {

  const {
    valor: values,
    ponentes: speakers,
    actividades: activities,
    evidencias: evidences,
  } = events;

  const document = new Document();

  const documentElements = [];

  const paragraph = generateText(`INFORME DEL EVENTO "${events.nombre}"`, 30, AlignmentType.CENTER, true);
  documentElements.push(paragraph, emptyLine());

  const {documentDescriptionTitle, documentDescriptionContent} = getEventDescriptionSection(events.descripcion);
  documentElements.push(documentDescriptionTitle, documentDescriptionContent, emptyLine());

  const eventTableDescription = getEventDescriptionTable(events);
  documentElements.push(eventTableDescription, emptyLine());

  const { paragraphValues, valuesTable} = getValuesSection(values);
  documentElements.push(paragraphValues, emptyLine(), valuesTable, emptyLine());

  const speakerParagraph = generateText('Ponente/s:', 25, AlignmentType.LEFT, true);
  documentElements.push(speakerParagraph, emptyLine());

  speakers.forEach(speaker => {
    documentElements.push(getSpeakerTable(speaker), emptyLine());
  });

  const activitiesParagraph = generateText('Actividad/es:', 25, AlignmentType.LEFT, true);
  documentElements.push(activitiesParagraph, emptyLine());

  activities.forEach(activity => {
    documentElements.push(getActivityTable(activity), emptyLine());
  });

  const [paragraphImage, eventImage]  = await getImageSection('Imagen del evento:', document, events.imagen);
  documentElements.push(paragraphImage, emptyLine(), eventImage, emptyLine());

  const [paragraphImageEventSchedule, eventScheduleImage]  = await getImageSection('Imagen del cronograma del evento:', document, events.cronograma);
  documentElements.push(paragraphImageEventSchedule, emptyLine(), eventScheduleImage, emptyLine());

  const paragraphEvidencesTitle = generateText('Evidencias del evento desarrollado', 25, AlignmentType.LEFT, true);
  documentElements.push(paragraphEvidencesTitle, emptyLine());
  
  for (const [index, evidence] of evidences.entries()) {
    const currentImageEvidence = await generateImage(document, evidence, `evidencia_${index + 1}`);
    documentElements.push(currentImageEvidence, emptyLine());
  }

  document.addSection({
    children: documentElements,
  });

  const documentInformation = getRandomDocumentName("eventos", events.id);
  documentPath = await generateDocument(document, "eventos", documentInformation.documentName);

  return documentInformation;
}

module.exports = {
  generateEventReport,
  generateEventsReport,
};
