const enum Grade {
  INCOMING = "Incoming",
  KINDERGARTEN = "Kindergarten",
  FIRST = "1st Grade",
  SECOND = "2nd Grade",
  THIRD = "3rd Grade",
  FOURTH = "4th Grade",
  FIFTH = "5th Grade",
  ALUMNI = "Alumni",
}

type Parent = {
  name: string;
  address: string;
  phone: string;
  email: string;
};

type IntakeRow = {
  program: string;
  grade: Grade;
  inclusion: string;
  studentFirstName: string;
  studentLastName: string;
  parent1Name: string;
  parent1Address: string;
  parent1Phone: string;
  parent1Email: string;
  parent2Name: string;
  parent2Address: string;
  parent2Phone: string;
  parent2Email: string;
};
type IntakeColumn = keyof IntakeRow;

type ExportRow = {
  studentId: string;
  gradeId: number;
  studentFirstName: string;
  studentLastName: string;
  parentFirstName: string;
  parentLastName: string;
  parentAddress: string;
  parentPhone: string;
  parentEmail: string;
  parentLanguage: string;
};
type ExportColumn = keyof ExportRow;

const exportHeaderRow: { [column in ExportColumn]: string } = {
  studentId: "Student ID",
  gradeId: "Grade",
  studentFirstName: "Student First",
  studentLastName: "Student Last",
  parentFirstName: "Parent First",
  parentLastName: "Parent Last",
  parentEmail: "Parent Email",
  parentPhone: "Parent Cell",
  parentAddress: "Parent Address",
  parentLanguage: "Parent Language",
};

function main() {
  // We assume the user has actually selected entire rows on the correct sheet
  // here. If not, terrible things will happen.
  const intakeRows = getSelectedIntakeRows();

  // Convert the data to our export format.
  const exportRows = convertIntakeRows(intakeRows);

  // Prefix with the header row and convert to 2D array.
  const exportValues = exportRowsTo2D([exportHeaderRow, ...exportRows]);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportSheet = ss.insertSheet("File > Download > .csv and delete me ");

  exportSheet
    .getRange(1, 1, exportValues.length, exportValues[0].length)
    .setValues(exportValues);
}

function getSelectedIntakeRows(): IntakeRow[] {
  const values = getSelectedValues();
  return values.map((row) => ({
    program: row[1],
    studentFirstName: row[2],
    studentLastName: row[3],
    grade: row[4],
    inclusion: row[6],
    parent1Name: row[7],
    parent1Address: row[8],
    parent1Phone: row[9],
    parent1Email: row[10],
    parent2Name: row[11],
    parent2Address: row[12],
    parent2Phone: row[13],
    parent2Email: row[14],
  }));
}

function exportRowsTo2D(rows: { [column in ExportColumn]: any }[]): any[][] {
  return rows.map((row) => [
    row.studentId,
    row.gradeId,
    row.studentFirstName,
    row.studentLastName,
    row.parentFirstName,
    row.parentLastName,
    row.parentEmail,
    row.parentPhone,
    row.parentAddress,
    row.parentLanguage,
  ]);
}

function convertIntakeRows(intakeRows: IntakeRow[]): ExportRow[] {
  // Exclude anyone indicating they don't want to be added to ParentSquare.
  const includedRows = intakeRows.filter((row) =>
    row.inclusion.includes("YES")
  );

  // The intake form includes both parents in the same row, but ParentSquare
  // wants to see them in separate rows (with the student duplicated).
  const parent1Rows = includedRows.map((row) => ({
    studentId: "",
    gradeId: getGradeId(row),
    studentFirstName: row.studentFirstName,
    studentLastName: row.studentLastName,
    parentFirstName: getParentFirstName(row.parent1Name),
    parentLastName: getParentLastName(row.parent1Name),
    parentAddress: row.parent1Address,
    parentPhone: row.parent1Phone,
    parentEmail: row.parent1Email,
    parentLanguage: "en",
  }));

  // Some rows may not have a second parent; ensure we filter them out first.
  const parent2Rows = includedRows
    .filter((row) => row.parent2Name.length > 0)
    .map((row) => ({
      studentId: "",
      gradeId: getGradeId(row),
      studentFirstName: row.studentFirstName,
      studentLastName: row.studentLastName,
      parentFirstName: getParentFirstName(row.parent2Name),
      parentLastName: getParentLastName(row.parent2Name),
      parentAddress: row.parent2Address,
      parentPhone: row.parent2Phone,
      parentEmail: row.parent2Email,
      parentLanguage: "en",
    }));

  return [...parent1Rows, ...parent2Rows];
}

const intakeGradeMap: { [grade in Grade]: number } = {
  Incoming: -1,
  Kindergarten: 0,
  "1st Grade": 1,
  "2nd Grade": 2,
  "3rd Grade": 3,
  "4th Grade": 4,
  "5th Grade": 5,
  Alumni: 6,
};

function getGradeId(intakeRow: IntakeRow): number {
  const programModifier = intakeRow.program.includes("JBBP") ? 0 : 8;
  const gradeModifier = intakeGradeMap[intakeRow.grade];

  return programModifier + gradeModifier;
}

function getParentFirstName(name: string) {
  return name.split(" ")[0];
}

function getParentLastName(name: string) {
  return name.split(" ")[1];
}

function getSelectedValues() {
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const selection = activeSheet.getSelection();
  return selection.getActiveRange().getValues();
}

function onOpen() {
  createMenu();
}

function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu("ParentSquare Import")
    .addItem("Convert selected rows", "main")
    .addToUi();
}
