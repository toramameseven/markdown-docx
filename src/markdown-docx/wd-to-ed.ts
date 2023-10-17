import { MessageType, ShowMessage } from "./common";
import { WdCommand, wdCommand } from "./wd0-to-wd";

class Cell {
  blockList: string[] = [];
  align: string;
  constructor(align: string) {
    this.align = align;
  }
}

class Code {
  hasCode = false;
  codes: string[];

  constructor() {
    this.codes = [];
  }

  addCode(command: string, params: string[]) {
    this.hasCode = true;
    this.codes.push(params.join(""));
  }

  initialize() {
    this.codes = [];
    this.hasCode = false;
  }

  createCode() {
    createLineBlank("createCode in");
    for (let r = 0; r < this.codes.length; r++) {
      const element = this.codes[r];
      const command = r === 0 || r === this.codes.length - 1 ? "code" : "";
      excelLines.push(`${command}\t\t${element}`);
    }
    if (this.codes.length === 1) {
      excelLines.push(`${wdCommand.code}\t\t`);
    }
    createLineBlank("createCode out");
    this.initialize();
  }
}

class Table {
  rowCount: number;
  columnCount: number;
  cells: Cell[][];
  row: number = 0;
  column: number = 0;
  columnWidth: number[] = [];

  constructor(rows: number, columns: number) {
    this.rowCount = rows;
    this.columnCount = columns;

    this.cells = [];
    let row: Cell[] = [];

    for (let c = 0; c < this.columnCount; c++) {
      this.columnWidth.push(0);
    }

    for (let r = 0; r < this.rowCount; r++) {
      for (let c = 0; c < this.columnCount; c++) {
        row.push(new Cell("left"));
      }
      this.cells.push(row);
      row = [];
    }
  }

  addCell(i: number, j: number, text: string) {
    this.row = i;
    this.column = j;
  }

  appendTextToCell(text: string) {
    this.cells[this.row][this.column].blockList.push(text);
    const len = this.cells[this.row][this.column].blockList.join("").length;
    this.columnWidth[this.column] =
      this.columnWidth[this.column] < len ? len : this.columnWidth[this.column];
  }

  value(i: number, j: number) {
    try {
      return this.cells[i][j];
    } catch {}
    return undefined;
  }

  createWordDownTable() {
    if (this.rowCount === 0) {
      return;
    }

    createLineBlank("createWordDownTable In");
    let row: string[] = [];
    for (let r = 0; r < this.rowCount; r++) {
      for (let c = 0; c < this.columnCount; c++) {
        row.push(this.cells[r][c].blockList.join(""));
      }
      const tableRow = row.join("\t");
      const tableCode = r === 0 || r === this.rowCount - 1 ? "table" : "";
      excelLines.push(`${tableCode}\t\t${tableRow}`);
      row = [];
    }
    createLineBlank("createWordDownTable Out");
  }

  initialize() {
    this.rowCount = 0;
    this.columnCount = 0;
    this.cells = [];
  }
}

function multibyteCharCount(str: string) {
  let len = 0;
  for (let i = 0; i < str.length; i++) {
    str[i].match(/[ -~]/) ? (len += 1) : (len += 2);
  }
  return len;
}


let showMessage: ShowMessage | undefined;

const _sp = "\t";

let excelLines: string[] = [""];

//   convert functions

function createLine(command: string, picture: string, text: string) {
  excelLines.push(`${command}\t${picture}\t${text}`);
}

function createLineBlank(info: string, isForce = false) {
  if (
    (excelLines.length && excelLines[excelLines.length - 1].trim() !== "") ||
    isForce
  ) {
    excelLines.push(`\t\t`);
  }
}

function createBlock() {
  //
}

function flushCommand(command: WdCommand = wdCommand.non, params = [""]) {
  if (
    mdCode?.hasCode &&
    !(
      command === "code" ||
      (command === "newLine" && params[0] === "convertCode")
    )
  ) {
    mdCode.createCode();
  }

  if (params[0] === 'convertHeading End' ){
    excelLines.push(textBuffer);
    createLineBlank("convertSection out");
    textBuffer = "";
  }

  if (lineCommand !== "" || textBuffer !== "") {
    if (!!getPreviousCommand(1) && lineCommand === "") {
      createLineBlank("flushCommand exist previous, no line command");
    }
    excelLines.push(`${lineCommand}\t\t${textBuffer}`);
  }


  mdTable.createWordDownTable();

  mdTable = new Table(0, 0);
  lineCommand = "";
  textBuffer = "";
}

function getPreviousCommand(n: number) {
  const line = excelLines[excelLines.length - n] ?? "";
  const r = line.split("\t")[0];
  return r;
}

function convertText(params: string[]) {
  textBuffer += params[0];
}

function convertNormalList(params: string[]) {
  flushCommand();
  const listLevel = parseInt(params[0], 10);
  lineCommand = "*".repeat(listLevel);
}

function convertOderList(params: string[]) {
  flushCommand();
  const listLevel = parseInt(params[0], 10);
  lineCommand = ".".repeat(listLevel);
}

function convertNewline(params: string[]) {
  flushCommand(wdCommand.newLine, params);
  const newLineInfo = params[0];
  if (newLineInfo === "convertParagraph") {
    createLineBlank("convertNewline convertParagraph");
  }
}

function convertSection(params: string[]) {
  flushCommand();
  const sections = parseInt(params[0], 10);
  const sectionTitle = params[1];
  createLineBlank("convertSection in");
  //createLine("#".repeat(sections), "", sectionTitle);
  textBuffer = "#".repeat(sections) + "\t\t";
  // createLineBlank("convertSection out");
}

function convertCode(params: string[]) {
  flushCommand(wdCommand.code, params);
  mdCode.addCode(wdCommand.code, params);
}

function convertImage(params: string[]) {
  flushCommand();

  createLineBlank("convertImage");
  const imagePath = params[0];
  const imageTitle = params[1];
  const hover = params[2];
  createLine("", imagePath, "");
}

function convertLink(params: string[]) {
  createLineBlank("convertLink");
  const linkPath = params[0];
  const hover = params[1];
  const linkTitle = params[2];

  convertText([`[${linkTitle}](${linkPath})`]);
}

function convertHr(params: string[]) {
  flushCommand();
  createLineBlank("convertHr");
  createLine("", "", "---");
}

function convertTableCreate(params: string[]) {
  flushCommand();
  createLineBlank("convertTableCreate");

  const rows = parseInt(params[0]);
  const columns = parseInt(params[1]);

  mdTable = new Table(rows, columns);
}

function convertTableContents(params: string[]) {
  const row = parseInt(params[0]);
  const column = parseInt(params[1]);

  mdTable.addCell(row, column, "left");
}

function convertTablecontentsList(params: string[]) {
  const command = params[0];
  const cellText = params[1];

  if (command === "text") {
    mdTable.appendTextToCell(cellText);
  }
}

let lineCommand = "";
let textBuffer = "";
let mdTable = new Table(0, 0);
let mdCode: Code = new Code();

export function wdToEd(wd: string, sm?: ShowMessage): string {
  showMessage = sm;
  excelLines = [];
  textBuffer = "";
  lineCommand = "";

  // convert main
  // wd0command p1 v1 p2 v2 ...
  const lines = wd.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const words = lines[i].split(_sp);
    const command = words[0];
    const params = words.slice(1);
    // convert
    const toCommand = command as WdCommand;
    resolveCommand(toCommand, params);
  }
  flushCommand();

  return excelLines.map((i) => i.replace(/^codeX/, "")).join("\n");
}

function resolveCommand(command: WdCommand, params: string[]) {
  switch (command) {
    case wdCommand.text:
      convertText(params);
      break;
    case wdCommand.NormalList:
      convertNormalList(params);
      break;
    case wdCommand.OderList:
      convertOderList(params);
      break;
    case wdCommand.newLine:
      convertNewline(params);
      break;
    case wdCommand.section:
      convertSection(params);
      break;
    case wdCommand.code:
      convertCode(params);
      break;
    case wdCommand.image:
      convertImage(params);
      break;
    case wdCommand.link:
      convertLink(params);
      break;
    case wdCommand.hr:
      convertHr(params);
      break;
    case wdCommand.tableCreate:
      convertTableCreate(params);
      break;
    case wdCommand.tablecontents:
      convertTableContents(params);
      break;
    case wdCommand.tablecontentslist:
      convertTablecontentsList(params);
      break;
    case wdCommand.export:
      break;
    default:
      const r = params[0];
      showMessage?.(
        MessageType.warn,
        `NO COMMAND!! [${command}][${r}]`,
        "wd-to-ed",
        false
      );
  }
}

// function createMdTable({ rawRows }) {
//   for (const row of rawRows) {
//     if (!row) {
//       continue;
//     }
//     //console.log(row);
//   }

//   /* Get Row Data */
//   const rows: string[][] = [];
//   let colWidth: number[] = [];
//   for (const row of rawRows) {
//     if (!row) {
//       continue;
//     }

//     /* Track columns */
//     const cols = row.split(" |").map((c, i) => {
//       c = c.trim();
//       const { value } = getRows(c);
//       if (colWidth.length < i + 1 || colWidth[i] < value.length) {
//         // if width 0, set 1
//         colWidth[i] = value.length + 1;
//       }

//       return c;
//     });

//     rows.push(cols);
//   }

//   if (rows.length < 1) {
//     //return PostProcessResult.RemoveNode;
//   }

//   /* Compose Table */
//   const maxCols = colWidth.length;

//   let res = "";
//   const caption = nodeMetadata.get(node)!.tableMeta!.caption;
//   if (caption) {
//     res += caption + "\n";
//   }

//   // rows info <!-- rows -->
//   let columnInfoPrevious: number[] = [];
//   for (let index = 0; index < maxCols; index++) {
//     columnInfoPrevious[index] = 1;
//   }

//   function getRows(cellInfo: string) {
//     const testMatch = cellInfo.match(/^<!--\s(?<rows>\d+)\s-->/i);
//     const rowsCount = testMatch?.groups?.rows ?? "1";
//     const value = cellInfo.replace(/^<!--\s\d+\s-->/i, "");
//     const rows = parseInt(rowsCount);
//     return { rows, value };
//   }

//   rows.forEach((cols, rowNumber) => {
//     const columnInfo: number[] = [];
//     res += "| ";

//     /* Add Columns */
//     let nodeColumn = 0;
//     for (let i = 0; i < maxCols; i++) {
//       let c = "";
//       if (columnInfoPrevious[i] > 1) {
//         columnInfo[i] = columnInfoPrevious[i] - 1;
//         c = `<!-- ${columnInfo[i]} -->`;
//       } else {
//         c = cols[nodeColumn] ?? "";
//         const { rows, value } = getRows(c);
//         columnInfo[i] = rows;
//         c = value;
//         nodeColumn++;
//       }
//       c += " ".repeat(Math.max(0, colWidth[i] - c.length)); // Pad to max length
//       res += c + " |" + (i < maxCols - 1 ? " " : "");
//     }
//     //
//     columnInfoPrevious = columnInfo;
//     res += "\n";

//     // Add separator row
//     if (rowNumber === 0) {
//       res +=
//         "|" +
//         colWidth
//           .map((w) => {
//             let ww = w;
//             if (ww === 0) {
//               ww = 1;
//             }
//             const r = " " + "-".repeat(ww) + " |";
//             return r;
//           })
//           .join("") +
//         "\n";
//     }
//   });

//   return res;
// }
