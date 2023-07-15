import { MessageType, ShowMessage } from "./common";

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
    textileLines.push("<pre>");
    for (let r = 0; r < this.codes.length; r++) {
      const element = this.codes[r];
      textileLines.push(`${element}`);
    }
    textileLines.push("</pre>");
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
      const tableRow = row.join(" | ");
      //const tableCode = r === 0 || r === this.rowCount - 1 ? "table" : "";
      textileLines.push(`| ${tableRow} |`);
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

const wdCommand = {
  text: "text",
  normalList: "NormalList",
  oderList: "OderList",
  newLine: "newLine",
  section: "section",
  code: "code",
  image: "image",
  link: "link",
  hr: "hr",
  tableCreate: "tableCreate",
  tablecontents: "tablecontents",
  tablecontentslist: "tablecontentslist",

  // marked
  title: "title",
  subTitle: "subTitle",
  heading: "heading",
  paragraph: "paragraph",
  list: "list",
  listitem: "listitem",
  blockquote: "blockquote",
  table: "table",
  tablerow: "tablerow",
  tablecell: "tablecell",
  html: "html",

  non: "non",

  // word down
  author: "author",
  date: "date",
  division: "division",
  docxEngine: "docxEngine",
  docxTemplate: "docxTemplate",
  pageSetup: "pageSetup",
  toc: "toc",

  crossRef: "crossRef",
  property: "property",
  clearContent: "clearContent",
  docNumber: "docNumber",
  indentPlus: "indentPlus",
  indentMinus: "indentMinus",
  endParagraph: "endParagraph",
  newPage: "newPage",
  htmlWdCommand: "htmlWdCommand",
  // table
  cols: "cols",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",
} as const;
type WdCommand = (typeof wdCommand)[keyof typeof wdCommand];

let showMessage: ShowMessage | undefined;

const _sp = "\t";

let textileLines: string[] = [""];

//   convert functions

function createLine(command: string, picture: string, text: string) {
  textileLines.push(`${command} ${text}`);
}

function createLineEx(command: string, picture: string, text: string) {
  //textileLines.push(`${command} ${text}`);
  textBuffer += `${command} ${text}`;
}

function createLineBlank(info: string, isForce = false) {
  //console.log(`${createLineBlank.name}==> ${info}`);
  if (
    (textileLines.length &&
      textileLines[textileLines.length - 1].trim() !== "") ||
    isForce
  ) {
    textileLines.push("");
  }
}

function createBlock() {
  //
}

function flushCommand(command: WdCommand = wdCommand.non, params = [""]) {
  
  // code
  if (
    mdCode?.hasCode &&
    !(
      command === "code" ||
      (command === "newLine" && params[0] === "convertCode")
    )
  ) {
    mdCode.createCode();
  }

  if (markupCommand || textBuffer) {
    const previousCommand = getPreviousCommand(1);
    if (previousCommand && markupCommand === "") {
      createLineBlank(`flushCommand exist previous:${previousCommand}, no markup Command`);
    }
    if (markupCommand) {
      textileLines.push(`${markupCommand} ${textBuffer}`);
    } else {
      textileLines.push(textBuffer);
    }
  }

  // table
  mdTable.createWordDownTable();
  mdTable = new Table(0, 0);
  markupCommand = "";
  textBuffer = "";
}

function getPreviousCommand(n: number) {
  const line = textileLines[textileLines.length - n] ?? "";
  const r = line.split("\t")[0];
  return r;
}

function getNextCommand(n: number) {
  const line = textileLines[textileLines.length + n] ?? "";
  const r = line.split("\t")[0];
  return r;
}

function convertText(params: string[]) {
  textBuffer += params[0];
}

function convertNormalList(params: string[]) {
  flushCommand();
  const listLevel = parseInt(params[0], 10);
  markupCommand = "*".repeat(listLevel);
}

function convertOderList(params: string[]) {
  flushCommand();
  const listLevel = parseInt(params[0], 10);
  markupCommand = "#".repeat(listLevel);
}

function convertNewline(params: string[]) {
  flushCommand(wdCommand.newLine, params);
  const newLineInfo = params[0];
  if (newLineInfo === "convertParagraph") {
    //createLineBlank("convertNewline convertParagraph");
  }
}

function convertSection(params: string[]) {
  flushCommand();
  const sections = parseInt(params[0], 10);
  const sectionTitle = params[1];
  createLineBlank("convertSection in");
  createLine(`h${sections}.`, "", sectionTitle);
  createLineBlank("convertSection out");
}

function convertCode(params: string[]) {
  flushCommand(wdCommand.code, params);
  mdCode.addCode(wdCommand.code, params);
}

function convertImage(params: string[]) {
  // for text
  // textBuffer += params[0];
  //flushCommand();

  //createLineBlank("convertImage");
  const imagePath = params[0];
  const imageTitle = params[1];
  const hover = params[2];
  createLineEx(`!${imagePath}!`, imagePath, "");
}

function convertLink(params: string[]) {
  createLineBlank("convertLink");
  const linkPath = params[0];
  const hover = params[1];
  const linkTitle = params[2];

  convertText([` "${linkTitle}":${linkPath} `]);
}

function convertHr(params: string[]) {
  flushCommand();
  createLineBlank("convertHr");
  createLine("---", "", "");
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

let markupCommand = "";
let textBuffer = "";
let mdTable = new Table(0, 0);
let mdCode: Code = new Code();

export function wdToTextile(wd: string, sm?: ShowMessage): string {
  showMessage = sm;
  textileLines = [];
  textBuffer = "";
  markupCommand = "";

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

  return textileLines.map((i) => i.replace(/^codeX/, "")).join("\n");
}

function resolveCommand(command: WdCommand, params: string[]) {
  switch (command) {
    case wdCommand.text:
      convertText(params);
      break;
    case wdCommand.normalList:
      convertNormalList(params);
      break;
    case wdCommand.oderList:
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
    default:
      const r = params[0];
      showMessage?.(
        MessageType.warn,
        `NO COMMAND!! [${command}][${r}]`,
        "wd-to-textile",
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
