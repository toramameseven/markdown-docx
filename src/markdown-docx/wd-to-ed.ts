import { MessageType, ShowMessage } from "./common";

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
type WdCommand = typeof wdCommand[keyof typeof wdCommand];

let showMessage: ShowMessage | undefined;

const _sp = "\t";

let excelLines: string[] = [""];

//   convert functions

function createLine(command: string, picture: string, text: string) {
  excelLines.push(`${command}\t${picture}\t${text}`);
}

function createLineBlank() {
  excelLines.push(`\t\t`);
}

function createBlock() {
  //
}

function flushCommand() {
  if (lineCommand !== "" || textBuffer !== "") {
    if (!!getPreviousCommand(1) && lineCommand === "") {
      createLineBlank();
    }
    excelLines.push(`${lineCommand}\t\t${textBuffer}`);
  }
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
  flushCommand();
  const newLineInfo = params[0];
  if (newLineInfo === "convertParagraph") {
    createLineBlank();
  }
}

function convertSection(params: string[]) {
  flushCommand();
  const sections = parseInt(params[0], 10);
  const sectionTitle = params[1];
  createLine("#".repeat(sections), "", sectionTitle);
  createLineBlank();
}

function convertCode(params: string[]) {
  if (getPreviousCommand(1) !== wdCommand.code) {
    createLineBlank();
  }
  flushCommand();
  const content = params[0];
  createLine(wdCommand.code, "", content);
}

function convertImage(params: string[]) {
  flushCommand();

  createLineBlank();
  const imagePath = params[0];
  const imageTitle = params[1];
  const hover = params[2];
  createLine("", imagePath, "");
}

function convertLink(params: string[]) {
  createLineBlank();
  const linkPath = params[0];
  const hover = params[1];
  const linkTitle = params[2];

  convertText([`[${linkTitle}](${linkPath})`]);
}

function convertHr(params: string[]) {
  flushCommand();
  createLineBlank();
  createLine("", "", "---");
}

let lineCommand = "";
let textBuffer = "";

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

  return excelLines.map((i) => i).join("\n");
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
      convertHr(params);
      break;
    case wdCommand.tablecontents:
      convertHr(params);
      break;
    case wdCommand.tablecontentslist:
      convertHr(params);
      break;
    default:
      const r = params[0];
      showMessage?.(
        MessageType.warn,
        `NO COMMAND!! [${command}][${r}]`,
        "wd-to-ed.ts",
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

class Cell {
  blockList: string[] = [];
  align: string;
  constructor(align: string) {
    this.align = align;
  }
}

class Table {
  rowCount: number;
  columnCount: number;
  cells: Cell[][];

  constructor(rows: number, columns: number) {
    this.rowCount = rows;
    this.columnCount = columns;
    this.cells = [];
  }

  addCell(i: number, j: number, text: string) {
    this.cells[i][j] = new Cell(text);
  }

  appendTextToCell(i: number, j: number, text: string) {
    this.cells[i][j].blockList.push(text);
  }

  value(i: number, j: number) {
    try {
      return this.cells[i][j];
    } catch {}
    return undefined;
  }

  createWordDownTable() {
    const commands: string[] = [];
    const commandContents: string[] = [];
    commands.push(`tableCreate\t${this.rowCount}\t${this.columnCount}`);

    // contents
    // delete html comments
    for (let i = 0; i < this.rowCount; i++) {
      for (let j = 0; j < this.columnCount; j++) {
        const blockList = this.cells[i][j].blockList;
        //const cellValue = blockList.join("\n");
        commandContents.push(
          `tablecontents\t${i}\t${j}\tnext\t${this.cells[i][j].align}`
        );
        const tableCellCommands = blockList.map(
          (i) => `tablecontentslist\t${i.trim().replace(/<!--.*?-->/g, "")}`
        );
        // detect end paragraph without newline
        tableCellCommands.push(`tablecontentslist\tendParagraph\t\ttm`);
        commandContents.push(...tableCellCommands);
      }
    }

    const r = [...commands, ...commandContents];
    return r;
  }

  initialize() {
    this.rowCount = 0;
    this.columnCount = 0;
    this.cells = [];
  }
}
