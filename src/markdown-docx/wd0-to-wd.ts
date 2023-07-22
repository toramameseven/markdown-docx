import { isInteger } from "lodash";
import { MessageType, ShowMessage, getWordDownMergeCommand } from "./common";

let showMessage: ShowMessage | undefined;

type DocxParam = { [x: string]: string };
const _sp = "\t";

let wordDownLines: string[] = [""];

type ListInfo = {
  ordered: boolean;
  start: number;
  indent: number;
};

let listInfos: ListInfo[] = [
  {
    start: 0,
    ordered: false,
    indent: 0,
  },
];

const wd0Command = {
  // marked
  title: "title",
  subTitle: "subTitle",
  heading: "heading",
  paragraph: "paragraph",
  list: "list",
  listitem: "listitem",
  code: "code",
  blockquote: "blockquote",
  table: "table",
  tablePos: "tablePos",
  tablerow: "tablerow",
  tablecell: "tablecell",
  text: "text",
  image: "image",
  link: "link",
  html: "html",

  non: "non",

  // word down
  param: "param",
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
  newLine: "newLine",
  newPage: "newPage",
  htmlWdCommand: "htmlWdCommand",
  hr: "hr",
  // table
  cols: "cols",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",
} as const;
type Wd0Command = typeof wd0Command[keyof typeof wd0Command];

interface BaseBlock {
  blockType: Wd0Command;
  blockList: string[];
}

class Base implements BaseBlock {
  blockType: Wd0Command;
  blockList: string[] = [];
  constructor(blockType: Wd0Command) {
    this.blockType = blockType;
  }
}
let blockInfos: BaseBlock[] = [new Base(wd0Command.non)];
// --------------------------------

class Cell implements BaseBlock {
  blockType: Wd0Command = wd0Command.tablecell;
  blockList: string[] = [];
  x: number = 0;
  y: number = 0;
  align: string;
  mergeTo: number[] = [];
  constructor(align: string) {
    this.align = align;
  }
}

class Table implements BaseBlock {
  blockType: Wd0Command;
  blockList: string[];
  rowCount: number;
  columnCount: number;
  row: Cell[];
  rows: Cell[][];
  tableWidthInfo: string;
  rowMerge: string;
  emptyMerge: boolean;
  tablePos: string;

  constructor() {
    this.blockType = wd0Command.table;
    this.blockList = [];
    this.rowCount = 0;
    this.columnCount = 0;
    this.rows = [];
    this.row = [];
    this.tableWidthInfo ='';
    //  "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1";
    this.rowMerge = "";
    this.tablePos = "";
    this.emptyMerge = false;
  }

  addCell(cell: Cell) {
    this.columnCount++;
    cell.x = this.rowCount;
    cell.y = this.columnCount;
    this.row.push(cell);
  }
  appendTextToLastCell(text: string) {
    const cell = this.row.pop();
    if (cell) {
      cell.blockList.push(text);
      this.row.push(cell);
    }
  }
  newRow() {
    this.rowCount++;
    this.columnCount = 0;
    this.row = [];
  }
  endRow() {
    this.rows.push([...this.row]);
    this.row = [];
  }

  value(i: number, j: number) {
    try {
      return this.rows[i][j];
    } catch {}
    return undefined;
  }

  createWordDownTable() {
    const commands: string[] = [];
    const commandContents: string[] = [];
    const commandMergeInfos: string[] = [];
    commands.push(`tableCreate\t${this.rowCount}\t${this.columnCount}`);

    // <!-- word cols 1,4  -->
    // tableWidthInfo = 1,4

    let outTableInfo = this.tableWidthInfo ? this.tableWidthInfo : this.getColumnSize();

    commands.push(`tableWidthInfo\t${outTableInfo}`);

    if (this.tablePos){
      commands.push(`tablePos\t${this.tablePos}`);
    }

    // merge info rows
    // <!-- word rowMerge 3-4,5-6,7-9 -->
    // rowMerge = 3-4,5-6,7-9
    // ${this.rowCount}-${this.rowCount} is for all rows
    this.rowMerge = !!this.rowMerge.length
      ? `${this.rowMerge},${this.rowCount}-${this.rowCount}`
      : `${this.rowCount}-${this.rowCount}`;

    const rowMergeList = this.rowMerge.split(",");
    let rowCellMerge0 = 0;

    // get merge info in the cells.
    for (let r = 0; r < this.rowCount; r++) {
      let lastEmptyColumn = -1;
      for (let c = 0; c < this.columnCount; c++) {
        let mergeColumnTo = -1;
        const mergeData =
          this.rows[r][c].blockList.length > 0
            ? this.rows[r][c].blockList[0]
            : "";
        lastEmptyColumn = mergeData ? -1 : c;
        const mergeCommand = getWordDownMergeCommand(mergeData);

        if (mergeCommand?.isMergeColumn) {
          mergeColumnTo = lastEmptyColumn - 1;
        }

        let mergeRowTo = -1;
        if (mergeCommand?.isMergeRow) {
          for (let r2 = r - 1; r2 > -1; r2--) {
            const mergeData =
              this.rows[r2][c].blockList.length > 0
                ? this.rows[r2][c].blockList[0]
                : "";
            if (!!mergeData) {
              mergeRowTo = r2;
            }
          }
        }
        this.rows[r][c].mergeTo = [mergeRowTo, mergeColumnTo];
        //console.log(`${r},${c} - ${mergeRowTo},${mergeColumnTo}`);
      }
    }

    // merge main
    for (let k = 0; k < rowMergeList.length; k++) {
      // rows[0]: start row, rows[1]:end row
      const rows = rowMergeList[k].split("-");
      // merge empty cell
      // if not emptyMerge, no columns loop
      // do merge in the range outside rowMerge.
      const columnLoopNum = this.emptyMerge ? this.columnCount : -1;
      for (let j = 0; j < columnLoopNum; j++) {
        // do emptyMerge(row direction)
        // ~~0 means zero base array.
        const beforeEnd0 = parseInt(rows[0]) - 1;
        let cellStartRow0 = rowCellMerge0;
        let cellEndRow0 = beforeEnd0;
        // rows
        for (let i = rowCellMerge0 + 1; i <= beforeEnd0; i++) {
          if (this.rows[i][j].blockList.length === 0) {
            // empty value so continue
          } else {
            // not empty value
            cellEndRow0 = i - 1;
            if (cellStartRow0 < cellEndRow0) {
              const mergeData =
                this.rows[cellStartRow0][j].blockList.length > 0
                  ? this.rows[cellStartRow0][j].blockList[0]
                  : "empty";
              commandMergeInfos.push(
                //`tableMarge\t${cellStartRow0}\t${j}\t${cellEndRow0}\t${j}\t${this.rows[cellStartRow0][j].blockList[0]}`
                `tableMarge\t${cellStartRow0}\t${j}\t${cellEndRow0}\t${j}\t${mergeData}`
              );
            }
            cellStartRow0 = i;
            cellEndRow0 = beforeEnd0;
          }
        }
        if (cellStartRow0 < cellEndRow0) {
          const mergeData =
            this.rows[cellStartRow0][j].blockList.length > 0
              ? this.rows[cellStartRow0][j].blockList[0]
              : "empty";
          commandMergeInfos.push(
            `tableMarge\t${cellStartRow0}\t${j}\t${cellEndRow0}\t${j}\t${mergeData}`
          );
        }
      } // for loop column

      // merge rows
      for (let j = 0; j < this.columnCount; j++) {
        const start = parseInt(rows[0]) - 1;
        const end = parseInt(rows[1]) - 1;
        if (start < end) {
          commandMergeInfos.push(
            `tableMarge\t${start}\t${j}\t${end}\t${j}\t${this.rows[start][
              j
            ].blockList.join("")}`
          );
        }
      }
      rowCellMerge0 = parseInt(rows[1]);
    }

    // contents
    // delete html comments
    for (let i = 0; i < this.rowCount; i++) {
      for (let j = 0; j < this.columnCount; j++) {
        const blockList = this.rows[i][j].blockList;
        //const cellValue = blockList.join("\n");
        commandContents.push(
          `tablecontents\t${i}\t${j}\tnext\t${this.rows[i][j].align}`
        );
        const tableCellCommands = blockList.map(
          (i) => `tablecontentslist\t${i.trim().replace(/<!--.*?-->/g, "")}`
          //(i) => `tablecontentslist\t${i}`
        );
        // detect end paragraph without newline
        tableCellCommands.push(`tablecontentslist\tendParagraph\t\ttm`);
        commandContents.push(...tableCellCommands);
      }
    }

    const ret = [...commands, ...commandContents, ...commandMergeInfos];
    return ret;
  }

  getColumnSize(){
    let columnSize = new Array<number[]>(this.columnCount);
    for(let  c = 0; c < this.columnCount; c++) {
      columnSize[c] = [];
    }

    for (let i = 0; i < this.rowCount; i++) {
      for (let j = 0; j < this.columnCount; j++) {
        let blockList = this.rows[i][j].blockList;
        const blockStringLen = blockList.map(
          (v) => count(v.trim().replace(/<!--.*?-->/g, ""))
        );
        columnSize[j] = [...columnSize[j], ...blockStringLen];
      }
    }

    let columnSizeS: string[] = [];
    for (let i = 0; i < this.columnCount; i++) {
      columnSizeS.push([averageLength(columnSize[i])].toString());
    }

    return columnSizeS.join(",");

    function count(s: string) {
      let len = 0;
      for (let i = 0; i < s.length; i++) {
        if (s[i].match(/[ -~]/)) {
          len += 0.5;
        }
        else {
          len += 1;
        }
      }
      return len;
    }

    function averageLength(dataset: number[]){

      if (dataset.length === 0){
        dataset = [3];
      }

      const sum = dataset.reduce((a, b) => { 
        return a + b;
       });
      
      let average = sum / dataset.length;

      if (average < 3){
        average = 3;
      }

      return average;

      // const deviation = dataset.map((a) => {
      //   const subtract = a - average; /*平均との差 */
      //   return subtract ** 2;
      //  });

      // const deviationSum = deviation.reduce((a, b) => {
      //   return a + b;
      //  });


      // return deviationSum / (dataset.length);
    }
  }



  initialize() {
    this.blockList = [];
    this.rowCount = 0;
    this.columnCount = 0;
    this.rows = [];
    this.row = [];
    this.tableWidthInfo = '';
    //  "1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1";
    this.rowMerge = "";
    this.emptyMerge = false;
  }
}

const table = new Table();

function popBlockInfo() {
  const popped = blockInfos.pop();
  return popped;
}

function pushBlockInfo(i: BaseBlock) {
  blockInfos.push(i);
}

function getBlockInfoTypeLast() {
  if (blockInfos.length === 0) {
    return wd0Command.non;
  }
  return blockInfos.slice(-1)[0].blockType;
}

//   convert functions

function addNewLine(info: string) {
  const r = [wd0Command.newLine, info, "tm"].join(_sp);
  outputWd(r);
}

function convertTitle(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  // main title
  const r = [wd0Command.title, params.title].join(_sp);
  outputWd(r);
  addNewLine("convertTitle");
}

function convertSubTitle(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = [wd0Command.subTitle, params.subTitle].join(_sp);
  outputWd(r);
  addNewLine("convertSubTitle");
}

function convertHeading(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    const r = popBlockInfo();
    r!.blockList.map((i) => {
      outputWd(i);
    });
    addNewLine("convertHeading End");
    return;
  }
  // const r = ["section", params.index, params.title, params.idTitle].join(_sp);
  const r = ["section", params.index, params.idTitle].join(_sp);
  outputWd(r);
  pushBlockInfo(new Base(wd0Command.heading));
  // addNewLine("convertHeading");
}

function convertParagraph(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    const r = popBlockInfo();
    r!.blockList.map((i) => {
      outputWd(i);
    });
    addNewLine("convertParagraph");
    return;
  }

  pushBlockInfo(new Base(wd0Command.paragraph));
}

function convertBlockquote(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    const r = popBlockInfo();
    r!.blockList.map((i) => {
      outputWd(i);
    });
    return;
  }

  const block = new Base(wd0Command.blockquote);
  pushBlockInfo(block);
}

function convertTable(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    popBlockInfo();
    const r = table.createWordDownTable();
    r.map((l) => {
      outputWd(l);
    });
    table.initialize();
    return;
  }
  pushBlockInfo(table);
}

function convertTableRow(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    table.endRow();
    popBlockInfo();
    return;
  }
  table.newRow();
  pushBlockInfo(new Base(wd0Command.tablerow));
}

function convertTableCell(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    const cell = blockInfos!.pop();
    table.addCell(cell as Cell);
    return;
  }

  const cell = new Cell(params.align ?? "left");
  pushBlockInfo(cell);
  if (params.content.trim()) {
    convertText(params.content, false);
  }
}

function convertList(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    listInfos!.pop();
    const r = popBlockInfo();
    r!.blockList.map((i) => {
      outputWd(i);
    });
    // when end list, not newline
    outputWd(wd0Command.indentMinus);
    return;
  }

  // if inside list, does not need new line.
  if (getBlockInfoTypeLast() === wd0Command.listitem) {
    addNewLine("convertList In");
  }

  //list ordered 0 orq start number body
  const r: ListInfo = {
    start: parseInt(params.start),
    ordered: params.ordered === "1",
    indent: listInfos![listInfos!.length - 1].indent + 1,
  };
  outputWd(wd0Command.indentPlus);
  listInfos!.push(r);
  pushBlockInfo(new Base(wd0Command.list));
}

function convertListItem(params: DocxParam, isCommandEnd?: boolean) {
  //task 0 checked 0 text
  // task, checked, text
  if (!listInfos) {
    return;
  }

  if (isCommandEnd) {
    const r = popBlockInfo();
    r!.blockList.map((i) => {
      outputWd(i);
    });
    addNewLine("convertListItem End");
    return;
  }

  // if inside list, does not need new line.
  if (getBlockInfoTypeLast() === wd0Command.listitem) {
    addNewLine("convertListItem In");
  }
  params.checked = "";
  const listType = listInfos[listInfos.length - 1].ordered
    ? "OderList"
    : "NormalList";

  const r = [
    listType,
    listInfos[listInfos.length - 1].indent.toString(),
    params.text,
  ].join(_sp);

  outputWd(r);
  pushBlockInfo(new Base(wd0Command.listitem));
}

function convertCode(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    const code = popBlockInfo();
    code!.blockList.forEach((i) => {
      const codeParam = i.split(_sp);
      const r = ["code", codeParam[1]].join(_sp);
      outputWd(r);
      addNewLine("convertCode");
    });
    return;
  }
  pushBlockInfo(new Base(wd0Command.code));
}

function convertHr(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["hr", params.href].join(_sp);
  outputWd(r);
}

function convertImage(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  //href images/main_window2.png, text caption sss, title image title
  const r = ["image", params.href, params.text, params.title, "tm"].join(_sp);
  outputWd(r);
}

function convertText(params: DocxParam | string, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }

  // get text
  let plainText = "";
  if (typeof params === "string") {
    plainText = params;
  } else if (typeof params === "object") {
    plainText = (params as DocxParam).text;
  }
  const wordDownText = ["text", plainText].join(_sp);
  outputWd(wordDownText);
}

function convertHtml(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }

  if (getBlockInfoTypeLast() === wd0Command.tablecell) {
    blockInfos.slice(-1)[0].blockList.push(["text", params.text].join(_sp));
    return;
  }

  // delete html comments
  const trimmed = params.text.replace(/<!--.*?-->/g, "").trim();
  if (trimmed === "") {
    return;
  }

  let r = "";
  if (params.text.match(/<br>/i)) {
    r = [wd0Command.newLine, "convertHtml <BR>", "tm"].join(_sp);
  } else {
    r = ["text", params.text].join(_sp);
  }
  outputWd(r);
}

function convertLink(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["link", params.href, params.title, params.text, "tm"].join(_sp);
  outputWd(r);
}
function convertNewPage(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["newPage", params.href, params.title, params.text, "tm"].join(_sp);
  outputWd(r);
}
function convertPageSetup(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = [
    "pageSetup",
    params.orientation,
    params.pagesize,
    "",
    "",
    "tm",
  ].join(_sp);
  outputWd(r);
}

function convertPatchParameter(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["param", params.key, params.value, "", "", "tm"].join(_sp);
  outputWd(r);
}

function convertEngine(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["docxEngine", params.docxEngine, "", "", "", "tm"].join(_sp);
  outputWd(r);
}
function convertTemplate(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["docxTemplate", params.docxTemplate, "", "", "", "tm"].join(_sp);
  outputWd(r);
}

function convertAuthor(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["author", params.author, "", "", "", "tm"].join(_sp);
  outputWd(r);
}

function convertDivision(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["division", params.division, "", "", "", "tm"].join(_sp);
  outputWd(r);
}

function convertDocNumber(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["docNumber", params.docNumber, "", "", "", "tm"].join(_sp);
  outputWd(r);
}

function convertProperty(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = [
    "property",
    params.propertyKey,
    params.propertyValue,
    "",
    "",
    "tm",
  ].join(_sp);
  outputWd(r);
}

function convertCrossRef(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["crossRef", params.crossRef, "", "", "", "tm"].join(_sp);
  outputWd(r);
}
function convertClearContent(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["clearContent", params.isClearContent, "", "", "", "tm"].join(_sp);
  outputWd(r);
}
function convertDate(params: DocxParam, isCommandEnd?: boolean) {
  if (isCommandEnd) {
    return;
  }
  const r = ["date", params.date, "", "", "", "tm"].join(_sp);
  outputWd(r);
}

function outputWd(wdText: string) {
  // if block exists, add all wdTexts to the block.
  if (getBlockInfoTypeLast() !== wd0Command.non) {
    blockInfos.slice(-1)[0].blockList.push(wdText);
    return;
  }
  // do not duplicate new lines.
  if (wdText.split(_sp)[0] === wd0Command.newLine) {
    if (
      wordDownLines.length > 0 &&
      wordDownLines.slice(-1)[0].split(_sp)[0] === wd0Command.newLine
    ) {
      return;
    }

    let line = "\t";
    for (let i = wordDownLines.length - 1; i > 0; i--) {
      line = wordDownLines[i];
      if (
        ![
          wd0Command.indentMinus as string,
          wd0Command.indentPlus as string,
        ].includes(line)
      ) {
        break;
      }
    }

    if (line.split(_sp)[0] === wd0Command.newLine) {
      return;
    }
  }

  if (isAddWordSeparator(wordDownLines, wdText)) {
    // add space between words. but is not needed for japanese.
    wordDownLines.push(["text", " ", "", "", "", "tm"].join(_sp));
  }
  wordDownLines.push(wdText);
}

function isAddWordSeparator(wdLines: string[], wdLine: string) {
  const r =
    wdLines.length > 0 &&
    wdLines.slice(-1)[0].split("\t")[0] === "text" &&
    wdLine.split("\t")[0] === "text";
  return r;
}

export function wd0ToDocx(wd0: string, sm?: ShowMessage): string {
  showMessage = sm;
  wordDownLines = [];
  blockInfos = [new Base(wd0Command.non)];

  // now not use front matter
  // option from front matter.
  // convertDocxEngine({
  //   docxEngine: optionsFromFrontmatter.docxEngine ?? "",
  //   docxTemplate: optionsFromFrontmatter.docxTemplate ?? "",
  // });

  // convert main
  // wd0command p1 v1 p2 v2 ...
  const lines = wd0.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const words = lines[i].split(_sp);
    const rawCommand = words.shift();
    const isCommandEnd = rawCommand?.[0] === "/";
    const command = rawCommand?.slice(isCommandEnd ? 1 : 0);

    // check params
    const params: DocxParam = {};
    const count = words.length;
    for (let j = 0; j < count; j += 2) {
      params[words[j]] = words[j + 1];
    }
    // convert
    const toCommand = command as Wd0Command;
    resolveCommand(toCommand, params, isCommandEnd);
  }

  return wordDownLines.map((i) => i).join("\n");
}

function resolveCommand(
  command: Wd0Command,
  params: DocxParam,
  isCommandEnd?: boolean
) {
  switch (command) {
    case wd0Command.title:
      convertTitle(params, isCommandEnd);
      break;
    case wd0Command.subTitle:
      convertSubTitle(params, isCommandEnd);
      break;
    case wd0Command.heading:
      convertHeading(params, isCommandEnd);
      break;
    case wd0Command.paragraph:
      convertParagraph(params, isCommandEnd);
      break;
    case wd0Command.list:
      convertList(params, isCommandEnd);
      break;
    case wd0Command.listitem:
      convertListItem(params, isCommandEnd);
      break;
    case wd0Command.text:
      convertText(params, isCommandEnd);
      break;
    case wd0Command.html:
      convertHtml(params, isCommandEnd);
      break;
    case wd0Command.link:
      convertLink(params, isCommandEnd);
      break;
    case wd0Command.image:
      convertImage(params, isCommandEnd);
      break;
    case wd0Command.table:
      convertTable(params, isCommandEnd);
      break;
    case wd0Command.blockquote:
      convertBlockquote(params, isCommandEnd);
      break;
    case wd0Command.tablerow:
      convertTableRow(params, isCommandEnd);
      break;
    case wd0Command.tablecell:
      convertTableCell(params, isCommandEnd);
      break;
    case wd0Command.code:
      convertCode(params, isCommandEnd);
      break;
    case wd0Command.cols:
      if (isCommandEnd) {
        return;
      }
      table.tableWidthInfo = params.cols;
      break;
    case wd0Command.tablePos:
      if (isCommandEnd) {
        return;
      }
      table.tablePos = params.tablePos;
      break;
    case wd0Command.rowMerge:
      if (isCommandEnd) {
        return;
      }
      table.rowMerge = params.rowMerge;
      break;
    case wd0Command.emptyMerge:
      if (isCommandEnd) {
        return;
      }
      table.emptyMerge = true;
      break;
    case wd0Command.newPage:
      if (isCommandEnd) {
        return;
      }
      convertNewPage(params, isCommandEnd);
      break;
    case wd0Command.pageSetup:
      if (isCommandEnd) {
        return;
      }
      convertPageSetup(params, isCommandEnd);
      break;
    case wd0Command.docxEngine:
      if (isCommandEnd) {
        return;
      }
      convertEngine(params, isCommandEnd);
      break;
    case wd0Command.param:
      if (isCommandEnd) {
        return;
      }
      convertPatchParameter(params, isCommandEnd);
      break;
    case wd0Command.author:
      if (isCommandEnd) {
        return;
      }
      convertAuthor(params, isCommandEnd);
      break;
    case wd0Command.division:
      if (isCommandEnd) {
        return;
      }
      convertDivision(params, isCommandEnd);
      break;
    case wd0Command.docNumber:
      if (isCommandEnd) {
        return;
      }
      convertDocNumber(params, isCommandEnd);
      break;
    case wd0Command.property:
      if (isCommandEnd) {
        return;
      }
      convertProperty(params, isCommandEnd);
      break;
    case wd0Command.crossRef:
      if (isCommandEnd) {
        return;
      }
      convertCrossRef(params, isCommandEnd);
      break;
    case wd0Command.clearContent:
      if (isCommandEnd) {
        return;
      }
      convertClearContent(params, isCommandEnd);
      break;
    case wd0Command.date:
      if (isCommandEnd) {
        return;
      }
      convertDate(params, isCommandEnd);
      break;
    case wd0Command.docxTemplate:
      if (isCommandEnd) {
        return;
      }
      convertTemplate(params, isCommandEnd);
      break;
    case wd0Command.toc:
      if (isCommandEnd) {
        return;
      }
      const tocCommand = `toc\t${params.tocTo}\t${params.tocCaption}\ttm`;
      outputWd(tocCommand);
      addNewLine("convertToc");
      break;
    case wd0Command.hr:
      convertHr(params, isCommandEnd);
      break;
    case wd0Command.newLine:
      if (!isCommandEnd) {
        addNewLine("wd0NewLine");
      }
      break;
    default:
      const r = params[0] ?? "";
      showMessage?.(
        MessageType.warn,
        `NO COMMAND!! [${command}][${r}]`,
        "wd0-to-wd",
        false
      );
  }
}
