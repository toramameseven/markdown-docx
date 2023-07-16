import {
  DocxOption,
  MessageType,
  ShowMessage,
  selectExistsPath,
  getFileContents
} from "./common";
const pptxgenx = require("pptxgenjs");
import PptxGenJS from "pptxgenjs";
import * as fs from "fs";
import * as Path from "path";
import * as imageSize from "image-size";
const TeXToSVG = require("tex-to-svg");

import {
  Bookmark,
  ExternalHyperlink,
  //HeadingLevel,
  ImageRun,
  //Indent,
  //InternalHyperlink,
  Paragraph,
  ParagraphChild,
  patchDocument,

  //PatchDocumentOptions,
  PatchType,
  Table,
  //TableCell,
  //TableRow,
  //TextDirection,
  TableOfContents,
  TextRun,
  //VerticalAlign,
  WidthType,
  // Document as DocumentDocx,
  //convertInchesToTwip,
  //PageReference,
  SimpleField,
} from "docx";
import { svg2imagePng } from "./svg-png-image";
import { runCommand } from "../tools/tools-common";


export async function wordDownToPptxBody(
  fileWd: string,
  wdBody: string,
  option: DocxOption
) {
  let body = wdBody;

  if (body === ''){
    body = getFileContents(fileWd);
  }

  try {
    await wdToPptxJs(body, "", "outPath", Path.dirname(fileWd), option);
  } catch (e) {
    option.message?.(MessageType.warn, `wdToPptxJs err: ${e}.`, "wd-to-pptx", false);
    return;
  }
  return;
}

/**
 *
 * @param wd
 * @param docxTemplatePath
 * @param docxOutPath
 * @param mdSourcePath
 */
export async function wdToPptxJs(
  wd: string,
  docxTemplatePath: string,
  docxOutPath: string,
  mdSourcePath: string,
  option?: DocxOption
): Promise<void> {
  // initialize pptx
  let pptx: PptxGenJS = new pptxgenx();

  pptx.title = "PptxGenJS Test Suite Presentation";
  pptx.subject = "PptxGenJS Test Suite Export";
  pptx.author = "Brent Ely";
  pptx.revision = "15";

  // FYI: use `headFontFace` and/or `bodyFontFace` to set the default font for the entire presentation (including slide Masters)
  // pptx.theme = { bodyFontFace: "Arial" };
  pptx.layout = "LAYOUT_WIDE";

  // parse lines
  const lines = (wd + "\nEndline").split(/\r?\n/);
  let currentParagraph = new DocParagraph(NodeType.text);
  let paragraphBlock: PptxGenJS.TextProps[] = [];
  let currentSlide = pptx.addSlide();
  let tableJs: TableJs | undefined = undefined;

  const documentInfo = {
    title: "",
    subTitle: "",
    division: "",
    date: "",
    author: "",
    docNumber: "",
  };

  for (let i = 0; i < lines.length; i++) {
    const wdCommandList = lines[i].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      // flush texts before.
      const paragraph = currentParagraph.createPptxParagraph(currentSlide);
      paragraphBlock.push(...paragraph);
      addParagraphBlock(currentSlide, paragraphBlock);
      paragraphBlock = [];

      // initialize table.
      tableJs = new TableJs(
        parseInt(wdCommandList[1]),
        parseInt(wdCommandList[2]),
        mdSourcePath
      );
      continue;
    }

    // table command
    if (wdCommandList[0].includes("table")) {
      tableJs!.doTableCommand(lines[i]);
    } else {
      // in not table command, create table.
      if (tableJs) {
        tableJs!.createTable(currentSlide);
        tableJs = undefined;
      }
    }

    // image command
    if (wdCommandList[0] === "image") {
      // flush texts before.
      const paragraph = currentParagraph.createPptxParagraph(currentSlide);
      paragraphBlock.push(...paragraph);
      addParagraphBlock(currentSlide, paragraphBlock);

      paragraphBlock = [];

      // initialize image
      const image = createImageChild(mdSourcePath, wdCommandList[1]);
      currentSlide.addImage(image);
      continue;
    }

    // html comment command <!-- word xxxx -->
    const isResolveCommand = resolveWordCommentsCommands(
      wdCommandList,
      documentInfo
    );

    if (isResolveCommand) {
      continue;
    }

    // body commands
    currentParagraph = await resolveWordDownCommandEx(
      lines[i],
      currentParagraph,
      mdSourcePath
    );

    // when paragraph end, flush paragraph
    if (currentParagraph.isFlush) {
      const isNewSheet = currentParagraph.isNewSheet;
      const paragraph = currentParagraph.createPptxParagraph(currentSlide);
      paragraphBlock.push(...paragraph);

      // reset paragraph. but keep the indent.
      currentParagraph = new DocParagraph(
        NodeType.text,
        { fontSize: 10 },
        currentParagraph.indent
      );

      if (isNewSheet) {
        currentSlide = pptx.addSlide();
        currentParagraph.isNewSheet = false;
      }
    }
  }
  // end loop lines

  // cerate paragraph block
  addParagraphBlock(currentSlide, paragraphBlock);

  //Export Presentation
  const ff = `PptxGenJS_Demo_${new Date()
    .toISOString()
    .replace(/\D/gi, "")}.pptx`;

  const outPathPPtx = Path.resolve(Path.dirname(mdSourcePath), ff);

  const r = await pptx.writeFile({
    fileName: outPathPPtx,
    compression: true,
  });

  // open ppt
  const pptExe = await selectExistsPath(
    [
      "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
      "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
    ],
    ""
  );

  runCommand(pptExe, r);
}


const _sp = "\t";
//
const NodeType = {
  non: "non",
  section: "section",
  heading: "heading",
  OderList: "OderList",
  NormalList: "NormalList",
  math: "math",
  note: "note",
  warning: "warning",

  // word down
  author: "author",
  date: "date",
  division: "division",
  docxEngine: "docxEngine",
  docxTemplate: "docxTemplate",
  pageSetup: "pageSetup",
  toc: "toc",

  //marked
  title: "title",
  subTitle: "subTitle",
  paragraph: "paragraph",
  list: "list",
  listitem: "listitem",
  code: "code",
  blockquote: "blockquote",
  table: "table",
  tablerow: "tablerow",
  tablecell: "tablecell",
  text: "text",
  image: "image",
  link: "link",
  html: "html",

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
type NodeType = (typeof NodeType)[keyof typeof NodeType];

const DocStyle = {
  "1": "1",
  Body: "body",
  Body1: "body1",
  nList1: "nList1",
  nList2: "nList2",
  nList3: "nList3",
  numList1: "numList1",
  numList2: "numList2",
  numList3: "numList3",
  code: "code",
  note1: "note1",
  warning1: "warn1",
  Error: "Error",
} as const;
type DocStyle = (typeof DocStyle)[keyof typeof DocStyle];

/**
 *
 */
class TableJs {
  rows: number;
  columns: number;
  // x, y, list();;
  cells: PptxGenJS.TextProps[][][];
  row: number;
  column: number;
  mergedCells: number[][][];
  tableWidthArray: number[];
  mdSourcePath: string;

  constructor(rows: number, columns: number, mdSourcePath: string) {
    this.row = 0;
    this.column = 0;
    this.rows = rows;
    this.columns = columns;
    this.mdSourcePath = mdSourcePath;
    this.cells = new Array(rows);
    for (let r = 0; r < rows; r++) {
      this.cells[r] = new Array(columns);
    }
    this.mergedCells = new Array(rows);
    for (let r = 0; r < rows; r++) {
      this.mergedCells[r] = new Array(columns);
    }
    this.tableWidthArray = [];
  }

  doTableCommand(line: string) {
    const words = line.split(_sp);
    switch (words[0]) {
      case "tableWidthInfo":
        const widthArray = words[1].split(",").map((l) => parseInt(l));
        const sums = widthArray.reduce(function (a, x) {
          return a + x;
        });
        this.tableWidthArray = widthArray.map((l) => (l / sums) * 100);

        return;
        break;

      case "tablecontents":
        this.row = parseInt(words[1]);
        this.column = parseInt(words[2]);
        this.cells[this.row][this.column] = [];
        this.mergedCells[this.row][this.column] = [this.row, this.column];
        return;
        break;
      case "tablecontentslist":
        if (words[1] === "endParagraph" || words[1] === "newLine") {
          //
        } else if (words[1] === "image") {
          // this.cells[this.row][this.column].push(
          //   new Paragraph({
          //     children: [createImageChild(this.mdSourcePath, words[2])],
          //   })
          // );
        } else {
          this.cells[this.row][this.column].push(...resolveEmphasis(words[2]));
          return;
        }
        break;
      case "tableMarge":
        const row = parseInt(words[1]);
        const column = parseInt(words[2]);
        const row2 = parseInt(words[3]);
        const column2 = parseInt(words[4]);
        this.mergedCells[row][column] = [row2, column2];
        for (let j = column; j <= column2; j++) {
          for (let i = row + 1; i <= row2; i++) {
            this.mergedCells[i][j] = [-1, -1];
            this.cells[row][column].push(...this.cells[i][j]);
          }
        }
        return;
        break;
      default:
        return;
        break;
    }
  }

  createTable(slide: PptxGenJS.Slide) {
    let rows = new Array(this.rows);
    for (let i = 0; i < this.rows; i++) {
      rows[i] = new Array<PptxGenJS.TableCell | null>(this.columns);
      for (let j = 0; j < this.columns; j++) {
        let thisWidth = {
          size: this.tableWidthArray[j],
          type: WidthType.PERCENTAGE,
        };

        if (this.tableWidthArray.length > 0) {
          thisWidth = {
            size: this.tableWidthArray[j],
            type: WidthType.PERCENTAGE,
          };
        }

        let pCell = this.cells[i][j];
        const [rowSpan, columnSpan] = this.createRawColumnSpan(
          i,
          j,
          this.mergedCells[i][j][0],
          this.mergedCells[i][j][1]
        );

        let tCell = {
          text: pCell,
        };

        if (
          this.mergedCells[i][j][0] === -1 &&
          this.mergedCells[i][j][1] === -1
        ) {
          rows[i][j] = null;
        } else {
          rows[i][j] = tCell;
        }
      }
    }

    slide.addTable(rows, {
      x: 0,
      y: 2.1,
      w: 7.0,
      rowH: 0.75,
      fill: { color: "F7F7F7" },
      color: "000000",
      fontSize: 16,
      valign: "middle",
      align: "center",
      //border: { pt: "1", color: "FFFFFF" },
    });
  }

  createRawColumnSpan(r1: number, c1: number, r2: number, c2: number) {
    return [r2 - r1 + 1, c2 - c1 + 1];
  }
}

class DocParagraph {
  nodeType: NodeType;
  isFlush: boolean;
  indent: number;
  children: PptxGenJS.TextProps[] = [];
  docStyle: DocStyle;
  isImage: boolean;
  textPropsOptions: PptxGenJS.TextPropsOptions;
  isNewSheet: boolean = false;

  constructor(
    nodeType: NodeType = NodeType.non,
    textPropsOptions: PptxGenJS.TextPropsOptions = { fontSize: 10 },
    indent: number = 0,
    docStyle: DocStyle = DocStyle.Body,
    child?: PptxGenJS.TextProps
  ) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.textPropsOptions = textPropsOptions;
    this.indent = indent;
    this.children = child ? [child] : [];
    this.docStyle = docStyle;
    this.isImage = false;
  }

  // export interface TextProps {
  // 	text?: string
  // 	options?: TextPropsOptions
  // }
  createPptxParagraph(slide: PptxGenJS.Slide) {
    let pStyle = this.isImage ? "picture1" : this.docStyle;

    if (pStyle === "body") {
      pStyle = `body${this.indent + 1}`;
    }

    const r = this.children.map((p) => {
      return {
        text: p.text,
        options: { ...this.textPropsOptions, ...p.options },
      };
    });

    return r;
  }

  addIndent() {
    this.indent++;
    if (this.indent > 3) {
      this.indent = 3;
    }
  }

  removeIndent() {
    this.indent--;
    if (this.indent < 0) {
      this.indent = 0;
    }
  }

  addChild(s: string | PptxGenJS.TextProps, isImage = false) {
    const ss = typeof s === "string" ? { text: s } : s;
    this.isImage = false;
    if (isImage && this.children.length === 1) {
      this.isImage = true;
    }
    this.children.push(ss);
  }

  addChildren(s: PptxGenJS.TextProps[]) {
    this.isImage = false;
    this.children.push(...s);
  }
}



function addParagraphBlock(
  currentSlide: PptxGenJS.Slide,
  paragraphBlock: PptxGenJS.TextProps[]
) {
  // cerate paragraph block
  currentSlide.addText(paragraphBlock, {
    x: 0.5,
    y: 0.5,
    w: 5.75,
    h: 2.0,
    //fill: { color: pptx.SchemeColor.background2 },
    //color: pptx.SchemeColor.accent1,
  });
}

// ############################################################

function resolveWordCommentsCommands(
  wdCommandList: string[],
  documentInfo: { [v: string]: string }
) {
  const documentInfoKeys = Object.keys(documentInfo);
  if (documentInfoKeys.includes(wdCommandList[0])) {
    documentInfo[wdCommandList[0]] = wdCommandList[1];
    return true;
  }

  return false;
}

async function resolveWordDownCommandEx(
  line: string,
  docParagraph: DocParagraph,
  mdSourcePath: string
) {
  const words = line.split(_sp);
  let current: DocParagraph;
  const nodeType = words[0] as NodeType;
  let style: DocStyle;
  let child: PptxGenJS.TextProps;

  switch (nodeType) {
    case "section":
      current = new DocParagraph(NodeType.section, { fontSize: 48 });
      resolveEmphasis(words[2]).forEach((x) => current.addChild(x));
      return current;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      docParagraph.addChild({
        text: " ",
        options: {
          fontSize: 32,
          bullet: true,
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
        },
      });
      return docParagraph;
      break;
    case NodeType.OderList:
      docParagraph.addChild({
        text: " ",
        options: {
          fontSize: 32,
          bullet: { type: "number", style: "romanLcPeriod" },
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
        },
      });
      return docParagraph;
      break;
    case "code":
      docParagraph.addChild({
        text: words[1],
        options: {
          fontSize: 36,
          fontFace: "Arial",
          //color: pptx.SchemeColor.accent5,
          highlight: "FFFF00",
        },
      });
      return docParagraph;
      break;
    case NodeType.link:
      // if (!words[3]) {
      //   let children = [];
      //   children.push(new SimpleField(` REF ${words[1]} \\w \\h `));
      //   children.push(new TextRun("---"));
      //   children.push(new SimpleField(` REF ${words[1]} \\h `));
      //   nodes.addChildren(children);
      // } else {
      //   child = new ExternalHyperlink({
      //     children: [
      //       new TextRun({
      //         text: words[3],
      //       }),
      //     ],
      //     link: words[1],
      //   });
      //   nodes.addChild(child);
      // }
      return docParagraph;
      break;
    case NodeType.image:
      // child = createImageChild(mdSourcePath, words[1]);
      // nodes.addChild(child, true);
      return docParagraph;
      break;
    case NodeType.hr:
      docParagraph.isFlush = true;
      docParagraph.isNewSheet = true;
      return docParagraph;
    case "text":
      const admonition = words[1].match(/^(note|warning)(:\s)(.*)/i);

      let s = words[1];
      if (admonition && admonition[3]) {
        s = admonition[3];
        const admonitionType = admonition[1];
        docParagraph.nodeType = admonitionType as NodeType;
        docParagraph.docStyle = resolveAdmonition(admonitionType);
      }

      const mathBlock = s.match(/^\$(.*)\$$/);
      // if (mathBlock?.length) {
      //   const child = await createMathImage(mathBlock[1]);
      //   nodes.addChild(child, true);
      // } else {
      //   resolveEmphasis(s).forEach((x) => nodes.addChild(x));
      // }

      resolveEmphasis(s).forEach((x) => docParagraph.addChild(x));

      return docParagraph;
      break;
    case "indentPlus":
      docParagraph.addIndent();
      return docParagraph;
      break;
    case "indentMinus":
      docParagraph.removeIndent();
      return docParagraph;
      break;
    case "newLine":
      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        docParagraph.addChild({ text: "", options: { breakLine: true } });
        docParagraph.isFlush = true;
      }
      return docParagraph;
    default:
      return docParagraph;
  }
}

function resolveAdmonition(s: string) {
  let admonition: DocStyle = DocStyle.note1;
  switch (s.toLocaleLowerCase()) {
    case "note":
      admonition = DocStyle.note1;
      break;
    case "warning":
      admonition = DocStyle.warning1;
      break;
    default:
      //
      break;
  }
  return admonition as DocStyle;
}

async function createMathImage(mathEq: string) {
  const options = {
    width: 1280,
    ex: 8,
    em: 16,
  };

  const svgStr = TeXToSVG(mathEq, options);

  // create png
  const pngArray = await svg2imagePng(svgStr);
  let pngBuffer = Buffer.from(pngArray);
  const sizeImageMath = imageSize.imageSize(pngBuffer);

  const child = new ImageRun({
    data: pngBuffer,
    transformation: {
      width: sizeImageMath.width!,
      height: sizeImageMath.height!,
    },
  });
  return child;
}

function createImageChild(mdSourcePath: string, imagePathR: string) {
  const imagePath = Path.resolve(mdSourcePath, imagePathR);

  const sizeImage = imageSize.imageSize(imagePath);

  const maxSize = 400; //convertInchesToTwip(3);

  let width = sizeImage.width ?? 100;
  let height = sizeImage.height ?? 100;

  if (width > maxSize || height > maxSize) {
    const r = maxSize / Math.max(width, height);
    width *= r;
    height *= r;
  }

  //slide.addImage({ path: IMAGE_PATHS.ccLogo.path, x: 9.9, y: 1.1, w: 2.5, h: 2.5, rounding: true });
  return { path: imagePath, x: 3, y: 3, w: 2.5, h: 2.5, rounding: true };
}

function resolveEmphasis(source: string) {
  let rg = /<(|\/)sub>|<(|\/)sup>|<(|\/)codespan>|<(|\/)i>|<(|\/)b>|<(|\/)~~>/g;

  let indexBefore = 0;
  const stack = [];
  let options: PptxGenJS.TextPropsOptions = {
    bold: false,
    italic: false,
    strike: false,
    subscript: false,
    superscript: false,
    highlight: "",
  };
  let result: any;
  while ((result = rg.exec(source)) !== null) {
    // text
    let text = source.substring(indexBefore, result.index);

    if (text) {
      stack.push({ text, options });
    }
    // tag
    text = source.substring(result.index + 1, rg.lastIndex - 1);
    const tag = text.replace("/", "");
    const isOn = text === tag;

    if (tag === "codespan") {
      if (isOn) {
        options = { ...options, highlight: "FF88CC" };
      } else {
        options = { ...options, highlight: "" };
      }
    } else {
      options = { ...options, [resolveEmphasisTag(tag)]: isOn };
    }
    indexBefore = rg.lastIndex;
  }

  // text
  let text = source.substring(indexBefore);
  if (text) {
    stack.push({ text, options });
  }
  return stack;

  function resolveEmphasisTag(tag: string) {
    if (tag === "b") {
      return "bold";
    }
    if (tag === "i") {
      return "italic";
    }
    if (tag === "sup") {
      return "superscript";
    }
    if (tag === "sub") {
      return "subscript";
    }
    if (tag === "~~") {
      return "strike";
    }
    return "";
  }
}
