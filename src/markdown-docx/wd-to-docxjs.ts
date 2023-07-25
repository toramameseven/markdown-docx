import { DocxOption, MessageType, fileExists } from "./common";
import * as fs from "fs";
import * as Path from "path";
import * as imageSize from "image-size";
const texToSvg = require("tex-to-svg");

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
  TableCell,
  TableRow,
  //TextDirection,
  TableOfContents,
  TextRun,
  //VerticalAlign,
  WidthType,
  // Document as DocumentDocx,
  //convertInchesToTwip,
  //PageReference,
  SimpleField,
  TableLayoutType,
  PageBreak,
} from "docx";
import { svg2imagePng } from "../tools/svg-png-image";

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
  cells: Paragraph[][][];
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

  async doTableCommand(line: string) {
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
          this.cells[this.row][this.column].push(
            new Paragraph({
              children: [await createImageChild(this.mdSourcePath, words[2])],
            })
          );
        } else {
          this.cells[this.row][this.column].push(
            new Paragraph({ children: resolveEmphasis(words[2]).stack })
          );
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

  createTable() {
    let rows = new Array(this.rows);
    for (let i = 0; i < this.rows; i++) {
      rows[i] = new Array<TableCell | null>(this.columns);
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

        let tCell = new TableCell({
          children: pCell,
          rowSpan,
          columnSpan,
          width: thisWidth,
        });

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

    let tableRaws = rows.map((r) => {
      const rr = r.filter(Boolean);
      return new TableRow({ children: rr });
    });

    const tableJs = new Table({
      layout: TableLayoutType.FIXED,
      rows: tableRaws,
      style: "styleN",
      // indent: {
      //   size: "20mm",
      //   type: WidthType.AUTO,
      // },
      width: {
        size: 95,
        type: WidthType.PERCENTAGE,
      },
    });
    return tableJs;
  }

  createRawColumnSpan(r1: number, c1: number, r2: number, c2: number) {
    return [r2 - r1 + 1, c2 - c1 + 1];
  }
}

class DocParagraph {
  nodeType: NodeType;
  isFlush: boolean;
  indent: number;
  children: ParagraphChild[] = [];
  docStyle: DocStyle;
  isImage: boolean;
  refId: string = "";
  refString: string = "";

  constructor(
    nodeType: NodeType = NodeType.non,
    indent: number = 0,
    docStyle: DocStyle = DocStyle.Body,
    childe: ParagraphChild = new TextRun("")
  ) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.indent = indent;
    this.children = [childe];
    this.docStyle = docStyle;
    this.isImage = false;
  }

  createDocxParagraph(): Paragraph | Table {
    let pStyle = this.isImage ? "picture1" : this.docStyle;

    if (pStyle === "body") {
      pStyle = `body${this.indent + 1}`;
    }

    const docxR = new Paragraph({
      children: this.children,
      style: pStyle,
    });
    return docxR;
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

  addChild(s: string | ParagraphChild, isImage = false) {
    const ss = typeof s === "string" ? new TextRun(s) : s;
    this.isImage = false;
    if (isImage && this.children.length === 1) {
      this.isImage = true;
    }
    this.children.push(ss);
  }

  addChildren(s: ParagraphChild[]) {
    this.isImage = false;
    this.children.push(...s);
  }
}

/**
 *
 * @param wd
 * @param docxTemplatePath
 * @param docxOutPath
 * @param mdSourcePath
 */
export async function wdToDocxJs(
  wd: string,
  docxTemplatePath: string,
  docxOutPath: string,
  mdSourcePath: string,
  option: DocxOption
): Promise<void> {
  let patches: (Paragraph | Table | TableOfContents)[] = [];

  const lines = (wd + "\nEndline").split(/\r?\n/);
  let currentParagraph = new DocParagraph(NodeType.text);
  let tableJs: TableJs;
  let insideTable = false;

  const documentInfo = {
    title: "",
    subTitle: "",
    division: "",
    date: "",
    author: "",
    docNumber: "",
    crossRef:"[[$n $t (p.$p)]]"
  };

  // patch parameter
  const params ={};

  for (let i = 0; i < lines.length; i++) {
    const wdCommandList = lines[i].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      tableJs = new TableJs(
        parseInt(wdCommandList[1]),
        parseInt(wdCommandList[2]),
        mdSourcePath
      );
      insideTable = true;
      continue;
    }

    // table command
    if (wdCommandList[0].includes("table")) {
      tableJs!.doTableCommand(lines[i]);
    } else {
      // in not table command, create table.
      if (insideTable) {
        patches.push(tableJs!.createTable());
        insideTable = false;
      }
    }

    // html comment command <!-- word xxxx -->
    const isResolveCommand = resolveWordCommentsCommands(
      wdCommandList,
      patches,
      documentInfo,
      params
    );

    if (isResolveCommand) {
      continue;
    }

    // body commands
    currentParagraph = await resolveWordDownCommandEx(
      lines[i],
      currentParagraph,
      mdSourcePath,
      option,
      documentInfo
    );

    // when paragraph end, flush paragraph
    if (currentParagraph.isFlush) {
      const p = currentParagraph.createDocxParagraph();
      patches.push(p);
      // reset paragraph. but keep the indent.
      currentParagraph = new DocParagraph(
        NodeType.text,
        currentParagraph.indent
      );
    }
  }
  await createDocxPatch(patches, docxTemplatePath, docxOutPath, documentInfo, params);
}

// ############################################################

function createListType(listType: NodeType, listOrder: number) {
  if (listType === NodeType.NormalList) {
    return `nList${listOrder}` as DocStyle;
  }
  if (listType === NodeType.OderList) {
    return `numList${listOrder}` as DocStyle;
  }
  return DocStyle.Error;
}

function resolveWordCommentsCommands(
  wdCommandList: string[],
  patches: (Paragraph | Table | TableOfContents)[],
  documentInfo: { [v: string]: string },
  params: { [v: string]: Params }
) {
  // table of contents
  if (wdCommandList[0] === "toc") {
    const ptitle = new Paragraph(wdCommandList[2]);
    patches.push(ptitle);

    const toc = new TableOfContents("Summary", {
      hyperlink: true,
      headingStyleRange: `1-${wdCommandList[1]}`,
      captionLabel: wdCommandList[2],
    });
    patches.push(toc);

    const p = new Paragraph({
      children: [
        new TextRun({
          text: "",
          break: 1,
        }),
      ],
    });

    patches.push(p);
    return true;
  }

  // patch words
  if (wdCommandList[0] === "param") {
    params[wdCommandList[1]] ={
      type: PatchType.PARAGRAPH,
      children: [new TextRun(wdCommandList[2])],
    },
    
    wdCommandList[2];
    return true;
  }

  // option
  const documentInfoKeys = Object.keys(documentInfo);
  if (documentInfoKeys.includes(wdCommandList[0])) {
    documentInfo[wdCommandList[0]] = wdCommandList[1];
    return true;
  }
  return false;
}

async function resolveWordDownCommandEx(
  line: string,
  currentParagraph: DocParagraph,
  mdSourcePath: string,
  option: DocxOption,
  documentInfo: { [v: string]: string }
) {
  const words = line.split(_sp);
  let current: DocParagraph;
  const nodeType = words[0] as NodeType;
  let style: DocStyle;
  let child: ParagraphChild;


  switch (nodeType) {
    case "section":
      // section 2  heading2(id)
      child = new Bookmark({
        id: words[2],
        //children: resolveEmphasis(words[2]),
        children:[],
      });
      current = new DocParagraph(
        nodeType,
        currentParagraph.indent,
        `hh${words[1]}` as DocStyle,
        // `${words[1]}` as DocStyle, // we do not know how this works.
        //child
      );
      current.refId = words[2];
      return current;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(
        NodeType.NormalList,
        currentParagraph.indent,
        style
      );
      return current;
      break;
    case NodeType.OderList:
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(
        NodeType.OderList,
        currentParagraph.indent,
        style
      );
      return current;
      break;
    case "code":
      child = new TextRun(words[1]);
      current = new DocParagraph(
        nodeType,
        currentParagraph.indent,
        "code",
        child
      );
      return current;
      break;
    case NodeType.link:
      if (!words[3]) {
        // internal link
        let children = resolveXref(words[1], documentInfo.crossRef);
        currentParagraph.addChildren(children);
      } else {
        child = new ExternalHyperlink({
          children: [
            new TextRun({
              text: words[3],
            }),
          ],
          link: words[1],
        });
        currentParagraph.addChild(child);
      }
      return currentParagraph;
      break;
    case NodeType.image:
      child = await createImageChild(mdSourcePath, words[1], option);
      currentParagraph.addChild(child, true);
      return currentParagraph;
      break;
    case "text":
      const admonition = words[1].match(/^(note|warning)(:\s)(.*)/i);

      let s = words[1];
      if (admonition && admonition[3]) {
        s = admonition[3];
        const admonitionType = admonition[1];
        currentParagraph.nodeType = admonitionType as NodeType;
        currentParagraph.docStyle = resolveAdmonition(admonitionType);
      }

      const mathBlock = s.match(/^\$(.*)\$$/);
      if (mathBlock?.length && option?.mathExtension) {
        const child = await createMathImage(mathBlock[1]);
        currentParagraph.addChild(child, true);
      } else {
        const {stack, sectionText} = resolveEmphasis(s);
        stack.forEach((x) => currentParagraph.addChild(x));
        currentParagraph.refString += sectionText;
      } 

      return currentParagraph;
      break;
    case "indentPlus":
      currentParagraph.addIndent();
      return currentParagraph;
      break;
    case "indentMinus":
      currentParagraph.removeIndent();
      return currentParagraph;
      break;
    case "newLine":
      if (words[1] === "convertHeading End"){
        child = new Bookmark({
          id: currentParagraph.refId,
          children: [new TextRun(currentParagraph.refString)]
        });
        current = new DocParagraph(
          nodeType,
          currentParagraph.indent,
          currentParagraph.docStyle,
          child
        );
        current.isFlush = true;
        return current;
      }
      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        currentParagraph.isFlush = true;
      }
      return currentParagraph;
    case "newPage":
      currentParagraph.addChild(new PageBreak());
      currentParagraph.isFlush = true;
      return currentParagraph;
    default:
      return currentParagraph;
  }
}

function resolveXref(linkRef: string, refFormat: string ="[[$n $t p.$p]]") {
  const refItems = [];
  for (let i = 0; i < refFormat.length; i++) {
    let t = refFormat.slice(i, i + 2);
    console.log(t);
    if (t.match(/\$n|\$p|\$t/)) {
      refItems.push(t);
      i++;
    } else {
      refItems.push(refFormat.slice(i, i + 1));
    }
  }

  const children = [];
  //  \w : full contents
  //  \h :
  for (let i = 0; i < refItems.length; i++) {

    switch (refItems[i]) {
      case '$n':
        children.push(new SimpleField(` REF ${linkRef} \\w \\h `));
        break;
      case '$t':
        children.push(new SimpleField(` REF ${linkRef} \\h `));
        break;
      case '$p':
        children.push(new SimpleField(` PAGEREF ${linkRef} \\h `));
        // Expected output: "Mangoes and papayas are $2.79 a pound."
        break;
      default:
        children.push(new TextRun(refItems[i]));
    }
  }
  return children;
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

  const svgStr: string = texToSvg(mathEq, options);

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

async function createImageChild(
  mdSourcePath: string,
  imagePathR: string,
  option?: DocxOption
) {
  const imagePath = Path.resolve(mdSourcePath, imagePathR);

  if (!(await fileExists(imagePath))) {
    option?.message?.(
      MessageType.err,
      `No image ${imagePath}`,
      "wd-to-docxjs",
      false
    );
    const errorChild = new TextRun(`[No Image: ${imagePath}]`);
    return errorChild;
  }

  const sizeImage = imageSize.imageSize(imagePath);

  const maxSize = 400; //convertInchesToTwip(3);

  let width = sizeImage.width ?? 100;
  let height = sizeImage.height ?? 100;

  if (width > maxSize || height > maxSize) {
    const r = maxSize / Math.max(width, height);
    width *= r;
    height *= r;
  }
  const child = new ImageRun({
    data: fs.readFileSync(imagePath),
    transformation: {
      width,
      height,
    },
  });
  return child;
}

type Params  = {
  type: PatchType.PARAGRAPH;
  children: TextRun[];
};

/**
 *
 * @param children array of paragraph or table
 * @param docxTemplatePath docx file path
 * @param docxOutPath docx path for output
 */
export async function createDocxPatch(
  children: (Paragraph | Table | TableOfContents)[],
  docxTemplatePath: string,
  docxOutPath: string,
  docInfo: { [v: string]: string },
  params: { [v: string]: Params }
) {

  const patchDoc = await patchDocument(fs.readFileSync(docxTemplatePath), {
    patches: {
      paragraphReplace: {
        type: PatchType.DOCUMENT,
        children: children,
      },
      ...params
    },
  });
  fs.writeFileSync(docxOutPath, patchDoc);
}

function resolveEmphasis(source: string) {
  let rg = /<(|\/)sub>|<(|\/)sup>|<(|\/)codespan>|<(|\/)i>|<(|\/)b>|<(|\/)~~>/g;

  let indexBefore = 0;
  let sectionText:string = "";
  const stack = [];
  let textProp = {
    bold: false,
    italics: false,
    strike: false,
    subScript: false,
    superScript: false,
    style: "",
  };
  let result: any;
  while ((result = rg.exec(source)) !== null) {
    // text
    let text = source.substring(indexBefore, result.index);
    if (text) {
      sectionText += text;
      stack.push(new TextRun({ text, ...textProp }));
    }
    // tag
    text = source.substring(result.index + 1, rg.lastIndex - 1);
    const tag = text.replace("/", "");
    const isOn = text === tag;

    if (tag === "codespan") {
      if (isOn) {
        textProp["style"] = "codespan";
      } else {
        textProp["style"] = "";
      }
    } else {
      textProp = { ...textProp, [resolveEmphasisTag(tag)]: isOn };
    }
    indexBefore = rg.lastIndex;
  }

  // text
  let text = source.substring(indexBefore);
  if (text) {
    sectionText += text;
    stack.push(new TextRun({ text, ...textProp }));
  }
  return {stack, sectionText};

  function resolveEmphasisTag(tag: string) {
    if (tag === "b") {
      return "bold";
    }
    if (tag === "i") {
      return "italics";
    }
    if (tag === "sup") {
      return "superScript";
    }
    if (tag === "sub") {
      return "subScript";
    }
    if (tag === "~~") {
      return "strike";
    }
    return "";
  }
}
