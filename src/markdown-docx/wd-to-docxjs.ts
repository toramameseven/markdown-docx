import {
  DocxOption,
  MessageType,
  ShowMessage,
  fileExists,
  getDocxDocTitleFromWd,
  getFloat,
} from "./common";
import * as fs from "fs";
import * as Path from "path";
import * as imageSize from "image-size";
const texToSvg = require("tex-to-svg");
let showMessageThis: ShowMessage | undefined;

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
  CheckBox,
  AlignmentType,
} from "docx";
import { svg2imagePng } from "./tools/svg-png-image";
import { WdCommand, wdCommand } from "./wd0-to-wd";
//import { initialize } from "svg2png-wasm";
import { Bookmarks } from "./bookmarks";
// import { OoxParameters } from "./markdown-to-wd0";
// import mermaid from "mermaid";

const _sp = "\t";

// eslint-disable-next-line @typescript-eslint/naming-convention
const DocStyle = {
  hh0: "hh0",
  hh1: "hh1",
  hh2: "hh2",
  hh3: "hh3",
  hh4: "hh4",
  hh5: "hh5",
  hh6: "hh6",
  // eslint-disable-next-line @typescript-eslint/naming-convention
  Body: "body",
  // eslint-disable-next-line @typescript-eslint/naming-convention
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
  imageCaption: "imageCaption",
  tableCaption: "tableCaption",
  // eslint-disable-next-line @typescript-eslint/naming-convention
  Error: "Error",
} as const;
type DocStyle = (typeof DocStyle)[keyof typeof DocStyle];

import { DocumentInfoParams } from "./types";

type DocumentInfo = {
  placeholders: { [v: string]: string };
  params: { [v in DocumentInfoParams]?: string };
  bookmarks: Bookmarks;
};

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

  async doTableCommand(line: string, documentInfo: DocumentInfo) {
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
          // align
          const alignInfo = (documentInfo.params.tableAlign ?? "") + "llllllllllllllllllllllllllll";
          let align: AlignmentType = AlignmentType.LEFT;
          if ((alignInfo[this.column]).toLowerCase() === 'c') {
            align = AlignmentType.CENTER;
          }
          if ((alignInfo[this.column]).toLowerCase() === 'r') {
            align = AlignmentType.RIGHT;
          }
          this.cells[this.row][this.column].push(
            new Paragraph({ children: await resolveEmphasis(words[2], documentInfo), alignment: align })
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

  applyToPatch(
    patches: (Paragraph | Table | TableOfContents)[],
    documentInfo: DocumentInfo
  ) {
    if (documentInfo.params.tableCaption && documentInfo.params.tablePrefix) {
      patches.push(
        this.createTableCaptionParagraph(
          documentInfo.params.tableCaption,
          documentInfo
        )
      );
    }
    // clear caption
    documentInfo.params.tableCaption = "";
    patches.push(this.createTable(documentInfo));
    patches.push(new Paragraph(" "));
  }

  createTable(documentInfo: DocumentInfo) {
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
      
      // clear table info
      documentInfo.params.tableAlign = undefined;
      documentInfo.params.tableStyle = undefined;
    }

    let tableRaws = rows.map((r) => {
      const rr = r.filter(Boolean);
      return new TableRow({ children: rr });
    });

    const tableWidth = getFloat(documentInfo.params.tableWidth, 20, 100);

    const tableJs = new Table({
      layout: TableLayoutType.FIXED,
      rows: tableRaws,
      style: "styleN",
      // indent: {
      //   size: "20mm",
      //   type: WidthType.AUTO,
      // },
      width: {
        size: tableWidth,
        type: WidthType.PERCENTAGE,
      },
    });
    return tableJs;
  }

  createRawColumnSpan(r1: number, c1: number, r2: number, c2: number) {
    return [r2 - r1 + 1, c2 - c1 + 1];
  }

  createTableCaptionParagraph(
    tableCaption: string,
    documentInfo: DocumentInfo
  ) {
    const tableCaptionId = documentInfo.bookmarks.slugify("table-" + tableCaption);
    const tableCaptionBookmark = new Bookmark({
      id: tableCaptionId,
      children: [
        new TextRun(documentInfo.params.tablePrefix + " "),
        new SimpleField(` SEQ GRIDTABLE \\* ARABIC  `),
      ],
    });

    const docxR = new Paragraph({
      children: [tableCaptionBookmark, new TextRun(" " + tableCaption)],
      style: DocStyle.tableCaption,
    });
    return docxR;
  }
}

class DocParagraph {
  nodeType: WdCommand;
  isFlush: boolean;
  indent: number;
  children: ParagraphChild[] = [];
  childrenRaw: string[] = [];
  docStyle: DocStyle;
  isImage: boolean;
  refId: string = "";
  imageCaption: string = "";
  isListBefore: boolean = false;

  constructor(
    nodeType: WdCommand = wdCommand.non,
    indent: number = 0,
    docStyle: DocStyle = DocStyle.Body,
    isListBefore: boolean = false,
    child?: ParagraphChild
  ) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.indent = indent;
    if (child) {
      this.children = [child];
    }
    this.docStyle = docStyle;
    this.isImage = false;
    this.isListBefore = isListBefore;
  }

  createDocxParagraph(): Paragraph | Table | undefined {
    if (this.children.length === 0) {
      return undefined;
    }
    let pStyle = this.isImage ? "picture1" : this.docStyle;

    // if more then one item, its not block images
    if (this.children.length > 1) {
      this.imageCaption = "";
    }

    if (pStyle === "body") {
      pStyle = `body${this.indent + 1}`;
    }

    const docxR = new Paragraph({
      children: this.children,
      style: pStyle,
    });

    this.initialize();
    return docxR;
  }

  createDocxParagraphAsImageCaption(
    documentInfo: DocumentInfo
  ): Paragraph | Table | undefined {
    if (this.imageCaption === "") {
      return undefined;
    }

    const imageCaptionId = documentInfo.bookmarks.slugify("fig-" + this.imageCaption);
    const imageCaptionBookmark = new Bookmark({
      id: imageCaptionId,
      children: [
        new TextRun(documentInfo.params.figurePrefix + " "),
        new SimpleField(` SEQ DOCXFIG \\* ARABIC  `),
      ],
    });

    const children = documentInfo.params.figurePrefix ? [imageCaptionBookmark, new TextRun(" " + this.imageCaption)] : [new TextRun(" ")];

    const docxR = new Paragraph({
      children,
      style: DocStyle.imageCaption,
    });

    this.imageCaption = "";
    this.initialize();
    return docxR;
  }

  createRawString(): string {
    if (this.childrenRaw.length === 0) {
      return "";
    }

    const r = this.childrenRaw.join("\n");
    this.initialize();
    return r;
  }

  initialize() {
    (this.nodeType = wdCommand.text), (this.isFlush = false);
    this.children = [];
    this.childrenRaw = [];
    this.docStyle = DocStyle.Body;
    this.isImage = false;
    this.refId = "";
    this.isListBefore = false;
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
    if (isImage && this.children.length === 0) {
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
  // message function
  option.message && (showMessageThis = option.message);

  // for placeholder
  let patches: (Paragraph | Table | TableOfContents)[] = [];

  // get tile from heading 1
  const title = getDocxDocTitleFromWd(wd);

  const lines = (wd + "\nEndline").split(/\r?\n/);
  let currentParagraph = new DocParagraph(wdCommand.text);
  let tableJs: TableJs | undefined = undefined;

  const documentInfo: DocumentInfo = {
    placeholders: { title },
    params: { refFormat: "[[$n $t (p.$p)]]", captionRefFormat: "($t)" },
    bookmarks: new Bookmarks(),
  };

  // initialize params
  documentInfo.params.tablePrefix = "Table";
  documentInfo.params.figurePrefix = "Fig.";
  documentInfo.params.useCheckBox = "true";

  // patch parameter

  const patchInfo: { [v: string]: PatchInfo } = {};
  for (let i = 0; i < lines.length; i++) {
    const wdCommandList = lines[i].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      if (tableJs) {
        tableJs.applyToPatch(patches, documentInfo);
        tableJs = undefined;
      }
      tableJs = new TableJs(
        parseInt(wdCommandList[1]),
        parseInt(wdCommandList[2]),
        mdSourcePath
      );
      continue;
    }

    // table command
    if (wdCommandList[0].includes("table")) {
      tableJs?.doTableCommand(lines[i], documentInfo);
    } else {
      // in not table command, create table.
      if (tableJs) {
        tableJs.applyToPatch(patches, documentInfo);
        tableJs = undefined;
      }
    }

    // html comment command <!-- word xxxx -->
    const isResolveCommand = resolveCommentCommand(
      wdCommandList,
      patches,
      documentInfo,
      patchInfo
    );

    if (isResolveCommand) {
      continue;
    }

    // body commands
    currentParagraph = await resolveWDCommandEx(
      lines[i],
      currentParagraph,
      mdSourcePath,
      option,
      documentInfo
    );

    // when paragraph end, flush paragraph
    if (currentParagraph.isFlush) {
      if (currentParagraph.isListBefore && currentParagraph.docStyle.substring(0, 2) === 'hh') {
        patches.push(new Paragraph(" "));
      }
      currentParagraph.isListBefore = false;
      const p = currentParagraph.createDocxParagraph();
      if (p) {
        patches.push(p);
      }

      // image figure caption
      const p2 =
        currentParagraph.createDocxParagraphAsImageCaption(documentInfo);
      if (p2) {
        patches.push(p2);
      }
    }
  } // end for

  if (title) {
    patchInfo["title"] = {
      type: PatchType.PARAGRAPH,
      children: [new TextRun(title)],
    };
  }

  // apply to placeholders
  await createDocxPatch(
    patches,
    docxTemplatePath,
    docxOutPath,
    documentInfo,
    patchInfo
  );
}

// ############################################################

function createListType(listType: WdCommand, listOrder: number) {
  if (listType === wdCommand.NormalList) {
    return `nList${listOrder}` as DocStyle;
  }
  if (listType === wdCommand.OderList) {
    return `numList${listOrder}` as DocStyle;
  }
  return DocStyle.Error;
}

function resolveCommentCommand(
  wdCommandList: string[],
  patches: (Paragraph | Table | TableOfContents)[],
  documentInfo: DocumentInfo,
  patchInfo: { [v: string]: PatchInfo }
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
  } // table of contents

  // placeholder words
  if (wdCommandList[0] === "placeholder") {
    patchInfo[wdCommandList[1]] = {
      type: PatchType.PARAGRAPH,
      children: [new TextRun(wdCommandList[2])],
    };
    return true;
  } // placeholder words

  // param words
  if (wdCommandList[0] === "param") {
    documentInfo.params[wdCommandList[1] as DocumentInfoParams] =
      wdCommandList[2];
    return true;
  } // param words

  return false;
}

async function resolveWDCommandEx(
  line: string,
  currentParagraph: DocParagraph,
  mdSourcePath: string,
  option: DocxOption,
  documentInfo: DocumentInfo
) {
  const words = line.split(_sp);
  let current: DocParagraph;
  const nodeType = words[0].split("/")[0] as WdCommand;
  let style: DocStyle;
  let child: ParagraphChild;

  switch (nodeType) {
    case wdCommand.section:
      // ## hh1
      // ### hh2
      // ###### hh5
      // markdownHeader - 1 = hh
      const hhHeader = parseInt(words[1]) - 1;
      current = new DocParagraph(
        nodeType,
        currentParagraph.indent,
        `hh${hhHeader}` as DocStyle, // `${words[1]}` as DocStyle, // we do not know how this works.
        currentParagraph.isListBefore
      );
      current.refId = words[2];
      documentInfo.bookmarks.slugify(current.refId);
      return current;
      break;
    case wdCommand.NormalList:
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(
        wdCommand.NormalList,
        currentParagraph.indent,
        style
      );
      return current;
      break;
    case wdCommand.OderList:
      style = createListType(nodeType, parseInt(words[1]));
      current = new DocParagraph(
        wdCommand.OderList,
        currentParagraph.indent,
        style
      );
      return current;
      break;
    case wdCommand.code:
      current = currentParagraph;
      if (current.docStyle === DocStyle.code) {
        child = new TextRun({ text: words[1], break: 1 });
      } else {
        child = new TextRun({ text: words[1] });
      }
      current.childrenRaw.push(words[1]);
      current.addChild(child);
      current.docStyle = "code";
      return current;
      break;
    case wdCommand.link:
      if (!words[3]) {
        // internal link
        let children = resolveXref(words[1], documentInfo);
        currentParagraph.addChildren(children);
      } else {
        // external link
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
    case wdCommand.image:
      child = await createImageChild(mdSourcePath, words[1], option, documentInfo);
      currentParagraph.addChild(child, true);
      currentParagraph.imageCaption = words[2];
      return currentParagraph;
      break;
    case "text":
      // admonition
      const admonition = words[1].match(/^(note|warning)(:\s)(.*)/i);

      let s = words[1];
      if (admonition && admonition[3]) {
        s = admonition[3];
        const admonitionType = admonition[1];
        currentParagraph.nodeType = admonitionType as WdCommand;
        currentParagraph.docStyle = resolveAdmonition(admonitionType);
      }

      // math $~~~~~$
      const mathBlock = s.match(/^\$(.+)\$$/);
      if (mathBlock?.length && option?.mathExtension) {
        const child = await createMathImage(mathBlock[1]);
        currentParagraph.addChild(child, true);
      } else {
        const stack = await resolveEmphasis(s, documentInfo);
        stack.forEach((x) => currentParagraph.addChild(x));
      }

      return currentParagraph;
      break;
    case wdCommand.indentPlus:
      currentParagraph.addIndent();
      return currentParagraph;
      break;
    case wdCommand.indentMinus:
      currentParagraph.removeIndent();
      currentParagraph.isListBefore = true;
      return currentParagraph;
      break;
    case wdCommand.newLine:
      if (words[1] === "convertHeading End") {
        if (currentParagraph.docStyle === "hh0") {
          current = new DocParagraph(wdCommand.text, 0);
          return current;
        }
        child = new Bookmark({
          id: currentParagraph.refId,
          children: currentParagraph.children,
        });
        current = new DocParagraph(
          nodeType,
          currentParagraph.indent,
          currentParagraph.docStyle,
          currentParagraph.isListBefore,
          child
        );
        current.isFlush = true;
        return current;
      }
      if (words[1] === "convertCode") {
        // if (words[2] === "mermaid") {
        //   // do render mermaid

        //   const mermaid = require("mermaid");

        //   const diagramCode = `
        //   graph LR
        //       A-->B
        //       B-->C
        //       C-->D
        //       D-->A
        //   `;
        //   const { svg } = await mermaid.render("diagramId", diagramCode);
        //   console.log(svg);
        // }
        if (words[2] === "math") {
          const child = await createMathImage(
            currentParagraph.createRawString()
          );
          currentParagraph.addChild(child, true);
          currentParagraph.isFlush = true;
          return currentParagraph;
        }
        currentParagraph.isFlush = true;
        return currentParagraph;
      }
      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        currentParagraph.isFlush = true;
      }
      return currentParagraph;
    case wdCommand.newPage:
      currentParagraph.addChild(new PageBreak());
      currentParagraph.isFlush = true;
      return currentParagraph;
    default:
      return currentParagraph;
  }
}

function getLinkType(linkRef: string): "section" | "caption" {
  if (linkRef.slice(0, "fig-".length) === "fig-") {
    return "caption";
  }
  if (linkRef.slice(0, "table-".length) === "table-") {
    return "caption";
  }
  return "section";
}

function resolveXref(linkRef: string, documentInfo: DocumentInfo) {

  let refFormat = getLinkType(linkRef) === "section" ? documentInfo.params.refFormat : documentInfo.params.captionRefFormat;
  refFormat = refFormat || "[[$n $t p.$p]]";

  const refItems = [];
  for (let i = 0; i < refFormat.length; i++) {
    let t = refFormat.slice(i, i + 2);
    //console.log(t);
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
      case "$n":
        children.push(new SimpleField(` REF ${linkRef} \\w \\h `));
        break;
      case "$t":
        children.push(new SimpleField(` REF ${linkRef} \\h `));
        break;
      case "$p":
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
  option?: DocxOption,
  documentInfo?: DocumentInfo
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
  const maxSizeStr = documentInfo?.params.imageWidth ?? '500';
  const maxSize = parseInt(maxSizeStr);

  let width = sizeImage.width ?? 100;
  let height = sizeImage.height ?? 100;

  const r = width / height;

  if (width > maxSize) {
    width = maxSize;
    height = width / r;
  }

  if (height > maxSize) {
    height = maxSize;
    width = height * r;
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

type PatchInfo = {
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
  docInfo: { placeholders: {}; params: {} },
  patchInfo: { [v: string]: PatchInfo }
) {
  const patchDoc = await patchDocument(fs.readFileSync(docxTemplatePath), {
    patches: {
      paragraphReplace: {
        type: PatchType.DOCUMENT,
        children: children,
      },
      ...patchInfo,
    },
  });
  fs.writeFileSync(docxOutPath, patchDoc);
}

async function resolveEmphasis(source: string, documentInfo: DocumentInfo) {
  let rg =
    /<(|\/)sub>|<(|\/)sup>|<(|\/)codespan>|<(|\/)i>|<(|\/)b>|<(|\/)~~>|☑|☐|\$/g;

  let indexBefore = 0;
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
  let insideMath: boolean = false;
  let mathString: string = "";

  // showMessageThis?.(
  //   MessageType.debug,
  //   `source: ${source}`,
  //   "resolveEmphasis",
  //   false
  // );

  while ((result = rg.exec(source)) !== null) {
    // text
    let text = source.substring(indexBefore, result.index);
    if (text) {
      // showMessageThis?.(
      //   MessageType.debug,
      //   `text: ${text}`,
      //   "resolveEmphasis",
      //   false
      // );

      if (insideMath) {
        mathString = text;
      } else {
        stack.push(new TextRun({ text, ...textProp }));
      }
    }
    // tag
    text = source.substring(result.index + 1, rg.lastIndex - 1);
    // showMessageThis?.(
    //   MessageType.debug,
    //   `tag : ${text}`,
    //   "resolveEmphasis",
    //   false
    // );
    const tag = text.replace("/", "");
    const isOn = text === tag;

    if (tag === "codespan") {
      flushMathStringAsNormalString();
      if (isOn) {
        textProp["style"] = "codespan";
      } else {
        textProp["style"] = "";
      }
    } else if (tag === "$") {
      if (!insideMath) {
        insideMath = true;
      } else if (textProp.style !== "codespan" && mathString) {
        stack.push(await createMathImage(mathString));
        mathString = "";
        insideMath = false;
      } else {
        flushMathStringAsNormalString(tag);
      }
    } else {
      flushMathStringAsNormalString();
      textProp = { ...textProp, [resolveEmphasisTag(tag)]: isOn };
    }
    indexBefore = rg.lastIndex;

    // add checkbox
    if (tag === "☑" || tag === "☐") {
      if (documentInfo.params.useCheckBox) {
        stack.push(new CheckBox({ checked: tag === "☑" }));
      } else {
        let a = tag === "☑" ? "☑" : "☐";
        stack.push(new TextRun(a));
      }
    }
  }

  // text
  let text = source.substring(indexBefore);
  if (text) {
    // showMessageThis?.(
    //   MessageType.debug,
    //   `text: ${text}`,
    //   "resolveEmphasis",
    //   false
    // );
    stack.push(new TextRun({ text, ...textProp }));
  }
  return stack;

  /// inner function
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

  function flushMathStringAsNormalString(endDollar: string = "") {
    if (mathString) {
      stack.push(
        new TextRun({ text: "$" + mathString + endDollar, ...textProp })
      );
    }
    mathString = "";
    insideMath = false;
  }
}
