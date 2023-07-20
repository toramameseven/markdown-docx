import pptxGen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";

export type PptStyle = {
  titleSlide: PptxGenJS.SlideMasterProps;
  masterSlide: PptxGenJS.SlideMasterProps;
  thanksSlide: PptxGenJS.SlideMasterProps;
  h1: PptxGenJS.TextPropsOptions;
  h2: PptxGenJS.TextPropsOptions;
  h3: PptxGenJS.TextPropsOptions;
  h4: PptxGenJS.TextPropsOptions;
  h5: PptxGenJS.TextPropsOptions;
  h6: PptxGenJS.TextPropsOptions;
  body: PptxGenJS.TextPropsOptions;
};

export type TextFrame = {
  textPropsArray: PptxGenJS.TextProps[];
  outputPosition: {};
};

const _sp = "\t";
//
export const WdNodeType = {
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
  tablePos: "tablePos",
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

export type WdNodeType = (typeof WdNodeType)[keyof typeof WdNodeType];

export const DocxStyle = {
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
export type DocxStyle = (typeof DocxStyle)[keyof typeof DocxStyle];

export type TableProps = {
  tableRows: pptxGen.TableRow[];
  options?: pptxGen.TableProps;
};

export const SheetObjectType = {
  text: "text",
  table: "table",
  image: "image",
} as const;
export type SheetObjectType =
  (typeof SheetObjectType)[keyof typeof SheetObjectType];

export type SheetObject = {
  type: SheetObjectType;
  sheetObject: pptxGen.ImageProps | TableProps | TextFrame;
};
/**
 *
 */
export class PptSheet {
  sheetObjects: SheetObject[] = [];

  slide?: pptxGen.Slide;
  currentTextPropsArray: PptxGenJS.TextProps[] = [];
  currentTextPropPosition: {} = {};
  defaultTextPropPosition: {} = {};
  pptx: PptxGenJS;

  constructor(pptx: PptxGenJS) {
    this.pptx = pptx;
  }

  addTitleSlide(documentInfo: { title?: string }) {
    this.slide = this.pptx.addSlide({ masterName: "TITLE_SLIDE" });
    this.slide.addText(documentInfo.title!, {
      placeholder: "title",
    });
  }

  addMasterSlide() {
    this.slide = this.pptx.addSlide({ masterName: "MASTER_SLIDE" });
    this.clear();
    this.currentTextPropPosition = this.defaultTextPropPosition;
    // add position
  }

  addHeader(
    textPropsArray: pptxGen.TextProps[],
    textPropsOptions: pptxGen.TextPropsOptions
  ) {
    this.slide!.addText(textPropsArray, {
      placeholder: "header",
      ...textPropsOptions,
    });
  }

  addImage(image: pptxGen.ImageProps) {
    this.sheetObjects.push({ type: "image", sheetObject: image });
  }

  addTable(table: TableProps) {
    this.sheetObjects.push({ type: "table", sheetObject: table });
  }

  addTextFrame(textPosition: {} = {}) {
    const sheetObject = {
      textPropsArray: this.currentTextPropsArray,
      outputPosition: { ...this.currentTextPropPosition, ...textPosition },
    };
    this.sheetObjects.push({ type: "text", sheetObject });
    this.currentTextPropsArray = [];
  }

  addTextPropsArray(...arr: PptxGenJS.TextProps[]) {
    this.currentTextPropsArray.push(...arr);
  }

  setDefaultPosition(position: {}) {
    this.defaultTextPropPosition = { ...position };
    this.currentTextPropPosition = { ...position };
  }
  setCurrentPosition(position: {}) {
    this.currentTextPropPosition = {
      ...this.defaultTextPropPosition,
      ...position,
    };
  }

  createSheet() {
    if (this.slide) {
      this.sheetObjects.forEach((x) => {
        switch (x.type) {
          case "text":
            const o = x.sheetObject as TextFrame;
            this.slide!.addText(o.textPropsArray, o.outputPosition);
            break;
          case "table":
            const t = x.sheetObject as TableProps;
            this.slide!.addTable(t.tableRows, t.options);
            break;
          case "image":
            const i = x.sheetObject as pptxGen.ImageProps;
            this.slide!.addImage(i);            
            break;
          default:
            break;
        }
      });
    }
  }

  clear() {
    this.sheetObjects = [];
  }
}

/**
 *
 */
export class TableJs {
  rows: number;
  columns: number;
  // x, y, list();;
  cells: PptxGenJS.TextProps[][][];
  row: number;
  column: number;
  mergedCells: number[][][];
  tableWidthArray: number[];
  mdSourcePath: string;
  tablePosition = {};

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
    this.tablePosition = getPositionPCT("10,10,80,80");
  }

  doTableCommand(line: string) {
    const words = line.split(_sp);
    switch (words[0]) {
      case "tablePos":
        this.tablePosition = getPositionPCT(words[1]);
        return;
        break;
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

  createTable() {
    let rows: pptxGen.TableRow[] = new Array(this.rows);
    for (let i = 0; i < this.rows; i++) {
      rows[i] = new Array<PptxGenJS.TableCell>(0);
      for (let j = 0; j < this.columns; j++) {
        let thisWidth = {
          size: this.tableWidthArray[j],
          //type: pptxGen .PERCENTAGE,
        };

        if (this.tableWidthArray.length > 0) {
          thisWidth = {
            size: this.tableWidthArray[j],
            //type: WidthType.PERCENTAGE,
          };
        }

        let pCell = this.cells[i][j];
        const [rowspan, colspan] = this.createRawColumnSpan(
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
          //rows[i][j] = null;
        } else {
          //rows[i][j] = {...tCell, options:{rowspan, colspan}};

          rows[i].push({ ...tCell, options: { rowspan, colspan } });
        }
      }
    }

    const options: pptxGen.TableProps = {
      x: 0,
      y: 2.1,
      w: "100%",
      rowH: 0.75,
      fill: { color: "F7F7F7" },
      color: "000000",
      fontSize: 16,
      valign: "middle",
      align: "center",
      border: { type: "solid", pt: 1, color: "000000" },
      ...this.tablePosition,
    };

    // { tableRows: pptxGen.TableRow[]; options?: pptxGen.TableProps }
    return { tableRows: rows, options };
    // slide.addTable(rows, {
    //   x: 0,
    //   y: 2.1,
    //   w: "100%",
    //   rowH: 0.75,
    //   fill: { color: "F7F7F7" },
    //   color: "000000",
    //   fontSize: 16,
    //   valign: "middle",
    //   align: "center",
    //   border: { type: "solid", pt: 1, color: "000000" },
    // });
  }

  createRawColumnSpan(r1: number, c1: number, r2: number, c2: number) {
    return [r2 - r1 + 1, c2 - c1 + 1];
  }
}

export class DocParagraph {
  nodeType: WdNodeType;
  isFlush: boolean;
  indent: number;
  // export interface TextProps {
  // 	text?: string
  // 	options?: TextPropsOptions
  // }
  children: PptxGenJS.TextProps[] = [];
  textPropsOptions: PptxGenJS.TextPropsOptions;
  docxStyle: DocxStyle;
  isImage: boolean;
  isNewSheet: boolean = false;
  currentFontSize: number = 24;

  constructor(
    nodeType: WdNodeType = WdNodeType.non,
    textPropsOptions: PptxGenJS.TextPropsOptions = {},
    indent: number = 0,
    docStyle: DocxStyle = DocxStyle.Body,
    child?: PptxGenJS.TextProps
  ) {
    this.nodeType = nodeType;
    this.isFlush = false;
    this.textPropsOptions = textPropsOptions;
    this.indent = indent;
    this.children = child ? [child] : [];
    this.docxStyle = docStyle;
    this.isImage = false;
  }

  createTextPropsArray(): PptxGenJS.TextProps[] {
    let pStyle = this.isImage ? "picture1" : this.docxStyle;

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

export function getPositionPCT(position: string) {
  const positions = position.split(",");
  return {
    x: `${positions[0]}%`,
    y: `${positions[1]}%`,
    w: `${positions[2]}%`,
    h: `${positions[3]}%`,
  };
}

/**
 *
 * @param source
 * @returns
 */
export function resolveEmphasis(source: string) {
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
    //valign: "top",
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
