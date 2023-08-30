import pptxGen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";

export type PptStyle = {
  titleSlide: PptxGenJS.SlideMasterProps;
  masterSlide: PptxGenJS.SlideMasterProps;
  h1: PptxGenJS.TextPropsOptions;
  h2: PptxGenJS.TextPropsOptions;
  h3: PptxGenJS.TextPropsOptions;
  h4: PptxGenJS.TextPropsOptions;
  h5: PptxGenJS.TextPropsOptions;
  h6: PptxGenJS.TextPropsOptions;
  body: PptxGenJS.TextPropsOptions;
  code: PptxGenJS.TextPropsOptions;
  codeSpan: PptxGenJS.TextPropsOptions;
  tableProps: PptxGenJS.TableProps;
  tableHeaderColor: string;
  tableHeaderFillColor: string;
  layout: string;
  bodyFontFace: PptxGenJS.TableCellProps;
  headFontFace: PptxGenJS.ThemeProps;
};

export type TextFrame = {
  textPropsArray: PptxGenJS.TextProps[];
  outputPosition: {};
};

const _sp = "\t";

export { wdCommand, WdCommand } from "./wd0-to-wd";

export const docxStyle = {
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
export type DocxStyle = (typeof docxStyle)[keyof typeof docxStyle];

export type TableProps = {
  tableRows: pptxGen.TableRow[];
  options?: pptxGen.TableProps;
};

export const sheetObjectType = {
  text: "text",
  table: "table",
  image: "image",
  shape: "shape",
} as const;
export type SheetObjectType =
  (typeof sheetObjectType)[keyof typeof sheetObjectType];

export type SheetObject = {
  type: SheetObjectType;
  sheetObject:
    | pptxGen.ImageProps
    | TableProps
    | TextFrame
    | pptxGen.TextPropsOptions;
};

/**
 *
 */
export class PptSheet {
  sheetObjects: SheetObject[] = [];
  slide?: pptxGen.Slide;
  currentTextPropsArray: PptxGenJS.TextProps[] = [];
  currentTextPropPositionPCT: {} = {};
  defaultTextPropPositionPCT: {} = {};
  pptx: PptxGenJS;
  pptxParagraph: PptParagraph;
  pptStyle: PptStyle;

  constructor(pptx: PptxGenJS, pptStyle: PptStyle) {
    this.pptx = pptx;
    this.pptStyle = pptStyle;
    this.pptxParagraph = new PptParagraph(
      pptStyle.body.fontSize ?? 0,
      pptStyle.body.lineSpacing ?? 0
    );
  }

  addTitleSlide(documentInfo: { [v: string]: string }) {
    this.slide = this.pptx.addSlide({ masterName: "TITLE_SLIDE" });
    this.slide.addText(documentInfo.title!, {
      placeholder: "title",
    });
    this.slide.addText(documentInfo.subTitle!, {
      placeholder: "subtitle",
    });
  }

  addDocumentSlide() {
    this.slide = this.pptx.addSlide({ masterName: "MASTER_SLIDE" });
    this.sheetObjects = [];
    this.currentTextPropPositionPCT = this.defaultTextPropPositionPCT;
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

  flushShapes() {
    const shapes = this.pptxParagraph.createTextCode();
    shapes.forEach((s) =>
      this.sheetObjects.push({ type: "shape", sheetObject: s })
    );
  }

  addTable(table: TableProps) {
    this.sheetObjects.push({ type: "table", sheetObject: table });
  }

  addTextFrame(textPosition: {} = {}) {
    if (this.currentTextPropsArray.length === 0) {
      return;
    }
    const sheetObject = {
      textPropsArray: this.currentTextPropsArray,
      outputPosition: { ...this.currentTextPropPositionPCT, ...textPosition },
    };
    this.sheetObjects.push({ type: "text", sheetObject });
    this.currentTextPropsArray = [];
  }

  addTextPropsArray() {
    const textPropsArray = this.pptxParagraph.createTextPropsArray();
    if (textPropsArray.length) {
      this.currentTextPropsArray.push(...textPropsArray);
    }
  }

  setDefaultPositionPCT(position: {}) {
    this.defaultTextPropPositionPCT = { ...position };
    this.setCurrentPositionPCT(position);
  }

  setCurrentPositionPCT(position: {}) {
    this.currentTextPropPositionPCT = {
      ...this.defaultTextPropPositionPCT,
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
          case "shape":
            const s = x.sheetObject as pptxGen.TextPropsOptions;

            // let sss = `{
            //   "objectName": "山本山",
            //   "shape": "rect",
            //   "x":0.5,
            //   "y":0.8,
            //   "w":1.5,
            //   "h":3.0,
            //   "fill":{ "color": "555555" },
            //   "align":"center",
            //   "fontSize":14
            // }`;

            const sname = s.objectName ?? "";
            delete s.objectName;
            this.slide!.addText(sname ?? "", s);
            break;
          default:
            break;
        }
      });
    }
  }
}

export class PptParagraph {
  isFlush: boolean = false;
  indent: number = 0;

  // export interface TextProps {
  // 	text?: string
  // 	options?: TextPropsOptions
  // }

  children: PptxGenJS.TextProps[] = [];
  childrenRaw: string[] = [];
  textPropsOptions: PptxGenJS.TextPropsOptions = {};
  codeLang: string = "";
  isNewSheet: boolean = false;
  defaultFontSize: number = 18;
  currentFontSize: number = 18;
  defaultLineSpacing: number = 0;
  currentLineSpacing: number = 0;
  insideSlideTitle: boolean = false;
  insideDocumentTitle: boolean = false;

  constructor(defaultFontSize: number, defaultLineSpacing: number) {
    this.defaultFontSize = defaultFontSize;
    this.currentFontSize = defaultFontSize;
    this.defaultLineSpacing = defaultLineSpacing;
    this.currentLineSpacing = defaultLineSpacing;
  }

  createTextPropsArray(): PptxGenJS.TextProps[] {
    if (this.children.length === 0) {
      return [];
    }
    const r = this.children.map((p) => {
      return {
        text: p.text,
        options: { ...this.textPropsOptions, ...p.options },
      };
    });
    this.children = [];
    this.childrenRaw = [];
    this.isFlush = false;
    this.codeLang = "";
    this.currentFontSize = this.defaultFontSize;
    this.currentLineSpacing = this.defaultLineSpacing;
    return r;
  }

  createTextCode() {
    const shapeJson: pptxGen.TextPropsOptions[] = JSON.parse(
      this.childrenRaw.join("")
    );

    const r = shapeJson;
    for (let i = 1; i < r.length; i++) {
      if (Object.keys(r[i]).length === 0) {
        // clear
        //rOut.push(r[i]);
      } else {
        // 
        r[i] = ({ ...r[i - 1], ...r[i] });
      }
    }

    this.children = [];
    this.childrenRaw = [];
    this.isFlush = false;
    this.codeLang = "";
    this.currentFontSize = this.defaultFontSize;
    this.currentLineSpacing = this.defaultLineSpacing;
    // remove empty obj
    return r.filter(ro =>(Object.keys(ro).length));
  }

  clear() {
    this.children = [];
    this.isFlush = false;
    this.currentFontSize = this.defaultFontSize;
    this.currentLineSpacing = this.defaultLineSpacing;
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

  addChild(s: string | PptxGenJS.TextProps) {
    if (typeof s === "string") {
      this.children.push({ text: s });
      this.childrenRaw.push(s);
    } else {
      this.children.push(s);
    }
  }
  addChildRaw(s: string) {
    this.childrenRaw.push(s);
  }
  addChildren(s: PptxGenJS.TextProps[]) {
    this.children.push(...s);
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

  doTableCommand(line: string, pptStyle: PptStyle) {
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
          this.cells[this.row][this.column].push(
            ...resolveEmphasis(words[2], pptStyle)
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

  createTable(pos: {}, tableProps?: PptxGenJS.TableProps) {
    let rows: pptxGen.TableRow[] = new Array(this.rows);
    for (let i = 0; i < this.rows; i++) {
      rows[i] = new Array<PptxGenJS.TableCell>(0);
      for (let j = 0; j < this.columns; j++) {
        // let thisWidth = {
        //   size: this.tableWidthArray[j],
        //   //type: pptxGen .PERCENTAGE,
        // };

        // if (this.tableWidthArray.length > 0) {
        //   thisWidth = {
        //     size: this.tableWidthArray[j],
        //     //type: WidthType.PERCENTAGE,
        //   };
        // }

        let pCell = this.cells[i][j];
        const [rowspan, colspan] = this.createRawColumnSpan(
          i,
          j,
          this.mergedCells[i][j][0],
          this.mergedCells[i][j][1]
        );

        let tCell = {
          text: pCell[0] ? pCell : [{ text: "" }],
        };

        // console.log(
        //   `${i}-${j}, ${pCell[0]}: ${this.mergedCells[i][j][0]}: ${this.mergedCells[i][j][1]}`
        // );

        if (
          this.mergedCells[i][j][0] === -1 &&
          this.mergedCells[i][j][1] === -1
        ) {
          //rows[i][j] = null;
        } else {
          const fillColor = i === 0 ? "676767" : "F7F7F7";
          const textColor = i === 0 ? "F7F7F7" : "676767";
          rows[i].push({
            ...tCell,
            options: {
              rowspan,
              colspan,
              valign: "middle",
              color: textColor,
              fill: { color: fillColor },
            },
          });
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
      ...tableProps,
      ...pos,
    };

    // { tableRows: pptxGen.TableRow[]; options?: pptxGen.TableProps }
    return { tableRows: rows, options };
  }

  createRawColumnSpan(r1: number, c1: number, r2: number, c2: number) {
    return [r2 - r1 + 1, c2 - c1 + 1];
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
export function resolveEmphasis(source: string, pptStyle: PptStyle) {
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
        options = { ...options, highlight: "FF88CC", ...pptStyle.codeSpan };
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
