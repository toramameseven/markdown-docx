// top word commands
export const wordCommand = {
  cols: "cols",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",

  newPage: "newPage",
  newLine: "newLine",
  toc: "toc",
  export: "export",

  placeholder: "placeholder",
  param: "param",
  tableParam: "tableParam"
} as const;


// https://chaika.hatenablog.com/entry/2021/11/03/083000
// <!-- word param|placeholder|tableParam params -->
export const documentInfoParams = [
  /** */
  "docxTemplate",
  "tablePrefix",
  "figurePrefix",
  "levelOffset",
  "refFormat",
  "captionRefFormat",

  "useCheckBox",

  "cols",
  "tableStyle",
  "rowMerge",
  "emptyMerge",
  "tableWidth",
  "tableAlign",
  "tableCaption",
  "tableCaptionId",

  "imageWidth",

  "pptxSettings",
  "position",
  "dpi",
] as const;
export type DocumentInfoParams = (typeof documentInfoParams)[number];




