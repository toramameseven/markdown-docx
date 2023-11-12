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
} as const;


// https://chaika.hatenablog.com/entry/2021/11/03/083000
// <!-- word param|placeholder params -->
export const documentInfoParams = [
  "pptxSettings",
  "position",
  "dpi",
  "docxTemplate",
  "refFormat",
  "captionRefFormat",
  "cols",
  "tableStyle",
  "rowMerge",
  "emptyMerge",
  "tableWidth",
  "tableAlign",
  "imageWidth",
  "levelOffset",
  "tableCaption",
  "tableCaptionId",
  "tablePrefix",
  "figurePrefix",
] as const;
export type DocumentInfoParams = (typeof documentInfoParams)[number];




