/**
 * enums.mjs
 * global variables
 */

import PptxGenJS from "pptxgenjs";

// LIBRARY
export const COMPRESS = true; // TEST: `compression` write prop

// CONST
export const CUST_NAME = "S.T.A.R. Laboratories";
export const USER_NAME = "Barry Allen";
export const ARRSTRBITES = [130];
export const CHARSPERLINE = 130; // "Open Sans", 13px, 900px-colW = ~19 words/line ~130 chars/line

// TABLES
export const BASE_TABLE_OPTS = { x: 0.5, y: 0.13, colW: [9, 3.33] }; // LAYOUT_WIDE w=12.33

// STYLES
export const BKGD_LTGRAY: PptxGenJS.Color = "F1F1F1";
export const COLOR_BLUE: PptxGenJS.Color = "0088CC";
export const CODE_STYLE: PptxGenJS.TextPropsOptions = {
  fill: { color: BKGD_LTGRAY },
  margin: 6,
  fontSize: 10,
  color: "696969",
};

export const header1: PptxGenJS.TextPropsOptions = {
  fill: { color: BKGD_LTGRAY },
  margin: 4,
  fontSize: 18,
  fontFace: "Segoe UI",
  color: COLOR_BLUE,
  valign: "top",
  align: "left",
};


export const TITLE_STYLE: PptxGenJS.TextPropsOptions = {
  fill: { color: BKGD_LTGRAY },
  margin: 4,
  fontSize: 18,
  fontFace: "Segoe UI",
  color: COLOR_BLUE,
  valign: "top",
  align: "center",
};

// OPTIONS
export const BASE_TEXT_OPTS_L: PptxGenJS.TableCellProps = {
  color: "9F9F9F",
  margin: 3,
  border: [
    { pt: 0, color: "CFCFCF" },
    { pt: 0, color: "CFCFCF" },
    { pt: 1, color: "CFCFCF" },
    { pt: 0, color: "CFCFCF" },
  ],
};

export const BASE_TEXT_OPTS_R: PptxGenJS.TextProps = {
  text: "PptxGenJS",
  options: {
    color: "9F9F9F",
    margin: 3,
    align: "right",
  },
};

export const FOOTER_TEXT_OPTS: PptxGenJS.TextPropsOptions = {
  x: 0.0,
  y: 7.16,
  w: "100%",
  h: 0.3,
  margin: 3,
  color: "9F9F9F",
  align: "center",
  fontSize: 10,
};

export const BASE_CODE_OPTS: PptxGenJS.TextPropsOptions = {
  color: "9F9F9F",
  margin: 3,
  //border: { pt: 1, color: "CFCFCF" },
  fill: { color: "F1F1F1" },
  fontFace: "Courier",
  fontSize: 12,
};

export const BASE_OPTS_SUBTITLE: PptxGenJS.TextPropsOptions = {
  x: 0.5,
  y: 0.7,
  w: 4,
  h: 0.3,
  fontSize: 18,
  fontFace: "Arial",
  color: "0088CC",
  fill: { color: "FFFFFF" },
};

export const DEMO_TITLE_TEXT: PptxGenJS.TextPropsOptions = {
  fontSize: 14,
  color: "0088CC",
  bold: true,
};

export const DEMO_TITLE_TEXTBK: PptxGenJS.TextPropsOptions = {
  fontSize: 14,
  color: "0088CC",
  bold: true,
  breakLine: true,
};

export const DEMO_TITLE_OPTS: PptxGenJS.TextPropsOptions = {
  fontSize: 13,
  color: "9F9F9F",
};
