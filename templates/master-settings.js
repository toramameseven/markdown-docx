/**
 * modified next package samples.  [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
 */

// images
let imageBack = {
  path: "C:\\starlabs_bkgd.jpg",
};

let imageLogo = {
  path: "C:\\starlabs_logo.png",
  x: 4.6,
  y: 3.5,
  w: 4,
  h: 1.8,
};

// common
// FYI: use `headFontFace` and/or `bodyFontFace` to set the default font for the entire presentation (including slide Masters)
// const headFontFace = { headFontFace: "Meiryo" };
const headFontFace = { headFontFace: "Arial" };
const bodyFontFace = { bodyFontFace: "Arial" };
// pptx.theme = { bodyFontFace: "Arial" };
const layout = "LAYOUT_WIDE";
const backgroundColor = { color: "E3E3E3", transparency: 50 };
const mainColor = { color: "568FAC" };
const subColor = { color: "5495A8" };
const mainFontFace = { fontFace: "Arial" };
const mainFontSize = { fontSize: 18 };
const mainFontColor = { color: "0F0F0F" };
const subFontColor = { color: "0F0F0F" };
const mainLineSpacing = 0;
const fontBaseHeading = {
  ...mainFontFace,
  ...mainFontColor,
  lineSpacing: mainLineSpacing,
};
const fontBase = {
  ...mainFontFace,
  ...mainFontColor,
  lineSpacing: mainLineSpacing,
};
const tableHeaderColor = "000000";
const tableHeaderFillColor = "FFFFFF";
// header testPropsOption
const h1 = { ...fontBaseHeading, fontSize: 48, lineSpacing: 0 };
const h2 = { ...fontBaseHeading, fontSize: 24, lineSpacing: 0 };
const h3 = { ...fontBaseHeading, fontSize: 22, lineSpacing: 0 };
const h4 = { ...fontBaseHeading, fontSize: 20, lineSpacing: 0 };
const h5 = { ...fontBaseHeading, fontSize: 18, lineSpacing: 0 };
const h6 = {
  ...fontBaseHeading,
  ...mainFontSize,
  lineSpacing: mainLineSpacing,
};

const body = {
  ...fontBaseHeading,
  ...mainFontSize,
  lineSpacing: mainLineSpacing,
};

const code = {
  ...fontBase,
  ...mainFontSize,
  highlight: "FFFF00",
  lineSpacing: mainLineSpacing,
};

const codeSpan = {
  ...fontBase,
  ...mainFontSize,
  highlight: "FFFF00",
  lineSpacing: mainLineSpacing,
};

// TITLE_SLIDE
const titleSlide = {
  title: "TITLE_SLIDE",
  background: backgroundColor,
  //background: objBkg,
  //background: { color: "46e0e0", transparency: 50 },
  //bkgd: objBkg, // TEST: @deprecated
  objects: [
    //{ 'line':  { x:3.5, y:1.0, w:6.0, h:0.0, line:{color:'0088CC'}, lineSize:5 } },
    //{ 'chart': { type:'PIE', data:[{labels:['R','G','B'], values:[10,10,5]}], options:{x:11.3, y:0.0, w:2, h:2, dataLabelFontSize:9} } },
    //{ 'image': { x:11.3, y:6.4, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
    {
      placeholder: {
        options: {
          name: "title",
          type: "title",
          x: "10%",
          y: "40%",
          w: "80%",
          h: "20%",
          margin: 0,
          align: "middle",
          valign: "middle",
          color: "404040",
          fontSize: 36,
          ...mainFontFace,
        },
        text: "", // USAGE: Leave blank to have powerpoint substitute default placeholder text (ex: "Click to add title")
      },
    },
    {
      placeholder: {
        options: {
          name: "subtitle",
          type: "title",
          x: "10%",
          y: "50%",
          w: "80%",
          h: "20%",
          margin: 0,
          align: "middle",
          valign: "middle",
          fontSize: 18,
          ...fontBase,
        },
        text: "", // USAGE: Leave blank to have powerpoint substitute default placeholder text (ex: "Click to add title")
      },
    },
    {
      rect: { x: 0.0, y: 5.7, w: "100%", h: 0.75, fill: { color: "F1F1F1" } },
    },
    {
      text: {
        text: "Global IT & Services :: Status Report",
        options: {
          x: 0.0,
          y: 5.7,
          w: "100%",
          h: 0.75,
          color: "363636",
          fontSize: 20,
          align: "center",
          valign: "middle",
          margin: 0,
          ...fontBase,
        },
      },
    },
  ],
};

// MASTER_SLIDE (MASTER_PLACEHOLDER)
const masterSlide = {
  title: "MASTER_SLIDE",
  background: backgroundColor,
  margin: [0.5, 0.25, 1.0, 0.25],
  slideNumber: {
    x: 0.6,
    y: 7.1,
    color: "FFFFFF",
    fontSize: 10,
    bold: true,
    ...fontBase,
  },
  objects: [
    //{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
    {
      rect: {
        x: 0.0,
        y: "90%",
        w: "100%",
        h: "10%",
        fill: { color: "003b75" },
      },
    },
    {
      text: {
        options: {
          x: 0,
          y: "90%",
          w: "100%",
          h: "10%",
          align: "center",
          valign: "middle",
          color: "FFFFFF",
          fontSize: 18,
          ...fontBase,
        },
        text: "S.T.A.R. Laboratories - Confidential",
      },
    },
    {
      placeholder: {
        options: {
          name: "header",
          type: "title",
          x: 0.6,
          y: 0.2,
          w: 12,
          h: 1.0,
          margin: 0,
          align: "middle",
          valign: "top",
          color: "404040",
          fontSize: 18,
          ...fontBase,
        },
        text: "", // USAGE: Leave blank to have powerpoint substitute default placeholder text (ex: "Click to add title")
      },
    },
    // {
    //   placeholder: {
    //     options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12, h: 5.25, fontSize: 28 },
    //     text: "(supports custom placeholder text!)",
    //   },
    // },
  ],
};

const tableProps = {
  x: 0,
  y: "30%",
  w: "100%",
  rowH: 0.75,
  fill: backgroundColor,
  color: "000000",
  valign: "middle",
  align: "left",
  border: { type: "solid", pt: 1, color: "000000" },
  ...mainFontFace,
  ...mainFontSize,
  ...mainFontColor,
};

const tablePropsArray = [
  {
    x: 0,
    y: "30%",
    w: "100%",
    rowH: 0.75,
    fill: backgroundColor,
    color: "000000",
    valign: "middle",
    align: "left",
    border: { type: "solid", pt: 1, color: "000000" },
    ...mainFontFace,
    ...mainFontSize,
    ...mainFontColor,
  },
];

/** "x,y,w,h" in percent */
const defaultPositionPCT = "10,15,80,70";

module.exports = {
  titleSlide: titleSlide,
  masterSlide: masterSlide,
  h1,
  h2,
  h3,
  h4,
  h5,
  h6,
  body,
  code,
  codeSpan,
  tableProps,
  layout,
  bodyFontFace,
  headFontFace,
  defaultPositionPCT,
  tablePropsArray,
};
