/**
 * clear.ppt.js
 * modified next package samples.  [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
 */

/**
 * @param {string} s
 * @returns {string}
 */
function s2color(s) {
  return s.substring(1);
}

// images
let imageBack = {
  path: "C:\\background.png",
};

let imageLogo = {
  path: "C:\\logo.png",
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

const layout = "LAYOUT_WIDE";
//const backgroundColor = { color: s2color("#E3E3E3E3"), transparency: 50 };

const mainColor = {color: s2color("#4469A6")};
const accentColor = {color: s2color("#6A8CC7")};
const codeHighlightColor = s2color("#FFFF00");

const mainFontFace = { fontFace: "Arial" };
const mainFontSize = { fontSize: 18 };
const mainFontColor =  mainColor;
const subFontColor = { color: s2color("#0F0F0F") };
const backgroundColor = { color: s2color("#E4E0BE")};
const mainLineSpacing = 0;

const fontBaseHeading = {
  ...mainFontFace,
  ...mainFontColor,
  lineSpacing: mainLineSpacing,
};

const fontBaseBody = {
  ...mainFontFace,
  ...mainFontColor,
  lineSpacing: mainLineSpacing,
};

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
  ...fontBaseBody,
  ...mainFontSize,
  highlight: codeHighlightColor,
  lineSpacing: mainLineSpacing,
};

const codeSpan = {
  ...fontBaseBody,
  ...mainFontSize,
  highlight: codeHighlightColor,
  lineSpacing: mainLineSpacing,
};

// TITLE_SLIDE
const titleSlide = {
  title: "TITLE_SLIDE",
  background: backgroundColor,
  //background: imageBack,
  //background: { color: "46e0e0", transparency: 50 },
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
          fontSize: 36,
          ...mainFontFace,
          ...mainFontColor,
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
          ...fontBaseBody,
        },
        text: "", // USAGE: Leave blank to have powerpoint substitute default placeholder text (ex: "Click to add title")
      },
    },
    {
      rect: { x: 0.0, y: 5.7, w: "100%", h: 0.75, fill: mainColor },
    },
    {
      text: {
        text: "Your Laboratories on the Title",
        options: {
          x: 0.0,
          y: 5.7,
          w: "100%",
          h: 0.75,
          fontSize: 20,
          align: "center",
          valign: "middle",
          margin: 0,
          ...fontBaseBody,
          color: s2color("#363636"),
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
    color: s2color("#FFFFFF"),
    fontSize: 10,
    bold: true,
    ...fontBaseBody,
  },
  objects: [
    //{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
    {
      rect: {
        x: 0.0,
        y: "90%",
        w: "100%",
        h: "10%",
        fill: mainColor,
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
          fontSize: 18,
          ...fontBaseBody,
          color: s2color("#FFFFFF"),
        },
        text: "Your Laboratories - Confidential",
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
          color: s2color("#404040"),
          fontSize: 18,
          ...fontBaseBody,
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
  ...mainColor,
  valign: "middle",
  align: "left",
  border: { type: "solid", pt: 1, ...accentColor },
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
    color: s2color("#000000"),
    valign: "middle",
    align: "left",
    border: { type: "solid", pt: 1, color: s2color("#FFFFFF") },
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
