

let objBkg = {
  path: "C:\\Users\\maru\\Desktop\\github\\markdown-docx\\src\\markdown-pptx\\starlabs_bkgd.jpg",
};

let objImg = {
  path: "C:\\Users\\maru\\Desktop\\github\\markdown-docx\\src\\markdown-pptx\\starlabs_logo.png",
  x: 4.6,
  y: 3.5,
  w: 4,
  h: 1.8,
};

// header testPropsOption
const mainFont = {fontFace: "Meiryo"};
const h1 = { ...mainFont, fontSize: 48 };
const h2 = { ...mainFont, fontSize: 42 };
const h3 = { ...mainFont, fontSize: 36 };
const h4 = { ...mainFont, fontSize: 30 };
const h5 = { ...mainFont, fontSize: 24 };
const h6 = { ...mainFont, fontSize: 18 };

const body = { ...mainFont, fontSize: 18 };

// TITLE_SLIDE
const titleSlide = {
  title: "TITLE_SLIDE",
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
          ...mainFont
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
          fontFace: "Arial",
          color: "363636",
          fontSize: 20,
          align: "center",
          valign: "middle",
          margin: 0,
        },
      },
    },
  ],
};

// MASTER_SLIDE (MASTER_PLACEHOLDER)
const masterSlide = {
  title: "MASTER_SLIDE",
  background: { color: "46e0e0", transparency: 50 },
  // background: { color: "E1E1E1", transparency: 50 },
  margin: [0.5, 0.25, 1.0, 0.25],
  slideNumber: {
    x: 0.6,
    y: 7.1,
    color: "FFFFFF",
    fontFace: "Arial",
    fontSize: 10,
    bold: true,
  },
  objects: [
    //{ 'image': { x:11.45, y:5.95, w:1.67, h:0.75, data:STARLABS_LOGO_SM } },
    {
      rect: { x: 0.0, y: "90%", w: "100%", h: "10%", fill: { color: "003b75" } },
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
          fontSize: 12,
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
          //fontSize: 18,
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

// THANKS_SLIDE (THANKS_PLACEHOLDER)
const thanksSlide = {
  title: "THANKS_SLIDE",
  background: { color: "36ABFF" }, // CORRECT WAY TO SET BACKGROUND COLOR
  //bkgd: "36ABFF", // [[BACKWARDS-COMPAT/DEPRECATED/UAT (`bkgd` will be removed in v4.x)]] **DO NOT USE THIS IN YOUR CODE**
  objects: [
    {
      rect: { x: 0.0, y: 3.4, w: "100%", h: 2.0, fill: { color: "FFFFFF" } },
    },
    //{ image: objImg },
    {
      placeholder: {
        options: {
          name: "thanksText",
          type: "title",
          x: 0.0,
          y: 0.9,
          w: "100%",
          h: 1,
          fontFace: "Arial",
          color: "FFFFFF",
          fontSize: 60,
          align: "center",
        },
      },
    },
    {
      placeholder: {
        options: {
          name: "body",
          type: "body",
          x: 0.0,
          y: 6.45,
          w: "100%",
          h: 1,
          fontFace: "Courier",
          color: "FFFFFF",
          fontSize: 32,
          align: "center",
        },
        text: "(add homepage URL)",
      },
    },
  ],
};

// style samples

const BASE_TABLE_OPTS = { x: 0.5, y: 0.13, colW: [9, 3.33] }; // LAYOUT_WIDE w=12.33

// STYLES
const BKGD_LTGRAY = "F1F1F1";
const COLOR_BLUE = "0088CC";
const CODE_STYLE = {
  fill: { color: BKGD_LTGRAY },
  margin: 6,
  fontSize: 10,
  color: "696969",
};
const TITLE_STYLE = {
  fill: { color: BKGD_LTGRAY },
  margin: 4,
  fontSize: 18,
  fontFace: "Segoe UI",
  color: COLOR_BLUE,
  valign: "top",
  align: "center",
};

// OPTIONS
const BASE_TEXT_OPTS_L = {
  color: "9F9F9F",
  margin: 3,
  border: [null, null, { pt: "1", color: "CFCFCF" }, null],
};
const BASE_TEXT_OPTS_R = {
  text: "PptxGenJS",
  options: {
    color: "9F9F9F",
    margin: 3,
    border: [0, 0, { pt: "1", color: "CFCFCF" }, 0],
    align: "right",
  },
};
const FOOTER_TEXT_OPTS = {
  x: 0.0,
  y: 7.16,
  w: "100%",
  h: 0.3,
  margin: 3,
  color: "9F9F9F",
  align: "center",
  fontSize: 10,
};
const BASE_CODE_OPTS = {
  color: "9F9F9F",
  margin: 3,
  border: { pt: "1", color: "CFCFCF" },
  fill: { color: "F1F1F1" },
  fontFace: "Courier",
  fontSize: 12,
};
const BASE_OPTS_SUBTITLE = {
  x: 0.5,
  y: 0.7,
  w: 4,
  h: 0.3,
  fontSize: 18,
  fontFace: "Arial",
  color: "0088CC",
  fill: { color: "FFFFFF" },
};
const DEMO_TITLE_TEXT = { fontSize: 14, color: "0088CC", bold: true };
const DEMO_TITLE_TEXTBK = {
  fontSize: 14,
  color: "0088CC",
  bold: true,
  breakLine: true,
};


module.exports = {
  titleSlide: titleSlide,
  masterSlide: masterSlide,
  thanksSlide: thanksSlide,
  h1,
  h2,
  h3,
  h4,
  h5,
  h6,
  body,
};
