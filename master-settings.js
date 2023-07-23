// images
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

// common
const mainFontFace = { fontFace: "Meiryo" };
const mainFontSize = {fontSize: 18 };

// header testPropsOption
const h1 = { ...mainFontFace, fontSize: 48 };
const h2 = { ...mainFontFace, fontSize: 42 };
const h3 = { ...mainFontFace, fontSize: 36 };
const h4 = { ...mainFontFace, fontSize: 30 };
const h5 = { ...mainFontFace, fontSize: 24 };
const h6 = { ...mainFontFace, ...mainFontSize };
const body = { ...mainFontFace, ...mainFontSize };

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
          color: "404040",
          fontSize: 18,
          ...mainFontFace,
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

const tableProps = {
  x: 0,
  y: "30%",
  w: "100%",
  rowH: 0.75,
  fill: { color: "F7F7F7" },
  color: "000000",
  valign: "middle",
  align: "center",
  border: { type: "solid", pt: 1, color: "000000" },
  ...mainFontFace,
  ...mainFontSize
};

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
  tableProps,
};
