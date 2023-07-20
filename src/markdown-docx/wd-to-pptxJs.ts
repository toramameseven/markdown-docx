/**
 * https://gitbrent.github.io/PptxGenJS/
 */
import pptxGen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import * as Path from "path";
import * as imageSize from "image-size";
const TeXToSVG = require("tex-to-svg");
import {
  DocParagraph,
  DocxStyle,
  PptSheet,
  PptStyle,
  TableJs,
  TextFrame,
  WdNodeType,
  resolveEmphasis,
} from "./pptxjs";

import { svg2imagePng } from "../tools/svg-png-image";
import { runCommand, selectExistsPath } from "../tools/tools-common";
import {
  DocxOption,
  MessageType,
  ShowMessage,
  getFileContents,
} from "./common";

let pptStyle: PptStyle = {
  titleSlide: { title: "" },
  masterSlide: { title: "" },
  thanksSlide: { title: "" },
  h1: {},
  h2: {},
  h3: {},
  h4: {},
  h5: {},
  h6: {},
  body: {},
};

const _sp = "\t";
//

// ############################################################
export async function wdToPptx(
  fileWd: string,
  wdBody: string,
  option: DocxOption
) {
  let body = wdBody;

  if (body === "") {
    body = getFileContents(fileWd);
  }

  try {
    await wdToPptxJs(body, Path.dirname(fileWd));
  } catch (e) {
    option.message?.(
      MessageType.warn,
      `wdToPptxJs err: ${e}.`,
      "wd-to-pptx",
      false
    );
    return;
  }
  return;
}

/**
 *
 * @param body
 * @param mdSourcePath
 */
export async function wdToPptxJs(
  body: string,
  mdSourcePath: string
): Promise<void> {
  // doc info
  const documentInfo = {
    title: "",
    subTitle: "",
    division: "",
    date: "",
    author: "",
    docNumber: "",
    position: "",
    pptxSettings: "",
  };

  // get lines
  const lines = (body + "\nEndline").split(/\r?\n/);

  // get document information
  const toInfoSearch = lines.length > 100 ? 100 : lines.length;
  for (let i = 0; i < toInfoSearch; i++) {
    const wdCommandList = lines[i].split(_sp);
    // html comment command <!-- word xxxx -->
    resolveWordCommentsCommands(wdCommandList, documentInfo);
  }

  // get ppt settings
  const settingPath = await selectExistsPath(
    [
      documentInfo.pptxSettings,
      Path.resolve(mdSourcePath, documentInfo.pptxSettings),
      "../master-settings.js",
      "../../master-settings.js",
    ],
    __dirname
  );
  pptStyle = require(settingPath);

  // initialize pptx
  let pptx: PptxGenJS = new pptxGen();

  // FYI: use `headFontFace` and/or `bodyFontFace` to set the default font for the entire presentation (including slide Masters)
  // pptx.theme = { bodyFontFace: "Arial" };
  pptx.layout = "LAYOUT_WIDE";

  // create master slide, some bugs in the slide number.
  createMasterSlides(pptx);

  // create sheet object
  const currentSheet = new PptSheet(pptx);
  currentSheet.setDefaultPosition({
    ...getPositionPCT("10,15,70,70"),
    valign: "top",
  });

  // add title slide
  if (documentInfo.title) {
    currentSheet.addTitleSlide(documentInfo);
  }

  currentSheet.addMasterSlide();

  let currentDocxParagraph = new DocParagraph(
    WdNodeType.text,
    pptStyle.body,
    0
  );
  let tableJs: TableJs | undefined = undefined;

  documentInfo.position = "";
  // main loop
  for (let i = 0; i < lines.length; i++) {
    const wdCommandList = lines[i].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      // flush texts before creating tables.
      const textPropsArray = currentDocxParagraph.createTextPropsArray();
      currentSheet.addTextPropsArray(...textPropsArray);
      currentSheet.addTextFrame(); //getPosition(documentInfo.position, pptx));

      // initialize table.
      tableJs = new TableJs(
        parseInt(wdCommandList[1]),
        parseInt(wdCommandList[2]),
        mdSourcePath
      );
      continue;
    }

    // table command
    if (wdCommandList[0].includes("table")) {
      tableJs!.doTableCommand(lines[i]);
    } else {
      // in not table command, create table.
      if (tableJs) {
        const r = tableJs!.createTable();
        currentSheet.addTable(r);
        tableJs = undefined;
      }
    }

    // image command
    if (wdCommandList[0] === "image") {
      //create text frame
      const textPropsArray = currentDocxParagraph.createTextPropsArray();
      currentSheet.addTextPropsArray(...textPropsArray);
      currentSheet.addTextFrame(); //getPosition(documentInfo.position, pptx));

      // initialize image
      const image = createImageChild(
        mdSourcePath,
        wdCommandList[1],
        wdCommandList[2],
        pptx,
        currentSheet.currentTextPropPosition
      );
      currentSheet.addImage(image);
      continue;
    }

    // html comment command <!-- word xxxx -->
    const isResolveCommand = resolveWordCommentsCommands(
      wdCommandList,
      documentInfo
    );

    if (isResolveCommand) {
      continue;
    }

    // body commands
    currentDocxParagraph = await resolveWordDownCommandEx(
      lines[i],
      currentDocxParagraph,
      currentSheet,
      mdSourcePath
    );

    // update current position
    if (documentInfo.position) {
      currentSheet.setCurrentPosition({
        ...getPositionPCT(documentInfo.position),
      });
      documentInfo.position = "";
    }

    // when paragraph end, flush paragraph
    const isNewSheet = currentDocxParagraph.isNewSheet;
    if (currentDocxParagraph.isFlush || isNewSheet) {
      const textPropsArray = currentDocxParagraph.createTextPropsArray();
      currentSheet.addTextPropsArray(...textPropsArray);

      // reset paragraph. but keep the indent.
      currentDocxParagraph = new DocParagraph(
        WdNodeType.text,
        pptStyle.body,
        currentDocxParagraph.indent
      );

      if (isNewSheet) {
        currentSheet.addTextFrame(); // getPosition(documentInfo.position, pptx));
        currentSheet.createSheet();

        // new sheet
        currentSheet.addMasterSlide();
        currentDocxParagraph.isNewSheet = false;
      }
    }
  }
  // end loop lines
  const textPropsArray = currentDocxParagraph.createTextPropsArray();
  currentSheet.addTextPropsArray(...textPropsArray);
  currentSheet.addTextFrame(); //getPosition(documentInfo.position, pptx));
  currentSheet.createSheet();

  //Export Presentation
  const ff = `PptxGenJS_Demo_${new Date()
    .toISOString()
    .replace(/\D/gi, "")}.pptx`;

  const outPathPPtx = Path.resolve(Path.dirname(mdSourcePath), ff);

  pptx.title = documentInfo.title; // "PptxGenJS Test Suite Presentation";
  pptx.subject = documentInfo.subTitle; // "PptxGenJS Test Suite Export";
  pptx.author = documentInfo.author; // ("Brent Ely");
  pptx.revision = "1";

  const r = await pptx.writeFile({
    fileName: outPathPPtx,
    compression: true,
  });

  // open ppt
  const pptExe = await selectExistsPath(
    [
      "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
      "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
    ],
    ""
  );

  runCommand(pptExe, r);
}

/**
 *
 * @param pptx
 */
function createMasterSlides(pptx: PptxGenJS) {
  // https://github.com/gitbrent/PptxGenJS/issues/1175
  // TITLE_SLIDE
  pptx.defineSlideMaster(JSON.parse(JSON.stringify(pptStyle.titleSlide)));

  // MASTER_SLIDE (MASTER_PLACEHOLDER)
  pptx.defineSlideMaster(JSON.parse(JSON.stringify(pptStyle.masterSlide)));

  // THANKS_SLIDE (THANKS_PLACEHOLDER)
  pptx.defineSlideMaster(JSON.parse(JSON.stringify(pptStyle.thanksSlide)));
}

/**
 *
 * @param wdCommandList
 * @param documentInfo
 * @returns
 */
function resolveWordCommentsCommands(
  wdCommandList: string[],
  documentInfo: { [v: string]: string }
) {
  const documentInfoKeys = Object.keys(documentInfo);
  if (
    wdCommandList[0] === "param" &&
    documentInfoKeys.includes(wdCommandList[1])
  ) {
    documentInfo[wdCommandList[1]] = wdCommandList[2];
    return true;
  }
  return false;
}

/**
 *
 * @param header
 * @returns
 */
function getHeaderStyle(header: string) {
  switch (parseInt(header)) {
    case 1:
      return pptStyle.h1;
    case 2:
      return pptStyle.h2;
    case 3:
      return pptStyle.h3;
    case 4:
      return pptStyle.h4;
    case 5:
      return pptStyle.h5;
    case 6:
      return pptStyle.h6;
    default:
      return pptStyle.h6;
  }
}

/**
 *
 * @param line
 * @param docParagraph
 * @param slide
 * @param mdSourcePath
 * @returns
 */
async function resolveWordDownCommandEx(
  line: string,
  docParagraph: DocParagraph,
  slide: PptSheet,
  mdSourcePath: string
) {
  const words = line.split(_sp);
  let current: DocParagraph;
  const nodeType = words[0] as WdNodeType;
  let style: DocxStyle;
  let child: PptxGenJS.TextProps;
  let fontSize: number | undefined;

  switch (nodeType) {
    case "section":
      // word
      // section|1|タイトル|タイトル(slug)
      const currentStyle = getHeaderStyle(words[1]);
      if (words[1] === "1") {
        current = new DocParagraph(WdNodeType.section);
        resolveEmphasis(words[2]).forEach((x) => current.addChild(x));
        slide.addHeader(current.createTextPropsArray(), currentStyle);
        return docParagraph;
      }

      current = new DocParagraph(WdNodeType.section, currentStyle);
      resolveEmphasis(words[2]).forEach((x) => current.addChild(x));
      docParagraph.currentFontSize = currentStyle.fontSize ?? 32;
      return current;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      docParagraph.addChild({
        text: " ",
        options: {
          bullet: true,
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
          ...pptStyle.body
        },
      });
      fontSize = getHeaderStyle("").fontSize;
      if (fontSize) {
        docParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      return docParagraph;
      break;
    case WdNodeType.OderList:
      docParagraph.addChild({
        text: " ",
        options: {
          bullet: { type: "number", style: "romanLcPeriod" },
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
          ...pptStyle.body
        },
      });
      fontSize = getHeaderStyle("").fontSize;
      if (fontSize) {
        docParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      return docParagraph;
      break;
    case "code":
      docParagraph.addChild({
        text: words[1],
        options: {
          fontFace: "Arial",
          //color: pptx.SchemeColor.accent5,
          highlight: "FFFF00",
          ...pptStyle.body
        },
      });
      return docParagraph;
      break;
    case WdNodeType.link:
      // if (!words[3]) {
      //   let children = [];
      //   children.push(new SimpleField(` REF ${words[1]} \\w \\h `));
      //   children.push(new TextRun("---"));
      //   children.push(new SimpleField(` REF ${words[1]} \\h `));
      //   nodes.addChildren(children);
      // } else {
      //   child = new ExternalHyperlink({
      //     children: [
      //       new TextRun({
      //         text: words[3],
      //       }),
      //     ],
      //     link: words[1],
      //   });
      //   nodes.addChild(child);
      // }
      return docParagraph;
      break;
    case WdNodeType.image:
      // child = createImageChild(mdSourcePath, words[1]);
      // nodes.addChild(child, true);
      return docParagraph;
      break;
    case WdNodeType.hr:
      docParagraph.isFlush = true;
      docParagraph.isNewSheet = true;
      return docParagraph;
    case "text":
      const admonition = words[1].match(/^(note|warning)(:\s)(.*)/i);

      let s = words[1];
      if (admonition && admonition[3]) {
        s = admonition[3];
        const admonitionType = admonition[1];
        docParagraph.nodeType = admonitionType as WdNodeType;
        docParagraph.docxStyle = resolveAdmonition(admonitionType);
      }

      const mathBlock = s.match(/^\$(.*)\$$/);
      // if (mathBlock?.length) {
      //   const child = await createMathImage(mathBlock[1]);
      //   nodes.addChild(child, true);
      // } else {
      //   resolveEmphasis(s).forEach((x) => nodes.addChild(x));
      // }

      resolveEmphasis(s).forEach((x) =>
        docParagraph.addChild({
          text: x.text,
          options: {
            ...x.options,
            fontSize: docParagraph.currentFontSize,
            valign: "top",
          },
        })
      );

      return docParagraph;
      break;
    case "indentPlus":
      docParagraph.addIndent();
      return docParagraph;
      break;
    case "indentMinus":
      docParagraph.removeIndent();
      return docParagraph;
      break;
    case "newLine":
      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        docParagraph.addChild({ text: "", options: { breakLine: true } });
        docParagraph.isFlush = true;
      }
      return docParagraph;
    default:
      return docParagraph;
  }
}

/**
 *
 * @param s
 * @returns
 */
function resolveAdmonition(s: string) {
  let admonition: DocxStyle = DocxStyle.note1;
  switch (s.toLocaleLowerCase()) {
    case "note":
      admonition = DocxStyle.note1;
      break;
    case "warning":
      admonition = DocxStyle.warning1;
      break;
    default:
      //
      break;
  }
  return admonition as DocxStyle;
}

/**
 *
 * @param mathEq
 * @returns
 */
async function createMathImage(mathEq: string) {
  const options = {
    width: 1280,
    ex: 8,
    em: 16,
  };

  const svgStr = TeXToSVG(mathEq, options);

  // create png
  const pngArray = await svg2imagePng(svgStr);
  let pngBuffer = Buffer.from(pngArray);
  const sizeImageMath = imageSize.imageSize(pngBuffer);

  // const child = new ImageRun({
  //   data: pngBuffer,
  //   transformation: {
  //     width: sizeImageMath.width!,
  //     height: sizeImageMath.height!,
  //   },
  // });
  // return child;
}

/**
 *
 * @param position x,y,w,h in percent
 * @param pptx for get slide size
 * @returns position{ pptxGen } in inches
 */
function getPositionInch(position: string, pptx: pptxGen) {
  const positions = position.split(",");

  // 914400; // One (1) inch (OfficeXML measures in EMU (English Metric Units))
  const pEmp = 1.093613298337708e-8; // 1/ 914400 * 0.01

  const x: number = parseFloat(positions[0]) * pptx.presLayout.width * pEmp;
  const y: number = parseFloat(positions[1]) * pptx.presLayout.height * pEmp;
  const w: number = parseFloat(positions[2]) * pptx.presLayout.width * pEmp;
  const h: number = parseFloat(positions[3]) * pptx.presLayout.height * pEmp;

  return { x, y, w, h };
}

/**
 * 
 * @param position x,y,w,h in percent
 * @returns 
 */
function getPositionPCT(position: string) {
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
 * @param mdSourcePath
 * @param imagePathR
 * @returns
 */
function createImageChild(
  mdSourcePath: string,
  imagePathR: string,
  imageAlt: string,
  pptx: pptxGen,
  pos: { [k: string]: string | number } = {}
) {
  const imagePath = Path.resolve(mdSourcePath, imagePathR);
  const sizeImage = imageSize.imageSize(imagePath);
  // max 6inch 15cm
  const maxSize = 600; //convertInchesToTwip(3);

  let width = sizeImage.width ?? 100;
  let height = sizeImage.height ?? 100;

  if (width > maxSize || height > maxSize) {
    const r = maxSize / Math.max(width, height);
    width *= r;
    height *= r;
  }

  let positions = {};

  if (imageAlt) {
    positions = getPositionPCT(imageAlt);
  }

  if (pos.x && pos.y) {
    positions = { x: pos.x, y: pos.y, h: height / 94, w: width / 94 };
  }

  return { path: imagePath, ...positions };
}
