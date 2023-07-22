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

let thisMessage: ShowMessage | undefined;

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
    await wdToPptxJs(body, Path.dirname(fileWd), option);
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
  mdSourcePath: string,
  option: DocxOption
): Promise<void> {
  const functionName = "wdToPptxJs";
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

  thisMessage = option.message;

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

  // create first slide
  currentSheet.addMasterSlide();

  // working table
  let tableJs: TableJs | undefined = undefined;
  // document information
  documentInfo.position = "";

  // main loop (wd)
  for (let i = 0; i < lines.length; i++) {
    const wdCommandList = lines[i].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      // flush texts before creating tables.
      currentSheet.addTextPropsArray();
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
        const r = tableJs!.createTable(currentSheet.currentTextPropPosition);
        currentSheet.addTable(r);
        tableJs = undefined;
      }
    }

    // image command
    if (wdCommandList[0] === "image") {
      //create text frame
      currentSheet.addTextPropsArray();
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
      // update current position
      if (documentInfo.position) {
        //create text frame
        currentSheet.addTextPropsArray();
        currentSheet.addTextFrame();

        currentSheet.setCurrentPosition({
          ...getPositionPCT(documentInfo.position),
        });
        documentInfo.position = "";
      }
      continue;
    }

    // body commands( main command)
    await resolveWordDownCommandEx(
      lines[i],
      //currentDocxParagraph,
      currentSheet
    );

    thisMessage?.(
      MessageType.debug,
      `${functionName}:currentDocxParagraph:${currentSheet.docxParagraph.children.length}`,
      "wd-to-pptxJs",
      false
    );

    // when paragraph end, flush paragraph
    const isNewSheet = currentSheet.docxParagraph.isNewSheet;
    if (currentSheet.docxParagraph.isFlush || isNewSheet) {
      currentSheet.addTextPropsArray();

      if (isNewSheet) {
        currentSheet.addTextFrame(); // getPosition(documentInfo.position, pptx));
        currentSheet.createSheet();

        // new sheet
        currentSheet.addMasterSlide();
        currentSheet.docxParagraph.isNewSheet = false;
        thisMessage?.(
          MessageType.debug,
          `${functionName}:add new slide`,
          "wd-to-pptxJs",
          false
        );
      }
    }
  }

  // end loop lines
  currentSheet.addTextPropsArray();
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
  const functionName = "resolveWordCommentsCommands";
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
 * @param slide.docxParagraph
 * @param slide
 * @param mdSourcePath
 * @returns
 */
async function resolveWordDownCommandEx(
  line: string,
  //slide.docxParagraph: DocParagraph,
  slide: PptSheet,
) {
  const functionName = "resolveWordDownCommandEx";
  const words = line.split(_sp);
  const nodeType = words[0] as WdNodeType;
  let fontSize: number | undefined;

  thisMessage?.(
    MessageType.debug,
    `${functionName}: ${nodeType}:${words[1]}`,
    "wd-to-pptxJs",
    false
  );

  switch (nodeType) {
    case "section":
      // word
      // section|1|タイトル(slug)
      const currentStyle = getHeaderStyle(words[1]);
      if (words[1] === "1") {
        slide.docxParagraph.insideSlideTitle = true;
      }

      /*
      if (words[1] === "1") {
        current = new DocParagraph();
        resolveEmphasis(words[2]).forEach((x) => current.addChild(x));
        slide.addHeader(current.createTextPropsArray(), currentStyle);
        return docParagraph;
      }

      current = new DocParagraph(currentStyle);
      resolveEmphasis(words[2]).forEach((x) => current.addChild(x));
      docParagraph.currentFontSize = currentStyle.fontSize ?? 32;
      return current;
*/
      //
      slide.docxParagraph.addChild({
        text: " ",
        options: {
          ...pptStyle.body,
        },
      });

      slide.docxParagraph.currentFontSize = currentStyle.fontSize ?? 32;
      return slide.docxParagraph;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      slide.docxParagraph.addChild({
        text: " ",
        options: {
          bullet: true,
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
          ...pptStyle.body,
        },
      });
      fontSize = getHeaderStyle("").fontSize;
      if (fontSize) {
        slide.docxParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      return slide.docxParagraph;
      break;
    case WdNodeType.OderList:
      slide.docxParagraph.addChild({
        text: " ",
        options: {
          bullet: { type: "number", style: "romanLcPeriod" },
          // color: PptxGenJS.SchemeColor.accent6,
          indentLevel: parseInt(words[1]),
          ...pptStyle.body,
        },
      });
      fontSize = getHeaderStyle("").fontSize;
      if (fontSize) {
        slide.docxParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      return slide.docxParagraph;
      break;
    case "code":
      slide.docxParagraph.addChild({
        text: words[1],
        options: {
          fontFace: "Arial",
          //color: pptx.SchemeColor.accent5,
          highlight: "FFFF00",
          ...pptStyle.body,
        },
      });
      return slide.docxParagraph;
      break;
    case WdNodeType.link:
      // link ref|bookmark hover text
      if (!words[3]) {
        slide.docxParagraph.addChild({
          text: words[1],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            fontSize: slide.docxParagraph.currentFontSize,
          },
        });
      } else {
        // outside link
        slide.docxParagraph.addChild({
          text: words[3],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            fontSize: slide.docxParagraph.currentFontSize,
          },
        });
      }
      return slide.docxParagraph;
      break;
    case WdNodeType.image:
      return slide.docxParagraph;
      break;
    case WdNodeType.hr:
      slide.docxParagraph.isFlush = true;
      slide.docxParagraph.isNewSheet = true;
      return slide.docxParagraph;
    case "text":
      const admonition = words[1].match(/^(note|warning)(:\s)(.*)/i);

      let s = words[1];
      if (admonition && admonition[3]) {
        s = admonition[3];
      }

      const mathBlock = s.match(/^\$(.*)\$$/);
      // if (mathBlock?.length) {
      //   const child = await createMathImage(mathBlock[1]);
      //   nodes.addChild(child, true);
      // } else {
      //   resolveEmphasis(s).forEach((x) => nodes.addChild(x));
      // }

      resolveEmphasis(s).forEach((x) =>
        slide.docxParagraph.addChild({
          text: x.text,
          options: {
            ...x.options,
            fontSize: slide.docxParagraph.currentFontSize,
            valign: "top",
          },
        })
      );
      return slide.docxParagraph;
      break;
    case "indentPlus":
      slide.docxParagraph.addIndent();
      return slide.docxParagraph;
      break;
    case "indentMinus":
      slide.docxParagraph.removeIndent();
      return slide.docxParagraph;
      break;
    case "newLine":
      if ("convertHeading End" === words[1] && slide.docxParagraph.insideSlideTitle) {
        const propArray = slide.docxParagraph.createTextPropsArray();
        slide.addHeader(propArray, {});
        slide.docxParagraph.insideSlideTitle = false;
        return slide.docxParagraph;
      }
      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        slide.docxParagraph.addChild({ text: "", options: { breakLine: true } });
        slide.docxParagraph.isFlush = true;
      }
      return slide.docxParagraph;
    default:
      return slide.docxParagraph;
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
