/**
 * https://gitbrent.github.io/PptxGenJS/
 */
import pptxGen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import * as Path from "path";
import * as imageSize from "image-size";
const TeXToSVG = require("tex-to-svg");
import {
  docxStyle,
  PptxDocument,
  PptStyle,
  TableJs,
  TextFrame,
  wdCommand,
  resolveEmphasis,
  WdCommand,
  DocxStyle,
  Position,
  initialPosition,
  PositionP,
  initialPositionP,
} from "./pptxjs";

import { svg2imagePng } from "./tools/svg-png-image";
import { runCommand, selectExistsPath } from "./tools/tools-common";
import {
  DocxOption,
  MessageType,
  ShowMessage,
  getFileContents,
} from "./common";

type Record = { [v: string]: string };

type DocumentInfoParam = {
  dpi?: string;
  position?: string;
  pptxSettings?: string;
};
type DocumentInfo = {
  placeholder: Record;
  param: DocumentInfoParam;
};

let thisMessage: ShowMessage | undefined;

let pptStyle: PptStyle = {
  titleSlide: { title: "" },
  masterSlide: { title: "" },
  h1: {},
  h2: {},
  h3: {},
  h4: {},
  h5: {},
  h6: {},
  body: {},
  code: {},
  codeSpan: {},
  tableProps: {},
  layout: "",
  headFontFace: {},
  bodyFontFace: {},
  tableHeaderColor: "000000",
  tableHeaderFillColor: "FFFFFF",
  defaultPositionPCT: "10,15,80,70",
  tablePropsArray: [],
};

// for delete the require cache.
let pptxSettingsFilePath = "";

// wd command separator
const _sp = "\t";

/** slide paging at section */
const isNewSlideAtSection: boolean = true;

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
    await wdToPptxJs(body, fileWd, option);
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
 * @param wdFullPath
 */
export async function wdToPptxJs(
  body: string,
  wdFullPath: string,
  option: DocxOption
): Promise<void> {
  const functionName = "wdToPptxJs";
  // doc info
  const documentInfo: DocumentInfo = { placeholder: {}, param: {} };
  const wdDirFullPath = Path.dirname(wdFullPath);

  thisMessage = option.message;

  // get lines
  const lines = (body + "\nend").split(/\r?\n/);

  // get document information
  const toInfoSearch = lines.length > 50 ? 49 : lines.length - 1;
  for (let i = 0; i < toInfoSearch; i++) {
    const wdCommandList = lines[i].split(_sp);
    const wdCommandList2 = lines[i + 1].split(_sp);
    // html comment command <!-- word xxxx -->
    resolveCommentCommand(wdCommandList, wdCommandList2, documentInfo);
  }

  // get ppt settings
  const settingPath = await selectExistsPath(
    [documentInfo.param.pptxSettings ?? "", "clear.ppt.js"],
    [wdDirFullPath, `${__dirname}/../templates`, `${__dirname}/../../templates`]
  );

  if (!settingPath) {
    option.message?.(
      MessageType.warn,
      `ppt settings: no setting is set.`,
      "wd-to-pptxJs",
      true
    );
    return;
  }

  // get pptStyle
  try {
    if (pptxSettingsFilePath === settingPath) {
      delete require.cache[pptxSettingsFilePath];
    }
    pptStyle = require(settingPath);
    pptxSettingsFilePath = settingPath;
    thisMessage?.(
      MessageType.info,
      `${functionName}:ppt setting is ${pptxSettingsFilePath}.`,
      "wd-to-pptxJs",
      false
    );
  } catch (e) {
    thisMessage?.(
      MessageType.err,
      `${functionName}:err read pptx settings.`,
      "wd-to-pptxJs",
      false
    );
    throw e;
  }

  // initialize pptx
  let pptx: PptxGenJS = new pptxGen();

  // may be for en-US
  pptx.theme = {
    ...pptx.theme,
    ...pptStyle.headFontFace,
    ...pptStyle.bodyFontFace,
  }; // { headFontFace: "Arial Light" };

  // FYI: use `headFontFace` and/or `bodyFontFace` to set the default font for the entire presentation (including slide Masters)
  // pptx.theme = { bodyFontFace: "Arial" };
  pptx.layout = pptStyle.layout; //"LAYOUT_WIDE";

  // create master slide, some bugs in the slide number.
  createMasterSlides(pptx);

  // create sheet object
  const pptDocument = new PptxDocument(pptx, pptStyle);
  pptDocument.setDefaultPositionPCT({
    ...getPositionPercent(pptStyle.defaultPositionPCT),
  });

  // add pptx document title (cover page)
  if (documentInfo.placeholder.title) {
    pptDocument.addPptxCover(documentInfo.placeholder);
  }

  // create first slide
  if (!isNewSlideAtSection) {
    pptDocument.addDocumentSlide();
  }

  // working table
  let tableJs: TableJs | undefined = undefined;

  // document information
  documentInfo.param.position = "";

  // main loop (wd)
  for (let i = 0; i < lines.length - 1; i++) {
    const wdCommandList = lines[i].split(_sp);
    const wdCommandList2 = lines[i + 1].split(_sp);

    // when find create table
    if (wdCommandList[0] === "tableCreate") {
      // flush texts before creating tables.
      pptDocument.addTextPropsArrayFromParagraph();
      pptDocument.addTextFrameToSheetObjects();

      // initialize table.
      tableJs = new TableJs(
        parseInt(wdCommandList[1]),
        parseInt(wdCommandList[2]),
        wdDirFullPath
      );
      continue;
    }

    // table command
    if (wdCommandList[0].includes("table")) {
      tableJs!.doTableCommand(lines[i], pptStyle);
    } else {
      // in not table command, create table.
      if (tableJs) {
        let slideWidth = percentToInches(
          pptDocument.currentTextPropPositionPCT.w,
          pptx
        );

        const r = tableJs!.createTable(
          pptDocument.currentTextPropPositionPCT,
          pptStyle.tableProps,
          slideWidth
        );
        pptDocument.addTableToSheetObjects(r);
        tableJs = undefined;
      }
    }

    // image command
    if (wdCommandList[0] === "image") {
      //create text frame
      pptDocument.addTextPropsArrayFromParagraph();
      pptDocument.addTextFrameToSheetObjects();

      // initialize image
      const image = createImageChild(
        wdDirFullPath,
        wdCommandList[1],
        wdCommandList[2],
        pptx,
        pptDocument.currentTextPropPositionPCT,
        parseInt(documentInfo.param.dpi ?? "96")
      );
      pptDocument.addImageToSheetObjects(image);
      continue;
    }

    // shape command
    if (
      wdCommandList[0].split("/")[0] === "code" &&
      wdCommandList[0].split("/")[1] === "json:ppt"
    ) {
      //create text frame
      pptDocument.addTextPropsArrayFromParagraph();
      pptDocument.addTextFrameToSheetObjects();

      pptDocument.addRawString(wdCommandList[1]);
      // initialize image
      continue;
    }

    // flush ppt shape
    if (wdCommandList[0] === "newLine" && wdCommandList[2] === "json:ppt") {
      pptDocument.addShapesToSheetObjects();
      continue;
    }

    // html comment command <!-- word xxxx -->
    const isResolveCommand = resolveCommentCommand(
      wdCommandList,
      wdCommandList2,
      documentInfo
    );

    if (isResolveCommand) {
      // update current position
      if (documentInfo.param.position) {
        //create text frame
        pptDocument.addTextPropsArrayFromParagraph();
        pptDocument.addTextFrameToSheetObjects();
        // update position
        pptDocument.setCurrentPositionPCT({
          ...getPositionPercent(documentInfo.param.position),
        });
        documentInfo.param.position = "";
      }
      continue;
    }

    // body commands( main command)
    await resolveWordDownCommandEx(lines[i], pptDocument);

    // only debug
    // thisMessage?.(
    //   MessageType.debug,
    //   `${functionName}:currentSheet.pptxParagraph.children.length: ${pptDocument.getCurrentParagraphNum()}`,
    //   "wd-to-pptxJs",
    //   false
    // );

    // when paragraph end, flush paragraph
    const isNewSheet = pptDocument.isNewSheet;
    if (pptDocument.isParagraphFlush || isNewSheet) {
      pptDocument.isParagraphFlush = false;
      pptDocument.addTextPropsArrayFromParagraph();

      if (isNewSheet) {
        pptDocument.addTextFrameToSheetObjects();
        pptDocument.createSheet();

        // new sheet
        pptDocument.addDocumentSlide();
        pptDocument.isNewSheet = false;
        // only debug
        // thisMessage?.(
        //   MessageType.debug,
        //   `${functionName}:add new slide`,
        //   "wd-to-pptxJs",
        //   false
        // );
      }
    }
  }

  // end loop lines
  pptDocument.addTextPropsArrayFromParagraph();
  pptDocument.addTextFrameToSheetObjects(); //getPosition(documentInfo.position, pptx));
  pptDocument.createSheet();

  //Export Presentation
  const outPathPPtx = Path.resolve(
    wdDirFullPath,
    Path.basename(wdFullPath, ".wd") + ".md.pptx"
  );

  pptx.title = documentInfo.placeholder.title ?? ""; // "PptxGenJS Test Suite Presentation";
  pptx.subject = documentInfo.placeholder.subTitle ?? ""; // "PptxGenJS Test Suite Export";
  pptx.author = documentInfo.placeholder.author ?? ""; // ("Brent Ely");
  pptx.revision = "1";

  let pptFilePath = "";
  try {
    pptFilePath = await pptx.writeFile({
      fileName: outPathPPtx,
      compression: true,
    });
  } catch (error) {
    thisMessage?.(
      MessageType.err,
      `${functionName}: ppt write error!!`,
      "wd-to-pptxJs",
      true
    );
    throw error;
  }

  // open ppt
  const pptExe = await selectExistsPath(
    [
      "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
      "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
    ],
    [""]
  );

  runCommand(pptExe, pptFilePath);
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
}

/**
 *
 * @param wdCommandList
 * @param documentInfo
 * @returns
 */
function resolveCommentCommand(
  wdCommandList: string[],
  wdCommandList2: string[],
  documentInfo: DocumentInfo
) {
  if (wdCommandList[0] === "placeholder") {
    documentInfo.placeholder[wdCommandList[1]] = wdCommandList[2];
    return true;
  }

  if (wdCommandList[0] === "param") {
    for (let i = 2; i < wdCommandList.length; i += 2) {
      switch (wdCommandList[i - 1]) {
        case "dpi":
          documentInfo.param.dpi = wdCommandList[i];
          break;
        case "pptxSettings":
          documentInfo.param.pptxSettings = wdCommandList[i];
          break;
        case "position":
          documentInfo.param.position = wdCommandList[i];
          break;
        default:
          // todo error
          break;
      }
    }
    return true;
  }

  // not comment command
  if (wdCommandList[0] === "section" && wdCommandList[1] === "1") {
    documentInfo.placeholder["title"] = wdCommandList2[1];
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
 * @param pptxDocument
 * @param mdSourcePath
 * @returns
 */
async function resolveWordDownCommandEx(
  line: string,
  pptxDocument: PptxDocument
) {
  const words = line.split(_sp);
  const nodeType = words[0].split("/")[0] as WdCommand;
  let fontSize: number | undefined;

  // only debug
  // thisMessage?.(
  //   MessageType.debug,
  //   `resolveWordDownCommandEx: ${nodeType}:${words[1]}`,
  //   "wd-to-pptxJs",
  //   false
  // );

  switch (nodeType) {
    case "section":
      // word
      // section|1|タイトル(slug)
      const currentStyle = getHeaderStyle(words[1]);

      if (words[1] === "2") {
        // ## is slide title
        if (isNewSlideAtSection) {
          pptxDocument.addTextFrameToSheetObjects();
          pptxDocument.createSheet();
          pptxDocument.addDocumentSlide();
        }
        pptxDocument.status = "insideSlideTitle";
      } else if (words[1] === "1") {
        // # is title for a cover
        pptxDocument.status = "insideCover";
      } else {
        // more ### headings
        pptxDocument.addTextProps({
          options: { breakLine: true },
        });
        pptxDocument.addTextPropsArrayFromParagraph();
      }

      pptxDocument.setTextPropsOptions({
        fontSize: currentStyle.fontSize ?? 32,
        lineSpacing: currentStyle.lineSpacing ?? 0,
      });
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      fontSize = getHeaderStyle("").fontSize ?? 10;
      pptxDocument.addTextProps({
        text: "\u200b",
        options: {
          bullet: true,
          indentLevel: parseInt(words[1]),
          ...pptStyle.body,
          fontSize,
        },
      });

      break;
    case wdCommand.OderList:
      fontSize = getHeaderStyle("").fontSize ?? 10;

      pptxDocument.addTextProps({
        text: "\u200b",
        options: {
          bullet: { type: "number", numberType: "arabicPeriod" },
          indentLevel: parseInt(words[1]),
          ...pptStyle.body,
          fontSize,
        },
      });

      break;
    case "code":
      if (words[1] === "") {
        // "end code" insert an empty line.
        return;
      }
      pptxDocument.addTextProps({
        text: words[1],
        options: {
          fontFace: "Arial",
          //color: pptx.SchemeColor.accent5,
          highlight: "FFFF00",
          ...pptStyle.body,
          ...pptStyle.code,
          fill: { color: "0088CC" },
          breakLine: true,
        },
      });
      pptxDocument.addRawString(words[1]);
      break;
    case wdCommand.link:
      // link ref|bookmark hover text
      if (!words[3]) {
        pptxDocument.addTextProps({
          text: words[1],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            fontSize: pptxDocument.textPropsOptions().fontSize ?? 18,
          },
        });
      } else {
        // outside link
        pptxDocument.addTextProps({
          text: words[3],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            //fontSize: slide.pptxParagraph.currentFontSize,
            fontSize: pptxDocument.textPropsOptions().fontSize ?? 18,
          },
        });
      }
      break;
    case wdCommand.image:
      break;
    case wdCommand.hr:
      if (!isNewSlideAtSection) {
        pptxDocument.isNewSheet = true;
      }
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

      resolveEmphasis(s, pptStyle).forEach((x) =>
        pptxDocument.addTextProps({
          text: x.text,
          options: {
            ...x.options,
            fontSize: pptxDocument.textPropsOptions().fontSize ?? 18,
            lineSpacing: pptxDocument.textPropsOptions().lineSpacing ?? 0,
            valign: "top",
          },
        })
      );
      break;
    case "indentPlus":
      pptxDocument.addIndent();
      break;
    case "indentMinus":
      pptxDocument.removeIndent();
      break;
    case "newLine":
      // slide title
      if (
        "convertHeading End" === words[1] &&
        pptxDocument.status === "insideSlideTitle"
      ) {
        pptxDocument.addSlideHeader({});
        pptxDocument.isParagraphFlush = true;
        pptxDocument.status = "non";
      }

      // # is document title, so this does not render #. do nothing here.
      // # is treated outside for loop.
      if (
        "convertHeading End" === words[1] &&
        pptxDocument.status === "insideCover"
      ) {
        pptxDocument.clearParagraph();
        pptxDocument.status = "non";
        return;
      }

      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph

        pptxDocument.addTextProps({
          options: { breakLine: true },
        });

        pptxDocument.isParagraphFlush = true;
      }
    default:
    // todo error;
  }
}

/**
 *
 * @param s
 * @returns
 */
function resolveAdmonition(s: string) {
  let admonition: DocxStyle = docxStyle.note1;
  switch (s.toLocaleLowerCase()) {
    case "note":
      admonition = docxStyle.note1;
      break;
    case "warning":
      admonition = docxStyle.warning1;
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

function percentToInches(lenPercent: number | string, pptx: pptxGen) {
  // 914400; // One (1) inch (OfficeXML measures in EMU (English Metric Units))
  let lenPercentR = 0;
  if (typeof lenPercent === "string") {
    lenPercentR = parseFloat(lenPercent);
  } else {
    lenPercentR = lenPercent;
  }
  const pEmp = 1.093613298337708e-8; // 1/ 914400 * 0.01
  const r: number = lenPercentR * pptx.presLayout.width * pEmp;
  return r;
}
/**
 *
 * @param position "x,y,w,h" in percent
 * @returns
 */
function getPositionPercent(position: string): PositionP {
  const functionName = "getPositionPCT";
  try {
    const positions = position.split(",");

    let x: number = parseFloat(positions[0]);
    let y: number = parseFloat(positions[1]);
    let w: number = parseFloat(positions[2]);
    let h: number = parseFloat(positions[3]);

    if (
      Number.isNaN(x) ||
      Number.isNaN(y) ||
      Number.isNaN(w) ||
      Number.isNaN(h) ||
      x > 100 ||
      x < 0 ||
      y > 100 ||
      y < 0 ||
      w > 100 ||
      w < 0 ||
      h > 100 ||
      h < 0
    ) {
      throw new Error("invalid number: position");
    }

    w = x + w > 100 ? 100 - x : w;
    h = y + h > 100 ? 100 - y : h;

    return {
      x: `${x}%`,
      y: `${y}%`,
      w: `${w}%`,
      h: `${h}%`,
    };
  } catch (e) {
    thisMessage?.(
      MessageType.err,
      `${functionName}: positions are not numeric: "${position}"`,
      "wd-to-pptxJs",
      true
    );
    return { x: "10%", y: "30%", w: "80%", h: "60%" };
  }
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
  pos: PositionP = { ...initialPositionP },
  dpi: number = 96
) {
  const imagePath = Path.resolve(mdSourcePath, imagePathR);

  // pixel > inch
  const sizeImage = imageSize.imageSize(imagePath);
  let imageWidthInch = (sizeImage.width ?? 100) / dpi;
  let imageHeightInch = (sizeImage.height ?? 100) / dpi;

  // slide size in (inch / 100)
  const pEmp = 1.093613298337708e-6; // 1/ 914400 * 0.01
  const slideWidth = pptx.presLayout.width * pEmp;
  const slideHeight = pptx.presLayout.height * pEmp;

  //frame size inch
  const frameWidth = slideWidth * parseFloat(pos.w.replace("%", "")) * 0.01;
  const frameHeight = slideHeight * parseFloat(pos.h.replace("%", "")) * 0.01;

  const r = imageWidthInch / imageHeightInch;
  if (imageWidthInch > frameWidth) {
    imageWidthInch = frameWidth;
    imageHeightInch = imageWidthInch / r;
  }

  if (imageHeightInch > frameHeight) {
    imageHeightInch = frameHeight;
    imageWidthInch = imageHeightInch * r;
  }

  let positions = {};

  if (pos.x && pos.y) {
    positions = { x: pos.x, y: pos.y, h: imageHeightInch, w: imageWidthInch };
  }

  return { path: imagePath, ...positions };
}
