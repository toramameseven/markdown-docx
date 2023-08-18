/**
 * https://gitbrent.github.io/PptxGenJS/
 */
import pptxGen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import * as Path from "path";
import * as imageSize from "image-size";
const TeXToSVG = require("tex-to-svg");
import {
  PptParagraph,
  docxStyle,
  PptSheet,
  PptStyle,
  TableJs,
  TextFrame,
  wdCommand,
  resolveEmphasis,
  WdCommand,
  DocxStyle,
} from "./pptxjs";

import { svg2imagePng } from "./tools/svg-png-image";
import { runCommand, selectExistsPath } from "./tools/tools-common";
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
  layout:"",
  headFontFace: {},
  bodyFontFace: {},
  tableHeaderColor: "000000",
  tableHeaderFillColor: "FFFFFF"
};

// for delete the require cache.
let pptxSettingsFilePath = "";

// wd command separator
const _sp = "\t";

// slide paging at section
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

type DocumentInfo = {
  placeholder: { [v: string]: string };
  param: { [v: string]: string };
};

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
  const lines = (body + "\nEndline").split(/\r?\n/);

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
    [documentInfo.param.pptxSettings ?? "", "master-settings.js"],
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

  pptx.theme = {...pptx.theme, ...pptStyle.headFontFace, ...pptStyle.bodyFontFace}; // { headFontFace: "Arial Light" };


  // FYI: use `headFontFace` and/or `bodyFontFace` to set the default font for the entire presentation (including slide Masters)
  // pptx.theme = { bodyFontFace: "Arial" };
  pptx.layout = pptStyle.layout; //"LAYOUT_WIDE";

  // create master slide, some bugs in the slide number.
  createMasterSlides(pptx);

  // create sheet object
  const currentSheet = new PptSheet(pptx, pptStyle);
  currentSheet.setDefaultPositionPCT({
    ...getPositionPercent("10,15,70,70"),
    valign: "top",
  });

  // add title slide
  if (documentInfo.placeholder.title) {
    currentSheet.addTitleSlide(documentInfo.placeholder);
  }

  // create first slide
  if (!isNewSlideAtSection) {
    currentSheet.addDocumentSlide();
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
      currentSheet.addTextPropsArray();
      currentSheet.addTextFrame();

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
        const r = tableJs!.createTable(
          currentSheet.currentTextPropPositionPCT,
          pptStyle.tableProps
        );
        currentSheet.addTable(r);
        tableJs = undefined;
      }
    }

    // image command
    if (wdCommandList[0] === "image") {
      //create text frame
      currentSheet.addTextPropsArray();
      currentSheet.addTextFrame();

      // initialize image
      const image = createImageChild(
        wdDirFullPath,
        wdCommandList[1],
        wdCommandList[2],
        pptx,
        currentSheet.currentTextPropPositionPCT,
        parseInt(documentInfo.param.dpi ?? "96")
      );
      currentSheet.addImage(image);
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
        currentSheet.addTextPropsArray();
        currentSheet.addTextFrame();
        // update position
        currentSheet.setCurrentPositionPCT({
          ...getPositionPercent(documentInfo.param.position),
        });
        documentInfo.param.position = "";
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
      `${functionName}:currentDocxParagraph:${currentSheet.pptxParagraph.children.length}`,
      "wd-to-pptxJs",
      false
    );

    // when paragraph end, flush paragraph
    const isNewSheet = currentSheet.pptxParagraph.isNewSheet;
    if (currentSheet.pptxParagraph.isFlush || isNewSheet) {
      currentSheet.addTextPropsArray();

      if (isNewSheet) {
        currentSheet.addTextFrame();
        currentSheet.createSheet();

        // new sheet
        currentSheet.addDocumentSlide();
        currentSheet.pptxParagraph.isNewSheet = false;
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
  const outPathPPtx = Path.resolve(
    wdDirFullPath,
    Path.basename(wdFullPath, ".wd") + ".md.pptx"
  );

  pptx.title = documentInfo.placeholder.title ?? ""; // "PptxGenJS Test Suite Presentation";
  pptx.subject = documentInfo.placeholder.subTitle ?? ""; // "PptxGenJS Test Suite Export";
  pptx.author = documentInfo.placeholder.author ?? ""; // ("Brent Ely");
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
    [""]
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
      if (wdCommandList[i - 1]) {
        documentInfo.param[wdCommandList[i - 1]] = wdCommandList[i];
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
 * @param slide
 * @param mdSourcePath
 * @returns
 */
async function resolveWordDownCommandEx(line: string, slide: PptSheet) {
  const words = line.split(_sp);
  const nodeType = words[0] as WdCommand;
  let fontSize: number | undefined;

  thisMessage?.(
    MessageType.debug,
    `resolveWordDownCommandEx: ${nodeType}:${words[1]}`,
    "wd-to-pptxJs",
    false
  );

  switch (nodeType) {
    case "section":
      // word
      // section|1|タイトル(slug)
      const currentStyle = getHeaderStyle(words[1]);
      // ## is slide title
      if (words[1] === "2") {
        if (isNewSlideAtSection) {
          slide.addTextFrame();
          slide.createSheet();
          slide.addDocumentSlide();
        }
        slide.pptxParagraph.insideSlideTitle = true;
      }
      // slide document title
      if (words[1] === "1") {
        slide.pptxParagraph.insideDocumentTitle = true;
      }
      //
      slide.pptxParagraph.addChild({
        text: " ",
        options: {
          ...pptStyle.body,
        },
      });

      slide.pptxParagraph.currentFontSize = currentStyle.fontSize ?? 32;
      slide.pptxParagraph.currentLineSpacing = currentStyle.lineSpacing ?? 0;
      break;
    case "NormalList":
      // OderList	1
      // text	Consectetur adipiscing elit
      // newLine	convertParagraph	tm
      slide.pptxParagraph.addChild({
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
        slide.pptxParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      break;
    case wdCommand.OderList:
      slide.pptxParagraph.addChild({
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
        slide.pptxParagraph.currentFontSize = fontSize;
      } else {
        //todo Error
      }
      break;
    case "code":
      if (words[1] === "") {
        // "end code" insert an empty line.
        return;
      }
      slide.pptxParagraph.addChild({
        text: words[1],
        options: {
          fontFace: "Arial",
        //color: pptx.SchemeColor.accent5,
          highlight: "FFFF00",
          ...pptStyle.body,
          ...pptStyle.code,
          fill:{ color:'0088CC' },
          breakLine:true
        },
      });
      break;
    case wdCommand.link:
      // link ref|bookmark hover text
      if (!words[3]) {
        slide.pptxParagraph.addChild({
          text: words[1],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            fontSize: slide.pptxParagraph.currentFontSize,
          },
        });
      } else {
        // outside link
        slide.pptxParagraph.addChild({
          text: words[3],
          options: {
            hyperlink: {
              url: words[1],
              tooltip: words[2],
            },
            fontSize: slide.pptxParagraph.currentFontSize,
          },
        });
      }
      break;
case wdCommand.image:
      break;
    case wdCommand.hr:
      if (!isNewSlideAtSection) {
        slide.pptxParagraph.isNewSheet = true;
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
        slide.pptxParagraph.addChild({
          text: x.text,
          options: {
            ...x.options,
            fontSize: slide.pptxParagraph.currentFontSize,
            lineSpacing: slide.pptxParagraph.currentLineSpacing,
            valign: "top",
          },
        })
      );
      break;
    case "indentPlus":
      slide.pptxParagraph.addIndent();
      break;
    case "indentMinus":
      slide.pptxParagraph.removeIndent();
      break;
    case "newLine":
      if (
        "convertHeading End" === words[1] &&
        slide.pptxParagraph.insideSlideTitle
      ) {
        const propArray = slide.pptxParagraph.createTextPropsArray();
        slide.addHeader(propArray, {});
        slide.pptxParagraph.insideSlideTitle = false;
      }

      // # is document title, so this does not render #.
      if (
        "convertHeading End" === words[1] &&
        slide.pptxParagraph.insideDocumentTitle
      ) {
        slide.pptxParagraph.insideDocumentTitle = false;
        slide.pptxParagraph.clear();
        return;
      }

      if (!["convertTitle", "convertSubTitle"].includes(words[1])) {
        // output paragraph
        slide.pptxParagraph.addChild({
          text: "",
          options: { breakLine: true },
        });
        slide.pptxParagraph.isFlush = true;
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

/**
 *
 * @param position "x,y,w,h" in percent
 * @returns
 */
function getPositionPercent(position: string) {
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
  pos: { [k: string]: string } = {},
  dpi: number = 96
) {
  const imagePath = Path.resolve(mdSourcePath, imagePathR);

  // pixel > inch
  const sizeImage = imageSize.imageSize(imagePath);
  let imageWidthInch = (sizeImage.width ?? 100) / dpi;
  let imageHeightInch = (sizeImage.height ?? 100) / dpi;

  // slide size in (inch / 100)
  const pEmp = 1.093613298337708e-8; // 1/ 914400 * 0.01
  const slideWidth = pptx.presLayout.width * pEmp;
  const slideHeight = pptx.presLayout.height * pEmp;

  //frame size inch
  const frameWidth = slideWidth * parseFloat(pos.w.replace("%", ""));
  const frameHeight = slideHeight * parseFloat(pos.w.replace("%", ""));

  const maxOutputSizeInch = Math.min(frameWidth, frameHeight);

  if (
    imageWidthInch > maxOutputSizeInch ||
    imageHeightInch > maxOutputSizeInch
  ) {
    const r = maxOutputSizeInch / Math.max(imageWidthInch, imageHeightInch);
    imageWidthInch *= r;
    imageHeightInch *= r;
  }

  let positions = {};

  if (pos.x && pos.y) {
    positions = { x: pos.x, y: pos.y, h: imageHeightInch, w: imageWidthInch };
  }

  return { path: imagePath, ...positions };
}
