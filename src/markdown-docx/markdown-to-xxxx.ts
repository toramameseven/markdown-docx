import { markdownToWd0 } from "./markdown-to-wd0";
import { wd0ToDocx as wd0ToWd } from "./wd0-to-wd";
import { createPath, rmFileIfExist } from "./common";
import * as Fs from "fs";
import * as Path from "path";
const fm = require("front-matter");
import * as fmModule from "front-matter";
import {
  DocxOption,
  getFileContents,
  MessageType,
  ShowMessage,
} from "./common";
import { wordDownToDocx } from "./wd-to-docx";
import { wdToEd } from "./wd-to-ed";
import { wdToTextile } from "./wd-to-textile";
import { createInlineHtml } from "../tools/createInlineHtml";
import { wordDownToPptxBody } from "./wd-to-pptxJs";


/**message function */
let showMessage: ShowMessage;

// front matter is not used
export type XFrontMatter = {
  dummy?: string;
};

// main
export async function markdownToDocx(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption
) {
  option.message && (showMessage = option.message);
  /** *.wd file path */
  let fileWd = "";

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      "Markdown to docx: start !!!",
      "markdown-to-xxxx",
      false
    );

    // wdPath: path wd file, wdBody: wd strings
    const { wdPath, wdBody } = await markdownToWd(
      pathMarkdown,
      selection,
      "docx",
      startLine,
      option.isDebug,
      !option.isUseDocxJs
    );

    // get template and engine from the body text. engine is only for vba(vbs).
    const docxEngineInsideWdBody = getDocxEngine(wdBody);
    const templateInsideWdBody = getDocxTemplate(wdBody);
    fileWd = wdPath;

    // get docx docxEngine and docxTemplate in a wd file or options.
    option.docxEngine = docxEngineInsideWdBody
      ? docxEngineInsideWdBody
      : option.docxEngine;
    option.docxTemplate = templateInsideWdBody
      ? templateInsideWdBody
      : option.docxTemplate;

    //create docx (docxJs or vbs)
    await wordDownToDocx(fileWd, wdBody, option);
  } catch (ex) {
    throw ex;
  } finally {
    if (!option.isDebug) {
      await rmFileIfExist(fileWd, { force: true });
    }
  }
}

export async function markdownToExcel(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption
) {
  option.message && (showMessage = option.message);
  const dirPath = Path.dirname(pathMarkdown);
  const fileNameMd = Path.basename(pathMarkdown).replace(/\.md$/i, "");
  const fileEd = await createPath(dirPath, fileNameMd, "ed", true);

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      "Markdown to ed: start !!!",
      "markdown-to-xxxx",
      false
    );

    const r = await markdownToWd(
      pathMarkdown,
      selection,
      "excel",
      startLine,
      option.isDebug
    );

    // wdToEd(wdBody, option);
    const edBody = wdToEd(r.wdBody);
    Fs.writeFileSync(fileEd, edBody);
  } catch (ex) {
    throw ex;
  }
}

export async function markdownToTextile(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption
) {
  option.message && (showMessage = option.message);
  const dirPath = Path.dirname(pathMarkdown);
  const fileNameMd = Path.basename(pathMarkdown).replace(/\.md$/i, "");
  const fileEd = await createPath(dirPath, fileNameMd, "textile", true);

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      "Markdown to textile: start !!!",
      "markdown-to-xxxx",
      false
    );

    const r = await markdownToWd(
      pathMarkdown,
      selection,
      "textile",
      startLine,
      option.isDebug
    );

    // wdToEd(wdBody, option);
    //const edBody = htmlToTextile("", r.wdBody);

    const edBody = wdToTextile(r.wdBody,showMessage);

    Fs.writeFileSync(fileEd, edBody);
  } catch (ex) {
    throw ex;
  }
}

export async function markdownToPptx(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption
) {
  option.message && (showMessage = option.message);
  let fileWd = "";
  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      "Markdown to pptx: start !!!",
      "markdown-to-xxxx",
      false
    );

    const r = await markdownToWd(
      pathMarkdown,
      selection,
      "docx",
      startLine,
      option.isDebug
    );

    fileWd = r.wdPath;

    //create docx (docxJs or vbs)
    await wordDownToPptxBody(fileWd, r.wdBody, option);
  } catch (ex) {
    throw ex;
  } finally {
    if (!option.isDebug) {
      await rmFileIfExist(fileWd, { force: true });
    }
  }
}

export async function markdownToHtml(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption,
  isInclude: boolean = false
) {
  option.message && (showMessage = option.message);
  const dirPath = Path.dirname(pathMarkdown);
  const fileNameMd = Path.basename(pathMarkdown).replace(/\.md$/i, "");
  const fileHtml = await createPath(dirPath, fileNameMd, "html", true);

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      `markdown to inline html:`,
      "markdown-to-xxxx",
      false
    );

    let r = await markdownToWd(
      pathMarkdown,
      selection,
      "html",
      startLine,
      option.isDebug
    );
    
    let rr = '';
    if (isInclude){
      rr = await createInlineHtml(pathMarkdown, r.wdBody);
    }

    const outHtml = rr ? rr: r.wdBody;
    Fs.writeFileSync(fileHtml, outHtml);

    return fileHtml;
  } catch (ex) {
    throw ex;
  }
}

export async function textileToMarkdown(
  pathMarkdown: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  option: DocxOption,
  isInclude: boolean = false
) {
  option.message && (showMessage = option.message);
  const dirPath = Path.dirname(pathMarkdown);
  const fileNameTextile = Path.basename(pathMarkdown).replace(/\.textile$/i, "");
  const fileHtml = await createPath(dirPath, fileNameTextile, "html", true);

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      `textile to html to markdown:`,
      "markdown-to-xxxx",
      false
    );

    let r = await markdownToWd(
      pathMarkdown,
      selection,
      "html",
      startLine,
      option.isDebug
    );
    
    let rr = '';
    if (isInclude){
      rr = await createInlineHtml(pathMarkdown, r.wdBody);
    }

    const outHtml = rr ? rr: r.wdBody;
    Fs.writeFileSync(fileHtml, outHtml);

    return fileHtml;
  } catch (ex) {
    throw ex;
  }
}


//creating wd file for test
export async function markdownToWd(
  filePath: string,
  selection: string,
  convertType: "docx" | "excel" | "html" | "textile" = "docx",
  /** zero base number */
  startLine = 0,
  isDebug: boolean = false,
  isSaveWd: boolean = false
) {
  const dirPath = Path.dirname(filePath);
  const fileNameMd = Path.basename(filePath).replace(/\.md$/i, "");
  const fileWd0 = await createPath(dirPath, fileNameMd, "wd0", true);
  const fileWd = await createPath(dirPath, fileNameMd, "wd", true);

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      `markdown to wd:`,
      "markdown-to-xxxx",
      false
    );

    // get file contents
    let fileContents = getFileContents(filePath);

    // get front matter and body
    // const { frontMatter, markdownBody } = getFrontMatterAndBody(
    const r = getFrontMatterAndBody(
      filePath,
      fileContents,
      selection,
      startLine
    );

    // for @ mark escape
    const markdownBodyX = r.markdownBody.replace(/@/g, "\\@");
    // if \s\s\n is, convert it newline.
    const markdownBodyXX = markdownBodyX.replace(/\s{2,}\n/g, "\n\n");

    // create wd0
    const wd0 = markdownToWd0(
      markdownBodyXX,
      convertType,
      { sanitize: false },
      showMessage
    ).replace(/(\r?\n)+/g, "\n");

    if (convertType === "html") {
      return { wdPath: fileWd, wdBody: wd0 };
    }

    // wd must vbLF
    // wd0: html docx pptx ed(excel) textile
    const wd = wd0ToWd(wd0, showMessage);

    // output wd0 file.
    isDebug && Fs.writeFileSync(fileWd0, wd0);
    // output wd file.
    (isDebug || isSaveWd)  &&   Fs.writeFileSync(fileWd, wd);

    return { wdPath: fileWd, wdBody: wd };
  } catch (ex) {
    throw ex;
  } finally {
    if (!isDebug) {
      await rmFileIfExist(fileWd0, { force: true });
    }
  }
}

function getFrontMatterAndBody(
  filePath: string,
  fileContents: string,
  selection: string,
  /**zero base number */
  startLine: number
) {
  let frontMatterAndBody: fmModule.FrontMatterResult<XFrontMatter> | undefined;
  let frontMatter: XFrontMatter = {};
  let markdownBody = selection;
  frontMatterAndBody = undefined;

  try {
    frontMatterAndBody = fm(fileContents);
  } catch (e) {
    console.log(e);
    return { frontMatter, markdownBody };
  }

  // get lines to convert
  // bodyBegin one base number
  const deleteRows = frontMatterAndBody!.bodyBegin - (startLine + 1);
  if (deleteRows > 0) {
    markdownBody = selection.split("\n").slice(deleteRows).join("\n");
  }
  if (markdownBody === "") {
    markdownBody = frontMatterAndBody!.body;
  }
  // get front matter
  frontMatter = frontMatterAndBody!.attributes ?? {};
  if (typeof frontMatter !== "object") {
    // for sometime attributes is string
    frontMatter = {};
  }
  return { frontMatter, markdownBody };
}

function getDocxEngine(wd: string) {
  const testMatch = wd.match(/^docxEngine\t(?<docxEngine>.*)\t/i);
  const docxEngine = testMatch?.groups?.docxEngine ?? "";
  return docxEngine;
}

function getDocxTemplate(wd: string) {
  const testMatch = wd.match(/^docxTemplate\t(?<docxTemplate>.*)\t/im);
  const docxTemplate = testMatch?.groups?.docxTemplate ?? "";
  return docxTemplate;
}
