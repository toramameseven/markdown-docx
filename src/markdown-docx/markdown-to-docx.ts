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
  let fileWd = "";
  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      `markdown to docx:`,
      "Markdown to docx: start !!!",
      false
    );

    const { wdPath, wdBody } = await markdownToWd(
      pathMarkdown,
      selection,
      startLine,
      option.isDebug
    );

    const docxEngineInsideWdBody = getDocxEngine(wdBody);
    const templateInsideWdBody = getDocxTemplate(wdBody);
    fileWd = wdPath;

    // get docx docxEngine and docxTemplate in a wd file.
    option.docxEngine = docxEngineInsideWdBody
      ? docxEngineInsideWdBody
      : option.docxEngine;
    option.docxTemplate = templateInsideWdBody
      ? templateInsideWdBody
      : option.docxTemplate;

    //create docx by vbs
    await wordDownToDocx(fileWd, option);
  } catch (ex) {
    throw ex;
  } finally {
    if (!option.isDebug) {
      await rmFileIfExist(fileWd, { force: true });
    }
  }
}

//creating wd file for test
export async function markdownToWd(
  filePath: string,
  selection: string,
  /** zero base number */
  startLine = 0,
  isDebug?: boolean,
  isDryRun?: boolean
) {
  const dirPath = Path.dirname(filePath);
  const fileNameMd = Path.basename(filePath).replace(/\.md$/i, "");
  const fileWd0 = await createPath(dirPath, fileNameMd, "wd0");
  const fileWd = await createPath(dirPath, fileNameMd, "wd");

  // convert markdown to docx
  try {
    showMessage?.(
      MessageType.info,
      `markdown to wd:`,
      "Markdown to Wd start !!!",
      false
    );

    // get file contents
    let fileContents = getFileContents(filePath);

    // get front matter and body
    const { frontMatter, markdownBody } = getFrontMatterAndBody(
      filePath,
      fileContents,
      selection,
      startLine
    );

    // for @ mark escape
    const markdownBodyX = markdownBody.replace(/@/g, "\\@");

    const wd0 = markdownToWd0(
      markdownBodyX,
      { sanitize: false },
      showMessage
    ).replace(/(\r?\n)+/g, "\n");

    // wd must vbLF
    const wd = wd0ToWd(wd0, showMessage);

    // if dry run, not save files
    if (!isDryRun) {
      // output wd0 file.
      if (isDebug) {
        Fs.writeFileSync(fileWd0, wd0);
      }
      // output wd file.
      Fs.writeFileSync(fileWd, wd);
    }

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
