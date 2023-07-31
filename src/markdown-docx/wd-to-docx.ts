import {
  DocxOption,
  getFileContents,
  MessageType,
  fileExists,
  vbsSpawn,
  docxTemplate001,
  templatesPath
} from "./common";
import { wdToDocxJs } from "./wd-to-docxjs";
import * as Path from "path";
import { runCommand, selectExistsPath } from "./tools/tools-common";


export async function wordDownToDocx(
  fileWd: string,
  wdBody: string,
  option: DocxOption
) {

  if (wdBody === "") {
    wdBody = getFileContents(fileWd);
  }

  const defaultTemplate = Path.resolve(
    __dirname,
    `../${templatesPath}/${docxTemplate001}`
  );
  const defaultTemplate2 = Path.resolve(
    __dirname,
    `../../${templatesPath}/${docxTemplate001}`
  );
  const template = await selectExistsPath(
    [
      getDocxTemplateFromWd(wdBody),
      option.docxTemplate ?? "",
      defaultTemplate,
      defaultTemplate2,
    ],
    Path.dirname(fileWd)
  );

  option.message?.(
    MessageType.info,
    `docx docxTemplate: ${template ? template : "use inside"}`,
    "wd-to-docx",
    false
  );

  if (option.isUseDocxJs) {
    // create docx from docxJs
    if (!template) {
      option.message?.(
        MessageType.warn,
        `docx template: no docx template is set.`,
        "wd-to-docx",
        false
      );
      return;
    }

    const outPath = Path.resolve(
      Path.dirname(fileWd),
      Path.basename(fileWd) + ".docx"
    ); //"C:\\home\\docx_temp\\___output.docx";

    if (!option.isOverWrite && (await fileExists(outPath))) {
      option.message?.(
        MessageType.warn,
        `docx exists: ${outPath}.`,
        "wd-to-docx",
        false
      );
      return;
    }

    option.message?.(
      MessageType.info,
      `create docx: ${outPath}.`,
      "wd-to-docx",
      false
    );

    // render by vba(vbs)
    try {
      await wdToDocxJs(wdBody, template, outPath, Path.dirname(fileWd), option);
    } catch (e) {
      option.message?.(
        MessageType.warn,
        `wdToDocxJs err: ${e}.`,
        "wd-to-docx",
        false
      );
      return;
    }

    if (!option.isOpenWord) {
      return;
    }

    option.message?.(
      MessageType.info,
      `open docx: ${outPath}.`,
      "wd-to-docx",
      false
    );

    const wordExe = await selectExistsPath(
      [
        option.wordPath ?? "",
        "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
        "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
      ],
      ""
    );
    runCommand(wordExe, outPath);
    return;
  }

  const docxEngineInsideBody = getDocxEngineFromWd(wdBody);
  const thisEngine = docxEngineInsideBody
    ? docxEngineInsideBody
    : option.docxEngine;

  //docxEngine options
  option.message?.(
    MessageType.info,
    `docx docxEngine: ${thisEngine ? thisEngine : "use inside"}`,
    "wd-to-docx",
    false
  );

  return await vbsSpawn(
    option.docxEngine ?? "wordDownToDocx.vbs",
    option.timeOut ?? 600000,
    [
      fileWd,
      option.docxTemplate ?? "",
      option.mathExtension ? "1" : "0",
      option.isDebug ? "1" : "0",
    ],
    option.ac,
    option.message
  );
}

function getDocxEngineFromWd(wd: string) {
  const testMatch = wd.match(/^docxEngine\t(?<docxEngine>.*)\t/i);
  const docxEngine = testMatch?.groups?.docxEngine ?? "";
  return docxEngine;
}

function getDocxTemplateFromWd(wd: string) {
  const testMatch = wd.match(/^docxTemplate\t(?<docxTemplate>.*?)\t/im);
  const docxTemplate = testMatch?.groups?.docxTemplate ?? "";
  return docxTemplate;
}
