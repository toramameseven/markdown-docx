import {
  DocxOption,
  getFileContents,
  MessageType,
  fileExists,
  selectExistsPath,
  vbsSpawn,
} from "./common";
import { wdToDocxJs } from "./wd-to-docxjs";
import * as Path from "path";
import { runCommand } from "../tools/tools-common";

export async function wordDownToDocx(fileWd: string, option: DocxOption) {
  const wdBody = getFileContents(fileWd);
  wordDownToDocxBody(fileWd, wdBody, option);
}

export async function wordDownToDocxBody(
  fileWd: string,
  wdBody: string,
  option: DocxOption
) {
  const defaultTemplate = Path.resolve(
    __dirname,
    "../vbs/sample-heder-js.docx"
  );
  const defaultTemplate2 = Path.resolve(
    __dirname,
    "../../vbs/sample-heder-js.docx"
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
    "main",
    false
  );

  if (option.isUseDocxJs) {
    // create docx from docxJs
    if (!template) {
      option.message?.(
        MessageType.warn,
        `docx template: no docx template is set.`,
        "main",
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
        "main",
        false
      );
      return;
    }

    // render by vba(vbs)
    try {
      await wdToDocxJs(wdBody, template, outPath, Path.dirname(fileWd), option);
    } catch (e) {
      option.message?.(
        MessageType.warn,
        `wdToDocxJs err: ${e}.`,
        "main",
        false
      );
      return;
    }

    if (!option.isOpenWord) {
      return;
    }

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
    "main",
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
