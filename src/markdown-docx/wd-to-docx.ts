import { vbsSpawn } from "./common";
import { DocxOption, getFileContents, MessageType } from "./common";

export async function wordDownToDocx(fileWd: string, option: DocxOption) {
  // get docxEngine and docxTemplate
  const wdBody = getFileContents(fileWd);
  const docxEngineInsideBody = getDocxEngineFromWd(wdBody);
  const templateInsideBody = getDocxTemplateFromWd(wdBody);

  const thisEngine = docxEngineInsideBody
    ? docxEngineInsideBody
    : option.docxEngine;
  const thisTemplate = templateInsideBody
    ? templateInsideBody
    : option.docxTemplate;

  //docxEngine options
  option.message?.(
    MessageType.info,
    `docx docxEngine: ${thisEngine ?? "use inside"}`,
    "main",
    false
  );
  option.message?.(
    MessageType.info,
    `docx docxTemplate: ${thisTemplate ?? "use inside"}`,
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
  const testMatch = wd.match(/^docxTemplate\t(?<docxTemplate>.*)\t/im);
  const docxTemplate = testMatch?.groups?.docxTemplate ?? "";
  return docxTemplate;
}
