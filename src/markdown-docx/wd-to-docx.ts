import {
  DocxOption,
  getFileContents,
  MessageType,
  fileExists,
  docxTemplate001,
  templatesPath,
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

  // search template
  const defaultTemplate = Path.resolve(__dirname, `../${templatesPath}`);
  const defaultTemplate2 = Path.resolve(__dirname, `../../${templatesPath}`);
  const templateInsideDocument = getDocxTemplateFromWd(wdBody);

  const template = await selectExistsPath(
    [
      templateInsideDocument,
      option.docxTemplate ?? "",
      docxTemplate001,
    ],
    [Path.dirname(fileWd), defaultTemplate, defaultTemplate2]
  );

  option.message?.(
    MessageType.info,
    `Used docxTemplate: ${template ? template : "use inside"}`,
    "wd-to-docx",
    false
  );

  // create docx from docxJs
  if (!template) {
    option.message?.(
      MessageType.warn,
      `docx template: no docx template is set.`,
      "wd-to-docx",
      true
    );
    return;
  }

  // create path for output
  const outPath = Path.resolve(
    Path.dirname(fileWd),
    Path.basename(fileWd) + ".docx"
  ); //"C:\\home\\docx_temp\\___output.docx";

  // file existence test
  if (!option.isOverWrite && (await fileExists(outPath))) {
    option.message?.(
      MessageType.warn,
      `docx exists: ${outPath}.`,
      "wd-to-docx",
      true
    );
    return;
  }

  option.message?.(
    MessageType.info,
    `create docx: ${outPath}.`,
    "wd-to-docx",
    false
  );

  // render
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

  // open the docx file
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
    [""]
  );
  runCommand(wordExe, outPath);
  return;
}

function getDocxTemplateFromWd(wd: string) {
  const testMatch = wd.match(/^param\tdocxTemplate\t(?<docxTemplate>.*?)\t/im);
  const docxTemplate = testMatch?.groups?.docxTemplate ?? "";
  return docxTemplate;
}
