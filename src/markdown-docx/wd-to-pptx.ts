import { DocxOption, getFileContents, MessageType } from "./common";
import * as Path from "path";
import { wdToPptxJs } from "./wd-to-pptxJs";

export async function wordDownToPptx(fileWd: string, option: DocxOption) {
  const wdBody = getFileContents(fileWd);
  wordDownToPptxBody(fileWd, wdBody, option);
}

export async function wordDownToPptxBody(
  fileWd: string,
  wdBody: string,
  option: DocxOption
) {
  // const outPath = Path.resolve(
  //   Path.dirname(fileWd),
  //   Path.basename(fileWd) + ".docx"
  // ); //"C:\\home\\docx_temp\\___output.docx";

  // if (!option.isOverWrite && (await fileExists(outPath))) {
  //   option.message?.(
  //     MessageType.warn,
  //     `docx exists: ${outPath}.`,
  //     "main",
  //     false
  //   );
  //   return;
  // }

  try {
    await wdToPptxJs(wdBody, "", "outPath", Path.dirname(fileWd), option);
  } catch (e) {
    option.message?.(MessageType.warn, `wdToPptxJs err: ${e}.`, "wd-to-pptx", false);
    return;
  }
  return;
}

