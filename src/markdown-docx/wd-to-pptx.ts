import { fileExists, selectExistsPath, vbsSpawn } from "./common";
import { DocxOption, getFileContents, MessageType } from "./common";
import { wdToDocxJs } from "./wd-to-docxjs";
import * as child_process from "child_process";
import * as util from "util";
import * as Path from "path";
import { wdToPptxJs } from "./wd-to-pptxJs";

async function runChildProc(exe: string, param: string) {
  const execFile = util.promisify(child_process.execFile);
  execFile(exe, [param])
    .then(() => {
      console.log("Successfully executed.");
    })
    .catch((err: any) => {
      console.log("Error!");
      console.log(err);
    });
}

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
    option.message?.(MessageType.warn, `wdToPptxJs err: ${e}.`, "main", false);
    return;
  }
  return;
}

