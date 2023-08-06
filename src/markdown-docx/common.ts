import path = require("path");
import * as Fs from "fs";
import Encoding = require("encoding-japanese");
import * as fs from "fs";
import * as iconv from "iconv-lite";
import { spawn } from "child_process";

const FOLDER_VBS = "vbs";

// front matter is not used
export type XFrontMatter = {
  dummy?: string;
};

export type DocxOption = {
  timeOut?: number;
  docxEngine?: string;
  docxTemplate?: string;
  mathExtension?: boolean;
  isDebug?: boolean;
  logInterval?: number;
  isUseDocxJs?: boolean;
  ac?: AbortController;
  isOverWrite?: boolean;
  wordPath?: string;
  isOpenWord?: boolean;
  message?: ShowMessage;
};

export function createDocxOption(option: DocxOption = {}) {
  const r = {
    timeOut: option.timeOut ?? 60000,
    docxEngine: option.docxEngine ?? "",
    docxTemplate: option.docxTemplate ?? "",
    mathExtension: option.mathExtension ?? true,
    isDebug: option.isDebug ?? false,
    logInterval: option.logInterval ?? 10,
    isUseDocxJs: option.isUseDocxJs ?? true,
    isOverWrite: option.isOverWrite ?? false,
    wordPath: option.wordPath ?? "",
    isOpenWord: option.isOpenWord ?? false,
    ac: option.ac,
    message: option.message,
  };
  return r;
}

// eslint-disable-next-line @typescript-eslint/naming-convention
export const MessageType = {
  info: "info",
  warn: "warn",
  err: "err",
  debug: "debug"
} as const;

export type MessageType = (typeof MessageType)[keyof typeof MessageType];

export type MessageProp = {
  msgType: MessageType,
  message: unknown,
  source: string,
  showNotification?: boolean
};


export type ShowMessage = (
  msgType: MessageType,
  message: unknown,
  source: string,
  showNotification?: boolean
) => void;

export type UpdateStatusBar = (isRunning: boolean) => void;
export const docxTemplate001 = "_template_001.docx";
export const templatesPath = "templates";

export async function createDocxTemplateFile(wfFsPath:string) {
  const folderOut =wfFsPath;
  const fileOut = path.resolve(folderOut, docxTemplate001);
  if (
    (await dirExists(folderOut)) &&
    !(await fileExists(fileOut))
  ) {
    fs.copyFileSync(
      path.resolve(__dirname, `../${templatesPath}/${docxTemplate001}`),
      fileOut
    );
  } else {
    throw new Error(`can not create a docx template!!.`);
  }
}

export function getFileContents(filePath: string) {
  const outLines: string[] = [];
  const dirPath = path.dirname(filePath);
  const buffer = Fs.readFileSync(filePath);
  let fileContents = Encoding.convert(buffer, {
    to: "UNICODE",
    type: "string",
  });
  
  // bom
  if (fileContents.charCodeAt(0) === 0xfeff) {
    fileContents = fileContents.substring(1);
  }

  const lines = fileContents.split(/\r?\n/g);

  for (let index = 0; index < lines.length; index++) {
    const line = lines[index];
    const wdCommand = getWordDownCommand(line);
    if (wdCommand?.command === "import") {
      outLines.push(
        getFileContents(path.resolve(dirPath, wdCommand?.params[0]))
      );
    } else {
      outLines.push(line);
    }
  }
  return outLines.join("\n");
}

export function getWordDownCommand(wd: string) {
  const testMatch = wd.match(/^<!--(?<name>.*)-->/i);
  const command = testMatch?.groups?.name ?? "";
  const commandList = command.trim().split(/\s(?=(?:[^"]*"[^"]*")*[^"]*$)/i);
  if (commandList[0].match(/word|ppt|oox/i) && commandList[1]) {
    const params = (commandList.slice(2) ?? []).map((l) =>
      l.replace(/\"/g, "")
    );
    return { command: commandList[1], params };
  }
  return undefined;
}

// wd: <!-- c m 1 3 xxx -->
// c: command
// m: merge
// row: 1
// column : 3
// xxx: param
export function getWordDownMergeCommand(wd: string) {
  const testMatch = wd.match(/<!--(?<name>.*)-->/i);
  const command = testMatch?.groups?.name ?? "";
  const commandList = command.trim().split(/\s(?=(?:[^"]*"[^"]*")*[^"]*$)/gi);
  if (commandList[0] === "c" && commandList[1]) {
    const params = (commandList.slice(2) ?? []).map((l) =>
      l.replace(/\"/g, "")
    );
    const isMergeRow = params.includes("^");
    const isMergeColumn = params.includes("<");
    return { isMergeRow, isMergeColumn };
  }
  return undefined;
}

/**
 * create uniq path
 * @param dir 
 * @param name 
 * @param ext 
 * @param isSame 
 * @returns 
 */
export async function createPath(
  dir: string,
  name: string,
  ext: string,
  isSame = false
) {
  for (let index = 0; index < 1000; index++) {
    const filePath =
      path.resolve(dir, name + (index > 0 ? index.toString() : "")) + "." + ext;
    if (isSame) {
      return filePath;
    }
    if (!(await fileExists(filePath)) && !(await dirExists(filePath))) {
      return filePath;
    }
  }
  throw new Error(`Can not create a file: ${name}.${ext}`);
}

export async function fileExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isFile();
    return res;
  } catch (e) {
    return false;
  }
}

export async function dirExists(filepath: string) {
  try {
    const res = (await fs.promises.lstat(filepath)).isDirectory();
    return res;
  } catch (e) {
    return false;
  }
}

/**
 * remove pathFolder
 * @param pathFolder 
 * @param option  option: { force: true }
 * @returns 
 */
export async function rmDirIfExist(pathFolder: string, option: {}) {
  try {
    const isExist = await dirExists(pathFolder);
    if (!isExist) {
      // no folder no delete
      return;
    }
    await fs.promises.rm(pathFolder, option);
  } catch (error) {
    throw error;
  }
}

/**
 * remove pathFile
 * @param pathFile 
 * @param option option: { force: true }
 * @returns void
 */
export async function rmFileIfExist(pathFile: string, option: { force: true }) {
  try {
    const isExist = await fileExists(pathFile);
    if (!isExist) {
      // no file no delete
      return;
    }
    await fs.promises.rm(pathFile, option);
  } catch (error) {
    throw error;
  }
}

function s2u(sb: Buffer) {
  // todo for non japanese language
  //const vbsEncode = vscode.workspace.getConfiguration("vbecm").get<string>("vbsEncode") || "windows-31j";
  // https://github.com/ashtuchkin/iconv-lite/wiki/Use-Buffers-when-decoding
  return iconv.decode(sb, "windows-31j");

  // if japanese, select below
  // const r = Encoding.convert(sb, {
  //   to: "UNICODE",
  //   type: "string",
  // });
  // return r;
}

export type VbsSpawn = typeof vbsSpawn;
export function vbsSpawn(
  script: string,
  timeout: number,
  param: string[],
  ac?: AbortController,
  showMessage?: ShowMessage
) {
  return new Promise<number>(async (resolve, reject) => {
    const { signal } = ac ?? new AbortController();

    let scriptPath = "";
    if (await fileExists(script)) {
      // optional docxEngine
      scriptPath = script;
    } else {
      const rootFolder = path.dirname(__dirname);
      // rootFolder is differ between debug and release
      // for release
      let vbsPath = path.resolve(rootFolder, FOLDER_VBS);
      if (!(await dirExists(vbsPath))) {
        // for debug
        vbsPath = path.resolve(rootFolder, "..", FOLDER_VBS);
      }
      scriptPath = path.resolve(vbsPath, "wordDownToDocx.vbs");
    }

    if (!(await fileExists(scriptPath))) {
      return reject(9991);
    }

    const p = spawn("cscript.exe", ["//Nologo", scriptPath, ...param], {
      timeout: timeout,
      signal,
    });

    p.stdout.on("data", (data) => {
      const r = s2u(data);
      r.split("\n")
        .filter((d) => d.trim())
        .forEach((d) => showMessage?.(MessageType.info, d, "vbs"));
    });
    p.stderr.on("data", (data) => {
      const r = s2u(data as Buffer);
      r.split("\n")
        .filter((d) => d.trim())
        .forEach((d) => showMessage?.(MessageType.err, d, "vbs"));
    });
    p.on("close", (code) => {
      const r = code ?? 9999;
      if (r === 0) {
        showMessage?.(MessageType.info, "complete!!", "vbs");
      } else if (ac?.signal.aborted) {
        showMessage?.(MessageType.info, "convert is aborted.", "common");
      } else {
        showMessage?.(
          MessageType.err,
          `some error happens. code: ${r} killed? : ${p.killed}`,
          "vbs"
        );
        return reject(r);
      }
      return resolve(r);
    });

    const cleanup = () => {
      showMessage?.(MessageType.info, `spawn kill pid: ${p.pid}`, "common");
      p.kill();
    };

    // for windows, they do not work. may be
    p.on("SIGINT", cleanup);
    p.on("SIGTERM", cleanup);
    p.on("SIGQUIT", cleanup);
  });
}


