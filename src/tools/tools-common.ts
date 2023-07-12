import { spawn } from "child_process";
import Encoding = require("encoding-japanese");
import path = require("path");
import * as Fs from "fs";

/**
 * run windows process
 * @param exe full path of the exe file.
 * @param params parameters for exe
 */
export function runCommand(exe: string, params: string) {
  const child = spawn(exe, [params], {
    stdio: "ignore",
    detached: true,
    env: process.env,
  });
  child.unref();
}

export function getFileContents(filePath: string) {
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

  return lines.join("\n");
}
