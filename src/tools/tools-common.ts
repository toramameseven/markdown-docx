import { spawn } from "child_process";

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
