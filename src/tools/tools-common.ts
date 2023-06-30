import { spawn } from "child_process";

export function runCommand(exe: string, param: string) {
  const child = spawn(exe, [param], {
    stdio: "ignore", // piping all stdio to /dev/null
    detached: true, // メインプロセスから切り離す設定
    env: process.env, // NODE_ENV を tick.js へ与えるため
  });
  child.unref(); // メインプロセスから切り離す
}