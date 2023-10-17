const childProcess = require("child_process");

function spawn(cmd: string, args: string[] | undefined = undefined) {
  return new Promise<void>((resolve) => {
    let p = childProcess.spawn(cmd, args);
    p.on("exit", (code: any) => {
      resolve();
    });
    p.stdout.setEncoding("utf-8");
    p.stdout.on("data", (data: any) => {
      console.log(data);
    });
    p.stderr.on("data", (data: any) => {
      console.log(data);
    });
  });
}

async function nnmain() {
  await spawn("dir");
  // await spawn("mkdir", ["newtmp"]);
  // await spawn("ls");
}

nnmain().catch((e) => console.log(e));
