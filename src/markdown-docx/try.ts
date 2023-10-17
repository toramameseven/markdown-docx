//import { run } from "@mermaid-js/mermaid-cli";
// const runx = require("@mermaid-js/mermaid-cli");
const util = require("util");
import childProcess = require("child_process");
const exec = util.promisify(childProcess.exec);

// const exe = childProcess.exec("dir", (error, stdout, stderr) => {
//   if (error) return console.error("ERROR", error);
//   console.log("STDOUT", stdout); // string
//   console.log("STDERR", stderr); // string
// });

// function main2() {
//   console.log("main", "start"); // string
//   exe.stdin?.write("");
//   exe.stdin?.end();
//   console.log("main", "end"); // string
// }

async function main() {
  exec.stdin?.write("--input echo--");
  //exec.stdin?.end();
  let res = await exec("echo");
  console.log(res.stdout);

  // await exec("mkdir tempdir");

  // let res2 = await exec("dir");
  // console.log(res2.stdout);
}

// main().catch((e) => console.log(e));

function spawn (cmd:string, args?: string[]){
  return new Promise<void>((resolve)=>{
    //let p = childProcess.spawn(cmd,args);

    const p = childProcess.exec(cmd, (error, stdout, stderr) => {
      if (error) {return console.error("ERROR", error);}
      console.log("STDOUT", stdout); // string
      console.log("STDERR", stderr); // string
    });

    p.stdin?.write("graph TD\nA[Client] --> B[Load Balancer]");
    p.stdin?.end();
    p.on('exit', (code)=>{
      resolve();
    });
    // p.stdout.setEncoding('utf-8');
    // p.stdout.on('data', (data)=>{
    //   console.log(data);
    // });
    // p.stderr.on('data', (data)=>{
    //   console.log(data);
    // });
  });
}

async function nnmain(){
  await spawn('mmdc -i - ');
  // await spawn('mkdir',['newtmp']);
  // await spawn('ls');
}

nnmain().catch(e=>console.log(e));


