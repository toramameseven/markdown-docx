import { getFileContents } from "./tools-common";
// const toTextile = require('to-textile');
const toHtml = require('textile-js');
// https://github.com/borgar/textile-js

// export function htmlToTextile(filePath: string, body: string){
//   let s = body;
//   if (s === ""){
//     s = getFileContents(filePath);
//   }
//   const textile = toTextile(s);
//   return textile;
// }

export function textileToHtml(filePath:string, body:string){
  let s = body;
  if (s === ""){
    s = getFileContents(filePath);
  }
  const html = toHtml(s);
  // console.log("========================>");
  // console.log(html);
  return html;
}
