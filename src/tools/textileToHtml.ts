import { getFileContents } from "./tools-common";

const textile = require('textile-js');

export function textileToHtml(file:string, body:string){
  let s = body;
  if (s === ""){
    s = getFileContents(file);
  }
  const r = textile(s);
  return r;
}

