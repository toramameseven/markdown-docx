const { inlineSource } = require("inline-source");
//import { inlineSource } from "inline-source";
//import * as inlineSource from "inline-source";

import * as path from "path";

export async function createInlineHtml(htmlPath: string, htmlBody: string) {
  try {
    const r = await inlineSource(htmlBody, {
      compress: false,
      attribute: false,
      rootpath: path.dirname(htmlPath),
    });
    return r as string;
  } catch (error) {
    console.log(error);
  } 
  return "";
}
