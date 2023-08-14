const { inlineSource } = require("inline-source");
//import { inlineSource } from "inline-source";
//import * as inlineSource from "inline-source";

import * as path from "path";
import { getFileContents, selectExistsPath } from "./tools-common";

export async function createInlineHtml(htmlPath: string, htmlBody: string) {
  try {
    const r = await inlineSource(htmlBody, {
      compress: false,
      attribute: false,
      rootpath: path.dirname(htmlPath),
    });

    const htmlTemplatePath = await selectExistsPath(
      ["../../templates/sample.html", "../templates/sample.html"],
      [__dirname]
    );

    let sampleHtml = getFileContents(htmlTemplatePath);

    const outHtml = sampleHtml.replace("[[html main contents]]", r);

    return outHtml;
  } catch (error) {
    console.log(error);
  }
  return "";
}
