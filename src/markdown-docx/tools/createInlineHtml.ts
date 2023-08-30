const { inlineSource } = require("inline-source");
//import { inlineSource } from "inline-source";
//import * as inlineSource from "inline-source";

import * as path from "path";
import { getFileContents, selectExistsPath } from "./tools-common";

export async function createInlineHtml(htmlPath: string, htmlBody: string) {
  try {
    const r:string = await inlineSource(htmlBody, {
      compress: false,
      attribute: false,
      rootpath: path.dirname(htmlPath),
    });

    const htmlTemplatePath = await selectExistsPath(
      ["../../templates/sample.html", "../templates/sample.html"],
      [__dirname]
    );

    let sampleHtml = getFileContents(htmlTemplatePath);

    const outHtml = sampleHtml.replace("673ab4838a71448692a56d601b77c818", r);

    return outHtml;
  } catch (error) {
    throw(error);
  }
  return "";
}
