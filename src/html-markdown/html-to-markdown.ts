import {
  NodeHtmlMarkdown,
  // NodeHtmlMarkdownOptions,
  // TranslatorConfig,
  // TranslatorConfigFactory,
  // TranslatorCollection,
  // PostProcessResult,
  // TranslatorConfigObject,
} from ".";

//import { wordDownTranslators } from "./wordDownConfig";

import * as Fs from "fs";
import { getFileContents } from "../markdown-docx/common";
/* ********************************************************* *
 * Single use
 * If using it once, you can use the static method
 * https://github.com/crosstype/node-html-markdown
 * ********************************************************* */
export function htmlToMarkdown(filePath: string, body: string) {
  let s = body;
  if (s === "") {
    s = getFileContents(filePath);
  }
  // Single file
  const r = NodeHtmlMarkdown.translate(
    /* html `<b>hello</b>`*/ s,
    /* options (optional) */ {},
    /* customTranslators (optional) */ undefined,
    /* customCodeBlockTranslators (optional) */ undefined
  );

  if (filePath !== "") {
    Fs.writeFileSync(filePath + ".md", r);
  }
  return r;
}
