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
export function htmlToMarkdown(filePath: string) {
  const s = getFileContents(filePath);
  // Single file
  const r = NodeHtmlMarkdown.translate(
    s /* html */,
    ///* html */ `<b>hello</b>`,
    /* options (optional) */ {},
    /* customTranslators (optional) */ undefined,
    /* customCodeBlockTranslators (optional) */ undefined
  );
  Fs.writeFileSync(filePath + ".md", r);
  return r;
}
