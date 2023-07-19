import {
  NodeHtmlMarkdown,
  // NodeHtmlMarkdownOptions,
  // TranslatorConfig,
  // TranslatorConfigFactory,
  // TranslatorCollection,
  // PostProcessResult,
  // TranslatorConfigObject,
} from ".";

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

function testme() {
  const source = `
  <h1>Lists ordered</h1>
<ol>
	<li>Lorem ipsum dolor sit amet fffffa2<sup>x</sup><sub>y</sub>affff</li>
	<li>Consectetur <b>adipiscing</b> elit</li>
</ol>
  `;

  console.log(htmlToMarkdown("", source));
}

//testme();
