import * as Fs from "fs";
import Encoding = require("encoding-japanese");
import { NodeHtmlMarkdown } from ".";
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

function getFileContents(filePath: string) {
  const buffer = Fs.readFileSync(filePath);
  let fileContents = Encoding.convert(buffer, {
    to: "UNICODE",
    type: "string",
  });

  // bom
  if (fileContents.charCodeAt(0) === 0xfeff) {
    fileContents = fileContents.substring(1);
  }

  const lines = fileContents.split(/\r?\n/g);

  return lines.join("\n");
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
