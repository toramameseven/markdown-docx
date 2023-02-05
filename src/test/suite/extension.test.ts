import * as assert from "assert";
import { markdownToWd0 } from "../../markdown-docx/markdown-to-wd0";
import { wd0ToDocx as wd0ToWd } from "../../markdown-docx/wd0-to-wd";

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
//import * as vscode from "vscode";
// import * as myExtension from '../../extension';

function mdToWd(marked: string) {
  const markdownBodyX = marked.replace(/@/g, "\\@");

  const wd0 = markdownToWd0(markdownBodyX, { sanitize: false }).replace(
    /(\r?\n)+/g,
    "\n"
  );
  // wd must vbCRLF
  const wd = wd0ToWd(wd0).replace("\r", "");
  return wd;
}

suite("Extension Test Suite", () => {
  //vscode.window.showInformationMessage("Start all tests.");

  test("init", () => {
    assert.strictEqual(-1, [1, 2, 3].indexOf(5));
  });

  //
  test("title", () => {
    assert.strictEqual(
      mdToWd("<!-- word title Markdown to Docx samples -->"),
      `title\tMarkdown to Docx samples\nnewLine\tconvertTitle\ttm`
    );
  });

  //
  test("toc", () => {
    assert.strictEqual(
      mdToWd("<!-- word toc 1 -->"),
      `toc	1
newLine	convertToc	tm`
    );
  });
  //
  test("heading1", () => {
    assert.strictEqual(
      mdToWd("# title"),
      `section	1	title	title
newLine	convertHeading	tm`
    );
  });
  //
  test("heading5", () => {
    assert.strictEqual(
      mdToWd("##### title"),
      `section	5	title	title
newLine	convertHeading	tm`
    );
  });
  //
  test("hr", () => {
    assert.strictEqual(
      mdToWd(`___

---

***`),
      `hr	
hr	
hr	`
    );
  });

  //
  test("BR", () => {
    // marked
    const marked = `
next line is \`<br>\`
<br>
upper line is \`<br>\``;
    // expect
    const expect = `text	next line is <codespan><br></codespan>
newLine	convertHtml <BR>	tm
text	upper line is <codespan><br></codespan>
newLine	convertParagraph	tm`;

    assert.strictEqual(mdToWd(marked), expect);
  });

  //
  test("new page", () => {
    // marked
    const marked = `<!-- word newPage -->`;
    // expect
    const expect = `newPage				tm`;
    // assert
    assert.strictEqual(mdToWd(marked), expect);
  });

  //
  test("at mark", () => {
    // marked
    const marked = `anonymous@com.com`;
    // expect
    const expect = "text\tanonymous@com.com\nnewLine\tconvertParagraph\ttm";
    // assert
    assert.strictEqual(mdToWd(marked), expect);
  });

  //
  test("emphasis1", () => {
    // marked
    const marked = `
**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>`;
    // expect
    const expect = `text	<b>This is bold text</b>
newLine	convertParagraph	tm
text	<b>This is bold text</b>
newLine	convertParagraph	tm
text	<i>This is italic text</i>
newLine	convertParagraph	tm
text	<i>This is italic text</i>
newLine	convertParagraph	tm
text	<~~>Strikethrough</~~>
newLine	convertParagraph	tm
text	2
text	<sup>
text	x
text	</sup>
text	<sub>
text	y
text	</sub>
newLine	convertParagraph	tm`;
    // assert
    assert.strictEqual(mdToWd(marked), expect);
  });

  //
  test("XXX", () => {
    // marked
    const marked = ``;
    // expect
    const expect = ``;
    // assert
    assert.strictEqual(mdToWd(marked), expect);
  });

  //
});
