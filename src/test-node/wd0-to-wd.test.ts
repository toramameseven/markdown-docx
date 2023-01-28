import * as assert from "assert";
import { markdownToWd0 } from "../markdown-docx/markdown-to-wd0";
import { wd0ToDocx as wd0ToWd } from "../markdown-docx/wd0-to-wd";
import { suite, test } from "mocha";
import { markdownToWd } from "../../src/markdown-docx/markdown-to-docx";
import path = require("path");
import * as Fs from "fs";
import { getFileContents } from "../markdown-docx/common";
//import { assert } from "chai";
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

async function mdFileToWd(file: string) {
  const r = await markdownToWd(file, "", 0, false, true);
  return r.wdBody;
}

function removeTopN(s: string) {
  if (s[0] === "\n") {
    return s.slice(1);
  }
  return s;
}

suite("Extension Test Suite", () => {
  // ==========================================
  test("title", () => {
    // marked
    const markdown = `
<!-- word title Markdown to Docx samples -->`;

    // expect
    const expect = `
title\tMarkdown to Docx samples
newLine\tconvertTitle\ttm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("toc", () => {
    // marked
    const markdown = `
<!-- word toc 1 -->`;

    // expect
    const expect = `
toc	1	TOC	tm
newLine	convertToc	tm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("heading1", () => {
    // marked
    const markdown = `
# heading1`;

    // expect
    // command title    idTitle
    // section heading1 heading1 (if no id, same to title)
    const expect = `
section	1	heading1	heading1
newLine	convertHeading	tm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("heading6", () => {
    // marked
    const markdown = `
###### heading6`;

    // expect
    // command title    idTitle
    // section heading1 heading1 (if no id, same to title)
    const expect = `
section	6	heading6	heading6
newLine	convertHeading	tm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("hr", () => {
    // marked
    const markdown = `
___

---

***`;

    // expect
    const expect = `
hr	
hr	
hr	`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("br", () => {
    // marked
    const markdown = `
next line is \`<br>\`
<br>
upper line is \`<br>\``;

    // expect
    const expect = `
text	next line is <codespan><br></codespan>
newLine	wd0NewLine	tm
text	upper line is <codespan><br></codespan>
newLine	convertParagraph	tm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("new page", () => {
    // marked
    const markdown = `
<!-- word newPage -->`;

    // expect
    const expect = `
newPage				tm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("at mark", () => {
    // marked
    const markdown = `
anonymous@com.com`;

    // expect
    const expect = `
text\tanonymous@com.com
newLine\tconvertParagraph\ttm`;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("emphasis1", () => {
    // marked
    const markdown = `
**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>`;

    // expect
    const expect = `
text	<b>This is bold text</b>
newLine	convertParagraph	tm
text	<b>This is bold text</b>
newLine	convertParagraph	tm
text	<i>This is italic text</i>
newLine	convertParagraph	tm
text	<i>This is italic text</i>
newLine	convertParagraph	tm
text	<~~>Strikethrough</~~>
newLine	convertParagraph	tm
text	2<sup>x</sup><sub>y</sub>
newLine	convertParagraph	tm`;
    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("XXXXX", () => {
    // marked
    const markdown = ``;

    // expect
    const expect = ``;

    // assert
    assert.strictEqual(mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("demo-admonition", async () => {
    // marked
    const mdFile = "demo-admonition";
    const mdPath = path.resolve(__dirname, "../../md_demo", mdFile + ".md");
    const wd = await mdFileToWd(mdPath);

    // expect
    const wdPath = path.resolve(__dirname, "../../md_demo/wd", mdFile + ".wd");
    const expect = getFileContents(wdPath);

    // assert
    assert.strictEqual(wd, expect);
  });

  //
});

suite("Demo Test Suite", () => {
  const r = path.resolve(__dirname, "../../md_demo");

  Fs.readdir(r, (err, files) => {
    files
      .filter((f) => f.match(/\.md$/i))
      .forEach(async (file) => {
        const baseName = path.basename(file).replace(/\.md$/i, "");
        testWd(baseName);
      });
  });

  function testWd(mdBaseName: string) {
    // ==========================================
    test(mdBaseName, async () => {
      // marked
      const mdFile = mdBaseName;
      const mdPath = path.resolve(__dirname, "../../md_demo", mdFile + ".md");
      const wd = await mdFileToWd(mdPath);

      // expect
      const wdPath = path.resolve(
        __dirname,
        "../../md_demo/wd",
        mdFile + ".wd"
      );
      const expect = getFileContents(wdPath);

      // assert
      assert.strictEqual(wd, expect);
    });
  }
  //
});
