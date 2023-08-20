import * as assert from "assert";
import { markdownToWd0 } from "../markdown-docx/markdown-to-wd0";
import { wd0ToDocx as wd0ToWd } from "../markdown-docx/wd0-to-wd";
import { suite, test } from "mocha";
import { markdownToWd } from "../../src/markdown-docx/markdown-to-xxxx";
import path = require("path");
import * as Fs from "fs";
import { getFileContents } from "../markdown-docx/common";
import { addTableSpanToMarkdown } from "../markdown-docx/add-table-span";
//import { assert } from "chai";
// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
//import * as vscode from "vscode";
// import * as myExtension from '../../extension';

async function mdToWd(marked: string) {
  const markdownBodyX = marked.replace(/@/g, "\\@");

  const wd0 = (
    await markdownToWd0(markdownBodyX, "docx", { sanitize: false })
  ).replace(/(\r?\n)+/g, "\n");
  // wd must vbCRLF
  const wd = wd0ToWd(wd0).replace("\r", "");
  return wd;
}

async function mdFileToWd(file: string) {
  const r = await markdownToWd(file, "", "docx", 0, false, true);
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
  test("toc", async () => {
    // marked
    const markdown = `
<!-- word toc 1 "table of content"-->`;

    // expect
    const expect = `
toc\t1\ttable of content\ttm
newLine\tconvertToc\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("heading1", async () => {
    // marked
    const markdown = `
# heading1`;

    // expect
    // command title    idTitle
    // section heading1 heading1 (if no id, same to title)
    const expect = `
section\t1\theading1
text\theading1
newLine\tconvertHeading End\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("heading6", async () => {
    // marked
    const markdown = `
###### heading6`;

    // expect
    // command title    idTitle
    // section heading1 heading1 (if no id, same to title)
    const expect = `
section\t6\theading6
text\theading6
newLine\tconvertHeading End\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("hr", async () => {
    // marked
    const markdown = `
___

---

***`;

    // expect
    const expect = `
hr\t
hr\t
hr\t`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("br", async () => {
    // marked
    const markdown = `
next line is \`<br>\`
<br>
upper line is \`<br>\``;

    // expect
    const expect = `
text\tnext line is <codespan><br></codespan>
newLine\twd0NewLine\t\ttm
text\tupper line is <codespan><br></codespan>
newLine\tconvertParagraph\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("new page", async () => {
    // marked
    const markdown = `
<!-- word newPage -->`;

    // expect
    const expect = `
newPage\t\t\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("at mark", async () => {
    // marked
    const markdown = `
anonymous@com.com`;

    // expect
    const expect = `
text\tanonymous@com.com
newLine\tconvertParagraph\t\ttm`;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("emphasis1", async () => {
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
text\t<b>This is bold text</b>
newLine\tconvertParagraph\t\ttm
text\t<b>This is bold text</b>
newLine\tconvertParagraph\t\ttm
text\t<i>This is italic text</i>
newLine\tconvertParagraph\t\ttm
text\t<i>This is italic text</i>
newLine\tconvertParagraph\t\ttm
text\t<~~>Strikethrough</~~>
newLine\tconvertParagraph\t\ttm
text\t2<sup>x</sup><sub>y</sub>
newLine\tconvertParagraph\t\ttm`;
    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("addTableSpanToMarkdown", async () => {
    // marked
    const markdown = `
<!-- word emptyMerge -->
<!-- word rowMerge 2-4  -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |
`;

    // expect
    const expect = `
<!-- word emptyMerge -->
<!-- word rowMerge 2-4  -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|          ^| data3-2                  ^|
| data4-1  ^| <!-- not merged -->      ^|
`;

    // assert
    assert.strictEqual(
      await addTableSpanToMarkdown("", removeTopN(markdown), undefined),
      removeTopN(expect)
    );
  });

  // ==========================================
  test("XXXXX", async () => {
    // marked
    const markdown = ``;

    // expect
    const expect = ``;

    // assert
    assert.strictEqual(await mdToWd(removeTopN(markdown)), removeTopN(expect));
  });

  // ==========================================
  test("demo-admonition", async () => {
    // marked
    const mdFile = "demo-admonition";
    const mdPath = path.resolve(__dirname, "../../md_demo/md", mdFile + ".md");
    const wd = await mdFileToWd(mdPath);

    // expect
    const wdPath = path.resolve(__dirname, "../../md_demo/wd", mdFile + ".wd");
    const expect = getFileContents(wdPath);

    // assert
    assert.strictEqual(wd, expect);
  });

  //
});

/**
 *
 */
suite("Demo Test Suite", () => {
  const r = path.resolve(__dirname, "../../md_demo/md");

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
      const mdPath = path.resolve(
        __dirname,
        "../../md_demo/md",
        mdFile + ".md"
      );
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
