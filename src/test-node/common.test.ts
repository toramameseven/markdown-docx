import * as assert from "assert";
import { markdownToWd0 } from "../markdown-docx/markdown-to-wd0";
import { wd0ToDocx as wd0ToWd } from "../markdown-docx/wd0-to-wd";
import * as common from "../markdown-docx/common";
import { suite, test } from "mocha";
import { markdownToWd } from "../../src/markdown-docx/markdown-to-xxxx";
import path = require("path");
import * as Fs from "fs";
import { getFileContents } from "../markdown-docx/common";
import { addTableSpanToMarkdown } from "../markdown-docx/add-table-span";
import { isEnableExperimentalFeature } from "../common-settings";

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


/**
 * test parts
 */
suite("Common Test Suite", () => {
  // ==========================================
  test("word param", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word param sss dddd -->");

    // expect
    const expect = {
      command: 'param',
      params: [
        'sss',
        'dddd'
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param cols", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word cols 1,2,3,4 -->");

    // expect
    const expect = {
      command: 'cols',
      params: [
        '1,2,3,4'
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param rowMerge", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word rowMerge 1-3,4-5 -->");

    // expect
    const expect = {
      command: 'rowMerge',
      params: [
        '1-3,4-5'
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param newPage", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word newPage -->");

    // expect
    const expect = {
      command: 'newPage',
      params: [
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param newLine", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word newLine -->");

    // expect
    const expect = {
      command: 'newLine',
      params: [
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param toc", async () => {
    // marked
    const res = common.getWordDownCommand(`<!-- word toc 1 "table of contents" -->`);

    // expect
    const expect = {
      command: 'toc',
      params: [
        "1",
        "table of contents"
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param placeholder", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word placeholder date  113456 -->");

    // expect
    const expect = {
      command: 'placeholder',
      params: [
        'date',
        "113456"
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param placeholder with space", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word placeholder date  \"\" -->");

    // expect
    const expect = {
      command: 'placeholder',
      params: [
        'date',
        ""
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

  // ==========================================
  test("word param tables", async () => {
    // marked
    const res = common.getWordDownCommand("<!-- word param cols 1,2,3 \nemptyMerge 1 rowMerge\n 1-2,2-4 -->");

    // expect
    const expect = {
      command: 'param',
      params: [
        "cols",
        "1,2,3",
        "emptyMerge",
        "1",
        "rowMerge",
        "1-2,2-4"
      ]
    };

    // assert
    assert.deepEqual(res, expect);
  });

});
