import * as Fs from "fs";
import * as path from "path";
import { getWordTitle, slugify } from "../markdown-docx/markdown-to-wd0";
import {
  getFileContents,
  getWordDownCommand,
  fileExists,
  dirExists,
} from "../markdown-docx/common";

export async function splitMarkdownExportCmd(filePath: string) {
  const markdown = getFileContents(filePath);
  await splitByExport(markdown, path.dirname(filePath));
}

export async function splitMarkdownToHugo(filePath: string) {
  const markdown = getFileContents(filePath);
  console.log(markdown);
  await splitForHugo(markdown, path.dirname(filePath));
}

const exportSplit = "{B71BFF78-FB1B-491A-9A37-946FB2D9D738}";

async function splitForHugo(marked: string, pathToSave: string = "") {
  let markdowns = marked
    .replace(/(^#\s.*)/gim, exportSplit + "<!-- $1 -->")
    .split(exportSplit);

  if (markdowns.length === 1) {
    const docTitle = getWordTitle(marked);
    markdowns = (`# ${docTitle}\n` + marked)
      .replace(/(^#\s.*)/gim, exportSplit + "<!-- $1 -->")
      .split(exportSplit);
  }

  // create index markdown _index.md
  /*
  ---
  title: "XXXXXX"
  chapter: false
  weight: 10
  ---
  this is a hugo front matter.
  */

  const dirHugoMds = path.resolve(pathToSave, "forHugo");

  if (!(await dirExists(dirHugoMds))) {
    Fs.mkdirSync(dirHugoMds);
  }

  for (let index = 0; index < markdowns.length; index++) {
    const element = markdowns[index];
    const title = getHeading1Title(element);
    if (title !== "") {
      // has section
      const fileName = path.resolve(dirHugoMds, slugify(title) + ".md");
      console.log(`try save: ${fileName}`);
      if (!(await fileExists(fileName))) {
        Fs.writeFileSync(fileName, createFrontMatter(title, index) + element);
      } else {
        console.log(`${fileName} already exists!!!`);
      }
    } else {
      // no section (title or hugo index)
      const docTitle = getWordTitle(element);
      const fileName = path.resolve(dirHugoMds, "_index.md");
      console.log(`try save: ${fileName}`);
      if (!(await fileExists(fileName))) {
        Fs.writeFileSync(
          fileName,
          createFrontMatter(docTitle, index) + element
        );
      } else {
        console.log(`${fileName} already exists!!!`);
      }
    }
  }
}

function getHeading1Title(wd: string) {
  const testMatch = wd.match(/(^<!-- #\s)(\d+\.)*\s*(?<title>.*)\s-->/i);
  const title = testMatch?.groups?.title ?? "";
  return title;
}

function createFrontMatter(title: string, index: number) {
  const weight = (index + 1) * 10;
  const r = `---
title: "${title}"
chapter: true
weight: ${weight}
---\n`;
  return r;
}

async function splitByExport(marked: string, pathToSave: string = "") {
  // create include markdown
  const markdownSplitted: string[] = [];

  const markdowns = marked
    .replace(/(<!--\s*word\s+export\s+.*-->)/gi, exportSplit + "$1")
    .split(exportSplit);

  console.log(`mds  ${markdowns.length}`);

  for (let index = 0; index < markdowns.length; index++) {
    const element = markdowns[index];
    const list = getWordDownCommand(element);
    if (list?.command === "export") {
      const fileName = path.resolve(pathToSave, list.params[0] ?? "zzzzz.md");
      console.log(`try save: ${fileName}`);
      if (!(await fileExists(fileName))) {
        Fs.writeFileSync(fileName, element);
        markdownSplitted.push(`<!-- word import ${list.params[0]} -->`);
      } else {
        console.log(`${fileName} already exists!!!`);
      }
    } else {
      // add direct
      markdownSplitted.push(element);
    }
  }
  // save splitted
  const splittedFileName = path.resolve(pathToSave, "splittedXXXXX.md");
  if (!(await fileExists(splittedFileName))) {
    Fs.writeFileSync(splittedFileName, markdownSplitted.join("\n"));
  } else {
    console.log(`${splittedFileName} already exists!!!`);
  }
}

// splitMarkdownExportCmd(
//   "C:\\home\\tora-hub\\markdown-to-xxxx\\md_demo\\demo.md"
// );
