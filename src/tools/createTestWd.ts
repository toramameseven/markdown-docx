import { markdownToWd } from "../../src/markdown-docx/markdown-to-docx";
import path = require("path");
import * as Fs from "fs";

async function createTestWd() {
  const pathDemoMd = "../../md_demo";
  const r = path.resolve(__dirname, pathDemoMd);
  Fs.readdir(r, (err, files) => {
    files
      .filter((f) => f.match(/\.md$/i))
      .forEach(async (file) => {
        const baseName = path.basename(file).replace(/\.md$/i, "");
        const mdPath = path.resolve(__dirname, pathDemoMd, baseName + ".md");
        const r = await markdownToWd(mdPath, "", 0, false, false);
        Fs.copyFileSync(
          r.wdPath,
          path.resolve(__dirname, pathDemoMd, "wd", baseName + ".wd")
        );
        Fs.unlinkSync(r.wdPath);
      });
  });
}

//createTestWd();
