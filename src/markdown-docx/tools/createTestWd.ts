import { markdownToWd } from "../markdown-to-xxxx";
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
        const r = await markdownToWd(mdPath, "", "docx", 0, false);
        Fs.copyFileSync(
          r.wdPath,
          path.resolve(__dirname, pathDemoMd, "wd", baseName + ".wd")
        );
        Fs.unlinkSync(r.wdPath);
      });
  });
}

//createTestWd();
