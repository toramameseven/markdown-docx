import { getFileContents } from "./tools/tools-common";
import { MessageType, getWordDownCommand } from "./common";

import { ShowMessage } from "./common";

const wordCommand = {
  title: "title",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",
} as const;

type TableInfo = {
  emptyMerge: boolean;
  rowMergeList: number[];
  currentRow: number;
  levelOffset: number;
};

export async function addTableSpanToMarkdown(
  mdFilePath: string,
  mdBody: string,
  showMessage: ShowMessage | undefined
) {
  try {
    let body = mdBody;
    if (body === "") {
      body = getFileContents(mdFilePath);
    }
    let lines = body.split(/\r?\n/);

    let tableInfo: TableInfo = {
      emptyMerge: false,
      rowMergeList: [],
      currentRow: -1,
      levelOffset: 0,
    };

    let currentLine = "";
    for (let i = 0; i < lines.length; i++) {
      currentLine = lines[i];
      tableInfo = resolveHtmlCommentEx(currentLine, tableInfo);

      if (
        currentLine[0] === "|" &&
        (tableInfo.emptyMerge || tableInfo.rowMergeList.length)
      ) {
        // update current row
        tableInfo = { ...tableInfo, currentRow: tableInfo.currentRow + 1 };

        // table merge
        let convertedLine = "|";
        const splitted = currentLine.split("|");

        let isMergeEntireRow: boolean = false;
        if (tableInfo.rowMergeList.includes(tableInfo.currentRow)) {
          isMergeEntireRow = true;
        }

        splitted.forEach((x, index) => {
          if (index === 0 || index === splitted.length - 1) {
            // this is not table cell.
            convertedLine += x;
          } else if (x.trim() === "" || isMergeEntireRow) {
            convertedLine += x + " ^|";
          } else {
            convertedLine += x + "|";
          }
        });
        lines[i] = convertedLine;
      } else {
        // outside table
        if (tableInfo.currentRow > -1) {
          tableInfo = {
            ...tableInfo,
            emptyMerge: false,
            rowMergeList: [],
            currentRow: -1,
          };
        }
      }
    }
    return lines.join("\n");
  } catch (error) {
    showMessage?.(MessageType.err, error, "addTableSpanToMarkdown", false);
  }
  return mdBody;
}

function resolveHtmlCommentEx(content: string, tableInfo: TableInfo) {
  // const testMatch = content.match(/<!--(?<name>.*)-->/i);
  // const command = testMatch?.groups?.name ?? "";
  // const params = command.trim().split(" ");

  const r = getWordDownCommand(content);

  // wordDown commands
  if (r) {
    const { command, params } = r;
    //if (params.length > 1 && params[0] === wordCommand.word) {
    switch (command) {
      // merge info rows
      // <!-- word rowMerge 3-4,5-6,7-9 -->
      // rowMerge = 3-4,5-6,7-9
      // ${this.rowCount}-${this.rowCount} is for all rows
      case wordCommand.rowMerge:
        if (params.length) {
          //[1,2], [4,5]
          const mList = params[0].split(",").map((x) => x.split("-"));
          const rowMergeList: number[] = [];
          mList.forEach((x) => {
            let [start, end] = x;
            let startI = parseInt(start);
            let endI = parseInt(end);
            for (let i = startI + 1; i < endI + 1; i++) {
              rowMergeList.push(i);
            }
          });
          return { ...tableInfo, rowMergeList };
        }
        break;
      case wordCommand.emptyMerge:
        return { ...tableInfo, emptyMerge: true };
      // case wordCommand.title:
      //   // default ''
      //   const title = params[0];
      //   return createBlockCommand(wordCommand.title, {
      //     title,
      //   });
      default:
        return tableInfo;
        break;
    }
  }
  return tableInfo;
}
