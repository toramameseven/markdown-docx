// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from "vscode";
import * as vscodeCommon from "./common-vscode";
import { splitMarkdownToHugo } from "./tools/split-markdown";
import { htmlToMarkdown } from "./html-markdown/html-to-markdown";
import {
  markdownToDocx,
  wordDownToDocx,
  DocxOption,
  MessageType,
} from "./markdown-docx";
import {
  markdownToExcel,
  markdownToHtml,
  markdownToPptx,
  markdownToTextile,
} from "./markdown-docx/markdown-to-xxxx";
// import { markdownToPptx } from "./markdown-pptx/markdown-to-pptx";
import { createDocxTemplateFile } from "./markdown-docx/common";
import { wordDownToPptx } from "./markdown-docx/wd-to-pptx";
import { getWorkingDirectory } from "./common-vscode";

export let isDebug = false;

// for cancel spawn
let ac = new AbortController();

function resetAbortController() {
  ac = new AbortController();
}

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {
  // activate event
  vscodeCommon.showMessage(
    MessageType.info,
    '"markdown-docx" is now active!',
    "main"
  );

  // explorer md wd to docx
  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.ExportDocx",
      exportDocxFromExplorer
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.mdToPptx", exportPptxFromExplorer)
  );

  // explorer html to docx
  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.ExportHtmlDocx", exportHtmlDocx)
  );

  //exportMarkdownEd
  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.mdToEd", exportMarkdownEd)
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.mdToTextile",
      exportMarkdownTextile
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.mdToHtml", exportMarkdownHtml)
  );

  // split md for hugo
  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.SplitForHugo",
      splitToHugoMarkdown
    )
  );

  //  main.createDocxTemplate
  context.subscriptions.push(
    vscode.commands.registerCommand(
      "main.createDocxTemplate",
      createDocxTemplate
    )
  );

  // editor md wd to docx
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand(
      "editor.ExportDocx",
      exportDocxFromEditor
    )
  );

  // editor html to docx
  // context.subscriptions.push(
  //   vscode.commands.registerTextEditorCommand(
  //     "editor.ExportHtmlDocx",
  //     exportDocxFromEditor
  //   )
  // );

  // cancel convert
  context.subscriptions.push(
    vscode.commands.registerCommand("editor.ExportStop", () => {
      ac.abort();
    })
  );
}

/**
 * convert html to markdown
 * @param uriFile
 */
function exportHtmlDocx(uriFile: vscode.Uri) {
  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.html$/i)) {
      // wordDown
      const r = htmlToMarkdown(filePath);
      vscodeCommon.showMessage(MessageType.info, r, "");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportMarkdownEd(uriFile: vscode.Uri) {
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      // wordDown
      //const r = markdownToExDown(filePath, "");
      const r = markdownToExcel(filePath, "", 0, thisOption);
      vscodeCommon.showMessage(MessageType.info, r, "");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportMarkdownTextile(uriFile: vscode.Uri) {
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      // wordDown
      const r = markdownToTextile(filePath, "", 0, thisOption);
      vscodeCommon.showMessage(MessageType.info, r, "");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportMarkdownHtml(uriFile: vscode.Uri) {
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      // wordDown
      //const r = markdownToExDown(filePath, "");
      const r = markdownToHtml(filePath, "", 0, thisOption);
      vscodeCommon.showMessage(MessageType.info, r, "");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

/**
 * convert a markdown to hugo markdowns
 * @param uriFile
 */
async function splitToHugoMarkdown(uriFile: vscode.Uri) {
  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$|\.mds$/i)) {
      // wordDown
      await splitMarkdownToHugo(filePath);
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function createDocxTemplate() {
  try {
    const wf = getWorkingDirectory();
    await createDocxTemplateFile(wf?.uri.fsPath ?? "");
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

/**
 * convert a md or a wd to a docx on a Explorer.
 * @param uriFile
 */
async function exportDocxFromExplorer(uriFile: vscode.Uri) {
  try {
    vscodeCommon.updateStatusBar(true);
    await exportDocxFromExplorerCore(uriFile);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportPptxFromExplorer(uriFile: vscode.Uri) {
  try {
    vscodeCommon.updateStatusBar(true);
    await exportPptxFromExplorerCore(uriFile);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

/**
 * convert file(md, wd) to docx
 * @param uriFile
 * @returns
 */
async function exportDocxFromExplorerCore(uriFile: vscode.Uri) {
  const filePath = uriFile.fsPath;

  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "main"
  );

  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wordDownToDocx(filePath, thisOption);
    return;
  }

  // markdown
  try {
    await markdownToDocx(filePath, "", 0, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "");
  }
}

async function exportPptxFromExplorerCore(uriFile: vscode.Uri) {
  const filePath = uriFile.fsPath;

  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "main"
  );

  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wordDownToPptx(filePath, thisOption);
    return;
  }

  // markdown
  try {
    await markdownToPptx(filePath, "", 0, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "");
  }
}

/**
 * convert a md or a wd to a docx on a editor.
 * @param textEditor
 * @param edit
 * @returns
 */
async function exportDocxFromEditor(
  textEditor: {
    document: { uri: { fsPath: any }; getText: (arg0: any) => any };
    selection: any;
  },
  edit: any
) {
  try {
    vscodeCommon.updateStatusBar(true);
    await exportDocxFromEditorCore(textEditor, edit);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "main");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

/**
 * convert text(md, wd) to docx
 * @param textEditor
 * @param edit
 * @returns
 */
async function exportDocxFromEditorCore(
  textEditor: {
    document: { uri: { fsPath: any }; getText: (arg0: any) => any };
    selection: any;
  },
  // eslint-disable-next-line no-unused-vars
  edit: any
) {
  vscodeCommon.outputTab.show();

  const filePath = textEditor.document.uri.fsPath;
  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "main"
  );
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wordDownToDocx(filePath, thisOption);
    return;
  }

  let selection = textEditor.selection;
  const isSelected = selection.start !== selection.end;
  const startLine = isSelected ? selection.start.line : 0;
  let text = isSelected ? textEditor.document.getText(selection) : "";
  try {
    await markdownToDocx(filePath, text, startLine, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "");
  }
}

// This method is called when your extension is deactivated
export function deactivate() {}

// options
function createDocxOptionExtension(option: DocxOption) {
  const r: DocxOption = {
    timeOut: getTimeout(),
    docxEngine: getDocxEngine(),
    docxTemplate: getDocxTemplate(),
    mathExtension: getMathExtension(),
    isDebug: getDebug(),
    logInterval: getLogInterval(),
    isUseDocxJs: getUseDocxJs(),
    isOverWrite: getIsOverWrite(),
    wordPath: getWordPath(),
    isWordOpen: getIsWordOpen(),
  };
  return { ...r, ...option };

  // get docx docxTemplate
  function getDocxTemplate() {
    const docxTemplate =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<string>("path.docxTemplate") ?? "";
    return docxTemplate;
  }

  // get docx vbs
  function getDocxEngine() {
    const docxEngine =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<string>("path.docxEngine") ?? "";
    return docxEngine;
  }

  // get docx convert timeout milliseconds.
  function getTimeout() {
    const timeout =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<number>("docxEngine.timeout") ?? 60000;
    return timeout;
  }

  // get debug mode
  function getDebug() {
    const isDebug =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.debug") ?? false;
    return isDebug;
  }

  // is math extension is enable.
  function getMathExtension() {
    const mathExtension =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.mathExtension") ?? false;
    return mathExtension;
  }

  // get docx convert timeout milliseconds.
  function getLogInterval() {
    const timeout =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<number>("docxEngine.logInterval") ?? 10;
    return timeout > 0 ? timeout : 10;
  }

  function getUseDocxJs() {
    const useDocxJs =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.docxJs") ?? true;
    return useDocxJs;
  }

  function getIsOverWrite() {
    const isOverWrite =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.isOverWrite") ?? true;
    return isOverWrite;
  }

  function getIsWordOpen() {
    const isWordOpen =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.isWordOpen") ?? true;
    return isWordOpen;
  }

  function getWordPath() {
    const wordPath =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<string>("docxEngine.wordExePath") ?? "";
    return wordPath;
  }
}
