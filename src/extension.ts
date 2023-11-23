// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from "vscode";
import * as vscodeCommon from "./common-vscode";
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
import { wdToPptx } from "./markdown-docx/wd-to-pptxJs";
import { createDocxTemplateFile } from "./markdown-docx/common";
import { getWorkingDirectory } from "./common-vscode";
import { textileToHtml } from "./markdown-docx/tools/toolsTextile";
import { isEnableExperimentalFeature } from "./common-settings";
export let isDebug = false;

//export const isEnableExperimentalFeature: boolean = true;

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
    '"markdown-docx" is initializing.',
    "extension"
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
    vscode.commands.registerCommand(
      "explorer.ExportHtmlMarkdown",
      exportMarkdownFromHtml
    )
  );

  //exportMarkdownEd
  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.mdToEd", exportEdFromMarkdown)
  );
  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.mdToTextile",
      exportTextileFromMarkdown
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand("explorer.mdToHtml", exportHtmlFromMarkdown)
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.htmlToInlineHtml",
      exportInlineHtmlFromMarkdown
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.htmlToInlineHtmlNoMenu",
      exportInlineHtmlFromMarkdownNoMenu
    )
  );

  context.subscriptions.push(
    vscode.commands.registerCommand(
      "explorer.textileToMarkdown",
      exportMarkdownFromTextile
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

  // editor md wd to pptx
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand(
      "editor.ExportPptx",
      exportPptxFromEditor
    )
  );

  // editor html to docx
  context.subscriptions.push(
    vscode.commands.registerTextEditorCommand(
      "editor.ExportHtmlDocx",
      exportDocxFromEditor
    )
  );

  // cancel convert
  context.subscriptions.push(
    vscode.commands.registerCommand("editor.ExportStop", () => {
      ac.abort();
    })
  );

  // this feature is only experiment!!
  if (isEnableExperimentalFeature) {
    enableExperimentFeature();
  }

  enableMainFeature();


  vscodeCommon.showMessage(
    MessageType.info,
    '"markdown-docx" is initialized.',
    "extension"
  );
}

/**
 * convert html to markdown
 * @param uriFile
 */
function exportMarkdownFromHtml(uriFile: vscode.Uri) {
  try {

    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.html$|\.htm$/i)) {
      // wordDown
      const r = htmlToMarkdown(filePath, "");
      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportEdFromMarkdown(uriFile: vscode.Uri) {
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      // wordDown
      //const r = markdownToExDown(filePath, "");
      const r = markdownToExcel(filePath, "", 0, thisOption);
      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportTextileFromMarkdown(uriFile: vscode.Uri) {
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
      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

function exportHtmlFromMarkdown(uriFile: vscode.Uri) {
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
      const r = markdownToHtml(filePath, "", 0, thisOption, false, false);
      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportInlineHtmlFromMarkdown(uriFile: vscode.Uri) {
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      const r = await markdownToHtml(filePath, "", 0, thisOption, true, true);

      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportInlineHtmlFromMarkdownNoMenu(uriFile: vscode.Uri) {
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.md$/i)) {
      const r = await markdownToHtml(filePath, "", 0, thisOption, true, false);

      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportMarkdownFromTextile(uriFile: vscode.Uri) {
  // todo implement.
  // const thisOption = createDocxOptionExtension({
  //   ac,
  //   message: vscodeCommon.showMessage,
  // });

  try {
    vscodeCommon.updateStatusBar(true);
    const filePath = uriFile.fsPath;
    if (filePath.match(/\.textile$/i)) {
      const rHtml = textileToHtml(filePath, "");

      const r = htmlToMarkdown(filePath, rHtml);

      vscodeCommon.showMessage(MessageType.info, r, "extension");
    }
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
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
    vscodeCommon.showMessage(MessageType.err, error, "extension");
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

  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "extension"
  );

  // wordDown
  if (filePath.match(/\.wd$/i)) {
    await wordDownToDocx(filePath, "", thisOption);
    return;
  }

  // markdown
  try {
    await markdownToDocx(filePath, "", 0, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  }
}

async function exportPptxFromExplorer(uriFile: vscode.Uri) {
  try {
    vscodeCommon.updateStatusBar(true);
    await exportPptxFromExplorerCore(uriFile);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportPptxFromExplorerCore(uriFile: vscode.Uri) {
  const filePath = uriFile.fsPath;

  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "extension"
  );

  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wdToPptx(filePath, "", thisOption);
    return;
  }

  // markdown
  try {
    await markdownToPptx(filePath, "", 0, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
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
    vscodeCommon.showMessage(MessageType.err, error, "extension");
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
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  const filePath = textEditor.document.uri.fsPath;
  vscodeCommon.showMessage(
    MessageType.info,
    `convert docx from ${filePath}`,
    "extension"
  );

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wordDownToDocx(filePath, "", thisOption);
    return;
  }

  let selection = textEditor.selection;
  const isSelected = selection.start !== selection.end;
  const startLine = isSelected ? selection.start.line : 0;
  let text = isSelected ? textEditor.document.getText(selection) : "";
  try {
    await markdownToDocx(filePath, text, startLine, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  }
}

async function exportPptxFromEditor(
  textEditor: {
    document: { uri: { fsPath: any }; getText: (arg0: any) => any };
    selection: any;
  },
  edit: any
) {
  try {
    vscodeCommon.updateStatusBar(true);
    await exportPptxFromEditorCore(textEditor, edit);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

async function exportPptxFromEditorCore(
  textEditor: {
    document: { uri: { fsPath: any }; getText: (arg0: any) => any };
    selection: any;
  },
  // eslint-disable-next-line no-unused-vars
  edit: any
) {
  resetAbortController();
  const thisOption = createDocxOptionExtension({
    ac,
    message: vscodeCommon.showMessage,
  });

  if (thisOption.isShowOutputTab) {
    vscodeCommon.outputTab.show();
  }

  const filePath = textEditor.document.uri.fsPath;
  vscodeCommon.showMessage(
    MessageType.info,
    `convert pptx from ${filePath}`,
    "extension"
  );

  if (filePath.match(/\.wd$/i)) {
    // wordDown
    await wdToPptx(filePath, "", thisOption);
    return;
  }

  let selection = textEditor.selection;
  const isSelected = selection.start !== selection.end;
  const startLine = isSelected ? selection.start.line : 0;
  let text = isSelected ? textEditor.document.getText(selection) : "";
  try {
    await markdownToPptx(filePath, text, startLine, thisOption);
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  }
}

// This method is called when your extension is deactivated
export function deactivate() { }

function enableExperimentFeature() {
  vscode.commands.executeCommand(
    "setContext",
    "markdown-docx.isExperimentFeature",
    true
  );
}

function enableMainFeature() {
  vscode.commands.executeCommand(
    "setContext",
    "markdown-docx.isMainFeature",
    true
  );
}

async function createDocxTemplate() {
  try {
    const wf = getWorkingDirectory();
    await createDocxTemplateFile(wf?.uri.fsPath ?? "extension");
  } catch (error) {
    vscodeCommon.showMessage(MessageType.err, error, "extension");
  } finally {
    vscodeCommon.updateStatusBar(false);
  }
}

// options
function createDocxOptionExtension(option: DocxOption) {
  const r: DocxOption = {
    timeOut: getTimeout(),
    docxEngine: getDocxEngine(),
    docxTemplate: getDocxTemplate(),
    mathExtension: getMathExtension(),
    isShowOutputTab: getIsShowOutputTab(),
    isDebug: getDebug(),
    logInterval: getLogInterval(),
    isUseDocxJs: getUseDocxJs(),
    isOverWrite: getIsOverWrite(),
    wordPath: getWordPath(),
    isOpenWord: getIsWordOpen(),
    isOpenPpt: getIsPptOpen(),
    pptPath: getPptPath(),
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

  function getIsShowOutputTab() {
    const isShowOutputTab =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.showOutputTab") ?? false;
    return isShowOutputTab;
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
    // const useDocxJs =
    //   vscode.workspace
    //     .getConfiguration("markdown-docx")
    //     .get<boolean>("docxEngine.docxJs") ?? true;
    // return useDocxJs;
    return true;
  }

  function getIsOverWrite() {
    const isOverWrite =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.isOverWrite") ?? true;
    return isOverWrite;
  }

  function getIsWordOpen() {
    const isOpenWord =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.isOpenWord") ?? true;
    return isOpenWord;
  }

  function getIsPptOpen() {
    const isPptWord =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<boolean>("docxEngine.isPptWord") ?? true;
    return isPptWord;
  }

  function getWordPath() {
    const wordPath =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<string>("docxEngine.wordExePath") ?? "";
    return wordPath;
  }
  function getPptPath() {
    const pptPath =
      vscode.workspace
        .getConfiguration("markdown-docx")
        .get<string>("docxEngine.pptExePath") ?? "";
    return pptPath;
  }
}
