import * as vscode from "vscode";
import { MessageType } from "./markdown-docx/common";

export const extensionName = "markdown-docx";
export const extensionNameShort = "m2x";
export const outputTab = vscode.window.createOutputChannel(extensionName);

export const codeStatusBar = vscode.window.createStatusBarItem(
  vscode.StatusBarAlignment.Right,
  100
);

/**
 * get vscode current folder
 * @returns 
 */
export function getWorkingDirectory(){
  let wf = vscode.workspace.workspaceFolders?.[0];
  return wf;
}

function createMessage(message: string | unknown, source: string) {
  if (message instanceof Error) {
    return source + " : " + message.message;
  } else if (typeof message === "string") {
    return source + " : " + message;
  }
  return "create message error.";
}

export function showMessage(
  msgType: MessageType,
  message: unknown,
  source: string,
  showNotification?: boolean
) {
  switch (msgType) {
    case MessageType.info:
    case MessageType.warn:
    case MessageType.debug:
    case MessageType.err:
      showMessageCore(msgType, message, source, showNotification);
      break;
    default:
      showMessageCore(MessageType.err, message, source, showNotification);
  }
}

function showMessageCore(msgType: MessageType, message: unknown, source: string, showNotification = false) {
  const messageOut = `[${msgType}   - ${new Date().toLocaleTimeString()}] ${createMessage(
    message,
    source
  )}`;
  outputTab.appendLine(messageOut.trim());
  if (showNotification) {
    switch (msgType) {
      case MessageType.info:
        vscode.window.showInformationMessage(messageOut);
        break;
      case MessageType.warn:
        vscode.window.showWarningMessage(messageOut);
        break;
      case MessageType.debug:
        //
        break;
      case MessageType.err:
        vscode.window.showErrorMessage(messageOut);
        break;
      default:
        vscode.window.showErrorMessage(messageOut);
        break;
    }
  }
}

/*
async function modalDialogShow(message: string, retValue?: boolean) {
  if (retValue !== undefined) {
    return retValue;
  }
  const ans = await vscode.window.showInformationMessage(
    message,
    { modal: true },
    { title: "No", isCloseAffordance: true, dialogValue: false },
    { title: "Yes", isCloseAffordance: false, dialogValue: true }
  );
  return ans?.dialogValue ?? false;
}
*/

export function updateStatusBar(isRunning: boolean): void {
  vscode.commands.executeCommand(
    "setContext",
    "markdown-docx.isRunning",
    isRunning
  );
  if (isRunning) {
    codeStatusBar.text = `$(sync~spin) ${extensionNameShort}`;
    codeStatusBar.show();
    return;
  }
  codeStatusBar.hide();
}
