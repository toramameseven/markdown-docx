import * as vscode from "vscode";
import { MessageType } from "./markdown-docx/common";

export const extensionName = "markdown-docx";
export const extensionNameShort = "m2x";
export const outputTab = vscode.window.createOutputChannel(extensionName);

export const codeStatusBar = vscode.window.createStatusBarItem(
  vscode.StatusBarAlignment.Right,
  100
);

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
      showInfo(message, source, showNotification);
      break;
    case MessageType.warn:
      showWarn(message, source, showNotification);
      break;
    case MessageType.err:
    default:
      showError(message, source, showNotification);
  }
}

function showInfo(message: unknown, source: string, showNotification = false) {
  const messageOut = `[Info   - ${new Date().toLocaleTimeString()}] ${createMessage(
    message,
    source
  )}`;
  outputTab.appendLine(messageOut.trim());
  if (showNotification) {
    vscode.window.showInformationMessage(messageOut);
  }
}

function showWarn(message: unknown, source: string, showNotification = true) {
  const messageOut = `[Warn   - ${new Date().toLocaleTimeString()}] ${createMessage(
    message,
    source
  )}`;
  outputTab.appendLine(messageOut.trim());
  if (showNotification) {
    vscode.window.showWarningMessage(messageOut);
  }
}

function showError(message: unknown, source: string, showNotification = true) {
  const messageOut = `[Error  - ${new Date().toLocaleTimeString()}] ${createMessage(
    message,
    source
  )}`;
  outputTab.appendLine(messageOut.trim());
  if (showNotification) {
    vscode.window.showErrorMessage(messageOut);
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
