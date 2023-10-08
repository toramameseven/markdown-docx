import type { marked as Marked } from "marked";
import { marked } from "marked";
import { unescape } from "lodash";
import {
  getWordDownCommand as parseHtmlComment,
  MessageType,
  ShowMessage,
} from "./common";
import { addTableSpanToMarkdown } from "./add-table-span";
import { spanTable } from "./marked-extended-tables";
import { Bookmarks } from "./bookmarks";

const source = "markdown-to-wd";
let showMessage: ShowMessage | undefined;

const bookmarks = new Bookmarks();

const markedCommand = {
  heading: "heading",
  paragraph: "paragraph",
  list: "list",
  listitem: "listitem",
  code: "code",
  blockquote: "blockquote",
  table: "table",
  tablerow: "tablerow",
  tablecell: "tablecell",
  html: "html",
  text: "text",
  image: "image",
  link: "link",
  codespan: "codespan",
  strong: "strong",
  em: "em",
  del: "~~",
  non: "non",
  hr: "hr",
} as const;
// type MarkedCommand = typeof markedCommand[keyof typeof markedCommand];

const wordCommand = {
  cols: "cols",
  rowMerge: "rowMerge",
  emptyMerge: "emptyMerge",
  newPage: "newPage",
  newLine: "newLine",
  toc: "toc",
  export: "export",
  placeholder: "placeholder",
  param: "param",
  } as const;

const ooxParameters = {
  pptxSettings: "pptxSettings",
  position: "position",
  dpi: "dpi",
  docxTemplate: "docxTemplate",
  refFormat: "refFormat",
} as const;

export type OoxParameters = (typeof ooxParameters)[keyof typeof ooxParameters];

export const wd0Command = {
  ...markedCommand,
  ...wordCommand,
} as const;
export type Wd0Command = (typeof wd0Command)[keyof typeof wd0Command];

const _sp = "\t";
const _newline = "\n";

// inline
const inline = (inlineType: string) => (content: string) => {
  if (inlineType) {
    return `<${inlineType}>${content}</${inlineType}>`;
  }
  return content;
};

// inline for *emphasis*
const inlineEx = (inlineType: string) => (content: string) => {
  if (inlineType) {
    return `${inlineType}${content}${inlineType}`;
  }
  return content;
};

// hr
const hr = () => "\nhr\n";

// newline
const newline = () => _newline;

//empty
const empty = () => "";

// block
const block = (blokType: string) => (content: string) => {
  const params = {
    text: content,
  };
  return createBlockCommand(blokType, params);
};

// block quote
/* does note use
const blockQuote = (blokType: string) => (content: string) => {
  const params = {
    text: content,
  };
  return createBlockCommand(blokType, params);
};
*/

//
const htmlBlock = (content: string) => {
  // word down command is a html comment. <!--
  const wd0 = resolveHtmlComment(content);
  if (wd0) {
    return wd0;
  }

  if (content.match(/<br>/i)) {
    return createBlockCommand(wd0Command.newLine, {
      info: "info new line",
    });
  }

  const usedHtml = /<br>|<sup>|<\/sup>|<sub>|<\/sub>|<!--.*/i;
  const nonUsedHtml = /<!--\s+(word|ppt|oox).*/i;

  if (content.match(usedHtml) && !content.match(nonUsedHtml)) {
    // output inlineF
    return content;
  }

  showMessage?.(
    MessageType.warn,
    `Next html is not allowed: ${content}`,
    source,
    false
  );
  return "";
};

const documentInfoParams = [
  "pptxSettings",
  "position",
  "dpi",
  "docxTemplate",
  "refFormat",
  "tableWidth",
  "levelOffset",
  "tableCaption",
  "tableCaptionId",
] as const;
type DocumentInfoParams = (typeof documentInfoParams)[number];
const isDocumentInfoParams = (name: string): name is DocumentInfoParams => {
  return documentInfoParams.some((value) => value === name);
};

function resolveHtmlComment(content: string) {
  const commandParams = parseHtmlComment(content);
  if (commandParams) {
    // <!-- word command param1 param2 param3 -->
    const { command, params } = commandParams;
    switch (command) {
      case wordCommand.cols:
        if (params.length) {
          return createBlockCommand(wd0Command.cols, { cols: params[0] });
        }
        break;
      case wordCommand.rowMerge:
        if (params.length) {
          return createBlockCommand(wd0Command.rowMerge, {
            rowMerge: params[0],
          });
        }
        break;
      case wordCommand.emptyMerge:
        return createBlockCommand(wd0Command.emptyMerge, { emptyMerge: "1" });

      case wordCommand.newPage:
        return createBlockCommand(wd0Command.newPage, {
          info: "info new page",
        });
      case wordCommand.newLine:
        return createBlockCommand(wd0Command.newLine, {
          info: "info new line",
        });
      case wordCommand.toc:
        // default toc is to 3 heading level
        const tocTo = params[0] ? params[0] : "3";
        const tocCaption = params[1] ?? "table of contents";
        return createBlockCommand(wd0Command.toc, {
          tocTo,
          tocCaption,
        });
      case wordCommand.export:
        return createBlockCommand(wordCommand.export, {
          info: "export",
        });
      case wordCommand.param:
      case wordCommand.placeholder:
        let r = "";
        for (let i = 1; i < params.length; i += 2) {
          if (params[i - 1]) {
            r += createBlockCommand(command, {
              key: params[i - 1],
              value: params[i],
            });
          }
        }
        return r;
      default:
        // todo error parameters
        let defaultParam = "";
        let defaultParams = [command, ...params];
        for (let i = 1; i < defaultParams.length; i += 2) {
          if (
            defaultParams[i - 1] &&
            isDocumentInfoParams(defaultParams[i - 1])
          ) {
            defaultParam += createBlockCommand("param", {
              key: defaultParams[i - 1],
              value: defaultParams[i],
            });
          } else {
            showMessage?.(
              MessageType.warn,
              `Next is not a parameter: ${defaultParams[i - 1]}`,
              "markdown-to-wd0::resolveHtmlComment",
              false
            );
          }
        }
        return defaultParam;
    }
  }
}


export function getWordTitle(wd: string) {
  const r = parseHtmlComment(wd);
  const title = r?.command === "title" ? r.params[0] : "no title";
  return title;
}



// -----------------------------------

const blockHeading = (content: string, index: number) => {
  //title for asciidoc type
  const isAdocTypeTitle = false;
  if (index === 1 && isAdocTypeTitle) {
    //TODO // title
  }

  // heading offset
  const headingOffset = 0;

  // headings
  const reg = /([^{}]+)(\s{#(\w+)})?/gm;
  const r = reg.exec(content);

  const mainTitle = (r && r[1]) ?? content;
  const mainTitleSplit = mainTitle.split(" ");
  let title = mainTitle;
  if (
    mainTitleSplit &&
    mainTitleSplit.length > 1 &&
    !Number.isNaN(parseInt(mainTitleSplit[0][0], 10))
  ) {
    title = title.slice(mainTitleSplit[0].length);
    showMessage?.(MessageType.warn, `${mainTitle} ==> ${title}`, source, false);
  }

  let idTitle = (r && r[4]) ?? title;
  idTitle = bookmarks.slugify(idTitle);

  const lines = splitBlockContents(content);
  const headings = {
    index: (index - headingOffset).toString(),
    //title: title,
    idTitle,
    text: lines.join(_newline),
  };
  return createBlockCommand(wd0Command.heading, headings);
};

function splitBlockContents(blockContents: string) {
  const lines = blockContents
    .split(_newline)
    .filter(Boolean)
    .map((s) => {
      //if (s.match(/\t|^\/.*/)) {
      if (s.match(/\t/)) {
        // with command
        return _newline + s + _newline;
      } else {
        // only text
        return `\n${wd0Command.text}\t${wd0Command.text}\t` + s + _newline;
      }
    });
  return lines;
}

// a code block needs empty lines
function splitCodeBlockContents(
  blockContents: string
) {
  const lines = blockContents.split(_newline).map((s) => {
    //if (s.match(/\t|^\/.*/)) {
    if (s.match(/\t/)) {
      // with command
      return _newline + s + _newline;
    } else {
      // only text
      return `\n${wd0Command.text}\t${wd0Command.text}\t${s}${_newline}`;
    }
  });
  return lines;
}

const blockCodeParagraph =
  (blokType: string) => (content: string, language: string | undefined) => {
    const lines = splitCodeBlockContents(content);
    const params = {
      language: language ?? "",
      body: lines.join(_newline),
    };
    return createBlockCommand(blokType, params);
  };

const blockParagraph =
  (blokType: string, isNeedEmptyLine = false) =>
  (content: string) => {
    const lines = isNeedEmptyLine
      ? splitCodeBlockContents(content)
      : splitBlockContents(content);
    const params = {
      body: lines.join(_newline),
    };
    return createBlockCommand(blokType, params);
  };

const blockList = (body: string, ordered: boolean, start: number) => {
  const params = {
    ordered: ordered ? "1" : "0",
    start: start.toString(),
    body: body,
  };
  return createBlockCommand(wd0Command.list, params);
};

const blockListItem = (content: string, task: boolean, checked: boolean) => {
  const lines = splitBlockContents(content);
  const params = {
    task: task ? "1" : "0",
    checked: checked ? "1" : "0",
    text: lines.join(_newline),
  };
  return createBlockCommand(wd0Command.listitem, params);
};

// [text] (href "title")
// command href, text, title
const blockImage =
  (blokType: string) =>
  (href: string | null, title: string | null, content: string) => {
    const params = {
      href: href ?? "",
      text: content,
      title: title ?? "",
    };
    return createBlockCommand(blokType, params);
  };

const blockTable = (header: string, body: string) => {
  const params = {
    header,
    body,
  };
  return createBlockCommandTable(wd0Command.table, params);
};

const blockTableCell = (
  content: string,
  flags: {
    header: boolean;
    align: "center" | "left" | "right" | null;
  }
) => {
  const lines = splitBlockContents(content);
  const params = {
    isHeder: flags.header ? "1" : "0",
    align: flags.align?.toString() ?? "left",
    content: lines.join(_newline),
  };
  return createBlockCommand(wd0Command.tablecell, params);
};

const blockLink = (
  href: string | null,
  title: string | null,
  content: string
) => {
  let convertRef = href ?? "";

  // if convertRef === "", what happens?  for index? what is index?
  if (convertRef.length === 1 && convertRef[0] === "#") {
    convertRef = "";
  }

  // [content](#href "title") .. word cross ref
  // when content === "", cross ref
  if (convertRef.length > 1 && convertRef[0] === "#") {
    convertRef = bookmarks.slugify(convertRef, true);
    content = "";
  }

  const params = {
    href: convertRef,
    title: title ?? "",
    text: content,
  };
  return createBlockCommand(wd0Command.link, params);
};

//type DocCommand = { command: string; params: DocxParam; isBlock: boolean };
type DocxParam = { [x: string]: string };

const createBlockCommand = (command: string, params: DocxParam) => {
  const xx = joinObjectToString(params);
  const r = `\n${command}${_sp}` + xx;
  // end command /command must have \t. this indicates the line is command. no \t line is a text.
  return `\n${r}\n/${command}\t\t\n`;
};

const createBlockCommandTable = (command: string, params: DocxParam) => {
  const xx = [params.header, params.body].join(_newline);
  // end command /command must have \t. this indicates the line is command. no \t line is a text.
  const r = `\n${command}${_sp}${_sp}\n` + xx;
  // end command /command must have \t. this indicates the line is command. no \t line is a text.
  return `\n${r}\n/${command}${_sp}${_sp}\n`;
};

// {kye: value} > kye\tvalue.....
function joinObjectToString(params: DocxParam) {
  const r: string[] = [];
  for (const [key, value] of Object.entries(params)) {
    r.push(key);
    r.push(value);
  }
  return r.join(_sp);
}

export async function markdownToWd0(
  markdown: string,
  convertType: "docx" | "excel" | "html" | "textile",
  options?: Marked.MarkedOptions,
  messageFunction?: ShowMessage
) {
  bookmarks.clear();
  showMessage = messageFunction;

  const typeOfConvert = {
    docx: docxRenderer,
    excel: excelRenderer,
    html: null,
    textile: textileRenderer,
  } as const;

  const render = typeOfConvert[convertType];
  let markedOptions = { ...options };
  if (render) {
    markedOptions = { ...markedOptions, renderer: render };
  }

  // get markdown levelOffset
  const offsetMatch = markdown.match(
    /<!--\s+(oox|word|ppt)\s+levelOffset\s+(?<name>.*)\s+-->/i
  );
  let levelOffset = parseInt(offsetMatch?.groups?.name ?? "0");
  levelOffset = Number.isNaN(levelOffset) ? 0 : levelOffset;

  const walkTokens = (token: any) => {
    if (token.type === "heading") {
      token.depth += levelOffset;
    }
  };

  // https://marked.js.org/
  let mdForMarked = markdown;
  if (convertType === "html") {
    marked.use(marked.getDefaults());
    marked.use(spanTable());
    marked.use({ walkTokens });
    mdForMarked = await addTableSpanToMarkdown("", mdForMarked, showMessage);
  } else {
    marked.use(marked.getDefaults());
    marked.use({ walkTokens });
  }

  const unmarked = marked(mdForMarked, markedOptions);
  const toOutput = convertType === "html" ? unmarked : unescape(unmarked);
  return toOutput.trim();
}

/**
 * docx renderer
 */
const docxRenderer: Marked.Renderer = {
  // Block elements
  heading: blockHeading,

  // normal paragraph
  paragraph: blockParagraph(markedCommand.paragraph),

  list: blockList,
  listitem: blockListItem,

  // ``` or tab
  code: blockCodeParagraph(markedCommand.code),

  // >
  blockquote: block(markedCommand.blockquote),

  table: blockTable,
  tablerow: block(markedCommand.tablerow),
  tablecell: blockTableCell,

  html: htmlBlock,
  hr: hr,
  checkbox: empty,

  // Inline elements
  image: blockImage(markedCommand.image),
  link: blockLink,
  text: inline(""),
  // `code`
  codespan: inline(markedCommand.codespan),
  // ** **
  strong: inline("b"),
  // _ _
  em: inline("i"),
  // <br>?
  br: newline,
  // ~~ ~~
  del: inline(markedCommand.del),
  // etc.
  options: {},
};

/**
 * excel Renderer
 */
const excelRenderer: Marked.Renderer = {
  // Block elements
  heading: blockHeading,

  // normal paragraph
  paragraph: blockParagraph(markedCommand.paragraph),

  list: blockList,
  listitem: blockListItem,

  // ``` or tab
  code: blockParagraph(markedCommand.code, true),

  // >
  blockquote: block(markedCommand.blockquote),

  table: blockTable,
  tablerow: block(markedCommand.tablerow),
  tablecell: blockTableCell,

  html: htmlBlock,
  hr: hr,
  checkbox: empty,

  // Inline elements
  image: blockImage(markedCommand.image),
  link: blockLink,
  text: inlineEx(""),
  // `code`
  codespan: inlineEx("`"),
  // ** **
  strong: inlineEx("**"),
  // _ _
  em: inlineEx("_"),
  // <br>?
  br: newline,
  // ~~ ~~
  del: inlineEx(markedCommand.del),
  // etc.
  options: {},
};

/**
 * textile Renderer
 */
const textileRenderer: Marked.Renderer = {
  // Block elements
  heading: blockHeading,

  // normal paragraph
  paragraph: blockParagraph(markedCommand.paragraph),

  list: blockList,
  listitem: blockListItem,

  // ``` or tab
  code: blockParagraph(markedCommand.code, true),

  // >
  blockquote: block(markedCommand.blockquote),

  table: blockTable,
  tablerow: block(markedCommand.tablerow),
  tablecell: blockTableCell,

  html: htmlBlock,
  hr: hr,
  checkbox: empty,

  // Inline elements
  image: blockImage(markedCommand.image),
  link: blockLink,
  text: inlineEx(""),
  // `code`
  codespan: inlineEx("`"),
  // ** **
  strong: inlineEx("**"),
  // _ _
  em: inlineEx("_"),
  // <br>?
  br: newline,
  // ~~ ~~
  del: inlineEx(markedCommand.del),
  // etc.
  options: {},
};
