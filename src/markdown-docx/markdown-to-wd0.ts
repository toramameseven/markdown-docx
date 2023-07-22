import { marked } from "marked";
import { unescape } from "lodash";
import { getWordDownCommand, MessageType, ShowMessage } from "./common";

const source = "markdown-to-wd";
let showMessage: ShowMessage | undefined;

const idMap = new Map();

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
  text: "text",
  image: "image",
  link: "link",
  codespan: "codespan",
  strong: "strong",
  em: "em",
  del: "~~",
  non: "non",
} as const;

//type MarkedCommand = typeof markedCommand[keyof typeof markedCommand];

const wordCommand = {
  word: "word",
  title: "title",
  subTitle: "subTitle",
  date: "date",
  docNumber: "docNumber",
  author: "author",
  division: "division",
  cols: "cols",
  rowMerge: "rowMerge",
  tablePos:"tablePos",
  emptyMerge: "emptyMerge",
  newPage: "newPage",
  newLine: "newLine",
  pageSetup: "pageSetup",
  wdOrientationLandscape: "wdOrientationLandscape",
  wdOrientationPortrait: "wdOrientationPortrait",
  wdSizeA4: "wdSizeA4",
  wdSizeA3: "wdSizeA3",
  toc: "toc",
  export: "export",
  docxEngine: "docxEngine",
  docxTemplate: "docxTemplate",
  property: "property",
  clearContent: "clearContent",
  crossRef: "crossRef",
  param: "param"
} as const;

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
  const wd0 = getWordCommand(content);
  if (wd0) {
    return wd0;
  }

  if (content.match(/<br>/i)) {
    return createBlockCommand(wordCommand.newLine, {
      info: "info new line",
    });
  }

  const usedHtml = /<br>|<sup>|<\/sup>|<sub>|<\/sub>|<!--.*/i;

  // return createBlockCommand(blokType, params);

  if (content.match(usedHtml)) {
    // output inline
    return content;
  }
  showMessage?.(
    MessageType.warn,
    `next html is not allow: ${content}`,
    source,
    false
  );
  return `[[ Next html is not allow: ${content} ]]`;
};

function getWordCommand(content: string) {
  // const testMatch = content.match(/<!--(?<name>.*)-->/i);
  // const command = testMatch?.groups?.name ?? "";
  // const params = command.trim().split(" ");

  const r = getWordDownCommand(content);

  // wordDown commands
  if (r) {
    const { command, params } = r;
    //if (params.length > 1 && params[0] === wordCommand.word) {
    switch (command) {
      case wordCommand.cols:
        if (params.length) {
          return createBlockCommand(wordCommand.cols, { cols: params[0] });
        }
        break;
      case wordCommand.rowMerge:
        if (params.length) {
          return createBlockCommand(wordCommand.rowMerge, {
            rowMerge: params[0],
          });
        }
        break;
      case wordCommand.tablePos:
        if (params.length) {
          return createBlockCommand(wordCommand.tablePos, {
            tablePos: params[0],
          });
        }
        break;
      case wordCommand.emptyMerge:
        return createBlockCommand(wordCommand.emptyMerge, { emptyMerge: "1" });

      case wordCommand.newPage:
        return createBlockCommand(wordCommand.newPage, {
          info: "info new page",
        });
      case wordCommand.newLine:
        return createBlockCommand(wordCommand.newLine, {
          info: "info new line",
        });
      case wordCommand.pageSetup:
        //if (params[0] === wordCommand.wdOrientationLandscape || params[0] === wordCommand.wdOrientationPortrait) {
        return createBlockCommand(wordCommand.pageSetup, {
          orientation: params[0],
          pagesize: params[1],
        });
        //}
        break;
      case wordCommand.toc:
        // default toc is to 3 heading level
        const tocTo = params[0] ? params[0] : "3";
        const tocCaption = params[1] ?? "table of contents";
        return createBlockCommand(wordCommand.toc, {
          tocTo,
          tocCaption,
        });
      case wordCommand.property:
        const propertyKey = params[0];
        const propertyValue = params[1];
        if (!propertyKey) {
          return;
        }
        return createBlockCommand(wordCommand.property, {
          propertyKey,
          propertyValue,
        });
      case wordCommand.crossRef:
        const crossRef = params[0];
        return createBlockCommand(wordCommand.crossRef, {
          crossRef,
        });
      case wordCommand.param:
        return createBlockCommand(wordCommand.param, {
          key: params[0],
          value: params[1],
        });
      case wordCommand.clearContent:
        const isClearContent = params[0] ?? false;
        return createBlockCommand(command, {
          isClearContent,
        });
      case wordCommand.title:
        // default ''
        const title = params[0];
        return createBlockCommand(wordCommand.title, {
          title,
        });
      case wordCommand.subTitle:
        // default ''
        const subTitle = params[0];
        return createBlockCommand(wordCommand.subTitle, {
          subTitle,
        });
      case wordCommand.author:
        // default ''
        const author = params[0];
        return createBlockCommand(wordCommand.author, {
          author,
        });
      case wordCommand.docNumber:
        // default ''
        const docNumber = params[0];
        return createBlockCommand(wordCommand.docNumber, {
          docNumber,
        });
      case wordCommand.date:
        // default ''
        const date = params[0];
        return createBlockCommand(wordCommand.date, {
          date,
        });
      case wordCommand.division:
        // default ''
        const division = params[0];
        return createBlockCommand(wordCommand.division, {
          division,
        });
      case wordCommand.docxEngine:
        // default ''
        const docxEngine = params[0] ? params[0] : "";
        return createBlockCommand(wordCommand.docxEngine, {
          docxEngine,
        });
      case wordCommand.docxTemplate:
        // default ''
        const docxTemplate = params[0] ? params[0] : "";
        return createBlockCommand(wordCommand.docxTemplate, {
          docxTemplate,
        });
      case wordCommand.export:
        // no operation
        break;
      default:
        showMessage?.(
          MessageType.warn,
          `No word command: ${command}`,
          source,
          false
        );
        break;
    }
  }
}

export function getWordTitle(wd: string) {
  const r = getWordDownCommand(wd);
  const title = r?.command === "title" ? r.params[0] : "no title";
  return title;
}

// https://qiita.com/satokaz/items/64582da4640898c4bf42
// slugify:
export function slugify(header: string, alowDuplicate = false) {
  //return encodeURI(
  let r = header
    .trim()
    .toLowerCase()
    .replace(
      /[\]\[\!\"\#\$\%\&\'\(\)\*\+\,\.\/\:\;\<\=\>\?\@\\\^\_\{\|\}\~＠＃＄％＾＆＊（）＿＋－＝｛｝”’＜＞［］「」・、。～]/g,
      ""
    )
    .replace(/\s+/g, "-") // Replace spaces with hyphens
    .replace(/\-+$/, ""); // Replace trailing hyphen

  if (alowDuplicate === false) {
    r = createNotDuplicateId(r, r);
  }

  return r;
}

function createNotDuplicateId(id: string, originalId: string, index = 0) {
  let testId = id;
  if (idMap.has(testId)) {
    testId = createNotDuplicateId(
      originalId + "-" + (index + 1).toString(),
      originalId,
      index + 1
    );
  }
  idMap.set(testId, testId);
  return testId;
}

//
const blockHeading = (content: string, index: number) => {
  //title for asciidoc type
  const isAdocTypeTitle = false;
  if (index === 1 && isAdocTypeTitle) {
    return createBlockCommand(wordCommand.title, {
      title: content,
      subTitle: "",
    });
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
  idTitle = slugify(idTitle);

  const lines = splitBlockContents(content);
  const headings = {
    index: (index - headingOffset).toString(),
    //title: title,
    idTitle,
    text: lines.join(_newline)
  };
  return createBlockCommand(markedCommand.heading, headings);
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
        return (
          `\n${markedCommand.text}\t${markedCommand.text}\t` + s + _newline
        );
      }
    });
  return lines;
}

// a code block needs empty lines
function splitCodeBlockContents(blockContents: string) {
  const lines = blockContents.split(_newline).map((s) => {
    //if (s.match(/\t|^\/.*/)) {
    if (s.match(/\t/)) {
      // with command
      return _newline + s + _newline;
    } else {
      // only text
      return `\n${markedCommand.text}\t${markedCommand.text}\t` + s + _newline;
    }
  });
  return lines;
}

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
  return createBlockCommand(markedCommand.list, params);
};

const blockListItem = (content: string, task: boolean, checked: boolean) => {
  const lines = splitBlockContents(content);
  const params = {
    task: task ? "1" : "0",
    checked: checked ? "1" : "0",
    text: lines.join(_newline),
  };
  return createBlockCommand(markedCommand.listitem, params);
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
  return createBlockCommandTable(markedCommand.table, params);
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
  return createBlockCommand(markedCommand.tablecell, params);
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
    convertRef = slugify(convertRef, true);
    content = "";
  }

  const params = {
    href: convertRef,
    title: title ?? "",
    text: content,
  };
  return createBlockCommand(markedCommand.link, params);
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


export function markdownToWd0(
  markdown: string,
  convertType: "docx"|"excel"|"html"|"textile",
  options?: marked.MarkedOptions,
  messageFunction?: ShowMessage
): string {
  idMap.clear();
  showMessage = messageFunction;

  const typeOfConvert ={
    docx: docxRenderer,
    excel: excelRenderer,
    html: null,
    textile: textileRenderer
  };

  const render = typeOfConvert[convertType];

  let markedOptions = { ...options };
  if (render){
    markedOptions = {...markedOptions, renderer:render};
  }
  
  const unmarked = marked(markdown, markedOptions);

  const unescaped = unescape(unmarked);
  const trimmed = unescaped.trim();
  return trimmed;
}

/**
 * docx renderer
 */
const docxRenderer: marked.Renderer = {
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
const excelRenderer: marked.Renderer = {
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
const textileRenderer: marked.Renderer = {
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
