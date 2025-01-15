<!-- word placeholder subTitle subTitle -->
<!-- word placeholder author author -->
<!-- word placeholder division division -->
<!-- word placeholder date date -->
<!-- word placeholder docNumber XXXX-XXXX -->

<!-- word docxTemplate _with_cover.docx -->

# Markdown Docx README

<!-- word toc 1 TOC -->
<!-- word newPage -->

This is the README **Markdown Docx**. 
This extension converts a markdown file to a Docx or a Pptx (experimental).
This uses next two excellent modules.

* [DOCX](https://docx.js.org/#/)
* [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)

The document is [here](https://toramameseven.github.io/markdown-docx-doc/en/)

## Requirements

* Windows 10

## Features

* **Markdown Docx** is a markdown converter to docx.
* **Markdown Docx** works for common mark md(s).
* Click **Convert Docx** at the context menu on the explore or the editor.
* In the editor, you can convert only the selection in the text.
* In the output tab, markdown-docx, the progress and the warns are displayed.

## Features(Experimental)

* Convert a markdown to pptx

 
## markdown vscode settings

like below

```markdown
  "[markdown]": {
    "editor.wordWrap": "off",
    "editor.quickSuggestions": {
      "other": true,
      "comments": false,
      "strings": false
    },
    "editor.snippetSuggestions": "top"
  },
```

## word template

It is better, set your language font.

### template files

You can see the some template in the [markdown-docx site](https://github.com/toramameseven/markdown-docx) templates folder.

* _with_cover.docx
* _no_cover.docx (default template)

In these template, you see the placeholder described at next section.

### place holder

[DOCX](https://docx.js.org/#/) type place folder is used.

Next place holders are used in the sample template.

* main content
  * `{{paragraphReplace}}
  * do not set this other information.

  ![](./images/main_placeholder.png)

* for cover
  * `{{title}}`
  * `{{subTitle}}`
  * `{{author}}`
  * `{{division}}`
  * `{{date}}`
  * `{{docNumber}}`

  ![](./images/cover_placeholder.png)

markdown  
```
<!-- word placeholder title "sample document" -->
```

docx template
```
{{title}}
```

`{{title}}` is replaced to "sample document".


## Known Issues

* Inline math does not work.
* HTML does not work.
* Block quote does no work.
* The indent of table of contents is not good.


## How to package

1. npm install -g @vscode/vsce
1. vsce package --target win32-x64
1. vsce publish

## Acknowledgments

We thank for the wonderful npm packages.

[Packages](usedModules.md)

some feature are not active now.

And we use some useful articles below. 
* [markdown-to-txt](https://www.npmjs.com/package/markdown-to-txt) tells us how to use **Marked**.
* [marked-extended-tables](https://github.com/calculuschild/marked-extended-tables) is for merged table.
* To Slugify, we use Mr. Sato 's code (https://qiita.com/satokaz/items/64582da4640898c4bf42)
* [node-html-markdown](https://github.com/crosstype/node-html-markdown)'s code is used, for converting html to markdown.


## Release Notes
* 0.0.5
  * add new line under a image.

* 0.0.4
  * experimental feature creating pptx.

* 0.0.3
  * add check box.
  
* 0.0.2
  * use [DOCX](https://docx.js.org/#/) for creating word files.
  * we do not support the vbs rendering on version `0.0.2`.
    
* 0.0.1
  * first Release.


