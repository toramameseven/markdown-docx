# Markdown Docx README

This is the README **Markdown Docx**. 

This extension uses a docx binary file for the template. 
If security check happens, you can download the template form this repo and set it as template in the settings.

## Requirements

* Windows 10
* Microsoft Word

## Features

* **Markdown Docx** is a markdown converter to docx.
* **Markdown Docx** works for common mark md(s).
* Click **Convert Docx** at the context menu on the explore or the editor.
* In the editor, you can convert only the selection in the text.
* In the output tab, markdown-docx, the progress and the warns are displayed.
  
## Extensions for word

##### general

`<!-- word [command] parameters -->` is used for word command.

* `<!-- word title Title -->`

    add title

* `<!-- word subTitle SubTitle -->`
  
    add sub title
* `<!-- word author Author -->`

    add author
* `<!-- word division Division -->`

    add division
* `<!-- word date Date -->`

    add Date
* `<!-- word toc 1 "table of contents" -->`

    * add toc
    * 1: levels of toc.
    * table of contents: toc caption

* `<!-- word import imported.md-->`

  imported.md will be imported.

* `<!-- word pageSetup wdOrientationLandscape wdSizeA4 -->`
  
    page setup sample. landscape and a4 size.

* `<!-- word pageSetup wdOrientationPortrait wdSizeA3 -->`

    page setup sample. portrait and a3 size

* `<!-- word newPage -->`

    insert new page

##### table

* `<!-- word cols 1,2 -->`

    columns width are 1:2

* `<!-- word rowMerge 1-4,5-6 -->`

    rows 1-4 and 5-6 are merged.

* `<!-- word emptyMerge -->`
  
    empty cells are merged. only row direction.

## sample markdown file

You can see the sample file in the [markdown-docx site](https://github.com/toramameseven/markdown-docx) md_demo folder.

## Extension Settings

### markdown docx

* markdown-docx.path.docxEngine

    Set your original docx rendering vbs.

* markdown-docx.path.docxTemplate

    Set your original docx file for template.

* markdown-docx.docxEngine.mathExtension
   
   If set true, `$x+1$` type math is rendered.

* markdown-docx.docxEngine.timeout

    60000 ms is default. docx rendering is so slow, you can set bigger value.

* markdown-docx.docxEngine.debug
   
    some debug option is enabled.

    * intermediate files *.wd0, *.wd are not deleted.
  
### markdown vscode settings

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

##### styles

next styles are created.

* author1
* blockA
* body1
* body2
* body3
* BodyTitle
* code
* codespan
* date1
* division1
* nList1
* nList2
* nList3
* note1
* numList1
* numList2
* numList3
* picture1
* styleN
* warn1
* wdHeading5

### user properties

* dNumber
  * number is displayed at header.
* dDivision
* dDate
* dAuthor

## Known Issues

* Inline Images do not work.
* HTML does not work.

## How to package

1. npm install -g vsce
1. vsce package --target win32-x64
1. vsce publish

## Acknowledgments

We thank for the wonderful npm packages.

And we use some useful articles below. 

* [Marked](https://www.npmjs.com/package/marked) is a very useful package for this extension.
* [markdown-to-txt](https://www.npmjs.com/package/markdown-to-txt) tells us how to use **Marked**.
* [koukimura's page](https://koukimra.com/) is used to resize pictures.
* [minnano macro page](https://www.wordvbalab.com/) is used for emphasis styles.
* To Slugify, we use Mr. Sato 's code (https://qiita.com/satokaz/items/64582da4640898c4bf42)



## Release Notes

* 0.0.1
  * first Release.


