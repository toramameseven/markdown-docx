# Markdown Docx README

This is the README **Markdown Docx**. 


## Requirements

* Windows 10
* Microsoft Word

## Features

* **Markdown Docx** Docx is a markdown converter to docx.
* **Markdown Docx** Docx works for common mark md(s).
* Click **convert Docx** at the context menu on the explore or the editor.
  
## extensions for word

##### general

`<!-- word [command] parameters -->` is used.

`<!-- word title Title -->`

`<!-- word subTitle SubTitle -->`

`<!-- word author Author -->`

`<!-- word division Division -->`

`<!-- word date Date -->`

`<!-- word toc 1 TOC -->`

    1: levels of toc.
    TOC: toc caption

`<!-- word import imported.md-->`

  imported.md will be imported.

`<!-- word pageSetup wdOrientationLandscape wdSizeA4 -->`
  
  page setup sample

`<!-- word pageSetup wdOrientationPortrait wdSizeA3 -->`

  page setup sample

`<!-- word newPage -->`

  insert new page

##### table

`<!-- word cols 1,2 -->`

    columns width are 1:2

`<!-- word rowMerge 1-4,5-6 -->`

  rows 1-4 and 5-6 are merged.

`<!-- word emptyMerge -->`
  
  empty cells are merged. only row direction.

## sample markdown file

You can see the sample file.

## Extension Settings

### markdown docx

* markdown-docx.path.docxEngine

    Set your original docx rendering vbs.

* markdown-docx.path.docxTemplate

    Set your original docx file for template.

* markdown-docx.docxEngine.mathExtension
   
   If set true, 

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
* [markdown-to-txt](https://www.npmjs.com/package/markdown-to-txt) tell us how to use **Marked**.
* [木村工の Office 仕事術](https://koukimra.com/) is used to resize pictures.
* [みんなのワードマクロ](https://www.wordvbalab.com/) is used for emphasis styles.
* To Slugify, we use Mr. Sato 's code (https://qiita.com/satokaz/items/64582da4640898c4bf42)



## Release Notes

* No Release


