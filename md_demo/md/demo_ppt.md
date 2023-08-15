<!-- word placeholder subTitle "sample document subTitle" -->
<!-- word placeholder division divisionXXX -->
<!-- word placeholder date dateXXX -->
<!-- word placeholder author authorXXX -->
<!-- word placeholder docNumber XXXX-0004 -->
<!-- word newLine -->
<!-- word toc 3 "toc caption" -->
<!-- word levelOffset 0 -->


<!-- https://markdown-it.github.io/ -->

# this is sample document

## title subtitle and table of content 

<!-- ppt position 5,20,45,80 -->

markdown

```
<!-- word title sample document title -->
<!-- word subTitle sample document subTitle -->
<!-- word toc 3 -->
```
<!-- ppt position 55,20,45,80 -->

The result is above.


This sample markdown is modified from the document of [markdown-it](https://markdown-it.github.io/).

<!-- word export demo-headings.md-->
## Heading


<!-- ppt position 5,20,45,80 -->
markdown

```
### Heading2

#### Heading3

##### Heading4

###### Heading5

```

<!-- ppt position 55,20,45,80 -->

result

### Heading2

#### Heading3

##### Heading4

###### Heading5





<!-- word export demo-Horizontal.md-->
## Horizontal Rules


<!-- ppt position 5,20,45,80 -->
markdown

```
___

---

***
```
 
Horizontal Rules do not work.



<!-- word export demo-br.md-->
## BR

<!-- ppt position 5,20,45,80 -->
markdown

```
next line is `<br>`

test<BR>

<BR>

upper line is `<br>`

```
<!-- ppt position 55,20,45,80 -->
result

next line is `<br>`

test<BR>

<BR>

upper line is `<br>`

<!-- word export demo-newpage.md-->
## new page

### markdown

markdown

```
next line is `<!-- word newPage -->`

<!-- word newPage -->

upper line is `<!-- word newPage -->`
```
result

`<!-- word newPage -->` does not work.



<!-- word export demo-Emphasis.md-->
## Emphasis

<!-- ppt position 5,20,45,80 -->
markdown

```
**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>
```
<!-- ppt position 55,20,45,80 -->
result

**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>

<!-- word export demo-Emphasis2.md-->
## Emphasis2

<!-- ppt position 5,20,45,80 -->
markdown

```
**This is bold text**
**This is bold text**
_This is italic text_
_This is italic text_
~~Strikethrough~~ 
2<sup>x</sup><sub>y</sub>
```
<!-- ppt position 55,20,45,80 -->
result

**This is bold text**
**This is bold text**
_This is italic text_
_This is italic text_
~~Strikethrough~~ 
2<sup>x</sup><sub>y</sub>


<!-- word export demo-Blockquotes.md-->
## Blockquotes
<!-- ppt position 5,20,45,80 -->
NOTE: **markdown to docx** does not support blockquote.



<!-- word export demo-list-unordered.md-->
## Lists Unordered

<!-- ppt position 5,20,45,80 -->


NOTE: **markdown to docx**  supports only three layers.

markdown

```
- Create a list by starting a line with `+`, `-`, or `*`
- Sub-lists are made by indenting 2 spaces:
  - Marker character change forces new list start:
    - Ac tristique libero volutpat at
    * Facilisis in pretium nisl aliquet
    - Nulla volutpat aliquam velit
- Very easy!

* look me [](#links) send you
```

<!-- ppt position 55,20,45,80 -->

result

- Create a list by starting a line with `+`, `-`, or `*`
- Sub-lists are made by indenting 2 spaces:
  - Marker character change forces new list start:
    - Ac tristique libero volutpat at
    * Facilisis in pretium nisl aliquet
    - Nulla volutpat aliquam velit
- Very easy!

* look me [](#links) send you

<!-- word export demo-list-orderd.md-->
## Lists ordered

<!-- ppt position 5,20,45,80 -->

markdown

```
1. Lorem ipsum dolor sit amet
2. Consectetur adipiscing elit
3. Integer molestie lorem at massa

4. You can use sequential numbers...
5. ...or keep all the numbers as `1.`

NOTE: **markdown to docx** does not support numbering with offset:

57. foo
1. bar
```

<!-- ppt position 55,20,45,80 -->

result

1. Lorem ipsum dolor sit amet
   fffffa2<sup>x</sup><sub>y</sub>affff
2. Consectetur adipiscing elit
3. Integer molestie lorem at massa

4. You can use sequential numbers...
5. ...or keep all the numbers as `1.`

NOTE: **markdown to docx** does not support numbering with offset:

57. foo
1. bar

<!-- word export demo-list-mixed.md-->
## Lists mixed


<!-- ppt position 5,20,45,80 -->

markdown

```
1.  ordered1
    - unordered2
        1. ordered3
```

<!-- ppt position 55,20,45,80 -->

result

1.  ordered1
    - unordered2
        1. ordered3


<!-- word export demo-code.md-->
## Code

<!-- ppt position 5,20,45,80 -->

markdown

```
Inline `code`
```
<!-- ppt position 55,20,45,80 -->

result

Inline `code`

<!-- word export demo-code-indented.md-->
## Indented code

<!-- ppt position 5,20,45,80 -->

markdown

```
    // Some comments
    line 1 of code
    line 2 of code
    line 3 of code
```
<!-- ppt position 55,20,45,80 -->
result

    // Some comments
    line 1 of code
    line 2 of code
    line 3 of code

<!-- word export demo-code-fences.md-->
## Block code syntax highlighting

<!-- ppt position 5,20,45,80 -->

NOTE: **markdown to docx** does not support Syntax highlighting.


<!-- word export demo-table.md-->
## Tables


## normal table

<!-- ppt position 5,20,45,80 -->

markdown

```
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |
```
<!-- ppt position 55,20,45,80 -->

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


<!-- word export demo-table-merge.md-->
## merge cells No.1
<!-- ppt position 5,20,45,80 -->

* cell(3,1) and cell(4,2) are merged.


markdown
```
<!-- word emptyMerge -->

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 |                         |
```

<!-- ppt position 55,20,45,80 -->

<!-- word emptyMerge -->

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 |                         |


<!-- word export demo-table-merge2.md-->
## merge cells No.2

<!-- ppt position 5,20,45,80 -->

markdown

```
<!-- word emptyMerge -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |
```
<!-- ppt position 55,20,45,80 -->


<!-- word emptyMerge -->

cell(4,2) is not merged. (comment cell)

<!-- ppt position 55,30,45,80 -->


| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |

## table with emphasis

<!-- ppt position 5,20,45,80 -->

markdown

```
| data1-1 | data1-2                   |
| ------- | --------------------  --- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |
```

<!-- ppt position 55,20,45,80 -->


| data1-1 | data1-2                   |
| ------- | ------------------------- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |


<!-- word export demo-table-columns.md-->
## table column width

<!-- ppt position 5,20,45,80 -->

markdown

```
<!-- word cols 2,1 -->
<!-- word emptyMerge -->
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |
```

<!-- ppt position 55,20,45,80 -->

<!-- word cols 2,1 -->
<!-- word emptyMerge -->
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


<!-- word export demo-table-rowmerge.md-->
## Right aligned and rows merge

<!-- ppt position 5,20,45,80 -->

markdown

```
<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |
```

<!-- ppt position 55,20,45,80 -->

<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


NOTE: aligned is not worked


## table with new line

<!-- ppt position 5,20,45,80 -->

markdown

```
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<bR>data3-2-2    |
| data4-1 | data4-2                 |
```

<!-- ppt position 55,20,45,80 -->


| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<BR>data3-2-2    |
| data4-1 | data4-2                 |



<!-- word export demo-normalLink.md-->
### normal link

<!-- ppt position 5,20,45,80 -->

```
look you [link text](http://dev.nodeca.com) see me

Jason Campbell <jasoncampbell@google.com> (http://twitter.com/jxson)

* [link with title](http://nodeca.github.io/pica/demo/ "title text!")

* Autoconverted link https://github.com/nodeca/pica (enable linkify to see)
```

<!-- ppt position 55,20,45,80 -->


look you [link text](http://dev.nodeca.com) see me

Jason Campbell <jasoncampbell@google.com> (http://twitter.com/jxson)

* [link with title](http://nodeca.github.io/pica/demo/ "title text!")

* Autoconverted link https://github.com/nodeca/pica (enable linkify to see)


<!-- word export demo-image.md-->
## Images

<!-- ppt position 5,20,45,80 -->

markdown

```
![logo](./markdown2docx.png)

This extension logo is  ![logo](./markdown2docx.png) .
```
<!-- ppt position 55,20,45,80 -->
result

NOTE: **markdown to docx**  supports only image files.

<!-- ppt position 55,30,45,80 -->
![logo](./markdown2docx.png)


<!-- word export demo-math.md-->
## Math

NOTE: math works so so.


<!-- word export demo-admonition.md-->
## Admonition

markdown

```
NOTE: admonition note.

WARNING: admonition warning.
```

result

NOTE: admonition note.

WARNING: admonition warning.
