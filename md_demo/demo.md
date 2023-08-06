<!-- word placeholder title "sample document" -->
<!-- word placeholder subTitle "sample document subTitle" -->
<!-- word placeholder division divisionXXX -->
<!-- word placeholder date dateXXX -->
<!-- word placeholder author authorXXX -->
<!-- word placeholder docNumber XXXX-0004 -->
<!-- word newLine -->
<!-- word toc 3 "toc caption" -->
<!-- word levelOffset 0 -->


<!-- https://markdown-it.github.io/ -->

## title subtitle and table of content 

markdown

```
<!-- word title sample document title -->
<!-- word subTitle sample document subTitle -->
<!-- word toc 3 -->
```

The result is above.


This sample markdown is modified from the document of [markdown-it](https://markdown-it.github.io/).

<!-- word export demo-headings.md-->
## Heading

markdown

```
## Heading1

### Heading2

#### Heading3

##### Heading4

###### Heading5

####### Heading6
```

result

### Heading2

#### Heading3

##### Heading4

###### Heading5

####### Heading6




<!-- word export demo-Horizontal.md-->
## Horizontal Rules

markdown

```
___

---

***
```

result
___

---

***


<!-- word export demo-br.md-->
## BR

markdown

```
next line is `<br>`

test<BR>

<BR>

upper line is `<br>`
```
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

next line is `<!-- word newPage -->`

<!-- word newPage -->

upper line is `<!-- word newPage -->`


<!-- word export demo-Typographic.md-->
## Typographic replacements

NOTE: All @ are replaced to "\\@" in **markdown to docx**.

anonymous@com.com


<!-- word export demo-Emphasis.md-->
## Emphasis

markdown

```
**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>
```
result

**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>

<!-- word export demo-Emphasis2.md-->
## Emphasis2

markdown

```
**This is bold text**
**This is bold text**
_This is italic text_
_This is italic text_
~~Strikethrough~~ 
2<sup>x</sup><sub>y</sub>
```

result

**This is bold text**
**This is bold text**
_This is italic text_
_This is italic text_
~~Strikethrough~~ 
2<sup>x</sup><sub>y</sub>

NOTE: Sometime Emphasis does not work. That time, please add some spaces between words.


<!-- word export demo-Blockquotes.md-->
## Blockquotes

NOTE: **markdown to docx** does not support blockquote.

markdown

```
> Blockquotes can also be nested...
>
> > ...by using additional greater-than *signs* _right_ next to each other...
> >
> > * eeeeeeeee
> > > ...or with spaces between arrows.
```
result

> Blockquotes can also be nested...
>
> > ...by using additional greater-than *signs* _right_ next to each other...
> >
> > * eeeeeeeee
> > > ...or with spaces between arrows.


<!-- word export demo-list-unordered.md-->
## Lists Unordered

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

markdown

```
1.  ordered1
    - unordered2
        1. ordered3
```
result

1.  ordered1
    - unordered2
        1. ordered3


<!-- word export demo-code.md-->
## Code

markdown

```
Inline `code`
```
result

Inline `code`

<!-- word export demo-code-indented.md-->
## Indented code

markdown

```
    // Some comments
    line 1 of code
    line 2 of code
    line 3 of code
```
result

    // Some comments
    line 1 of code
    line 2 of code
    line 3 of code

<!-- word export demo-code-fences.md-->
## Block code syntax highlighting

NOTE: **markdown to docx** does not support Syntax highlighting.

```js
var foo = function (bar) {
  return bar++;
};
console.log(foo(5));
```


<!-- word export demo-table.md-->
## Tables

### normal table

markdown

```
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |
```

result

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


<!-- word export demo-table-merge.md-->
### merge cells No.1

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

result


<!-- word emptyMerge -->

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 |                         |


<!-- word export demo-table-merge2.md-->
### merge cells No.2

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
result

<!-- word emptyMerge -->

cell(4,2) is not merged. (comment cell)

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |

### table with emphasis

markdown

```
| data1-1 | data1-2                   |
| ------- | --------------------  --- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |
```

result

| data1-1 | data1-2                   |
| ------- | ------------------------- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |


<!-- word export demo-table-columns.md-->
### table column width

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

result

<!-- word cols 2,1 -->
<!-- word emptyMerge -->
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


<!-- word export demo-table-rowmerge.md-->
### Right aligned and rows merge

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

result

<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |

NOTE: aligned is not worked


### table with new line

markdown

```
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<bR>data3-2-2    |
| data4-1 | data4-2                 |
```
result

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<BR>data3-2-2    |
| data4-1 | data4-2                 |

NOTE: does not work



## Links

<!-- word export demo-links-xref.md-->
### cross reference in a docx

markdown

```
### refs1

### refs2 with space

* look me [refs1](#refs1) send you
* look me [refs2 with space](#refs2-with-space) send you
```

result

### refs1

### refs2 with space

* look me [refs1](#refs1) send you
* look me [refs2 with space](#refs2-with-space) send you


<!-- word export demo-normalLink.md-->
### normal link

```
look you [link text](http://dev.nodeca.com) see me

Jason Campbell <jasoncampbell@google.com> (http://twitter.com/jxson)

* [link with title](http://nodeca.github.io/pica/demo/ "title text!")

* Autoconverted link https://github.com/nodeca/pica (enable linkify to see)
```

look you [link text](http://dev.nodeca.com) see me

Jason Campbell <jasoncampbell@google.com> (http://twitter.com/jxson)

* [link with title](http://nodeca.github.io/pica/demo/ "title text!")

* Autoconverted link https://github.com/nodeca/pica (enable linkify to see)


<!-- word export demo-image.md-->
## Images

markdown

```
![logo](./markdown2docx.png)

This extension logo is  ![logo](./markdown2docx.png) .
```

result

![logo](./markdown2docx.png)

This extension logo is  ![logo](./markdown2docx.png) .

NOTE: **markdown to docx**  supports only image files.

NOTE: Inline images do not work well.

<!-- word export demo-math.md-->
## Math

markdown

```
This sentence uses dollar sign delimiters to show math inline: $\sqrt{3x-1}+(1+x)^2$  `$\sqrt{3x-1}+(1+x)^2$`

$\sqrt{3x-1}+(1+x)^2$
```

result

This sentence uses dollar sign delimiters to show math inline: $\sqrt{3x-1}+(1+x)^2$  `$\sqrt{3x-1}+(1+x)^2$`

$\sqrt{3x-1}+(1+x)^2$

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
