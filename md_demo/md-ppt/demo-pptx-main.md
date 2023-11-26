<!-- ppt parm pptxSettings clear.ppt.js -->
# Demo PPtx

## Typographic replacements

NOTE: All @ are replaced to "\\@" in **markdown to docx**.

anonymous@com.com

## Emphasis1


**This is bold text**

**This is bold text**

_This is italic text_

_This is italic text_

~~Strikethrough~~

2<sup>x</sup><sub>y</sub>

## Emphasis2

**This is bold text**
**This is bold text**
_This is italic text_
_This is italic text_
~~Strikethrough~~ 
2<sup>x</sup><sub>y</sub>

## Blockquotes

> Blockquotes can also be nested...
>
> > ...by using additional greater-than *signs* _right_ next to each other...
> >
> > * eeeeeeeee
> > > ...or with spaces between arrows.

## Lists Unordered

NOTE: **markdown to docx**  supports only three layers.


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


1. Lorem ipsum dolor sit amet
   fffffa2<sup>x</sup><sub>y</sub>affff
2. Consectetur adipiscing elit
3. Integer molestie lorem at massa

4. You can use sequential numbers...
5. ...or keep all the numbers as `1.`

NOTE: **markdown to docx** does not support numbering with offset:

57. foo
1. bar


## Lists mixed


1.  ordered1
    - unordered2
        1. ordered3

## Code

markdown


Inline `code`

    line 1 of code
    line 2 of code
    line 3 of code


## Block code syntax highlighting

NOTE: **markdown to docx** does not support Syntax highlighting.

```js
var foo = function (bar) {
  return bar++;
};
console.log(foo(5));
```



## normal table


| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |


## merge cells No.1
<!-- word emptyMerge -->

| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 |                         |


## merge cells No.2
<!-- word emptyMerge -->




| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
|         | data3-2                 |
| data4-1 | <!-- not merged -->     |

## table with emphasis



| data1-1 | data1-2                   |
| ------- | ------------------------- |
| data2-1 | _This is italic text_     |
|         | 2<sup>x</sup><sub>y</sub> |
| data4-1 | **This is bold text**     |



## table column width

<!-- word cols 2,1 -->
<!-- word emptyMerge -->
| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |



## Right aligned and rows merge



<!-- word cols 1,3 -->
<!-- word rowMerge 3-4 -->
| data1-1 | data1-2                 |
| -------:| -----------------------:|
| data2-1 | data2-2 XXXXX           |
| data3-1 | data3-2                 |
| data4-1 | data4-2                 |




## table with new line



| data1-1 | data1-2                 |
| ------- | ----------------------- |
| data2-1 | data2-2                 |
| data3-1 | data3-2<BR>data3-2-2    |
| data4-1 | data4-2                 |

## normal link



look you [link text](http://dev.nodeca.com) see me

Jason Campbell <jasoncampbell@google.com> (http://twitter.com/jxson)

* [link with title](http://nodeca.github.io/pica/demo/ "title text!")

* Autoconverted link https://github.com/nodeca/pica (enable linkify to see)

## cross refs2

* look me [refs1](#refs1) send you
* look me [refs2 with space](#refs2-with-space) send you

## Images

<!-- ppt position 10,10,10,10-->
![logo](./markdown2docx.png)

<!-- ppt position 50,50,10,10-->
This extension logo is

<!-- ppt position 80,80,10,10-->
![logo](./markdown2docx.png) .

## Math


This sentence uses dollar sign delimiters to show math inline: $\sqrt{3x-1}+(1+x)^2$  `$\sqrt{3x-1}+(1+x)^2$`

$\sqrt{3x-1}+(1+x)^2$
