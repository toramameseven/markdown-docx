import {
  Document,
  Bookmark,
  ExternalHyperlink,
  //HeadingLevel,
  ImageRun,
  //Indent,
  //InternalHyperlink,
  Paragraph,
  ParagraphChild,
  patchDocument,

  //PatchDocumentOptions,
  PatchType,
  Table,
  TableCell,
  TableRow,
  //TextDirection,
  TableOfContents,
  TextRun,
  //VerticalAlign,
  WidthType,
  // Document as DocumentDocx,
  //convertInchesToTwip,
  //PageReference,
  SimpleField,
  TableLayoutType,
  PageBreak,
  UnderlineType,
  AlignmentType,
  convertInchesToTwip,
  LevelFormat,
  HeadingLevel,
  Packer,
  ShadingType,
} from "docx";
import { IPropertiesOptions } from "docx/build/file/core-properties";
import { FileChild } from "docx/build/file/file-child";


const propOptions:IPropertiesOptions = {
  creator: "me",
  title: "Sample Document",
  description: "Sample Document description",
  styles: {
    default: {
      heading1: {
        run: {
          size: 28,
          // bold: true,
          // italics: true,
          // color: "FF0000",
        },
        paragraph: {
          spacing: {
            after: 120,
          },
        },
      },
      heading2: {
        run: {
          size: 26,
          bold: true,
          // underline: {
          //     type: UnderlineType.DOUBLE,
          //     color: "FF0000",
          // },
        },
        paragraph: {
          spacing: {
            before: 240,
            after: 120,
          },
        },
      },
      listParagraph: {
        run: {
          color: "#FF0000",
        },
      },
      document: {
        run: {
          size: "11pt",
          font: {
            ascii: "Courier New", // Can also use minorHAnsi
            eastAsia: "ＭＳ 明朝", // Can also use minorEastAsia
            cs: "minorBidi",
            hAnsi: "Courier New",
          },
        },
        paragraph: {
          alignment: AlignmentType.LEFT,
        },
      },
    },
    paragraphStyles: [
      {
        id: "body1",
        name: "body1",
        basedOn: "document",
        next: "body1",
      },
      {
        id: "code",
        name: "code",
        basedOn: "document",
        next: "code",
        paragraph: {
        },
        run: {
          shading: {
            type: ShadingType.REVERSE_DIAGONAL_STRIPE,
            color: "00FFFF",
            fill: "FF0000",
        },
        },
      },
      {
        id: "hh1",
        name: "Hh1",
        basedOn: "Heading1",
        next: "Normal",
        paragraph: {
          numbering: {
            reference: "markHeader",
            level: 0,
          },
        },
      },
      {
        id: "hh2",
        name: "Hh2",
        basedOn: "Heading2",
        next: "Normal",
        paragraph: {
          numbering: {
            reference: "markHeader",
            level: 1,
          },
        },
      },
      {
        id: "numList1",
        name: "numList1",
        basedOn: "Heading7",
        next: "numList1",
        paragraph: {
          numbering: {
            reference: "markHeader",
            level: 6,
          },
        },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "markHeader",
        levels: [
          {
            level: 0,
            format: LevelFormat.DECIMAL,
            text: "%1.",
            alignment: AlignmentType.LEFT,
          },
          {
            level: 1,
            format: LevelFormat.DECIMAL,
            text: "%1.%2.",
            alignment: AlignmentType.LEFT,
          },
          {
            level: 6,
            format: LevelFormat.DECIMAL,
            text: "%7.",
            alignment: AlignmentType.LEFT,
          },
        ],
      },
    ],
  },
  sections: [
    {
      children: [],
    },
  ],
};

export function createDocx(children: FileChild[]) {
  let doc = new Document({...propOptions, sections:[{children: children}]});
  return doc;
}
