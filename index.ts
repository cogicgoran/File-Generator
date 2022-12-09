import * as docx from "docx";
import { createFileSync, writeFileSync } from "fs-extra";
import { loremIpsum } from "lorem-ipsum";

const pathname = "./dist/myDocx.docx";

const SIZE_KiB_MIN = 8;
const SIZE_FILL_TO_UNIT = 8 * 1024 - 6654;
const SIZE_KiB = 9;
const KiB_bytes = 1024;

function createParagraph() {
  return loremIpsum({
    count: 1,
    units: "paragraph",
  });
}

function isSizeOverkill(text: string, sizeLimit: number) {
  return text.length >= sizeLimit;
}

function trimToLength(text: string, length: number) {
  return text.slice(0, length - 1).concat(".");
}

function createKiBSentence() {
  let textString = "";
  while (!isSizeOverkill(textString, KiB_bytes)) {
    textString = textString.concat(createParagraph());
  }
  return trimToLength(textString, KiB_bytes);
}

function createFillToFirstKiBSentence() {
  let textString = "";
  while (!isSizeOverkill(textString, SIZE_FILL_TO_UNIT)) {
    textString = textString.concat(createParagraph());
  }

  const a = trimToLength(textString, SIZE_FILL_TO_UNIT);
  console.log(a.length);
  
  return a;
}

async function createDocxDocument() {
  const { Document, Packer, Paragraph } = docx;
  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            text: '3.5.2',
          }),
        ],
      },
    ],
  });

  try {
    const b64string = await Packer.toBase64String(document);
    const documentBuffer = Buffer.from(b64string, "base64");
    console.log("Creating file...");
    createFileSync(pathname)
    writeFileSync(pathname, documentBuffer,{encoding:'utf-8'});
    console.log("File completed");
  } catch (error) {
    console.log("Error");
    console.log(error);
  }
}

createDocxDocument();
