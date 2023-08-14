import { svg2png, initialize } from "svg2png-wasm";
import { readFileSync } from "fs";
import { selectExistsPath } from "./tools-common";

let initialized = false;

export async function svg2imagePng(svg: string) {
  const wasmPath = await selectExistsPath(
    ["../wasm/svg2png_wasm_bg.wasm", "../../wasm/svg2png_wasm_bg.wasm","../../../wasm/svg2png_wasm_bg.wasm"],
    [__dirname]
  );

  try {
    if (!initialized) {
      await initialize(readFileSync(wasmPath));
      initialized = true;
    }
  } catch (e) {
    throw e;
  } finally {
    //
  }

  /** @type {Uint8Array} */
  const png: Uint8Array = await svg2png(svg, {
    // scale: 2, // optional
    // width: 400, // optional
    // height: 400, // optional
    // backgroundColor: "white", // optional
    // fonts: [
    //   // optional
    //   readFileSync('./Roboto.ttf'), // require, If you use text in svg
    // ],
    // defaultFontFamily: {
    //   // optional
    //   sansSerif: 'Roboto',
    // },
  });
  const pb = Buffer.from(png);
  return pb;
}
