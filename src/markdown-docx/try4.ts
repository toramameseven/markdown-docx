import Module = require('module');
import { dirname } from 'path';

const requireFromString = (code: string, filename: string = '') => {
  //const parent = module.parent || undefined;
  const m = new Module("");

  //m.filename = filename;

  // @ts-ignore
  // const paths: string[] = Module._nodeModulePaths(dirname(filename));
  // m.paths = paths;

  // @ts-ignore
  m._compile(code, filename);

  return m.exports;
};

const code = `
let array = [
  {
    type: "string",
    params:[
    [
      { text: "Sub" },
      { text: "Subscript", options: { subscript: true } },
      { text: " // Super" },
      { text: "Superscript", options: { superscript: true } },
    ],
    { x: 10, y: 6.3, w: 3.3 }
  ]
  },
];
module.exports = {array}
`;

const {array} = requireFromString(code);

console.log(array); // 3