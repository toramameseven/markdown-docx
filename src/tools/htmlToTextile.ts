var toTextile = require('to-textile');

export function htmlToTextile(filePath: string, body: string){

  const r = toTextile(body);
  return r;
}


console.log(htmlToTextile('','<h1>dddddddd</>'));
