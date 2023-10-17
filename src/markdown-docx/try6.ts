function getLinkType(linkRef: string): "section"|"caption" {
  if (linkRef.slice(0, "fig.".length) === "fig."){
    return "caption";
  }
  if (linkRef.slice(0, "table.".length) === "table."){
    return "caption";
  }
  return "section";
}


// console.log(getLinkType("table.xxxxx"));
// console.log(getLinkType("fig.xxxxx"));
// console.log(getLinkType("figxxxxx"));

console.log("abcde".substring(0,3));
