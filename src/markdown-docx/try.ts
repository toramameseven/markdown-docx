const isNumber = (value: unknown): value is number => {
  return typeof value === "number";
};

function numberToStirng(value: unknown) {
  if (isNumber(value)) {
    console.log("number");
    return value.toString();
  }
  console.log("not number");
  return value;
}


console.log(numberToStirng(1));
console.log(numberToStirng("s"));