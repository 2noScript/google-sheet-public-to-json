import sheetPublicToJson from "index.js";

const data = await sheetPublicToJson(
  "https://docs.google.com/spreadsheets/d/1vd2XOjo_dk069cEOUvGDAEKgnjdqcUMnbA_JaVbWboE/edit#gid=0"
);

console.log(data);
