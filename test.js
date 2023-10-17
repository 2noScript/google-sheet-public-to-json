import { getUrlExcelSheet } from "./src/utils.js";
import XLSX from "xlsx";

const handleGenerateData = (arr, lengthHeader) => {
  const [firstRow, ...restRow] = [...arr];
  const finalHeader = firstRow;
  const finalData = [];

  restRow.forEach((row, indexRow) => {
    const objData = {};
    if (indexRow <= lengthHeader - 2) {
      // Tìm finalHeader 0 - 1 - 2 = 4 - 2
      row.forEach((item, index) => {
        if (item !== finalHeader[index]) {
          finalHeader[index] = [finalHeader[index], item].join(".");
        }
      });
    } else {
      row.forEach((item, index) => {
        objData[`${finalHeader[index]}`.replace(/ /g, "")] = item;
      });
      finalData.push(objData);
    }
  });

  return finalData;
};

const linkSheetPublic =
  "https://docs.google.com/spreadsheets/d/1K6HonBg-o2x0riE9_Ba-1hckJWl8Alg8-SXHjgT_ZWc/edit#gid=0";

const linkXLSX = getUrlExcelSheet(linkSheetPublic);
const response = await fetch(linkXLSX);
const arrayBuffer = await response.arrayBuffer();
const data = new Uint8Array(arrayBuffer);
const workbook = XLSX.read(data, { type: "array" });

const sheetNames = workbook.SheetNames;
let roa;
for (let worksheet of sheetNames) {
  if (workbook.Sheets[worksheet]["!merges"]) {
    workbook.Sheets[worksheet]["!merges"].map((merge) => {
      const value = XLSX.utils.encode_range(merge).split(":")[0];
      for (let col = merge.s.c; col <= merge.e.c; col++)
        for (let row = merge.s.r; row <= merge.e.r; row++) {
          workbook.Sheets[worksheet][
            String.fromCharCode(65 + col) + (row + 1)
          ] = workbook.Sheets[worksheet][value];
        }
    });
  }

  roa = XLSX.utils.sheet_to_json(workbook.Sheets[worksheet], {
    header: 1,
  });
}

console.log(handleGenerateData(roa, 3));
