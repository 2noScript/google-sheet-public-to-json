import { getUrlExcelSheet } from "./src/utils.js";
import XLSX from "xlsx";

//___________Developed by 2noscript_____________

export const getHeaderRowCount = (workbook) => {
  const sheetNameList = workbook.SheetNames;
  const sheet = workbook.Sheets[sheetNameList[0]];

  const headers = {};
  let headerRowCount = 0;

  for (const cellAddress in sheet) {
    if (cellAddress[0] === "!") continue;
    const cell = sheet[cellAddress];
    if (cell && cell.v) {
      const header = cell.v.toString().trim();
      if (!headers[cellAddress[0]]) {
        headers[cellAddress[0]] = header;
        headerRowCount++;
      } else if (headers[cellAddress[0]] !== header) {
        break;
      }
    }
  }
  return headerRowCount;
};

const sheetPublicToJson = async (linkSheetPublic) => {
  const linkXLSX = getUrlExcelSheet(linkSheetPublic);
  const response = await fetch(linkXLSX);
  const arrayBuffer = await response.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });

  const headerRowCount = getHeaderRowCount(workbook);

  const sheetNameList = workbook.SheetNames;
  const sheet = workbook.Sheets[sheetNameList[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headersArr = [];

  for (let R = range.s.r; R < headerRowCount; R++) {
    const headers = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];
      headers[C] = (headers[C] ? headers[C] : "") + (cell ? cell.v : "");
    }
    let i = 0;
    while (i < range.e.c) {
      try {
        if (headers[i] === "") headers[i] = headers[i - 1];
      } catch {}
      i++;
    }
    headersArr.push(headers);
  }

  const header = {};
  for (let i = 0; i < headersArr.length; i++) {
    for (const key in headersArr[i]) {
      let textCurrentHeader = "";
      if (headersArr[i][key] !== undefined && headersArr[i][key] !== "") {
        textCurrentHeader =
          i == 0 ? headersArr[i][key] : "." + headersArr[i][key];
      }
      header[key] =
        header[key] !== undefined
          ? header[key] + textCurrentHeader
          : textCurrentHeader;
      if (header[key].startsWith(""))
        header[key] = header[key].replace("-", "");
    }
  }
  const dataJson = [];

  for (let R = 3; R <= range.e.r; R++) {
    const row = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];
      row[header[C]] = cell ? cell.v : "";
    }
    dataJson.push(row);
  }

  return dataJson;
};

export default sheetPublicToJson;