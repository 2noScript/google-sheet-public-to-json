import XLSX from "xlsx";

//___________Developed by 2noscript_____________

interface DynamicKeyObject {
  [key: string | number]: any;
}

export function getUrlExcelSheet(url: string) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    const sheetKey = match[1];
    const baseUrl = url.split(sheetKey)[0];
    return `${baseUrl}${sheetKey}/export?format=xlsx`;
  }
  return "ERROR";
}

const getHeaderRowCount = (workbook: any) => {
  const sheetNameList = workbook.SheetNames;
  const sheet = workbook.Sheets[sheetNameList[0]];

  const headers: DynamicKeyObject = {};
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

export default async function sheetPublicToJson(linkSheetPublic: string) {
  const linkXLSX = getUrlExcelSheet(linkSheetPublic);
  const response = await fetch(linkXLSX);
  const arrayBuffer = await response.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });

  const headerRowCount = getHeaderRowCount(workbook);

  const sheetNameList = workbook.SheetNames;
  const sheet: DynamicKeyObject = workbook.Sheets[sheetNameList[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headersArr = [];

  for (let R = range.s.r; R < headerRowCount; R++) {
    const headers: DynamicKeyObject = {};
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

  const header: DynamicKeyObject = {};
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
    const row: DynamicKeyObject = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];
      row[header[C]] = cell ? cell.v : "";
    }
    dataJson.push(row);
  }

  return dataJson;
}
