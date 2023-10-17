import XLSX from "xlsx";

//___________Developed by 2noscript_____________

export function getUrlExcelSheet(url: string) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    const sheetKey = match[1];
    const baseUrl = url.split(sheetKey)[0];
    return `${baseUrl}${sheetKey}/export?format=xlsx`;
  }
  return "ERROR";
}

export const getHeaderRowCount = (workbook: any) => {
  const sheetNameList = workbook.SheetNames;
  const sheet = workbook.Sheets[sheetNameList[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);

  let parsedNumHeaders = 0;

  for (let R = range.s.r; R <= range.e.r; R++) {
    let isParsed = true;
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];
      if (cell && cell.t === "n") {
        isParsed = false;
        break;
      }
    }

    if (isParsed) {
      parsedNumHeaders++;
    } else {
      break;
    }
  }
  return parsedNumHeaders;
};

export default async function sheetPublicToJson(linkSheetPublic: string) {
  const linkXLSX = getUrlExcelSheet(linkSheetPublic);
  const response = await fetch(linkXLSX);
  const arrayBuffer = await response.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: "array" });

  const headerRowCount = getHeaderRowCount(workbook);
  const sheetNameList = workbook.SheetNames;
  const sheet: any = workbook.Sheets[sheetNameList[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headersArr: any = [];

  for (let R = range.s.r; R < headerRowCount; R++) {
    const headers: any = {};
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
    headersArr.push(headers);
  }

  let header: any = {};
  for (let i = 0; i < headersArr.length; i++) {
    for (const key in headersArr[i]) {
      const index = Number(key);
      if (i === 0) {
        if (index === 0 || headersArr[i][key] !== "") {
          header[key] = headersArr[i][key];
        } else {
          let k = 0;
          while (headersArr[i][index - k] === "") {
            header[key] = headersArr[i][index - k - 1];
            k++;
          }
        }
      } else {
        if (headersArr[i][key] !== "") {
          header[key] += "." + headersArr[i][key];
        }
        if (headersArr[i][key] === "") {
          if (headersArr[0][key] === "") {
            let k = 0;
            while (headersArr[i][index - k] === "") {
              header[key] += "." + headersArr[i][index - k - 1];
              k++;
            }
          }
        }
      }
    }
  }
  const dataJson = [];

  for (let R = 3; R <= range.e.r; R++) {
    const row: any = {};
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
sheetPublicToJson(
  "https://docs.google.com/spreadsheets/d/1vd2XOjo_dk069cEOUvGDAEKgnjdqcUMnbA_JaVbWboE/edit#gid=0"
).then((data) => console.log(data));
