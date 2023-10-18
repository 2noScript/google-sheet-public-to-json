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

export const getHeaderRowCount = (workbook: XLSX.WorkBook) => {
  const sheetNameList = workbook.SheetNames;
  const sheet: any = workbook.Sheets[sheetNameList[0]];
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

export const sheetPublicToJson = async (linkSheetPublic: string) => {
  const linkXLSX = getUrlExcelSheet(linkSheetPublic);
  const response = await fetch(linkXLSX);
  const arrayBuffer = await response.arrayBuffer();
  const data = new Uint8Array(arrayBuffer);
  const workbook: XLSX.WorkBook = XLSX.read(data, { type: "array" });

  const headerRowCount = getHeaderRowCount(workbook);
  const sheetNameList = workbook.SheetNames;
  const sheet: any = workbook.Sheets[sheetNameList[0]];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headersArr = [];

  for (let R = range.s.r; R < headerRowCount; R++) {
    const headers: any = {};
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellAddress = { c: C, r: R };
      const cellRef = XLSX.utils.encode_cell(cellAddress);
      const cell = sheet[cellRef];

      headers[C] = (headers[C] ? headers[C] : "") + (cell ? cell.v : "");
    }
    headersArr.push(headers);
  }

  let header: any = {};
  for (let i = 0; i < headersArr.length; i++) {
    for (const index in headersArr[i]) {
      const key: number = Number(index);
      if (i === 0) {
        if (key == 0 || headersArr[i][key] !== "") {
          header[key] = headersArr[i][key];
        } else {
          let k = 0;
          while (headersArr[i][key - k] === "") {
            header[key] = headersArr[i][key - k - 1];
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
            while (headersArr[i][key - k] === "") {
              header[key] += "." + headersArr[i][key - k - 1];
              k++;
            }
          }
        }
      }
    }
  }
  const dataJson: any = [];

  for (let R = headerRowCount; R <= range.e.r; R++) {
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
};

sheetPublicToJson(
  "https://docs.google.com/spreadsheets/d/1K6HonBg-o2x0riE9_Ba-1hckJWl8Alg8-SXHjgT_ZWc/edit#gid=0"
).then((data) => console.log(data[0]));
