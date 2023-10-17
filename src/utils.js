import XLSX from "xlsx";

export function getUrlJsonSheet(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    const sheetKey = match[1];
    const baseUrl = url.split(sheetKey)[0];
    return `${baseUrl}${sheetKey}/gviz/tq?tqx=out:json`;
  }
  return "ERROR";
}

export function getUrlExcelSheet(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    const sheetKey = match[1];
    const baseUrl = url.split(sheetKey)[0];
    return `${baseUrl}${sheetKey}/export?format=xlsx`;
  }
  return "ERROR";
}
//_________________
export const hanldeGenerateData = (arr, lengthHeader) => {
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
      // Sau khi đã có finalHeader -> tìm finalData
      row.forEach((item, index) => {
        objData[`${finalHeader[index]}`.replace(/ /g, "")] = item;
      });
      finalData.push(objData);
    }
  });

  return finalData;
};

export const flatMergedCellExcel = (workbook) => {
  const sheetNames = workbook.SheetNames;
  let roa;
  let headerLength = 1;

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
    // Option 2, calc Header length by BgColor
    // Destructing to get obj not has !ref, !cols, !merges
    const {
      "!ref": ref,
      "!cols": col,
      "!merges": merges,
      ...rest
    } = workbook.Sheets[worksheet];

    // Object.entries(rest).forEach((item) => {
    //   // item[0] is key item[1] is obj
    //   // Check if style has props bgColor
    //   const hasKey = Object.keys(item[1]?.s).includes("bgColor");
    //   // Example: A1223 -> row is 1223
    //   const row = Number(item[0].match(/\d+/g)[0]);
    //   if (hasKey && row > headerLength) {
    //     headerLength = row;
    //   }
    // });

    roa = XLSX.utils.sheet_to_json(workbook.Sheets[worksheet], {
      header: 1,
    });
  }
  return { roa, headerLength };
};
