"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.sheetPublicToJson = exports.getHeaderRowCount = exports.getUrlExcelSheet = void 0;
const xlsx_1 = __importDefault(require("xlsx"));
//___________Developed by 2noscript_____________
function getUrlExcelSheet(url) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
        const sheetKey = match[1];
        const baseUrl = url.split(sheetKey)[0];
        return `${baseUrl}${sheetKey}/export?format=xlsx`;
    }
    return "ERROR";
}
exports.getUrlExcelSheet = getUrlExcelSheet;
const getHeaderRowCount = (workbook) => {
    const sheetNameList = workbook.SheetNames;
    const sheet = workbook.Sheets[sheetNameList[0]];
    const range = xlsx_1.default.utils.decode_range(sheet["!ref"]);
    let parsedNumHeaders = 0;
    for (let R = range.s.r; R <= range.e.r; R++) {
        let isParsed = true;
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = { c: C, r: R };
            const cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
            const cell = sheet[cellRef];
            if (cell && cell.t === "n") {
                isParsed = false;
                break;
            }
        }
        if (isParsed) {
            parsedNumHeaders++;
        }
        else {
            break;
        }
    }
    return parsedNumHeaders;
};
exports.getHeaderRowCount = getHeaderRowCount;
const sheetPublicToJson = (linkSheetPublic) => __awaiter(void 0, void 0, void 0, function* () {
    const linkXLSX = getUrlExcelSheet(linkSheetPublic);
    const response = yield fetch(linkXLSX);
    const arrayBuffer = yield response.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = xlsx_1.default.read(data, { type: "array" });
    const headerRowCount = (0, exports.getHeaderRowCount)(workbook);
    const sheetNameList = workbook.SheetNames;
    const sheet = workbook.Sheets[sheetNameList[0]];
    const range = xlsx_1.default.utils.decode_range(sheet["!ref"]);
    const headersArr = [];
    for (let R = range.s.r; R < headerRowCount; R++) {
        const headers = {};
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = { c: C, r: R };
            const cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
            const cell = sheet[cellRef];
            headers[C] = (headers[C] ? headers[C] : "") + (cell ? cell.v : "");
        }
        headersArr.push(headers);
    }
    let header = {};
    for (let i = 0; i < headersArr.length; i++) {
        for (const index in headersArr[i]) {
            const key = Number(index);
            if (i === 0) {
                if (key == 0 || headersArr[i][key] !== "") {
                    header[key] = headersArr[i][key];
                }
                else {
                    let k = 0;
                    while (headersArr[i][key - k] === "") {
                        header[key] = headersArr[i][key - k - 1];
                        k++;
                    }
                }
            }
            else {
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
    const dataJson = [];
    for (let R = 3; R <= range.e.r; R++) {
        const row = {};
        for (let C = range.s.c; C <= range.e.c; C++) {
            const cellAddress = { c: C, r: R };
            const cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
            const cell = sheet[cellRef];
            row[header[C]] = cell ? cell.v : "";
        }
        dataJson.push(row);
    }
    return dataJson;
});
exports.sheetPublicToJson = sheetPublicToJson;
