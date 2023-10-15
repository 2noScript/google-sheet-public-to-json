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
exports.getUrlExcelSheet = void 0;
const xlsx_1 = __importDefault(require("xlsx"));
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
    const headers = {};
    let headerRowCount = 0;
    for (const cellAddress in sheet) {
        if (cellAddress[0] === "!")
            continue;
        const cell = sheet[cellAddress];
        if (cell && cell.v) {
            const header = cell.v.toString().trim();
            if (!headers[cellAddress[0]]) {
                headers[cellAddress[0]] = header;
                headerRowCount++;
            }
            else if (headers[cellAddress[0]] !== header) {
                break;
            }
        }
    }
    return headerRowCount;
};
function sheetPublicToJson(linkSheetPublic) {
    return __awaiter(this, void 0, void 0, function* () {
        const linkXLSX = getUrlExcelSheet(linkSheetPublic);
        const response = yield fetch(linkXLSX);
        const arrayBuffer = yield response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = xlsx_1.default.read(data, { type: "array" });
        const headerRowCount = getHeaderRowCount(workbook);
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
            let i = 0;
            while (i < range.e.c) {
                try {
                    if (headers[i] === "")
                        headers[i] = headers[i - 1];
                }
                catch (_a) { }
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
                const cellRef = xlsx_1.default.utils.encode_cell(cellAddress);
                const cell = sheet[cellRef];
                row[header[C]] = cell ? cell.v : "";
            }
            dataJson.push(row);
        }
        return dataJson;
    });
}
exports.default = sheetPublicToJson;
