import XLSX from "xlsx";
export declare function getUrlExcelSheet(url: string): string;
export declare const getHeaderRowCount: (workbook: XLSX.WorkBook) => number;
export declare const sheetPublicToJson: (linkSheetPublic: string) => Promise<any>;
