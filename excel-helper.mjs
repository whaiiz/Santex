import fs from 'fs';
import XLSX from 'xlsx';
import XlsxPopulate from 'xlsx-populate';

const { fromDataAsync, fromBlankAsync } = XlsxPopulate;
const { readFileSync } = fs;

export const createExcel = async () => {
    const workbook = await fromBlankAsync();
    return workbook;
}

export const createMerges = (merges, worksheet) => {
    for (const { startAddress, endAddress } of merges) {
        const mergeRange = `${startAddress}:${endAddress}`;
        worksheet.range(mergeRange).merged(true);
    }
}

export const createWorksheet = (config) => {
    const { workbook, merges, cells, worksheetName } = config;
    const worksheet = workbook.addSheet(worksheetName);

    createMerges(merges, worksheet)
    createCells(cells, worksheet);
}

export const createCells = (cells, worksheet) => {
    for (const { value, address } of cells) { 
        const cell = worksheet.cell(address)
        cell.value(value);

        cell.style('horizontalAlignment', 'center');
        cell.style('verticalAlignment', 'center');
    }
}

export const getWorksheet = async (fileName, workSheetName) => {
    const fileData = readFileSync(fileName);
    const workbook = await fromDataAsync(fileData);
    const worksheet = workbook.sheet(workSheetName);

    return worksheet ? worksheet: workbook.sheets[0];
}

export const getCell = (address, worksheet) => {
    return worksheet.cell(address).value();
}

export const deleteWorksheet = (workbook, sheetName) => {
    workbook.deleteSheet(sheetName);
}

export const xlsToXlsx = (fileName) => {
    const workbookXLS = XLSX.readFile(fileName);
    const newFileName = fileName.slice(0, fileName.lastIndexOf('.')) + '.xlsx'

    XLSX.writeFile(workbookXLS, `${newFileName}`);
    return newFileName;
}

export const countFilledRows = (worksheet) => {
    const usedRange = worksheet.usedRange();
    return usedRange._numRows;
}

export const saveExcel = async (workbook, fileName) => {
    return await workbook.toFileAsync(fileName);
}