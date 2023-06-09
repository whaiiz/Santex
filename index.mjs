import { createExcel, 
         getWorksheet, 
         countFilledRows, 
         getCell,
         xlsToXlsx, 
         createWorksheet,
         saveExcel,
         deleteWorksheet} 
         from './excel-helper.mjs';
import { parseDate, monthNumberToString } from './date-helper.mjs';
import { groupBy } from './generic-helper.mjs';
import fs from 'fs';

let inputFile = '';
let outputFile = '';

const readSpents = async () => {
    const worksheet = await getWorksheet(xlsToXlsx(inputFile), "Saldos e Movimentos");
    const rowsCount = countFilledRows(worksheet);
    let result = [];

    for (let row = 8; row < rowsCount; row++) {
        result.push({
            date: parseDate(getCell(`A${row}`, worksheet)),
            description: getCell(`C${row}`, worksheet),
            cost: getCell(`D${row}`, worksheet)
        });
    }

    result = groupBy(result, column => column.date.getFullYear());

    Object.keys(result).forEach(key => {
        result[key] = groupBy(result[key], column => column.date.getMonth()); 
    })

    return result;
}

const getTotals = (spents, initialPosition, finalPosition) => {
    const loses = spents.reduce((accumulator, { cost }) => cost < 0 ? accumulator + cost : accumulator + 0, 0);
    const gains = spents.reduce((accumulator, { cost }) => cost > 0 ? accumulator + cost : accumulator + 0, 0);
    const totalsMerges = ['D','E','F'].map(p => { 
        return { startAddress: p + initialPosition, endAddress: p + finalPosition }
    });
    const totalsCells = [
        { value: loses, address: 'D' + initialPosition },
        { value: gains, address: 'E' + initialPosition }, 
        { value: gains + loses, address: 'F' + initialPosition }
    ];

    return { totalsMerges, totalsCells }
}

const getCellsForSpents = (spents, initialPosition) => {
    const finalPosition = initialPosition + spents.length - 1;
    const { totalsCells, totalsMerges } = getTotals(spents, initialPosition, finalPosition)    
    const cellSpents = spents.flatMap(spent => {
        const { description, cost } = spent;
        const cells = [
            { value: description, address: 'B' + initialPosition },
            { value: cost, address: 'C' + initialPosition }
        ]

        initialPosition += 1;
        
        return cells;             
    });

    return { cellSpents: cellSpents.concat(totalsCells), 
             spentMerges: totalsMerges.concat() } ;
}

const getCellsForMonthlySpents = (monthlySpents, initialPosition = 1) => {
    const monthCells = [];
    const monthMerges = [];

    for (const [month, spents] of Object.entries(monthlySpents)) {
        const finalPosition = initialPosition + Object.keys(spents).length - 1;
        const { cellSpents, spentMerges } = getCellsForSpents(spents, initialPosition);

        monthCells.push({value: monthNumberToString(month), address: `A${initialPosition}`}, ...cellSpents);
        monthMerges.push({ startAddress: `A${initialPosition}`, endAddress: `A${finalPosition}`}, ...spentMerges);

        initialPosition = finalPosition + 1;
    }

    return { monthCells, monthMerges }
}

const populateExcel = (data, workbook) => {
    for (const [year, monthlySpents] of Object.entries(data)) {
        let { monthCells, monthMerges } = getCellsForMonthlySpents(monthlySpents)
        createWorksheet({ cells: monthCells, merges: monthMerges, workbook, worksheetName: year });
    }

    deleteWorksheet(workbook, "Sheet1");
}

const init = async () => {
    const workbook = await createExcel();
    const spents = await readSpents();

    populateExcel(spents, workbook);
    saveExcel(workbook, outputFile);
}

if (process.argv.length < 4 || !fs.existsSync(process.argv[2])) {
    console.error("node <path> <input-file> <output-file>")
} else {
    inputFile = process.argv[2];
    outputFile = process.argv[3];
    init();
}