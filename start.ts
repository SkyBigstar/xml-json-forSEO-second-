import * as XLSX from 'xlsx';
import { writeFileSync } from 'fs';

const workbook: XLSX.WorkBook = XLSX.readFile('data/1.xlsx');

const worksheet_roadmap: XLSX.WorkSheet = workbook.Sheets[workbook.SheetNames[4]];

// import * as ExcelJS from 'exceljs';

// const _workbook = new ExcelJS.Workbook();
// const myfunc = async () => {
//     await _workbook.xlsx.readFile('data/1.xlsx');
//     const worksheet = _workbook.getWorksheet(workbook.SheetNames[0]);
//     const cell = worksheet.getCell('A17');
    
//     const style = cell.style;
//     console.log(style);
// }
// myfunc();

const data: any[] = XLSX.utils.sheet_to_json(worksheet_roadmap, { header: 1 });

interface Result {
    Recommendation: string
    SpearmanCorrelation: number
    PearsonCorrelation: number
    BestofBothCorrelation: number
    FactorID: string
    FactorName: string
    BestofBothByPage: number
    SharedBoB: number
    SharedBoBbyPage: number
    SharedPage1Avg: number
    Shared1Max: number
    SharedPcent: number // is green?
    Page1Avg: number // is red?
    Result1: number
    Result2: number
    Result3: number
    Result4: number
    Deficit: number
    Goal: number
    OverallMax: number
    Usage: number
    Class: string
}

interface Phase {
    title: string;
    results: Result[];
}

interface Roadmap_data {
    phases: Phase[];
}

const ret_data: Roadmap_data = {phases: []}; // the return value of this method

let cnt_phases = 0;

// Loop through each row of data
for (let i = 0; i < data.length; i++) {
    const rowData = data[i];
    if(!rowData[0]) continue;
    // When meet the first line for Phase paragraph
    if (rowData[0].startsWith("Recommendation")) {
        const _data : Phase = { title: rowData[0], results: [] }
        ret_data.phases.push(_data);
        cnt_phases ++;
        continue;
    }

    if(!cnt_phases) continue;
    
    // Get the cell address
    // const cellAddress = XLSX.utils.encode_cell({ r: i, c: 1 })
    // Get the cell style
    // const cellStyle = worksheet_roadmap[cellAddress]?.s;
    // If the cell has a style, get the font color
    // let _strongFactor: boolean = false;
    // let _strongFactorExcluding: boolean = false;

    // if (cellStyle !== undefined && cellStyle !== null && cellStyle.font !== undefined && cellStyle.font !== null) {
    //     const fontColor = cellStyle.font.color;
    //     // If the font color is not black, add it to the cell data
    //     if (fontColor && fontColor.rgb && (fontColor.rgb === 'ff0000' || fontColor.rgb === 'FF0000')) {
    //         _strongFactor = true;
    //     }
    //     if (fontColor && fontColor.rgb && (fontColor.rgb === '00ff00' || fontColor.rgb === '00FF00')) {
    //         _strongFactorExcluding = true;
    //     }
    // }
    
    const _result: Result = {
        Recommendation: rowData[0],
        SpearmanCorrelation: parseInt(rowData[1]),
        PearsonCorrelation: parseInt(rowData[2]),
        BestofBothCorrelation: parseInt(rowData[3]),
        FactorID: rowData[4],
        FactorName: rowData[5],
        BestofBothByPage: rowData[6],
        SharedBoB: parseInt(rowData[7]),
        SharedBoBbyPage: parseInt(rowData[8]),
        SharedPage1Avg: parseInt(rowData[9]),
        Shared1Max: parseInt(rowData[10]),
        SharedPcent: parseInt(rowData[11]),
        Page1Avg: parseInt(rowData[12]),
        Result1: parseInt(rowData[13]),
        Result2: parseInt(rowData[14]),
        Result3: parseInt(rowData[15]),
        Result4: parseInt(rowData[16]),
        Deficit: parseInt(rowData[17]),
        Goal: parseInt(rowData[18]),
        OverallMax: parseInt(rowData[19]),
        Usage: parseInt(rowData[20]),
        Class: rowData[21],
    }
    ret_data.phases[cnt_phases - 1].results.push(_result);
}
writeFileSync("1.txt",JSON.stringify(ret_data),{
    flag: 'w',
});
