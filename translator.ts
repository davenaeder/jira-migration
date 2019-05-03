import * as XLSX from "xlsx";
import * as fs from "fs";
import {
    ISSUE_SHEETS,
    USER_TABLE,
    USER_COLS,
    WANTED_COLS,
    PROJ_OLD_KEY,
    PROJ_NEW_KEY,
} from "./const";

const WIDTH = 702;

const generateColumnAddresses = () => {
    const addrs: string[] = [];
    for (let i = 0; i < WIDTH; i++) {
        const code1 = Math.floor(i / 26) === 0 ? 0 : Math.floor(i / 26) + 64;
        const code2 = (i % 26) + 65;
        // const letter1 = String.fromCharCode(Math.floor(i / 26) + 65);
        const value = [code1, code2]
            .filter(Boolean)
            .map(n => String.fromCharCode(n))
            .join("");
        addrs.push(value);
    }
    return addrs;
};

const intToExcelCol = (number: number) => {
    let colName = "",
        dividend = Math.floor(Math.abs(number)),
        rest;

    while (dividend > 0) {
        rest = Math.floor((dividend - 1) % 26);
        colName = String.fromCharCode(65 + rest) + colName;
        dividend = Math.floor((dividend - rest) / 26);
    }
    return colName;
};

const excelColToInt = (colName: string) => {
    let digits = colName.toUpperCase().split(""),
        number = 0;

    for (var i = 0; i < digits.length; i++) {
        number +=
            (digits[i].charCodeAt(0) - 64) *
            Math.pow(26, digits.length - i - 1);
    }

    return number;
};

interface TablePointer {
    x: number;
    y: number;
    c: string;
    p?: string;
}

const analyzeRange = (colonRange: string) => {
    const [left, right] = colonRange.split(":");
    const [a, leftCol, leftRow] = RE_ADDR.exec(left);
    const [b, rightCol, rightRow] = RE_ADDR.exec(right);
    return {
        start: {
            x: excelColToInt(leftCol),
            y: parseInt(leftRow),
            c: leftCol,
        } as TablePointer,
        end: {
            x: excelColToInt(rightCol),
            y: parseInt(rightRow),
            c: rightCol,
        } as TablePointer,
    };
};

const visitCells = (
    sheet: XLSX.WorkSheet,
    visitor: (cell: XLSX.CellObject, addr: TablePointer) => void,
) => {
    const range = analyzeRange(sheet["!ref"]);
    for (let r = range.start.y; r <= range.end.y; r++) {
        for (let c = range.start.x; c <= range.end.x; c++) {
            const col = intToExcelCol(c);
            const loc = `${col}${r}`;
            // console.log(loc);
            visitor(sheet[loc], { x: c, y: r, c: col, p: loc });
        }
    }
};

const COLS = generateColumnAddresses();
const RE_ADDR = /([A-Z]+)([0-9]+)/;
const RE_PROJ_KEY = new RegExp(`${PROJ_OLD_KEY}\-([0-9]+)`);

function main() {
    var workbook = XLSX.readFile("issues.xlsx");

    const getColumnNames = (sheet: XLSX.WorkSheet) =>
        COLS.map(addr => [addr, sheet[`${addr}1`]])
            .filter(([addr, value]) => Boolean(value))
            .map(([addr, value]) => [addr, value.w] as [string, string]);

    ISSUE_SHEETS.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const names = getColumnNames(sheet);
        const newSheet: string[][] = [];
        const allCells = Object.keys(sheet).filter(k => k[0] !== "!");
        // Filter all columns to the ones in our whitelist.
        const columnsWeWant = names.filter(([addr, n]) =>
            WANTED_COLS.includes(n),
        );

        // Get the column addresses from the know username columns
        const usernameColAddrs = columnsWeWant
            .filter(([addr, column]) => USER_COLS.includes(column))
            .reduce((p, c) => [...p, c[0]], [] as string[]);

        // Quick function for generating a new row
        const newRow = (id: number) => {
            if (newSheet[id - 1]) {
                return newSheet[id - 1];
            } else {
                const newRow = new Array<string>();
                newSheet.push(newRow);
                return newRow;
            }
        };

        // Create the first row, which is the header row.
        const row = newRow(1);
        columnsWeWant.forEach(([addr, name], i) => (row[i] = name));

        console.log("Sheet", sheetName);

        visitCells(sheet, (cell, addr) => {
            if (addr.y <= 1) {
                // DO not look at header row for values. We've already done this.
                return;
            }
            let value: string = undefined;

            const isWanted = columnsWeWant.find(([a, name]) => a === addr.c);
            if (isWanted) {
                if (cell) {
                    value = cell.w;
                    value = value.replace(RE_PROJ_KEY, `${PROJ_NEW_KEY}-$1`);

                    const [_, colName] = isWanted;
                    if (usernameColAddrs.indexOf(addr.c) >= 0) {
                        // NOTE: good spot to push users into a set if you're looking to generate a list
                        const translatedUser = USER_TABLE[value.toLowerCase()];
                        value = translatedUser ? translatedUser : value;
                    } else if (colName === "Comment") {
                        const [
                            commentDate,
                            commentUser,
                            commentText,
                        ] = value.split(";");
                        const translatedUser =
                            USER_TABLE[commentUser.toLowerCase()];
                        if (translatedUser) {
                            const newCommentText = [
                                commentDate,
                                translatedUser,
                                commentText,
                            ].join(";");
                            value = newCommentText;
                        }
                    }
                }
                console.log(addr.p, value);
                newRow(addr.y).push(value);
            }
        });

        const newWorkbook = XLSX.utils.book_new();
        // console.log(newSheet);
        const newXLSheet = XLSX.utils.aoa_to_sheet(newSheet);
        fs.writeFileSync(`translated-${sheetName}.csv`, XLSX.utils.sheet_to_csv(newXLSheet));
        // XLSX.utils.book_append_sheet(newWorkbook, newXLSheet);
        // XLSX.writeFile(newWorkbook, "issues-translated.xlsx");
    });
}

main();
