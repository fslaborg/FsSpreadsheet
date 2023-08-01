import { map } from "./fable_modules/fable-library.4.1.3/Seq.js";
import { disposeSafe, getEnumerator } from "./fable_modules/fable-library.4.1.3/Util.js";
import { FsAddress_$ctor_Z721C83C5, FsAddress__get_Address } from "./FsSpreadsheet/FsAddress.js";
import { value as value_6, some } from "./fable_modules/fable-library.4.1.3/Option.js";
import { printf, toText } from "./fable_modules/fable-library.4.1.3/String.js";
import { fromFsTable } from "./Table.js";
import { FsWorksheet } from "./FsSpreadsheet/FsWorksheet.js";
import { DataType, FsCell } from "./FsSpreadsheet/Cells/FsCell.js";
import { toString } from "./fable_modules/fable-library.4.1.3/Types.js";
import { parse } from "./fable_modules/fable-library.4.1.3/Double.js";
import { parse as parse_1 } from "./fable_modules/fable-library.4.1.3/Date.js";
import { parse as parse_2 } from "./fable_modules/fable-library.4.1.3/Boolean.js";
import { FsRangeAddress_$ctor_Z721C83C5 } from "./FsSpreadsheet/Ranges/FsRangeAddress.js";
import { FsTable } from "./FsSpreadsheet/Tables/FsTable.js";

export function addFsWorksheet(wb, fsws) {
    fsws.RescanRows();
    const rows = map((x) => x.Cells, fsws.Rows);
    const ws = wb.addWorksheet(fsws.Name);
    const enumerator = getEnumerator(rows);
    try {
        while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
            const enumerator_1 = getEnumerator(enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]());
            try {
                while (enumerator_1["System.Collections.IEnumerator.MoveNext"]()) {
                    const cell = enumerator_1["System.Collections.Generic.IEnumerator`1.get_Current"]();
                    const c = ws.getCell(FsAddress__get_Address(cell.Address));
                    const matchValue = cell.DataType;
                    switch (matchValue.tag) {
                        case 1: {
                            c.value = some(cell.ValueAsBool());
                            break;
                        }
                        case 2: {
                            c.value = some(cell.ValueAsFloat());
                            break;
                        }
                        case 3: {
                            c.value = some(cell.ValueAsDateTime());
                            break;
                        }
                        case 0: {
                            c.value = some(cell.Value);
                            break;
                        }
                        default: {
                            const msg = toText(printf("ValueType \'%A\' is not fully implemented in FsSpreadsheet and is handled as string input."))(matchValue);
                            console.log(msg);
                            c.value = some(cell.Value);
                        }
                    }
                }
            }
            finally {
                disposeSafe(enumerator_1);
            }
        }
    }
    finally {
        disposeSafe(enumerator);
    }
    const enumerator_2 = getEnumerator(map((table) => fromFsTable(fsws.CellCollection, table), fsws.Tables));
    try {
        while (enumerator_2["System.Collections.IEnumerator.MoveNext"]()) {
            const table_1 = enumerator_2["System.Collections.Generic.IEnumerator`1.get_Current"]();
            ws.addTable(table_1);
        }
    }
    finally {
        disposeSafe(enumerator_2);
    }
}

export function addJsWorksheet(wb, jsws) {
    const fsws = new FsWorksheet(jsws.name);
    jsws.eachRow(function(row, rowNumber){((tupledArg) => {
        tupledArg[0].eachCell(function(cell, rowIndex){((tupledArg_1) => {
            const c = tupledArg_1[0];
            if (c.value != null) {
                const t = (c.type) | 0;
                const fsadress = FsAddress_$ctor_Z721C83C5(c.address);
                const createFscell = (dt, v) => (new FsCell(v, dt, fsadress));
                const vTemp = toString(value_6(c.value));
                let fscell;
                switch (t) {
                    case 2: {
                        const arg_3 = parse(vTemp);
                        fscell = createFscell(new DataType(2, []), arg_3);
                        break;
                    }
                    case 3: {
                        fscell = createFscell(new DataType(0, []), vTemp);
                        break;
                    }
                    case 4: {
                        const arg_5 = parse_1(vTemp);
                        fscell = createFscell(new DataType(3, []), arg_5);
                        break;
                    }
                    case 9: {
                        const arg_1 = parse_2(vTemp);
                        fscell = createFscell(new DataType(1, []), arg_1);
                        break;
                    }
                    default: {
                        const msg = toText(printf("ValueType \'%A\' is not fully implemented in FsSpreadsheet and is handled as string input."))(t);
                        console.log(msg);
                        fscell = createFscell(new DataType(0, []), vTemp);
                    }
                }
                fsws.AddCell(fscell);
            }
        })([cell, rowIndex])});
    })([row, rowNumber])});
    const arr = jsws.getTables();
    for (let idx = 0; idx <= (arr.length - 1); idx++) {
        const jstableref = arr[idx];
        const table = value_6(jstableref.table);
        const tableRef = FsRangeAddress_$ctor_Z721C83C5(table.tableRef);
        const table_1 = new FsTable(table.name, tableRef, table.totalsRow, table.headerRow);
        fsws.AddTable(table_1);
    }
    fsws.RescanRows();
    wb.AddWorksheet(fsws);
}

