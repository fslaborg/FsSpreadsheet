import { choose, transpose, map, iterateIndexed, iterate, ofArrayWithTail, empty, reverse, singleton, tail as tail_25, cons, head, isEmpty } from "../../fable_modules/fable-library.4.1.3/List.js";
import { RowIndex__get_Index, ColumnIndex__get_Index, TableElement__get_IsColumn, SheetElement } from "./Types.js";
import { FsRangeColumn__Cell_Z4232C216 } from "../Ranges/FsRangeColumn.js";
import { FsTableField__get_Column } from "../Tables/FsTableField.js";
import { add, FSharpSet__Contains, ofSeq } from "../../fable_modules/fable-library.4.1.3/Set.js";
import { comparePrimitives } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { max } from "../../fable_modules/fable-library.4.1.3/Seq.js";
import { FsCellsCollection__get_MaxRowNumber } from "../Cells/FsCellsCollection.js";
import { FsRangeAddress_$ctor_7E77A4A0 } from "../Ranges/FsRangeAddress.js";
import { FsAddress_$ctor_Z37302880 } from "../FsAddress.js";
import { toText, printf, toFail } from "../../fable_modules/fable-library.4.1.3/String.js";
import { FsWorkbook } from "../FsWorkbook.js";
import { FsWorksheet } from "../FsWorksheet.js";

export function splitRowsAndColumns(els) {
    const loop = (inRows_mut, inColumns_mut, current_mut, remaining_mut, agg_mut) => {
        loop:
        while (true) {
            const inRows = inRows_mut, inColumns = inColumns_mut, current = current_mut, remaining = remaining_mut, agg = agg_mut;
            if (!isEmpty(remaining)) {
                switch (head(remaining).tag) {
                    case 4:
                        if (inColumns) {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = cons(new SheetElement(4, [head(remaining).fields[0]]), current);
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                        else if (inRows) {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = singleton(new SheetElement(4, [head(remaining).fields[0]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = cons(["Rows", reverse(current)], agg);
                            continue loop;
                        }
                        else {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = singleton(new SheetElement(4, [head(remaining).fields[0]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                    case 3:
                        if (inColumns) {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = cons(new SheetElement(3, [head(remaining).fields[0], head(remaining).fields[1]]), current);
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                        else if (inRows) {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = singleton(new SheetElement(3, [head(remaining).fields[0], head(remaining).fields[1]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = cons(["Rows", reverse(current)], agg);
                            continue loop;
                        }
                        else {
                            inRows_mut = false;
                            inColumns_mut = true;
                            current_mut = singleton(new SheetElement(3, [head(remaining).fields[0], head(remaining).fields[1]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                    case 2:
                        if (inRows) {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = cons(new SheetElement(2, [head(remaining).fields[0]]), current);
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                        else if (inColumns) {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = singleton(new SheetElement(2, [head(remaining).fields[0]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = cons(["Columns", reverse(current)], agg);
                            continue loop;
                        }
                        else {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = singleton(new SheetElement(2, [head(remaining).fields[0]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                    case 1:
                        if (inRows) {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = cons(new SheetElement(1, [head(remaining).fields[0], head(remaining).fields[1]]), current);
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                        else if (inColumns) {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = singleton(new SheetElement(1, [head(remaining).fields[0], head(remaining).fields[1]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = cons(["Columns", reverse(current)], agg);
                            continue loop;
                        }
                        else {
                            inRows_mut = true;
                            inColumns_mut = false;
                            current_mut = singleton(new SheetElement(1, [head(remaining).fields[0], head(remaining).fields[1]]));
                            remaining_mut = tail_25(remaining);
                            agg_mut = agg;
                            continue loop;
                        }
                    case 0:
                        if (inRows) {
                            inRows_mut = false;
                            inColumns_mut = false;
                            current_mut = empty();
                            remaining_mut = tail_25(remaining);
                            agg_mut = ofArrayWithTail([["Table", singleton(new SheetElement(0, [head(remaining).fields[0], head(remaining).fields[1]]))], ["Rows", reverse(current)]], agg);
                            continue loop;
                        }
                        else if (inColumns) {
                            inRows_mut = false;
                            inColumns_mut = false;
                            current_mut = empty();
                            remaining_mut = tail_25(remaining);
                            agg_mut = ofArrayWithTail([["Table", singleton(new SheetElement(0, [head(remaining).fields[0], head(remaining).fields[1]]))], ["Columns", reverse(current)]], agg);
                            continue loop;
                        }
                        else {
                            inRows_mut = false;
                            inColumns_mut = false;
                            current_mut = empty();
                            remaining_mut = tail_25(remaining);
                            agg_mut = cons(["Table", singleton(new SheetElement(0, [head(remaining).fields[0], head(remaining).fields[1]]))], agg);
                            continue loop;
                        }
                    default:
                        throw new Error("Unknown element combination when grouping Sheet elements");
                }
            }
            else if (inRows) {
                return cons(["Rows", reverse(current)], agg);
            }
            else if (inColumns) {
                return cons(["Columns", reverse(current)], agg);
            }
            else {
                return agg;
            }
            break;
        }
    };
    return reverse(loop(false, false, empty(), els, empty()));
}

export function FsSpreadsheet_DSL_Workbook__Workbook_parseTable_Static(cellCollection, table, els) {
    iterate((col_2) => {
        if (!isEmpty(col_2)) {
            const field = table.Field(head(col_2)[1], cellCollection);
            iterateIndexed((i, tupledArg) => {
                const cell_4 = FsRangeColumn__Cell_Z4232C216(FsTableField__get_Column(field), i + 2, cellCollection);
                cell_4.DataType = tupledArg[0];
                cell_4.Value = tupledArg[1];
            }, tail_25(col_2));
        }
        else {
            throw new Error("Empty column");
        }
    }, TableElement__get_IsColumn(head(els)) ? map((col) => {
        if (col.tag === 1) {
            return map((cell) => {
                if (cell.tag === 1) {
                    return cell.fields[0];
                }
                else {
                    throw new Error("Indexed cells not supported in column transformation");
                }
            }, col.fields[0]);
        }
        else {
            throw new Error("Indexed columns not supported in table transformation");
        }
    }, els) : transpose(map((row) => {
        if (row.tag === 0) {
            return map((cell_2) => {
                if (cell_2.tag === 1) {
                    return cell_2.fields[0];
                }
                else {
                    throw new Error("Indexed cells not supported in row transformation");
                }
            }, row.fields[0]);
        }
        else {
            throw new Error("Indexed rows not supported in table transformation");
        }
    }, els)));
}

export function FsSpreadsheet_DSL_Workbook__Workbook_parseRow_Static(cellCollection, row, els) {
    let cellIndexSet = ofSeq(choose((el) => {
        if (el.tag === 0) {
            return ColumnIndex__get_Index(el.fields[0]);
        }
        else {
            return void 0;
        }
    }, els), {
        Compare: comparePrimitives,
    });
    iterate((el_1) => {
        let i_1;
        if (el_1.tag === 1) {
            const cell_1 = row.Item((i_1 = 1, ((() => {
                while (FSharpSet__Contains(cellIndexSet, i_1)) {
                    i_1 = ((i_1 + 1) | 0);
                }
            })(), (cellIndexSet = add(i_1, cellIndexSet), i_1))));
            cell_1.DataType = el_1.fields[0][0];
            cell_1.Value = el_1.fields[0][1];
        }
        else {
            const cell = row.Item(ColumnIndex__get_Index(el_1.fields[0]));
            cell.DataType = el_1.fields[1][0];
            cell.Value = el_1.fields[1][1];
        }
    }, els);
}

export function FsSpreadsheet_DSL_Workbook__Workbook_parseSheet_Static(sheet, els) {
    let rowIndexSet = add(0, ofSeq(choose((el) => {
        if (el.tag === 1) {
            return RowIndex__get_Index(el.fields[0]);
        }
        else {
            return void 0;
        }
    }, els), {
        Compare: comparePrimitives,
    }));
    iterate((_arg) => {
        let matchResult, l, name, tableElements, l_1, s;
        switch (_arg[0]) {
            case "Columns": {
                matchResult = 0;
                l = _arg[1];
                break;
            }
            case "Table": {
                if (!isEmpty(_arg[1])) {
                    if (head(_arg[1]).tag === 0) {
                        if (isEmpty(tail_25(_arg[1]))) {
                            matchResult = 1;
                            name = head(_arg[1]).fields[0];
                            tableElements = head(_arg[1]).fields[1];
                        }
                        else {
                            matchResult = 3;
                            s = _arg[0];
                        }
                    }
                    else {
                        matchResult = 3;
                        s = _arg[0];
                    }
                }
                else {
                    matchResult = 3;
                    s = _arg[0];
                }
                break;
            }
            case "Rows": {
                matchResult = 2;
                l_1 = _arg[1];
                break;
            }
            default: {
                matchResult = 3;
                s = _arg[0];
            }
        }
        switch (matchResult) {
            case 0: {
                const columns = l;
                const baseRowIndex = (1 + max(rowIndexSet, {
                    Compare: comparePrimitives,
                })) | 0;
                let columnIndexSet = ofSeq(choose((col) => {
                    if (col.tag === 3) {
                        return ColumnIndex__get_Index(col.fields[0]);
                    }
                    else {
                        return void 0;
                    }
                }, columns), {
                    Compare: comparePrimitives,
                });
                iterate((col_1) => {
                    let i_3;
                    let patternInput;
                    switch (col_1.tag) {
                        case 3: {
                            patternInput = [ColumnIndex__get_Index(col_1.fields[0]), col_1.fields[1]];
                            break;
                        }
                        case 4: {
                            patternInput = [(i_3 = 1, ((() => {
                                while (FSharpSet__Contains(columnIndexSet, i_3)) {
                                    i_3 = ((i_3 + 1) | 0);
                                }
                            })(), (columnIndexSet = add(i_3, columnIndexSet), i_3))), col_1.fields[0]];
                            break;
                        }
                        default:
                            throw new Error("Expected column elements");
                    }
                    const elements_2 = patternInput[1];
                    const colI = patternInput[0] | 0;
                    let cellIndexSet = ofSeq(choose((el_1) => {
                        if (el_1.tag === 0) {
                            return RowIndex__get_Index(el_1.fields[0]);
                        }
                        else {
                            return void 0;
                        }
                    }, elements_2), {
                        Compare: comparePrimitives,
                    });
                    iterate((el_2) => {
                        let i_6;
                        if (el_2.tag === 1) {
                            const row_1 = sheet.Row((((i_6 = 1, ((() => {
                                while (FSharpSet__Contains(cellIndexSet, i_6)) {
                                    i_6 = ((i_6 + 1) | 0);
                                }
                            })(), (cellIndexSet = add(i_6, cellIndexSet), i_6)))) + baseRowIndex) - 1);
                            rowIndexSet = add(row_1.Index, rowIndexSet);
                            const cell_1 = row_1.Item(colI);
                            cell_1.DataType = el_2.fields[0][0];
                            cell_1.Value = el_2.fields[0][1];
                        }
                        else {
                            const i_7 = el_2.fields[0];
                            const row = sheet.Row((RowIndex__get_Index(i_7) + baseRowIndex) - 1);
                            rowIndexSet = add(RowIndex__get_Index(i_7), rowIndexSet);
                            const cell = row.Item(colI);
                            cell.DataType = el_2.fields[1][0];
                            cell.Value = el_2.fields[1][1];
                        }
                    }, elements_2);
                }, columns);
                break;
            }
            case 1: {
                const maxRow = (FsCellsCollection__get_MaxRowNumber(sheet.CellCollection) + 1) | 0;
                const range = FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(maxRow, 1), FsAddress_$ctor_Z37302880(maxRow, 1));
                const table = sheet.Table(name, range);
                FsSpreadsheet_DSL_Workbook__Workbook_parseTable_Static(sheet.CellCollection, table, tableElements);
                break;
            }
            case 2: {
                iterate((_arg_1) => {
                    let i_1;
                    switch (_arg_1.tag) {
                        case 1: {
                            const row_2 = sheet.Row(RowIndex__get_Index(_arg_1.fields[0]));
                            FsSpreadsheet_DSL_Workbook__Workbook_parseRow_Static(sheet.CellCollection, row_2, _arg_1.fields[1]);
                            break;
                        }
                        case 2: {
                            const row_3 = sheet.Row((i_1 = 1, ((() => {
                                while (FSharpSet__Contains(rowIndexSet, i_1)) {
                                    i_1 = ((i_1 + 1) | 0);
                                }
                            })(), (rowIndexSet = add(i_1, rowIndexSet), i_1))));
                            FsSpreadsheet_DSL_Workbook__Workbook_parseRow_Static(sheet.CellCollection, row_3, _arg_1.fields[0]);
                            break;
                        }
                        default:
                            throw new Error("Expected row elements");
                    }
                }, l_1);
                break;
            }
            case 3: {
                toFail(printf("Invalid sheet element %s"))(s);
                break;
            }
        }
    }, splitRowsAndColumns(els));
}

export function FsSpreadsheet_DSL_Workbook__Workbook_Parse(self) {
    const workbook = new FsWorkbook();
    iterateIndexed((i, wbEl) => {
        let arg;
        if (wbEl.tag === 1) {
            const worksheet_1 = new FsWorksheet(wbEl.fields[0]);
            FsSpreadsheet_DSL_Workbook__Workbook_parseSheet_Static(worksheet_1, wbEl.fields[1]);
            const value_1 = workbook.AddWorksheet(worksheet_1);
        }
        else {
            const worksheet = new FsWorksheet((arg = ((i + 1) | 0), toText(printf("Sheet%i"))(arg)));
            FsSpreadsheet_DSL_Workbook__Workbook_parseSheet_Static(worksheet, wbEl.fields[0]);
            const value = workbook.AddWorksheet(worksheet);
        }
    }, self.fields[0]);
    return workbook;
}

