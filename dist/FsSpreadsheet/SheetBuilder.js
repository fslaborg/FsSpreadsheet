import { disposeSafe, getEnumerator, curry2, defaultOf } from "../fable_modules/fable-library.4.1.3/Util.js";
import { addToDict, tryGetValue } from "../fable_modules/fable-library.4.1.3/MapUtil.js";
import { Record, FSharpRef } from "../fable_modules/fable-library.4.1.3/Types.js";
import { toNullable, some } from "../fable_modules/fable-library.4.1.3/Option.js";
import { FsCell, FsCell_$reflection } from "./Cells/FsCell.js";
import { record_type, string_type, bool_type, option_type, float64_type, list_type, lambda_type } from "../fable_modules/fable-library.4.1.3/Reflection.js";
import { iterate, isEmpty, map as map_1, concat, singleton, append, empty as empty_1 } from "../fable_modules/fable-library.4.1.3/List.js";
import { newGuid } from "../fable_modules/fable-library.4.1.3/Guid.js";
import { FsRangeAddress_$ctor_7E77A4A0, FsRangeAddress__get_FirstAddress } from "./Ranges/FsRangeAddress.js";
import { FsRangeBase__get_RangeAddress } from "./Ranges/FsRangeBase.js";
import { FsAddress_$ctor_Z37302880, FsAddress__get_RowNumber } from "./FsAddress.js";
import { find, indexed } from "../fable_modules/fable-library.4.1.3/Seq.js";
import { FsRangeColumn__Cell_Z4232C216 } from "./Ranges/FsRangeColumn.js";
import { FsTableField__get_Column } from "./Tables/FsTableField.js";
import { FsWorksheet } from "./FsWorksheet.js";
import { FsWorkbook } from "./FsWorkbook.js";

export function Dictionary_tryGetValue(k, dict) {
    let patternInput;
    let outArg = defaultOf();
    patternInput = [tryGetValue(dict, k, new FSharpRef(() => outArg, (v) => {
        outArg = v;
    })), outArg];
    if (patternInput[0]) {
        return some(patternInput[1]);
    }
    else {
        return void 0;
    }
}

export function Dictionary_length(dict) {
    return dict.size;
}

export class FieldMap$1 extends Record {
    constructor(CellTransformers, HeaderTransformers, ColumnWidth, RowHeight, AdjustToContents, Hash) {
        super();
        this.CellTransformers = CellTransformers;
        this.HeaderTransformers = HeaderTransformers;
        this.ColumnWidth = ColumnWidth;
        this.RowHeight = RowHeight;
        this.AdjustToContents = AdjustToContents;
        this.Hash = Hash;
    }
}

export function FieldMap$1_$reflection(gen0) {
    return record_type("FsSpreadsheet.SheetBuilder.FieldMap`1", [gen0], FieldMap$1, () => [["CellTransformers", list_type(lambda_type(gen0, lambda_type(FsCell_$reflection(), FsCell_$reflection())))], ["HeaderTransformers", list_type(lambda_type(gen0, lambda_type(FsCell_$reflection(), FsCell_$reflection())))], ["ColumnWidth", option_type(float64_type)], ["RowHeight", option_type(lambda_type(gen0, option_type(float64_type)))], ["AdjustToContents", bool_type], ["Hash", string_type]]);
}

export function FieldMap$1_empty() {
    let copyOfStruct;
    return new FieldMap$1(empty_1(), empty_1(), void 0, void 0, false, (copyOfStruct = newGuid(), copyOfStruct));
}

export function FieldMap$1_create_Z3BCCB7EB(mapRow) {
    const empty = FieldMap$1_empty();
    return new FieldMap$1(append(empty.CellTransformers, singleton(curry2(mapRow))), empty.HeaderTransformers, empty.ColumnWidth, empty.RowHeight, empty.AdjustToContents, empty.Hash);
}

export function FieldMap$1__header_Z721C83C5(self, name) {
    return new FieldMap$1(self.CellTransformers, append(self.HeaderTransformers, singleton((_arg) => ((cell) => {
        cell.SetValueAs(name);
        return cell;
    }))), self.ColumnWidth, self.RowHeight, self.AdjustToContents, self.Hash);
}

export function FieldMap$1__header_54F19B58(self, mapHeader) {
    return new FieldMap$1(self.CellTransformers, append(self.HeaderTransformers, singleton((value) => ((cell) => {
        cell.SetValueAs(mapHeader(value));
        return cell;
    }))), self.ColumnWidth, self.RowHeight, self.AdjustToContents, self.Hash);
}

export function FieldMap$1__adjustToContents(self) {
    return new FieldMap$1(self.CellTransformers, self.HeaderTransformers, self.ColumnWidth, self.RowHeight, true, self.Hash);
}

export function FieldMap$1_field_Z5C94BCA1(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(map(row));
        return cell;
    });
}

export function FieldMap$1_field_54F19B58(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(map(row));
        return cell;
    });
}

export function FieldMap$1_field_477D79EC(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(map(row));
        return cell;
    });
}

export function FieldMap$1_field_Z1D55A0D7(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(map(row));
        return cell;
    });
}

export function FieldMap$1_field_298D8758(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(map(row));
        return cell;
    });
}

export function FieldMap$1_field_Z78848E24(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(toNullable(map(row)));
        return cell;
    });
}

export function FieldMap$1_field_4C14CE8F(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(toNullable(map(row)));
        return cell;
    });
}

export function FieldMap$1_field_Z404517D6(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(toNullable(map(row)));
        return cell;
    });
}

export function FieldMap$1_field_Z14464045(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        cell.SetValueAs(toNullable(map(row)));
        return cell;
    });
}

export function FieldMap$1_field_33F53BB(map) {
    return FieldMap$1_create_Z3BCCB7EB((row, cell) => {
        const matchValue = map(row);
        if (matchValue != null) {
            const text = matchValue;
            cell.SetValueAs(text);
            return cell;
        }
        else {
            return cell;
        }
    });
}

export function FsSpreadsheet_FsTable__FsTable_Populate_526F9CF7(self, cells, data, fields) {
    const headersAvailable = !isEmpty(concat(map_1((field) => field.HeaderTransformers, fields)));
    if (headersAvailable && (self.ShowHeaderRow === false)) {
        self.ShowHeaderRow = headersAvailable;
    }
    const startAddress = FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self));
    const startRowIndex = (headersAvailable ? (FsAddress__get_RowNumber(startAddress) + 1) : FsAddress__get_RowNumber(startAddress)) | 0;
    const enumerator = getEnumerator(indexed(data));
    try {
        while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
            const forLoopVar = enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]();
            const row = forLoopVar[1];
            const activeRowIndex = (forLoopVar[0] + startRowIndex) | 0;
            const enumerator_1 = getEnumerator(fields);
            try {
                while (enumerator_1["System.Collections.IEnumerator.MoveNext"]()) {
                    const field_1 = enumerator_1["System.Collections.Generic.IEnumerator`1.get_Current"]();
                    const headerCell = FsCell.createEmpty();
                    const enumerator_2 = getEnumerator(field_1.HeaderTransformers);
                    try {
                        while (enumerator_2["System.Collections.IEnumerator.MoveNext"]()) {
                            enumerator_2["System.Collections.Generic.IEnumerator`1.get_Current"]()(row)(headerCell);
                        }
                    }
                    finally {
                        disposeSafe(enumerator_2);
                    }
                    const headerString = (headerCell.Value === "") ? field_1.Hash : headerCell.Value;
                    const activeCell = FsRangeColumn__Cell_Z4232C216(FsTableField__get_Column(self.Field(headerString, cells)), activeRowIndex, cells);
                    const enumerator_3 = getEnumerator(field_1.CellTransformers);
                    try {
                        while (enumerator_3["System.Collections.IEnumerator.MoveNext"]()) {
                            enumerator_3["System.Collections.Generic.IEnumerator`1.get_Current"]()(row)(activeCell);
                        }
                    }
                    finally {
                        disposeSafe(enumerator_3);
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
}

export function FsSpreadsheet_FsTable__FsTable_populate_Static_Z735E4C44(table, cells, data, fields) {
    FsSpreadsheet_FsTable__FsTable_Populate_526F9CF7(table, cells, data, fields);
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_Populate_Z2A1350BF(self, data, fields) {
    const headersAvailable = !isEmpty(concat(map_1((field) => field.HeaderTransformers, fields)));
    const headers = new Map([]);
    const enumerator = getEnumerator(indexed(data));
    try {
        while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
            const forLoopVar = enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]();
            const row = forLoopVar[1];
            const startRowIndex = (headersAvailable ? 2 : 1) | 0;
            const activeRow = self.Row(forLoopVar[0] + startRowIndex);
            const enumerator_1 = getEnumerator(fields);
            try {
                while (enumerator_1["System.Collections.IEnumerator.MoveNext"]()) {
                    const field_1 = enumerator_1["System.Collections.Generic.IEnumerator`1.get_Current"]();
                    const headerCell = FsCell.createEmpty();
                    const enumerator_2 = getEnumerator(field_1.HeaderTransformers);
                    try {
                        while (enumerator_2["System.Collections.IEnumerator.MoveNext"]()) {
                            enumerator_2["System.Collections.Generic.IEnumerator`1.get_Current"]()(row)(headerCell);
                        }
                    }
                    finally {
                        disposeSafe(enumerator_2);
                    }
                    let index;
                    const patternInput = (headerCell.Value === "") ? [false, field_1.Hash] : [true, headerCell.Value];
                    const headerString = patternInput[1];
                    const matchValue = Dictionary_tryGetValue(headerString, headers);
                    if (matchValue == null) {
                        const i = (headers.size + 1) | 0;
                        addToDict(headers, headerString, i);
                        if (patternInput[0]) {
                            const value = self.Row(1).Item(i).CopyFrom(headerCell);
                        }
                        index = i;
                    }
                    else {
                        index = matchValue;
                    }
                    const activeCell = activeRow.Item(index);
                    const enumerator_3 = getEnumerator(field_1.CellTransformers);
                    try {
                        while (enumerator_3["System.Collections.IEnumerator.MoveNext"]()) {
                            enumerator_3["System.Collections.Generic.IEnumerator`1.get_Current"]()(row)(activeCell);
                        }
                    }
                    finally {
                        disposeSafe(enumerator_3);
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
    self.SortRows();
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_populate_Static_Z2578A106(sheet, data, fields) {
    FsSpreadsheet_FsWorksheet__FsWorksheet_Populate_Z2A1350BF(sheet, data, fields);
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_createFrom_Static_Z2DEFA746(name, data, fields) {
    const sheet = new FsWorksheet(name);
    FsSpreadsheet_FsWorksheet__FsWorksheet_populate_Static_Z2578A106(sheet, data, fields);
    return sheet;
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_createFrom_Static_Z2A1350BF(data, fields) {
    return FsSpreadsheet_FsWorksheet__FsWorksheet_createFrom_Static_Z2DEFA746("Sheet1", data, fields);
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_PopulateTable_59F260F9(self, tableName, startAddress, data, fields) {
    const headersAvailable = !isEmpty(concat(map_1((field) => field.HeaderTransformers, fields)));
    FsSpreadsheet_FsTable__FsTable_Populate_526F9CF7(self.Table(tableName, FsRangeAddress_$ctor_7E77A4A0(startAddress, startAddress), headersAvailable), self.CellCollection, data, fields);
    self.SortRows();
}

export function FsSpreadsheet_FsWorksheet__FsWorksheet_createTableFrom_Static_30234BE1(name, tableName, data, fields) {
    const sheet = new FsWorksheet(name);
    FsSpreadsheet_FsWorksheet__FsWorksheet_PopulateTable_59F260F9(sheet, tableName, FsAddress_$ctor_Z37302880(1, 1), data, fields);
    return sheet;
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_Populate_Z2DEFA746(self, name, data, fields) {
    self.InitWorksheet(name);
    FsSpreadsheet_FsWorksheet__FsWorksheet_populate_Static_Z2578A106(find((s) => (s.Name === name), self.GetWorksheets()), data, fields);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_populate_Static_Z7163EB39(workbook, name, data, fields) {
    FsSpreadsheet_FsWorkbook__FsWorkbook_Populate_Z2DEFA746(workbook, name, data, fields);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_createFrom_Static_Z2DEFA746(name, data, fields) {
    const workbook = new FsWorkbook();
    FsSpreadsheet_FsWorkbook__FsWorkbook_populate_Static_Z7163EB39(workbook, name, data, fields);
    return workbook;
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_createFrom_Static_Z2A1350BF(data, fields) {
    return FsSpreadsheet_FsWorkbook__FsWorkbook_createFrom_Static_Z2DEFA746("Sheet1", data, fields);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_createFrom_Static_Z2E36FAE7(sheets) {
    const workbook = new FsWorkbook();
    iterate((sheet) => {
        const value = workbook.AddWorksheet(sheet);
    }, sheets);
    return workbook;
}

