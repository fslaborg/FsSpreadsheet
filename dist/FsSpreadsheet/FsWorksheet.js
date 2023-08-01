import { FsCellsCollection__RemoveValueAt_Z37302880, FsCellsCollection__TryRemoveValueAt_Z37302880, FsCellsCollection__RemoveCellAt_Z37302880, FsCellsCollection__Add_2E78CE33, FsCellsCollection__Add_Z21F271A4, FsCellsCollection__Add_Z334DF64D, FsCellsCollection__TryGetCell_Z37302880, FsCellsCollection__Copy, FsCellsCollection__GetCells, FsCellsCollection_$ctor } from "./Cells/FsCellsCollection.js";
import { empty, maxBy, filter, exists, find, tryFind, iterate, last, head, sortBy, map } from "../fable_modules/fable-library.4.1.3/Seq.js";
import { FsColumn } from "./FsColumn.js";
import { clear, safeHash, equals, disposeSafe, getEnumerator, numberHash, comparePrimitives } from "../fable_modules/fable-library.4.1.3/Util.js";
import { FsAddress__get_ColumnNumber, FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "./FsAddress.js";
import { FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "./Ranges/FsRangeAddress.js";
import { groupBy } from "../fable_modules/fable-library.4.1.3/Seq2.js";
import { FsRow } from "./FsRow.js";
import { toConsole, printf, toFail } from "../fable_modules/fable-library.4.1.3/String.js";
import { FsRangeBase__set_RangeAddress_6A2513BC } from "./Ranges/FsRangeBase.js";
import { addRangeInPlace, removeInPlace } from "../fable_modules/fable-library.4.1.3/Array.js";
import { tryFind as tryFind_1, ofSeq } from "../fable_modules/fable-library.4.1.3/Map.js";
import { value as value_3, defaultArg } from "../fable_modules/fable-library.4.1.3/Option.js";
import { FsTable } from "./Tables/FsTable.js";
import { iterate as iterate_1 } from "../fable_modules/fable-library.4.1.3/List.js";
import { FsCell } from "./Cells/FsCell.js";
import { class_type } from "../fable_modules/fable-library.4.1.3/Reflection.js";

export class FsWorksheet {
    constructor(name, fsRows, fsTables, fsCellsCollection) {
        this.name = name;
        this._name = this.name;
        this._rows = defaultArg(fsRows, []);
        this._tables = defaultArg(fsTables, []);
        this._cells = defaultArg(fsCellsCollection, FsCellsCollection_$ctor());
    }
    static init(name) {
        return new FsWorksheet(name, [], [], FsCellsCollection_$ctor());
    }
    get Name() {
        const self = this;
        return self._name;
    }
    set Name(name) {
        const self = this;
        self._name = name;
    }
    get CellCollection() {
        const self = this;
        return self._cells;
    }
    get Tables() {
        const self = this;
        return self._tables;
    }
    get Rows() {
        const self = this;
        return self._rows;
    }
    get Columns() {
        const self = this;
        return map((tupledArg) => {
            let tupledArg_1, cells_1, c_2, c_3;
            const columnIndex = tupledArg[0] | 0;
            return new FsColumn((tupledArg_1 = ((cells_1 = sortBy((c_1) => c_1.RowNumber, tupledArg[1], {
                Compare: comparePrimitives,
            }), [FsAddress_$ctor_Z37302880((c_2 = head(cells_1), c_2.RowNumber), columnIndex), FsAddress_$ctor_Z37302880((c_3 = last(cells_1), c_3.RowNumber), columnIndex)])), FsRangeAddress_$ctor_7E77A4A0(tupledArg_1[0], tupledArg_1[1])), self.CellCollection);
        }, groupBy((c) => c.ColumnNumber, FsCellsCollection__GetCells(self._cells), {
            Equals: (x, y) => (x === y),
            GetHashCode: numberHash,
        }));
    }
    Copy() {
        let n, n_1;
        const self = this;
        const fcc = FsCellsCollection__Copy(self.CellCollection);
        return new FsWorksheet(self.Name, (n = [], (iterate((r) => {
            const arg = r.Copy();
            void (n.push(arg));
        }, self.Rows), n)), (n_1 = [], (iterate((t) => {
            const arg_1 = t.Copy();
            void (n_1.push(arg_1));
        }, self.Tables), n_1)), fcc);
    }
    static copy(sheet) {
        return sheet.Copy();
    }
    Row(rowIndex) {
        const self = this;
        const matchValue = tryFind((row) => (row.Index === rowIndex), self._rows);
        if (matchValue == null) {
            const row_2 = FsRow.createAt(rowIndex, self.CellCollection);
            void (self._rows.push(row_2));
            return row_2;
        }
        else {
            return matchValue;
        }
    }
    RowWithRange(rangeAddress) {
        const self = this;
        if (FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(rangeAddress)) !== FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(rangeAddress))) {
            toFail(printf("Row may not have a range address spanning over different row indices"));
        }
        FsRangeBase__set_RangeAddress_6A2513BC(self.Row(FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(rangeAddress))), rangeAddress);
    }
    static appendRow(row, sheet) {
        sheet.Row(row.Index);
        return sheet;
    }
    static getRows(sheet) {
        return sheet.Rows;
    }
    static getRowAt(rowIndex, sheet) {
        return find((arg_1) => (rowIndex === FsRow.getIndex(arg_1)), FsWorksheet.getRows(sheet));
    }
    static tryGetRowAt(rowIndex, sheet) {
        return tryFind((arg_1) => (rowIndex === FsRow.getIndex(arg_1)), sheet.Rows);
    }
    static tryGetRowAfter(rowIndex, sheet) {
        return tryFind((r) => (r.Index >= rowIndex), sheet.Rows);
    }
    InsertBefore(row, refRow) {
        const self = this;
        let enumerator = getEnumerator(self._rows);
        try {
            while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
                const row_1 = enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]();
                if (row_1.Index >= refRow.Index) {
                    row_1.Index = ((row_1.Index + 1) | 0);
                }
            }
        }
        finally {
            disposeSafe(enumerator);
        }
        self.Row(row.Index);
        return self;
    }
    static insertBefore(row, refRow, sheet) {
        return sheet.InsertBefore(row, refRow);
    }
    ContainsRowAt(rowIndex) {
        const self = this;
        return exists((t) => (t.Index === rowIndex), self.Rows);
    }
    static containsRowAt(rowIndex, sheet) {
        return sheet.ContainsRowAt(rowIndex);
    }
    static countRows(sheet) {
        return sheet.Rows.length;
    }
    RemoveRowAt(rowIndex) {
        const self = this;
        const enumerator = getEnumerator(filter((r) => (r.Index === rowIndex), self._rows));
        try {
            while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
                removeInPlace(enumerator["System.Collections.Generic.IEnumerator`1.get_Current"](), self._rows, {
                    Equals: equals,
                    GetHashCode: safeHash,
                });
            }
        }
        finally {
            disposeSafe(enumerator);
        }
    }
    static removeRowAt(rowIndex, sheet) {
        sheet.RemoveRowAt(rowIndex);
        return sheet;
    }
    TryRemoveAt(rowIndex) {
        const self = this;
        if (self.ContainsRowAt(rowIndex)) {
            self.RemoveRowAt(rowIndex);
        }
        return self;
    }
    static tryRemoveAt(rowIndex, sheet) {
        if (FsWorksheet.containsRowAt(rowIndex, sheet)) {
            sheet.RemoveRowAt(rowIndex);
        }
    }
    SortRows() {
        const self = this;
        const sorted = Array.from(sortBy((r) => r.Index, self._rows, {
            Compare: comparePrimitives,
        }));
        clear(self._rows);
        addRangeInPlace(sorted, self._rows);
    }
    MapRowsInPlace(f) {
        const self = this;
        for (let i = 0; i <= (self._rows.length - 1); i++) {
            const r = self._rows[i];
            self._rows[i] = f(r);
        }
        return self;
    }
    static mapRowsInPlace(f, sheet) {
        return sheet.MapRowsInPlace(f);
    }
    GetMaxRowIndex() {
        const self = this;
        if (self.Rows.length === 0) {
            throw new Error("The FsWorksheet has no FsRows.");
        }
        return maxBy((r) => r.Index, self.Rows, {
            Compare: comparePrimitives,
        });
    }
    static getMaxRowIndex(sheet) {
        return sheet.GetMaxRowIndex();
    }
    GetRowValuesAt(rowIndex) {
        const self = this;
        return self.ContainsRowAt(rowIndex) ? map((c) => c.Value, self.Row(rowIndex).Cells) : empty();
    }
    static getRowValuesAt(rowIndex, sheet) {
        return sheet.GetRowValuesAt(rowIndex);
    }
    TryGetRowValuesAt(rowIndex) {
        const self = this;
        return self.ContainsRowAt(rowIndex) ? self.GetRowValuesAt(rowIndex) : void 0;
    }
    static tryGetRowValuesAt(rowIndex, sheet) {
        return sheet.TryGetRowValuesAt(rowIndex);
    }
    RescanRows() {
        const self = this;
        const rows = ofSeq(map((r) => [r.Index, r], self._rows), {
            Compare: comparePrimitives,
        });
        iterate((tupledArg) => {
            let c_2, c_3;
            const rowIndex = tupledArg[0] | 0;
            let newRange;
            let tupledArg_1;
            const cells_1 = sortBy((c_1) => c_1.ColumnNumber, tupledArg[1], {
                Compare: comparePrimitives,
            });
            tupledArg_1 = [FsAddress_$ctor_Z37302880(rowIndex, (c_2 = head(cells_1), c_2.ColumnNumber)), FsAddress_$ctor_Z37302880(rowIndex, (c_3 = last(cells_1), c_3.ColumnNumber))];
            newRange = FsRangeAddress_$ctor_7E77A4A0(tupledArg_1[0], tupledArg_1[1]);
            const matchValue = tryFind_1(rowIndex, rows);
            if (matchValue == null) {
                self.RowWithRange(newRange);
            }
            else {
                FsRangeBase__set_RangeAddress_6A2513BC(matchValue, newRange);
            }
        }, groupBy((c) => c.RowNumber, FsCellsCollection__GetCells(self._cells), {
            Equals: (x_1, y_1) => (x_1 === y_1),
            GetHashCode: numberHash,
        }));
    }
    Column(columnIndex) {
        const self = this;
        return FsColumn.createAt(columnIndex, self.CellCollection);
    }
    ColumnWithRange(rangeAddress) {
        const self = this;
        if (FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(rangeAddress)) !== FsAddress__get_ColumnNumber(FsRangeAddress__get_LastAddress(rangeAddress))) {
            toFail(printf("Column may not have a range address spanning over different column indices"));
        }
        FsRangeBase__set_RangeAddress_6A2513BC(self.Column(FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(rangeAddress))), rangeAddress);
    }
    static getColumns(sheet) {
        return sheet.Columns;
    }
    static getColumnAt(columnIndex, sheet) {
        return find((arg_1) => (columnIndex === FsColumn.getIndex(arg_1)), FsWorksheet.getColumns(sheet));
    }
    static tryGetColumnAt(columnIndex, sheet) {
        return tryFind((arg_1) => (columnIndex === FsColumn.getIndex(arg_1)), sheet.Columns);
    }
    Table(tableName, rangeAddress, showHeaderRow) {
        const self = this;
        const showHeaderRow_1 = defaultArg(showHeaderRow, true);
        const matchValue = tryFind((table) => (table.Name === self.name), self._tables);
        if (matchValue == null) {
            const table_2 = new FsTable(tableName, rangeAddress, showHeaderRow_1);
            void (self._tables.push(table_2));
            return table_2;
        }
        else {
            return matchValue;
        }
    }
    static tryGetTableByName(tableName, sheet) {
        return tryFind((t) => (t.Name === tableName), sheet.Tables);
    }
    static getTableByName(tableName, sheet) {
        try {
            return value_3(tryFind((t) => (t.Name === tableName), sheet.Tables));
        }
        catch (matchValue) {
            throw new Error(`FsTable with name ${tableName} is not presen in the FsWorksheet ${sheet.Name}.`);
        }
    }
    AddTable(table) {
        const self = this;
        if (exists((t) => (t.Name === table.Name), self.Tables)) {
            toConsole(`FsTable ${table.Name} could not be appended as an FsTable with this name is already present in the FsWorksheet ${self.Name}.`);
        }
        else {
            void (self._tables.push(table));
        }
        return self;
    }
    static addTable(table, sheet) {
        return sheet.AddTable(table);
    }
    AddTables(tables) {
        const self = this;
        iterate_1((arg_1) => {
            self.AddTable(arg_1);
        }, tables);
        return self;
    }
    static addTables(tables, sheet) {
        return sheet.AddTables(tables);
    }
    TryGetCellAt(rowIndex, colIndex) {
        const self = this;
        return FsCellsCollection__TryGetCell_Z37302880(self.CellCollection, rowIndex, colIndex);
    }
    static tryGetCellAt(rowIndex, colIndex, sheet) {
        return sheet.TryGetCellAt(rowIndex, colIndex);
    }
    GetCellAt(rowIndex, colIndex) {
        const self = this;
        return value_3(self.TryGetCellAt(rowIndex, colIndex));
    }
    static getCellAt(rowIndex, colIndex, sheet) {
        return sheet.GetCellAt(rowIndex, colIndex);
    }
    AddCell(cell) {
        const self = this;
        const value = FsCellsCollection__Add_Z334DF64D(self.CellCollection, cell);
        return self;
    }
    AddCells(cells) {
        const self = this;
        const value = FsCellsCollection__Add_Z21F271A4(self.CellCollection, cells);
        return self;
    }
    InsertValueAt(value, rowIndex, colIndex) {
        const self = this;
        const cell = new FsCell(value);
        FsCellsCollection__Add_2E78CE33(self.CellCollection, rowIndex, colIndex, cell);
    }
    static insertValueAt(value, rowIndex, colIndex, sheet) {
        sheet.InsertValueAt(value, rowIndex, colIndex);
    }
    SetValueAt(value, rowIndex, colIndex) {
        const self = this;
        const matchValue = FsCellsCollection__TryGetCell_Z37302880(self.CellCollection, rowIndex, colIndex);
        if (matchValue == null) {
            const value_2 = FsCellsCollection__Add_2E78CE33(self.CellCollection, rowIndex, colIndex, value);
            return self;
        }
        else {
            const c = matchValue;
            const value_1 = c.SetValueAs(value);
            return self;
        }
    }
    static setValueAt(value, rowIndex, colIndex, sheet) {
        return sheet.SetValueAt(value, rowIndex, colIndex);
    }
    RemoveCellAt(rowIndex, colIndex) {
        const self = this;
        FsCellsCollection__RemoveCellAt_Z37302880(self.CellCollection, rowIndex, colIndex);
        return self;
    }
    static removeCellAt(rowIndex, colIndex, sheet) {
        return sheet.RemoveCellAt(rowIndex, colIndex);
    }
    TryRemoveValueAt(rowIndex, colIndex) {
        const self = this;
        FsCellsCollection__TryRemoveValueAt_Z37302880(self.CellCollection, rowIndex, colIndex);
    }
    static tryRemoveValueAt(rowIndex, colIndex, sheet) {
        sheet.TryRemoveValueAt(rowIndex, colIndex);
    }
    RemoveValueAt(rowIndex, colIndex) {
        const self = this;
        FsCellsCollection__RemoveValueAt_Z37302880(self.CellCollection, rowIndex, colIndex);
    }
    static removeValueAt(rowIndex, colIndex, sheet) {
        sheet.RemoveValueAt(rowIndex, colIndex);
    }
    static addCell(cell, sheet) {
        return sheet.AddCell(cell);
    }
    static addCells(cell, sheet) {
        return sheet.AddCells(cell);
    }
}

export function FsWorksheet_$reflection() {
    return class_type("FsSpreadsheet.FsWorksheet", void 0, FsWorksheet);
}

export function FsWorksheet_$ctor_7FDA5F7A(name, fsRows, fsTables, fsCellsCollection) {
    return new FsWorksheet(name, fsRows, fsTables, fsCellsCollection);
}

