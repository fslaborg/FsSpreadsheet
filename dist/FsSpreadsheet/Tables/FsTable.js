import { disposeSafe, getEnumerator, comparePrimitives, int32ToString, defaultOf, equals } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { FsRangeBase_$reflection, FsRangeBase, FsRangeBase__set_RangeAddress_6A2513BC, FsRangeBase__ColumnCount, FsRangeBase__get_RangeAddress } from "../Ranges/FsRangeBase.js";
import { exists, toList, choose, tryPick, tryFind, find, max, length, iterate, iterateIndexed, map, reduce, singleton, collect, delay, initialize } from "../../fable_modules/fable-library.4.1.3/Seq.js";
import { FsRange_$ctor_6A2513BC, FsRange__FirstRow } from "../Ranges/FsRange.js";
import { FsColumn } from "../FsColumn.js";
import { FsRangeAddress__Copy, FsRangeAddress__Union_6A2513BC, FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "../Ranges/FsRangeAddress.js";
import { FsAddress__get_ColumnNumber, FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "../FsAddress.js";
import { rangeDouble } from "../../fable_modules/fable-library.4.1.3/Range.js";
import { FsTableField__set_Index_Z524259A4, FsTableField_$ctor_Z18115A39, FsTableField__HeaderCell_10EBE01C, FsTableField__get_Index, FsTableField_$ctor_726EFFB, FsTableField__get_Name, FsTableField_$ctor_6547009B, FsTableField__get_Column } from "./FsTableField.js";
import { getItemFromDict, addToDict, addToSet } from "../../fable_modules/fable-library.4.1.3/MapUtil.js";
import { FsRangeColumn__get_Index, FsRangeColumn_$ctor_6A2513BC, FsRangeColumn_$ctor_Z524259A4 } from "../Ranges/FsRangeColumn.js";
import { FsCellsCollection__TryGetCell_Z37302880, FsCellsCollection__GetCellsInColumn_Z524259A4, Dictionary_tryGet } from "../Cells/FsCellsCollection.js";
import { replace } from "../../fable_modules/fable-library.4.1.3/String.js";
import { defaultArg, value as value_2 } from "../../fable_modules/fable-library.4.1.3/Option.js";
import { FsRangeRow__Cells_Z2740B3CA } from "../Ranges/FsRangeRow.js";
import { DataType } from "../Cells/FsCell.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";

export class FsTable extends FsRangeBase {
    constructor(name, rangeAddress, showTotalsRow, showHeaderRow) {
        super(rangeAddress);
        this._name = replace(name.trim(), " ", "_");
        this._lastRangeAddress = rangeAddress;
        this._showTotalsRow = defaultArg(showTotalsRow, false);
        this._showHeaderRow = defaultArg(showHeaderRow, true);
        this._fieldNames = (new Map([]));
        this._uniqueNames = (new Set([]));
    }
    get Name() {
        const self = this;
        return self._name;
    }
    GetFieldNames(cellsCollection) {
        const self = this;
        if ((!equals(self._fieldNames, defaultOf()) && !equals(self._lastRangeAddress, defaultOf())) && equals(self._lastRangeAddress, FsRangeBase__get_RangeAddress(self))) {
            return self._fieldNames;
        }
        else {
            self._lastRangeAddress = FsRangeBase__get_RangeAddress(self);
            return self._fieldNames;
        }
    }
    GetFields(cellsCollection) {
        const self = this;
        return initialize(FsRangeBase__ColumnCount(self), (i) => self.GetFieldAt(i, cellsCollection));
    }
    get ShowHeaderRow() {
        const self = this;
        return self._showHeaderRow;
    }
    set ShowHeaderRow(showHeaderRow) {
        const self = this;
        self._showHeaderRow = showHeaderRow;
    }
    HeadersRow() {
        const self = this;
        return !self.ShowHeaderRow ? defaultOf() : FsRange__FirstRow(FsRange_$ctor_6A2513BC(FsRangeBase__get_RangeAddress(self)));
    }
    GetColumns(cellsCollection) {
        const self = this;
        return delay(() => collect((i) => singleton(new FsColumn(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))), i), FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self))), i)), cellsCollection)), rangeDouble(FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))), 1, FsAddress__get_ColumnNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self))))));
    }
    RescanRange() {
        const self = this;
        FsRangeBase__set_RangeAddress_6A2513BC(self, reduce(FsRangeAddress__Union_6A2513BC, map((v) => FsRangeBase__get_RangeAddress(FsTableField__get_Column(v)), self._fieldNames.values())));
    }
    static rescanRange(table) {
        table.RescanRange();
        return table;
    }
    GetUniqueName(originalName, initialOffset, enforceOffset) {
        const this$ = this;
        let name = originalName + (enforceOffset ? int32ToString(initialOffset) : "");
        if (this$._uniqueNames.has(name)) {
            let i = initialOffset;
            name = (originalName + int32ToString(i));
            while (this$._uniqueNames.has(name)) {
                i = ((i + 1) | 0);
                name = (originalName + int32ToString(i));
            }
        }
        addToSet(name, this$._uniqueNames);
        return name;
    }
    static getUniqueNames(originalName, initialOffset, enforceOffset, table) {
        return table.GetUniqueName(originalName, initialOffset, enforceOffset);
    }
    InitFields(fieldNames) {
        const this$ = this;
        iterateIndexed((i, fn) => {
            const tableField = FsTableField_$ctor_6547009B(fn, i, FsRangeColumn_$ctor_Z524259A4(i));
            addToDict(this$._fieldNames, fn, tableField);
        }, fieldNames);
    }
    static initFields(fieldNames, table) {
        table.InitFields(fieldNames);
        return table;
    }
    AddFields(tableFields) {
        const this$ = this;
        iterate((tf) => {
            addToDict(this$._fieldNames, FsTableField__get_Name(tf), tf);
        }, tableFields);
    }
    static addFields(tableFields, table) {
        table.AddFields(tableFields);
        return table;
    }
    Field(name, cellsCollection) {
        let s, offset;
        const self = this;
        const matchValue = Dictionary_tryGet(name, self._fieldNames);
        if (matchValue == null) {
            const newField = FsTableField_$ctor_726EFFB(name, ((s = map(FsTableField__get_Index, self._fieldNames.values()), (length(s) === 0) ? 0 : max(s, {
                Compare: comparePrimitives,
            }))) + 1, FsRangeColumn_$ctor_6A2513BC((offset = (self._fieldNames.size | 0), FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))), FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))) + offset), FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self))), FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))) + offset)))), defaultOf(), defaultOf());
            if (self.ShowHeaderRow) {
                const value = FsTableField__HeaderCell_10EBE01C(newField, cellsCollection, true).SetValueAs(name);
            }
            addToDict(self._fieldNames, name, newField);
            self.RescanRange();
            return newField;
        }
        else {
            return matchValue;
        }
    }
    GetField(name, cellsCollection) {
        const self = this;
        const name_1 = replace(name, "\r\n", "\n");
        try {
            return getItemFromDict(self.GetFieldNames(cellsCollection), name_1);
        }
        catch (matchValue) {
            throw new Error(("The header row doesn\'t contain field name \'" + name_1) + "\'.");
        }
    }
    static getField(name, cellsCollection, table) {
        return table.GetField(name, cellsCollection);
    }
    GetFieldAt(index, cellsCollection) {
        const self = this;
        try {
            return find((ftf) => (FsTableField__get_Index(ftf) === index), self.GetFieldNames(cellsCollection).values());
        }
        catch (matchValue) {
            throw new Error(`FsTableField with index ${index} does not exist in the FsTable.`);
        }
    }
    GetFieldIndex(name, cellsCollection) {
        const self = this;
        return FsTableField__get_Index(self.GetField(name, cellsCollection)) | 0;
    }
    RenameField(oldName, newName) {
        const this$ = this;
        const matchValue = Dictionary_tryGet(oldName, this$._fieldNames);
        if (matchValue == null) {
            throw new Error("The FsTabelField does not exist in this FsTable", "oldName");
        }
        else {
            const field = matchValue;
            this$._fieldNames.delete(oldName);
            addToDict(this$._fieldNames, newName, field);
        }
    }
    static renameField(oldName, newName, table) {
        table.RenameField(oldName, newName);
        return table;
    }
    TryGetHeaderCellOfColumnAt(cellsCollection, colIndex) {
        const this$ = this;
        const fstRowIndex = FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(this$))) | 0;
        return tryFind((c) => (c.RowNumber === fstRowIndex), FsCellsCollection__GetCellsInColumn_Z524259A4(cellsCollection, colIndex));
    }
    static tryGetHeaderCellOfColumnIndexAt(cellsCollection, colIndex, table) {
        return table.TryGetHeaderCellOfColumnAt(cellsCollection, colIndex);
    }
    TryGetHeaderCellOfColumn(cellsCollection, column) {
        const this$ = this;
        return this$.TryGetHeaderCellOfColumnAt(cellsCollection, FsRangeColumn__get_Index(column));
    }
    static tryGetHeaderCellOfColumn(cellsCollection, column, table) {
        return table.TryGetHeaderCellOfColumn(cellsCollection, column);
    }
    GetHeaderCellOfColumnAt(cellsCollection, colIndex) {
        const this$ = this;
        return value_2(this$.TryGetHeaderCellOfColumnAt(cellsCollection, colIndex));
    }
    static getHeaderCellOfColumnIndexAt(cellsCollection, colIndex, table) {
        return table.GetHeaderCellOfColumnAt(cellsCollection, colIndex);
    }
    GetHeaderCellOfColumn(cellsCollection, column) {
        const this$ = this;
        return value_2(this$.TryGetHeaderCellOfColumn(cellsCollection, column));
    }
    static getHeaderCellOfColumn(cellsCollection, column, table) {
        return table.GetHeaderCellOfColumn(cellsCollection, column);
    }
    GetHeaderCellOfTableField(cellsCollection, tableField) {
        const this$ = this;
        return FsTableField__HeaderCell_10EBE01C(tableField, cellsCollection, this$.ShowHeaderRow);
    }
    static getHeaderCellOfTableField(cellsCollection, tableField, table) {
        return table.GetHeaderCellOfTableField(cellsCollection, tableField);
    }
    TryGetHeaderCellOfTableFieldAt(cellsCollection, tableFieldIndex) {
        const this$ = this;
        return tryPick((tf) => {
            if (FsTableField__get_Index(tf) === tableFieldIndex) {
                return FsTableField__HeaderCell_10EBE01C(tf, cellsCollection, this$.ShowHeaderRow);
            }
            else {
                return void 0;
            }
        }, this$._fieldNames.values());
    }
    static tryGetHeaderCellOfTableFieldIndexAt(cellsCollection, tableFieldIndex, table) {
        return table.TryGetHeaderCellOfTableFieldAt(cellsCollection, tableFieldIndex);
    }
    GetHeaderCellOfTableFieldAt(cellsCollection, tableFieldIndex) {
        const this$ = this;
        return value_2(this$.TryGetHeaderCellOfTableFieldAt(cellsCollection, tableFieldIndex));
    }
    static getHeaderCellOfTableFieldIndexAt(cellsCollection, tableFieldIndex, table) {
        return table.GetHeaderCellOfTableFieldAt(cellsCollection, tableFieldIndex);
    }
    TryGetHeaderCellByFieldName(cellsCollection, fieldName) {
        const this$ = this;
        const matchValue = Dictionary_tryGet(fieldName, this$._fieldNames);
        return (matchValue == null) ? void 0 : FsTableField__HeaderCell_10EBE01C(matchValue, cellsCollection, this$.ShowHeaderRow);
    }
    static tryGetHeaderCellByFieldName(cellsCollection, fieldName, table) {
        return table.TryGetHeaderCellByFieldName(cellsCollection, fieldName);
    }
    GetDataCellsOfColumnAt(cellsCollection, colIndex) {
        const this$ = this;
        return choose((ri) => FsCellsCollection__TryGetCell_Z37302880(cellsCollection, ri, colIndex), toList(rangeDouble(FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(this$))) + 1, 1, FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(this$))))));
    }
    static getDataCellsOfColumnIndexAt(cellsCollection, colIndex, table) {
        return table.GetDataCellsOfColumnAt(cellsCollection, colIndex);
    }
    Copy() {
        const this$ = this;
        const ra = FsRangeAddress__Copy(FsRangeBase__get_RangeAddress(this$));
        return new FsTable(this$.Name, ra, false, this$.ShowHeaderRow);
    }
    static copy(table) {
        return table.Copy();
    }
    RescanFieldNames(cellsCollection) {
        const this$ = this;
        if (this$.ShowHeaderRow) {
            const oldFieldNames = this$._fieldNames;
            this$._fieldNames = (new Map([]));
            let cellPos = 0;
            const enumerator = getEnumerator(FsRangeRow__Cells_Z2740B3CA(this$.HeadersRow(), cellsCollection));
            try {
                while (enumerator["System.Collections.IEnumerator.MoveNext"]()) {
                    const cell = enumerator["System.Collections.Generic.IEnumerator`1.get_Current"]();
                    let name = cell.Value;
                    const matchValue = Dictionary_tryGet(name, oldFieldNames);
                    if (matchValue == null) {
                        if ((name === defaultOf()) !== (name === "")) {
                            name = this$.GetUniqueName("Column", cellPos + 1, true);
                            const value = cell.SetValueAs(name);
                            cell.DataType = (new DataType(0, []));
                        }
                        if (this$._fieldNames.has(name)) {
                            throw new Error(("The header row contains more than one field name \'" + name) + "\'.");
                        }
                        addToDict(this$._fieldNames, name, FsTableField_$ctor_Z18115A39(name, cellPos));
                        cellPos = ((cellPos + 1) | 0);
                    }
                    else {
                        const tableField = matchValue;
                        FsTableField__set_Index_Z524259A4(tableField, cellPos);
                        addToDict(this$._fieldNames, name, tableField);
                        cellPos = ((cellPos + 1) | 0);
                    }
                }
            }
            finally {
                disposeSafe(enumerator);
            }
        }
        else {
            const colCount = FsRangeBase__ColumnCount(this$) | 0;
            for (let i = 1; i <= colCount; i++) {
                if (!exists((v) => (FsTableField__get_Index(v) === (i - 1)), this$._fieldNames.values())) {
                    const name_1 = "Column" + int32ToString(i);
                    addToDict(this$._fieldNames, name_1, FsTableField_$ctor_Z18115A39(name_1, i - 1));
                }
            }
        }
    }
}

export function FsTable_$reflection() {
    return class_type("FsSpreadsheet.FsTable", void 0, FsTable, FsRangeBase_$reflection());
}

export function FsTable_$ctor_30096B47(name, rangeAddress, showTotalsRow, showHeaderRow) {
    return new FsTable(name, rangeAddress, showTotalsRow, showHeaderRow);
}

