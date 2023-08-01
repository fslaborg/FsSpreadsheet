import { addToSet, addToDict, getItemFromDict } from "../../fable_modules/fable-library.4.1.3/MapUtil.js";
import { some } from "../../fable_modules/fable-library.4.1.3/Option.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { minBy, min, singleton, empty, delay, collect, max, isEmpty, iterate, map } from "../../fable_modules/fable-library.4.1.3/Seq.js";
import { FsAddress_$ctor_Z37302880, FsAddress__get_ColumnNumber, FsAddress__get_RowNumber } from "../FsAddress.js";
import { comparePrimitives } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { rangeDouble } from "../../fable_modules/fable-library.4.1.3/Range.js";

export function Dictionary_tryGet(k, dict) {
    if (dict.has(k)) {
        return some(getItemFromDict(dict, k));
    }
    else {
        return void 0;
    }
}

export class FsCellsCollection {
    constructor() {
        this._columnsUsed = (new Map([]));
        this._deleted = (new Map([]));
        this._rowsCollection = (new Map([]));
        this._maxColumnUsed = 0;
        this._maxRowUsed = 0;
        this._rowsUsed = (new Map([]));
        this._count = 0;
    }
}

export function FsCellsCollection_$reflection() {
    return class_type("FsSpreadsheet.FsCellsCollection", void 0, FsCellsCollection);
}

export function FsCellsCollection_$ctor() {
    return new FsCellsCollection();
}

export function FsCellsCollection__get_Count(this$) {
    return this$._count;
}

function FsCellsCollection__set_Count_Z524259A4(this$, count) {
    this$._count = (count | 0);
}

/**
 * The highest rowIndex in The FsCellsCollection.
 */
export function FsCellsCollection__get_MaxRowNumber(this$) {
    return this$._maxRowUsed;
}

/**
 * The highest columnIndex in The FsCellsCollection.
 */
export function FsCellsCollection__get_MaxColumnNumber(this$) {
    return this$._maxColumnUsed;
}

/**
 * Creates a deep copy of the FsCellsCollection.
 */
export function FsCellsCollection__Copy(this$) {
    return FsCellsCollection_createFromCells_Z21F271A4(map((c) => c.Copy(), FsCellsCollection__GetCells(this$)));
}

/**
 * Returns a deep copy of a given FsCellsCollection.
 */
export function FsCellsCollection_copy_Z2740B3CA(cellsCollection) {
    return FsCellsCollection__Copy(cellsCollection);
}

function FsCellsCollection_IncrementUsage_71185086(dictionary, key) {
    const matchValue = Dictionary_tryGet(key, dictionary);
    if (matchValue == null) {
        addToDict(dictionary, key, 1);
    }
    else {
        const count = matchValue | 0;
        dictionary.set(key, count + 1);
    }
}

function FsCellsCollection_DecrementUsage_71185086(dictionary, key) {
    const matchValue = Dictionary_tryGet(key, dictionary);
    if (matchValue == null) {
        return false;
    }
    else if (matchValue > 1) {
        const count_1 = matchValue | 0;
        dictionary.set(key, count_1 - 1);
        return false;
    }
    else {
        dictionary.delete(key);
        return true;
    }
}

/**
 * Creates an FsCellsCollection from the given FsCells.
 */
export function FsCellsCollection_createFromCells_Z21F271A4(cells) {
    const fcc = FsCellsCollection_$ctor();
    FsCellsCollection__Add_Z21F271A4(fcc, cells);
    return fcc;
}

/**
 * Empties the whole FsCellsCollection.
 */
export function FsCellsCollection__Clear(this$) {
    this$._count = 0;
    this$._rowsUsed.clear();
    this$._columnsUsed.clear();
    this$._rowsCollection.clear();
    this$._maxRowUsed = 0;
    this$._maxColumnUsed = 0;
    return this$;
}

/**
 * Adds an FsCell of given rowIndex and columnIndex to the FsCellsCollection.
 */
export function FsCellsCollection__Add_2E78CE33(this$, row, column, cell) {
    let matchValue, columnsCollection_1;
    cell.RowNumber = (row | 0);
    cell.ColumnNumber = (column | 0);
    this$._count = ((this$._count + 1) | 0);
    FsCellsCollection_IncrementUsage_71185086(this$._rowsUsed, row);
    FsCellsCollection_IncrementUsage_71185086(this$._columnsUsed, column);
    addToDict((matchValue = Dictionary_tryGet(row, this$._rowsCollection), (matchValue == null) ? ((columnsCollection_1 = (new Map([])), (addToDict(this$._rowsCollection, row, columnsCollection_1), columnsCollection_1))) : matchValue), column, cell);
    if (row > this$._maxRowUsed) {
        this$._maxRowUsed = (row | 0);
    }
    if (column > this$._maxColumnUsed) {
        this$._maxColumnUsed = (column | 0);
    }
    const matchValue_1 = Dictionary_tryGet(row, this$._deleted);
    if (matchValue_1 == null) {
    }
    else {
        const delHash = matchValue_1;
        delHash.delete(column);
    }
}

/**
 * Adds an FsCell of given rowIndex and columnIndex to an FsCellsCollection.
 */
export function FsCellsCollection_addCellWithIndeces(rowIndex, colIndex, cell, cellsCollection) {
    FsCellsCollection__Add_2E78CE33(cellsCollection, rowIndex, colIndex, cell);
}

/**
 * Adds an FsCell to the FsCellsCollection.
 */
export function FsCellsCollection__Add_Z334DF64D(this$, cell) {
    FsCellsCollection__Add_2E78CE33(this$, FsAddress__get_RowNumber(cell.Address), FsAddress__get_ColumnNumber(cell.Address), cell);
}

/**
 * Adds an FsCell to an FsCellsCollection.
 */
export function FsCellsCollection_addCell(cell, cellsCollection) {
    FsCellsCollection__Add_Z334DF64D(cellsCollection, cell);
    return cellsCollection;
}

/**
 * Adds FsCells to the FsCellsCollection.
 */
export function FsCellsCollection__Add_Z21F271A4(this$, cells) {
    iterate((arg_1) => {
        const value = FsCellsCollection__Add_Z334DF64D(this$, arg_1);
    }, cells);
}

/**
 * Adds FsCells to an FsCellsCollection.
 */
export function FsCellsCollection_addCells(cells, cellsCollection) {
    FsCellsCollection__Add_Z21F271A4(cellsCollection, cells);
    return cellsCollection;
}

/**
 * Checks if an FsCell exists at given row- and columnIndex.
 */
export function FsCellsCollection__ContainsCellAt_Z37302880(this$, rowIndex, colIndex) {
    const matchValue = Dictionary_tryGet(rowIndex, this$._rowsCollection);
    if (matchValue == null) {
        return false;
    }
    else {
        const colsCollection = matchValue;
        return colsCollection.has(colIndex);
    }
}

/**
 * Checks if an FsCell exists at given row- and columnIndex of a given FsCellsCollection.
 */
export function FsCellsCollection_containsCellAt(rowIndex, colIndex, cellsCollection) {
    return FsCellsCollection__ContainsCellAt_Z37302880(cellsCollection, rowIndex, colIndex);
}

/**
 * Removes an FsCell of given rowIndex and columnIndex from the FsCellsCollection.
 */
export function FsCellsCollection__RemoveCellAt_Z37302880(this$, row, column) {
    let delHash;
    this$._count = ((this$._count - 1) | 0);
    const rowRemoved = FsCellsCollection_DecrementUsage_71185086(this$._rowsUsed, row);
    const columnRemoved = FsCellsCollection_DecrementUsage_71185086(this$._columnsUsed, column);
    if (rowRemoved && (row === this$._maxRowUsed)) {
        this$._maxRowUsed = ((!isEmpty(this$._rowsUsed.keys()) ? max(this$._rowsUsed.keys(), {
            Compare: comparePrimitives,
        }) : 0) | 0);
    }
    if (columnRemoved && (column === this$._maxColumnUsed)) {
        this$._maxColumnUsed = ((!isEmpty(this$._columnsUsed.keys()) ? max(this$._columnsUsed.keys(), {
            Compare: comparePrimitives,
        }) : 0) | 0);
    }
    const matchValue = Dictionary_tryGet(row, this$._deleted);
    if (matchValue == null) {
        const delHash_3 = new Set([]);
        addToSet(column, delHash_3);
        addToDict(this$._deleted, row, delHash_3);
    }
    else if ((delHash = matchValue, delHash.has(column))) {
        const delHash_1 = matchValue;
    }
    else {
        const delHash_2 = matchValue;
        addToSet(column, delHash_2);
    }
    const matchValue_1 = Dictionary_tryGet(row, this$._rowsCollection);
    if (matchValue_1 == null) {
    }
    else {
        const columnsCollection = matchValue_1;
        columnsCollection.delete(column);
        if (columnsCollection.size === 0) {
            this$._rowsCollection.delete(row);
        }
    }
}

/**
 * Removes an FsCell of given rowIndex and columnIndex from an FsCellsCollection.
 */
export function FsCellsCollection_removeCellAt(rowIndex, colIndex, cellsCollection) {
    FsCellsCollection__RemoveCellAt_Z37302880(cellsCollection, rowIndex, colIndex);
    return cellsCollection;
}

/**
 * Removes the value of an FsCell at given row- and columnIndex if it exists from the FsCollection.
 */
export function FsCellsCollection__TryRemoveValueAt_Z37302880(this$, rowIndex, colIndex) {
    const matchValue = Dictionary_tryGet(rowIndex, this$._rowsCollection);
    if (matchValue == null) {
    }
    else {
        const colsCollection = matchValue;
        try {
            getItemFromDict(colsCollection, colIndex).Value = "";
        }
        catch (matchValue_1) {
        }
    }
}

/**
 * Removes the value of an FsCell at given row- and columnIndex if it exists from a given FsCollection.
 */
export function FsCellsCollection_tryRemoveValueAt(rowIndex, colIndex, cellsCollection) {
    FsCellsCollection__TryRemoveValueAt_Z37302880(cellsCollection, rowIndex, colIndex);
    return cellsCollection;
}

/**
 * Removes the value of an FsCell at given row- and columnIndex from the FsCollection.
 */
export function FsCellsCollection__RemoveValueAt_Z37302880(this$, rowIndex, colIndex) {
    getItemFromDict(getItemFromDict(this$._rowsCollection, rowIndex), colIndex).Value = "";
}

/**
 * Removes the value of an FsCell at given row- and columnIndex from a given FsCollection.
 */
export function FsCellsCollection_removeValueAt(rowIndex, colIndex, cellsCollection) {
    FsCellsCollection__RemoveValueAt_Z37302880(cellsCollection, rowIndex, colIndex);
    return cellsCollection;
}

/**
 * Returns all FsCells of the FsCellsCollection.
 */
export function FsCellsCollection__GetCells(this$) {
    return collect((columnsCollection) => columnsCollection.values(), this$._rowsCollection.values());
}

/**
 * Returns all FsCells of the FsCellsCollection.
 */
export function FsCellsCollection_getCells_Z2740B3CA(cellsCollection) {
    return FsCellsCollection__GetCells(cellsCollection);
}

/**
 * Returns the FsCells from given rowStart to rowEnd and columnStart to columnEnd and fulfilling the predicate.
 */
export function FsCellsCollection__GetCells_6611854F(this$, rowStart, columnStart, rowEnd, columnEnd, predicate) {
    const finalRow = ((rowEnd > this$._maxRowUsed) ? this$._maxRowUsed : rowEnd) | 0;
    const finalColumn = ((columnEnd > this$._maxColumnUsed) ? this$._maxColumnUsed : columnEnd) | 0;
    return delay(() => collect((ro) => {
        const matchValue = Dictionary_tryGet(ro, this$._rowsCollection);
        if (matchValue == null) {
            return empty();
        }
        else {
            const columnsCollection = matchValue;
            return collect((co) => {
                const matchValue_1 = Dictionary_tryGet(co, columnsCollection);
                let matchResult, cell_1;
                if (matchValue_1 != null) {
                    if (predicate(matchValue_1)) {
                        matchResult = 0;
                        cell_1 = matchValue_1;
                    }
                    else {
                        matchResult = 1;
                    }
                }
                else {
                    matchResult = 1;
                }
                switch (matchResult) {
                    case 0:
                        return singleton(cell_1);
                    default: {
                        return empty();
                    }
                }
            }, rangeDouble(columnStart, 1, finalColumn));
        }
    }, rangeDouble(rowStart, 1, finalRow)));
}

/**
 * Returns the FsCells from an FsCellsCollection with given rowStart to rowEnd and columnStart to columnEnd and fulfilling the predicate.
 */
export function FsCellsCollection_filterCellsFromTo(rowStart, columnStart, rowEnd, columnEnd, predicate, cellsCollection) {
    return FsCellsCollection__GetCells_6611854F(cellsCollection, rowStart, columnStart, rowEnd, columnEnd, predicate);
}

/**
 * Returns the FsCells from given startAddress to lastAddress and fulfilling the predicate.
 */
export function FsCellsCollection__GetCells_24D826EF(this$, startAddress, lastAddress, predicate) {
    return FsCellsCollection__GetCells_6611854F(this$, FsAddress__get_RowNumber(startAddress), FsAddress__get_ColumnNumber(startAddress), FsAddress__get_RowNumber(lastAddress), FsAddress__get_ColumnNumber(lastAddress), predicate);
}

/**
 * Returns the FsCells from an FsCellsCollection with given startAddress to lastAddress and fulfilling the predicate.
 */
export function FsCellsCollection_filterCellsFromToAddress(startAddress, lastAddress, predicate, cellsCollection) {
    return FsCellsCollection__GetCells_24D826EF(cellsCollection, startAddress, lastAddress, predicate);
}

/**
 * Returns the FsCells from given rowStart to rowEnd and columnStart to columnEnd.
 */
export function FsCellsCollection__GetCells_Z6C21C500(this$, rowStart, columnStart, rowEnd, columnEnd) {
    const finalRow = ((rowEnd > this$._maxRowUsed) ? this$._maxRowUsed : rowEnd) | 0;
    const finalColumn = ((columnEnd > this$._maxColumnUsed) ? this$._maxColumnUsed : columnEnd) | 0;
    return delay(() => collect((ro) => {
        const matchValue = Dictionary_tryGet(ro, this$._rowsCollection);
        if (matchValue == null) {
            return empty();
        }
        else {
            const columnsCollection = matchValue;
            return collect((co) => {
                const matchValue_1 = Dictionary_tryGet(co, columnsCollection);
                if (matchValue_1 != null) {
                    return singleton(matchValue_1);
                }
                else {
                    return empty();
                }
            }, rangeDouble(columnStart, 1, finalColumn));
        }
    }, rangeDouble(rowStart, 1, finalRow)));
}

/**
 * Returns the FsCells from an FsCellsCollection with given rowStart to rowEnd and columnStart to columnEnd.
 */
export function FsCellsCollection_getCellsFromTo(rowStart, columnStart, rowEnd, columnEnd, cellsCollection) {
    return FsCellsCollection__GetCells_Z6C21C500(cellsCollection, rowStart, columnStart, rowEnd, columnEnd);
}

/**
 * Returns the FsCells from given startAddress to lastAddress.
 */
export function FsCellsCollection__GetCells_7E77A4A0(this$, startAddress, lastAddress) {
    return FsCellsCollection__GetCells_Z6C21C500(this$, FsAddress__get_RowNumber(startAddress), FsAddress__get_ColumnNumber(startAddress), FsAddress__get_RowNumber(lastAddress), FsAddress__get_ColumnNumber(lastAddress));
}

/**
 * Returns the FsCells from an FsCellsCollection with given startAddress to lastAddress.
 */
export function FsCellsCollection_getCellsFromToAddress(startAddress, lastAddress, cellsCollection) {
    return FsCellsCollection__GetCells_7E77A4A0(cellsCollection, startAddress, lastAddress);
}

/**
 * Returns the FsCell at given rowIndex and columnIndex if it exists. Otherwise returns None.
 */
export function FsCellsCollection__TryGetCell_Z37302880(this$, row, column) {
    if ((row > this$._maxRowUsed) ? true : (column > this$._maxColumnUsed)) {
        return void 0;
    }
    else {
        const matchValue = Dictionary_tryGet(row, this$._rowsCollection);
        if (matchValue == null) {
            return void 0;
        }
        else {
            const matchValue_1 = Dictionary_tryGet(column, matchValue);
            if (matchValue_1 == null) {
                return void 0;
            }
            else {
                return matchValue_1;
            }
        }
    }
}

/**
 * Returns the FsCell from an FsCellsCollection at given rowIndex and columnIndex if it exists. Otherwise returns None.
 */
export function FsCellsCollection_tryGetCell(rowIndex, colIndex, cellsCollection) {
    return FsCellsCollection__TryGetCell_Z37302880(cellsCollection, rowIndex, colIndex);
}

/**
 * Returns all FsCells in the given columnIndex.
 */
export function FsCellsCollection__GetCellsInColumn_Z524259A4(this$, colIndex) {
    return FsCellsCollection__GetCells_Z6C21C500(this$, 1, colIndex, this$._maxRowUsed, colIndex);
}

/**
 * Returns all FsCells in an FsCellsCollection with the given columnIndex.
 */
export function FsCellsCollection_getCellsInColumn(colIndex, cellsCollection) {
    return FsCellsCollection__GetCellsInColumn_Z524259A4(cellsCollection, colIndex);
}

/**
 * Returns all FsCells in the given rowIndex.
 */
export function FsCellsCollection__GetCellsInRow_Z524259A4(this$, rowIndex) {
    return FsCellsCollection__GetCells_Z6C21C500(this$, rowIndex, 1, rowIndex, this$._maxColumnUsed);
}

/**
 * Returns all FsCells in an FsCellsCollection with the given rowIndex.
 */
export function FsCellsCollection_getCellsInRow(rowIndex, cellsCollection) {
    return FsCellsCollection__GetCellsInRow_Z524259A4(cellsCollection, rowIndex);
}

/**
 * Returns the upper left corner of the FsCellsCollection.
 */
export function FsCellsCollection__GetFirstAddress(this$) {
    let d_1;
    if (isEmpty(this$._rowsCollection) ? true : isEmpty(this$._rowsCollection.keys())) {
        return FsAddress_$ctor_Z37302880(0, 0);
    }
    else {
        return FsAddress_$ctor_Z37302880(min(this$._rowsCollection.keys(), {
            Compare: comparePrimitives,
        }), (d_1 = minBy((d) => min(d.keys(), {
            Compare: comparePrimitives,
        }), this$._rowsCollection.values(), {
            Compare: comparePrimitives,
        }), min(d_1.keys(), {
            Compare: comparePrimitives,
        })));
    }
}

/**
 * Returns the upper left corner of a given FsCellsCollection.
 */
export function FsCellsCollection_getFirstAddress_Z2740B3CA(cells) {
    return FsCellsCollection__GetFirstAddress(cells);
}

/**
 * Returns the lower right corner of the FsCellsCollection.
 */
export function FsCellsCollection__GetLastAddress(this$) {
    return FsAddress_$ctor_Z37302880(FsCellsCollection__get_MaxRowNumber(this$), FsCellsCollection__get_MaxColumnNumber(this$));
}

/**
 * Returns the lower right corner of a given FsCellsCollection.
 */
export function FsCellsCollection_getLastAddress_Z2740B3CA(cells) {
    return FsCellsCollection__GetLastAddress(cells);
}

