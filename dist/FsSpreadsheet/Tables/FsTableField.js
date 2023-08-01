import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { equals, defaultOf } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { FsRangeColumn__Cells_Z2740B3CA, FsRangeColumn__Copy, FsRangeColumn__FirstCell_Z2740B3CA, FsRangeColumn_$ctor_6A2513BC } from "../Ranges/FsRangeColumn.js";
import { FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "../Ranges/FsRangeAddress.js";
import { FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "../FsAddress.js";
import { FsRangeBase__get_RangeAddress } from "../Ranges/FsRangeBase.js";
import { printf, toFail } from "../../fable_modules/fable-library.4.1.3/String.js";
import { skip } from "../../fable_modules/fable-library.4.1.3/Seq.js";

export class FsTableField {
    constructor(name, index, column, totalsRowLabel, totalsRowFunction) {
        this._totalsRowsFunction = totalsRowFunction;
        this._totalsRowLabel = totalsRowLabel;
        this._column = column;
        this._index = (index | 0);
        this._name = name;
        this["Column@"] = this._column;
    }
}

export function FsTableField_$reflection() {
    return class_type("FsSpreadsheet.FsTableField", void 0, FsTableField);
}

export function FsTableField_$ctor_726EFFB(name, index, column, totalsRowLabel, totalsRowFunction) {
    return new FsTableField(name, index, column, totalsRowLabel, totalsRowFunction);
}

/**
 * Creates an empty FsTableField.
 */
export function FsTableField_$ctor() {
    return FsTableField_$ctor_726EFFB("", 0, defaultOf(), defaultOf(), defaultOf());
}

/**
 * Creates an FsTableField with the given name.
 */
export function FsTableField_$ctor_Z721C83C5(name) {
    return FsTableField_$ctor_726EFFB(name, 0, defaultOf(), defaultOf(), defaultOf());
}

/**
 * Creates an FsTableField with the given name and index.
 */
export function FsTableField_$ctor_Z18115A39(name, index) {
    return FsTableField_$ctor_726EFFB(name, index, defaultOf(), defaultOf(), defaultOf());
}

/**
 * Creates an FsTableField with the given name, index, and FsRangeColumn.
 */
export function FsTableField_$ctor_6547009B(name, index, column) {
    return FsTableField_$ctor_726EFFB(name, index, column, defaultOf(), defaultOf());
}

/**
 * Gets or sets the FsRangeColumn of this FsTableField.
 */
export function FsTableField__get_Column(__) {
    return __["Column@"];
}

/**
 * Gets or sets the FsRangeColumn of this FsTableField.
 */
export function FsTableField__set_Column_Z7F7BA1C4(__, v) {
    __["Column@"] = v;
}

/**
 * Gets or sets the 0-based index of the FsTableField inside the associated FsTable.
 * Sets the associated FsRangeColumn's column index accordingly.
 */
export function FsTableField__get_Index(this$) {
    return this$._index;
}

/**
 * Gets or sets the 0-based index of the FsTableField inside the associated FsTable.
 * Sets the associated FsRangeColumn's column index accordingly.
 */
export function FsTableField__set_Index_Z524259A4(this$, index) {
    if (index === this$._index) {
    }
    else {
        this$._index = (index | 0);
        if (equals(this$._column, defaultOf())) {
        }
        else {
            const indDiff = (index - this$._index) | 0;
            FsTableField__set_Column_Z7F7BA1C4(this$, FsRangeColumn_$ctor_6A2513BC(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(FsTableField__get_Column(this$)))), this$._index + indDiff), FsAddress_$ctor_Z37302880(FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(FsTableField__get_Column(this$)))), this$._index + indDiff))));
        }
    }
}

/**
 * The name of this FsTableField.
 */
export function FsTableField__get_Name(this$) {
    return this$._name;
}

/**
 * Sets the name of the FsTableField. If `showHeaderRow` is true, takes the respective FsCellsCollection and renames the header cell
 * according to the name of the FsTableField.
 */
export function FsTableField__SetName_103577A7(this$, name, cellsCollection, showHeaderRow) {
    this$._name = name;
    if (showHeaderRow) {
        FsRangeColumn__FirstCell_Z2740B3CA(FsTableField__get_Column(this$), cellsCollection).SetValueAs(name);
    }
}

/**
 * Sets the name of a given FsTableField. If `showHeaderRow` is true, takes the respective FsCellsCollection and renames the header cell
 * according to the name of the FsTableField.
 */
export function FsTableField_setName(name, cellsCollection, showHeaderRow, tableField) {
    FsTableField__SetName_103577A7(tableField, name, cellsCollection, showHeaderRow);
    return tableField;
}

/**
 * Creates a deep copy of this FsTableField.
 */
export function FsTableField__Copy(this$) {
    const col = FsRangeColumn__Copy(FsTableField__get_Column(this$));
    const ind = FsTableField__get_Index(this$) | 0;
    return FsTableField_$ctor_726EFFB(FsTableField__get_Name(this$), ind, col, defaultOf(), defaultOf());
}

/**
 * Returns a deep copy of a given FsTableField.
 */
export function FsTableField_copy_Z5343F797(tableField) {
    return FsTableField__Copy(tableField);
}

/**
 * Returns the header cell (taken from a given FsCellsCollection) for the FsTableField if `showHeaderRow` is true. Else fails.
 */
export function FsTableField__HeaderCell_10EBE01C(this$, cellsCollection, showHeaderRow) {
    if (!showHeaderRow) {
        const arg = this$._name;
        return toFail(printf("tried to get header cell of table field \"%s\" even though showHeaderRow is set to zero"))(arg);
    }
    else {
        return FsRangeColumn__FirstCell_Z2740B3CA(FsTableField__get_Column(this$), cellsCollection);
    }
}

/**
 * Returns the header cell (taken from an FsCellsCollection) for a given FsTableField if `showHeaderRow` is true. Else fails.
 */
export function FsTableField_getHeaderCell(cellsCollection, showHeaderRow, tableField) {
    return FsTableField__HeaderCell_10EBE01C(tableField, cellsCollection, showHeaderRow);
}

/**
 * Gets the collection of data cells for this FsTableField. Excludes the header and footer cells.
 */
export function FsTableField__DataCells_Z2740B3CA(this$, cellsCollection) {
    return skip(1, FsRangeColumn__Cells_Z2740B3CA(FsTableField__get_Column(this$), cellsCollection));
}

/**
 * Gets the collection of data cells for a given FsTableField. Excludes the header and footer cells.
 */
export function FsTableField_getDataCells(cellsCollection, tableField) {
    return FsTableField__DataCells_Z2740B3CA(tableField, cellsCollection);
}

