import { FsRangeBase__Cells_Z2740B3CA, FsRangeBase__Cell_Z3407A44B, FsRangeBase__get_RangeAddress, FsRangeBase_$reflection, FsRangeBase } from "./FsRangeBase.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { FsRangeAddress__Copy, FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "./FsRangeAddress.js";
import { FsAddress__get_RowNumber, FsAddress__set_ColumnNumber_Z524259A4, FsAddress__get_ColumnNumber, FsAddress_$ctor_Z37302880 } from "../FsAddress.js";
import { map, toList } from "../../fable_modules/fable-library.4.1.3/Seq.js";
import { rangeDouble } from "../../fable_modules/fable-library.4.1.3/Range.js";

export class FsRangeColumn extends FsRangeBase {
    constructor(rangeAddress) {
        super(rangeAddress);
    }
}

export function FsRangeColumn_$reflection() {
    return class_type("FsSpreadsheet.FsRangeColumn", void 0, FsRangeColumn, FsRangeBase_$reflection());
}

export function FsRangeColumn_$ctor_6A2513BC(rangeAddress) {
    return new FsRangeColumn(rangeAddress);
}

export function FsRangeColumn_$ctor_Z524259A4(index) {
    return FsRangeColumn_$ctor_6A2513BC(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(0, index), FsAddress_$ctor_Z37302880(0, index)));
}

export function FsRangeColumn__get_Index(self) {
    return FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)));
}

export function FsRangeColumn__set_Index_Z524259A4(self, i) {
    FsAddress__set_ColumnNumber_Z524259A4(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)), i);
    FsAddress__set_ColumnNumber_Z524259A4(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self)), i);
}

export function FsRangeColumn__Cell_Z4232C216(self, rowIndex, cellsCollection) {
    return FsRangeBase__Cell_Z3407A44B(self, FsAddress_$ctor_Z37302880((rowIndex - FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))) + 1, 1), cellsCollection);
}

export function FsRangeColumn__FirstCell_Z2740B3CA(self, cells) {
    return FsRangeBase__Cell_Z3407A44B(self, FsAddress_$ctor_Z37302880(1, 1), cells);
}

export function FsRangeColumn__Cells_Z2740B3CA(self, cellsCollection) {
    return FsRangeBase__Cells_Z2740B3CA(self, cellsCollection);
}

export function FsRangeColumn_fromRangeAddress_6A2513BC(rangeAddress) {
    return FsRangeColumn_$ctor_6A2513BC(rangeAddress);
}

/**
 * Creates a deep copy of this FsRangeColumn.
 */
export function FsRangeColumn__Copy(self) {
    return FsRangeColumn_$ctor_6A2513BC(FsRangeAddress__Copy(FsRangeBase__get_RangeAddress(self)));
}

/**
 * Returns a deep copy of a given FsRangeColumn.
 */
export function FsRangeColumn_copy_Z7F7BA1C4(rangeColumn) {
    return FsRangeColumn__Copy(rangeColumn);
}

/**
 * Takes an FsRangeAddress and returns, for every column the range address spans, an FsRangeColumn.
 */
export function FsSpreadsheet_FsRangeAddress__FsRangeAddress_toRangeColumns_Static_6A2513BC(rangeAddress) {
    const columns = toList(rangeDouble(FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(rangeAddress)), 1, FsAddress__get_ColumnNumber(FsRangeAddress__get_LastAddress(rangeAddress))));
    const fstRow = FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(rangeAddress)) | 0;
    const lstRow = FsAddress__get_RowNumber(FsRangeAddress__get_LastAddress(rangeAddress)) | 0;
    return map((c) => FsRangeColumn_$ctor_6A2513BC(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(fstRow, c), FsAddress_$ctor_Z37302880(lstRow, c))), columns);
}

