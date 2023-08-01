import { FsRangeBase__Cells_Z2740B3CA, FsRangeBase__Cell_Z3407A44B, FsRangeBase__get_RangeAddress, FsRangeBase_$reflection, FsRangeBase } from "./FsRangeBase.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "./FsRangeAddress.js";
import { FsAddress__get_ColumnNumber, FsAddress__set_RowNumber_Z524259A4, FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "../FsAddress.js";

export class FsRangeRow extends FsRangeBase {
    constructor(rangeAddress) {
        super(rangeAddress);
    }
}

export function FsRangeRow_$reflection() {
    return class_type("FsSpreadsheet.FsRangeRow", void 0, FsRangeRow, FsRangeBase_$reflection());
}

export function FsRangeRow_$ctor_6A2513BC(rangeAddress) {
    return new FsRangeRow(rangeAddress);
}

export function FsRangeRow_$ctor_Z524259A4(index) {
    return FsRangeRow_$ctor_6A2513BC(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(index, 0), FsAddress_$ctor_Z37302880(index, 0)));
}

export function FsRangeRow__get_Index(self) {
    return FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)));
}

export function FsRangeRow__set_Index_Z524259A4(self, i) {
    FsAddress__set_RowNumber_Z524259A4(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)), i);
    FsAddress__set_RowNumber_Z524259A4(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self)), i);
}

export function FsRangeRow__Cell_Z4232C216(self, columnIndex, cells) {
    return FsRangeBase__Cell_Z3407A44B(self, FsAddress_$ctor_Z37302880(1, (columnIndex - FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))) + 1), cells);
}

export function FsRangeRow__Cells_Z2740B3CA(self, cells) {
    return FsRangeBase__Cells_Z2740B3CA(self, cells);
}

