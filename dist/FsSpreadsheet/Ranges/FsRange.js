import { FsRangeBase__get_RangeAddress, FsRangeBase_$reflection, FsRangeBase } from "./FsRangeBase.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { int32ToString, defaultOf } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { FsAddress__get_ColumnNumber, FsAddress_$ctor_Z37302880, FsAddress__get_RowNumber } from "../FsAddress.js";
import { FsRangeAddress__get_LastAddress, FsRangeAddress_$ctor_7E77A4A0, FsRangeAddress__get_FirstAddress } from "./FsRangeAddress.js";
import { printf, toText } from "../../fable_modules/fable-library.4.1.3/String.js";
import { FsRangeRow_$ctor_6A2513BC } from "./FsRangeRow.js";

export class FsRange extends FsRangeBase {
    constructor(rangeAddress, styleValue) {
        super(rangeAddress);
    }
}

export function FsRange_$reflection() {
    return class_type("FsSpreadsheet.FsRange", void 0, FsRange, FsRangeBase_$reflection());
}

export function FsRange_$ctor_Z1F5897D9(rangeAddress, styleValue) {
    return new FsRange(rangeAddress, styleValue);
}

export function FsRange_$ctor_6A2513BC(rangeAddress) {
    return FsRange_$ctor_Z1F5897D9(rangeAddress, defaultOf());
}

export function FsRange_$ctor_65705F1F(rangeBase) {
    return FsRange_$ctor_Z1F5897D9(FsRangeBase__get_RangeAddress(rangeBase), defaultOf());
}

export function FsRange__Row_Z524259A4(self, row) {
    if ((row <= 0) ? true : (((row + FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))) - 1) > 1048576)) {
        throw new Error(int32ToString(row), toText(printf("Row number must be between 1 and %i"))(1048576));
    }
    return FsRangeRow_$ctor_6A2513BC(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880((FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))) + row) - 1, FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))), FsAddress_$ctor_Z37302880((FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))) + row) - 1, FsAddress__get_ColumnNumber(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self))))));
}

export function FsRange__FirstRow(self) {
    return FsRange__Row_Z524259A4(self, 1);
}

