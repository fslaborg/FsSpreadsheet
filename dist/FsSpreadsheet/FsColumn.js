import { FsRangeAddress__Copy, FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress_$ctor_7E77A4A0 } from "./Ranges/FsRangeAddress.js";
import { FsAddress__set_ColumnNumber_Z524259A4, FsAddress__get_ColumnNumber, FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "./FsAddress.js";
import { FsCellsCollection__Add_Z21F271A4, FsCellsCollection__GetCellsInColumn_Z524259A4, FsCellsCollection_$ctor } from "./Cells/FsCellsCollection.js";
import { map, maxBy, minBy, length } from "../fable_modules/fable-library.4.1.3/Seq.js";
import { toIterator, getEnumerator, comparePrimitives } from "../fable_modules/fable-library.4.1.3/Util.js";
import { FsRangeBase_$reflection, FsRangeBase, FsRangeBase__Cell_Z3407A44B, FsRangeBase__get_RangeAddress, FsRangeBase__Cells_Z2740B3CA } from "./Ranges/FsRangeBase.js";
import { class_type } from "../fable_modules/fable-library.4.1.3/Reflection.js";

export class FsColumn extends FsRangeBase {
    constructor(rangeAddress, cells) {
        super(rangeAddress);
        this["cells@18"] = cells;
    }
    static empty() {
        return new FsColumn(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(0, 0), FsAddress_$ctor_Z37302880(0, 0)), FsCellsCollection_$ctor());
    }
    static createAt(index, cells) {
        const getIndexBy = (f) => {
            if (length(FsCellsCollection__GetCellsInColumn_Z524259A4(cells, index)) === 0) {
                return 1;
            }
            else {
                return FsAddress__get_RowNumber(f((c) => FsAddress__get_RowNumber(c.Address))(FsCellsCollection__GetCellsInColumn_Z524259A4(cells, index)).Address) | 0;
            }
        };
        const minRowIndex = getIndexBy((projection) => ((source_1) => minBy(projection, source_1, {
            Compare: comparePrimitives,
        }))) | 0;
        const maxRowIndex = getIndexBy((projection_1) => ((source_2) => maxBy(projection_1, source_2, {
            Compare: comparePrimitives,
        }))) | 0;
        return new FsColumn(FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z37302880(minRowIndex, index), FsAddress_$ctor_Z37302880(maxRowIndex, index)), cells);
    }
    get Cells() {
        const self = this;
        return FsRangeBase__Cells_Z2740B3CA(self, self["cells@18"]);
    }
    get Index() {
        const self = this;
        return FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self))) | 0;
    }
    set Index(i) {
        const self = this;
        FsAddress__set_ColumnNumber_Z524259A4(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)), i);
        FsAddress__set_ColumnNumber_Z524259A4(FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self)), i);
    }
    Copy() {
        const self = this;
        const ra = FsRangeAddress__Copy(FsRangeBase__get_RangeAddress(self));
        const cells = map((c) => c.Copy(), self.Cells);
        const fcc = FsCellsCollection_$ctor();
        FsCellsCollection__Add_Z21F271A4(fcc, cells);
        return new FsColumn(ra, fcc);
    }
    static copy(column) {
        return column.Copy();
    }
    static getIndex(column) {
        return column.Index;
    }
    Item(rowIndex) {
        const this$ = this;
        return FsRangeBase__Cell_Z3407A44B(this$, FsAddress_$ctor_Z37302880(rowIndex, 1), this$["cells@18"]);
    }
    static item(rowIndex, column) {
        return column.Item(rowIndex);
    }
    GetEnumerator() {
        const this$ = this;
        return getEnumerator(this$.Cells);
    }
    [Symbol.iterator]() {
        return toIterator(getEnumerator(this));
    }
    "System.Collections.IEnumerable.GetEnumerator"() {
        const this$ = this;
        return getEnumerator(this$);
    }
}

export function FsColumn_$reflection() {
    return class_type("FsSpreadsheet.FsColumn", void 0, FsColumn, FsRangeBase_$reflection());
}

export function FsColumn_$ctor_7678C70A(rangeAddress, cells) {
    return new FsColumn(rangeAddress, cells);
}

