import { equals, defaultOf } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { FsRangeAddress__get_LastAddress, FsRangeAddress__get_FirstAddress, FsRangeAddress__Extend_6D30B323 } from "./FsRangeAddress.js";
import { FsAddress__get_FixedColumn, FsAddress__get_FixedRow, FsAddress_$ctor_Z4C746FC0, FsAddress__get_ColumnNumber, FsAddress__get_RowNumber } from "../FsAddress.js";
import { FsCellsCollection__GetCells_7E77A4A0, FsCellsCollection__Add_2E78CE33, FsCellsCollection__TryGetCell_Z37302880, FsCellsCollection__get_MaxColumnNumber, FsCellsCollection__get_MaxRowNumber } from "../Cells/FsCellsCollection.js";
import { printf, toFail } from "../../fable_modules/fable-library.4.1.3/String.js";
import { FsCell } from "../Cells/FsCell.js";

export class FsRangeBase {
    constructor(rangeAddress) {
        this._sortRows = defaultOf();
        this._sortColumns = defaultOf();
        this._rangeAddress = rangeAddress;
        let _id;
        FsRangeBase.IdCounter = ((FsRangeBase.IdCounter + 1) | 0);
        _id = FsRangeBase.IdCounter;
    }
}

export function FsRangeBase_$reflection() {
    return class_type("FsSpreadsheet.FsRangeBase", void 0, FsRangeBase);
}

export function FsRangeBase_$ctor_6A2513BC(rangeAddress) {
    return new FsRangeBase(rangeAddress);
}

(() => {
    FsRangeBase.IdCounter = 0;
})();

export function FsRangeBase__Extend_6D30B323(self, address) {
    FsRangeAddress__Extend_6D30B323(self._rangeAddress, address);
}

export function FsRangeBase__get_RangeAddress(self) {
    return self._rangeAddress;
}

export function FsRangeBase__set_RangeAddress_6A2513BC(self, rangeAdress) {
    if (!equals(rangeAdress, self._rangeAddress)) {
        const oldAddress = self._rangeAddress;
        self._rangeAddress = rangeAdress;
    }
}

export function FsRangeBase__Cell_Z3407A44B(self, cellAddressInRange, cells) {
    const absRow = ((FsAddress__get_RowNumber(cellAddressInRange) + FsAddress__get_RowNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))) - 1) | 0;
    const absColumn = ((FsAddress__get_ColumnNumber(cellAddressInRange) + FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)))) - 1) | 0;
    if ((absRow <= 0) ? true : (absRow > 1048576)) {
        const arg = FsCellsCollection__get_MaxRowNumber(cells) | 0;
        toFail(printf("Row number must be between 1 and %i"))(arg);
    }
    if ((absColumn <= 0) ? true : (absColumn > 16384)) {
        const arg_1 = FsCellsCollection__get_MaxColumnNumber(cells) | 0;
        toFail(printf("Column number must be between 1 and %i"))(arg_1);
    }
    const cell = FsCellsCollection__TryGetCell_Z37302880(cells, absRow, absColumn);
    if (cell == null) {
        const absoluteAddress = FsAddress_$ctor_Z4C746FC0(absRow, absColumn, FsAddress__get_FixedRow(cellAddressInRange), FsAddress__get_FixedColumn(cellAddressInRange));
        const newCell = FsCell.createEmptyWithAdress(absoluteAddress);
        FsRangeBase__Extend_6D30B323(self, absoluteAddress);
        const value = FsCellsCollection__Add_2E78CE33(cells, absRow, absColumn, newCell);
        return newCell;
    }
    else {
        return cell;
    }
}

export function FsRangeBase__Cells_Z2740B3CA(self, cells) {
    return FsCellsCollection__GetCells_7E77A4A0(cells, FsRangeAddress__get_FirstAddress(FsRangeBase__get_RangeAddress(self)), FsRangeAddress__get_LastAddress(FsRangeBase__get_RangeAddress(self)));
}

export function FsRangeBase__ColumnCount(self) {
    return (FsAddress__get_ColumnNumber(FsRangeAddress__get_LastAddress(self._rangeAddress)) - FsAddress__get_ColumnNumber(FsRangeAddress__get_FirstAddress(self._rangeAddress))) + 1;
}

