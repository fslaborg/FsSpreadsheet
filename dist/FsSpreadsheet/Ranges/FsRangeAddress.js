import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { toConsole, split, printf, toText } from "../../fable_modules/fable-library.4.1.3/String.js";
import { FsAddress__get_Address, FsAddress_$ctor_Z37302880, FsAddress__set_ColumnNumber_Z524259A4, FsAddress__get_ColumnNumber, FsAddress__set_RowNumber_Z524259A4, FsAddress__get_RowNumber, FsAddress_$ctor_Z721C83C5, CellReference_moveHorizontal, CellReference_toIndices } from "../FsAddress.js";

export class FsRangeAddress {
    constructor(firstAddress, lastAddress) {
        this.firstAddress = firstAddress;
        this.lastAddress = lastAddress;
        this._firstAddress = this.firstAddress;
        this._lastAddress = this.lastAddress;
    }
    toString() {
        const self = this;
        return FsRangeAddress__get_Range(self);
    }
}

export function FsRangeAddress_$reflection() {
    return class_type("FsSpreadsheet.FsRangeAddress", void 0, FsRangeAddress);
}

export function FsRangeAddress_$ctor_7E77A4A0(firstAddress, lastAddress) {
    return new FsRangeAddress(firstAddress, lastAddress);
}

/**
 * Given A1-based top left start and bottom right end indices, returns a "A1:A1"-style area-
 */
export function Range_ofBoundaries(fromCellReference, toCellReference) {
    return toText(printf("%s:%s"))(fromCellReference)(toCellReference);
}

/**
 * Given a "A1:A1"-style area, returns A1-based cell start and end cellReferences.
 */
export function Range_toBoundaries(area) {
    const a = split(area, [":"], void 0, 0);
    return [a[0], a[1]];
}

/**
 * Gets the right boundary of the area.
 */
export function Range_rightBoundary(area) {
    return CellReference_toIndices(Range_toBoundaries(area)[1])[0];
}

/**
 * Gets the left boundary of the area.
 */
export function Range_leftBoundary(area) {
    return CellReference_toIndices(Range_toBoundaries(area)[0])[0];
}

/**
 * Gets the Upper boundary of the area.
 */
export function Range_upperBoundary(area) {
    return CellReference_toIndices(Range_toBoundaries(area)[0])[1];
}

/**
 * Gets the lower boundary of the area.
 */
export function Range_lowerBoundary(area) {
    return CellReference_toIndices(Range_toBoundaries(area)[1])[1];
}

/**
 * Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
 */
export function Range_moveHorizontal(amount, area) {
    let tupledArg_1;
    const tupledArg = Range_toBoundaries(area);
    tupledArg_1 = [CellReference_moveHorizontal(amount, tupledArg[0]), CellReference_moveHorizontal(amount, tupledArg[1])];
    return Range_ofBoundaries(tupledArg_1[0], tupledArg_1[1]);
}

/**
 * Moves both start and end of the area by the given amount (positive amount moves area to right and vice versa).
 */
export function Range_moveVertical(amount, area) {
    let tupledArg_1;
    const tupledArg = Range_toBoundaries(area);
    tupledArg_1 = [CellReference_moveHorizontal(amount, tupledArg[0]), CellReference_moveHorizontal(amount, tupledArg[1])];
    return Range_ofBoundaries(tupledArg_1[0], tupledArg_1[1]);
}

/**
 * Extends the right boundary of the area by the given amount (positive amount increases area to right and vice versa).
 */
export function Range_extendRight(amount, area) {
    let tupledArg_1;
    const tupledArg = Range_toBoundaries(area);
    tupledArg_1 = [tupledArg[0], CellReference_moveHorizontal(amount, tupledArg[1])];
    return Range_ofBoundaries(tupledArg_1[0], tupledArg_1[1]);
}

/**
 * Extends the left boundary of the area by the given amount (positive amount decreases the area to left and vice versa).
 */
export function Range_extendLeft(amount, area) {
    let tupledArg_1;
    const tupledArg = Range_toBoundaries(area);
    tupledArg_1 = [CellReference_moveHorizontal(amount, tupledArg[0]), tupledArg[1]];
    return Range_ofBoundaries(tupledArg_1[0], tupledArg_1[1]);
}

/**
 * Returns true if the column index of the reference exceeds the right boundary of the area.
 */
export function Range_referenceExceedsAreaRight(reference, area) {
    return CellReference_toIndices(reference)[0] > Range_rightBoundary(area);
}

/**
 * Returns true if the column index of the reference exceeds the left boundary of the area.
 */
export function Range_referenceExceedsAreaLeft(reference, area) {
    return CellReference_toIndices(reference)[0] < Range_leftBoundary(area);
}

/**
 * Returns true if the column index of the reference exceeds the upper boundary of the area.
 */
export function Range_referenceExceedsAreaAbove(reference, area) {
    return CellReference_toIndices(reference)[1] > Range_upperBoundary(area);
}

/**
 * Returns true if the column index of the reference exceeds the lower boundary of the area.
 */
export function Range_referenceExceedsAreaBelow(reference, area) {
    return CellReference_toIndices(reference)[1] < Range_lowerBoundary(area);
}

/**
 * Returns true if the reference does not lie in the boundary of the area.
 */
export function Range_referenceExceedsArea(reference, area) {
    if ((Range_referenceExceedsAreaRight(reference, area) ? true : Range_referenceExceedsAreaLeft(reference, area)) ? true : Range_referenceExceedsAreaAbove(reference, area)) {
        return true;
    }
    else {
        return Range_referenceExceedsAreaBelow(reference, area);
    }
}

/**
 * Returns true if the A1:A1-style area is of correct format.
 */
export function Range_isCorrect(area) {
    try {
        const hor = Range_leftBoundary(area) <= Range_rightBoundary(area);
        const ver = Range_upperBoundary(area) <= Range_lowerBoundary(area);
        if (!hor) {
            toConsole(printf("Right area boundary must be higher or equal to left area boundary."));
        }
        if (!ver) {
            toConsole(printf("Lower area boundary must be higher or equal to upper area boundary."));
        }
        return hor && ver;
    }
    catch (err) {
        const arg_1 = err.message;
        toConsole(printf("Area \"%s\" could not be parsed: %s"))(area)(arg_1);
        return false;
    }
}

export function FsRangeAddress_$ctor_Z721C83C5(rangeAddress) {
    const patternInput = Range_toBoundaries(rangeAddress);
    return FsRangeAddress_$ctor_7E77A4A0(FsAddress_$ctor_Z721C83C5(patternInput[0]), FsAddress_$ctor_Z721C83C5(patternInput[1]));
}

/**
 * Creates a deep copy of this FsRangeAddress.
 */
export function FsRangeAddress__Copy(self) {
    return FsRangeAddress_$ctor_Z721C83C5(FsRangeAddress__get_Range(self));
}

/**
 * Returns a deep copy of a given FsRangeAddress.
 */
export function FsRangeAddress_copy_6A2513BC(rangeAddress) {
    return FsRangeAddress__Copy(rangeAddress);
}

export function FsRangeAddress__Extend_6D30B323(self, address) {
    if (FsAddress__get_RowNumber(address) < FsAddress__get_RowNumber(self._firstAddress)) {
        FsAddress__set_RowNumber_Z524259A4(self._firstAddress, FsAddress__get_RowNumber(address));
    }
    if (FsAddress__get_RowNumber(address) > FsAddress__get_RowNumber(self._lastAddress)) {
        FsAddress__set_RowNumber_Z524259A4(self._lastAddress, FsAddress__get_RowNumber(address));
    }
    if (FsAddress__get_ColumnNumber(address) < FsAddress__get_ColumnNumber(self._firstAddress)) {
        FsAddress__set_ColumnNumber_Z524259A4(self._firstAddress, FsAddress__get_ColumnNumber(address));
    }
    if (FsAddress__get_ColumnNumber(address) > FsAddress__get_ColumnNumber(self._lastAddress)) {
        FsAddress__set_ColumnNumber_Z524259A4(self._lastAddress, FsAddress__get_ColumnNumber(address));
    }
}

export function FsRangeAddress__Normalize(self) {
    const patternInput = (FsAddress__get_RowNumber(self.firstAddress) < FsAddress__get_RowNumber(self.lastAddress)) ? [FsAddress__get_RowNumber(self.firstAddress), FsAddress__get_RowNumber(self.lastAddress)] : [FsAddress__get_RowNumber(self.lastAddress), FsAddress__get_RowNumber(self.firstAddress)];
    const patternInput_1 = (FsAddress__get_RowNumber(self.firstAddress) < FsAddress__get_RowNumber(self.lastAddress)) ? [FsAddress__get_RowNumber(self.firstAddress), FsAddress__get_RowNumber(self.lastAddress)] : [FsAddress__get_RowNumber(self.lastAddress), FsAddress__get_RowNumber(self.firstAddress)];
    self._firstAddress = FsAddress_$ctor_Z37302880(patternInput[0], patternInput_1[0]);
    self._lastAddress = FsAddress_$ctor_Z37302880(patternInput[1], patternInput_1[1]);
}

export function FsRangeAddress__get_Range(self) {
    return Range_ofBoundaries(FsAddress__get_Address(self._firstAddress), FsAddress__get_Address(self._lastAddress));
}

export function FsRangeAddress__set_Range_Z721C83C5(self, address) {
    const patternInput = Range_toBoundaries(address);
    self._firstAddress = FsAddress_$ctor_Z721C83C5(patternInput[0]);
    self._lastAddress = FsAddress_$ctor_Z721C83C5(patternInput[1]);
}

export function FsRangeAddress__get_FirstAddress(self) {
    return self._firstAddress;
}

export function FsRangeAddress__get_LastAddress(self) {
    return self._lastAddress;
}

export function FsRangeAddress__Union_6A2513BC(self, rangeAddress) {
    FsRangeAddress__Extend_6D30B323(self, FsRangeAddress__get_FirstAddress(rangeAddress));
    FsRangeAddress__Extend_6D30B323(self, FsRangeAddress__get_LastAddress(rangeAddress));
    return self;
}

