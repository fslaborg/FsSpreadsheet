import { toFail, printf, toText } from "../fable_modules/fable-library.4.1.3/String.js";
import { create, match } from "../fable_modules/fable-library.4.1.3/RegExp.js";
import { parse } from "../fable_modules/fable-library.4.1.3/Int32.js";
import { toUInt32, fromInt32, fromUInt32, op_Addition, toInt64 } from "../fable_modules/fable-library.4.1.3/BigInt.js";
import { class_type } from "../fable_modules/fable-library.4.1.3/Reflection.js";

/**
 * Transforms excel column string indices (e.g. A, B, Z, AA, CD) to index number (starting with A = 1).
 */
export function CellReference_colAdressToIndex(columnAdress) {
    const length = columnAdress.length | 0;
    let sum = 0;
    for (let i = 0; i <= (length - 1); i++) {
        let c;
        const arg = columnAdress[(length - 1) - i];
        c = arg.toLocaleUpperCase();
        let factor;
        const value = Math.pow(26, i);
        factor = (value >>> 0);
        sum = (sum + ((c.charCodeAt(0) - 64) * factor));
    }
    return sum;
}

/**
 * Transforms number index to excel column string indices (e.g. A, B, Z, AA, CD) (starting with A = 1).
 */
export function CellReference_indexToColAdress(i) {
    const loop = (index_mut, acc_mut) => {
        loop:
        while (true) {
            const index = index_mut, acc = acc_mut;
            if (index === 0) {
                return acc;
            }
            else {
                const mod26 = (index - 1) % 26;
                index_mut = ~~((index - 1) / 26);
                acc_mut = (String.fromCharCode(65 + mod26) + acc);
                continue loop;
            }
            break;
        }
    };
    return loop(i, "");
}

/**
 * Maps 1 based column and row indices to "A1" style reference.
 */
export function CellReference_ofIndices(column, row) {
    const arg = CellReference_indexToColAdress(column);
    return toText(printf("%s%i"))(arg)(row);
}

/**
 * Maps a "A1" style excel cell reference to a column * row index tuple (1 Based indices).
 */
export function CellReference_toIndices(reference) {
    const inp = reference.toLocaleUpperCase();
    const regex = match(create("([A-Z]*)(\\d*)"), inp);
    if (regex != null) {
        const a = regex;
        return [CellReference_colAdressToIndex(a[1] || ""), parse(a[2] || "", 511, true, 32)];
    }
    else {
        return toFail(printf("Reference %s does not match Excel A1-style"))(reference);
    }
}

/**
 * Maps a "A1" style excel cell reference to a column (1 Based indices).
 */
export function CellReference_toColIndex(reference) {
    return CellReference_toIndices(reference)[0];
}

/**
 * Maps a "A1" style excel cell reference to a row (1 Based indices).
 */
export function CellReference_toRowIndex(reference) {
    return CellReference_toIndices(reference)[1];
}

/**
 * Sets the column index (1 Based indices) of a "A1" style excel cell reference.
 */
export function CellReference_setColIndex(colI, reference) {
    return CellReference_ofIndices(colI, CellReference_toIndices(reference)[1]);
}

/**
 * Sets the row index (1 Based indices) of a "A1" style excel cell reference.
 */
export function CellReference_setRowIndex(rowI, reference) {
    return CellReference_ofIndices(CellReference_toIndices(reference)[0], rowI);
}

/**
 * Changes the column portion of a "A1"-style reference by the given amount (positive amount = increase and vice versa).
 */
export function CellReference_moveHorizontal(amount, reference) {
    let value;
    let tupledArg_1;
    const tupledArg = CellReference_toIndices(reference);
    tupledArg_1 = [(value = toInt64(op_Addition(toInt64(fromUInt32(tupledArg[0])), toInt64(fromInt32(amount)))), toUInt32(value) >>> 0), tupledArg[1]];
    return CellReference_ofIndices(tupledArg_1[0], tupledArg_1[1]);
}

/**
 * Changes the row portion of a "A1"-style reference by the given amount (positive amount = increase and vice versa).
 */
export function CellReference_moveVertical(amount, reference) {
    let value;
    let tupledArg_1;
    const tupledArg = CellReference_toIndices(reference);
    tupledArg_1 = [tupledArg[0], (value = toInt64(op_Addition(toInt64(fromUInt32(tupledArg[1])), toInt64(fromInt32(amount)))), toUInt32(value) >>> 0)];
    return CellReference_ofIndices(tupledArg_1[0], tupledArg_1[1]);
}

export class FsAddress {
    constructor(rowNumber, columnNumber, fixedRow, fixedColumn) {
        this._fixedRow = fixedRow;
        this._fixedColumn = fixedColumn;
        this._rowNumber = (rowNumber | 0);
        this._columnNumber = (columnNumber | 0);
        this._trimmedAddress = "";
    }
}

export function FsAddress_$reflection() {
    return class_type("FsSpreadsheet.FsAddress", void 0, FsAddress);
}

export function FsAddress_$ctor_Z4C746FC0(rowNumber, columnNumber, fixedRow, fixedColumn) {
    return new FsAddress(rowNumber, columnNumber, fixedRow, fixedColumn);
}

export function FsAddress_$ctor_Z668C30D9(rowNumber, columnLetter, fixedRow, fixedColumn) {
    return FsAddress_$ctor_Z4C746FC0(rowNumber, ~~CellReference_colAdressToIndex(columnLetter), fixedRow, fixedColumn);
}

export function FsAddress_$ctor_Z37302880(rowNumber, columnNumber) {
    return FsAddress_$ctor_Z4C746FC0(rowNumber, columnNumber, false, false);
}

export function FsAddress_$ctor_Z721C83C5(cellAddressString) {
    const patternInput = CellReference_toIndices(cellAddressString);
    return FsAddress_$ctor_Z37302880(~~patternInput[1], ~~patternInput[0]);
}

export function FsAddress__get_ColumnNumber(self) {
    return self._columnNumber;
}

export function FsAddress__set_ColumnNumber_Z524259A4(self, colI) {
    self._columnNumber = (colI | 0);
}

export function FsAddress__get_RowNumber(self) {
    return self._rowNumber;
}

export function FsAddress__set_RowNumber_Z524259A4(self, rowI) {
    self._rowNumber = (rowI | 0);
}

export function FsAddress__get_Address(self) {
    return CellReference_ofIndices(self._columnNumber >>> 0, self._rowNumber >>> 0);
}

export function FsAddress__set_Address_Z721C83C5(self, address) {
    const patternInput = CellReference_toIndices(address);
    self._rowNumber = (~~patternInput[1] | 0);
    self._columnNumber = (~~patternInput[0] | 0);
}

export function FsAddress__get_FixedRow(self) {
    return false;
}

export function FsAddress__get_FixedColumn(self) {
    return false;
}

/**
 * Creates a deep copy of the FsAddress.
 */
export function FsAddress__Copy(this$) {
    const colNo = FsAddress__get_ColumnNumber(this$) | 0;
    return FsAddress_$ctor_Z4C746FC0(FsAddress__get_RowNumber(this$), colNo, FsAddress__get_FixedRow(this$), FsAddress__get_FixedColumn(this$));
}

/**
 * Returns a deep copy of the given FsAddress.
 */
export function FsAddress_copy_6D30B323(address) {
    return FsAddress__Copy(address);
}

/**
 * Updates the row- and columnIndex respective to the given indices.
 */
export function FsAddress__UpdateIndices_Z37302880(self, rowIndex, colIndex) {
    self._columnNumber = (colIndex | 0);
    self._rowNumber = (rowIndex | 0);
}

/**
 * Updates the row- and columnIndex of a given FsAddress respective to the given indices.
 */
export function FsAddress_updateIndices(rowIndex, colIndex, address) {
    FsAddress__UpdateIndices_Z37302880(address, rowIndex, colIndex);
    return address;
}

/**
 * Returns the row- and the columnIndex of the FsAddress.
 */
export function FsAddress__ToIndices(self) {
    return [self._rowNumber, self._columnNumber];
}

/**
 * Returns the row- and the columnIndex of a given FsAddress.
 */
export function FsAddress_toIndices_6D30B323(address) {
    return FsAddress__ToIndices(address);
}

/**
 * Compares the FsAddress with a given other one.
 */
export function FsAddress__Compare_6D30B323(self, address) {
    if ((((FsAddress__get_Address(self) === FsAddress__get_Address(address)) && (FsAddress__get_ColumnNumber(self) === FsAddress__get_ColumnNumber(address))) && (FsAddress__get_RowNumber(self) === FsAddress__get_RowNumber(address))) && (FsAddress__get_FixedColumn(self) === FsAddress__get_FixedColumn(address))) {
        return FsAddress__get_FixedRow(self) === FsAddress__get_FixedRow(address);
    }
    else {
        return false;
    }
}

/**
 * Checks if 2 FsAddresses are equal.
 */
export function FsAddress_compare(address1, address2) {
    return FsAddress__Compare_6D30B323(address1, address2);
}

