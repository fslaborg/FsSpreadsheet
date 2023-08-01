import { toString as toString_1, Union } from "../../fable_modules/fable-library.4.1.3/Types.js";
import { class_type, union_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { FsAddress__Copy, CellReference_indexToColAdress, FsAddress__get_ColumnNumber, FsAddress__get_RowNumber, FsAddress_$ctor_Z37302880 } from "../FsAddress.js";
import { int64ToString, int16ToString, int32ToString } from "../../fable_modules/fable-library.4.1.3/Util.js";
import { parse as parse_4, toString } from "../../fable_modules/fable-library.4.1.3/Decimal.js";
import Decimal from "../../fable_modules/fable-library.4.1.3/Decimal.js";
import { parse } from "../../fable_modules/fable-library.4.1.3/Boolean.js";
import { parse as parse_1 } from "../../fable_modules/fable-library.4.1.3/Double.js";
import { parse as parse_2 } from "../../fable_modules/fable-library.4.1.3/Int32.js";
import { toUInt64, toInt64 } from "../../fable_modules/fable-library.4.1.3/BigInt.js";
import { parse as parse_3 } from "../../fable_modules/fable-library.4.1.3/Long.js";
import { parse as parse_5 } from "../../fable_modules/fable-library.4.1.3/Date.js";
import { parse as parse_6 } from "../../fable_modules/fable-library.4.1.3/Guid.js";
import { parse as parse_7 } from "../../fable_modules/fable-library.4.1.3/Char.js";
import { map, defaultArg } from "../../fable_modules/fable-library.4.1.3/Option.js";

export class DataType extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["String", "Boolean", "Number", "Date", "Empty"];
    }
}

export function DataType_$reflection() {
    return union_type("FsSpreadsheet.DataType", [], DataType, () => [[], [], [], [], []]);
}

export class FsCell {
    constructor(value, dataType, address) {
        this._cellValue = toString_1(value);
        this._dataType = defaultArg(dataType, new DataType(0, []));
        this._comment = "";
        this._hyperlink = "";
        this._richText = "";
        this._formulaA1 = "";
        this._formulaR1C1 = "";
        this._rowIndex = (defaultArg(map(FsAddress__get_RowNumber, address), 0) | 0);
        this._columnIndex = (defaultArg(map(FsAddress__get_ColumnNumber, address), 0) | 0);
    }
    static empty() {
        return new FsCell("", new DataType(4, []), FsAddress_$ctor_Z37302880(0, 0));
    }
    get Value() {
        const self = this;
        return self._cellValue;
    }
    set Value(value) {
        const self = this;
        self._cellValue = value;
    }
    get DataType() {
        const self = this;
        return self._dataType;
    }
    set DataType(dataType) {
        const self = this;
        self._dataType = dataType;
    }
    get ColumnNumber() {
        const self = this;
        return self._columnIndex | 0;
    }
    set ColumnNumber(colI) {
        const self = this;
        self._columnIndex = (colI | 0);
    }
    get RowNumber() {
        const self = this;
        return self._rowIndex | 0;
    }
    set RowNumber(rowI) {
        const self = this;
        self._rowIndex = (rowI | 0);
    }
    get Address() {
        const self = this;
        return FsAddress_$ctor_Z37302880(self._rowIndex, self._columnIndex);
    }
    set Address(address) {
        const self = this;
        self._rowIndex = (FsAddress__get_RowNumber(address) | 0);
        self._columnIndex = (FsAddress__get_ColumnNumber(address) | 0);
    }
    static create(rowNumber, colNumber, value) {
        let patternInput;
        const value_1_1 = value;
        patternInput = ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : ((typeof value_1_1 === "boolean") ? (value_1_1 ? [new DataType(1, []), "True"] : [new DataType(1, []), "False"]) : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int32ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int16ToString(value_1_1)] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), int64ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((value_1_1 instanceof Decimal) ? [new DataType(2, []), toString(value_1_1)] : ((value_1_1 instanceof Date) ? [new DataType(3, []), toString_1(value_1_1)] : ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : [new DataType(0, []), toString_1(value_1_1)])))))))))))))));
        return new FsCell(patternInput[1], patternInput[0], FsAddress_$ctor_Z37302880(rowNumber, colNumber));
    }
    static createEmpty() {
        return new FsCell("", new DataType(4, []), FsAddress_$ctor_Z37302880(0, 0));
    }
    static createWithAdress(adress, value) {
        let patternInput;
        const value_1_1 = value;
        patternInput = ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : ((typeof value_1_1 === "boolean") ? (value_1_1 ? [new DataType(1, []), "True"] : [new DataType(1, []), "False"]) : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int32ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int16ToString(value_1_1)] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), int64ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((value_1_1 instanceof Decimal) ? [new DataType(2, []), toString(value_1_1)] : ((value_1_1 instanceof Date) ? [new DataType(3, []), toString_1(value_1_1)] : ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : [new DataType(0, []), toString_1(value_1_1)])))))))))))))));
        return new FsCell(patternInput[1], patternInput[0], adress);
    }
    static createEmptyWithAdress(adress) {
        return new FsCell("", new DataType(4, []), adress);
    }
    static createWithDataType(dataType, rowNumber, colNumber, value) {
        return new FsCell(value, dataType, FsAddress_$ctor_Z37302880(rowNumber, colNumber));
    }
    toString() {
        const self = this;
        return `${CellReference_indexToColAdress(self.ColumnNumber >>> 0)}${self.RowNumber} : ${self.Value} | ${self.DataType}`;
    }
    CopyFrom(otherCell) {
        const self = this;
        self.DataType = otherCell.DataType;
        self.Value = otherCell.Value;
    }
    CopyTo(target) {
        const self = this;
        target.DataType = self.DataType;
        target.Value = self.Value;
    }
    static copyFromTo(sourceCell, targetCell) {
        targetCell.DataType = sourceCell.DataType;
        targetCell.Value = sourceCell.Value;
        return targetCell;
    }
    Copy() {
        const self = this;
        return new FsCell(self.Value, self.DataType, FsAddress__Copy(self.Address));
    }
    static copy(cell) {
        return cell.Copy();
    }
    ValueAsString() {
        const self = this;
        return self.Value;
    }
    static getValueAsString(cell) {
        return cell.ValueAsString();
    }
    ValueAsBool() {
        const self = this;
        return parse(self.Value);
    }
    static getValueAsBool(cell) {
        return cell.ValueAsBool();
    }
    ValueAsFloat() {
        const self = this;
        return parse_1(self.Value);
    }
    static getValueAsFloat(cell) {
        return cell.ValueAsFloat();
    }
    ValueAsInt() {
        const self = this;
        return parse_2(self.Value, 511, false, 32) | 0;
    }
    static getValueAsInt(cell) {
        return cell.ValueAsInt();
    }
    ValueAsUInt() {
        const self = this;
        return parse_2(self.Value, 511, true, 32);
    }
    static getValueAsUInt(cell) {
        return cell.ValueAsUInt();
    }
    ValueAsLong() {
        const self = this;
        return toInt64(parse_3(self.Value, 511, false, 64));
    }
    static getValueAsLong(cell) {
        return cell.ValueAsLong();
    }
    ValueAsULong() {
        const self = this;
        return toUInt64(parse_3(self.Value, 511, true, 64));
    }
    static getValueAsULong(cell) {
        return cell.ValueAsULong();
    }
    ValueAsDouble() {
        const self = this;
        return parse_1(self.Value);
    }
    static getValueAsDouble(cell) {
        return cell.ValueAsDouble();
    }
    ValueAsDecimal() {
        const self = this;
        return parse_4(self.Value);
    }
    static getValueAsDecimal(cell) {
        return cell.ValueAsDecimal();
    }
    ValueAsDateTime() {
        const self = this;
        return parse_5(self.Value);
    }
    static getValueAsDateTime(cell) {
        return cell.ValueAsDateTime();
    }
    ValueAsGuid() {
        const self = this;
        return parse_6(self.Value);
    }
    static getValueAsGuid(cell) {
        return cell.ValueAsGuid();
    }
    ValueAsChar() {
        const self = this;
        return parse_7(self.Value);
    }
    static getValueAsChar(cell) {
        return cell.ValueAsChar();
    }
    SetValueAs(value) {
        const self = this;
        let patternInput;
        const value_1_1 = value;
        patternInput = ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : ((typeof value_1_1 === "boolean") ? (value_1_1 ? [new DataType(1, []), "True"] : [new DataType(1, []), "False"]) : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int32ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), int16ToString(value_1_1)] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), int64ToString(value_1_1)] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "bigint") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((typeof value_1_1 === "number") ? [new DataType(2, []), value_1_1.toString()] : ((value_1_1 instanceof Decimal) ? [new DataType(2, []), toString(value_1_1)] : ((value_1_1 instanceof Date) ? [new DataType(3, []), toString_1(value_1_1)] : ((typeof value_1_1 === "string") ? [new DataType(0, []), value_1_1] : [new DataType(0, []), toString_1(value_1_1)])))))))))))))));
        self._dataType = patternInput[0];
        self._cellValue = patternInput[1];
    }
    static setValueAs(value, cell) {
        cell.SetValueAs(value);
        return cell;
    }
}

export function FsCell_$reflection() {
    return class_type("FsSpreadsheet.FsCell", void 0, FsCell);
}

export function FsCell_$ctor_Z44506C7(value, dataType, address) {
    return new FsCell(value, dataType, address);
}

