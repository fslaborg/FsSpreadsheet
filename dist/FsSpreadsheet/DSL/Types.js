import { Union } from "../../fable_modules/fable-library.4.1.3/Types.js";
import { tuple_type, int32_type, list_type, union_type, class_type, string_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { empty, pick, exists, map, reduce } from "../../fable_modules/fable-library.4.1.3/List.js";
import { printf, toConsole } from "../../fable_modules/fable-library.4.1.3/String.js";
import { DataType_$reflection } from "../Cells/FsCell.js";

export class Message extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["Text", "Exception"];
    }
}

export function Message_$reflection() {
    return union_type("FsSpreadsheet.DSL.Message", [], Message, () => [[["Item", string_type]], [["Item", class_type("System.Exception")]]]);
}

export function Message_message_Z721C83C5(s) {
    return new Message(0, [s]);
}

export function Message_message_2BC701FD(e) {
    return new Message(1, [e]);
}

export function Message__MapText_11D407F6(this$, m) {
    if (this$.tag === 1) {
        return this$;
    }
    else {
        return new Message(0, [m(this$.fields[0])]);
    }
}

export function Message__AsString(this$) {
    if (this$.tag === 1) {
        return this$.fields[0].message;
    }
    else {
        return this$.fields[0];
    }
}

export function Message__TryText(this$) {
    if (this$.tag === 0) {
        return this$.fields[0];
    }
    else {
        return void 0;
    }
}

export function Message__TryException(this$) {
    if (this$.tag === 1) {
        return this$.fields[0];
    }
    else {
        return void 0;
    }
}

export function Message__get_IsTxt(this$) {
    if (this$.tag === 0) {
        return true;
    }
    else {
        return false;
    }
}

export function Message__get_IsExc(this$) {
    if (this$.tag === 0) {
        return true;
    }
    else {
        return false;
    }
}

export function Messages_format(ms) {
    return reduce((a, b) => ((a + ";") + b), map(Message__AsString, ms));
}

export function Messages_fail(ms) {
    const s = Messages_format(ms);
    if (exists(Message__get_IsExc, ms)) {
        toConsole(printf("s"));
        throw pick(Message__TryException, ms);
    }
    else {
        throw new Error(s);
    }
}

export class SheetEntity$1 extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["Some", "NoneOptional", "NoneRequired"];
    }
}

export function SheetEntity$1_$reflection(gen0) {
    return union_type("FsSpreadsheet.DSL.SheetEntity`1", [gen0], SheetEntity$1, () => [[["Item1", gen0], ["Item2", list_type(Message_$reflection())]], [["Item", list_type(Message_$reflection())]], [["Item", list_type(Message_$reflection())]]]);
}

export function SheetEntity$1_some_2B595(v) {
    return new SheetEntity$1(0, [v, empty()]);
}

/**
 * Get messages
 */
export function SheetEntity$1__get_Messages(this$) {
    switch (this$.tag) {
        case 1:
            return this$.fields[0];
        case 2:
            return this$.fields[0];
        default:
            return this$.fields[1];
    }
}

export class ColumnIndex extends Union {
    constructor(Item) {
        super();
        this.tag = 0;
        this.fields = [Item];
    }
    cases() {
        return ["Col"];
    }
}

export function ColumnIndex_$reflection() {
    return union_type("FsSpreadsheet.DSL.ColumnIndex", [], ColumnIndex, () => [[["Item", int32_type]]]);
}

export function ColumnIndex__get_Index(self) {
    return self.fields[0];
}

export class RowIndex extends Union {
    constructor(Item) {
        super();
        this.tag = 0;
        this.fields = [Item];
    }
    cases() {
        return ["Row"];
    }
}

export function RowIndex_$reflection() {
    return union_type("FsSpreadsheet.DSL.RowIndex", [], RowIndex, () => [[["Item", int32_type]]]);
}

export function RowIndex__get_Index(self) {
    return self.fields[0];
}

export class ColumnElement extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["IndexedCell", "UnindexedCell"];
    }
}

export function ColumnElement_$reflection() {
    return union_type("FsSpreadsheet.DSL.ColumnElement", [], ColumnElement, () => [[["Item1", RowIndex_$reflection()], ["Item2", tuple_type(DataType_$reflection(), string_type)]], [["Item", tuple_type(DataType_$reflection(), string_type)]]]);
}

export class RowElement extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["IndexedCell", "UnindexedCell"];
    }
}

export function RowElement_$reflection() {
    return union_type("FsSpreadsheet.DSL.RowElement", [], RowElement, () => [[["Item1", ColumnIndex_$reflection()], ["Item2", tuple_type(DataType_$reflection(), string_type)]], [["Item", tuple_type(DataType_$reflection(), string_type)]]]);
}

export class TableElement extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["UnindexedRow", "UnindexedColumn"];
    }
}

export function TableElement_$reflection() {
    return union_type("FsSpreadsheet.DSL.TableElement", [], TableElement, () => [[["Item", list_type(RowElement_$reflection())]], [["Item", list_type(ColumnElement_$reflection())]]]);
}

export function TableElement__get_IsRow(this$) {
    if (this$.tag === 0) {
        return true;
    }
    else {
        return false;
    }
}

export function TableElement__get_IsColumn(this$) {
    if (this$.tag === 1) {
        return true;
    }
    else {
        return false;
    }
}

export class SheetElement extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["Table", "IndexedRow", "UnindexedRow", "IndexedColumn", "UnindexedColumn", "IndexedCell", "UnindexedCell"];
    }
}

export function SheetElement_$reflection() {
    return union_type("FsSpreadsheet.DSL.SheetElement", [], SheetElement, () => [[["Item1", string_type], ["Item2", list_type(TableElement_$reflection())]], [["Item1", RowIndex_$reflection()], ["Item2", list_type(RowElement_$reflection())]], [["Item", list_type(RowElement_$reflection())]], [["Item1", ColumnIndex_$reflection()], ["Item2", list_type(ColumnElement_$reflection())]], [["Item", list_type(ColumnElement_$reflection())]], [["Item1", RowIndex_$reflection()], ["Item2", ColumnIndex_$reflection()], ["Item3", tuple_type(DataType_$reflection(), string_type)]], [["Item", tuple_type(DataType_$reflection(), string_type)]]]);
}

export class WorkbookElement extends Union {
    constructor(tag, fields) {
        super();
        this.tag = tag;
        this.fields = fields;
    }
    cases() {
        return ["UnnamedSheet", "NamedSheet"];
    }
}

export function WorkbookElement_$reflection() {
    return union_type("FsSpreadsheet.DSL.WorkbookElement", [], WorkbookElement, () => [[["Item", list_type(SheetElement_$reflection())]], [["Item1", string_type], ["Item2", list_type(SheetElement_$reflection())]]]);
}

export class Workbook extends Union {
    constructor(Item) {
        super();
        this.tag = 0;
        this.fields = [Item];
    }
    cases() {
        return ["Workbook"];
    }
}

export function Workbook_$reflection() {
    return union_type("FsSpreadsheet.DSL.Workbook", [], Workbook, () => [[["Item", list_type(WorkbookElement_$reflection())]]]);
}

