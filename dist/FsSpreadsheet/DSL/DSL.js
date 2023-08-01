import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { SheetEntity$1 } from "./Types.js";
import { singleton } from "../../fable_modules/fable-library.4.1.3/List.js";

export class DSL {
    constructor() {
    }
}

export function DSL_$reflection() {
    return class_type("FsSpreadsheet.DSL.DSL", void 0, DSL);
}

/**
 * Transforms any given missing element to an optional.
 */
export function DSL_opt_Z6AB9374C(elem) {
    switch (elem.tag) {
        case 1:
            return new SheetEntity$1(1, [elem.fields[0]]);
        case 2:
            return new SheetEntity$1(1, [elem.fields[0]]);
        default:
            return elem;
    }
}

/**
 * Drops the cell with the given message
 */
export function DSL_dropCell_6CC5727E(message) {
    return new SheetEntity$1(2, [singleton(message)]);
}

/**
 * Drops the row with the given message
 */
export function DSL_dropRow_6CC5727E(message) {
    return new SheetEntity$1(2, [singleton(message)]);
}

/**
 * Drops the column with the given message
 */
export function DSL_dropColumn_6CC5727E(message) {
    return new SheetEntity$1(2, [singleton(message)]);
}

/**
 * Drops the sheet with the given message
 */
export function DSL_dropSheet_6CC5727E(message) {
    return new SheetEntity$1(2, [singleton(message)]);
}

/**
 * Drops the workbook with the given message
 */
export function DSL_dropWorkbook_6CC5727E(message) {
    return new SheetEntity$1(2, [singleton(message)]);
}

