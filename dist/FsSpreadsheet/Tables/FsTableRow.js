import { FsRangeRow_$reflection, FsRangeRow } from "../Ranges/FsRangeRow.js";
import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";

export class FsTableRow extends FsRangeRow {
    constructor(rangeAddress) {
        super(rangeAddress);
    }
}

export function FsTableRow_$reflection() {
    return class_type("FsSpreadsheet.FsTableRow", void 0, FsTableRow, FsRangeRow_$reflection());
}

export function FsTableRow_$ctor_6A2513BC(rangeAddress) {
    return new FsTableRow(rangeAddress);
}

