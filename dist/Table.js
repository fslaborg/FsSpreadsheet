import { tail, mapIndexed, collect, map, delay, toArray } from "./fable_modules/fable-library.4.1.3/Seq.js";
import { FsRangeRow__Cells_Z2740B3CA } from "./FsSpreadsheet/Ranges/FsRangeRow.js";
import { sortBy, map as map_1 } from "./fable_modules/fable-library.4.1.3/Array.js";
import { Array_groupBy } from "./fable_modules/fable-library.4.1.3/Seq2.js";
import { printf, toText } from "./fable_modules/fable-library.4.1.3/String.js";
import { comparePrimitives, numberHash } from "./fable_modules/fable-library.4.1.3/Util.js";
import { FsRangeAddress__get_Range } from "./FsSpreadsheet/Ranges/FsRangeAddress.js";
import { FsRangeBase__get_RangeAddress } from "./FsSpreadsheet/Ranges/FsRangeBase.js";

export function fromFsTable(fscellcollection, fsTable) {
    const columns = fsTable.ShowHeaderRow ? toArray(delay(() => map((headerCell) => ({
        name: headerCell.Value,
    }), FsRangeRow__Cells_Z2740B3CA(fsTable.HeadersRow(), fscellcollection)))) : [];
    const rows = map_1((tupledArg) => map_1((tuple_2) => tuple_2[1], tupledArg[1]), sortBy((tuple_1) => tuple_1[0], Array_groupBy((tuple) => tuple[0], toArray(delay(() => collect((col) => mapIndexed((i, c) => {
        let matchValue, msg;
        return [i, (matchValue = c.DataType, (matchValue.tag === 1) ? c.ValueAsBool() : ((matchValue.tag === 2) ? c.ValueAsFloat() : ((matchValue.tag === 3) ? c.ValueAsDateTime() : ((matchValue.tag === 0) ? c.Value : ((msg = toText(printf("ValueType \'%A\' is not fully implemented in FsSpreadsheet and is handled as string input."))(matchValue), (console.log(msg), c.Value)))))))];
    }, fsTable.ShowHeaderRow ? tail(col.Cells) : col.Cells), fsTable.GetColumns(fscellcollection)))), {
        Equals: (x, y) => (x === y),
        GetHashCode: numberHash,
    }), {
        Compare: comparePrimitives,
    }));
    return {
        name: fsTable.Name,
        ref: FsRangeAddress__get_Range(FsRangeBase__get_RangeAddress(fsTable)),
        columns: columns,
        rows: rows,
        headerRow: fsTable.ShowHeaderRow,
    };
}

