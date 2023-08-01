import { Union } from "../../fable_modules/fable-library.4.1.3/Types.js";
import { class_type, union_type, char_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { DataType } from "../Cells/FsCell.js";
import { head, tail, isEmpty, append, empty, map, reduce } from "../../fable_modules/fable-library.4.1.3/List.js";
import { SheetEntity$1__get_Messages, Message__MapText_11D407F6, SheetEntity$1 } from "./Types.js";
import { printf, toText } from "../../fable_modules/fable-library.4.1.3/String.js";
import { OptionalSource$1__get_Source, RequiredSource$1__get_Source, OptionalSource$1_$ctor_2B595, RequiredSource$1_$ctor_2B595 } from "./Expression.js";

export class ReduceOperation extends Union {
    constructor(Item) {
        super();
        this.tag = 0;
        this.fields = [Item];
    }
    cases() {
        return ["Concat"];
    }
}

export function ReduceOperation_$reflection() {
    return union_type("FsSpreadsheet.DSL.ReduceOperation", [], ReduceOperation, () => [[["Item", char_type]]]);
}

export function ReduceOperation__Reduce_Z26A4E3C2(this$, values) {
    return [new DataType(0, []), reduce((a, b) => (`${a}${this$.fields[0]}${b}`), map((arg) => arg[1], values))];
}

export class CellBuilder {
    constructor() {
        this.reducer = (new ReduceOperation(","));
    }
}

export function CellBuilder_$reflection() {
    return class_type("FsSpreadsheet.DSL.CellBuilder", void 0, CellBuilder);
}

export function CellBuilder_$ctor() {
    return new CellBuilder();
}

export function CellBuilder_get_Empty() {
    return new SheetEntity$1(1, [empty()]);
}

export function CellBuilder__SignMessages_3F89FFFC(this$, messages) {
    return map((m) => {
        let clo;
        return Message__MapText_11D407F6(m, (clo = toText(printf("In Cell: %s")), clo));
    }, messages);
}

export function CellBuilder__Yield_36B9E420(this$, ro) {
    this$.reducer = ro;
    return new SheetEntity$1(1, [empty()]);
}

export function CellBuilder__Combine_4F8B6A0(this$, wx1, wx2) {
    let matchResult, l1, l2, messages1, messages2, messages2_1, messages1_1, f1, messages1_2, messages2_2, f2, messages1_3, messages2_3, messages1_4, messages2_4;
    switch (wx1.tag) {
        case 2: {
            if (wx2.tag === 2) {
                matchResult = 1;
                messages2_1 = wx2.fields[0];
            }
            else {
                matchResult = 2;
                messages1_1 = wx1.fields[0];
            }
            break;
        }
        case 1: {
            switch (wx2.tag) {
                case 0: {
                    matchResult = 4;
                    f2 = wx2.fields[0];
                    messages1_3 = wx1.fields[0];
                    messages2_3 = wx2.fields[1];
                    break;
                }
                case 1: {
                    matchResult = 5;
                    messages1_4 = wx1.fields[0];
                    messages2_4 = wx2.fields[0];
                    break;
                }
                default: {
                    matchResult = 1;
                    messages2_1 = wx2.fields[0];
                }
            }
            break;
        }
        default:
            switch (wx2.tag) {
                case 2: {
                    matchResult = 1;
                    messages2_1 = wx2.fields[0];
                    break;
                }
                case 1: {
                    matchResult = 3;
                    f1 = wx1.fields[0];
                    messages1_2 = wx1.fields[1];
                    messages2_2 = wx2.fields[0];
                    break;
                }
                default: {
                    matchResult = 0;
                    l1 = wx1.fields[0];
                    l2 = wx2.fields[0];
                    messages1 = wx1.fields[1];
                    messages2 = wx2.fields[1];
                }
            }
    }
    switch (matchResult) {
        case 0:
            return new SheetEntity$1(0, [append(l1, l2), append(messages1, messages2)]);
        case 1:
            return new SheetEntity$1(2, [append(SheetEntity$1__get_Messages(wx1), messages2_1)]);
        case 2:
            return new SheetEntity$1(2, [append(messages1_1, SheetEntity$1__get_Messages(wx2))]);
        case 3:
            return new SheetEntity$1(0, [f1, append(messages1_2, messages2_2)]);
        case 4:
            return new SheetEntity$1(0, [f2, append(messages1_3, messages2_3)]);
        default:
            return new SheetEntity$1(1, [append(messages1_4, messages2_4)]);
    }
}

export function CellBuilder__Combine_2C6C5E1E(this$, wx1, wx2) {
    return RequiredSource$1_$ctor_2B595(wx2);
}

export function CellBuilder__Combine_Z40A4EDE2(this$, wx1, wx2) {
    return RequiredSource$1_$ctor_2B595(wx1);
}

export function CellBuilder__Combine_3F9D1279(this$, wx1, wx2) {
    return OptionalSource$1_$ctor_2B595(wx2);
}

export function CellBuilder__Combine_7AD05659(this$, wx1, wx2) {
    return OptionalSource$1_$ctor_2B595(wx1);
}

export function CellBuilder__Combine_Z674DE191(this$, wx1, wx2) {
    return RequiredSource$1_$ctor_2B595(CellBuilder__Combine_4F8B6A0(this$, RequiredSource$1__get_Source(wx1), wx2));
}

export function CellBuilder__Combine_1F389C8F(this$, wx1, wx2) {
    return RequiredSource$1_$ctor_2B595(CellBuilder__Combine_4F8B6A0(this$, wx1, RequiredSource$1__get_Source(wx2)));
}

export function CellBuilder__Combine_512E85C8(this$, wx1, wx2) {
    return OptionalSource$1_$ctor_2B595(CellBuilder__Combine_4F8B6A0(this$, OptionalSource$1__get_Source(wx1), wx2));
}

export function CellBuilder__Combine_6390E688(this$, wx1, wx2) {
    return OptionalSource$1_$ctor_2B595(CellBuilder__Combine_4F8B6A0(this$, wx1, OptionalSource$1__get_Source(wx2)));
}

export function CellBuilder__AsCellElement_825BA8D(this$, children) {
    let matchResult, messages, v, messages_1, vals, messages_2, messages_3;
    switch (children.tag) {
        case 2: {
            matchResult = 2;
            messages_2 = children.fields[0];
            break;
        }
        case 1: {
            matchResult = 3;
            messages_3 = children.fields[0];
            break;
        }
        default:
            if (!isEmpty(children.fields[0])) {
                if (isEmpty(tail(children.fields[0]))) {
                    matchResult = 0;
                    messages = children.fields[1];
                    v = head(children.fields[0]);
                }
                else {
                    matchResult = 1;
                    messages_1 = children.fields[1];
                    vals = children.fields[0];
                }
            }
            else {
                matchResult = 1;
                messages_1 = children.fields[1];
                vals = children.fields[0];
            }
    }
    switch (matchResult) {
        case 0:
            return new SheetEntity$1(0, [[v, void 0], messages]);
        case 1:
            return new SheetEntity$1(0, [[ReduceOperation__Reduce_Z26A4E3C2(this$.reducer, vals), void 0], messages_1]);
        case 2:
            return new SheetEntity$1(2, [messages_2]);
        default:
            return new SheetEntity$1(1, [messages_3]);
    }
}

