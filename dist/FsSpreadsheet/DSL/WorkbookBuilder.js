import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";
import { SheetEntity$1__get_Messages, SheetEntity$1, Message__MapText_11D407F6, SheetEntity$1_some_2B595 } from "./Types.js";
import { append, map, empty } from "../../fable_modules/fable-library.4.1.3/List.js";
import { printf, toText } from "../../fable_modules/fable-library.4.1.3/String.js";

export class WorkbookBuilder {
    constructor() {
    }
}

export function WorkbookBuilder_$reflection() {
    return class_type("FsSpreadsheet.DSL.WorkbookBuilder", void 0, WorkbookBuilder);
}

export function WorkbookBuilder_$ctor() {
    return new WorkbookBuilder();
}

export function WorkbookBuilder_get_Empty() {
    return SheetEntity$1_some_2B595(empty());
}

export function WorkbookBuilder__SignMessages_3F89FFFC(this$, messages) {
    return map((m) => {
        let clo;
        return Message__MapText_11D407F6(m, (clo = toText(printf("In Workbook: %s")), clo));
    }, messages);
}

export function WorkbookBuilder__Combine_49CE2EC0(this$, wx1, wx2) {
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

