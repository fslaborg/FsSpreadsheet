import { class_type } from "../../fable_modules/fable-library.4.1.3/Reflection.js";

export class OptionalSource$1 {
    constructor(s) {
        this.s = s;
    }
}

export function OptionalSource$1_$reflection(gen0) {
    return class_type("FsSpreadsheet.DSL.Expression.OptionalSource`1", [gen0], OptionalSource$1);
}

export function OptionalSource$1_$ctor_2B595(s) {
    return new OptionalSource$1(s);
}

export function OptionalSource$1__get_Source(this$) {
    return this$.s;
}

export class RequiredSource$1 {
    constructor(s) {
        this.s = s;
    }
}

export function RequiredSource$1_$reflection(gen0) {
    return class_type("FsSpreadsheet.DSL.Expression.RequiredSource`1", [gen0], RequiredSource$1);
}

export function RequiredSource$1_$ctor_2B595(s) {
    return new RequiredSource$1(s);
}

export function RequiredSource$1__get_Source(this$) {
    return this$.s;
}

export class ExpressionSource$1 {
    constructor(s) {
        this.s = s;
    }
}

export function ExpressionSource$1_$reflection(gen0) {
    return class_type("FsSpreadsheet.DSL.Expression.ExpressionSource`1", [gen0], ExpressionSource$1);
}

export function ExpressionSource$1_$ctor_2B595(s) {
    return new ExpressionSource$1(s);
}

export function ExpressionSource$1__get_Source(this$) {
    return this$.s;
}

