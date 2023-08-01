import { Xlsx } from "./Xlsx.js";

export function FsSpreadsheet_FsWorkbook__FsWorkbook_fromXlsxFile_Static_Z721C83C5(path) {
    return Xlsx.fromXlsxFile(path);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_fromXlsxStream_Static_4D976C1A(stream) {
    return Xlsx.fromXlsxStream(stream);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_fromBytes_Static_Z3F6BC7B1(bytes) {
    return Xlsx.fromBytes(bytes);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_toFile_Static(path, wb) {
    return Xlsx.toFile(path, wb);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_toStream_Static(stream, wb) {
    return Xlsx.toStream(stream, wb);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_toBytes_Static_32154C9D(wb) {
    return Xlsx.toBytes(wb);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_ToFile_Z721C83C5(this$, path) {
    return FsSpreadsheet_FsWorkbook__FsWorkbook_toFile_Static(path, this$);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_ToStream_4D976C1A(this$, stream) {
    return FsSpreadsheet_FsWorkbook__FsWorkbook_toStream_Static(stream, this$);
}

export function FsSpreadsheet_FsWorkbook__FsWorkbook_ToBytes(this$) {
    return FsSpreadsheet_FsWorkbook__FsWorkbook_toBytes_Static_32154C9D(this$);
}

