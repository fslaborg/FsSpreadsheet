namespace FsSpreadsheet.Interactive

open System
open System.Threading.Tasks
open Microsoft.DotNet.Interactive
open Microsoft.DotNet.Interactive.Formatting
open FsSpreadsheet

type FormatterKernelExtension() =

    interface IKernelExtension with
        member _.OnLoadAsync _ =

            Formatter.Register<FsWorkbook>(
                Action<_, _>(fun workbook (writer: IO.TextWriter) -> writer.Write(Formatters.formatWorkbook workbook)),
                "text/html"
            )

            Formatter.Register<FsWorksheet>(
                Action<_, _>(fun worksheet (writer: IO.TextWriter) -> writer.Write(Formatters.formatWorksheet worksheet)),
                "text/html"
            )

            Task.CompletedTask
