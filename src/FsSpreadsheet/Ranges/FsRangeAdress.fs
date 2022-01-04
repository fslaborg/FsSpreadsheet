namespace rec FsSpreadsheet

type FsRangeAddress(worksheet : FsWorksheet, rangeAddress : string) =

    let mutable _address = rangeAddress

    member self.Address 
        with get() = _address
        and set(address) = _address <- address