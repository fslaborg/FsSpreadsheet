namespace FsSpreadsheet


/// A FsSparseMatrix
type FsSparseMatrix<'T>(defaultEmptyValue: 'T,sparseValues : 'T array, sparseRowOffsets : int array, ncols:int, columnValues: int array) =     
    member m.NumCols = ncols
    member m.NumRows = sparseRowOffsets.Length - 1
    member m.SparseColumnValues = columnValues
    member m.SparseRowOffsets =  sparseRowOffsets (* nrows + 1 elements *)
    member m.SparseValues =  sparseValues

    member m.MinIndexForRow i = m.SparseRowOffsets.[i]
    member m.MaxIndexForRow i = m.SparseRowOffsets.[i+1]
            

    member m.Item 
        with get (i,j) = 
            let imax = m.NumRows
            let jmax = m.NumCols
            if j < 0 || j >= jmax || i < 0 || i >= imax then raise (new System.ArgumentOutOfRangeException()) else
            let kmin = m.MinIndexForRow i
            let kmax = m.MaxIndexForRow i
            let rec loopRow k =
                (* note: could do a binary chop here *)
                if k >= kmax then defaultEmptyValue else
                let j2 = columnValues.[k]
                if j < j2 then defaultEmptyValue else
                if j = j2 then sparseValues.[k] else 
                loopRow (k+1)
            loopRow kmin


    /// Creates a matrix from a sparse sequence 
    static member init<'T> (emptyValue:'T) (maxi:int) (maxj:int) (s:seq<int*int*'T>) = 

        (* nb. could use sorted dictionary but that is in System.dll *)
        let tab = Array.create maxi null
        let count = ref 0
        for (i,j,v) in s do
            if i < 0 || i >= maxi || j <0 || j >= maxj then failwith "initial value out of range";
            count := !count + 1;
            let tab2 = 
                match tab.[i] with 
                | null -> 
                    let tab2 = new System.Collections.Generic.Dictionary<_,_>(3) 
                    tab.[i] <- tab2;
                    tab2
                | tab2 -> tab2
            tab2.[j] <- v
        // optimize this line....
        let offsA =  
            let rowsAcc = Array.zeroCreate (maxi + 1)
            let mutable acc = 0 
            for i = 0 to maxi-1 do 
                rowsAcc.[i] <- acc;
                acc <- match tab.[i] with 
                        | null -> acc
                        | tab2 -> acc+tab2.Count
            rowsAcc.[maxi] <- acc;
            rowsAcc
            
        let colsA,valsA = 
            let colsAcc = new ResizeArray<_>(!count)
            let valsAcc = new ResizeArray<_>(!count)
            for i = 0 to maxi-1 do 
                match tab.[i] with 
                | null -> ()
                | tab2 -> tab2 |> Seq.toArray |> Array.sortBy (fun kvp -> kvp.Key) |> Array.iter (fun kvp -> colsAcc.Add(kvp.Key); valsAcc.Add(kvp.Value));
            colsAcc.ToArray(), valsAcc.ToArray()

        FsSparseMatrix(emptyValue, sparseValues=valsA, sparseRowOffsets=offsA, ncols=maxj, columnValues=colsA)

    #if FABLE_COMPILER
    #else
    /// Returns the SparseMatrix as Array2D
    static member toArray2D (m:FsSparseMatrix<'T>) =
        Array2D.init m.NumRows m.NumCols (fun i j -> m.[i,j])
    #endif