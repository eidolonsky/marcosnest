Sub CreateTable()
    Dim mRange As Range
    Set mRange = ActiveDocument.Range
    mRange.SetRange Start:=ActiveDocument.Range.End, End:=ActiveDocument.Range.End
    Set SelfGenTable = ActiveDocument.Tables.Add(Range:=mRange, NumRows:=2, NumColumns:=8, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed)

    Dim xRange, yRange, zRange As Range
    With SelfGenTable
    
        Set xRange = .Cell(1, 6).Range
             xRange.End = .Cell(1, 8).Range.End
             xRange.Cells.Merge
        Set yRange = .Cell(2, 2).Range
             yRange.End = .Cell(2, 8).Range.End
             yRange.Cells.Merge
        
        SelfGenTable.Cell(Row:=1, Column:=1).Range.InsertAfter "场次"
        SelfGenTable.Cell(Row:=1, Column:=3).Range.InsertAfter "时间"
        SelfGenTable.Cell(Row:=1, Column:=5).Range.InsertAfter "地点"
        SelfGenTable.Cell(Row:=2, Column:=1).Range.InsertAfter "人物"
        
        Set zRange = .Cell(1, 2).Range
            zRange.End = zRange.End - 1
            zRange.Fields.Add zRange, Type:=wdFieldAutoNum, PreserveFormatting:=False
            zRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With
End Sub

