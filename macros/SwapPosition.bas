Attribute VB_Name = "SwapPosition"
Sub SwapPosition()
    Dim l1, t1, l2, t2 As Long
    
    With ActiveWindow.Selection.ShapeRange
        If .Count <> 2 Then
            MsgBox "Please select two shapes"
            
        Else
            l1 = .Item(1).Left
            t1 = .Item(1).Top
            l2 = .Item(2).Left
            t2 = .Item(2).Top
            
            .Item(1).Left = l2
            .Item(1).Top = t2
            
            .Item(2).Left = l1
            .Item(2).Top = t1
        End If
    End With
End Sub
