Attribute VB_Name = "Module1"
Sub remove_bullet_periods()
    Dim osld As Slide
    Dim oshp As Shape
    Dim otxR As TextRange
    Dim i As Integer
    Set osld = Application.ActiveWindow.View.Slide
    On Error Resume Next
    With osld
        For Each oshp In osld.Shapes
            If oshp.HasTextFrame Then
                With oshp.TextFrame.TextRange
                    For i = 1 To .Paragraphs.Count
                        If .Paragraphs(i).ParagraphFormat.Bullet = True Then
                            Set otxR = .Paragraphs(i)
                                If InStrRev(otxR, ".") > Len(otxR) - 3 Then
                                    otxR.Characters(InStrRev(otxR, ".")).Delete
                                End If
                                If InStrRev(otxR, "¡£") > Len(otxR) - 3 Then
                                    otxR.Characters(InStrRev(otxR, "¡£")).Delete
                                End If
                        End If
                    Next i
                End With
            End If
        Next oshp
    End With
End Sub
