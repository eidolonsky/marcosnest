Attribute VB_Name = "Module3"
Sub bullet3()
With ActiveWindow.Selection.TextRange
    .IndentLevel = 3
With .Paragraphs.ParagraphFormat.bullet
.Visible = True
        .RelativeSize = 0.6
        .Character = 9658
            With .Font
            .Name = "Monotype Corsiva"
            .Color.RGB = RGB(127, 127, 127)
            End With
End With
End With
End Sub
