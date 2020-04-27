Attribute VB_Name = "Module4"
Sub bullet4()
With ActiveWindow.Selection.TextRange
    .IndentLevel = 4
With .Paragraphs.ParagraphFormat.bullet
.Visible = True
        .RelativeSize = 1
        .Character = 2013
            With .Font
            .Name = "Wingdings"
            .Color.RGB = RGB(127, 127, 127)
            End With
End With
End With
End Sub
