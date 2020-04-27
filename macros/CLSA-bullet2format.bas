Attribute VB_Name = "Module2"
Sub bullet2()
With ActiveWindow.Selection.TextRange
With .Paragraphs.ParagraphFormat.bullet
.Visible = True
        .RelativeSize = 0.6
        .Character = 108
            With .Font
            .Name = "Wingdings"
            .Color.RGB = RGB(127, 127, 127)
            End With
End With
End With
End Sub

