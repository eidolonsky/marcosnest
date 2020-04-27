Attribute VB_Name = "Module21"
Sub bullet1()
With ActiveWindow.Selection.TextRange
With .Paragraphs.ParagraphFormat.bullet
.Visible = True
        .RelativeSize = 0.6
        .Character = 110
            With .Font
            .Name = "Wingdings"
            .Color.RGB = RGB(0, 0, 83)
            End With
End With
End With
End Sub

