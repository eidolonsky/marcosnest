Attribute VB_Name = "splitbyparagraph"
Sub splitbyparagraph()
Dim oshp As Shape
Dim oTB As Shape
Dim osld As Slide
Dim L As Long
Dim sngT As Single
Dim sngW As Single
Dim sngL As Single
Dim otxR1 As TextRange
Dim FM As Single
Dim LM As Single

Dim fontsz As Long
Set oshp = ActiveWindow.Selection.ShapeRange(1)
sngL = oshp.Left
sngW = oshp.Width
Set osld = oshp.Parent
fontsz = oshp.TextFrame.TextRange.Paragraphs(1).Font.Size
Set otxR1 = oshp.TextFrame.TextRange.Paragraphs(1)
LM = oshp.TextFrame.Ruler.Levels(1).LeftMargin
FM = oshp.TextFrame.Ruler.Levels(1).FirstMargin
For L = 1 To oshp.TextFrame.TextRange.Paragraphs.Count

If L = 1 Then
sngT = oshp.Top
Else
sngT = otxR1.BoundTop + otxR1.BoundHeight
End If

Set oTB = osld.Shapes.AddTextbox(msoTextOrientationHorizontal, sngL, sngT, sngW, 10)
oTB.TextFrame.TextRange.Text = oshp.TextFrame2.TextRange.Paragraphs(L).Text

Set otxR1 = oTB.TextFrame.TextRange
While Right(otxR1.Text, 1) = Chr(13)
otxR1.Text = Left(otxR1.Text, Len(otxR1.Text) - 1)
Wend

otxR1.ParagraphFormat.Bullet.Visible = _
oshp.TextFrame2.TextRange.Paragraphs(L).ParagraphFormat.Bullet.Visible
otxR1.ParagraphFormat.Bullet.Type = _
oshp.TextFrame.TextRange.Paragraphs(L).ParagraphFormat.Bullet.Type
oTB.TextFrame.TextRange.Font.Size = fontsz
oTB.TextFrame.Ruler.Levels(1).FirstMargin = FM
oTB.TextFrame.Ruler.Levels(1).LeftMargin = LM
Next L
oshp.TextFrame.DeleteText
oshp.Delete
Exit Sub
End Sub

