Attribute VB_Name = "Module1"
Sub ShapeSelection()
  
  Dim sh As Shape
  Dim shapeCollection() As String
  
  Set sh = ActiveWindow.Selection.ShapeRange(1)
  
  ReDim Preserve shapeCollection(0)
  
  shapeCollection(0) = sh.Name
  
  Dim otherShape As Shape
  Dim iShape As Integer
  
  iShape = 1
  Dim j As Integer
  
  On Error Resume Next
  
  For Each otherShape In ActiveWindow.View.Slide.Shapes
  
    If StripNumber(otherShape.Name) = StripNumber(sh.Name) Then
        If (otherShape.Name <> sh.Name) Then
         ReDim Preserve shapeCollection(1 + iShape)
         shapeCollection(iShape) = otherShape.Name
         iShape = iShape + 1
        End If
    End If
  Next otherShape
  
  ActiveWindow.View.Slide.Shapes.Range(shapeCollection).Select
  
End Sub

Function StripNumber(stdText As String)
Dim str As String, i As Integer

    stdText = Trim(stdText)

    For i = 1 To Len(stdText)
        If Not IsNumeric(Mid(stdText, i, 1)) Then
            str = str & Mid(stdText, i, 1)
        End If
    Next i

        StripNumber = str

End Function

