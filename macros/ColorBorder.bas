Attribute VB_Name = "Module1"
Sub ColorBorderSelection()
  
  Dim sh As Shape
  Dim shapeCollection() As String
  
  Set sh = ActiveWindow.Selection.ShapeRange(1)
  
  ReDim Preserve shapeCollection(0)
  
  shapeCollection(0) = sh.Name
  
  Dim otherShape As Shape
  Dim iShape As Integer
  
  iShape = 1
  
  On Error Resume Next
  
  For Each otherShape In ActiveWindow.View.Slide.Shapes
    If otherShape.Line.ForeColor = sh.Line.ForeColor _
    And otherShape.Fill.ForeColor = sh.Fill.ForeColor _
    And otherShape.Type <> msoPlaceholder _
    Then
    If (otherShape.Name <> sh.Name) Then
      ReDim Preserve shapeCollection(1 + iShape)
      shapeCollection(iShape) = otherShape.Name
      iShape = iShape + 1
    End If
    End If
  Next otherShape
  ActiveWindow.View.Slide.Shapes.Range(shapeCollection).Select
  
End Sub

