Attribute VB_Name = "Module2"
Sub SpaceHorizontal()
'Automatically space and align shapes

Dim shp As Shape
Dim i As Long
Dim dTop As Double
Dim dLeft As Double
Dim dWidth As Double
Const dSPACE As Double = 0

  'Check if shapes are selected
  If TypeName(Selection) = "Range" Then
    MsgBox "Please select first."
    Exit Sub
  End If
  
  'Set variables
  i = 1
  
  'Loop through selected shapes (charts, slicers, timelines, etc.)
  For Each shp In ActiveWindow.Selection.ShapeRange
    With shp
      'If not first shape then move it below previous shape and align left.
      If i > 1 Then
        .Top = dTop
        .Left = dLeft + dWidth + dSPACE
      End If
      
      'Store properties of shape for use in moving next shape in the collection.
      dTop = .Top
      dLeft = .Left
      dWidth = .Width
    End With
    
    'Add to shape counter
    i = i + 1
    
  Next shp

End Sub

