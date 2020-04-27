VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} verticalAlign 
   Caption         =   "Vertical Gap"
   ClientHeight    =   2295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2145
   OleObjectBlob   =   "VerticalGap.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "verticalAlign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CommandButton1_Click()
    Dim shp As Shape
    Dim i As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double


    Dim oInput As Double


    On Error Resume Next

  'Set variables
    i = 1
  
  'Loop through selected shapes (charts, slicers, timelines, etc.)
  For Each shp In ActiveWindow.Selection.ShapeRange
    With shp
      'If not first shape then move it below previous shape and align left.
      If i > 1 Then
      
          oInput = TextBox1.Value
    
        If OptionButton1.Value = "True" Then
            oInput = oInput * 72
        ElseIf OptionButton2.Value = "True" Then
            oInput = oInput * 28.35
        ElseIf OptionButton3.Value = "True" Then
            oInput = oInput
        End If
      
        .Top = dTop + dHeight + oInput
        .Left = dLeft
      End If
      
      'Store properties of shape for use in moving next shape in the collection.
      dTop = .Top
      dLeft = .Left
      dHeight = .Height
    End With
    
    'Add to shape counter
    i = i + 1
    
  Next shp
  
End Sub

Public Sub CommandButton2_Click()
    Unload Me
End Sub

Public Sub OptionButton1_Click()
    OptionButton1.Value = "True"
End Sub

Public Sub OptionButton2_Click()
    OptionButton2.Value = "True"
End Sub

Public Sub OptionButton3_Click()
    OptionButton3.Value = "True"
End Sub

Public Sub TextBox1_Change()
    TextBox1.Text = TextBox1.Value
End Sub

Public Sub UserForm_Initialize()
    TextBox1.Value = ""
    
    OptionButton1.Value = "True"
    
    TextBox1.SetFocus
End Sub
