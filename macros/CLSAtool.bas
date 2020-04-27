Attribute VB_Name = "CLSATOOL"
Option Explicit

Public oWidth As Variant
Public oHeight As Variant

Const MyToolbar As String = "CLSA TOOLS"

Dim oToolbar As CommandBar

Sub Auto_Open()

    Dim oButton As CommandBarButton
    Call AddMe


    On Error Resume Next


    On Error GoTo ErrorHandler


    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)

    With oButton

         .DescriptionText = "Font-Change"


         .Caption = "Font-Change"


         .OnAction = "fontcheck"


         .Style = msoButtonIcon
   

         .FaceId = 80

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Bullet Level 1"


         .Caption = "Bullet Level 1"


         .OnAction = "bullet1"


         .Style = msoButtonIcon
   

         .FaceId = 71

    End With
    

Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    With oButton

         .DescriptionText = "Bullet Level 2"


         .Caption = "Bullet Level 2"


         .OnAction = "bullet2"


         .Style = msoButtonIcon
   

         .FaceId = 72

    End With
    
Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Bullet Level 3"


         .Caption = "Bullet Level 3"


         .OnAction = "bullet3"


         .Style = msoButtonIcon
   

         .FaceId = 73

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Bullet Level 4"


         .Caption = "Bullet Level 4"


         .OnAction = "bullet4"


         .Style = msoButtonIcon
   

         .FaceId = 74

    End With
    
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Bullet Period Removal"


         .Caption = "Bullet Period Removal"


         .OnAction = "remove_bullet_periods"


         .Style = msoButtonIcon
   

         .FaceId = 770

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Select with Color & Border"


         .Caption = "Select with Color & Border"


         .OnAction = "ColorBorderSelection"


         .Style = msoButtonIcon
   

         .FaceId = 962

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Select with Shape"


         .Caption = "Select with Shape"


         .OnAction = "ShapeSelection"


         .Style = msoButtonIcon
   

         .FaceId = 689

    End With
    
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "AlignVertical"


         .Caption = "Align Vertical"


         .OnAction = "GapVertical"


         .Style = msoButtonIcon
   

         .FaceId = 360

    End With
    
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Align Horizontal"


         .Caption = "Align Horizontal"


         .OnAction = "GapHorizontal"


         .Style = msoButtonIcon
   

         .FaceId = 39

    End With
    
        Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Swap Position"


         .Caption = "Swap Position"


         .OnAction = "SwapPosition"


         .Style = msoButtonIcon
   

         .FaceId = 525

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Split textbox by paragraph"


         .Caption = "Split textbox by paragraph"


         .OnAction = "SplitByParagraph"


         .Style = msoButtonIcon
   

         .FaceId = 520

    End With
    
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Copy the size of object"


         .Caption = "Copy the size of object"


         .OnAction = "CopySize"


         .Style = msoButtonIcon
   

         .FaceId = 342

    End With
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    
    With oButton

         .DescriptionText = "Paste the size of object"


         .Caption = "Paste the size of object"


         .OnAction = "PasteSize"


         .Style = msoButtonIcon
   

         .FaceId = 352

    End With
    
        oToolbar.Top = 150
        oToolbar.Left = 150
        oToolbar.Visible = True
NormalExit:
    Exit Sub

ErrorHandler:
     MsgBox err.Number & vbCrLf & err.Description
     Resume NormalExit:
End Sub

Private Sub RemoveMe()
' Removes the toobar if it already exists:
    On Error Resume Next
    CommandBars(MyToolbar).Delete
End Sub

Private Sub AddMe()
    ' If the toolbar already exists, remove it
    Call RemoveMe

    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True
End Sub

Sub fontcheck()
    Dim oShape As Shape
    Dim oSlide As Slide
    Dim oRng As TextRange
    Dim oTable As Table
    Dim oChart As Chart
    Dim oCell As Cell
    Dim oSmtart As SmartArt
    Dim oNode As SmartArtNode
    Dim oPPT As Presentation
    Dim iCol
    Dim iRow
    Dim i, j, iSa, oShapes, oShape1
    Set oPPT = PowerPoint.ActivePresentation
    With oPPT
        For Each oSlide In .Slides
            With oSlide
                For Each oShape In .Shapes
                    With oShape
                        If .HasTextFrame Then
                            Set oRng = .TextFrame.TextRange
                            With oRng.Font
                                 .NameAscii = "Arial"
                                 .NameFarEast = "KaiTi_GB2312"
                            End With
                        End If
                        If .HasTable Then
                            Set oTable = .Table
                            With oTable
                                iCol = .Columns.Count
                                iRow = .Rows.Count
                                For i = 1 To iRow
                                    For j = 1 To iCol
                                        Set oRng = .Cell(i, j).Shape.TextFrame.TextRange
                                        With oRng.Font
                                 .NameAscii = "Arial"
                                 .NameFarEast = "KaiTi_GB2312"
                                        End With
                                    Next j
                                Next i
                            End With
                        End If
                        If .HasChart Then
                            Set oChart = .Chart
                            With oChart.ChartArea.Format.TextFrame2.TextRange.Font
                                 .NameAscii = "Arial"
                                 .NameFarEast = "KaiTi_GB2312"
                            End With
                   
                            
                        End If
                        If .HasSmartArt Then
                            Set oSmtart = .SmartArt
                            With oSmtart
                                iSa = .AllNodes.Count
                                For i = 1 To iSa
                                    With .AllNodes(i).TextFrame2.TextRange.Font
                                     .NameAscii = "Arial"
                                     .NameFarEast = "KaiTi_GB2312"
                                    End With
                                Next i
                            End With
                        End If
                        
                        If .Type = msoGroup Then
                        Set oShapes = .GroupItems
                        For Each oShape1 In oShapes
                            With oShape1
                                 If .HasTextFrame Then
                            Set oRng = .TextFrame.TextRange
                            With oRng.Font
                                 .NameAscii = "Arial"
                                 .NameFarEast = "KaiTi_GB2312"
                            End With
                        End If
                            End With
                        Next
                        End If
                    End With
                Next
            End With
        Next
    End With
End Sub

Sub bullet1()
On Error Resume Next
With ActiveWindow.Selection.TextRange
.IndentLevel = 1
With .Paragraphs.ParagraphFormat.Bullet
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

Sub bullet2()
On Error Resume Next
With ActiveWindow.Selection.TextRange
.IndentLevel = 2
With .Paragraphs.ParagraphFormat.Bullet
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

Sub bullet3()
On Error Resume Next
With ActiveWindow.Selection.TextRange
.IndentLevel = 3
With .Paragraphs.ParagraphFormat.Bullet
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

Sub bullet4()
On Error Resume Next
With ActiveWindow.Selection.TextRange
.IndentLevel = 4
With .Paragraphs.ParagraphFormat.Bullet
.Visible = True
        .RelativeSize = 1
        .Character = 8211
            With .Font
            .Name = "Monotype Corsiva"
            .Color.RGB = RGB(127, 127, 127)
            End With
End With
End With
End Sub

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

Sub GapVertical()
    verticalAlign.Show
End Sub

Sub GapHorizontal()
    horizontalAlign.Show
End Sub

Sub SwapPosition()
    Dim l1, t1, l2, t2 As Long
    
    With ActiveWindow.Selection.ShapeRange
        If .Count <> 2 Then
            MsgBox "Please select two shapes"
            
        Else
            l1 = .Item(1).Left
            t1 = .Item(1).Top
            l2 = .Item(2).Left
            t2 = .Item(2).Top
            
            .Item(1).Left = l2
            .Item(1).Top = t2
            
            .Item(2).Left = l1
            .Item(2).Top = t1
        End If
    End With
End Sub

Sub SplitByParagraph()
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


Sub CopySize()
    With ActiveWindow.Selection.ShapeRange(1)
        oWidth = .Width
        oHeight = .Height
    End With
End Sub
Sub PasteSize()
    With ActiveWindow.Selection.ShapeRange
        .Width = oWidth
        .Height = oHeight
    End With
End Sub
