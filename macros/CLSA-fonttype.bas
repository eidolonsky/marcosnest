Attribute VB_Name = "Module11"
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
    Set oPTT = PowerPoint.ActivePresentation
    With oPTT
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
                                iSA = .AllNodes.Count
                                For i = 1 To iSA
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
