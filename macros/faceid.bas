Attribute VB_Name = "Module1"
Sub ShowFaceIDs()
    Dim NewToolbar As CommandBar
    Dim NewButton As CommandBarButton
    Dim i As Integer, IDStart As Integer, IDStop As Integer
    
    On Error Resume Next
    Application.CommandBars("FaceIds").Delete
    On Error GoTo 0
    
    Set NewToolbar = Application.CommandBars.Add _
        (Name:="FaceIds", temporary:=True)
    NewToolbar.Visible = True
    
    IDStart = 1
    IDStop = 1000
    
    For i = IDStart To IDStop
        Set NewButton = NewToolbar.Controls.Add _
            (Type:=msoControlButton, ID:=2950)
        NewButton.FaceId = i
        NewButton.Caption = "FaceID = " & i
    Next i
    NewToolbar.Width = 600
End Sub

