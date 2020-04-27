Attribute VB_Name = "Toolbar_A"
Const MyToolbar As String = "Tool_A"

Dim oToolbar As CommandBar


Sub ToolBarA()
    Dim oToolbar As CommandBar
    Dim NewButton As CommandBarButton
    Call AddMe
    
    On Error Resume Next
    Application.CommandBars("Tool_A").Delete
    On Error GoTo 0
    
    Set oToolbar = Application.CommandBars.Add _
        (Name:="Tool_A", Temporary:=True)
    oToolbar.Visible = True
    
    
        Set NewButton = oToolbar.Controls.Add _
            (Type:=msoControlButton, ID:=2950)
        NewButton.FaceId = 950
        NewButton.Caption = "Auto Size Comments"
        NewButton.OnAction = "autoSizeComment"

NormalExit:
    Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Private Sub RemoveMe()
    On Error Resume Next
    CommandBars(MyToolbar).Delete
End Sub

Private Sub AddMe()

    Call RemoveMe

    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)

End Sub

Sub autoSizeComment()

    Dim xComment As Comment
    
    For Each xComment In Application.ActiveSheet.Comments
        xComment.Shape.TextFrame.AutoSize = True
        
Next
    
End Sub


