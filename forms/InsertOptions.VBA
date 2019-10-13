Private Sub Cancel_Click()
    'Close
    Unload InsertOptions
End Sub

Private Sub Continue_Click()
    If InsertOptions.Add.Value = True Then
        'Take To UserForm
        AddUser.Show
    ElseIf InsertOptions.File.Value = True Then
        'Set Active Sheet
        Worksheets("Stats").Activate
        'Delet The Previous Chart
        If Worksheets("Stats").ChartObjects.Count > 0 Then
            Worksheets("Stats").ChartObjects.Delete
        End If
        'Call The Module Where SQL Takes Place
        Call Module2.DisplayUserInfo
        'Call The Module Where The Chart Is Made
        Call Module3.CreateChart
        'Close
        Unload InsertOptions
        Unload Welcome
    Else
        MsgBox "Please Pick An Option!", vbOKOnly, "Error" 'Show Error Message
    End If
End Sub
