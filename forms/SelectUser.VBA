Private Sub Cancel_Click()
    'Close Form
    Unload SelectUser
End Sub

Private Sub Continue_Click()
    'If Index Greater Than 0
    If UserList.ListIndex >= 0 And MonthList.ListIndex >= 0 Then
        userID = UserList.Value
        monthID = MonthList.Value
        'Set Value Of Cells To The List Box Info
        Cells(1, 2).Value = UserList.List(UserList.ListIndex, 1)
        Cells(1, 5).Value = MonthList.List(MonthList.ListIndex, 1)
        'Close Form
        Unload SelectUser
    'Tell User What To Do
    Else
        MsgBox "Please Select A User And Month.", vbOKOnly + vbInformation, "Select User And Month."
    End If
End Sub
