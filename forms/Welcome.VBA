Private Sub Cancel_Click()
    'Close Form
    Unload Welcome
End Sub

Private Sub Chart_Click()
    'Take To Chart
    Worksheets("BMIChart").Activate
    'Close
    Unload Welcome
End Sub
Private Sub Continue_Click()
    If Welcome.Insert.Value = True Then
        InsertOptions.Show 'Take To A UserForm
    ElseIf Welcome.View.Value = True Then
        Worksheets("Stats").Activate 'Take To The Sheet
        Unload Welcome 'Close
    ElseIf Welcome.BMI.Value = True Then
        BMICalc.Show 'Take To A UserForm
    ElseIf Welcome.Calories.Value = True Then
        CaloriesCalc.Show 'Take To A UserForm
    Else
        MsgBox "Please Pick An Option!", vbOKOnly, "Error" 'Show Error Message
    End If
End Sub
