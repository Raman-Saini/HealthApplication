Private Sub Calculate_Click()
    'Initalize Variables
    Dim userWeight As Integer
    Dim userHeightCM As Integer
    Dim userHeightIN As Integer
    Dim userBMI As Integer
    'Error Handaling - Check If User Enters Numbers
    If Not IsNumeric(BMICalc.Weight.Value) Or Not IsNumeric(userHeightCM = BMICalc.High.Value) Then
        MsgBox "Please Try Again! Reminder Only Enter Numbers!", vbOKOnly, "Error"
        BMICalc.Weight.Value = ""
        BMICalc.High.Value = ""
        BMICalc.Result.Value = ""
    'If Checks Out Continue With Calculation
    Else
        'Set Variables
        userWeight = BMICalc.Weight.Value
        userHeightCM = BMICalc.High.Value
        'Convert
        userHeightIN = userHeightCM / 2.54
        'Get BMI
        userBMI = userWeight / (userHeightIN ^ 2) * 703
        'Display Result In Result Box
        BMICalc.Result.Value = Round(userBMI, 2)
    End If
End Sub

Private Sub Cancel_Click()
    'Close The Sub
    Unload BMICalc
End Sub

Private Sub Clear_Click()
    'Clear All The Boxes
    BMICalc.Weight.Value = ""
    BMICalc.High.Value = ""
    BMICalc.Result.Value = ""
End Sub
