Private Sub Calculate_Click()
    Dim userWeight As Variant
    Dim userHeight As Variant
    Dim userAge As Variant
    Dim userCalorieInatke As Variant
    
    If Not IsNumeric(CaloriesCalc.Weight.Value) Or Not IsNumeric(CaloriesCalc.High.Value) Or Not IsNumeric(CaloriesCalc.Age.Value) Then
        'Error Message
        MsgBox "Please Try Again! Reminder Only Enter Numbers!", vbOKOnly, "Error"
        'Clear All Boxes
        CaloriesCalc.Weight.Value = ""
        CaloriesCalc.High.Value = ""
        CaloriesCalc.Age.Value = ""
        CaloriesCalc.Result.Value = ""
    Else
        'Set Variables
        userWeight = CaloriesCalc.Weight.Value
        userHeight = CaloriesCalc.High.Value
        userAge = CaloriesCalc.Age.Value
        'If Gender Make
        If CaloriesCalc.MaleOption.Value = True Then
            userCalorieIntake = ((10 * userWeight) + (6.25 * userHeight) - (5 * userAge) + 5)
        'Else Female
        ElseIf CaloriesCalc.FemaleOption.Value = True Then
            userCalorieIntake = ((10 * userWeight) + (6.25 * userHeight) - (5 * userAge) - 161)
        Else
            'Error Message
            MsgBox "Pleae Pick A Gender To Continue!", vbOKOnly, "Error"
            'Clear All Boxes
            CaloriesCalc.Weight.Value = ""
            CaloriesCalc.High.Value = ""
            CaloriesCalc.Age.Value = ""
            CaloriesCalc.Result.Value = ""
        End If
        'Display Result In Result Box
        CaloriesCalc.Result.Value = Round(userCalorieIntake, 2)
    End If
End Sub
Private Sub Cancel_Click()
    'Close Form
    Unload CaloriesCalc
End Sub

Private Sub Clear_Click()
    'Clear All Boxes
    CaloriesCalc.Weight.Value = ""
    CaloriesCalc.High.Value = ""
    CaloriesCalc.Age.Value = ""
    CaloriesCalc.Result.Value = ""
End Sub
