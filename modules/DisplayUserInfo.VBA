Option Explicit
'===========================================================+
' Program title: DisplayUserInfo
' Description: Allows User To Open Database File And Then Select
'              A User And Month And Returns Weight, Height, BMI
'              Info For User Along With The Week.
'===========================================================+
Public userID As String, monthID As String
    Sub DisplayUserInfo()
        'Initiate Variables
        Dim rowLast As Integer
        Dim filePath As String, fileExt As String, fileChoice As Integer
        Dim isValid As Boolean
        Dim connection As ADODB.connection, recSetUser As ADODB.Recordset, recSetMonth As ADODB.Recordset
        Dim userSQL As String, monthSQL As String
        Dim userInput As Object
        Dim i As Integer
        Dim recSetResult As ADODB.Recordset
        Dim resultSQL As String
        Dim cell As Range
        Dim strDate As String
        'Reset Cells For Next Batch Of Data
        Cells(1, 2).Clear
        Cells(1, 5).Clear
        rowLast = Cells(Rows.Count, 1).End(xlUp).Row
        If rowLast <> 3 Then
            Range(Cells(4, 1), Cells(rowLast, 4)).Clear
        End If
    'While Valid
    Do Until isValid
        'Open File Dialog Box
        With Application.FileDialog(msoFileDialogOpen)
            .AllowMultiSelect = False 'Only Select One File
            .Title = "Select Database(.mdb) File" 'Title
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
            .Filters.Add "Access Database Files", "*.mdb" 'Type Of File To Select
        End With
    
        'Make Dialog Box Visible
        fileChoice = Application.FileDialog(msoFileDialogOpen).Show
        'Get File Path
        If fileChoice <> 0 Then
            filePath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
            'Check To See If Correct File
            fileExt = Right(filePath, Len(filePath) - InStrRev(filePath, "."))
            'If Valid
            If StrComp(fileExt, "mdb") = 0 Then
                isValid = True
            'If Not Valid.
            Else
                MsgBox "Please Select A .mdb File.", vbOKOnly + vbInformation, "Error"
            End If
        'Else No File Selected
        Else
            MsgBox "No File Selected. Application Closing...", vbOKOnly + vbInformation, "Terminated"
            Exit Sub
        End If
    Loop
    
        'Open Connection
        Set connection = New ADODB.connection
        With connection
            .Provider = "Microsoft.Jet.OLEDB.4.0"
            .Open filePath
        End With
        
        'Build SQL String To Open Form
        userSQL = "SELECT UsersID, UserFirstName, UserLastName FROM Users;"
        monthSQL = "SELECT MonthsID, MonthName FROM Months;"
        
        'Define User Recordset And Get Info
        Set recSetUser = New ADODB.Recordset
        recSetUser.CursorLocation = adUseServer
        recSetUser.Open Source:=userSQL, _
            ActiveConnection:=connection, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdText
        recSetUser.MoveFirst 'Move To First Value
        
        'Define Month Recordset And Get Info
        Set recSetMonth = New ADODB.Recordset
        recSetMonth.CursorLocation = adUseServer
        recSetMonth.Open Source:=monthSQL, _
            ActiveConnection:=connection, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdText
        recSetMonth.MoveFirst 'Move To First Value
        
        'Populate The User And Month List Boxes In The Userform
        Set userInput = SelectUser
        
        'User List
        With SelectUser.UserList
            .ColumnCount = recSetUser.Fields.Count 'Number Of Entries
            i = 0
            'Add The Users To List Box
            Do
                .AddItem 'New Row
                .List(i, 0) = recSetUser!UsersID
                .List(i, 1) = recSetUser!UserFirstName & " " & recSetUser!UserLastName
                i = i + 1
                recSetUser.MoveNext 'Next Entry
                
            Loop Until recSetUser.EOF
            
            .ColumnWidths = "1 cm; "
        
        End With
        
        'Month List
        With SelectUser.MonthList
            .ColumnCount = recSetMonth.Fields.Count 'Number Of Entries
            i = 0
            'Add The Months To List Box
            Do
                .AddItem 'Add A New Row
                .List(i, 0) = recSetMonth!MonthsID
                .List(i, 1) = recSetMonth!MonthName
                i = i + 1
                recSetMonth.MoveNext 'Next Entry
                
            Loop Until recSetMonth.EOF
            
            .ColumnWidths = "1 cm; "
            
        End With
        
        'Prompt User To Choose Current User And Month
        isValid = False
        
        userInput.Show
        
        Do Until isValid
            If IsEmpty(Cells(1, 2)) And IsEmpty(Cells(1, 5)) Then
                MsgBox "No Options Selected! Application Closing...", vbOKOnly + vbInformation, "Terminated"
                Exit Sub
            Else
                isValid = True
            End If
        Loop
        
        'Build SQL String For Results
        resultSQL = "SELECT Stats.WeekOne, Stats.WeekTwo, Stats.WeekThree, Stats.WeekFour, Stats.WeekOneWeight, Stats.WeekTwoWeight, Stats.WeekThreeWeight, Stats.WeekFourWeight, Stats.WeekOneHeight, Stats.WeekTwoHeight, Stats.WeekThreeHeight, Stats.WeekFourHeight, Stats.WeekOneBMI, Stats.WeekTwoBMI, Stats.WeekThreeBMI, Stats.WeekFourBMI " _
        & "FROM Stats " _
        & "WHERE Stats.UsersID = " & userID & " AND Stats.MonthsID = " & monthID
        
        'Define Resutls Recordset And Get Info
        Set recSetResult = New ADODB.Recordset
        recSetResult.CursorLocation = adUseServer
        recSetResult.Open Source:=resultSQL, _
            ActiveConnection:=connection, _
            CursorType:=adOpenDynamic, _
            LockType:=adLockOptimistic, _
            Options:=adCmdText
        
        'Print Dates To Worksheet
        Worksheets("Stats").Cells(4, 1).Value = recSetResult("WeekOne").Value
        Worksheets("Stats").Cells(5, 1).Value = recSetResult("WeekTwo").Value
        Worksheets("Stats").Cells(6, 1).Value = recSetResult("WeekThree").Value
        Worksheets("Stats").Cells(7, 1).Value = recSetResult("WeekFour").Value
        'Print Weights To Worksheet
        Worksheets("Stats").Cells(4, 2).Value = recSetResult("WeekOneWeight").Value
        Worksheets("Stats").Cells(5, 2).Value = recSetResult("WeekTwoWeight").Value
        Worksheets("Stats").Cells(6, 2).Value = recSetResult("WeekThreeWeight").Value
        Worksheets("Stats").Cells(7, 2).Value = recSetResult("WeekFourWeight").Value
        'Print Heights To Worksheet
        Worksheets("Stats").Cells(4, 3).Value = recSetResult("WeekOneHeight").Value
        Worksheets("Stats").Cells(5, 3).Value = recSetResult("WeekTwoHeight").Value
        Worksheets("Stats").Cells(6, 3).Value = recSetResult("WeekThreeHeight").Value
        Worksheets("Stats").Cells(7, 3).Value = recSetResult("WeekFourHeight").Value
        'Print BMIs To Worksheet
        Worksheets("Stats").Cells(4, 4).Value = recSetResult("WeekOneBMI").Value
        Worksheets("Stats").Cells(5, 4).Value = recSetResult("WeekTwoBMI").Value
        Worksheets("Stats").Cells(6, 4).Value = recSetResult("WeekThreeBMI").Value
        Worksheets("Stats").Cells(7, 4).Value = recSetResult("WeekFourBMI").Value
        
        'Close Connections
        recSetUser.Close
        recSetMonth.Close
        recSetResult.Close
        connection.Close
    'End
    End Sub


