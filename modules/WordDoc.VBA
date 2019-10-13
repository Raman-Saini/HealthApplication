Option Explicit
'===========================================================+
' Program title: WordDoc
' Description: Generates Outputs User Data To Word Document
'===========================================================+
Sub wordDoc()
'Get The User's Stats
    Worksheets("Stats").Range("A3:D7").Copy
    'Create Variables
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim docName As String
    'Ignore Error
    On Error Resume Next
    'Creat Application
    Set wordApp = GetObject(, "Word.Application")
    'Clear If It Comes Up
    If Err.Number = 429 Then
        Err.Clear
    End If
    'Show
    wordApp.Visible = True
    'Path
    docName = "C:\Users\raman\Documents\CP212 Projects\Assignment 5\Output.docx"
    'If Not There
    If docName = "" Then
        MsgBox "This File " & docName & " Does Not Exsist"
        Exit Sub
    End If
    'Activate
    wordApp.Activate
    'Set Document
    Set wordDoc = wordApp.Documents(docName)
    'Open
    If wordDoc Is Nothing Then Set wordDoc = wordApp.Documents.Open(docName)
    'Set
    wordDoc.Activate 'Activate
    wordDoc.Range.Paste 'Paste The Range Copied Earlier
    wordDoc.Save 'Save
    wordApp.Quit 'Close
    
    Set wordDoc = Nothing 'Reset
    Set wordApp = Nothing 'Reset
    Application.CutCopyMode = False 'Can't Change
'End
End Sub
