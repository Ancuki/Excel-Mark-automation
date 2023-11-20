Option Explicit
'the main function that will open the userform and _
well as operate the specific task from the userform, used both to pratice using classes and for storing class marks
Type class
    name As String
    tMarks As Double
    numStudents As Integer
End Type
'main sub that is called when button is pressed, opens the main userform for operation
Sub Main()
    'turning off some setting for better user experience
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'initlizing the main userform
    Call userfrmMain.Initialize
    'turning settings back on
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
'one for the button, another for the porgram since the screen updating will be change back to true _
which we don't want if the program is using it
'deletes every sheet but the main sheet
Sub deleteSheets_Program()
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        If sheet.name <> MainSheet.name Then
            sheet.Delete
        End If
    Next sheet
End Sub
'for certain worksheet that is not main, finds and delete sheet
Sub deleteSheets_Program_Specific(name As String)
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        'if the sheet is not the main sheet and the name of the sheet is the same oe given, delete
        If sheet.name = name And sheet.name <> MainSheet.name Then
            sheet.Delete
        End If
    Next sheet
End Sub
'look if sheet exist, returns boolean if it does
Public Function sheetExist(name As String) As Boolean
    'default initilize value of boolean is false
    Dim sheet As Worksheet, a As Integer
    a = 1
    'do until a=number of sheet and sheetExist is false
    Do While a <= Worksheets.count And Not sheetExist
        If Worksheets(a).name = name Then
            sheetExist = True
        End If
        a = a + 1
    Loop
End Function
'delete all sheets but main, button on the sheet, different since this one needs to make screen updating off and such
Sub deleteSheets_Button()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim sheet As Worksheet
    For Each sheet In Worksheets
        If sheet.name <> MainSheet.name Then
            sheet.Delete
        End If
    Next sheet
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub



