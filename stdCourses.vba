
Option Explicit

Private Cancel As Boolean
'if user press cancle
Private Sub cmdCancel_Click()
    'make it do nothing
    Me.Hide
    Cancel = True
End Sub
' in case user press x button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'same as pressing cancel, calls the cancel_click
  If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub
'public function
Public Function Initialize(cn As ADODB.Connection) As Boolean
    On Error GoTo ErrorHandler
    Dim SQL As String, rng As Range, rowCount As Integer
    Dim rs As New ADODB.Recordset 'needed to take info from database
    Dim stdArray(49, 3) As String '50 students with fname, lname and id
    SQL = "SELECT * FROM students"
    rs.Open SQL, cn
    rowCount = 0
    With rs
        Do Until .EOF
            stdArray(rowCount, 0) = .Fields("FirstName")
            stdArray(rowCount, 1) = .Fields("LastName")
            stdArray(rowCount, 2) = .Fields("studentID")
            rowCount = rowCount + 1
            .MoveNext
        Loop
    End With
    'making the list appear from the array from database
    lbStd.List = stdArray
    'automatically selecting first option
    lbStd.ListIndex = 0
    Me.Show
    Initialize = Not Cancel
    If (Initialize) Then
        'using cn, courseName and courseID to get info from grades
        Call makeSheet(cn, stdArray(lbStd.ListIndex, 0), stdArray(lbStd.ListIndex, 1), stdArray(lbStd.ListIndex, 2))
    Else
        MsgBox "You have canceled the program.", vbInformation
        Unload Me
    End If
    On Error GoTo 0
    'closing record set and database
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorHandler: 'if there was an error, it was due to the database having incorrect infoamtion
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'closing record set and database
    rs.Close
    Set rs = Nothing
End Function
'if user press ok
Private Sub cmdOk_Click()
    Me.Hide
    Cancel = False
End Sub
'sub that makes the sheet based on the data selected from user
Private Sub makeSheet(cn As ADODB.Connection, fName As String, lName As String, ID As String)
    'error handler in case an error occurs
    On Error GoTo ErrorHandler
    Dim rng As Range, SQL As String
    Dim rs As New ADODB.Recordset
    Dim rowCount As Integer
    rowCount = 0
    'helper sub to set up sheet
    Call SetUpSheet(fName, lName, ID)
    'getting info from database
    Set rng = Range("A6")
    SQL = "SELECT * FROM grades"
    'opening the record set with sql command
    rs.Open SQL, cn
    'with the record set, do until reaches the end of file character or until all courses has been set to true
    With rs
        Do Until .EOF
            'if the fields is the same student user selected
            If .Fields("studentID") = ID Then
                'putting data on the sheet
                rng.Offset(rowCount, 0).Value = .Fields("course")
                rng.Offset(rowCount, 1).Value = .Fields("A1")
                rng.Offset(rowCount, 2).Value = .Fields("A2")
                rng.Offset(rowCount, 3).Value = .Fields("A3")
                rng.Offset(rowCount, 4).Value = .Fields("A4")
                rng.Offset(rowCount, 5).Value = .Fields("MidTerm")
                rng.Offset(rowCount, 6).Value = .Fields("Exam")
                rowCount = rowCount + 1
            End If
            .MoveNext
        Loop
        'adding labels 2 rows after table for min,max and marks for each task
    End With
    'putting info onto sheet based on booleans
    Call CalculateStats
    'adding the chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    'since each course has the same amount of students, this source data will always be the same place until data base get's change
    'getting avg data range
    Set rng = Range("C17:J17")
    ActiveChart.SetSourceData Source:=rng
    'getting titles for x index
    Set rng = Range("C5:I5")
    ActiveChart.FullSeriesCollection(1).XValues = rng
    'setting title
    ActiveChart.ChartTitle.Text = "Averages for " & fName & " " & lName
    'adding outline + making gapwidth 0
    ActiveChart.ChartGroups(1).GapWidth = 0
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(38, 38, 38)
    End With
    'moving the chart so does not block data
    ActiveSheet.Shapes("Chart 1").IncrementLeft 200.75
    ActiveSheet.Shapes("Chart 1").IncrementTop 2.25
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'calling helper sub that finds stats for each grade
    Range("A1").Select
    On Error GoTo 0
    'closing record set and database
    rs.Close
    Set rs = Nothing
    Exit Sub
'if there was an error, it was due to the database having incorrect infoamtion
ErrorHandler:
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'closing record set and database
    rs.Close
    Set rs = Nothing
End Sub
'sub to set up the sheet
Private Sub SetUpSheet(fName As String, lName As String, ID As String)
    Dim rng As Range
    'making sure the sheet does not already exist
    Call deleteSheets_Program_Specific(fName & " " & lName & " " & ID & " Report")
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = fName & " " & lName & " " & ID & " Report"
    Set rng = Range("A1")
    'setting up details on page, can't use with/loops since values keep changing
    With rng
        .Value = fName & " " & lName & " " & ID & " Report"
        .Offset(1, 0).Value = "First Name"
        .Offset(1, 1).Value = "Last Name"
        .Offset(1, 2).Value = "ID"
        ActiveSheet.UsedRange.Font.Bold = True    'bolding important infomation
        .Offset(2, 0).Value = fName
        .Offset(2, 1).Value = lName
        .Offset(2, 2).Value = ID
        .Offset(4, 0).Value = "Course"
        .Offset(4, 1).Value = "Course Code"
        .Offset(4, 2).Value = "A1"
        .Offset(4, 2).HorizontalAlignment = xlRight
        .Offset(4, 3).Value = "A2"
        .Offset(4, 4).Value = "A3"
        .Offset(4, 5).Value = "A4"
        .Offset(4, 6).Value = "A5"
        .Offset(4, 8).Value = "Final Mark"
        'from looking at the data base it seems every student has taken all of the courses, can assume _
        where the stats should be
        .Offset(14, 1).Value = "Minimum Mark"
        .Offset(15, 1).Value = "Maximum Mark"
        .Offset(16, 1).Value = "Average Mark"
    End With
End Sub
'a private/ support sub for the makeSheet function, used to calcualte the min,max,avg for each assign,midterm and final
Private Sub CalculateStats()
    On Error GoTo ErrorHandler
    'constant mark weights
    Const assign As Double = 0.05
    Const Midterm As Double = 0.3
    Const Exam As Double = 0.5
    'before this method is called, sheets should be completed and with makeSheets
    'rng is the A1 of the latest / last ID in the database
    Dim cell As Range, finalGrade As Double, calcRange As Range, rng As Range
    'rng of latest set of data will be the same since each student has taken all courses in the database
    Set rng = Range("C13:G13")
    For Each cell In rng
        'setting the calcRange to the range of values needed
        Set calcRange = Range(cell, cell.End(xlUp))
        'finding min,max and avg and offsetting to put the value in the corresponding cells
        With cell
            .Offset(2, 0).Value = WorksheetFunction.Min(calcRange)
            .Offset(3, 0).Value = WorksheetFunction.Max(calcRange)
            .Offset(4, 0).Value = WorksheetFunction.Average(calcRange)
                'making it to 2 deciaml places
            .Offset(4, 0).NumberFormat = "0.00"
        End With
    Next cell
    'finding min, max and avg of the indivual students
    Set rng = Range(Range("G6"), Range("G6").End(xlDown))
    For Each cell In rng
        'finding final mark for the students
        'final, midterm a4,a3,a2,a1, assuming each assignment is 5%, midterm is 30% and final is 50%
        finalGrade = cell.Value * Exam + cell.Offset(0, -1).Value * Midterm + cell.Offset(0, -2).Value * assign + _
        cell.Offset(0, -3).Value * assign + cell.Offset(0, -4).Value * assign + cell.Offset(0, -5).Value * assign
        cell.Offset(0, 2).Value = finalGrade
        cell.NumberFormat = "0.00"
    Next cell
    
    'finding min, max and avg final marks
    Set rng = Range("I6").End(xlDown)
    'setting the calcRange to the range of values needed (final scores)
    Set calcRange = Range(rng, rng.End(xlUp))
    With rng
        .Offset(2, 0) = WorksheetFunction.Min(calcRange)
        .Offset(3, 0) = WorksheetFunction.Max(calcRange)
        .Offset(4, 0) = WorksheetFunction.Average(calcRange)
            'making it to 2 deciaml places
        .Offset(4, 0).NumberFormat = "0.00"
    End With
    On Error GoTo 0
    Exit Sub
'if there was an error, it was due to the database having incorrect infoamtion
ErrorHandler:
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
End Sub








