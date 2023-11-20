Option Explicit
'if the userform got canceled
Private Cancel As Boolean
'if user press cancle
Private Sub cmdCancel_Click()
    'make it do nothing
    Me.Hide
    Cancel = True
End Sub
' in case user press x button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        'same as pressing cancel
  If CloseMode = vbFormControlMenu Then cmdCancel_Click
End Sub

Public Function Initialize(cn As ADODB.Connection) As Boolean
    On Error GoTo ErrorHandler
    Dim SQL As String, rng As Range, rowCount As Integer
    Dim rs As New ADODB.Recordset 'needed to take info from database
    Dim courseArray(7, 2) As String '8 courses with 3 detail columns
    SQL = "SELECT * FROM courses"
    rs.Open SQL, cn
    rowCount = 0
    With rs
        Do Until .EOF
            courseArray(rowCount, 2) = .Fields("ID")
            courseArray(rowCount, 1) = .Fields("CourseCode")
            courseArray(rowCount, 0) = .Fields("CourseName")
            rowCount = rowCount + 1
            .MoveNext
        Loop
    End With
    'making the list appear from the array from database
    lbCourses.List = courseArray
    'automatically selecting first option
    lbCourses.ListIndex = 0
    Me.Show
    Initialize = Not Cancel
    If (Initialize) Then
        'using cn, courseName and courseID to get info from grades
        Call makeSheet(cn, courseArray(lbCourses.ListIndex, 0), courseArray(lbCourses.ListIndex, 1), _
        courseArray(lbCourses.ListIndex, 2))
        Range("A1").Select
    Else
        MsgBox "You have canceled the program.", vbInformation
        Unload Me
    End If
    On Error GoTo 0
    Unload Me
    'closing record set
    rs.Close
    Set rs = Nothing
    Exit Function
ErrorHandler: 'if there was an error, it was due to the database having incorrect infoamtion
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'closing record set
    rs.Close
    Set rs = Nothing
End Function
'if user press ok
Private Sub cmdOk_Click()
    Me.Hide
    Cancel = False
End Sub
'sub that makes the sheet based on the data selected from user
Private Sub makeSheet(cn As ADODB.Connection, name As String, Course As String, ID As String)
    'error handler in case an error occurs
    On Error GoTo ErrorHandler
    Dim rng As Range, SQL As String, rowCount As Integer
    Dim rs As New ADODB.Recordset
    'the actual info on the report
    Dim maxGrd As Integer, minGrd As Integer, avg As Integer
    'making sure the sheet does not already exist
    Call deleteSheets_Program_Specific(Course & " Report")
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = Course & " Report"
    Set rng = Range("A1")
    'setting up details on page, can't use with/loops since values keep changing
    With rng
        .Value = name & " Report"
        .Offset(1, 0).Value = "Course name"
        .Offset(1, 1).Value = "Course code"
        .Offset(1, 2).Value = "Course ID"
        ActiveSheet.UsedRange.Font.Bold = True
        .Offset(2, 0).Value = name
        .Offset(2, 1).Value = Course
        .Offset(2, 2).Value = ID
        .Offset(4, 0).Value = "ID"
        .Offset(4, 1).Value = "studentID"
        'dont need these, was just here for testing
        'rng.Offset(4, 2).Value = "course"
        .Offset(4, 2).Value = "A1"
        .Offset(4, 3).Value = "A2"
        .Offset(4, 4).Value = "A3"
        .Offset(4, 5).Value = "A4"
        .Offset(4, 6).Value = "Midterm"
        .Offset(4, 7).Value = "Exam"
        .Offset(4, 9).Value = "Final Mark"
    End With
    'bolding important infomation
    'getting info from database
    Set rng = Range("A6")
    rowCount = 0
    SQL = "SELECT * FROM grades"
    'opening the record set with sql command
    rs.Open SQL, cn
    'with the record set, do until reaches the end of file character
    With rs
        Do Until .EOF
            'if the fields is the same course user selected
            If .Fields("course") = Course Then
                rng.Offset(rowCount, 0).Value = .Fields("ID")
                rng.Offset(rowCount, 1).Value = .Fields("studentID")
                'rng.Offset(rowCount, 2).Value = .Fields("course") 'for testing purposes to see it's _
                all from the same course
                rng.Offset(rowCount, 2).Value = .Fields("A1")
                rng.Offset(rowCount, 3).Value = .Fields("A2")
                rng.Offset(rowCount, 4).Value = .Fields("A3")
                rng.Offset(rowCount, 5).Value = .Fields("A4")
                rng.Offset(rowCount, 6).Value = .Fields("Midterm")
                rng.Offset(rowCount, 7).Value = .Fields("Exam")
                rowCount = rowCount + 1
            End If
            .MoveNext
        Loop
        'adding labels 2 rows after table for min,max and marks for each task
        With rng
            .Offset(rowCount + 1, 1).Value = "Minimum Mark"
            .Offset(rowCount + 2, 1).Value = "Maximum Mark"
            .Offset(rowCount + 3, 1).Value = "Average Mark"
            .Offset(rowCount + 3, 1).NumberFormat = "0.00"
        End With
        
    End With
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'calling helper sub that finds stats for each grade
    Call CalculateStats(Range(rng.Offset(rowCount - 1, 2), rng.Offset(rowCount - 1, 2).End(xlToRight)))
    'adding the chart
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    'since each course has the same amount of students, this source data will always be the same place until data base get's change
    'getting avg data range
    Set rng = Range("C59:J59")
    ActiveChart.SetSourceData Source:=rng
    'getting titles for x index
    Set rng = Range("C5:J5")
    ActiveChart.FullSeriesCollection(1).XValues = rng
    'setting title
    ActiveChart.ChartTitle.Text = "Averages"
    'adding outline + making gapwidth 0
    ActiveChart.ChartGroups(1).GapWidth = 0
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(38, 38, 38)
    End With
    'moving the chart so does not block data
    ActiveSheet.Shapes("Chart 1").IncrementLeft 250
    ActiveSheet.Shapes("Chart 1").IncrementTop 2.25
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
'a private/ support sub for the makeSheet function, used to calcualte the min,max,avg for each assign,midterm and final
Private Sub CalculateStats(rng As Range)
    On Error GoTo ErrorHandler
    'constant mark weights
    Const assign As Double = 0.05
    Const Midterm As Double = 0.3
    Const Exam As Double = 0.5
    'before this method is called, sheets should be completed and with makeSheets
    'rng is the A1 of the latest / last ID in the database
    Dim cell As Range, finalGrade As Double, calcRange As Range
    For Each cell In rng
        'setting the calcRange to the range of values needed
        Set calcRange = Range(cell, cell.End(xlUp))
        'finding min,max and avg and offsetting to put the value in the corresponding cells
        With cell
            .Offset(2, 0).Value = WorksheetFunction.Min(calcRange)
            .Offset(3, 0).Value = WorksheetFunction.Max(calcRange)
            .Offset(4, 0).Value = WorksheetFunction.Average(calcRange)
        End With
    Next cell
    'finding min, max and avg of the indivual students
    Set rng = Range(Range("H6"), Range("H6").End(xlDown))
    For Each cell In rng
        'finding final mark for the students
        'final, midterm a4,a3,a2,a1, assuming each assignment is 5%, midterm is 30% and final is 50%
        finalGrade = cell.Value * Exam + cell.Offset(0, -1).Value * Midterm + cell.Offset(0, -2).Value * assign + _
        cell.Offset(0, -3).Value * assign + cell.Offset(0, -4).Value * assign + cell.Offset(0, -5).Value * assign
        cell.Offset(0, 2).Value = finalGrade
        cell.NumberFormat = "0.00"
    Next cell
    
    'finding min, max and avg final marks
    Set rng = Range("J6").End(xlDown)
    'setting the calcRange to the range of values needed (final scores)
    Set calcRange = Range(rng, rng.End(xlUp))
    With rng
        .Offset(2, 0) = WorksheetFunction.Min(calcRange)
        .Offset(3, 0) = WorksheetFunction.Max(calcRange)
        .Offset(4, 0) = WorksheetFunction.Average(calcRange)
        .Offset(4, 0).NumberFormat = "0.00"
    End With

    On Error GoTo 0
    Exit Sub
'if there was an error, it was due to the database having incorrect infoamtion
ErrorHandler:
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
End Sub






