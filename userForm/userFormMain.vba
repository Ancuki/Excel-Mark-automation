Dim Cancel As Boolean
Option Explicit

'if user press ok
Private Sub cmdContinue_Click()
    'if user does not press cancel or exit // if user presses continue
    If Cancel = False Then
    'setting up sub with error hadnling and variables, including connection to the database
        On Error GoTo Wrongfile
        Dim selected As Boolean
        Dim cn As New ADODB.Connection
        Dim fd As Office.FileDialog
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        selected = True
        fd.InitialFileName = ThisWorkbook.Path & "\Registrar.mdb"
        'opening and setting up the connection
        With cn
            .ConnectionString = "Data Source=" & fd.InitialFileName
            .Provider = "Microsoft.ACE.OLEDB.12.0"
            .Open
        End With
        'if user wants to import data
        If optImportData.Value = True Then
            Call FilePicker 'special function to pick the data file
        'is user picks list courses
        ElseIf optListCourse.Value = True Then
            Me.Hide 'no longer need this userform, if user does not cancel the new userform for listCourse
            If (ListCourse.Initialize(cn)) Then
                'msgbox to say that function is done and sheet can be viewed
                MsgBox "Sheet is complete", vbInformation, "Done"
            Else
                Unload Me 'unload if user does not want to continue
            End If
            'if user presses generate report
        ElseIf optGenerateReport.Value = True Then
             Call GenerateReport
             'if user presses course enrroll
        ElseIf optCourseEnroll.Value = True Then
            'if user does not cancel the new userform for course enrollment
            If (stdCourse.Initialize(cn)) Then
                MsgBox "Sheet is complete", vbInformation, "Done"
            Else
                'unload if user does not want to continue
                Unload Me
            End If
        Else 'if use does not select an option and presses continue
            MsgBox "Please select an option", vbExclamation, "No Selection"
            Exit Sub
        End If
        'making objects nothing for security reasons
        fd.InitialFileName = ""
        cn.Close
        Set fd = Nothing
        Set cn = Nothing
    End If
    Unload Me
    Exit Sub
Wrongfile:
    MsgBox "The file " & fd.InitialFileName & " could not be found", vbCritical, "Error"
    'making objects nothing for security reasons
    fd.InitialFileName = ""
    Set fd = Nothing
    cn.Close
    Set cn = Nothing
    On Error GoTo 0
    Unload Me
End Sub

'if user press cancle
Private Sub cmdCancel_Click()
    'make it do nothing
    MsgBox "You have canceled the program.", vbInformation
    Me.Hide
    Cancel = True
End Sub

' in case user press x button
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = vbFormControlMenu Then cmdCancel_Click
'  If CloseMode = 0 Then
'    MsgBox "You must not live in a Big Ten Provinces.",vbInformation
'    Unload Me
'    End
'  End If
End Sub
'this here in case wanting anything to be set upon activasion
Private Sub UserForm_Initialize()
    Cancel = False
    'setting up userform properties
End Sub
'public sub that main can call
Public Sub Initialize()
    Me.Show
End Sub
'sub for import courses, allows user to select file for database
Private Sub FilePicker()
    On Error GoTo fileError 'if the file is not a database, error will occur to tell user
    Dim fd As Office.FileDialog
    Dim notCancel As Boolean
    'from slides to select file from pc
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'ask user to select the correct database
    MsgBox "Please select the Register.mdb file from the folder", vbDefaultButton1, "Select File"
    With fd
        'if not canceled
        notCancel = .Show
        If notCancel Then
            'index 1 holds the file string/path from example in slides + textbook
            Call ImportData(.SelectedItems(1)) 'method to set up connection and call functions to make the sheets
        Else
            MsgBox "Process has been canceled", vbOKOnly, "Cancled"
        End If
    End With
    Exit Sub
fileError:
    MsgBox "The file(s) you selected does not match the database for this project, please pick" & _
    " the correct single file or check for errors", vbCritical, "Error"
End Sub
''method to set up connection and call functions to make the sheets
Private Sub ImportData(file As Variant)
    Dim cn As New ADODB.Connection
    'making the connection from the file given by the user
    With cn
        .ConnectionString = "Data Source=" & file
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    'deleting previous sheets if there is one
    Call deleteSheets_Program_Specific("Grades")
    Call deleteSheets_Program_Specific("Courses")
    Call deleteSheets_Program_Specific("Students")
    'importing all data from the database
    Call importCourses(cn)
    Call importGrades(cn)
    Call importStudents(cn)
    'msgbox saying the sheets have been made
    MsgBox "3 Sheets has been made with the data" & vbNewLine & "Please go to the desired sheet for the data", vbInformation, "Task Complete"
    'selects the mainsheet afterwards
    MainSheet.Select
    'closing and making connection = to nothing
    cn.Close
    Set cn = Nothing
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
    'rng of latest set of data will be the same since each student has taken all courses in the database
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
    'finding min, max and avg of the indivual students and courses
    Set rng = Range(Range("I2"), Range("I2").End(xlDown))
    For Each cell In rng
        'finding final mark for the students
        'final, midterm a4,a3,a2,a1, assuming each assignment is 5%, midterm is 30% and final is 50%
        finalGrade = cell.Value * Exam + cell.Offset(0, -1).Value * Midterm + cell.Offset(0, -2).Value * assign + _
        cell.Offset(0, -3).Value * assign + cell.Offset(0, -4).Value * assign + cell.Offset(0, -5).Value * assign
        cell.Offset(0, 1).Value = finalGrade
        cell.Offset(0, 1).NumberFormat = "0.00"
    Next cell
    
    'finding min, max and avg final marks
    Set rng = Range("J2").End(xlDown)
    'setting the calcRange to the range of values needed (final scores)
    Set calcRange = Range(rng, rng.End(xlUp))
    With rng
        .Offset(2, 0) = WorksheetFunction.Min(calcRange)
        .Offset(3, 0) = WorksheetFunction.Max(calcRange)
        .Offset(4, 0) = WorksheetFunction.Average(calcRange)
            'making it to 2 deciaml places
        .Offset(4, 0).NumberFormat = "0.00"
    End With
    'make charts
    'chart1, for max, min and avg of everything
    Range("C403:J405").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    'title of graph
    ActiveChart.ChartTitle.Text = "Min, Max, Avg For All Courses"
    'making biggest scale to 100
    ActiveChart.Axes(xlValue).MaximumScale = 100
    'making labels
    ActiveChart.FullSeriesCollection(1).XValues = "=Grades!$D$1:$J$1"
    'moving chart to proper spot
    'making chart 2
    Range("J2:J401").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    'making labels
    'title of graph
    ActiveChart.ChartTitle.Text = "Final Marks"
    'making biggest scale to 100
    ActiveChart.Axes(xlValue).MaximumScale = 100
    'deleting this x label
    ActiveChart.Axes(xlCategory).Delete
    'goback to a1 + moving charts to visuable locations
    ActiveSheet.Shapes("Chart 1").IncrementLeft 350
    ActiveSheet.Shapes("Chart 1").IncrementTop -74.25
    ActiveSheet.Shapes("Chart 2").IncrementLeft 350
    ActiveSheet.Shapes("Chart 2").IncrementTop 132.75
    Range("A1").Select
    On Error GoTo 0
    Exit Sub
'if there was an error, it was due to the database having incorrect infoamtion
ErrorHandler:
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
End Sub
Private Sub importGrades(cn As ADODB.Connection)
    'error handler in case an error occurs
    On Error GoTo ErrorHandler
    Dim SQL As String, rng As Range
    Dim rs As New ADODB.Recordset 'needed to take info from database
    'sq command
    SQL = "SELECT * FROM grades"
    'adding sheet to the right most side and naming it Grades
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = "Grades"
    Set rng = Range("A1")
    'this can be used to import the whole database, but I decided to do it indivualy to ensure the db is the correct one
'    rs.Open SQL, cn
'    rng.CopyFromRecordset rs
    'adding labels for detailed sheet
    With rng
        .Value = "ID"
        .Offset(0, 1).Value = "studentID"
        .Offset(0, 2).Value = "course"
        .Offset(0, 3).Value = "A1"
        .Offset(0, 4).Value = "A2"
        .Offset(0, 5).Value = "A3"
        .Offset(0, 6).Value = "A4"
        .Offset(0, 7).Value = "MidTerm"
        .Offset(0, 8).Value = "Exam"
    End With
    With Range("A1:I1")
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    rng.Offset(0, 9).Value = "Final"
    'move rng down one row
    Set rng = rng.Offset(1, 0)
    'going to database and copying data into new sheet row by row
    With rs
        .Open SQL, cn    'opening the record set with sql command
        Do Until .EOF    'with the record set, do until reaches the end of file character
            rng.Value = .Fields("ID")
            rng.Offset(0, 1).Value = .Fields("studentID")
            rng.Offset(0, 2).Value = .Fields("course")
            rng.Offset(0, 3).Value = .Fields("A1")
            rng.Offset(0, 4).Value = .Fields("A2")
            rng.Offset(0, 5).Value = .Fields("A3")
            rng.Offset(0, 6).Value = .Fields("A4")
            rng.Offset(0, 7).Value = .Fields("MidTerm")
            rng.Offset(0, 8).Value = .Fields("Exam")
            Set rng = rng.Offset(1, 0)
            .MoveNext
        Loop
    End With
    'adding min max and avg labels
    rng.Offset(1, 2).Value = "Min:"
    rng.Offset(2, 2).Value = "Max:"
    rng.Offset(3, 2).Value = "Avg:"
    Set rng = Range(rng.Offset(-1, 3), rng.Offset(-1, 3).End(xlToRight))
    'calculating stats for marks
    Call CalculateStats(rng)
    'autofiting all cells in the used range
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'closing record set and database
    rs.Close
    Set rs = Nothing
    On Error GoTo 0
    Exit Sub
ErrorHandler: 'if there was an error, it was due to the database having incorrect infoamtion
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'deleting errror sheet
    ActiveSheet.Delete
    'selecting main sheet again
    MainSheet.Select
    'closing record set and database
    rs.Close
    Set rs = Nothing
End Sub
Private Sub importStudents(cn As ADODB.Connection)
    'error handler in case an error occurs
    On Error GoTo ErrorHandler
    Dim SQL As String, rng As Range
    Dim rs As New ADODB.Recordset 'needed to take info from database
    'making sure the sheet does not exist via deletion
    SQL = "SELECT * FROM students" 'sql command
    'adding sheet to the right most side and naming it Students
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = "Students"
    'adding labels for details
    Set rng = Range("A1")
    rng.Value = "FirstName"
    rng.Offset(0, 1).Value = "LastName"
    rng.Offset(0, 2).Value = "studentID"
    With Range("A1:C1")
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    Set rng = rng.Offset(1, 0)
    'going to database and copying data into new sheet row by row
    With rs
        .Open SQL, cn    'opening the record set with sql command
        'with the record set, do until reaches the end of file character
        Do Until .EOF
            rng.Value = .Fields("FirstName")
            rng.Offset(0, 1).Value = .Fields("LastName")
            rng.Offset(0, 2).Value = .Fields("studentID")
            Set rng = rng.Offset(1, 0)
            .MoveNext
        Loop
    End With
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'closing record set and database
    rs.Close
    Set rs = Nothing
    
    On Error GoTo 0
    Exit Sub
ErrorHandler: 'if there was an error, it was due to the database having incorrect infoamtion
    MsgBox "Please check that the database has the correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'deleting the incorrect sheet
    ActiveSheet.Delete
    'selecting main sheet again
    MainSheet.Select
    'closing record set and database
    rs.Close
    Set rs = Nothing
End Sub

Private Sub importCourses(cn As ADODB.Connection)
    'error handler in case an error occurs
    On Error GoTo ErrorHandler
    Dim SQL As String, rng As Range
    Dim rs As New ADODB.Recordset 'needed to take info from database
    'making sure the sheet does not exist via deletion
    'sql command
    SQL = "SELECT * FROM courses"
    'adding sheet to the right most side and naming it Courses
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = "Courses"
    Set rng = Range("A1")
    'adding labels for table headers
    rng.Value = "ID"
    rng.Offset(0, 1).Value = "CourseCode"
    rng.Offset(0, 2).Value = "CourseName"
    With Range("A1:C1")
        .Font.Bold = True
        .Interior.ColorIndex = 40
    End With
    Set rng = rng.Offset(1, 0)
    'going to database and copying data into new sheet row by row
    With rs
        'opening the record set with sql command
        .Open SQL, cn
        'with the record set, do until reaches the end of file character
        Do Until .EOF
            rng.Value = .Fields("ID")
            rng.Offset(0, 1).Value = .Fields("CourseCode")
            rng.Offset(0, 2).Value = .Fields("CourseName")
            Set rng = rng.Offset(1, 0)
            .MoveNext
        Loop
    End With
    'autofitting all cells in the used range
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'closing record set and database
    rs.Close
    Set rs = Nothing

    On Error GoTo 0
    Exit Sub
'if there was an error, it was due to the database having incorrect infoamtion
ErrorHandler:
    MsgBox "Please check that the database had correct infomation: currently there is an error", _
    vbExclamation, "Error"
    'deleting the incorrect sheet
    ActiveSheet.Delete
    'closing record set and database + selecting main sheet
    MainSheet.Select
    rs.Close
    Set rs = Nothing
End Sub
'function to generate report
Private Sub GenerateReport()
    On Error GoTo ErrorHandler
    Dim wdDoc As Word.Document
    Dim WordApp As Object
    Dim wdSel As Word.Selection
    ' create a new instance of Word  .
    Dim wdApp As New Word.Application
    Dim file As String
    file = ThisWorkbook.Path & "\" & _
    "vuon8730_Doc_Output_Generated" & ".docx"
    'checking if there is already a report and then deleting if there is, file needs to be close though
    With New FileSystemObject
        If .FileExists(file) Then
            .DeleteFile file
        End If
    End With
    ' See Word application.
    wdApp.Visible = True
     ' Add a new document. Select the active window
     Set wdDoc = wdApp.Documents.Add
     Set wdSel = wdDoc.ActiveWindow.Selection
     ' Insert text that is wanted
    wdSel.TypeText Text:="User Guide Generated"
    wdSel.TypeParagraph
    wdSel.TypeText Text:="Date: 2022-12-03"
     ' Add new line by pressing Enter.
    wdSel.TypeParagraph
    wdSel.TypeText "Author: Andy Vuong"
    wdSel.TypeParagraph
    wdSel.TypeText Text:="ID: 210868730"
    wdSel.TypeParagraph
    wdSel.TypeText Text:="This is a sample output from the Class Average results " & _
    "These results are acquired from pressing the Display Class Averages button on the main user form."
    wdSel.TypeParagraph
    wdSel.TypeText Text:="Class Averages"
    wdSel.TypeParagraph
    wdSel.TypeText Text:="From the datasheet, every student has taken every class which makes the total amount " & _
    "of students that has taken the course equal to 50, which can be confirm by counting the number of students in " & _
    "the database."
    wdSel.TypeParagraph
    'making the sheet to get data and graph, delete AverageSheet exist, delete first to ensure there is no errors
    Call makeAverageSheet
    Worksheets("Average Report").Select
    'copy and paste the graphs
    ActiveSheet.Shapes("Chart 1").CopyPicture
    wdSel.Paste
    wdSel.TypeParagraph
    wdSel.TypeText Text:="From the histogram above (from the Display Class Averages results)," & _
    "cp102 average is " & Range("B3").Value & ", cp104 is " & Range("B4").Value & ", cp212 is " & _
     Range("B5").Value & ", as101 is " & Range("B6").Value & ", pc120 is " & _
     Range("B7").Value & ", pc131 is " & Range("B8").Value & ", and cp411 is " & Range("B9").Value
     'detling the report sheet
    wdSel.TypeParagraph
    Worksheets("Grades").Select
    'copy and paste the graphs
    ActiveSheet.Shapes("Chart 1").CopyPicture
    wdSel.Paste
    wdSel.TypeParagraph
    wdSel.TypeText Text:="Here is the min, max, and average for every assignment, midterm, and exam for all " & _
    "courses."
    wdSel.TypeParagraph
    'copy and paste the graphs
    ActiveSheet.Shapes("Chart 2").CopyPicture
    wdSel.Paste
    wdSel.TypeParagraph
    wdSel.TypeText Text:="Here are all the final marks for all students taking all the courses. Refer to the Grades " & _
    "worksheet by using the import data option for more details as puttig all infomation for the graphs above would take a lot of space."
    Call deleteSheets_Program_Specific("Average Report")
    'see if import data has been called, and if it has then don't delete the sheet
    If sheetExist("Courses") = False And sheetExist("Students") = False Then
        Call deleteSheets_Program_Specific("Grades")
    End If
    'going back onto the mainsheet
    MainSheet.Select
    ' Save the document and close it.
    'moving cursor back to home/start
    wdApp.Selection.HomeKey Unit:=wdStory, Extend:=wdMove
    wdDoc.SaveAs file
    MsgBox "The file can be found at " & file & vbNewLine & vbNewLine & "The file is opened in the backgroud!"
    ' wdDoc.Close
    Exit Sub
ErrorHandler:
    MsgBox "An error has occured, please close the report file so the program can delete it to remake the docuement", vbCritical, "Close File"
End Sub
Private Sub cmdAverages_Click()
    'this is a function so when the user wants a report, the program can use this and control if making the sheet _
    will close exit the sub early or not
    Call makeAverageSheet
    Unload Me
    MsgBox "Average Report Sheet has been made", vbDefaultButton1, "Sheet Made"
    Exit Sub
End Sub
Private Sub makeAverageSheet()
    'setting up sub with error hadnling and variables
    On Error GoTo Wrongfile
    Dim selected As Boolean
    Dim cn As New ADODB.Connection
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    selected = True
    fd.InitialFileName = ThisWorkbook.Path & "\Registrar.mdb"
    'opening the connection
    With cn
        .ConnectionString = "Data Source=" & fd.InitialFileName
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open
    End With
    'see if Grade sheet has been made, and if it hasn't, then make it as it would be needed in the report
    If sheetExist("Grades") = False Then
        Call importGrades(cn)
    End If
    'making the sheets needed for the doc + calling the functions
    Call calculateAverages(cn)
    'if the sheet does not exist, then make it
    On Error GoTo 0
    'setting objects back to nothing
    cn.Close
    fd.InitialFileName = ""
    Set cn = Nothing
    Set fd = Nothing
    Exit Sub
Wrongfile:
    MsgBox "The file " & fd.InitialFileName & "could not be found", vbCritical, "Error"
    'making objects nothing for security reasons
    cn.Close
    fd.InitialFileName = ""
    Set fd = Nothing
    Set cn = Nothing
    Unload Me
End Sub
Private Sub calculateAverages(cn As ADODB.Connection)
    Dim courses As String, SQL As String, counter As Integer
    Dim rs As New ADODB.Recordset 'needed to take info from database
    Dim cp102 As class, cp104 As class, cp212 As class, as101 As class, pc120 As class, pc131 As class, pc141 As class, cp411 As class
    SQL = "SELECT * FROM grades"
    'making different course types to keep data
    With cp102
        .name = "CP102"
        .tMarks = 0
        .numStudents = 0
    End With
    With cp104
        .name = "CP104"
        .tMarks = 0
        .numStudents = 0
    End With
    With cp212
        .name = "CP212"
        .tMarks = 0
        .numStudents = 0
    End With
    With as101
        .name = "AS101"
        .tMarks = 0
        .numStudents = 0
    End With
    With pc120
        .name = "PC120"
        .tMarks = 0
        .numStudents = 0
    End With
    With pc131
        .name = "PC131"
        .tMarks = 0
        .numStudents = 0
    End With
    With pc141
        .name = "PC141"
        .tMarks = 0
        .numStudents = 0
    End With
    With cp411
        .name = "CP411"
        .tMarks = 0
        .numStudents = 0
    End With
    'going to database and copying data into string row by row
    With rs
        .Open SQL, cn
        Do Until .EOF
            If .Fields("course") = "CP102" Then
                'adding std by 1 (counting total std)
                cp102.numStudents = cp102.numStudents + 1
                'calling private function to find the final mark per student taking the course
                cp102.tMarks = cp102.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "CP104" Then
                'adding std by 1 (counting total std)
                cp104.numStudents = cp104.numStudents + 1
                'calling private function to find the final mark per student taking the course
                cp104.tMarks = cp104.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "CP212" Then
                'adding std by 1 (counting total std)
                cp212.numStudents = cp212.numStudents + 1
                'calling private function to find the final mark per student taking the course
                cp212.tMarks = cp212.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "AS101" Then
                'adding std by 1 (counting total std)
                as101.numStudents = as101.numStudents + 1
                'calling private function to find the final mark per student taking the course
                as101.tMarks = as101.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "PC120" Then
                'adding std by 1 (counting total std)
                pc120.numStudents = pc120.numStudents + 1
                'calling private function to find the final mark per student taking the course
                pc120.tMarks = pc120.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "PC131" Then
                'adding std by 1 (counting total std)
                pc131.numStudents = pc131.numStudents + 1
                'calling private function to find the final mark per student taking the course
                pc131.tMarks = pc131.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "PC141" Then
                'adding std by 1 (counting total std)
                pc141.numStudents = pc141.numStudents + 1
                'calling private function to find the final mark per student taking the course
                pc141.tMarks = pc141.tMarks + findFinal(rs)
            ElseIf .Fields("course") = "CP411" Then
                'adding std by 1 (counting total std)
                cp411.numStudents = cp411.numStudents + 1
                'calling private function to find the final mark per student taking the course
                cp411.tMarks = cp411.tMarks + findFinal(rs)
            End If
            .MoveNext
        Loop
    End With
    '   Call DispayAverages(Array(cp102, cp104, cp212, as101, pc120, pc131, pc141, cp411))
    '   helper sub to make new sheet and display averages
    Call DisplayAverages(cp102, cp104, cp212, as101, pc120, pc131, pc141, cp411)
    rs.Close
    'setting record set to nothing
    Set rs = Nothing
End Sub
'private function used to calcualte the final mark
Private Function findFinal(rs As ADODB.Recordset) As Double
'constant weight of each thing
    Const assign As Double = 0.05
    Const Midterm As Double = 0.3
    Const Exam As Double = 0.5
    With rs
        findFinal = .Fields("A1") * assign + .Fields("A2") * assign + .Fields("A3") * assign + .Fields("A4") * assign + _
        .Fields("MidTerm") * Midterm + .Fields("Exam") * Exam
    End With
End Function
Private Sub DisplayAverages(cp102 As class, cp104 As class, cp212 As class, as101 As class, pc120 As class, _
pc131 As class, pc141 As class, cp411 As class)
'classes() As Variant 'tried using an array but got an error, keeping here in case of change
    Dim class As Variant, count As Integer, rng As Range, chrt As ChartObject
    'deleting report if there was one previously and adding it to the right most side
    Call deleteSheets_Program_Specific("Average Report")
    Worksheets.Add After:=Sheets(Sheets.count)
    ActiveSheet.name = "Average Report" 'naming it report
    Set rng = Range("A1")
    'adding detais (avg) from diving total marks by total students in each class class
    With rng
        .Value = "Class Average Report"
        .Offset(1, 0) = "Class"
        .Offset(1, 1) = "Average"
        .Offset(2, 0) = cp102.name
        .Offset(2, 1) = cp102.tMarks / cp102.numStudents
        .Offset(3, 0) = cp104.name
        .Offset(3, 1) = cp104.tMarks / cp104.numStudents
        .Offset(4, 0) = cp212.name
        .Offset(4, 1) = cp212.tMarks / cp212.numStudents
        .Offset(5, 0) = as101.name
        .Offset(5, 1) = as101.tMarks / as101.numStudents
        .Offset(6, 0) = pc120.name
        .Offset(6, 1) = pc120.tMarks / pc120.numStudents
        .Offset(7, 0) = pc131.name
        .Offset(7, 1) = pc131.tMarks / pc131.numStudents
        .Offset(8, 0) = pc141.name
        .Offset(8, 1) = pc141.tMarks / pc141.numStudents
        .Offset(9, 0) = cp411.name
        .Offset(9, 1) = cp411.tMarks / cp411.numStudents
    End With
    'making it to 2 deciaml places
    Range(Range("b3"), Range("b3").End(xlDown)).NumberFormat = "0.00"
    'adding chart to the sheet
    'vba record marco for graph
    rng = Worksheets("Average Report").Range("A2:B10")
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.ChartGroups(1).GapWidth = 0
    ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(38, 38, 38)
    End With
    'autofiting all columns to ensure all data will look good and is readable
    ActiveSheet.UsedRange.EntireColumn.AutoFit 'can also just be columns.autofit
    'selecting the a1 cell at the end
    Range("A1").Select
End Sub


