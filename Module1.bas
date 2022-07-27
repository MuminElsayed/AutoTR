Attribute VB_Name = "Module1"
'HOW TO RUN:
'Remove any course section numbers (eg: ENGL 101-5) --> (ENGL 101)
'Put the number of semesters to go through in the top left corner (Cell A-1 in AutoTR sheet)
'Press Ctrl+Q to run
'NOTES:
'Maximum number of courses to copy in each semester is 6
'To add a third page, just copy any old page
'Semester order is now from left to right, not one up one down
'Don't move any cells wa shokran


Sub autotr()

Dim sheetcount As Integer
Dim studentName As Variant
Dim studentID As Variant
Dim currentSemester As Variant
Dim currentCourseCode As Variant
Dim currentCourseName As Variant
Dim currentCourseCH As Variant
Dim currentCourseQP As Variant
Dim currentCourseGrade As Variant

    If IsEmpty(ThisWorkbook.Sheets(1).Range("A1").Value) Then
        MsgBox "Please specify the number of semesters to copy (Put number in cell A-1 in AutoTR sheet)", , "Error"
        Exit Sub
    End If
sheetcount = ThisWorkbook.Sheets(1).Range("A1").Value
ThisWorkbook.Sheets(1).Range("A1").Value = Null

studentID = Sheets(1).Range("C3").Value
studentName = Sheets(1).Range("C2").Value

ThisWorkbook.Sheets(1).Range("B2").Value = "Student Name: " & studentName 'puts student name
ThisWorkbook.Sheets(1).Range("B3").Value = "Student ID: " & studentID 'puts student Id

'Variables for inputting courses/semester names'
Dim currentCellColumn As Integer
Dim currentCellRow As Integer
Dim cellRowAdder As Integer
Dim x As Integer
Dim y As Integer
Dim outputCell As Range
''''''''''''''''''
currentCellColumn = 10
x = -1
cellRowAdder = 12
currentCellRow = -19
y = 1
''''''''''''''''''

Dim i As Integer
    For i = 1 To sheetcount 'Goes through specified number of sheets
        currentSemester = Sheets(i).Range("C4").Value
        'logic to place semesterNames
        ''''''''''''
        'For columns
        currentCellColumn = currentCellColumn + (8 * x)
        x = x * -1
        'For rows
        If i Mod 2 = 1 Then
            cellRowAdder = cellRowAdder + (17 * y)
            y = y * -1
            currentCellRow = currentCellRow + cellRowAdder
        End If
        ''''''''''''
        ThisWorkbook.Sheets(1).Cells(currentCellRow, currentCellColumn).Value = currentSemester
        For j = 1 To 6 'Goes through courses in the sheet
            
            Dim currentCell As Range
            Set currentCell = Sheets(i).Cells(6 + j, 2) 'Starts at first course cell
            
            currentCourseCode = Mid(currentCell, 1, 8) 'Takes first part as course code
            currentCourseName = Mid(currentCell, 9, 99) 'Takes second part as course name
            currentCourseCH = Mid(currentCell.Offset(0, 1).Value, 1, 1) 'Takes course Credit Hours
            currentCourseGrade = currentCell.Offset(0, 2).Value 'Takes course Grade
            currentCourseQP = currentCell.Offset(0, 3).Value 'Takes course Quality Points
            
            ThisWorkbook.Sheets(1).Cells(currentCellRow + j, currentCellColumn).Value = currentCourseCode 'Puts course code
            ThisWorkbook.Sheets(1).Cells(currentCellRow + j, currentCellColumn + 1).Value = currentCourseName 'Puts course name
            ThisWorkbook.Sheets(1).Cells(currentCellRow + j, currentCellColumn + 3).Value = currentCourseGrade 'Puts course Grade
            ThisWorkbook.Sheets(1).Cells(currentCellRow + j, currentCellColumn + 5).Value = currentCourseCH 'Puts course Credit Hours
            ThisWorkbook.Sheets(1).Cells(currentCellRow + j, currentCellColumn + 7).Value = currentCourseQP 'Puts course Quality Points
            
        Next j
    Next i
    ThisWorkbook.SaveAs Filename:=studentName & " - AutoTR", FileFormat:=-4143
    MsgBox "AutoTR done"
End Sub

