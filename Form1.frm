VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Process Excel Schedule"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lbLog 
      Height          =   3180
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   8655
   End
   Begin MSComDlg.CommonDialog cdGetExcelFile 
      Left            =   0
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnSelectFile 
      Caption         =   "Select File"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnProcess 
      Caption         =   "Process"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblFileName 
      Caption         =   "Select file to process"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   8775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

Dim ExcelFileName As String

' not the best way to do this by an stretch but it'll work for grabbing data for now
Dim EmpName() As String
Dim EmpWorkType() As String
Dim EmpSD() As Date
Dim EmpED() As Date


Private Sub ReadOverview(ws As Excel.Worksheet)

    ' traverse the table area for data
    Dim Col As Integer  ' the columber to read from
    Dim Row As Integer  ' the row to read from

    Dim recCount As Integer   ' track how many items that have been read
    
    Col = 1
    Row = 4
    recCount = 1
    
    ' fill the simple arrays with the data from the first sheet
    With ws
        Do While Not (.Cells(Row, Col) = "")
            ' using this technique because the number of records is unknown, this expenads the arrays as needed
            ReDim Preserve EmpName(recCount), EmpWorkType(recCount), EmpSD(recCount), EmpED(recCount)
            
            EmpName(recCount) = .Cells(Row, Col)
            EmpWorkType(recCount) = .Cells(Row, Col + 1)
            EmpSD(recCount) = .Cells(Row, Col + 2)
            EmpED(recCount) = .Cells(Row, Col + 3)
            
            WriteToListboxLog ("Reading Table 1: Found " & EmpName(recCount) & " doing " & EmpWorkType(recCount) & " for " & EmpSD(recCount) & " thru " & EmpED(recCount))
            
            ' incremement the row number to read from and the record counter
            recCount = recCount + 1
            Row = Row + 1
            
        Loop
                
    End With
    
    WriteToListboxLog ("Finished reading first table")
    
End Sub


Private Sub btnProcess_Click()

    Set objExcel = New Excel.Application
    Set objWorkbook = objExcel.Workbooks.Open(ExcelFileName)
    WriteToListboxLog ("Opened " & ExcelFileName)
        
    ' check to make sure the 2 sheets required are present
    Dim OverviewWS As Excel.Worksheet
    Dim DetailWS As Excel.Worksheet
    Dim WSheet As Excel.Worksheet
    Dim SaveTheWorkbook As Boolean   ' this indicates whether or not to save the file when closing
    SaveTheWorkbook = False
    
    For i = 1 To objExcel.Sheets.Count
        Set WSheet = objExcel.Sheets(i)
        Select Case WSheet.Name
            Case "Overview"
                Set OverviewWS = WSheet
            Case "Employee Schedule Detail"
                Set DetailWS = WSheet
        End Select
    Next
    
    If OverviewWS Is Nothing Or DetailWS Is Nothing Then
        ' a required worksheet is missing from the selected file
        ' this process can not continue
        WriteToListboxLog ("The required sheets were not found, the selected file can not be processed. Exiting.")
    Else
        ' what we expect is present - continue
        SaveTheWorkbook = True
    End If
    
    
    If SaveTheWorkbook Then
   
        WriteToListboxLog ("The required sheets are present, continuing to process the data.")
        
        ' read the overview sheet
        ReadOverview OverviewWS
        
        ' write the data read in the previous step to the detail sheet
        WriteToDetail DetailWS
        
    End If
    
    ' call the sub that saves, close and clears the excel workbook. tell if it should save the file
    AllDone SaveTheWorkbook
    
End Sub

Private Sub AllDone(SaveTheWorkbook As Boolean)

    ' this is called when the process of reading and writing is finished or a file has been selected that
    ' isn't correct for this processing

    If SaveTheWorkbook Then
        objWorkbook.Save
        WriteToListboxLog ("Saved " & ExcelFileName)
    End If
    
    objExcel.ActiveWorkbook.Close False, ExcelFileName
    objExcel.Quit
    
    Set ojbWorkbook = Nothing
    
    btnProcess.Enabled = False  ' the user can't process anything until there is a file selected again
End Sub

Private Sub WriteToDetail(ws As Excel.Worksheet)

    Dim Col As Integer
    Dim Row As Integer
    Dim EName As String
    Dim DateBetween As Date
    
    Dim EmployeeNameHeaderRow As Integer
    EmployeeNameHeaderRow = 2  ' the row containing the employee name
    
    Col = 1
    Dim CurrentEmpName As String
    CurrentEmpName = ""
    
    For i = 1 To UBound(EmpName)
    
        If EmpName(i) <> CurrentEmpName Then
            ' if the employee stays the same then there is no need to find them again in the sheet
            FindEmployeeSection ws, EmpName(i), Row, Col
            
            If Col = -1 Then
                ' -1 indicates that the current employee was not found in the detail sheet and must be skipped
                WriteToListboxLog ("The employee with this record wasn't found in the detail sheet " & EmpName(i) & " doing " & EmpWorkType(i) & " for " & EmpSD(i) & " thru " & EmpED(i))
            Else
                CurrentEmpName = EmpName(i)
                WriteToListboxLog ("Writing data for " & EmpName(i) & " doing " & EmpWorkType(i) & " for " & EmpSD(i) & " thru " & EmpED(i))
            End If
        Else
            'was was left off at the last place that data was written to the sheet, as the name hasn't changed increment
            ' the row to the next line
            Row = Row + 1
        End If
        
        If Col > -1 Then
            ' only write the data to the detail if the employee was found there
            ws.Cells(Row, Col) = EmpSD(i)
            ws.Cells(Row, Col + 1) = EmpWorkType(i)
            WriteToListboxLog ("   writing " & EmpWorkType(i) & " for " & EmpSD(i) & " thru " & EmpED(i))
    
            If EmpSD(i) <> EmpED(i) Then
                ' the employee is performing this role for more than 1 day
                If CDate(EmpED(i)) > CDate(EmpSD(i)) Then
                    'as long as the dates are in the correct order create a row for each date between them
                    DateBetween = EmpSD(i)
                    Do Until DateBetween = EmpED(i)
                        DateBetween = DateBetween + 1
                        Row = Row + 1
                        ws.Cells(Row, Col) = DateBetween
                        ws.Cells(Row, Col + 1) = EmpWorkType(i)
                        WriteToListboxLog ("   writing " & EmpWorkType(i) & " for " & DateBetween)
                    Loop
                    
                End If
            End If
        End If
        
    Next

End Sub


Private Sub FindEmployeeSection(ws As Excel.Worksheet, EmployeeName As String, _
        ByRef Row As Integer, ByRef Col As Integer)

    ' byref allows me to modify the value of the Col and Row so on return the calling procedure can use it
    ' this sub procedure finds the emplyee name in the detail sheet
    
    Dim EmployeeNameHeaderRow As Integer
    EmployeeNameHeaderRow = 2  ' the row containing the employee names
    
    ' with the format of the sheet I have if there are 3 blank spaces on row 2 then there are no more
    ' employees to match to
    Dim EmptyCellCount As Integer
    EmptyCellCount = 0
    
    Col = 1 ' reset the column when searching in case emplyee names are not in order
    Row = 4 ' in my format this is the row number the data should start writting on

    With ws
    
        ' find the employee header on the employee name header row
        Do While True
            If .Cells(EmployeeNameHeaderRow, Col) = "" Then
                EmptyCellCount = EmptyCellCount + 1
                
                If EmptyCellCount = 3 Then
                    Col = -1
                    Exit Do
                End If
                
                Col = Col + 1
            Else
                EmptyCellCount = 0
                If EmployeeName = .Cells(EmployeeNameHeaderRow, Col) Then
                    ' found the correct employee, now scroll down to find the next empty row
                    Do While .Cells(Row, Col) <> ""
                        Row = Row + 1
                    Loop
                    Exit Do
                Else
                    ' move to the next column to continue to check for the employee name
                    Col = Col + 1
                End If
            End If
        Loop
    End With
    
End Sub

Private Sub WriteToListboxLog(LogLine As String)

    ' take the passed text and add it to the listbox to show the user what is going on
    lbLog.AddItem (LogLine)
    DoEvents
    
End Sub

Private Sub ResetForm()

    ' reset the controls on the form
    lblFileName.Caption = ""
    btnProcess.Enabled = False
    lbLog.Clear
    
End Sub

Private Sub btnSelectFile_Click()
    
    ResetForm
    
    On Error Resume Next ' if an error occurs simply ignore it and continue to the next line of code
    
    cdGetExcelFile.ShowOpen   ' open the file dialog box to allow the user to select which excel file to use
    
    If Err.Number = cdlCancel Then
        ' The user canceled the dialog
        Exit Sub
    ElseIf Err.Number <> 0 Then
        ' Unknown error
        MsgBox "Error " & Format$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
        Exit Sub
    End If
    
    On Error GoTo 0  ' resume normal error handling
    
    ExcelFileName = cdGetExcelFile.FileName
    lblFileName.Caption = cdGetExcelFile.FileName
    
    ' enable the process button now that they have selected a file
    btnProcess.Enabled = True

End Sub

Private Sub Form_Load()

    ' setup the file open dialog
    With cdGetExcelFile
        .Filter = "Excel Files (*.xlsx)|*.xlsx"
        .DefaultExt = "xlsx"
        .DialogTitle = "Select Excel File"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNExplorer
        .CancelError = True
    End With
    
    ' was used for testing ExcelFileName = "C:\Documents and Settings\Administrator\My Documents\employee scheduling.xlsx"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objExcel = Nothing

End Sub
