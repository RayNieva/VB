Attribute VB_Name = "Module1"
Sub OpenCSVFile()
Attribute OpenCSVFile.VB_ProcData.VB_Invoke_Func = " \n14"
'
' OpenCSVFile Macro
'

'
    'Application.Run "UI_report.xlsm!OpenCSVFile"
    Workbooks.Open Filename:="C:\Users\Ray\Documents\Projects\Ruby\the_file.csv"
End Sub


Sub MacroSaveAsXLFileUI_Report()
'
' MacroSaveAsXLFileUI_Report Macro
'

'
    ChDir "C:\Users\Ray\Documents\Projects\Ruby"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Ray\Documents\Projects\Ruby\UI_report.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub


Sub MacroSave2JobSearchLogs()
'
' MacroSave2JobSearchLogs Macro
'

'
    ChDir "C:\Users\Ray\Documents\JobSearch\JobSearchLogs"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\Users\Ray\Documents\JobSearch\JobSearchLogs\UI_report.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub

Sub StringTest()
Dim reportDate As Date
Dim weekNumber As Integer

reportDate = 4 / 6 / 2018


Debug.Print reportDate
'Debug.Print "This is the date: " & reportDate





End Sub

Sub addDays()



Dim firstDate As Date, secondDate As Date

firstDate = DateValue("April 08, 2018")
secondDate = DateAdd("d", 7, firstDate)

Debug.Print firstDate

Debug.Print secondDate

End Sub


Sub LoopAddDates()


Dim firstDate As Date, secondDate As Date, Total As Integer

'Debug.Print "Week" & 1


firstDate = DateValue("April 08, 2018")
'Debug.Print firstDate

firstDate2 = Format(CDate(firstDate), "mm-dd-yy")

endOfWeekDate = DateAdd("d", 6, firstDate)
endOFWeekDate2 = Format(CDate(endOfWeekDate), "mm-dd-yy")

'Debug.Print "Week" & 1 & "...." & firstDate2 & "...." & endOFWeekDate2
'Debug.Print "JobSearchLogWeek" & 1 & "-" & firstDate2 & ".xlsx"
Debug.Print "copy JobSearchLogHeaderTemplate2.xlsx " & "JobSearchLogWeek" & 1 & "-" & firstDate2 & ".xlsx"
'Week 28 Starting Sunday (date)___10-15-17__________________________ Through Saturday (date)__10-21-17__________
Debug.Print "Week " & 1 & " Starting Sunday (date)___" & firstDate2 & "__________________________ Through Saturday (date)__" & endOFWeekDate2


Total = 1

For i = 2 To 52
    Total = Total + 1
    
    'Debug.Print "Week" & Total
    

    secondDate = DateAdd("d", 7, firstDate)
    secondDate2 = Format(CDate(secondDate), "mm-dd-yy")
    
    endOfWeekDate = DateAdd("d", 6, secondDate)
    endOFWeekDate2 = Format(CDate(endOfWeekDate), "mm-dd-yy")

    'Debug.Print firstDate
    

    'Debug.Print secondDate
    'Debug.Print "Week" & Total & "...." & secondDate2 & "...." & endOFWeekDate2
    Debug.Print "JobSearchLogWeek" & Total & "-" & secondDate2 & ".xlsx"
    
    'Debug.Print "copy JobSearchLogHeaderTemplate2.xlsx " & "JobSearchLogWeek" & Total & "-" & secondDate2 & ".xlsx"
    
    Debug.Print "Week " & Total & " Starting Sunday (date)___" & secondDate2 & "__________________________ Through Saturday (date)__" & endOFWeekDate2
    
    firstDate = secondDate

    
Next i



End Sub

Sub SetValue()

Worksheets(A13).Activate
ActiveCell.Value = "This is worklog info"

End Sub


Sub Example3()
Dim objWorkbook As Workbook
Dim objFileDialog As FileDialog
Dim strPath As String

Set objFileDialog = Application.FileDialog(msoFileDialogOpen)
objFileDialog.Show
If objFileDialog.SelectedItems.Count <> 0 Then
strPath = objFileDialog.SelectedItems.Item(1)
Debug.Print strPath
Set objWorkbook = Workbooks.Open(strPath)
objWorkbook.Worksheets.Item(1).Cells(1, 1) = "Job Search Logs"
'saves the file at the current location
objWorkbook.Save
End If
End Sub


