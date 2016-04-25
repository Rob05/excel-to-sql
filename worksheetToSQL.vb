
Option Explicit

Dim nFile1 As Object

Sub ConfigurationSQL_Click()
    Dim filename1 As String
    Dim path, tbs As String
    Dim currentWorksheet As Integer
    Dim strLine As String
    Dim strReportType As String
    Dim strCurrencyCode As String
    Dim strTmp As String
    Dim wrkCurrent As Excel.Worksheet
    Dim strIgnoreWorksheet As String
    dim fs as Object
    
    filename1 = Excel.Worksheets("Intro").Cells(10, "B").Text
    strIgnoreWorksheet = Excel.Worksheets("Intro").Cells(11, "B").Text
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    tbs = Application.GetOpenFilename("Select any file in desired output directory:, *.*")
    
    path = fs.GetParentFolderNAme(tbs)
    If Len(tbs) < 6 Then
        MsgBox ("Path not provided. File not generated.")
        Exit Sub
    End If
    
    Set nFile1 = fs.CreateTextFile(filename1, True, 1)
    
    For Each wrkCurrent In Excel.Worksheets
        If InStr(strIgnoreWorksheet, wrkCurrent.Name) = 0 Then
            WriteLine "-- " + wrkCurrent.Name + " --"
            WriteLine "DELETE FROM " + wrkCurrent.Name
            WriteLine "GO"
            CreateSQL wrkCurrent
            WriteLine "GO"
        End If
    Next
    
    nFile1.Close
    Excel.Worksheets("Intro").Activate
    
    MsgBox ("Cash Blotter SQL generated successfully. Review file system.")

End Sub

Function WriteLine(strLine As String)
    nFile1.WriteLine strLine
End Function

Function CreateSQL(wrkTable As Excel.Worksheet)
    Dim strInsert As String
    Dim strColumn As String
    Dim aryColumn As Variant
    Dim strSQLCmd As String
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim blnInsert As Boolean
    
    i = 1
    strColumn = ""
    
    wrkTable.Activate
    wrkTable.Range("A1").Select
    
            sqlString = "INSERT INTO CSH_ACCTG_CFG VALUES(" + Chr(34) + ReportTypeCode + Chr(34) + "," + Chr(34) + CashAccountingCategoryCode + Chr(34) + ");"

    strInsert = "INSERT INTO [" + wrkTable.Name + "] ("
    
    Do Until IsEmpty(ActiveCell)
        'If InStr(LCase(ActiveCell.Text), "description") = 0 Then
        If ActiveCell.Comment Is Nothing Then
            strInsert = strInsert + "[" + Replace(ActiveCell.Text, "'", "''") + "],"
            strColumn = strColumn + CStr(ActiveCell.Column) + ","
        End If
        ActiveCell.Offset(0, 1).Select
    Loop
    strInsert = Left(strInsert, Len(strInsert) - 1) + ") VALUES ("
    strColumn = Left(strColumn, Len(strColumn) - 1)
    
    aryColumn = Split(strColumn, ",")
    
    'wrkTable.Range("A2").Select
    
    iRow = 2
    
    Do While wrkTable.Cells(iRow, 2) > 0
        blnInsert = True
        
        If LCase(wrkTable.Name) = "DenominationList" Then
            blnInsert = IIf(wrkTable.Cells(iRow, "D") = "Y", True, False)
        End If
        
        If blnInsert Then
            strSQLCmd = strInsert
            For iColumn = 0 To UBound(aryColumn)
                Select Case TypeName(wrkTable.Cells(iRow, CInt(aryColumn(iColumn))).Value)
                Case "Date"
                    strSQLCmd = strSQLCmd + "'" + sqlDateTime(wrkTable.Cells(iRow, CInt(aryColumn(iColumn)))) + "',"
                Case "String"
                    strSQLCmd = strSQLCmd + "'" + Trim(Replace(wrkTable.Cells(iRow, CInt(aryColumn(iColumn))), "'", "''")) + "',"
                Case Else
                    strSQLCmd = strSQLCmd + CStr(wrkTable.Cells(iRow, CInt(aryColumn(iColumn)))) + ","
                End Select
            Next
            
            strSQLCmd = Left(strSQLCmd, Len(strSQLCmd) - 1)
            strSQLCmd = strSQLCmd + ")"
            WriteLine strSQLCmd
        End If
        iRow = iRow + 1
    Loop
    
End Function

Public Function sqlDateTime(IN_DateTime As Date) As String
    sqlDateTime = “ '” & sqlDate(IN_DateTime) & ” ” & sqlTime(IN_DateTime) & “'”
End Function
