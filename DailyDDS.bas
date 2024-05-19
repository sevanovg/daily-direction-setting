Attribute VB_Name = "DailyDDS"
Option Explicit
Option Compare Text

''Sheets names in Daily DDS template
Public Const ESCALATIONS_SHEET_NAME As String = "Escalations"
Public Const ESCALATION_MATRIX_SHEET_NAME As String = "Escalation matrix"
Public Const DAILY_DDS_SHEET_NAME As String = "Daily DDS"
Public Const LINKS_SHEET_NAME As String = "Links"
Public Const HYPER_CARE_SHEET_NAME As String = "X-doc Hyper Care"
Public Const DDS_RESTATE_SHEET_NAME As String = "Restate"
Public Const LOG_SHEET_NAME As String = "Error log"
Public Const DAILY_ARCHIVE_SHEET_NAME As String = "Actions Archive"

''Table names in the daily DDS sheet
Public Const CFR_TABLE_NAME As String = "CFRTable"
Public Const SAMBC_TABLE_NAME As String = "SAMBCTable"
Public Const IN_PROCESS_TABLE_NAME As String = "InProcessTable"
Public Const MATRIX_TABLE_NAME As String = "Escalations"

''Location of rows in daily dds file
Public Const DAILY_ACTION_PLAN_ROW As Integer = 7
Public Const WEEKLY_ACTION_PLAN_ROW As Integer = 204

''Column numbers in Daily DDS CFR table
Private Const DAILY_DDS_CFR_TREND_COLUMN_NUMBER = 3
Private Const DAILY_DDS_CFR_MTD_COLUMN_NUMBER = 4
Private Const DAILY_DDS_CFR_FIRST_DAY_COLUMN_NUMBER = 5

''Column numbers in Daily DDS SAMBC table
Private Const DAILY_DDS_SAMBC_CUSTOMER_COLUMN_NUMBER = 1
Private Const DAILY_DDS_SAMBC_TREND_COLUMN_NUMBER = 3
Private Const DAILY_DDS_SAMBC_MTD_COLUMN_NUMBER = 4
Private Const DAILY_DDS_SAMBC_FIRST_DAY_COLUMN_NUMBER = 5

''Column numbers in the In Process table
Private Const IN_PROCESS_MTD_COLUMN_NUMBER = 4
Private Const IN_PROCESS_FIRST_DAY_COLUMN_NUMBER = 5

''Row numbers in In Process table (starting from 2nd row, header row doesn't count)
Private Const IN_PROCESS_ATTENDANCE_FIRST_ROW = 1
Private Const IN_PROCESS_ATTENDANCE_LAST_ROW = 14
Private Const IN_PROCESS_CSOW1_ROW = 15
Private Const IN_PROCESS_CSOW2_ROW = 20
Private Const IN_PROCESS_CSO_LAST_ROW = 24
Private Const IN_PROCESS_SAFETY_ROW = 25
Private Const IN_PROCESS_QUALITY_ROW = 26
Private Const IN_PROCESS_OTS_ROW = 27
Private Const IN_PROCESS_OTS_TRANSPORT_ROW = 28
Private Const IN_PROCESS_OTS_DC_ROW = 29
Private Const IN_PROCESS_OTD_ROW = 30
Private Const IN_PROCESS_OTD_NA_ROW = 31
Private Const IN_PROCESS_OTD_RD_ROW = 32
Private Const IN_PROCESS_BACKLOGS_FIRST_ROW = 33
Private Const IN_PROCESS_BACKLOGS_LAST_ROW = 40
Private Const IN_PROCESS_POSTPONED_FIRST_ROW = 41
Private Const IN_PROCESS_POSTPONED_LAST_ROW = 42
Private Const IN_PROCESS_TRANSPORT_ISSUES_ROW = 43
Private Const IN_PROCESS_DC_ISSUES_ROW = 44
Private Const IN_PROCESS_CATEGORY_FIRST_ROW = 45
Private Const IN_PROCESS_CATEGORY_LAST_ROW = 58
Private Const IN_PROCESS_RESTATEMENTS = 109

''Column numbers of the action plans
Public Const ACTION_PLAN_DATE_COLUMN = 2
Public Const ACTION_PLAN_ACTION_COLUMN = 3
Public Const ACTION_PLAN_OWNER_COLUMN = 11
Public Const ACTION_PLAN_DEADLINE_COLUMN = 12
Public Const ACTION_PLAN_STATUS_COLUMN = 13
Public Const ACTION_PLAN_COMMENT_COLUMN = 14

Public Const FOR_DISCUSSION_FALSE = 1 'Actions that should not be discussed



Public Function updateCfr(ByRef dailyDdsSht As Worksheet, ByRef infoPageCfrSheet As Worksheet)
'
' Function that updates CFR data in the daily DDS sheet
' Arguments:    dailyDdsSht: main sheet of daily DDS
'               infoPageCfrSheet: DAILY CFR sheet from infopage extract
'

''Reference the table with CFR data
    On Error GoTo error_table_reference
        Dim dailyDdsCfrTable As ListObject
        Set dailyDdsCfrTable = dailyDdsSht.ListObjects(CFR_TABLE_NAME) ' "EECAR CFR Status" table
    On Error GoTo 0
    
''Insert new columns for all missing CFR dates in the daily DDS sheet
    Dim lastCFRUpdate As Date
    lastCFRUpdate = DateValue(dailyDdsCfrTable.Range(1, DAILY_DDS_CFR_FIRST_DAY_COLUMN_NUMBER)) 'latest CFR update
    
    Dim dDate As Date
    For dDate = lastCFRUpdate + 1 To Date - 1
        Run insertNewColumn(dailyDdsCfrTable, dDate, 5, False, False)
    Next dDate
    
''Put CFR data to Daily DDS CFR table
    '' Iterate through all rows and columns of the CFR table. On error leave cell blank
        On Error GoTo errCfr
        Dim r As Integer, c As Integer
        For c = DAILY_DDS_CFR_MTD_COLUMN_NUMBER To dailyDdsCfrTable.Range.Columns.Count
            For r = 2 To dailyDdsCfrTable.Range.Rows.Count
                If dailyDdsCfrTable.Range(r, 1) <> "" Then
                
                    ''get CFR data for each needed cell via the helper function by key/date
                    Dim curDate As String, key As String
                    key = dailyDdsSht.Cells(dailyDdsCfrTable.Range.Rows(r).row, 1)
                    curDate = CStr(dailyDdsCfrTable.HeaderRowRange(c))
                    dailyDdsCfrTable.Range(r, c) = Utils.getCfr(infoPageCfrSheet, "", "", "", "", key, curDate)
                    
                End If
            Next r
        Next c
        On Error GoTo 0
    
''Update Trend formula to compate D-2 vs. MTD numbers in CFR table
    dailyDdsCfrTable.ListColumns(DAILY_DDS_CFR_TREND_COLUMN_NUMBER).DataBodyRange.FormulaR1C1 = "=RC[3]-RC[1]"
           
Exit Function

''Error handler for when CFR for date/category/customer combination wasn't found
errCfr:
    ''Leave blank and resume to next cell
    dailyDdsCfrTable.Range(r, c) = ""
    Resume Next
    
''Error handler for when CFR table can not be referenced
error_table_reference:
    ''Log error
    Debug.Print ("CFR table with name: " & CFR_TABLE_NAME & " cannot be found in sheet: " & dailyDdsSht.Name)

End Function

Public Function updateSambc(ByRef dailyDdsSht As Worksheet, ByRef infoPageCfrProxyWb As Workbook)
'
' Function that updates CFR proxy data for SAMBC customers
' Arguments:    dailyDdsSht: main sheet of daily DDS
'               infoPageCfrProxyWb: DAILY CFR proxy workbook (extract from infopage)
'

'' Reference EECAR SAMBC Status table in the daily DDS sheet
    On Error GoTo error_table_reference
        Dim dailyDdsSambcTable As ListObject
        Set dailyDdsSambcTable = dailyDdsSht.ListObjects(SAMBC_TABLE_NAME)
    On Error GoTo 0
    
''Insert new columns for all missing SAMBC dates in the weekly DDS sheet
    Dim lastSAMBCUpdate As Date, dDate As Date
    lastSAMBCUpdate = DateValue(dailyDdsSambcTable.Range(1, DAILY_DDS_SAMBC_FIRST_DAY_COLUMN_NUMBER)) 'latest SAMBC update
           
    For dDate = lastSAMBCUpdate + 1 To Date - 1
        Run insertNewColumn(dailyDdsSambcTable, dDate, 5, False, False)
    Next dDate
    
'' put SAMBC CFR proxy data to Daily DDS SAMBC table
    '' Iterate through all rows and columns of the SAMBC table. On error leave cell blank
        Dim c As Integer, r As Integer
        On Error GoTo errSambc
        For c = DAILY_DDS_SAMBC_MTD_COLUMN_NUMBER To dailyDdsSambcTable.Range.Columns.Count
            For r = 2 To dailyDdsSambcTable.Range.Rows.Count
                ''get CFR data for each needed cell via the helper function by customer key/date
                Dim custKey As String, unitOfMeasure As String, curDate As String
                custKey = dailyDdsSht.Cells(dailyDdsSambcTable.Range.Rows(r).row, 1)
                curDate = CStr(dailyDdsSambcTable.HeaderRowRange(c))
                unitOfMeasure = Right(dailyDdsSambcTable.Range(r, DAILY_DDS_SAMBC_CUSTOMER_COLUMN_NUMBER), 2)
                dailyDdsSambcTable.Range(r, c) = Utils.getProxyCfr(infoPageCfrProxyWb, curDate, unitOfMeasure, , , , custKey)
            Next r
        Next c
        On Error GoTo 0

''Update Trend formula to compate D-2 vs. MTD numbers in CFR table and SAMBC tablle in the daily dds sheet
    dailyDdsSambcTable.ListColumns(DAILY_DDS_SAMBC_TREND_COLUMN_NUMBER).DataBodyRange.FormulaR1C1 = "=RC[3]-RC[1]"

Exit Function

''Error handler for when CFR proxy for customer wasn't found
errSambc:
    ''Leave blank and resume to next cell
    dailyDdsSambcTable.Range(r, c) = ""
    Resume Next
    
''Error handler for when SAMBC table can not be referenced
error_table_reference:
    ''Log error
    Debug.Print ("SAMBC table with name: " & SAMBC_TABLE_NAME & " cannot be found in sheet: " & dailyDdsSht.Name)

End Function

Public Function updateInProcessMeasures(ByRef dailyDdsSht As Worksheet, csoInputSht As Worksheet, _
    dailyOtdSht As Worksheet, backlogSht As Worksheet)
'
' Function that updates table with In Process Measure in the daily DDS
'
'

'' Reference all needed tables in the daily DDS sheet
    On Error GoTo error_table_reference
        Dim inProcessTable As ListObject
        Set inProcessTable = dailyDdsSht.ListObjects(IN_PROCESS_TABLE_NAME)  ' "In Process Measures" table
    On Error GoTo 0

''Insert new column for In Process Measures table (only today's date if missing)

    Dim lastInProcessUpdate As Date
    lastInProcessUpdate = DateValue(inProcessTable.Range(1, IN_PROCESS_FIRST_DAY_COLUMN_NUMBER)) 'latest In Process Table date
    
    If lastInProcessUpdate < Date Then Run insertNewColumn(inProcessTable, Date, 5, False, True)
        
'Put CSO formulas
    
    On Error GoTo errCso
    
    'CSO ranges for putting formulas
    Dim csoW1MeasuresRng As Range, csoW2MeasuresRng As Range, csoW1ResultRng As Range, csoW2ResultRng As Range
    
    'Define CSO ranges for putting formulas
    Set csoW1MeasuresRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_CSOW1_ROW + 1, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_CSOW2_ROW - 1, inProcessTable.Range.Columns.Count))
    Set csoW2MeasuresRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_CSOW2_ROW + 1, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_CSO_LAST_ROW, inProcessTable.Range.Columns.Count))
    Set csoW1ResultRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_CSOW1_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_CSOW1_ROW, inProcessTable.Range.Columns.Count))
    Set csoW2ResultRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_CSOW2_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_CSOW2_ROW, inProcessTable.Range.Columns.Count))
    
    
    'CSO W1 measures
    Dim sCell As Range
    For Each sCell In csoW1MeasuresRng
        sCell = Application.WorksheetFunction.Index( _
            csoInputSht.Cells, _
            Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), Range(csoInputSht.Cells(1, 1), csoInputSht.Cells(7, 1)), 0), _
            Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), csoInputSht.Rows(1), 0))
            
        'Modify formula in case didn't work (remove CLng)
        If sCell = "" Then
        sCell = Application.WorksheetFunction.Index( _
            csoInputSht.Cells, _
            Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), Range(csoInputSht.Cells(1, 1), csoInputSht.Cells(7, 1)), 0), _
            Application.WorksheetFunction.Match(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1)), csoInputSht.Rows(1), 0))
        End If
        
    Next sCell
    
    'CSO W2 measures
    For Each sCell In csoW2MeasuresRng
        sCell = Application.WorksheetFunction.Index( _
            csoInputSht.Cells, _
            Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), Range(csoInputSht.Cells(8, 1), csoInputSht.Cells(14, 1)), 0) + 7, _
            Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), csoInputSht.Rows(1), 0))
            
        'Modify formula in case didn't work (remove CLng)
        If sCell = "" Then
            sCell = Application.WorksheetFunction.Index( _
            csoInputSht.Cells, _
            Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), Range(csoInputSht.Cells(8, 1), csoInputSht.Cells(14, 1)), 0) + 7, _
            Application.WorksheetFunction.Match(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1)), csoInputSht.Rows(1), 0))
        End If
            
    Next sCell
    
'' Put PD formulas
    
    On Error GoTo errTransportDds

    'Define ranges for PD data
    Dim otsRng As Range, otsDcRng As Range, otsTransportRng As Range, otdRng As Range, otdRdRng As Range, _
        otdNaRng As Range, backlogRng As Range, postponedRng As Range
    
    Set otsRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTS_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTS_ROW, inProcessTable.Range.Columns.Count))
    Set otsDcRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTS_DC_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTS_DC_ROW, inProcessTable.Range.Columns.Count))
    Set otsTransportRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTS_TRANSPORT_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTS_TRANSPORT_ROW, inProcessTable.Range.Columns.Count))
    Set otdRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTD_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTD_ROW, inProcessTable.Range.Columns.Count))
    Set otdRdRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTD_RD_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTD_RD_ROW, inProcessTable.Range.Columns.Count))
    Set otdNaRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_OTD_NA_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_OTD_NA_ROW, inProcessTable.Range.Columns.Count))
    Set backlogRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_BACKLOGS_FIRST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_BACKLOGS_LAST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2))
    Set postponedRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_POSTPONED_FIRST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2), inProcessTable.DataBodyRange(IN_PROCESS_POSTPONED_LAST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER + 2))
    'OTS total formula
    'For Each sCell In otsRng
        'sCell = Application.WorksheetFunction.Index( _
        'transportationDdsSht.Cells, _
        'Application.WorksheetFunction.Match("OTS", transportationDdsSht.Columns(1), 0), _
        'Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), transportationDdsSht.Rows(3), 0))
    'Next sCell
    
    'OTS Transportation formula
    'For Each sCell In otsTransportRng
        'sCell = Application.WorksheetFunction.Index( _
        'transportationDdsSht.Cells, _
        'Application.WorksheetFunction.Match("Transport", transportationDdsSht.Columns(1), 0), _
        'Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), transportationDdsSht.Rows(3), 0))
    'Next sCell
    
    'OTS DC formula
    'For Each sCell In otsDcRng
        'sCell = Application.WorksheetFunction.Index( _
        'transportationDdsSht.Cells, _
        'Application.WorksheetFunction.Match("DC", transportationDdsSht.Columns(1), 0), _
        'Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), transportationDdsSht.Rows(3), 0))
    'Next sCell
    
    On Error GoTo errOtd
    
    'OTD TOTAL formula
    For Each sCell In otdRng
        sCell = Application.WorksheetFunction.Index( _
        dailyOtdSht.Cells, _
        Application.WorksheetFunction.Match("OTD Total", dailyOtdSht.Columns(1), 0), _
        Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), dailyOtdSht.Rows(3), 0))
    Next sCell
    
    'OTD NA formula
    For Each sCell In otdNaRng
        sCell = Application.WorksheetFunction.Index( _
        dailyOtdSht.Cells, _
        Application.WorksheetFunction.Match("DD", dailyOtdSht.Columns(1), 0), _
        Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), dailyOtdSht.Rows(3), 0))
    Next sCell
    
    'OTD RD formula
    For Each sCell In otdRdRng
        sCell = Application.WorksheetFunction.Index( _
        dailyOtdSht.Cells, _
        Application.WorksheetFunction.Match("RD", dailyOtdSht.Columns(1), 0), _
        Application.WorksheetFunction.Match(CLng(DateValue(inProcessTable.Range(1, sCell.Column - inProcessTable.Range(1, 1).Column + 1))), dailyOtdSht.Rows(3), 0))
    Next sCell
    
    'Backlog formula
    Dim rowNum As Integer, backlogs As Integer, postponed As Integer
    For Each sCell In backlogRng
        rowNum = Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), backlogSht.Columns(1), 0)
        backlogs = Val(backlogSht.Cells(rowNum, 9)) + Val(backlogSht.Cells(rowNum, 10)) + Val(backlogSht.Cells(rowNum, 11))
        postponed = Val(backlogSht.Cells(rowNum, 3)) + Val(backlogSht.Cells(rowNum, 4)) + Val(backlogSht.Cells(rowNum, 5))
        sCell.NumberFormat = "@"
        sCell = CStr(backlogs) & " / " & CStr(postponed) 'fill the data to the cell
    Next sCell
    
    'Postponed formula
    For Each sCell In postponedRng
        rowNum = Application.WorksheetFunction.Match(inProcessTable.Range(sCell.row - inProcessTable.HeaderRowRange.row + 1, 2), backlogSht.Columns(1), 0)
        sCell.NumberFormat = "@"
        postponed = Val(backlogSht.Cells(rowNum, 3)) + Val(backlogSht.Cells(rowNum, 4)) + Val(backlogSht.Cells(rowNum, 5))
        sCell = CStr(postponed)
    Next sCell
    
    On Error GoTo 0
    

''Update MTD formulas in the In Process table

    ''Find first day of current month that is present in the In Process table
    Dim firstDayInCurrentMonth As Date
    firstDayInCurrentMonth = DateValue(inProcessTable.HeaderRowRange(1, IN_PROCESS_FIRST_DAY_COLUMN_NUMBER))
    For Each sCell In inProcessTable.HeaderRowRange.Cells
        On Error Resume Next
        If DateValue(sCell) < firstDayInCurrentMonth And month(DateValue(sCell)) = month(Date) Then
            firstDayInCurrentMonth = DateValue(sCell)
        ElseIf DateValue(sCell) < firstDayInCurrentMonth And month(DateValue(sCell)) <> month(Date) Then
            Exit For 'exit the loop once iteration has reached the next month
        End If
        On Error GoTo 0
    Next sCell
    
    
    ''Reference ranges
    Dim attendanceRng As Range, csoRng As Range, safetyQualityRng As Range, pdRng As Range, categoryRng As Range
    Set attendanceRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_ATTENDANCE_FIRST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER), inProcessTable.DataBodyRange(IN_PROCESS_ATTENDANCE_LAST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER))
    Set csoRng = Union(inProcessTable.DataBodyRange(IN_PROCESS_CSOW1_ROW, IN_PROCESS_MTD_COLUMN_NUMBER), inProcessTable.DataBodyRange(IN_PROCESS_CSOW2_ROW, IN_PROCESS_MTD_COLUMN_NUMBER))
    Set safetyQualityRng = Union(inProcessTable.DataBodyRange(IN_PROCESS_SAFETY_ROW, IN_PROCESS_MTD_COLUMN_NUMBER), inProcessTable.DataBodyRange(IN_PROCESS_QUALITY_ROW, IN_PROCESS_MTD_COLUMN_NUMBER))
    Set pdRng = Union(inProcessTable.DataBodyRange(IN_PROCESS_TRANSPORT_ISSUES_ROW, IN_PROCESS_MTD_COLUMN_NUMBER), inProcessTable.DataBodyRange(IN_PROCESS_DC_ISSUES_ROW, IN_PROCESS_MTD_COLUMN_NUMBER))
    'Set categoryRng = Range(inProcessTable.DataBodyRange(IN_PROCESS_CATEGORY_FIRST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER), inProcessTable.DataBodyRange(IN_PROCESS_CATEGORY_LAST_ROW, IN_PROCESS_MTD_COLUMN_NUMBER))
    
    ''TODO add ranges for newely added measures
    
    ''Attendance/CSO MTD formula
    On Error Resume Next
    Union(attendanceRng, csoRng).Formula = _
        "=SUM(" & inProcessTable.Name & "[@[" & Date & "]:[" & firstDayInCurrentMonth & "]])" & _
        "/COUNT(" & inProcessTable.Name & "[@[" & Date & "]:[" & firstDayInCurrentMonth & "]])"
    
    ''Safety/Quality/PD issues/Category issues formula
    Union(safetyQualityRng, pdRng).Formula = _
        "=SUM(" & inProcessTable.Name & "[@[" & Date & "]:[" & firstDayInCurrentMonth & "]])"
    On Error GoTo 0
    
''Fill data for pending restatements (SU sum of all pending restatements)
    Dim dailyDdsWb As Workbook
    Set dailyDdsWb = dailyDdsSht.Parent
    inProcessTable.DataBodyRange(IN_PROCESS_RESTATEMENTS, IN_PROCESS_FIRST_DAY_COLUMN_NUMBER).Value = _
                Application.WorksheetFunction.Sum(dailyDdsWb.Sheets(DDS_RESTATE_SHEET_NAME).Columns(RESTATE_QTY_COLUMN_NUMBER))

Exit Function

errCso:
Resume Next
errTransportDds:
Resume Next
errOtd:
Resume Next

''Error handler for when SAMBC table can not be referenced
error_table_reference:
    ''Log error
    Debug.Print ("In Process table with name: " & IN_PROCESS_TABLE_NAME & " cannot be found in sheet: " & dailyDdsSht.Name)

End Function

Public Function archiveActions(dailyDdsSht As Worksheet, dailyActionsArchiveSht As Worksheet)
'
' Function that archives done actions in the daily dds
' Arguments:    dailyDdsSht: main sheet of daily DDS
'               dailyActionsArchiveSht: sheet where actions should be archived to
'

    '' Find last rows
    Dim lastRowDDSActions As Integer, lastRowDailyArchive As Long
    lastRowDDSActions = dailyDdsSht.Range("B" & DAILY_ACTION_PLAN_ROW).End(xlDown).row 'last row in daily dds actions section
    lastRowDailyArchive = dailyActionsArchiveSht.Range("B1").End(xlDown).row 'last row in "Actions Archive" sheet
    
    '' Set the range to be archived and to where it should be archived
    Dim dailyDdsActionsRng As Range, archiveRng As Range
    
    Set dailyDdsActionsRng = Range(dailyDdsSht.Cells(DAILY_ACTION_PLAN_ROW + 2, ACTION_PLAN_DATE_COLUMN), _
                                    dailyDdsSht.Cells(lastRowDDSActions, ACTION_PLAN_COMMENT_COLUMN))
                                    
    Set archiveRng = dailyActionsArchiveSht.Cells(lastRowDailyArchive + 1, 2)
    
    '' Execute archiving via the helper function
    Utils.archiveActions dailyDdsActionsRng, _
                    archiveRng, _
                    (ACTION_PLAN_STATUS_COLUMN - ACTION_PLAN_DATE_COLUMN + 1)

End Function

Private Function insertNewColumn(targetTable As ListObject, headerDate As Date, columnPosition As Integer, copyFromPreviousColumn As Boolean, copyFormat As Boolean)
'
' Function that adds new column to the table with the specified date header and fills the formula if needed
'

'
Dim newColumn As ListColumn
Set newColumn = targetTable.ListColumns.Add(Position:=columnPosition) 'add and reference new column
newColumn.Name = headerDate 'put header

'copy formulas to the new column if needed
If copyFromPreviousColumn Then
    targetTable.ListColumns(columnPosition + 1).DataBodyRange.AutoFill Destination:=Union(newColumn.DataBodyRange, targetTable.ListColumns(columnPosition + 1).DataBodyRange)
End If

'copy formatting from the right column if needed
If copyFormat Then
    targetTable.ListColumns(columnPosition + 1).DataBodyRange.Copy
    newColumn.DataBodyRange.PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End If

End Function
