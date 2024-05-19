Attribute VB_Name = "WeeklyDDS"
Option Explicit
Option Compare Text

''day on which weekly DDS is held (1 to 7, Mon to Sun)
Private Const WEEKLY_DDS_DAY As Integer = 2

''day on which weekly DDS file should be cleaned (1 to 7, Mon to Sun)
Private Const WEEKLY_DDS_CLEAN_UP_DAY As Integer = 1

''File path for weekly DDS file
Public Const WEEKLY_DDS_FILEPATH As String = _
    "https://pgone.sharepoint.com/sites/eecarweeklydds/Shared Documents/SNO Weekly DDS.xlsx"

''Sheet names in weekly DDS file
Public Const WEEKLY_DDS_MAIN_SHEET As String = "Weekly SMO DDS Template 2.0"
Public Const WEEKLY_DDS_ARCHIVE_SHEET As String = "Archive"

''Row/column numbers in the weekly DDS sheet
Private Const ACTION_PLAN_ISSUE_COLUMN As Integer = 2
Private Const ACTION_PLAN_STATUS_COLUMN As Integer = 10
Private Const ACTION_PLAN_COMMENT_COLUMN As Integer = 11

Private Const WEEKLY_DDS_CFR_FIRST_ROW As Integer = 3
Private Const WEEKLY_DDS_CFR_LAST_ROW As Integer = 12
Private Const ACTIONS_FIRST_ROW As Integer = 103

''Table names
Private Const CFR_TABLE_NAME As String = "CFR_Table"
Private Const CFR_OUTLOOK_TABLE_NAME As String = "CFR_Outlook_table"
Private Const SEGMENTATION_TABLE_NAME As String = "Segmentation_table"
Private Const IBA_TABLE_NAME As String = "IBA_table"
Private Const TRANSPORTATION_TABLE_NAME As String = "Transportation_table"
Private Const SAMBC_OUTLOOK_TABLE_NAME As String = "SAMBC_Table"
Private Const BI_TABLE_NAME As String = "BI_table"
Private Const INITIATIVES_TABLE_NAME As String = "Initiatives_table"
Private Const NPI_TABLE_NAME As String = "NPI_Table"
Private Const IDE_TABLE_NAME As String = "IDE_Table"
Private Const SPO_TABLE_NAME As String = "SPO_table"

''Column numbers of all tables
Private Const MEASURE_COLUMN As Integer = 1
Private Const OWNER_COLUMN As Integer = 2
Private Const TARGET_COLUMN As Integer = 3
Private Const EECAR_TOTAL_COLUMN As Integer = 4
Private Const REASON_COMMENT_COLUMN As Integer = 10
Private Const ACTION_HELP_COLUMN As Integer = 11

''Column numbers of SAMBC Table
Private Const SAMBC_TABLE_CUSTOMER_COLUMN As Integer = 1
Private Const SAMBC_TABLE_CFR_RESULT_COLUMN As Integer = 4

''Key words
Private Const ACTION_PLAN_HEADER As String = "Action points from last week"



Public Function updateCfr(ByRef weeklyDdsSht As Worksheet, ByRef infoPageCfrSheet As Worksheet)
'
' Function that updates CFR data in the weekly DDS sheet
' Arguments:    weeklyDdsSht: main sheet of weekly DDS
'               infoPageCfrSheet: DAILY CFR sheet from infopage extract
'

'' Get Table with CFR result
Dim cfrTable As ListObject
Set cfrTable = weeklyDdsSht.ListObjects(CFR_TABLE_NAME)

''Fill CFR result to all needed columns via the helper function
Dim cell As Range
For Each cell In cfrTable.DataBodyRange


    ''Relative position of a cell in the table
    Dim cellRow As Integer, cellColumn As Integer
    cellRow = cell.row - cfrTable.HeaderRowRange.row + 1
    cellColumn = cell.Column - cfrTable.ListColumns(1).Range.Column + 1
    
    ''Only columns where CFR data is needed
    If cellColumn >= EECAR_TOTAL_COLUMN And cellColumn < REASON_COMMENT_COLUMN Then
    
        ''Clean old data if any
        cell.Value2 = ""

        ''Determine category and geography
        Dim category As String, geography As String
        category = cfrTable.Range(cellRow, MEASURE_COLUMN)
        geography = cfrTable.HeaderRowRange.Columns(cellColumn)
    
        ''Get the cfr for all needed columns
        cell.Value2 = Utils.getCfr(infoPageCfrSheet, category, geography, "", "", "", MONTH_TO_DATE_LABEL)
    
    End If

Next cell
    
        
End Function

Public Function updateSambc(ByRef weeklyDdsSht As Worksheet, ByRef infoPageCfrProxyWb As Workbook)
'
' Function that updates CFR proxy data for SAMBC customers
' Arguments:    weeklyDdsSht: main sheet of weekly DDS
'               infoPageCfrProxyWb: DAILY CFR proxy workbook (extract from infopage)
'

'' Get Table with SAMBC Outlook
Dim sambcTable As ListObject
Set sambcTable = weeklyDdsSht.ListObjects(SAMBC_OUTLOOK_TABLE_NAME)

Dim row As Range
For Each row In sambcTable.DataBodyRange.Rows
    
    ''Read key and unit of measure
    Dim unitOfMeasure As String, custKey As String
    custKey = weeklyDdsSht.Cells(row.row, 1)
    unitOfMeasure = Right(row.Columns(SAMBC_TABLE_CUSTOMER_COLUMN), 2)
    
    ''Clean old data if needed
    row.Columns(SAMBC_TABLE_CFR_RESULT_COLUMN) = ""
    
    ''get CFR MTD for each customer
    row.Columns(SAMBC_TABLE_CFR_RESULT_COLUMN) = _
        Utils.getProxyCfr(infoPageCfrProxyWb, MONTH_TO_DATE_LABEL, unitOfMeasure, "", "", "", custKey)

Next row


End Function

Public Function clearWeeklyDds(ByRef weeklyDdsSht As Worksheet)
'
' Function that clears weekly dds template from old data
' Arguments:    weeklyDdsSht: main sheet of weekly DDS
'

''Range to be cleared
    Dim clearRng As Range

'' Iterate through all tables in the sheet
    Dim table As ListObject, col As Integer
    For Each table In weeklyDdsSht.ListObjects
        On Error GoTo error_results_range
        
        ''Iterate through all the columns with results data and add data from the column to the range to be cleared
        For col = EECAR_TOTAL_COLUMN To REASON_COMMENT_COLUMN - 1
            If Not clearRng Is Nothing Then
                Set clearRng = Union(clearRng, table.ListColumns(col).DataBodyRange)
            Else
                Set clearRng = table.ListColumns(col).DataBodyRange
            End If
        Next col
        
        ''Add action/help column to the range to be cleared
        Set clearRng = Union(clearRng, table.ListColumns(ACTION_HELP_COLUMN).DataBodyRange)
        
        On Error GoTo 0
    Next table
    
'' Clear the range
    clearRng.ClearContents

Exit Function

''Error handler for when a column fron ListObject can't be added to the results range
error_results_range:
    Debug.Print ("Error adding column to the range to be cleared. ListObject: " & table.Name)
    Resume Next

End Function
Public Function isWeeklyDdsDay() As Boolean
'
' Function that returns true if today is a weekly DDS day
'
'

If Weekday(Date, vbMonday) = WEEKLY_DDS_DAY Then
    isWeeklyDdsDay = True
Else
    isWeeklyDdsDay = True 'False - temporarily set to True to test update every day
End If

End Function

Public Function isWeeklyDdsCleanUpDay() As Boolean
'
' Function that returns true if today is a weekly DDS clean-up day
'
'

If Weekday(Date, vbMonday) = WEEKLY_DDS_CLEAN_UP_DAY Then
    isWeeklyDdsCleanUpDay = True
Else
    isWeeklyDdsCleanUpDay = False
End If

End Function

Public Function archiveActions(ByRef weeklyDdsSheet As Worksheet)
'
' Function that archives done actions in the weekly dds
' Arguments:    weeklyDdsSht: main sheet of weekly DDS
'

''Reference the archive sheet
Dim archiveSht As Worksheet
Set archiveSht = weeklyDdsSheet.Parent.Sheets(WEEKLY_DDS_ARCHIVE_SHEET)

''Find first/last row of the action plan
Dim actionsFirstRow As Integer, actionsLastRow As Integer
actionsFirstRow = weeklyDdsSheet.Columns(ACTION_PLAN_ISSUE_COLUMN).Find(ACTION_PLAN_HEADER).row + 2
actionsLastRow = weeklyDdsSheet.Cells(weeklyDdsSheet.Rows.Count, ACTION_PLAN_ISSUE_COLUMN).End(xlUp).row

''Reference the range containing all actions
Dim actionsRng As Range
Set actionsRng = Range(weeklyDdsSheet.Cells(actionsFirstRow, ACTION_PLAN_ISSUE_COLUMN), _
                        weeklyDdsSheet.Cells(actionsLastRow, ACTION_PLAN_COMMENT_COLUMN))

''Find last row in the archive sheet
Dim archiveLastRow As Long
archiveLastRow = archiveSht.Cells(archiveSht.Rows.Count, 1).End(xlUp).row

''Define the left/top cell where the actions need to be copied
Dim pasteRng As Range
Set pasteRng = archiveSht.Cells(archiveLastRow + 1, 1)

''Find the relative position of a status column in the action plan range
Dim statusColumnRelativePosition As Integer
statusColumnRelativePosition = ACTION_PLAN_STATUS_COLUMN - ACTION_PLAN_ISSUE_COLUMN + 1

''Execute the archiving via the helper function
Utils.archiveActions actionsRng, pasteRng, statusColumnRelativePosition


End Function
