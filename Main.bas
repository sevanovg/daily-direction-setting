Attribute VB_Name = "Main"

Option Compare Text
Option Explicit

''Sheet names of the source excel files
Public Const TRANSPORTATION_DDS_SHEET_NAME As String = "Measures"
Public Const CSO_INPUT_SHEET_NAME As String = "CSO"
Public Const DAILY_CFR_SHEET_NAME As String = "report"
Public Const DAILY_OTD_SHEET_NAME As String = "DD.RD"
Public Const BACKLOG_SHEET_NAME As String = "Backlog update"

''Range addresses to store temporary filepaths
Public Const TRANSPORTATION_DDS_PATH As String = "A9990"
Public Const CSO_INPUT_PATH As String = "A9991"
Public Const DAILY_CFR_PATH As String = "A9992"
Public Const DAILY_PROXY_PATH As String = "A9993"
Public Const OTD_PATH As String = "A9994"

''column numbers on escalation sheet
Public Const DATE_COLUMN As Integer = 1
Public Const ESC_BY_COLUMN As Integer = 2
Public Const ESC_TO_COLUMN As Integer = 3
Public Const FPC_COLUMN As Integer = 4
Public Const DESCR_COLUMN As Integer = 5
Public Const PLANT_COLUMN As Integer = 6
Public Const HELP_COLUMN As Integer = 7
Public Const FOR_DISCUSSION_COLUMN As Integer = 8
Public Const CHECKED_COLUMN As Integer = 9

''number of columns in customers/categories escalations section
Public Const ESC_COLUMNS_NUM As Integer = 7

''Checked marker on escalations list
Public Const CHECKED_MARKER As String = "Checked"

''Location of files on sharepoint
Private Const ESCALATIONS_FILE_LOCATION As String = "https://pgone-my.sharepoint.com/personal/novgorodtsev_vm_pg_com/Documents/SNO DDS Escalations.xlsx"
Private Const RESTATEMENTS_FILE_LOCATION As String = "https://pgone.sharepoint.com/sites/eecarsnodds/Shared Documents/Restatement Process Template 1.02.xlsx"
Private Const OTD_FILE_LOCATION As String = "https://pgone.sharepoint.com/sites/a_pd/Shared Documents/08. Transport Operations/01. Operational/Daily OTD Tracking May'17 (+DD.RD).xlsx"

''Sheet names in restatements file
Private Const FOR_RESTATE_SHEET_NAME As String = "SMO Template"
Private Const RESTATE_ARCHIVE_SHEET_NAME As String = "Archive"

''Sheet names in weekly DDS file
Public Const WEEKLY_DDS_MAIN_SHEET As String = "Weekly SMO DDS Template 2.0"

''Approval status strings in restatements file
Public Const APPROVED_STATUS As String = "APPROVED"
Public Const DENIED_STATUS As String = "DENIED"
Public Const PENDING_STATUS As String = "PENDING"

''Column numbers in restatements file
Public Const APPROVAL_STATUS_COLUMN_NUMBER As Integer = 9
Public Const RESTATE_QTY_COLUMN_NUMBER As Integer = 3

'' Other
Public Const ESCALATION_CHAIN_MARKER = "EscMarker"
Public Const MONTH_TO_DATE_LABEL As String = "MTD"
Public Const MISSING_DATA_LABEL As String = "missing"

Sub DataUpdate()
'
' Macro that updates daily/weekly DDS data
'
'

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'' Call the userform so that the user specifies the paths to the helper files
    FileSelectionForm.Show

'' Reference daily DDS workbook and all needed sheets
    Dim dailyDdsWb As Workbook
    Dim dailyDdsSht As Worksheet, dailyActionsArchiveSht As Worksheet, escalationsSht As Worksheet
    
    Set dailyDdsWb = ActiveWorkbook
    Set dailyDdsSht = dailyDdsWb.Sheets(DAILY_DDS_SHEET_NAME)
    Set dailyActionsArchiveSht = dailyDdsWb.Sheets(DAILY_ARCHIVE_SHEET_NAME)

'' Read the file paths provided by the user and open helper workbooks based on the specified paths
    Dim transportationdailyDdsWb As Workbook, csoInputWb As Workbook, dailyCfrWb As Workbook, _
        dailyOtdWb As Workbook, dailyProxyWb As Workbook
    
    On Error GoTo error_open_helper_wb
        'Set transportationdailyDdsWb = Workbooks.Open(dailyDdsSht.Range(TRANSPORTATION_DDS_PATH).Value2)
        Set csoInputWb = Workbooks.Open(dailyDdsSht.Range(CSO_INPUT_PATH).Value2)
        Set dailyCfrWb = Workbooks.Open(dailyDdsSht.Range(DAILY_CFR_PATH).Value2)
        Set dailyOtdWb = Workbooks.Open(Filename:=OTD_FILE_LOCATION, ReadOnly:=True)
        Set dailyProxyWb = Workbooks.Open(dailyDdsSht.Range(DAILY_PROXY_PATH).Value2)
    On Error GoTo 0

'' Reference all needed sheets in the hepler workbooks
    Dim transportationDdsSht As Worksheet, csoInputSht As Worksheet, dailyCfrSht As Worksheet, _
        dailyOtdSht As Worksheet, backlogSht As Worksheet
    
    On Error GoTo error_missing_sheet_helper_wb
        Set csoInputSht = csoInputWb.Sheets(CSO_INPUT_SHEET_NAME)
        Set dailyCfrSht = dailyCfrWb.Sheets(DAILY_CFR_SHEET_NAME)
        Set dailyOtdSht = dailyOtdWb.Sheets(DAILY_OTD_SHEET_NAME)
        'Set transportationDdsSht = transportationdailyDdsWb.Sheets(TRANSPORTATION_DDS_SHEET_NAME)
        Set backlogSht = dailyOtdWb.Sheets(BACKLOG_SHEET_NAME)
    On Error GoTo 0

''Update data tables in Daily DDS
    DailyDDS.updateCfr dailyDdsSht, dailyCfrSht
    DailyDDS.updateSambc dailyDdsSht, dailyProxyWb
    DailyDDS.updateInProcessMeasures dailyDdsSht, csoInputSht, dailyOtdSht, backlogSht

'' Archive done actions of the daily DDS
    DailyDDS.archiveActions dailyDdsSht, dailyActionsArchiveSht

'' Update restatements
    On Error GoTo error_restatements
        ProcessRestatements dailyDdsWb
    On Error GoTo 0

'' Update escalations
    On Error GoTo error_escalations
        Set escalationsSht = DownloadEscalations(dailyDdsWb)
    On Error GoTo 0

'' Update the matrix
    On Error GoTo error_update_matrix
        BuildMatrix dailyDdsSht, escalationsSht
    On Error GoTo 0
    
GoTo skip_weekly_dds_processing ''Skip processing of weekly DDS as not needed anymore

'' If today is Weekly DDS day then update Weekly DDS data
    If WeeklyDDS.isWeeklyDdsDay() Or WeeklyDDS.isWeeklyDdsCleanUpDay() Then
            
        ''Open weekly DDS workbook and reference necessary sheets
        On Error GoTo error_open_weekly_dds
        Dim weeklyDdsWb As Workbook
        Dim weeklyDdsSht As Worksheet
        Set weeklyDdsWb = Workbooks.Open(WEEKLY_DDS_FILEPATH)
        On Error Resume Next
            weeklyDdsWb.LockServerFile
        On Error GoTo error_open_weekly_dds
        Set weeklyDdsSht = weeklyDdsWb.Sheets(WEEKLY_DDS_MAIN_SHEET)
        
        ''On DDS clean-up day clean up the sheet and archive actions
        On Error GoTo error_weekly_dds_cleanup
        If WeeklyDDS.isWeeklyDdsCleanUpDay() Then
            'temporary turn off clean up
            'WeeklyDDS.clearWeeklyDds weeklyDdsSht
            WeeklyDDS.archiveActions weeklyDdsSht
        End If
        
        ''On either day update CFR/SAMBC data
        On Error GoTo error_updating_weekly_cfr
        WeeklyDDS.updateCfr weeklyDdsSht, dailyCfrSht
        WeeklyDDS.updateSambc weeklyDdsSht, dailyProxyWb
        
        On Error GoTo 0
                       
    End If
  
skip_weekly_dds_processing:

''Close all helper workbooks
    On Error Resume Next
        csoInputWb.Close SaveChanges:=False
        dailyCfrWb.Close SaveChanges:=False
        dailyOtdWb.Close SaveChanges:=False
        dailyProxyWb.Close SaveChanges:=False
    On Error GoTo 0

''Save/close weekly dds workbook
    If Not weeklyDdsWb Is Nothing Then weeklyDdsWb.Close SaveChanges:=True
        
''Activate daily DDS sheet
    dailyDdsSht.Activate

Application.DisplayAlerts = True
Application.ScreenUpdating = True

Exit Sub

'' Error handler for when weekly DDS file can't be opened
error_open_weekly_dds:
    Debug.Print ("Error opening weekly DDS file")
    Resume skip_weekly_dds_processing
   
'' Error handler for when weekly DDS CFR/SAMBC can't be opened
error_updating_weekly_cfr:
    Debug.Print ("Error updating CFR/SAMBC in the weekly DDS file")
    Resume Next

'' Error handler for when weekly DDS sheet can't be properly cleaned
error_weekly_dds_cleanup:
    Debug.Print ("Error cleaning weekly DDS sheet")
    Resume Next
    
'' Error handler for when some of the helper files can't be opened
error_open_helper_wb:
    
    ''Display the message
    MsgBox ("Один или несколько указанных файлов невозможно открыть. Убедитесь, что адреса указаны верно")
    
    ''Reopen file selection window
    FileSelectionForm.Show
    
    ''Resume to retry opening
    Resume
   
'' Error handler for when one of the sheets of the helper workbooks can't be found
error_missing_sheet_helper_wb:
   
   ''Display warning message and keep executing
    MsgBox ("Одна или несколько страниц изменили своё название. Макрос будет продолжен, но необходимо уведомить DTLM команду")
    Resume Next
    
'' Error handler for when one reference to daily dds tables is not working
error_daily_table_reference:
    Debug.Print ("Error referencing tables in the daily DDS sheet " & Err.Source)
    Resume Next
    
'' Error handler for when restatements could not be properly updated
error_restatements:
    Debug.Print ("Error updating restatements")
    Resume Next
    
'' Error handler for when escalations could not be properly updated
error_escalations:
    Debug.Print ("Error with escalations")
    Resume Next

'' Error handler for when escalations matrix could not be properly updated
error_update_matrix:
    Debug.Print ("Error with building the matrix")
    Resume Next


End Sub


Function BuildMatrix(ByRef dailyDdsSht As Worksheet, ByRef escalationsSht As Worksheet)
'
' Funtion that builds an escalation matrix for the daily DDS
' Arguments:    dailyDdsSht - a worksheet where the matrix should be build
'               escalationsSht - a worksheet which contains escalations to be built into the matrix
'
    
''Clean old escalations from the matrix
    CleanMatrix dailyDdsSht

'' Reference the range with the matrix (top left cell of the matrix)
    Dim matrixTbl As ListObject
    Set matrixTbl = dailyDdsSht.ListObjects(DailyDDS.MATRIX_TABLE_NAME)
    
'' Find last row in the escalations sheet
    Dim lastRowEscalations As Integer
    lastRowEscalations = escalationsSht.Cells(escalationsSht.Rows.Count, "A").End(xlUp).row
   
'' Collection of escalation chains to avoid repetitive arrows
    Dim escalationCollection As New Collection
    
'' Iterate through all rows with escalations and add today's relevant escalations to the matrix
    Dim i As Integer
    For i = 2 To lastRowEscalations 'iterate through all escalation rows
    
        Dim escalationDate As Date
        Dim escBy As String, escTo As String, fpc As String, plant As String, descr As String, helpNeeded As String
        Dim forDiscussion As Integer 'specifies whether escalation is for discussion or not (0 - no, any other - yes)
        
        'store all escalation details in variables
        escalationDate = escalationsSht.Cells(i, DATE_COLUMN)
        escBy = Trim(escalationsSht.Cells(i, ESC_BY_COLUMN))
        escTo = Trim(escalationsSht.Cells(i, ESC_TO_COLUMN))
        fpc = escalationsSht.Cells(i, FPC_COLUMN)
        descr = escalationsSht.Cells(i, DESCR_COLUMN)
        plant = escalationsSht.Cells(i, PLANT_COLUMN)
        helpNeeded = escalationsSht.Cells(i, HELP_COLUMN)
        forDiscussion = escalationsSht.Cells(i, FOR_DISCUSSION_COLUMN)
               
        If (escalationDate = Date) And IndexOf(escalationCollection, escBy & escTo) = 0 And _
        forDiscussion <> DailyDDS.FOR_DISCUSSION_FALSE Then 'consider only today's escalations and not repetetive
            
                escalationCollection.Add (escBy & escTo) 'add the chain to the collection
            
                Dim escalatorRng As Range, victimRng As Range 'range reference for escalator and victim
                Dim escalationCell  As Range 'intersection cell of escalator and victim
                
                Dim escalationDirection As Boolean 'true = row to col, false = col to row
                
                Set escalatorRng = matrixTbl.Range.Columns(1).Find(escBy)  'look for escalator in rows
                Set victimRng = matrixTbl.Range.Rows(1).Find(escTo)
                escalationDirection = True
                
                'Set escalatorRng = matrixRng.EntireColumn.Find(escBy, After:=matrixRng)  'look for escalator in rows
                'Set victimRng = matrixRng.EntireRow.Find(escTo, matrixRng) 'look for victim in columns
                'escalationDirection = True
                
                On Error Resume Next
                
                'If (escalatorRng Is Nothing) Or (victimRng Is Nothing) Or (escalatorRng.row < matrixRng.row) Then 'if escalator is not found in rows or victim not found in columns than swap
                    'Set escalatorRng = matrixRng.EntireRow.Find(escBy, matrixRng)
                    'Set victimRng = matrixRng.EntireColumn.Find(escTo, matrixRng)
                    'escalationDirection = False
                'End If
                If (escalatorRng Is Nothing) Or (victimRng Is Nothing) Then 'if escalator is not found in rows or victim not found in columns than swap
                    Set escalatorRng = matrixTbl.Range.Rows(1).Find(escBy)
                    Set victimRng = matrixTbl.Range.Columns(1).Find(escTo)
                    escalationDirection = False
                End If
                
                If escalationDirection Then
                    Set escalationCell = Cells(escalatorRng.row, victimRng.Column) 'reference intersection cell
                Else
                    Set escalationCell = Cells(victimRng.row, escalatorRng.Column)
                End If
            
                ' add an escalation arrow
                Dim arrow As Shape
                Dim arrowLeft As Single, arrowTop As Single, arrowHeightWidth As Single 'arrow coordinates
                Dim arrowColor As Integer
                Dim arrodailyDdsWbent As Integer
                
                'set arrow coordinates within an escalation cell
                arrowTop = escalationCell.Top + 2
                arrowHeightWidth = escalationCell.Height / 1.5
                
                'if escalation from row to column then put in the left part of the cell
                If escalationDirection Then
                    arrowLeft = escalationCell.Left + 5
                Else
                    arrowLeft = escalationCell.Left + escalationCell.Width / 2 + 5
                End If
                                 
                Set arrow = dailyDdsSht.Shapes.AddShape(msoShapeBentArrow, arrowLeft, arrowTop, arrowHeightWidth, arrowHeightWidth) 'add an arrow
                arrow.OnAction = "OnArrowClick" 'put action to do when arrow is clicked
                
                If escalationDirection Then 'if escalation from row to column then flip and rotate in the needed direction
                    arrow.Flip msoFlipHorizontal
                    arrow.IncrementRotation 90
                Else 'if escalation from column to row then flip and rotate in the needed direction and color yellow
                    arrow.IncrementRotation 180
                    arrow.ShapeStyle = msoShapeStylePreset31
                End If
                
                On Error GoTo 0
                
        'If action is not for discussion then add directly to the action plan
        ElseIf escalationDate = Date And forDiscussion = DailyDDS.FOR_DISCUSSION_FALSE Then
                
                'Build an action string
                Dim actionString As String
                actionString = escTo & " - " & helpNeeded & " (Escalated by " & escBy & " | FPC: " & fpc _
                    & " | Plant: " & plant & ")"
                
                'Add action to the daily action plan
                Utils.AddAction dailyDdsSht.Range("B" & DailyDDS.DAILY_ACTION_PLAN_ROW), _
                                Date, _
                                actionString, _
                                escTo, _
                                Utils.NextWorkingDay(Date), _
                                Utils.OPEN_STATUS, _
                                DailyDDS.FOR_DISCUSSION_FALSE
        End If
        
    Next i

End Function

Function OnArrowClick()
'
' Function that shows a UserForm with escalation details when the arrow is clicked
'
'

 'reference the table with escalation matrix
 Dim matrixTbl As ListObject
 Set matrixTbl = ActiveSheet.ListObjects(DailyDDS.MATRIX_TABLE_NAME)

 'determine row and column numbers of clicked arrow
 Dim rowNum As Integer, colNum As Integer
 rowNum = ActiveSheet.Shapes(Application.Caller).TopLeftCell.row
 colNum = ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column
 
 'determine escalation chain
 Dim escalationDirection As Boolean
 If ActiveSheet.Shapes(Application.Caller).Rotation = 90 Then escalationDirection = True
 
 'determine escalator, victim and recommended owner of action
    Dim escBy As String, escTo, owner As String
    If escalationDirection Then
       escBy = Cells(rowNum, matrixTbl.Range.Columns(1).Column)
       escTo = Cells(matrixTbl.Range.Rows(1).row, colNum)
       owner = Cells(matrixTbl.Range.Rows(2).row, colNum)
    Else
       escBy = Cells(matrixTbl.Range.Rows(1).row, colNum)
       escTo = Cells(rowNum, matrixTbl.Range.Columns(1).Column)
       owner = Cells(rowNum, matrixTbl.Range.Columns(2).Column)
    End If
 
'' Initialize the escalation form
    Dim escForm As EscalationForm
    Set escForm = New EscalationForm
 
'' Pass escalation data to the form
    escForm.escalatedBy = escBy
    escForm.escalatedTo = escTo
    escForm.arrName = Application.Caller
    escForm.owner = owner
 
'' Show the escalation form
    escForm.Show
 
End Function


Function DownloadEscalations(dailyDdsWb As Workbook) As Worksheet
'
' Function that downloads escalations for the daily DDS
' Arguements:   dailyDdsWb: workbook to which the escalations needed to be downloaded
' Return:       Reference to a worksheet with escalations

'' Turn off alerts and screen updating
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

'' Reference needed workbooks and worksheets

    Dim escalationsWb As Workbook
    Dim escalationsSht As Worksheet, linksSht As Worksheet, dailyDdsSht As Worksheet
    
    Set dailyDdsSht = dailyDdsWb.Sheets(DailyDDS.DAILY_DDS_SHEET_NAME)
    Set linksSht = dailyDdsWb.Sheets(DailyDDS.LINKS_SHEET_NAME)

'' Delete current escalations sheet if exists
    If Not dailyDdsWb.Sheets(DailyDDS.ESCALATIONS_SHEET_NAME) Is Nothing Then _
    dailyDdsWb.Sheets(DailyDDS.ESCALATIONS_SHEET_NAME).Delete
    
'' Create a new escalations sheet
    Set escalationsSht = dailyDdsWb.Sheets.Add(Before:=dailyDdsWb.Sheets(1))
    escalationsSht.Name = DailyDDS.ESCALATIONS_SHEET_NAME

'' Download and reference escalations file from sharepoint
    On Error Resume Next
        Set escalationsWb = Workbooks.Open(Filename:=ESCALATIONS_FILE_LOCATION, ReadOnly:=True)
    On Error GoTo 0
    
'' If file couldn't be opened then log and skip
    If escalationsWb Is Nothing Then
        Debug.Print ("SharePoint file with escalations couldn't be found")
    Else
        escalationsWb.Sheets("Escalations").Cells.Copy
        escalationsSht.Range("A1").PasteSpecial xlPasteValues
        escalationsSht.Range("A1").PasteSpecial xlPasteFormats
        escalationsWb.Close SaveChanges:=False
    End If

''Download escalations from daily DDSs
    
    Dim lastRowLinksSht As Long
    lastRowLinksSht = linksSht.Range("A1").End(xlDown).row
    
    Dim i As Long
    For i = 2 To lastRowLinksSht
        Dim lastRowEscalationsSht As Long
        lastRowEscalationsSht = escalationsSht.Cells(escalationsSht.Rows.Count, "A").End(xlUp).row
        
        Call DailyDDSEscalationsDownload(linksSht.Cells(i, 1), _
                                          linksSht.Cells(i, 2), _
                                          linksSht.Cells(i, 3), _
                                          linksSht.Cells(i, 4), _
                                          lastRowEscalationsSht + 1, _
                                          escalationsSht, _
                                          linksSht.Cells(i, 6), _
                                          linksSht.Cells(i, 5))
    Next i
    

'' Activate main sheet
    dailyDdsSht.Activate

'' Turn on alerts and screen updating
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
'' Return the reference to the escalations sheet
    Set DownloadEscalations = escalationsSht

End Function

Function DailyDDSEscalationsDownload(escalator As String, path As String, shtName As String, cellAddress As String, rowNum As Integer, escalationsSht As Worksheet, statusRng As Range, dateProvided As String)

'
' Sub that downloads escalations from customer/category dds file
'

Dim ddsWb As Workbook, ddsSht As Worksheet

' If file does not exist then put the according status
If Len(Dir(Environ("USERPROFILE") & "\" & path)) = 0 Then
    statusRng = "File does not exist"
    Exit Function
End If

' If file was not updated today then don't import escalations
If FileLastModified(Environ("USERPROFILE") & "\" & path) < Date Then
    statusRng = "Not updated. Last update: " & FileLastModified(Environ("USERPROFILE") & "\" & path)
    Exit Function
End If

On Error GoTo ErrOpening
Set ddsWb = Workbooks.Open(Environ("USERPROFILE") & "\" & path) 'customer/category DDS file path
Set ddsSht = ddsWb.Sheets(shtName)

Dim lastEscalationRow As Long
lastEscalationRow = ddsSht.Range(cellAddress).End(xlDown).row 'last row of escalations in DDS sheet

Dim escColumn As Integer
escColumn = ddsSht.Range(cellAddress).Column 'first column of escalations

Dim copyRng As Range

'copy paste escalations to main workbook
If ddsSht.Range(cellAddress).Offset(1, 0) <> "" Then

    Dim rangeWidth As Integer 'width of escalations table that will be copied
    
    Set copyRng = Range(ddsSht.Range(cellAddress).Offset(1, 0), ddsSht.Cells(lastEscalationRow, escColumn + ESC_COLUMNS_NUM)) 'escalations range in customer DDS
    copyRng.Copy escalationsSht.Cells(rowNum, ESC_TO_COLUMN) 'copy/paste esclations

    'fill escalator
    Range(escalationsSht.Cells(rowNum, ESC_BY_COLUMN), escalationsSht.Cells(rowNum + copyRng.Rows.Count - 1, ESC_BY_COLUMN)) = escalator
    
    'fill date (if not provided in DDS file then today's date)
    If dateProvided = "yes" Then
        Set copyRng = Range(ddsSht.Range(cellAddress).Offset(1, -1), ddsSht.Cells(lastEscalationRow, escColumn - 1)) 'range with date in customer DDS
        copyRng.Copy Range(escalationsSht.Cells(rowNum, DATE_COLUMN), escalationsSht.Cells(rowNum + copyRng.Rows.Count - 1, DATE_COLUMN)) 'copy/paste
    Else
        Range(escalationsSht.Cells(rowNum, DATE_COLUMN), escalationsSht.Cells(rowNum + copyRng.Rows.Count - 1, DATE_COLUMN)).Value = Date
    End If
    
    'log the status
    statusRng = "Added"
Else
    statusRng = "No input" 'log the status
End If

ddsWb.Close SaveChanges:=False

On Error GoTo 0
Exit Function

ErrOpening:
statusRng = "Error opening"
Resume Next


End Function


Function CleanMatrix(dailyDdsSht As Worksheet)
'
' Function that deletes all arrows (shapes) from the worksheet
' Arguments:    dailyDdsSht - sheet from which to delete the shapes
'
 
    Dim curArrow As Shape
    For Each curArrow In dailyDdsSht.Shapes
        If curArrow.AutoShapeType = msoShapeBentArrow Then curArrow.Delete
    Next curArrow
    
End Function


Sub EmailRecap()
'
' Sub that emails recap of daily DDS to the specified mailing list
'
'
Dim wb As Workbook
Dim dailyDdsSht As Worksheet, restateSht As Worksheet
    
'reference workbook and worksheets
Set wb = ActiveWorkbook
Set dailyDdsSht = wb.Sheets(DAILY_DDS_SHEET_NAME)
Set restateSht = wb.Sheets(DDS_RESTATE_SHEET_NAME)

'define daily action plan range
Dim lastRowDailyActionPlan As Integer
Dim dailyApRng As Range
    
lastRowDailyActionPlan = dailyDdsSht.Range("B" & DAILY_ACTION_PLAN_ROW).End(xlDown).row
Set dailyApRng = dailyDdsSht.Range(Cells(DAILY_ACTION_PLAN_ROW + 1, 2), Cells(lastRowDailyActionPlan, 18))

'define pending restatements range
Dim lastRowRestate As Integer
Dim pendingRestateRng As Range

lastRowRestate = restateSht.Range("B1").End(xlDown).row
Set pendingRestateRng = Range(restateSht.Cells(2, 1), restateSht.Cells(lastRowRestate, APPROVAL_STATUS_COLUMN_NUMBER))

'Open a new mail item (late binding)
Dim outlookApp As Object
Set outlookApp = CreateObject("Outlook.Application")
Dim outMail As Object
Const olMailItem As Long = 0
Set outMail = outlookApp.CreateItem(olMailItem)

'Set parameters of the letter
With outMail
    .Subject = "[" & Date & "] SNO DDS"
    .HTMLBody = "Dear All," & "<br>" & "<br>" & _
                "Please see below summary of today DDS." & "<br>" & _
                "Updated file is at box at normal location: https://pg.box.com/v/EecarSnoDailyDds" & _
                "<br>" & "<br>" & _
                RangetoHTML(dailyApRng, True) & "<br>" & _
                "Pending restatement requests:" & "<br>" & _
                RangetoHTML(pendingRestateRng, False)
    .Display
End With

End Sub

Function ProcessRestatements(wb As Workbook)
'
' Function that archives approved restatements in the restatement workbook and moves pending restatements to the daily DDS
' Arguments:    wb - main Daily DDS workbook
'

''
Application.DisplayAlerts = False
Application.ScreenUpdating = False

''Reference the restatement workbook
    Dim restatementsWb As Workbook
    On Error Resume Next
        Set restatementsWb = Workbooks.Open(Filename:=RESTATEMENTS_FILE_LOCATION, ReadOnly:=False, UpdateLinks:=False)
    On Error GoTo 0
    
''If restatement workbook failed to open then exit the function without doing anything
    If restatementsWb Is Nothing Then
        Debug.Print "Not able to open restatements workbook. Restatements were not processed"
        Exit Function
    End If

''Lock server file to edit the workbook
    'On Error Resume Next
        'restatementsWb.LockServerFile
    'On Error GoTo 0

''Reference needed sheets in the restatement workbook
    Dim forRestateSht, archiveSht As Worksheet
    Set forRestateSht = restatementsWb.Sheets(FOR_RESTATE_SHEET_NAME)
    Set archiveSht = restatementsWb.Sheets(RESTATE_ARCHIVE_SHEET_NAME)
    
''Clear all possible autofilters in the restatements/archive sheets
    forRestateSht.AutoFilterMode = False
    archiveSht.AutoFilterMode = False
    
''Delete previous sheet with restatements if exists
    If Not wb.Sheets(DDS_RESTATE_SHEET_NAME) Is Nothing Then wb.Sheets(DDS_RESTATE_SHEET_NAME).Delete

    
''Calculate last used rows
    Dim lastRowRestateSheet, lastRowArchiveSheet As Long
    lastRowRestateSheet = forRestateSht.Range("A1").End(xlDown).row 'last row in the sheet pending for restatement
    lastRowArchiveSheet = archiveSht.Range("A1").End(xlDown).row 'last row in the archive sheet
    
''Iterate through rows with missing Approval status put pending status, all approved/denied move to archive
    
    Dim i As Long
    For i = lastRowRestateSheet To 3 Step -1
        Select Case forRestateSht.Cells(i, APPROVAL_STATUS_COLUMN_NUMBER)
            Case APPROVED_STATUS, DENIED_STATUS
                forRestateSht.Rows(i).Copy archiveSht.Rows(lastRowArchiveSheet + 1) 'copy row to the archive sheet
                forRestateSht.Rows(i).Delete 'delete the row
                lastRowArchiveSheet = archiveSht.Range("A1").End(xlDown).row 'recalculate last row in the archive sheet
            Case PENDING_STATUS
            Case Else
                forRestateSht.Cells(i, APPROVAL_STATUS_COLUMN_NUMBER) = PENDING_STATUS 'if status is missing - put PENDING
        End Select
    Next i
    
''Copy the sheet with pending restatements to the SNO DDS File
    Dim restSht As Worksheet
    Set restSht = wb.Sheets.Add
    restSht.Name = DDS_RESTATE_SHEET_NAME
    forRestateSht.Cells.Copy
    restSht.Range("A1").PasteSpecial xlPasteValues
    restSht.Range("A1").PasteSpecial xlPasteFormats
    
''Save/close restatements workbook
    restatementsWb.Close SaveChanges:=True
  
''
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
  

End Function




