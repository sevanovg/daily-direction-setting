Attribute VB_Name = "Utils"
Option Explicit
Option Compare Text

'' Status constants
Public Const OPEN_STATUS = "OPEN"
Public Const DONE_STATUS = "DONE"
Public Const OVERDUE_STATUS = "OVERDUE"

Public Function getCfr(ByRef dailyCfrSheet As Worksheet, Optional product As String, Optional geography As String, _
        Optional shipFrom As String, Optional customer As String, Optional key As String, Optional dDate As String) As Variant
'
' Function that returns CFR result for specified category/geography from Daily CFR report workbook
' Arguments:    dailyCfrSheet: reference to the sheet with CFR data in Daily CFR report
'               product: product category as Product 5005
'               geography:  geography string as in Daily CFR report
'               customer:   customer string as in Daily CFR report
'               shipFrom:   Ship From string as in Daily CFR report
'               key:    key String if match by key is needed
'               dDate:  date for which CFR is needed or MTD label if MTD data is needed
' Function exits if needed data was not found

''Constants for default strings if not provided
    Const DEFAULT_PRODUCT As String = "EECAR Total"
    Const DEFAULT_GEOGRAPHY As String = "EECAR"
    Const DEFAULT_SHIP_FROM As String = "ALL LOCATIONS"
    Const DEFAULT_CUSTOMER As String = "TOTAL CUSTOMERS"
    
''Constants that determine column numbers in CFR report
    Const KEY_COL_NUMBER As Integer = 1
    Const PRODUCT_COL_NUMBER As Integer = 3
    Const GEOGRAPHY_COL_NUMBER As Integer = 4
    Const SHIP_FROM_COL_NUMBER As Integer = 5
    Const CUSTOMER_COL_NUMBER As Integer = 6
    Const CFR_MTD_COLUMN As Integer = 9
    
''Constant for row numbers in Daily CFR file
    Const DATE_ROW_NUMBER As Integer = 1
    Const LABEL_ROW_NUMBER As Integer = 2

''Other constants
    Const CFR_LABEL As String = "% CFR"
    
    
''Assign default values if not provided
    If product = "" Then product = DEFAULT_PRODUCT
    If geography = "" Then geography = DEFAULT_GEOGRAPHY
    If shipFrom = "" Then shipFrom = DEFAULT_SHIP_FROM
    If customer = "" Then customer = DEFAULT_CUSTOMER
    
''Convert product to infopage standard
    If product = "Total" Then product = "EECAR Total"
    If product = "Beauty Care" Then product = "PersonalCare"
    If product = "Fem Care" Then product = "Feminine Care"

''If date was formatted in the text form then convert it to date
    Dim dDateConverted As Date
    If Not dDate = MONTH_TO_DATE_LABEL Then dDateConverted = DateValue(dDate)
    
''Find column number based on a needed date. If MTD needed then use MTD column
    Dim colNum As Integer, cCell As Range, lastCol As Integer, i As Integer
    If dDate = MONTH_TO_DATE_LABEL Then
        colNum = CFR_MTD_COLUMN
    Else
        ''Find last column
        lastCol = dailyCfrSheet.Cells(DATE_ROW_NUMBER, dailyCfrSheet.Columns.Count).End(xlToLeft).Column
        
        ''Iterate through each column and find CFR column with a needed date
        For i = 1 To lastCol
            If dailyCfrSheet.Cells(DATE_ROW_NUMBER, i).Value2 = dDateConverted And dailyCfrSheet.Cells(LABEL_ROW_NUMBER, i) = CFR_LABEL Then
                colNum = i
                Exit For
            End If
        Next i
    End If
        
''Find row number based on criteria
    Dim rowNum As Integer, lastRow As Integer
    lastRow = dailyCfrSheet.Cells(dailyCfrSheet.Rows.Count, PRODUCT_COL_NUMBER).End(xlUp).row
    
    ''Iterate through each row and match by criteria
    For i = 2 To lastRow
    
        ''If key was provided then match by key
        If key <> "" Then
            If Trim(dailyCfrSheet.Cells(i, KEY_COL_NUMBER)) = Trim(key) Then
                rowNum = i
                Exit For
            End If
                
        ''If key was not provided then match by multiple criteria
        Else
            If Trim(dailyCfrSheet.Cells(i, PRODUCT_COL_NUMBER)) = Trim(product) And _
                Trim(dailyCfrSheet.Cells(i, GEOGRAPHY_COL_NUMBER)) = Trim(geography) And _
                Trim(dailyCfrSheet.Cells(i, SHIP_FROM_COL_NUMBER)) = Trim(shipFrom) And _
                Trim(dailyCfrSheet.Cells(i, CUSTOMER_COL_NUMBER)) = Trim(customer) _
            Then
                rowNum = i
                Exit For
            End If
        End If
    Next i

''Return cfr for found row/column. If not found then return blank
    If rowNum = 0 Or colNum = 0 Then
        getCfr = ""
    Else
        getCfr = dailyCfrSheet.Cells(rowNum, colNum).Value2
    End If


End Function

Public Function getProxyCfr(ByRef cfrProxyWb As Workbook, dDate As String, unitOfMeasure As String, _
     Optional geography As String, Optional product As String, Optional customer As String, Optional key As String) As Variant
'
' Function that returns proxy CFR result for specified customer / date. Additional criteria can be applied if needed.
' Arguments:    dailyCfrProxyWb: reference to the Daily CFR proxy workbook
'               product: product category as Product 5005
'               geography:  geography string as in Daily CFR report
'               customer:   customer string as in Daily CFR report
'               shipFrom:   Ship From string as in Daily CFR report
'               key:    key String if match by key is needed
'               dDate:  date for which CFR is needed or MTD label if MTD data is needed
' Function exits if needed data was not found

''Constants for page names in the proxy workbook
    Const DAILY_PROXY_SU_SHEET_NAME As String = "SU"
    Const DAILY_PROXY_IT_SHEET_NAME As String = "IT"

''Public constants for units of measure
    Const UNIT_OF_MEASURE_SU = "SU"
    Const UNIT_OF_MEASURE_IT = "IT"

''Constants for default strings if not provided
    Const DEFAULT_PRODUCT As String = "EECAR Total products"
    Const DEFAULT_CUSTOMER As String = "TOTAL - ALL CUSTOMERS [9900000001]"
    
''Constants that determine column numbers in CFR report
    Const PRODUCT_COL_NUMBER As Integer = 1
    Const CUSTOMER_COL_NUMBER As Integer = 2
    Const GEOGRAPHY_COL_NUMBER As Integer = 3
    Const RCA_COL_NUMBER As Integer = 4 ' "Root Cause - Case Fill Rate" column
    Const CFR_MTD_COLUMN As Integer = 8
    
''Constant for row numbers in Daily CFR file
    Const DATE_ROW_NUMBER As Integer = 7
    Const LABEL_ROW_NUMBER As Integer = 8

''Other constants
    Const CFR_LABEL As String = "% Case Fill Rate"
    Const SPLIT_KEY As String = "|" 'symbol used to separate values in a key string for the SAMBC table

''If CSV key was provided then get values from this key (Format: Product,Customer,Geography)
       
    If key <> "" Then
        ''Get an array of Strings separated by a split key
        Dim keyArray() As String
        keyArray = Split(key, SPLIT_KEY)
        
        ''Assign needed values from an array
        product = keyArray(0)
        customer = keyArray(1)
        geography = keyArray(2)
    Else
        ''Assign default values if not provided
        If product = "" Then product = DEFAULT_PRODUCT
        If customer = "" Then customer = DEFAULT_CUSTOMER
    End If

''Reference the needed sheet depending on unit of measure
    Dim cfrProxySht As Worksheet
    
    If unitOfMeasure = UNIT_OF_MEASURE_SU Then
        Set cfrProxySht = cfrProxyWb.Sheets(DAILY_PROXY_SU_SHEET_NAME)
    ElseIf unitOfMeasure = UNIT_OF_MEASURE_IT Then
        Set cfrProxySht = cfrProxyWb.Sheets(DAILY_PROXY_IT_SHEET_NAME)
    Else
        Exit Function
    End If
        

''If date was formatted in the text form then convert it to date
    Dim dDateConverted As Date
    If Not dDate = MONTH_TO_DATE_LABEL Then dDateConverted = DateValue(dDate)
    
''Find column number based on a needed date. If MTD needed then use MTD column
    Dim colNum As Integer, cCell As Range, lastCol As Integer, i As Integer
    If dDate = MONTH_TO_DATE_LABEL Then
        colNum = CFR_MTD_COLUMN
    Else
        ''Find last column
        lastCol = cfrProxySht.Cells(LABEL_ROW_NUMBER, cfrProxySht.Columns.Count).End(xlToLeft).Column
        
        ''Iterate through each column and find CFR column with a needed date
        For i = 1 To lastCol
            ''If date matches and label contains CFR label
            If convertCfrProxyDate(cfrProxySht.Cells(DATE_ROW_NUMBER, i).MergeArea.Cells(1, 1)) = dDateConverted And _
                    InStr(1, cfrProxySht.Cells(LABEL_ROW_NUMBER, i), CFR_LABEL) > 0 Then
                colNum = i
                Exit For
            End If
        Next i
    End If
        
''Find row number based on criteria
    Dim fullRowNum As Integer, blankRowNum As Integer, lastRow As Integer
    lastRow = cfrProxySht.Cells(cfrProxySht.Rows.Count, PRODUCT_COL_NUMBER).End(xlUp).row
    
    ''Iterate through each row and match by criteria
    For i = 2 To lastRow
                
        ''Match by multiple criteria (product/geography/customer)
            If Trim(cfrProxySht.Cells(i, PRODUCT_COL_NUMBER).MergeArea.Cells(1, 1)) = Trim(product) And _
                Trim(cfrProxySht.Cells(i, GEOGRAPHY_COL_NUMBER).MergeArea.Cells(1, 1)) = Trim(geography) And _
                Trim(cfrProxySht.Cells(i, CUSTOMER_COL_NUMBER).MergeArea.Cells(1, 1)) = Trim(customer) _
            Then
                If cfrProxySht.Cells(i, RCA_COL_NUMBER) = "" Then
                    '' If row is blank ("Root Cause - Case Fill Rate" is blank) then mark as blank row.
                    blankRowNum = i
                Else
                    ''If full then mark as full
                    fullRowNum = i
                End If
                
                ''If blank and full rows were both found then exit search loop
                If fullRowNum > 0 And blankRowNum > 0 Then Exit For
                
            Else
                ''If blank/full row was found but already moved to another customer then exit search loop
                If fullRowNum > 0 Or blankRowNum > 0 Then Exit For
            End If
    Next i

''Return cfr for found row/column. If not found then return blank
    If (fullRowNum = 0 And blankRowNum = 0) Or colNum = 0 Then
        getProxyCfr = ""
    Else
        ''Return the max CFR value between "blank" and "full" rows
        If fullRowNum = 0 Then getProxyCfr = cfrProxySht.Cells(blankRowNum, colNum)
        If blankRowNum = 0 Then getProxyCfr = cfrProxySht.Cells(fullRowNum, colNum)
        If fullRowNum > 0 And blankRowNum > 0 Then
            getProxyCfr = Application.Max(cfrProxySht.Cells(blankRowNum, colNum), cfrProxySht.Cells(fullRowNum, colNum))
        End If
    End If

End Function

Private Function convertCfrProxyDate(dateString As String) As Date
'
' Function that converts day as in CFR PROXY workbook to a date format (e.g.DAY305, WEDNESDAY, NOV 1 2017 to 01.11.2017)
' Arguments:    dateString: date string as in CFR PROXY workbook (DAY305, WEDNESDAY, NOV 1 2017)
' Returns -1 when date can't be converted

Dim day As Integer, month As Integer, year As Integer, dayNumber As Integer

On Error GoTo date_err

'' Parse year
    year = CInt(Right(dateString, 4))

'' Parse day number (1-365) (starting from 4th symbol till first coma)
    dayNumber = CInt(Mid(dateString, 4, InStr(1, dateString, ",") - 4))

'' Convert day number to date of current year (31.12 of previous year + day number)
    convertCfrProxyDate = DateSerial(year - 1, 12, 31) + dayNumber

On Error GoTo 0

Exit Function

''Error handler for when incoming string can't be converted
date_err:
    convertCfrProxyDate = -1
    Exit Function

End Function


Public Function archiveActions(ByRef sourceRng As Range, ByRef destinationRng As Range, statusColumnNum As Integer)
'
' Function that moves actions with status "DONE" to the daily DDS archive
' Arguments:    sourceRng: range with actions that should be archived
'               statusColumnNum: column number in the actions range which contains the status
'               destinationRng: top left cell to which the archived actions should be copied
'


'' Iterate through each row of the source actions range
    Dim rowNum As Integer
    For rowNum = sourceRng.Rows.Count To 1 Step -1
        '' If action has "DONE" status then move the row with action to the archive
        If sourceRng(rowNum, statusColumnNum) = DONE_STATUS Then
            
            '' Copy the row with action to the archive
            sourceRng.Rows(rowNum).Copy destinationRng
                
            '' Move the destination range to one cell below
            Set destinationRng = destinationRng.Offset(1, 0)
            
            '' Delete the copied row from source range
            sourceRng.Rows(rowNum).EntireRow.Delete
        End If
    Next rowNum

End Function


Public Function IndexOf(ByVal coll As Collection, ByVal Item As Variant) As Long
'
' Function that determines the index of an item in the collection. Returns 0 if item does not exist
' Arguments:    coll: Collection to search for an element
'               Item: Item of the collection for which the index should be determined
'
    Dim i As Long
    For i = 1 To coll.Count
        If coll(i) = Item Then
            IndexOf = i
            Exit Function
        End If
    Next
End Function

Public Function FileLastModified(strFullFileName As String)
'
' Function that returns last modified date of a file
' Arguments:    strFullFileName: string that represents the full path to the file
'
    Dim fs As Object, f As Object, s As String
     
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(strFullFileName)
    
    FileLastModified = f.DateLastModified
     
    Set fs = Nothing: Set f = Nothing
     
End Function

Public Function RangetoHTML(rng As Range, Optional actionPlanFormat As Boolean)
'
'Function that converts input range to HTML format.
'Set actionPlanFormat to True if the input table is daily action plan that needs to be specially formatted
'
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    
    If actionPlanFormat Then
        'format the action plan table to look nicely in outlook message
        TempWB.Sheets(1).Range("A:A").ColumnWidth = 10 'date
        TempWB.Sheets(1).Range("B:B").ColumnWidth = 100 'action
        TempWB.Sheets(1).Range("C:I").ColumnWidth = 0 'action hidden columns
        TempWB.Sheets(1).Range("J:L").ColumnWidth = 10 'owner/status/due date
        TempWB.Sheets(1).Range("M:M").ColumnWidth = 40 'comment
        TempWB.Sheets(1).Range("N:Q").ColumnWidth = 0 'comment hidden
        
        TempWB.Sheets(1).Range("B:Q").WrapText = True
        TempWB.Sheets(1).Cells.VerticalAlignment = xlCenter
    End If


    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Function AddAction(actionPlanRng As Range, addedDate As Date, action As String, owner As String, dueDate As Date, _
    status As String, Optional forDiscussion As Integer)
'
' Function that adds a new action to the action plan
' Arguments:    actionPlanRng: top left cell of an action plan to which the action should be added
'               addedDate: date when action was added
'               action: text description of an action
'               owner: owner of an action
'               dueDate: action deadline date
'               status: action status
'               forDiscussion: whether issue is for discussion or not (1 - without discussion)
'

    Dim lastRowDailyActionPlan As Integer
    
    'Calculate the last filled row in the action plan
    lastRowDailyActionPlan = actionPlanRng.End(xlDown).row
    
    'Insert an empty row below for the newly added action
    Dim newActionRow As Range
    actionPlanRng.Worksheet.Rows(lastRowDailyActionPlan + 1).Insert Shift:=xlDown
    Set newActionRow = actionPlanRng.Worksheet.Rows(lastRowDailyActionPlan + 1)
    
    'copy formatting from the row above
    With newActionRow
        .Offset(-1, 0).Copy 'Copy above row
        .PasteSpecial Paste:=xlPasteFormats 'Paste formats
        .Cells(, DailyDDS.ACTION_PLAN_ACTION_COLUMN).Resize(, 8).Merge 'Merge 8 cells for an action description
    End With
    'dailySht.Rows(lastRowDailyActionPlan).Copy
    'dailySht.Rows(lastRowDailyActionPlan + 1).PasteSpecial Paste:=xlPasteFormats
    'Range(dailySht.Cells(lastRowDailyActionPlan + 1, 3), dailySht.Cells(lastRowDailyActionPlan + 1, 10)).Merge 'merge the cells with action
    
    'put all the action details
    With newActionRow
        .Columns(DailyDDS.ACTION_PLAN_DATE_COLUMN) = addedDate
        .Columns(DailyDDS.ACTION_PLAN_ACTION_COLUMN) = action
        .Columns(DailyDDS.ACTION_PLAN_OWNER_COLUMN) = owner
        .Columns(DailyDDS.ACTION_PLAN_DEADLINE_COLUMN) = dueDate
        .Columns(ACTION_PLAN_STATUS_COLUMN) = status
        ' append *** if action was not for discussion
        If forDiscussion = DailyDDS.FOR_DISCUSSION_FALSE Then
            .Columns(DailyDDS.ACTION_PLAN_ACTION_COLUMN) = "*** " & .Columns(DailyDDS.ACTION_PLAN_ACTION_COLUMN)
        End If
    End With
End Function

Function NextWorkingDay(day As Date) As Date
'
' Function that returns the next working date
' Arguments:    day: the date for which the next working day needs to be calculated
'

NextWorkingDay = day + 1
If Weekday(day + 1, vbMonday) = 6 Then NextWorkingDay = day + 3
If Weekday(day + 1, vbMonday) = 7 Then NextWorkingDay = day + 2


End Function
