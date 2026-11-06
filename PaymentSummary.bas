Attribute VB_Name = "PaymentSummary"
Option Explicit

'============================================================
' USER-CONFIGURABLE SETTINGS
'============================================================

Private Const OUTPUT_SHEET_NAME As String = "Payment Summary"
Private Const PROCESSING_LOG_SHEET_NAME As String = "Processing Log"

Private Const FIELD_DOCUMENT_CURRENCY As String = "Currency"
Private Const FIELD_TRANSACTION_TYPE As String = "Types"
Private Const FIELD_DOCUMENT_AMOUNT As String = "Amount"

Private Const OUTPUT_START_ROW As Long = 2
Private Const OUTPUT_START_COLUMN As Long = 2
Private Const PAYMENT_ID_COLUMN As Long = 4

'============================================================
' PROCESSING LOG STATE
'============================================================

Private processingLogWs As Worksheet
Private processingLogNextRow As Long

'============================================================
' MAIN PROCEDURE
'============================================================

Public Sub GeneratePaymentSummary()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outputWs As Worksheet

    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable

    Dim lastRow As Long
    Dim lastCol As Long

    Dim cleanName As String
    Dim bank As String
    Dim paymentId As String

    Dim nextRow As Long
    Dim pivotBottom As Long
    Dim pivotStart As Long

    Dim currentSheetName As String
    Dim skipReason As String

    Dim processedCount As Long
    Dim skippedCount As Long
    Dim warningCount As Long

    Dim previousCalculation As XlCalculation
    Dim previousScreenUpdating As Boolean
    Dim previousEnableEvents As Boolean
    Dim previousDisplayAlerts As Boolean

    '--------------------------------------------------------
    ' Save the user's current Excel application settings.
    ' These settings are restored before the procedure exits.
    '--------------------------------------------------------

    previousScreenUpdating = Application.ScreenUpdating
    previousEnableEvents = Application.EnableEvents
    previousCalculation = Application.Calculation
    previousDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    '--------------------------------------------------------
    ' The target workbook is intentionally ActiveWorkbook.
    ' This allows the macro to be stored outside the workbook,
    ' including in PERSONAL.XLSB.
    '--------------------------------------------------------

    Set wb = Application.ActiveWorkbook

    If wb Is Nothing Then
        Err.Raise _
            Number:=vbObjectError + 1, _
            Source:="GeneratePaymentSummary", _
            Description:="No active workbook was found."
    End If

    If UCase$(wb.Name) = "PERSONAL.XLSB" Then
        Err.Raise _
            Number:=vbObjectError + 2, _
            Source:="GeneratePaymentSummary", _
            Description:="PERSONAL.XLSB is active. Activate the target workbook first."
    End If

    '--------------------------------------------------------
    ' Remove and recreate the output sheet on every run.
    ' This ensures that old PivotTables and previous output
    ' do not remain in the workbook.
    '--------------------------------------------------------

    Set outputWs = RecreateOutputSheet(wb, OUTPUT_SHEET_NAME)
    outputWs.Visible = xlSheetVisible
    
    InitializeProcessingLog wb
    
    nextRow = OUTPUT_START_ROW

    LogMessage _
        "STARTED", _
        vbNullString, _
        "GeneratePaymentSummary started for workbook: " & wb.Name

    '--------------------------------------------------------
    ' Process worksheets.
    '
    ' Business rules:
    '   1. Only visible worksheets are processed.
    '   2. Configured administrative sheets are excluded.
    '   3. Worksheet names must use one of these formats:
    '
    '          Bank-ID
    '          Bank - ID
    '
    '      Spaces around the hyphen are optional.
    '      Parentheses around the Payment ID are not allowed.
    '      Exactly one hyphen must be present.
    '
    '      Correct:   Delta Haven Bank-PAY4464
    '      Correct:   Delta Haven Bank - PAY4464
    '      Incorrect: Delta Haven Bank-(PAY4464)
    '      Incorrect: Delta Haven Bank - (PAY4464)
    '      Incorrect: Delta Haven Bank-PAY-4464
    '
    '   4. The source data starts in row 1.
    '   5. Column A determines the last source row.
    '--------------------------------------------------------

    For Each ws In wb.Worksheets

        currentSheetName = ws.Name
        skipReason = vbNullString

        If ws.Name = outputWs.Name Then

            'The macro-created output sheet is not a source worksheet.
            'It is excluded from processing and from the skipped count.
            LogMessage _
                "EXCLUDED", _
                currentSheetName, _
                "Macro output sheet"
        
        ElseIf ws.Name = PROCESSING_LOG_SHEET_NAME Then
        
            'The processing log is an administrative worksheet.
            'It is excluded from processing and from the skipped count.
            LogMessage _
                "EXCLUDED", _
                currentSheetName, _
                "Macro processing log"
        
        ElseIf ws.Visible <> xlSheetVisible Then
        
            skipReason = "Worksheet is hidden"
        
        ElseIf IsExcludedSheet(ws.Name) Then
        
            skipReason = "Worksheet is in the exclusion list"
        
        Else

            '------------------------------------------------
            ' Pass the complete worksheet name to the parser.
            '
            ' Do not remove text beginning with a parenthesis
            ' here. The parser explicitly rejects parentheses
            ' and records a clear reason for the skipped sheet.
            '------------------------------------------------

            cleanName = Trim$(ws.Name)

            If Not TryParseBankAndId( _
                cleanName, _
                bank, _
                paymentId) Then

                skipReason = _
                    "Invalid worksheet name. Use Bank-ID or " & _
                    "Bank - ID without parentheses and with " & _
                    "exactly one hyphen. Example: " & _
                    "Delta Haven Bank - PAY4464"

            Else

                lastRow = ws.Cells( _
                    ws.Rows.Count, _
                    1).End(xlUp).Row

                lastCol = ws.Cells( _
                    1, _
                    ws.Columns.Count).End(xlToLeft).Column

                If lastRow < 2 Or lastCol < 1 Then

                    skipReason = _
                        "Worksheet does not contain usable source data"

                Else

                    Set pivotCache = wb.PivotCaches.Create( _
                        SourceType:=xlDatabase, _
                        SourceData:=ws.Range( _
                            ws.Cells(1, 1), _
                            ws.Cells(lastRow, lastCol)))

                    '----------------------------------------
                    ' Add the bank title above the PivotTable.
                    '----------------------------------------

                    With outputWs.Cells( _
                        nextRow, _
                        OUTPUT_START_COLUMN)

                        .Value = BankDisplayName(bank)
                        .Font.Bold = True
                        .HorizontalAlignment = xlLeft
                        .VerticalAlignment = xlCenter
                        .Interior.Color = RGB(255, 255, 0)

                    End With

                    nextRow = nextRow + 1
                    pivotStart = nextRow

                    '----------------------------------------
                    ' Create the PivotTable.
                    '----------------------------------------

                    Set pivotTable = _
                        pivotCache.CreatePivotTable( _
                            TableDestination:=outputWs.Cells( _
                                nextRow, _
                                OUTPUT_START_COLUMN), _
                            TableName:=GetUniquePivotName( _
                                wb, _
                                "PT_" & bank & "_" & paymentId))

                    With pivotTable

                        .ManualUpdate = True

                        If PivotFieldExists( _
                            pivotTable, _
                            FIELD_DOCUMENT_CURRENCY) Then

                            .PivotFields( _
                                FIELD_DOCUMENT_CURRENCY _
                            ).Orientation = xlRowField

                        Else

                            LogMissingField _
                                currentSheetName, _
                                FIELD_DOCUMENT_CURRENCY

                            warningCount = warningCount + 1

                        End If

                        If PivotFieldExists( _
                            pivotTable, _
                            FIELD_TRANSACTION_TYPE) Then

                            .PivotFields( _
                                FIELD_TRANSACTION_TYPE _
                            ).Orientation = xlRowField

                        Else

                            LogMissingField _
                                currentSheetName, _
                                FIELD_TRANSACTION_TYPE

                            warningCount = warningCount + 1

                        End If

                        If PivotFieldExists( _
                            pivotTable, _
                            FIELD_DOCUMENT_AMOUNT) Then
                        
                            .AddDataField _
                                (.PivotFields(FIELD_DOCUMENT_AMOUNT)), _
                                "Sum of " & FIELD_DOCUMENT_AMOUNT, _
                                xlSum
                        
                        Else
                        
                            LogMissingField _
                                currentSheetName, _
                                FIELD_DOCUMENT_AMOUNT
                        
                            warningCount = warningCount + 1
                        
                        End If

                        .RowAxisLayout xlCompactRow
                        .ManualUpdate = False

                    End With

                    outputWs.Columns( _
                        OUTPUT_START_COLUMN).AutoFit

                    pivotBottom = _
                        pivotTable.TableRange2.Row + _
                        pivotTable.TableRange2.Rows.Count - 1

                    '----------------------------------------
                    ' Add the Payment ID heading.
                    '----------------------------------------

                    Dim pivotHeaderCell As Range

                    Set pivotHeaderCell = pivotTable.TableRange2.Cells(1, 1)
                    
                    With outputWs.Cells( _
                        pivotStart, _
                        PAYMENT_ID_COLUMN)
                    
                        .Value = "Payment ID"
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .Font.Bold = True
                        .Font.Color = pivotHeaderCell.DisplayFormat.Font.Color
                    
                        With .Interior
                            .Pattern = pivotHeaderCell.DisplayFormat.Interior.Pattern
                            .Color = pivotHeaderCell.DisplayFormat.Interior.Color
                            .PatternColor = _
                                pivotHeaderCell.DisplayFormat.Interior.PatternColor
                        End With
                    
                    End With

                    '----------------------------------------
                    ' Merge the Payment ID cells beside the
                    ' PivotTable.
                    '----------------------------------------

                    If pivotBottom >= pivotStart + 1 Then

                        With outputWs.Range( _
                            outputWs.Cells( _
                                pivotStart + 1, _
                                PAYMENT_ID_COLUMN), _
                            outputWs.Cells( _
                                pivotBottom, _
                                PAYMENT_ID_COLUMN))

                            .Merge
                            .Value = paymentId
                            .HorizontalAlignment = xlCenter
                            .VerticalAlignment = xlCenter
                            .Font.Bold = True
                            .Interior.ColorIndex = xlNone
                            .Borders.LineStyle = xlNone

                        End With

                        outputWs.Columns( _
                            PAYMENT_ID_COLUMN).AutoFit

                    End If

                    nextRow = pivotBottom + 2
                    processedCount = processedCount + 1

                    LogMessage _
                        "PROCESSED", _
                        currentSheetName, _
                        "Bank: " & bank & "; Payment ID: " & paymentId

                End If

            End If

        End If

        If Len(skipReason) > 0 Then

            skippedCount = skippedCount + 1

            LogMessage _
                "SKIPPED", _
                currentSheetName, _
                skipReason
                
        End If

    Next ws

    '--------------------------------------------------------
    ' Move the newly created Payment Summary sheet
    ' to the beginning of the workbook.
    '--------------------------------------------------------

    outputWs.Move Before:=wb.Worksheets(1)

    LogMessage _
        "COMPLETED", _
        vbNullString, _
        processedCount & " worksheet(s) processed; " & _
        skippedCount & " worksheet(s) skipped; " & _
        warningCount & " field warning(s)."
    
    FormatProcessingLog

    ShowCompletionMessage _
        processedCount, _
        skippedCount, _
        warningCount, _
        processingLogWs
        
CleanExit:

    Application.ScreenUpdating = previousScreenUpdating
    Application.EnableEvents = previousEnableEvents
    Application.Calculation = previousCalculation
    Application.DisplayAlerts = previousDisplayAlerts

    Set processingLogWs = Nothing

    Exit Sub

CleanFail:


    Call LogMessage( _
        "ERROR", _
        currentSheetName, _
        "Error " & Err.Number & ": " & Err.Description)

    MsgBox _
        "The report could not be completed." & _
        vbCrLf & vbCrLf & _
        "Worksheet: " & _
        GetDisplaySheetName(currentSheetName) & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Generate Payment Summary Report"

    Resume CleanExit

End Sub

'============================================================
' WORKSHEET HELPERS
'============================================================

Private Function RecreateOutputSheet( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    Dim existingSheet As Worksheet
    Dim displayAlertsState As Boolean

    displayAlertsState = Application.DisplayAlerts

    On Error GoTo CleanFail

    Set existingSheet = GetWorksheetIfExists( _
        wb, _
        sheetName)

    If Not existingSheet Is Nothing Then

        Application.DisplayAlerts = False
        existingSheet.Delete
        Application.DisplayAlerts = displayAlertsState

    End If

    Set RecreateOutputSheet = wb.Worksheets.Add( _
        After:=wb.Worksheets(wb.Worksheets.Count))

    RecreateOutputSheet.Name = sheetName

CleanExit:

    Application.DisplayAlerts = displayAlertsState
    Exit Function

CleanFail:

    Application.DisplayAlerts = displayAlertsState

    Err.Raise _
        Number:=Err.Number, _
        Source:="RecreateOutputSheet", _
        Description:=Err.Description

End Function

Private Function GetWorksheetIfExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetWorksheetIfExists = wb.Worksheets(sheetName)
    On Error GoTo 0

End Function

Private Function IsExcludedSheet( _
    ByVal sheetName As String) As Boolean

    '--------------------------------------------------------
    ' CONFIGURATION POINT:
    '
    ' Add or remove worksheet names from this list.
    '
    ' OUTPUT_SHEET_NAME does not need to be listed because
    ' it is handled separately by the main procedure.
    '--------------------------------------------------------

    Select Case UCase$(Trim$(sheetName))

        Case "RAW"
            IsExcludedSheet = True

    End Select

End Function

'============================================================
' WORKSHEET NAME PARSING
'============================================================

Private Function TryParseBankAndId( _
    ByVal cleanName As String, _
    ByRef bank As String, _
    ByRef paymentId As String) As Boolean

    Dim separatorPosition As Long
    Dim hyphenCount As Long

    bank = vbNullString
    paymentId = vbNullString

    cleanName = Trim$(cleanName)

    '--------------------------------------------------------
    ' Required worksheet naming convention:
    '
    '     Bank-ID
    '     Bank - ID
    '
    ' Spaces around the hyphen are optional.
    '
    ' Correct examples:
    '     Delta Haven Bank-PAY4464
    '     Delta Haven Bank - PAY4464
    '
    ' Invalid examples:
    '     Delta Haven Bank-(PAY4464)
    '     Delta Haven Bank - (PAY4464)
    '     Delta Haven Bank-PAY-4464
    '     Delta Haven Bank
    '
    ' Validation rules:
    '   1. Exactly one hyphen must be present.
    '   2. The bank name must not be blank.
    '   3. The Payment ID must not be blank.
    '   4. Parentheses are not permitted.
    '--------------------------------------------------------

    If Len(cleanName) = 0 Then
        Exit Function
    End If

    '--------------------------------------------------------
    ' Reject parentheses anywhere in the worksheet name.
    '--------------------------------------------------------

    If InStr( _
        1, _
        cleanName, _
        "(", _
        vbBinaryCompare) > 0 Or _
       InStr( _
        1, _
        cleanName, _
        ")", _
        vbBinaryCompare) > 0 Then

        Exit Function

    End If

    '--------------------------------------------------------
    ' Require exactly one hyphen.
    '
    ' This prevents Bank-PAY-4464 from being interpreted as
    ' bank "Bank" and Payment ID "PAY".
    '--------------------------------------------------------

    hyphenCount = _
        Len(cleanName) - _
        Len(Replace$(cleanName, "-", vbNullString))

    If hyphenCount <> 1 Then
        Exit Function
    End If

    separatorPosition = InStr( _
        1, _
        cleanName, _
        "-", _
        vbBinaryCompare)

    If separatorPosition = 0 Then
        Exit Function
    End If

    '--------------------------------------------------------
    ' Extract and trim the bank name and Payment ID.
    '--------------------------------------------------------

    bank = Trim$(Left$( _
        cleanName, _
        separatorPosition - 1))

    paymentId = Trim$(Mid$( _
        cleanName, _
        separatorPosition + 1))

    '--------------------------------------------------------
    ' Both values are required.
    '--------------------------------------------------------

    If Len(bank) = 0 Or Len(paymentId) = 0 Then

        bank = vbNullString
        paymentId = vbNullString

        Exit Function

    End If

    TryParseBankAndId = True

End Function

'============================================================
' BANK DISPLAY NAME CONFIGURATION
'============================================================

Private Function BankDisplayName( _
    ByVal bankCode As String) As String

    Dim mappings As Variant
    Dim mapping As Variant
    Dim normalizedBankCode As String

    normalizedBankCode = UCase$(Trim$(bankCode))

    '--------------------------------------------------------
    ' OPTIONAL BANK DISPLAY NAME MAPPINGS
    '
    ' This feature can be used when a worksheet name contains
    ' a shortened bank code, but the Payment Summary should
    ' show a longer or friendlier bank name.
    '
    ' Format:
    '     Array("SOURCE CODE OR NAME", "DISPLAY NAME")
    '
    ' Fictional examples:
    '     AVB - ASTER VAULT BANK
    '     BBK - BOREALIS BANK
    '
    ' If no mapping matches, the bank name from the worksheet
    ' is displayed unchanged. Therefore, these mappings do not
    ' need to be used for the macro to work.
    '--------------------------------------------------------

    mappings = Array( _
        Array("AVB", "Aster Vault Bank"), _
        Array("BBK", "Borealis Bank"), _
        Array("CFB", "Crescent Forge Bank"), _
        Array("DHB", "Delta Haven Bank") _
    )

    For Each mapping In mappings

        If normalizedBankCode = _
            UCase$(CStr(mapping(0))) Then

            BankDisplayName = CStr(mapping(1))
            Exit Function

        End If

    Next mapping

    '--------------------------------------------------------
    ' No mapping was found.
    ' Use the original bank name exactly as parsed from the
    ' worksheet name.
    '--------------------------------------------------------

    BankDisplayName = Trim$(bankCode)

End Function

'============================================================
' PIVOTTABLE HELPERS
'============================================================

Private Function PivotFieldExists( _
    ByVal pivotTable As pivotTable, _
    ByVal fieldName As String) As Boolean

    Dim pivotField As pivotField

    On Error Resume Next
    Set pivotField = pivotTable.PivotFields(fieldName)
    On Error GoTo 0

    PivotFieldExists = Not pivotField Is Nothing

End Function

Private Function GetUniquePivotName( _
    ByVal wb As Workbook, _
    ByVal baseName As String) As String

    Dim sanitizedName As String
    Dim candidateName As String
    Dim suffix As Long

    sanitizedName = SanitizePivotName(baseName)

    If Len(sanitizedName) > 200 Then
        sanitizedName = Left$(sanitizedName, 200)
    End If

    If Not PivotNameExists(wb, sanitizedName) Then

        GetUniquePivotName = sanitizedName
        Exit Function

    End If

    suffix = 1

    Do

        candidateName = sanitizedName & "_" & suffix

        If Not PivotNameExists(wb, candidateName) Then

            GetUniquePivotName = candidateName
            Exit Function

        End If

        suffix = suffix + 1

    Loop

End Function

Private Function PivotNameExists( _
    ByVal wb As Workbook, _
    ByVal pivotName As String) As Boolean

    Dim ws As Worksheet
    Dim pivotTable As pivotTable

    For Each ws In wb.Worksheets

        For Each pivotTable In ws.PivotTables

            If StrComp( _
                pivotTable.Name, _
                pivotName, _
                vbTextCompare) = 0 Then

                PivotNameExists = True
                Exit Function

            End If

        Next pivotTable

    Next ws

End Function

Private Function SanitizePivotName( _
    ByVal sourceName As String) As String

    Dim unsupportedCharacters As Variant
    Dim character As Variant

    unsupportedCharacters = Array( _
        ":", "\", "/", "?", "*", "[", "]")

    sourceName = Replace$( _
        sourceName, _
        " ", _
        vbNullString)

    For Each character In unsupportedCharacters

        sourceName = Replace$( _
            sourceName, _
            CStr(character), _
            "_")

    Next character

    SanitizePivotName = sourceName

End Function

'============================================================
' LOGGING AND USER MESSAGES
'============================================================

Private Sub LogMissingField( _
    ByVal sheetName As String, _
    ByVal fieldName As String)

    LogMessage _
        "WARNING", _
        sheetName, _
        "Pivot field not found: " & fieldName

End Sub

Private Sub InitializeProcessingLog(ByVal wb As Workbook)

    Set processingLogWs = RecreateOutputSheet( _
        wb, _
        PROCESSING_LOG_SHEET_NAME)

    processingLogWs.Visible = xlSheetVisible
    processingLogNextRow = 2

    With processingLogWs

        .Range("A1:D1").Value = Array( _
            "Date and Time", _
            "Status", _
            "Worksheet", _
            "Details")

        With .Range("A1:D1")
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        .Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Columns("A").ColumnWidth = 20
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 32
        .Columns("D").ColumnWidth = 70

        .Range("A1:D1").AutoFilter

    End With

End Sub

Private Sub LogMessage( _
    ByVal statusText As String, _
    ByVal sheetName As String, _
    ByVal details As String)

    'Keep the Immediate Window output for developers.
    Debug.Print _
        statusText & _
        " | " & sheetName & _
        " | " & details

    'If the worksheet log has not yet been initialized,
    'the Immediate Window output above is still retained.
    If processingLogWs Is Nothing Then
        Exit Sub
    End If

    With processingLogWs

        .Cells(processingLogNextRow, 1).Value = Now
        .Cells(processingLogNextRow, 2).Value = statusText
        .Cells(processingLogNextRow, 3).Value = sheetName
        .Cells(processingLogNextRow, 4).Value = details

        ApplyLogStatusFormatting .Cells(processingLogNextRow, 2), _
            statusText

    End With

    processingLogNextRow = processingLogNextRow + 1

End Sub

Private Sub ApplyLogStatusFormatting( _
    ByVal targetCell As Range, _
    ByVal statusText As String)

    With targetCell

        .Font.Bold = True
        .HorizontalAlignment = xlCenter

        Select Case UCase$(Trim$(statusText))

            Case "PROCESSED", "COMPLETED"
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)

            Case "SKIPPED"
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 101, 0)

            Case "WARNING"
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)

            Case "ERROR"
                .Interior.Color = RGB(192, 0, 0)
                .Font.Color = RGB(255, 255, 255)

            Case "EXCLUDED"
                .Interior.Color = RGB(217, 217, 217)
                .Font.Color = RGB(89, 89, 89)

            Case "STARTED"
                .Interior.Color = RGB(221, 235, 247)
                .Font.Color = RGB(31, 78, 121)

        End Select

    End With

End Sub

Private Sub FormatProcessingLog()

    Dim lastLogRow As Long

    If processingLogWs Is Nothing Then
        Exit Sub
    End If

    lastLogRow = processingLogWs.Cells( _
        processingLogWs.Rows.Count, _
        1).End(xlUp).Row

    With processingLogWs

        .Columns("A:D").VerticalAlignment = xlTop
        .Columns("D").WrapText = True

        If lastLogRow >= 2 Then

            With .Range("A1:D" & lastLogRow)
                .Borders.LineStyle = xlContinuous
                .Borders.Color = RGB(217, 217, 217)
                .Borders.Weight = xlThin
            End With

        End If

        .Rows(1).RowHeight = 24

    End With

End Sub

Private Sub ShowCompletionMessage( _
    ByVal processedCount As Long, _
    ByVal skippedCount As Long, _
    ByVal warningCount As Long, _
    ByVal logWs As Worksheet)

    Dim messageText As String
    Dim messageStyle As VbMsgBoxStyle
    Dim messageTitle As String
    Dim userResponse As VbMsgBoxResult

    messageText = _
        processedCount & " worksheet(s) processed." & vbCrLf & _
        skippedCount & " worksheet(s) skipped." & vbCrLf & _
        warningCount & " field warning(s)."

    If skippedCount > 0 Or warningCount > 0 Then

        messageText = _
            messageText & vbCrLf & vbCrLf & _
            "Some worksheets require attention." & _
            vbCrLf & vbCrLf & _
            "Required worksheet naming format:" & vbCrLf & _
            "Bank-ID or Bank - ID" & vbCrLf & _
            "Example: Delta Haven Bank - PAY4464" & _
            vbCrLf & vbCrLf & _
            "Do not put parentheses around the Payment ID." & _
            vbCrLf & vbCrLf & _
            "Would you like to open the Processing Log?"

        messageStyle = vbInformation + vbYesNo
        messageTitle = "Report Completed with Warnings"

        userResponse = MsgBox( _
            Prompt:=messageText, _
            Buttons:=messageStyle, _
            Title:=messageTitle)

        If userResponse = vbYes Then

            If Not logWs Is Nothing Then
                logWs.Visible = xlSheetVisible
                logWs.Activate
                logWs.Range("A1").Select
            End If

        End If

    Else

        messageText = _
            messageText & vbCrLf & vbCrLf & _
            "The Payment Summary was generated successfully."

        messageStyle = vbInformation
        messageTitle = "Report Generation Complete"

        MsgBox _
            Prompt:=messageText, _
            Buttons:=messageStyle, _
            Title:=messageTitle

    End If

End Sub

Private Function GetDisplaySheetName( _
    ByVal sheetName As String) As String

    If Len(sheetName) = 0 Then

        GetDisplaySheetName = "(not yet determined)"

    Else

        GetDisplaySheetName = sheetName

    End If

End Function

