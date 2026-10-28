Attribute VB_Name = "mTools"
Option Explicit

' Progress bar, log sheet and the Remove Duplicates / Clear Data buttons

' Maximum number of rows written to the Log sheet
Public Const MAX_LOG_ROWS As Long = 5000

Private mBar1 As Shape
Private mBar2 As Shape
Private mBarWidth As Double

Private mLogContact() As String
Private mLogField() As String
Private mLogReason() As String
Private mLogValue() As String
Private mLogN As Long

' Shows the progress bar on the Data sheet (shp_rec1 is the grey track, shp_rec2 the filled part)
Public Sub ProgressStart()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DataSheet()
    Set mBar1 = ws.Shapes("shp_rec1")
    Set mBar2 = ws.Shapes("shp_rec2")
    mBarWidth = mBar1.Width
    mBar2.Left = mBar1.Left
    mBar2.Top = mBar1.Top
    mBar2.Height = mBar1.Height
    mBar2.Width = 0
    mBar2.TextFrame.Characters.text = vbNullString
    mBar1.Visible = msoTrue
    mBar2.Visible = msoTrue
End Sub

Public Sub ProgressUpdate(ByVal pct As Double)
    If mBar2 Is Nothing Then Exit Sub
    On Error Resume Next
    mBar2.Width = mBarWidth * pct
    mBar2.TextFrame.Characters.text = Format$(pct, "0%")
    DoEvents
End Sub

' Safe to call at any time, also when an error occurs before the bar was shown
Public Sub ProgressEnd()
    On Error Resume Next
    mBar1.Visible = msoFalse
    mBar2.Visible = msoFalse
    Set mBar1 = Nothing
    Set mBar2 = Nothing
    Application.StatusBar = False
End Sub

Public Sub StatusText(ByVal s As String)
    Application.StatusBar = s
    DoEvents
End Sub

' Last row that has data in columns 1..lastCol at or below firstRow; 0 when there is none
Public Function FindLastRow(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastCol As Long) As Long
    Dim found As Range
    Set found = ws.Range(ws.Cells(firstRow, 1), ws.Cells(ws.Rows.Count, lastCol)).Find( _
        What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not found Is Nothing Then FindLastRow = found.Row
End Function

Public Sub LogClear()
    ReDim mLogContact(0 To 63)
    ReDim mLogField(0 To 63)
    ReDim mLogReason(0 To 63)
    ReDim mLogValue(0 To 63)
    mLogN = 0
End Sub

Public Sub LogAdd(ByVal contact As String, ByVal fieldName As String, ByVal reason As String, ByVal value As String)
    If mLogN >= MAX_LOG_ROWS Then Exit Sub
    If mLogN > UBound(mLogContact) Then
        ReDim Preserve mLogContact(0 To mLogN * 2)
        ReDim Preserve mLogField(0 To mLogN * 2)
        ReDim Preserve mLogReason(0 To mLogN * 2)
        ReDim Preserve mLogValue(0 To mLogN * 2)
    End If
    mLogContact(mLogN) = contact
    mLogField(mLogN) = fieldName
    mLogReason(mLogN) = reason
    mLogValue(mLogN) = Left$(value, 200)
    mLogN = mLogN + 1
End Sub

' Writes the collected entries to the Log sheet (created when missing); returns True when something was written
Public Function LogFlush(ByVal opName As String) As Boolean
    Dim ws As Worksheet, out() As Variant
    Dim i As Long

    If mLogN = 0 Then Exit Function
    If Not SettingYes("write_log") Then Exit Function

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Log"
    End If
    ws.Cells.Clear

    ReDim out(1 To mLogN + 1, 1 To 5)
    out(1, 1) = Txt(S_LOG_OPERATION)
    out(1, 2) = Txt(S_LOG_CONTACT)
    out(1, 3) = Txt(S_LOG_FIELD)
    out(1, 4) = Txt(S_LOG_REASON)
    out(1, 5) = Txt(S_LOG_VALUE)
    For i = 0 To mLogN - 1
        out(i + 2, 1) = opName
        out(i + 2, 2) = mLogContact(i)
        out(i + 2, 3) = mLogField(i)
        out(i + 2, 4) = mLogReason(i)
        out(i + 2, 5) = mLogValue(i)
    Next i
    ws.Range(ws.Cells(1, 1), ws.Cells(mLogN + 1, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(mLogN + 1, 5)).value = out
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:E").AutoFit
    LogFlush = True

    ' Adding or filling the Log sheet activates it; return to the Data sheet
    On Error Resume Next
    DataSheet().Activate
End Function

' Removes duplicate contact rows with Excel's own Remove Duplicates (all columns or the name columns).
' If Excel rejects the call, the same result is produced by ManualRemove.
Public Sub Remove_Duplicates()
    Dim ws As Worksheet, rng As Range
    Dim cols() As Variant
    Dim firstRow As Long, lastRow As Long, lastCol As Long, before As Long, after As Long
    Dim n As Long, i As Long, k As Long, idx As Long
    Dim names As Variant
    Dim stage As String, errText As String, msg As String, removeErr As String
    Dim failed As Boolean

    On Error GoTo ErrHandler

    stage = "fields"
    LoadFields
    Set ws = DataSheet()
    If Not ActiveSheet Is ws Then ws.Activate
    If ws.FilterMode Then
        MsgBox Txt(S_CLEAR_FILTER), vbExclamation
        Exit Sub
    End If

    stage = "range"
    firstRow = HEADER_ROW + 1
    lastCol = LastFieldCol()
    lastRow = FindLastRow(ws, firstRow, lastCol)
    If lastRow <= firstRow Then
        MsgBox Txt(S_DEDUP_NONE), vbInformation
        Exit Sub
    End If

    If MsgBox(Txt(S_DEDUP_CONFIRM) & vbCrLf & vbCrLf & OptionLabel("dup_match"), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    stage = "columns"
    If LCase$(SettingText("dup_match")) = "name columns" Then
        ReDim cols(0 To 2)
        names = Array("N_GIVEN", "N_MIDDLE", "N_FAMILY")
        For i = 0 To 2
            idx = FieldIndexByFamily(CStr(names(i)))
            If idx > 0 Then
                cols(n) = CInt(gFldCol(idx))
                n = n + 1
            End If
        Next i
        If n = 0 Then Exit Sub
        ReDim Preserve cols(0 To n - 1)
    Else
        ReDim cols(0 To lastCol - 1)
        For k = 1 To lastCol
            cols(k - 1) = CInt(k)
        Next k
    End If

    before = lastRow - firstRow + 1
    Set rng = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol))

    stage = "RemoveDuplicates"
    Err.Clear
    On Error Resume Next
    rng.RemoveDuplicates Columns:=AsVariantArray(cols), header:=xlNo
    failed = (Err.Number <> 0)
    If failed Then removeErr = Err.Description
    Err.Clear
    On Error GoTo ErrHandler
    If failed Then
        stage = "manual"
        ManualRemove rng, cols
    End If

    stage = "result"
    after = FindLastRow(ws, firstRow, lastCol) - firstRow + 1
    If after < 0 Then after = 0

    ' The rows freed at the end keep their old fill color, so they are deleted
    If before - after > 0 Then ws.Rows((firstRow + after) & ":" & lastRow).Delete

    If before - after > 0 Then
        msg = (before - after) & " " & Txt(S_DEDUP_DONE)
        If failed Then msg = msg & vbCrLf & "[manual: " & removeErr & "]"
        MsgBox msg, vbInformation
    Else
        MsgBox Txt(S_DEDUP_NONE), vbInformation
    End If
    Exit Sub

ErrHandler:
    errText = Err.Description
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText & " [" & stage & "]", vbExclamation
End Sub

' Keeps the first row of each duplicate group and moves the others out; the freed rows at the end stay empty
Private Sub ManualRemove(ByVal rng As Range, ByRef cols As Variant)
    Dim data As Variant, out() As Variant, keep() As Boolean
    Dim n As Long, m As Long, i As Long, c As Long, w As Long

    data = rng.value
    n = UBound(data, 1)
    m = MarkUniqueRows(data, cols, keep)
    If m = n Then Exit Sub

    ReDim out(1 To n, 1 To UBound(data, 2))
    For i = 1 To n
        If keep(i) Then
            w = w + 1
            For c = 1 To UBound(data, 2)
                out(w, c) = data(i, c)
            Next c
        End If
    Next i
    rng.value = out
End Sub

' Flags the first occurrence of every key (case-insensitive like Excel) built from the given columns;
' returns how many rows are kept.
Public Function MarkUniqueRows(ByRef data As Variant, ByRef cols As Variant, ByRef keep() As Boolean) As Long
    Dim seen As Collection
    Dim n As Long, i As Long, k As Long, kept As Long
    Dim key As String, part As String
    Dim dup As Boolean

    n = UBound(data, 1)
    ReDim keep(1 To n)
    Set seen = New Collection
    For i = 1 To n
        key = vbNullString
        For k = LBound(cols) To UBound(cols)
            If IsError(data(i, cols(k))) Then
                part = "#ERR"
            Else
                part = CStr(data(i, cols(k)))
            End If
            key = key & part & Chr$(1)
        Next k
        On Error Resume Next
        seen.Add i, key
        dup = (Err.Number <> 0)
        Err.Clear
        On Error GoTo 0
        If Not dup Then
            keep(i) = True
            kept = kept + 1
        End If
    Next i
    MarkUniqueRows = kept
End Function

' Deletes all contact rows after confirmation
Public Sub Clear_Data()
    Dim ws As Worksheet
    Dim firstRow As Long, lastRow As Long
    Dim errText As String

    On Error GoTo ErrHandler

    LoadFields
    Set ws = DataSheet()
    If Not ActiveSheet Is ws Then ws.Activate
    If ws.FilterMode Then
        MsgBox Txt(S_CLEAR_FILTER), vbExclamation
        Exit Sub
    End If

    ' UsedRange also covers empty rows that still have a fill color or format (e.g. left behind by Remove Duplicates).
    ' Rows are deleted up to the last used one rather than to the end of the sheet, so the validation and
    ' conditional formatting that cover the lower rows stay in place.
    firstRow = HEADER_ROW + 1
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow < firstRow Then
        MsgBox Txt(S_NO_DATA), vbInformation
        Exit Sub
    End If

    If MsgBox(Txt(S_CLEAR_CONFIRM), vbYesNo + vbExclamation + vbDefaultButton2) = vbNo Then Exit Sub
    ws.Rows(firstRow & ":" & lastRow).Delete
    MsgBox (lastRow - firstRow + 1) & " " & Txt(S_CLEAR_DONE), vbInformation
    Exit Sub

ErrHandler:
    errText = Err.Description
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText, vbExclamation
End Sub

Private Function AsVariantArray(ByRef values() As Variant) As Variant
    AsVariantArray = values
End Function
