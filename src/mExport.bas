Attribute VB_Name = "mExport"
Option Explicit

' Export of the Data sheet to a vCard 3.0 file

Private Const MAX_LISTED_ROWS As Long = 20

' Field indexes of the name parts and special fields, resolved once per export
Private mGiven As Long
Private mMiddle As Long
Private mFamily As Long
Private mPrefix As Long
Private mSuffix As Long
Private mFn As Long
Private mBday As Long
Private mOrg As Long

Private mNormalize As Boolean
Private mCountry As String
Private mTrunk As String
Private mFnLastFirst As Boolean
Private mFold As Boolean

Public Sub Create_vCard_File()
    Dim ws As Worksheet, fso As Object
    Dim data As Variant, cards() As String, include() As Boolean
    Dim Target As Variant, path As String
    Dim firstRow As Long, lastRow As Long, lastCol As Long, nRows As Long
    Dim i As Long, stepRows As Long, saved As Long, picked As Long
    Dim noName As Long, badBirthday As Long, filtered As Long
    Dim noNameRows As String, badBirthdayRows As String
    Dim card As String, birthdayFailed As Boolean, filterOn As Boolean, selectedOnly As Boolean
    Dim msg As String, errText As String

    On Error GoTo ErrHandler

    LoadFields
    Set ws = DataSheet()
    If Not ActiveSheet Is ws Then ws.Activate

    firstRow = HEADER_ROW + 1
    lastCol = LastFieldCol()
    lastRow = FindLastRow(ws, firstRow, lastCol)
    If lastRow < firstRow Then
        MsgBox Txt(S_NO_DATA), vbInformation
        Exit Sub
    End If
    nRows = lastRow - firstRow + 1

    selectedOnly = (LCase$(SettingText("export_scope")) = "selected rows")
    If selectedOnly Then
        ReDim include(1 To nRows)
        picked = MarkSelectedRows(ws, firstRow, nRows, include)
        If picked = 0 Then
            MsgBox Txt(S_NO_SELECTION), vbInformation
            Exit Sub
        End If
    End If

    Target = Application.GetSaveAsFilename(FileFilter:=Txt(S_FILE_TYPE) & " (*.vcf),*.vcf", Title:=Txt(S_SAVE_PROMPT))
    If VarType(Target) = vbBoolean Then Exit Sub
    path = CStr(Target)
    If LCase$(Right$(path, 4)) <> ".vcf" Then path = path & ".vcf"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        If MsgBox(path & vbCrLf & Txt(S_FILE_EXISTS) & vbCrLf & Txt(S_OVERWRITE), vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If

    ResolveExportFields
    LogClear
    ProgressStart

    ' Read the whole range at once; cell-by-cell access is much slower
    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).value
    ReDim cards(0 To nRows - 1)
    ' Rows hidden by an active AutoFilter are not exported
    filterOn = ws.FilterMode
    stepRows = nRows \ 100
    If stepRows < 1 Then stepRows = 1

    For i = 1 To nRows
        If RowIsBlank(data, i, lastCol) Then GoTo NextRow
        If selectedOnly Then
            If Not include(i) Then GoTo NextRow
        End If
        If filterOn Then
            If ws.Rows(firstRow + i - 1).Hidden Then
                filtered = filtered + 1
                GoTo NextRow
            End If
        End If

        ' An empty card means the row has neither a name nor an organization
        card = BuildCard(data, i, birthdayFailed)
        If Len(card) = 0 Then
            AddRowNumber noName, noNameRows, firstRow + i - 1
            LogAdd "#" & (firstRow + i - 1), vbNullString, Txt(S_R_NO_NAME), vbNullString
        Else
            cards(saved) = card
            saved = saved + 1
            If birthdayFailed Then
                AddRowNumber badBirthday, badBirthdayRows, firstRow + i - 1
                LogAdd "#" & (firstRow + i - 1), "BDAY", Txt(S_R_BAD_BIRTHDAY), CellText(data(i, gFldCol(mBday)))
            End If
        End If

NextRow:
        ' Redraw the progress bar about once per 1% of rows
        If i Mod stepRows = 0 Or i = nRows Then ProgressUpdate i / nRows
    Next i

    If saved > 0 Then
        ReDim Preserve cards(0 To saved - 1)
        SaveText path, Join(cards, vbNullString), SettingYes("export_bom")
    End If

    ProgressEnd

    If saved > 0 Then
        msg = saved & " " & Txt(S_EXPORTED_TO) & vbCrLf & path
    Else
        msg = Txt(S_NO_DATA)
    End If
    If noName > 0 Then
        msg = msg & vbCrLf & vbCrLf & noName & " " & Txt(S_NO_NAME) & vbCrLf & Txt(S_ROW_NUMBERS) & ": " & noNameRows
    End If
    If badBirthday > 0 Then
        msg = msg & vbCrLf & vbCrLf & badBirthday & " " & Txt(S_BAD_BIRTHDAY) & vbCrLf & Txt(S_ROW_NUMBERS) & ": " & badBirthdayRows
    End If
    If filtered > 0 Then
        msg = msg & vbCrLf & vbCrLf & filtered & " " & Txt(S_FILTERED)
    End If
    If LogFlush(Txt(S_OP_EXPORT)) Then msg = msg & vbCrLf & vbCrLf & Txt(S_LOG_WRITTEN)

    If saved > 0 Then Beep
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    errText = Err.Description
    ProgressEnd
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText, vbExclamation
End Sub

' Resolves the name and special fields and reads the export settings
Private Sub ResolveExportFields()
    mGiven = FieldIndexByFamily("N_GIVEN")
    mMiddle = FieldIndexByFamily("N_MIDDLE")
    mFamily = FieldIndexByFamily("N_FAMILY")
    mPrefix = FieldIndexByFamily("N_PREFIX")
    mSuffix = FieldIndexByFamily("N_SUFFIX")
    mFn = FieldIndexByFamily("FN")
    mBday = FieldIndexByFamily("BDAY")
    mOrg = FieldIndexByFamily("ORG")

    mNormalize = PhoneFormatEnabled("export")
    mCountry = SettingText("country_code")
    mTrunk = TrunkPrefix()
    mFnLastFirst = (LCase$(SettingText("fn_format")) = "last first")
    mFold = SettingYes("fold_lines")
End Sub

' Marks the data rows inside the current selection; returns how many rows were marked
Private Function MarkSelectedRows(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal nRows As Long, ByRef include() As Boolean) As Long
    Dim area As Range
    Dim r As Long, lo As Long, hi As Long, k As Long, cnt As Long

    If Not ActiveSheet Is ws Then Exit Function
    If TypeName(Selection) <> "Range" Then Exit Function
    For Each area In Selection.Areas
        lo = area.Row
        If lo < firstRow Then lo = firstRow
        hi = area.Row + area.Rows.Count - 1
        If hi > firstRow + nRows - 1 Then hi = firstRow + nRows - 1
        For r = lo To hi
            k = r - firstRow + 1
            If Not include(k) Then
                include(k) = True
                cnt = cnt + 1
            End If
        Next r
    Next area
    MarkSelectedRows = cnt
End Function

Private Function RowIsBlank(ByRef data As Variant, ByVal i As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    For c = 1 To lastCol
        If Len(CellText(data(i, c))) > 0 Then Exit Function
    Next c
    RowIsBlank = True
End Function

Private Function FieldValue(ByRef data As Variant, ByVal i As Long, ByVal FieldIdx As Long) As String
    If FieldIdx = 0 Then Exit Function
    FieldValue = CellText(data(i, gFldCol(FieldIdx)))
End Function

' Builds one vCard 3.0 entry for data row i; returns "" if the row has no name and no organization
Private Function BuildCard(ByRef data As Variant, ByVal i As Long, ByRef birthdayFailed As Boolean) As String
    Dim givenName As String, middleName As String, familyName As String, prefixName As String, suffixName As String
    Dim fnValue As String, composed As String, fn As String, org As String, birthday As String
    Dim v As String, line As String, sb As String
    Dim f As Long

    birthdayFailed = False
    givenName = FieldValue(data, i, mGiven)
    middleName = FieldValue(data, i, mMiddle)
    familyName = FieldValue(data, i, mFamily)
    prefixName = FieldValue(data, i, mPrefix)
    suffixName = FieldValue(data, i, mSuffix)
    fnValue = FieldValue(data, i, mFn)
    org = FieldValue(data, i, mOrg)

    If mFnLastFirst Then
        composed = JoinWords(familyName, givenName, middleName)
    Else
        composed = JoinWords(givenName, middleName, familyName)
    End If
    fn = fnValue
    If Len(fn) = 0 Then fn = composed
    If Len(fn) = 0 Then fn = org
    If Len(fn) = 0 Then Exit Function

    ' N is Family;Given;Middle;Prefix;Suffix. A contact with only an organization gets an empty N and
    ' the company as FN; X-ABShowAs makes Apple contacts display it as a company.
    sb = "BEGIN:VCARD" & vbCrLf & "VERSION:3.0" & vbCrLf
    sb = sb & "N:" & Esc(familyName) & ";" & Esc(givenName) & ";" & Esc(middleName) & ";" & Esc(prefixName) & ";" & Esc(suffixName) & vbCrLf
    sb = sb & Folded("FN:" & Esc(fn)) & vbCrLf
    If Len(composed) = 0 And Len(fnValue) = 0 Then sb = sb & "X-ABShowAs:COMPANY" & vbCrLf

    If mBday > 0 Then
        v = CellText(data(i, gFldCol(mBday)))
        If Len(v) > 0 Then
            If TryFormatBirthday(data(i, gFldCol(mBday)), birthday) Then
                sb = sb & "BDAY:" & birthday & vbCrLf
            Else
                birthdayFailed = True
            End If
        End If
    End If

    For f = 1 To gFldCount
        If gFldCol(f) > 0 Then
            Select Case gFldProp(f)
                Case "N", "FN", "BDAY"
                Case Else
                    v = CellText(data(i, gFldCol(f)))
                    If Len(v) > 0 Then
                        line = FieldLine(f, v)
                        If Len(line) > 0 Then sb = sb & Folded(line) & vbCrLf
                    End If
            End Select
        End If
    Next f

    BuildCard = sb & "END:VCARD" & vbCrLf
End Function

' Formats one property line. ADR is a structured value (street is the 3rd component),
' URLs are not text-escaped, and CATEGORIES is a comma-separated list.
Private Function FieldLine(ByVal f As Long, ByVal v As String) As String
    Dim prefix As String
    prefix = gFldProp(f)
    If Len(gFldParams(f)) > 0 Then prefix = prefix & ";" & gFldParams(f)
    prefix = prefix & ":"
    Select Case gFldProp(f)
        Case "ADR"
            FieldLine = prefix & ";;" & Esc(v) & ";;;;"
        Case "URL"
            FieldLine = prefix & Replace(Replace(v, vbCr, vbNullString), vbLf, vbNullString)
        Case "CATEGORIES"
            v = EscCategories(v)
            If Len(v) > 0 Then FieldLine = prefix & v
        Case "TEL"
            If mNormalize Then v = NormalizePhone(v, mCountry, mTrunk)
            FieldLine = prefix & Esc(v)
        Case Else
            FieldLine = prefix & Esc(v)
    End Select
End Function

Private Function Folded(ByVal line As String) As String
    If mFold Then
        Folded = FoldLine(line)
    Else
        Folded = line
    End If
End Function

' Counts a row but keeps only the first MAX_LISTED_ROWS numbers for display
Private Sub AddRowNumber(ByRef n As Long, ByRef list As String, ByVal r As Long)
    n = n + 1
    If n <= MAX_LISTED_ROWS Then
        If Len(list) > 0 Then list = list & ", "
        list = list & r
    ElseIf n = MAX_LISTED_ROWS + 1 Then
        list = list & ", ..."
    End If
End Sub
