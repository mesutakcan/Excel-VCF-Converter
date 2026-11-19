Attribute VB_Name = "mMergeApply"
' Mesut Akcan
' https://github.com/mesutakcan/Excel-VCF-Converter
' youtube.com/mesutakcan
Option Explicit

' Carries out the plan made by mMerge: the records of every group of the same person are merged into the main
' record and the other rows are deleted. Rows that are identical to another row are deleted in every case.
' Everything is computed in memory first; the deleted rows are copied to the MergeBackup sheet before the Data
' sheet is changed, and the rows are deleted last.

Private Const METHOD_KEEP As Long = 1
Private Const METHOD_FILL As Long = 2
Private Const METHOD_COMBINE As Long = 3

Private Const KIND_SINGLE As Long = 0
Private Const KIND_SLOT As Long = 1
Private Const KIND_NOTE As Long = 2
Private Const KIND_LABELS As Long = 3

' Columns that one merge may add for values that do not fit
Private Const MAX_NEW_COLS As Long = 30
' Columns in front of the Data columns on the MergeBackup sheet
Private Const BK_LEAD As Long = 4

Private mMethod As Long
Private mAddColumns As Boolean
Private mUnplaced As Long
Private mStamp As String

Private mRemove() As Boolean
Private mRestored() As Boolean

' Columns added by this merge: a field whose column was deleted is restored (xField > 0), otherwise a new field is added
Private xCount As Long
Private xField() As Long
Private xKey() As String
Private xProp() As String
Private xParams() As String
Private xFamily() As String
Private xHdrEN() As String
Private xHdrTR() As String
Private xCol() As Long
Private xVal() As String

' Cells of the Data sheet that change (record, column)
Private chRec() As Long
Private chCol() As Long
Private chN As Long

' Lines of the MergeBackup block; line 1 is the header
Private bk() As Variant
Private bkN As Long

' ---------------------------------------------------------------- value rules

Private Function MergeKind(ByVal family As String) As Long
    If Left$(family, 4) = "TEL|" Or Left$(family, 6) = "EMAIL|" Or Left$(family, 4) = "ADR|" Or family = "URL" Then
        MergeKind = KIND_SLOT
    ElseIf family = "NOTE" Then
        MergeKind = KIND_NOTE
    ElseIf family = "CATEGORIES" Then
        MergeKind = KIND_LABELS
    Else
        MergeKind = KIND_SINGLE
    End If
End Function

' Values of these groups are compared across all fields of the group (a number written under another label is the same number)
Private Function IdGroup(ByVal prop As String) As Long
    Select Case prop
        Case "TEL": IdGroup = 1
        Case "EMAIL": IdGroup = 2
        Case "ADR": IdGroup = 3
        Case "URL": IdGroup = 4
    End Select
End Function

Private Function ValueId(ByVal group As Long, ByVal v As String) As String
    Dim t As String
    Select Case group
        Case 1
            ValueId = PhoneIdentity(v)
        Case 2
            ValueId = LowerAscii(v)
        Case 4
            t = FoldText(v)
            If Left$(t, 6) = "https " Then t = Mid$(t, 7)
            If Left$(t, 5) = "http " Then t = Mid$(t, 6)
            If Left$(t, 4) = "www " Then t = Mid$(t, 5)
            ValueId = t
        Case Else
            ValueId = FoldText(v)
    End Select
End Function

Private Sub AddId(ByRef ids() As String, ByVal group As Long, ByVal v As String)
    Dim id As String
    If group < 1 Then Exit Sub
    id = ValueId(group, v)
    If Len(id) > 0 Then
        If InStr(ids(group), "|" & id & "|") = 0 Then ids(group) = ids(group) & id & "|"
    End If
End Sub

Private Function HasId(ByRef ids() As String, ByVal group As Long, ByVal v As String) As Boolean
    Dim id As String
    If group < 1 Then Exit Function
    id = ValueId(group, v)
    If Len(id) > 0 Then HasId = (InStr(ids(group), "|" & id & "|") > 0)
End Function

' Identities of the phone numbers, e-mails, addresses and websites the main record has now
Private Sub LoadIds(ByVal p As Long, ByRef ids() As String)
    Dim f As Long, g As Long, k As Long
    Dim v As String

    For g = 1 To 4
        ids(g) = "|"
    Next g
    For f = 1 To gFldCount
        If gFldCol(f) > 0 Then
            If MergeKind(gFldFamily(f)) = KIND_SLOT Then
                v = CellText(gData(p, gFldCol(f)))
                If Len(v) > 0 Then AddId ids, IdGroup(gFldProp(f)), v
            End If
        End If
    Next f
    For k = 1 To xCount
        If Len(xVal(p, k)) > 0 Then AddId ids, IdGroup(xProp(k)), xVal(p, k)
    Next k
End Sub

Private Function FieldHeader(ByVal f As Long) As String
    If LangNo() = 2 And Len(gFldHdrTR(f)) > 0 Then
        FieldHeader = gFldHdrTR(f)
    Else
        FieldHeader = gFldHdrEN(f)
    End If
End Function

Private Sub LogDifferent(ByVal p As Long, ByVal f As Long, ByVal v As String)
    LogAdd RecordName(gData, p), FieldHeader(f), Txt(S_MRG_LOG_DIFF), v
End Sub

' ---------------------------------------------------------------- changing the main record

Private Sub SetCell(ByVal p As Long, ByVal c As Long, ByVal v As Variant)
    gData(p, c) = v
    chN = chN + 1
    If chN > UBound(chRec) Then
        GrowLongs chRec, UBound(chRec) * 2
        GrowLongs chCol, UBound(chCol) * 2
    End If
    chRec(chN) = p
    chCol(chN) = c
End Sub

Private Function KeyInUse(ByVal key As String) As Boolean
    Dim k As Long
    If FieldIndexByKey(key) > 0 Then
        KeyInUse = True
        Exit Function
    End If
    For k = 1 To xCount
        If xKey(k) = key Then
            KeyInUse = True
            Exit Function
        End If
    Next k
End Function

Private Function NewColumn(ByVal restoreField As Long, ByVal key As String, ByVal prop As String, ByVal params As String, _
                           ByVal family As String, ByVal hdrEN As String, ByVal hdrTR As String) As Long
    If xCount = 0 Then
        ReDim xField(1 To MAX_NEW_COLS)
        ReDim xKey(1 To MAX_NEW_COLS)
        ReDim xProp(1 To MAX_NEW_COLS)
        ReDim xParams(1 To MAX_NEW_COLS)
        ReDim xFamily(1 To MAX_NEW_COLS)
        ReDim xHdrEN(1 To MAX_NEW_COLS)
        ReDim xHdrTR(1 To MAX_NEW_COLS)
        ReDim xVal(1 To gN, 1 To MAX_NEW_COLS)
    End If
    xCount = xCount + 1
    xField(xCount) = restoreField
    xKey(xCount) = key
    xProp(xCount) = prop
    xParams(xCount) = params
    xFamily(xCount) = family
    xHdrEN(xCount) = hdrEN
    xHdrTR(xCount) = hdrTR
    NewColumn = xCount
End Function

' Another column for the family: a field whose column was deleted is restored first, otherwise a field is added that is
' cloned from the last column of the family. Returns its number among the added columns, 0 when none can be added.
Private Function AddColumnFor(ByVal family As String) As Long
    Dim f As Long, k As Long, found As Boolean
    Dim key As String, prop As String, params As String, hdrEN As String, hdrTR As String

    If xCount >= MAX_NEW_COLS Then Exit Function

    For f = 1 To gFldCount
        If gFldFamily(f) = family And gFldCol(f) = 0 And Not mRestored(f) Then
            mRestored(f) = True
            AddColumnFor = NewColumn(f, gFldKey(f), gFldProp(f), gFldParams(f), family, gFldHdrEN(f), gFldHdrTR(f))
            Exit Function
        End If
    Next f

    For k = xCount To 1 Step -1
        If xFamily(k) = family Then
            key = xKey(k)
            prop = xProp(k)
            params = xParams(k)
            hdrEN = xHdrEN(k)
            hdrTR = xHdrTR(k)
            found = True
            Exit For
        End If
    Next k
    If Not found Then
        For f = gFldCount To 1 Step -1
            If gFldFamily(f) = family And gFldCol(f) > 0 Then
                key = gFldKey(f)
                prop = gFldProp(f)
                params = gFldParams(f)
                hdrEN = gFldHdrEN(f)
                hdrTR = gFldHdrTR(f)
                found = True
                Exit For
            End If
        Next f
    End If
    If Not found Then Exit Function

    key = IncrementTrailing(key)
    Do While KeyInUse(key)
        key = IncrementTrailing(key)
    Loop
    AddColumnFor = NewColumn(0, key, prop, params, family, IncrementTrailing(hdrEN), IncrementTrailing(hdrTR))
End Function

' Empty cell of the family in the main record: a column of the Data sheet first, then a column added by this merge
Private Function FindFreeSlot(ByVal p As Long, ByVal family As String, ByRef slotField As Long, ByRef slotX As Long) As Boolean
    Dim f As Long, k As Long

    slotField = 0
    slotX = 0
    For f = 1 To gFldCount
        If gFldFamily(f) = family And gFldCol(f) > 0 Then
            If Len(CellText(gData(p, gFldCol(f)))) = 0 Then
                slotField = f
                FindFreeSlot = True
                Exit Function
            End If
        End If
    Next f
    For k = 1 To xCount
        If xFamily(k) = family Then
            If Len(xVal(p, k)) = 0 Then
                slotX = k
                FindFreeSlot = True
                Exit Function
            End If
        End If
    Next k
End Function

' Puts a value into a free cell of its family; "Combine all" adds a column when none is free (and new columns are allowed)
Private Function PlaceSlot(ByVal p As Long, ByVal family As String, ByVal v As String) As Boolean
    Dim slotField As Long, slotX As Long

    If Not FindFreeSlot(p, family, slotField, slotX) Then
        If mMethod = METHOD_COMBINE And mAddColumns Then slotX = AddColumnFor(family)
    End If
    If slotField > 0 Then
        SetCell p, gFldCol(slotField), v
        PlaceSlot = True
    ElseIf slotX > 0 Then
        xVal(p, slotX) = v
        PlaceSlot = True
    End If
End Function

Private Sub MergeNote(ByVal p As Long, ByVal f As Long, ByVal c As Long, ByVal v As String)
    Dim cur As String
    cur = CellText(gData(p, c))
    If Len(cur) = 0 Then
        SetCell p, c, v
    ElseIf mMethod = METHOD_COMBINE Then
        If InStr(FoldText(cur), FoldText(v)) = 0 Then SetCell p, c, cur & vbLf & v
    ElseIf FoldText(cur) <> FoldText(v) Then
        LogDifferent p, f, v
    End If
End Sub

Private Sub MergeLabels(ByVal p As Long, ByVal f As Long, ByVal c As Long, ByVal v As String)
    Dim cur As String, res As String, seen As String, id As String
    Dim parts() As String
    Dim k As Long

    cur = CellText(gData(p, c))
    If Len(cur) = 0 Then
        SetCell p, c, v
    ElseIf mMethod <> METHOD_COMBINE Then
        If FoldText(cur) <> FoldText(v) Then LogDifferent p, f, v
    Else
        seen = "|"
        parts = Split(Replace(cur, ";", ","), ",")
        For k = 0 To UBound(parts)
            seen = seen & FoldText(parts(k)) & "|"
        Next k
        res = cur
        parts = Split(Replace(v, ";", ","), ",")
        For k = 0 To UBound(parts)
            id = FoldText(parts(k))
            If Len(id) > 0 Then
                If InStr(seen, "|" & id & "|") = 0 Then
                    res = res & ", " & Trim$(parts(k))
                    seen = seen & id & "|"
                End If
            End If
        Next k
        If res <> cur Then SetCell p, c, res
    End If
End Sub

' Takes what record s has and the main record p lacks
Private Sub MergeFrom(ByVal p As Long, ByVal s As Long, ByRef ids() As String)
    Dim f As Long, c As Long, g As Long
    Dim v As String, cur As String

    For f = 1 To gFldCount
        c = gFldCol(f)
        If c > 0 Then
            v = CellText(gData(s, c))
            If Len(v) > 0 Then
                Select Case MergeKind(gFldFamily(f))
                    Case KIND_SLOT
                        g = IdGroup(gFldProp(f))
                        If Not HasId(ids, g, v) Then
                            If PlaceSlot(p, gFldFamily(f), v) Then
                                AddId ids, g, v
                            Else
                                mUnplaced = mUnplaced + 1
                                LogAdd RecordName(gData, p), FieldHeader(f), Txt(S_R_NO_FREE), v
                            End If
                        End If
                    Case KIND_NOTE
                        MergeNote p, f, c, v
                    Case KIND_LABELS
                        MergeLabels p, f, c, v
                    Case Else
                        cur = CellText(gData(p, c))
                        If Len(cur) = 0 Then
                            SetCell p, c, gData(s, c)
                        ElseIf FoldText(cur) <> FoldText(v) Then
                            LogDifferent p, f, v
                        End If
                End Select
            End If
        End If
    Next f
End Sub

' Keeps the first occurrence of each phone, e-mail, address or URL in the main
' record.  The same value may be stored under different labels (for example,
' mobile and work), so identities are compared across all fields of the type.
Private Sub RemoveDuplicateSlots(ByVal p As Long)
    Dim ids(1 To 4) As String
    Dim f As Long, c As Long, g As Long
    Dim v As String

    For g = 1 To 4
        ids(g) = "|"
    Next g
    For c = 1 To gLastCol
        For f = 1 To gFldCount
            If gFldCol(f) = c And MergeKind(gFldFamily(f)) = KIND_SLOT Then
                g = IdGroup(gFldProp(f))
                v = CellText(gData(p, c))
                If Len(v) > 0 And g > 0 Then
                    If HasId(ids, g, v) Then
                        SetCell p, c, vbNullString
                    Else
                        AddId ids, g, v
                    End If
                End If
                Exit For
            End If
        Next f
    Next c
End Sub

Private Sub MergeCluster(ByVal root As Long)
    Dim ids(1 To 4) As String
    Dim p As Long, s As Long, k As Long

    p = gPrimary(root)
    If mMethod <> METHOD_KEEP Then
        LoadIds p, ids
        For k = 1 To MemberCount(root)
            s = MemberAt(root, k)
            If s <> p And Not gIsCopy(s) Then MergeFrom p, s, ids
        Next k
    End If
    RemoveDuplicateSlots p
End Sub

' ---------------------------------------------------------------- backup and log lines

Private Function DistinctCount(ByVal root As Long) As Long
    Dim k As Long
    For k = 1 To MemberCount(root)
        If Not gIsCopy(MemberAt(root, k)) Then DistinctCount = DistinctCount + 1
    Next k
End Function

Private Sub BackupLine(ByVal i As Long, ByVal reason As String, ByVal keptRow As Long)
    Dim c As Long
    bkN = bkN + 1
    bk(bkN, 1) = mStamp
    bk(bkN, 2) = RecordName(gData, keptRow)
    bk(bkN, 3) = reason
    bk(bkN, 4) = gFirstRow + i - 1
    For c = 1 To gLastCol
        bk(bkN, BK_LEAD + c) = CellText(gData(i, c))
    Next c
End Sub

' Copies a row that will be deleted to the backup and notes it in the log
Private Sub RemoveLine(ByVal s As Long, ByVal keptRow As Long)
    If gIsCopy(s) Then
        BackupLine s, Txt(S_MRG_BK_COPY), keptRow
        LogAdd RecordName(gData, s), Txt(S_MRG_COL_ROW) & " " & (gFirstRow + s - 1), Txt(S_MRG_LOG_COPY) & " " & (gFirstRow + gRepOf(s) - 1), RecordPhones(gData, s)
    Else
        BackupLine s, Txt(S_MRG_BK_MERGED), keptRow
        LogAdd RecordName(gData, s), Txt(S_MRG_COL_ROW) & " " & (gFirstRow + s - 1), Txt(S_MRG_LOG_INTO) & " " & (gFirstRow + keptRow - 1), RecordPhones(gData, s)
    End If
End Sub

Private Function MethodFromSetting() As Long
    Select Case LCase$(SettingText("merge_method"))
        Case "keep main only": MethodFromSetting = METHOD_KEEP
        Case "fill empty fields": MethodFromSetting = METHOD_FILL
        Case Else: MethodFromSetting = METHOD_COMBINE
    End Select
End Function

' Decides which rows go, collects the backup and log lines and works out the new values of the main records (memory only)
Private Sub PrepareMerge()
    Dim root As Long, i As Long, k As Long, s As Long, p As Long, lines As Long

    mMethod = MethodFromSetting()
    mAddColumns = SettingYes("add_missing")
    mStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    xCount = 0
    chN = 0
    mUnplaced = 0
    ReDim chRec(1 To 64)
    ReDim chCol(1 To 64)
    ReDim mRestored(1 To gFldCount)
    ReDim mRemove(1 To gN)
    LogClear

    For i = 1 To gN
        If gIsCopy(i) Then mRemove(i) = True
    Next i
    For root = 1 To gN
        If gRoot(root) = root Then
            If gStatus(root) = ST_AUTO Then
                p = gPrimary(root)
                For k = 1 To MemberCount(root)
                    s = MemberAt(root, k)
                    If s <> p Then mRemove(s) = True
                Next k
                If DistinctCount(root) > 1 Then lines = lines + 1
            End If
        End If
    Next root
    For i = 1 To gN
        If mRemove(i) Then lines = lines + 1
    Next i
    ReDim bk(1 To lines + 1, 1 To BK_LEAD + gLastCol)
    bkN = 1

    For root = 1 To gN
        If gRoot(root) = root Then
            If gStatus(root) = ST_AUTO Then
                p = gPrimary(root)
                If DistinctCount(root) > 1 Then BackupLine p, Txt(S_MRG_BK_MAIN), p
                For k = 1 To MemberCount(root)
                    s = MemberAt(root, k)
                    If s <> p Then RemoveLine s, p
                Next k
                MergeCluster root
            ElseIf MemberCount(root) > 1 Then
                For k = 1 To MemberCount(root)
                    s = MemberAt(root, k)
                    If gIsCopy(s) Then RemoveLine s, gRepOf(s)
                Next k
            End If
        End If
    Next root
End Sub

' ---------------------------------------------------------------- changing the sheets

Private Sub WriteBackup(ByVal ws As Worksheet)
    Dim bs As Worksheet
    Dim startRow As Long, c As Long, cols As Long, lastUsed As Long
    Dim rng As Range

    cols = BK_LEAD + gLastCol
    bk(1, 1) = Txt(S_MRG_BK_RUN)
    bk(1, 2) = Txt(S_MRG_BK_KEPT)
    bk(1, 3) = Txt(S_LOG_REASON)
    bk(1, 4) = Txt(S_MRG_COL_ROW)
    For c = 1 To gLastCol
        bk(1, BK_LEAD + c) = CStr(ws.Cells(HEADER_ROW, c).value)
    Next c

    Set bs = SheetByName("MergeBackup", True)
    lastUsed = bs.Cells(bs.rows.count, 1).End(xlUp).Row
    If lastUsed = 1 And Len(CStr(bs.Cells(1, 1).value)) = 0 Then
        startRow = 1
    Else
        startRow = lastUsed + 2
    End If
    Set rng = bs.Range(bs.Cells(startRow, 1), bs.Cells(startRow + bkN - 1, cols))
    rng.NumberFormat = "@"
    rng.value = bk
    bs.Range(bs.Cells(startRow, 1), bs.Cells(startRow, cols)).Font.Bold = True
    bs.Range("A:D").Columns.AutoFit
End Sub

Private Sub CreateColumns()
    Dim k As Long, idx As Long
    Dim restored As Boolean

    If xCount = 0 Then Exit Sub
    For k = 1 To xCount
        If xField(k) > 0 Then
            RestoreFieldColumn xKey(k)
            restored = True
        End If
    Next k
    If restored Then LoadFields
    For k = 1 To xCount
        If xField(k) = 0 Then AddFieldRow xKey(k), xProp(k), xParams(k), xFamily(k), xHdrEN(k), xHdrTR(k)
    Next k
    ReDim xCol(1 To xCount)
    For k = 1 To xCount
        idx = FieldIndexByKey(xKey(k))
        xCol(k) = gFldCol(idx)
    Next k
    RefreshAutoFilter
End Sub

Private Sub WriteCells(ByVal ws As Worksheet)
    Dim j As Long, i As Long, k As Long
    Dim cel As Range
    Dim v As Variant

    For j = 1 To chN
        i = chRec(j)
        Set cel = ws.Cells(gFirstRow + i - 1, chCol(j))
        v = gData(i, chCol(j))
        If VarType(v) = vbString Then
            If cel.NumberFormat = "General" Then cel.NumberFormat = "@"
        End If
        cel.value = v
    Next j
    For k = 1 To xCount
        For i = 1 To gN
            If Len(xVal(i, k)) > 0 Then ws.Cells(gFirstRow + i - 1, xCol(k)).value = xVal(i, k)
        Next i
    Next k
End Sub

' Deletes the rows from the bottom up, one block of neighbouring rows at a time
Private Sub DeleteRows(ByVal ws As Worksheet)
    Dim r As Long, r2 As Long, done As Long, total As Long
    Dim lastPct As Double, curPct As Double

    For r = 1 To gN
        If mRemove(r) Then total = total + 1
    Next r
    If total = 0 Then Exit Sub

    r = gN
    Do While r >= 1
        If mRemove(r) Then
            r2 = r
            Do While r2 > 1
                If Not mRemove(r2 - 1) Then Exit Do
                r2 = r2 - 1
            Loop
            ws.rows((gFirstRow + r2 - 1) & ":" & (gFirstRow + r - 1)).Delete
            done = done + (r - r2 + 1)
            curPct = 0.6 + 0.25 * done / total
            If curPct - lastPct >= 0.05 Or done = total Then
                ProgressUpdate curPct
                lastPct = curPct
            End If
            r = r2 - 1
        Else
            r = r - 1
        End If
    Loop
End Sub

' Merges the records of the same person (the groups on the Merge sheet that are not marked for review) using the
' method chosen on the Settings sheet. Deleted rows are copied to the MergeBackup sheet first.
Public Sub Merge_Apply()
    Dim ws As Worksheet, rep As Worksheet
    Dim plan As Long, reviewGroups As Long
    Dim groupsDone As Long, removedRows As Long, copyRows As Long, columnsAdded As Long, unplaced As Long
    Dim logWritten As Boolean
    Dim msg As String, errText As String, stage As String

    On Error GoTo ErrHandler

    ' Planning can take as long as applying the changes.  Start the bar before
    ' BuildPlan so the user gets feedback during duplicate detection as well.
    Set ws = DataSheet()
    ProgressStart
    StatusText Txt(S_ST_PARSING)
    ProgressUpdate 0.05
    stage = "plan"
    plan = BuildPlan()
    If plan = PLAN_FILTERED Then
        ProgressEnd
        MsgBox Txt(S_CLEAR_FILTER), vbExclamation
        Exit Sub
    End If
    If plan = PLAN_NO_DATA Then
        ProgressEnd
        MsgBox Txt(S_NO_DATA), vbInformation
        Exit Sub
    End If
    If gAutoGroups = 0 And gCopyRows = 0 Then
        ProgressEnd
        MsgBox Txt(S_MRG_NOTHING), vbInformation
        Exit Sub
    End If
    ProgressUpdate 1#
    ProgressEnd

    msg = Txt(S_MRG_APPLY_Q) & vbCrLf & vbCrLf & _
          gAutoGroups & " " & Txt(S_MRG_GROUPS_AUTO) & vbCrLf & _
          gRowsGone & " " & Txt(S_MRG_ROWS_GONE) & " (" & gCopyRows & " " & Txt(S_MRG_IDENT_OF) & ")" & vbCrLf & _
          Txt(S_SET_MMETHOD) & ": " & OptionLabel("merge_method") & vbCrLf & vbCrLf & _
          Txt(S_MRG_BACKUP_NOTE)
    If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub

    groupsDone = gAutoGroups
    removedRows = gRowsGone
    copyRows = gCopyRows

    ' The progress bar is drawn on Data; screen updating must stay on while it
    ' is being updated.
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    StatusText Txt(S_ST_WRITING)
    ProgressStart

    stage = "merge"
    PrepareMerge
    ProgressUpdate 0.2

    stage = "backup"
    WriteBackup ws
    ProgressUpdate 0.35

    stage = "columns"
    CreateColumns
    columnsAdded = xCount
    unplaced = mUnplaced
    ProgressUpdate 0.45

    stage = "cells"
    WriteCells ws
    ProgressUpdate 0.55

    stage = "rows"
    DeleteRows ws
    ProgressUpdate 0.85

    stage = "log"
    logWritten = LogFlush(Txt(S_OP_MERGE))
    ProgressUpdate 0.9

    stage = "report"
    If BuildPlan() = PLAN_OK Then reviewGroups = WriteReport()
    ProgressUpdate 1#
    ProgressEnd
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    msg = groupsDone & " " & Txt(S_MRG_DONE) & vbCrLf & _
          removedRows & " " & Txt(S_MRG_ROWS_REMOVED) & " (" & copyRows & " " & Txt(S_MRG_IDENT_OF) & ")"
    If columnsAdded > 0 Then msg = msg & vbCrLf & columnsAdded & " " & Txt(S_MRG_COLS_ADDED)
    If unplaced > 0 Then msg = msg & vbCrLf & unplaced & " " & Txt(S_MRG_UNPLACED)
    If reviewGroups > 0 Then msg = msg & vbCrLf & reviewGroups & " " & Txt(S_MRG_REVIEW_GROUPS)
    msg = msg & vbCrLf & vbCrLf & Txt(S_MRG_BACKUP_WRITTEN)
    If logWritten Then msg = msg & vbCrLf & Txt(S_LOG_WRITTEN)

    ' Creating or filling report/backup sheets can activate them.  Always
    ' return to Data after a merge, including when groups need review.
    ws.Activate
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    errText = Err.Description
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ProgressEnd
    On Error Resume Next
    DataSheet().Activate
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText & " [" & stage & "]", vbExclamation
End Sub
