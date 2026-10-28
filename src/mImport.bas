Attribute VB_Name = "mImport"
Option Explicit

' Import of vCard files into the Data sheet

' Maximum number of columns that one import may add
Private Const MAX_NEW_FIELDS As Long = 30
' Maximum number of skipped field names listed in the confirmation
Private Const MAX_LISTED_FIELDS As Long = 15
' Number of contact names shown in the confirmation
Private Const PREVIEW_COUNT As Long = 5

Private Const KIND_SINGLE As Long = 0
Private Const KIND_SLOT As Long = 1
Private Const KIND_JOIN As Long = 2

' Parsed entries: card number, import family and value
Private eCard() As Long
Private eFam() As String
Private eVal() As String
Private eCount As Long
Private cFn() As String
Private cCount As Long

' Values per contact and field; field numbers continue after the Fields sheet for columns that will be added
Private mOut() As Variant
Private nvKey() As String
Private nvProp() As String
Private nvParams() As String
Private nvFamily() As String
Private nvHdrEN() As String
Private nvHdrTR() As String
Private nvCount As Long
' Total number of field numbers in mOut: existing fields plus columns to be added
Private mFieldTotal As Long
' Fields whose Data column was deleted and will be recreated when the import is confirmed
Private mRestore() As Boolean
Private mRestoreCount As Long

' Family cache for slot assignment: name, field numbers (element 0 holds the count) and use counter per card
Private fcName() As String
Private fcSlots() As Variant
Private fcUsed() As Long
Private fcN As Long

' Phone formatting during import: settings, counters and a few before/after examples for the confirmation
Private mFormatPhones As Boolean
Private mCountryCode As String
Private mTrunk As String
Private mPhoneFormatted As Long
Private mPhoneKept As Long
Private mPhoneBefore(1 To 3) As String
Private mPhoneAfter(1 To 3) As String
Private mPhoneSamples As Long

Private mSkipKeys() As String
Private mSkipCounts() As Long
Private mSkipN As Long

' Log entries are kept until the contact names are known
Private lgCard() As Long
Private lgField() As String
Private lgReason() As String
Private lgValue() As String
Private lgN As Long

Public Sub Import_vCard_File()
    Dim ws As Worksheet
    Dim Target As Variant
    Dim text As String, encodingUsed As String
    Dim startRow As Long, written As Long, icon As Long
    Dim msg As String, errText As String

    On Error GoTo ErrHandler

    LoadFields
    Set ws = DataSheet()
    If Not ActiveSheet Is ws Then ws.Activate

    Target = Application.GetOpenFilename(FileFilter:=Txt(S_FILE_TYPE) & " (*.vcf),*.vcf", Title:=Txt(S_OPEN_PROMPT))
    If VarType(Target) = vbBoolean Then Exit Sub

    ResetState
    StatusText Txt(S_ST_READING)
    text = ReadTextFile(CStr(Target), SettingText("import_encoding"), AnsiCharset(), encodingUsed)

    StatusText Txt(S_ST_PARSING)
    ProgressStart
    ParseVCards text
    ProgressEnd

    If cCount = 0 Then
        MsgBox Txt(S_IMP_NO_CONTACTS), vbInformation
        Exit Sub
    End If
    ResolveEntries SettingYes("add_missing")

    ' The confirmation always shows the encoding and a preview so wrong characters are noticed before writing
    icon = vbQuestion
    If mSkipN > 0 Or nvCount > 0 Or mRestoreCount > 0 Or UsedAnsiFallback(encodingUsed) Then icon = vbExclamation
    If MsgBox(ImportSummary(encodingUsed), vbYesNo + icon + vbDefaultButton2) = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    StatusText Txt(S_ST_WRITING)
    written = WriteContacts(ws, startRow)
    FlushImportLog
    Application.ScreenUpdating = True
    Application.StatusBar = False

    msg = written & " " & Txt(S_IMP_DONE) & " " & startRow & "-" & (startRow + written - 1)
    If LogFlush(Txt(S_OP_IMPORT)) Then msg = msg & vbCrLf & vbCrLf & Txt(S_LOG_WRITTEN)
    MsgBox msg, vbInformation
    Exit Sub

ErrHandler:
    errText = Err.Description
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText, vbExclamation
End Sub

Private Sub ResetState()
    ReDim eCard(0 To 1023)
    ReDim eFam(0 To 1023)
    ReDim eVal(0 To 1023)
    eCount = 0
    ReDim cFn(0 To 255)
    cCount = 0
    nvCount = 0
    mRestoreCount = 0
    ReDim nvKey(1 To MAX_NEW_FIELDS)
    ReDim nvProp(1 To MAX_NEW_FIELDS)
    ReDim nvParams(1 To MAX_NEW_FIELDS)
    ReDim nvFamily(1 To MAX_NEW_FIELDS)
    ReDim nvHdrEN(1 To MAX_NEW_FIELDS)
    ReDim nvHdrTR(1 To MAX_NEW_FIELDS)
    fcN = 0
    ReDim fcName(1 To 64)
    ReDim fcSlots(1 To 64)
    ReDim fcUsed(1 To 64)
    ReDim mSkipKeys(0 To 31)
    ReDim mSkipCounts(0 To 31)
    mSkipN = 0
    ReDim lgCard(0 To 63)
    ReDim lgField(0 To 63)
    ReDim lgReason(0 To 63)
    ReDim lgValue(0 To 63)
    lgN = 0
    LogClear
    mFormatPhones = PhoneFormatEnabled("import")
    mCountryCode = SettingText("country_code")
    mTrunk = TrunkPrefix()
    mPhoneFormatted = 0
    mPhoneKept = 0
    mPhoneSamples = 0
End Sub

' Converts a phone number to the international format when phone formatting is on for imports;
' numbers that are not recognized stay as they are and are written to the log.
Private Function FormatPhone(ByVal v As String) As String
    Dim res As String
    FormatPhone = v
    If Not mFormatPhones Or Len(v) = 0 Then Exit Function
    If TryNormalizePhone(v, mCountryCode, mTrunk, res) Then
        If res <> v Then
            mPhoneFormatted = mPhoneFormatted + 1
            If mPhoneSamples < 3 Then
                mPhoneSamples = mPhoneSamples + 1
                mPhoneBefore(mPhoneSamples) = v
                mPhoneAfter(mPhoneSamples) = res
            End If
            FormatPhone = res
        End If
    Else
        mPhoneKept = mPhoneKept + 1
        LogLater cCount, "TEL", Txt(S_R_PHONE), v
    End If
End Function

' True when the file was not UTF-8 and the ANSI code page was applied automatically
Private Function UsedAnsiFallback(ByVal encodingUsed As String) As Boolean
    UsedAnsiFallback = (LCase$(SettingText("import_encoding")) = "auto") And (encodingUsed = AnsiCharset())
End Function

' Counts a vCard field that cannot be stored on the Data sheet
Private Sub AddSkipped(ByVal key As String)
    Dim k As Long
    For k = 0 To mSkipN - 1
        If mSkipKeys(k) = key Then
            mSkipCounts(k) = mSkipCounts(k) + 1
            Exit Sub
        End If
    Next k
    If mSkipN > UBound(mSkipKeys) Then
        ReDim Preserve mSkipKeys(0 To mSkipN * 2)
        ReDim Preserve mSkipCounts(0 To mSkipN * 2)
    End If
    mSkipKeys(mSkipN) = key
    mSkipCounts(mSkipN) = 1
    mSkipN = mSkipN + 1
End Sub

Private Sub LogLater(ByVal card As Long, ByVal fieldName As String, ByVal reason As String, ByVal value As String)
    If lgN > UBound(lgCard) Then
        ReDim Preserve lgCard(0 To lgN * 2)
        ReDim Preserve lgField(0 To lgN * 2)
        ReDim Preserve lgReason(0 To lgN * 2)
        ReDim Preserve lgValue(0 To lgN * 2)
    End If
    lgCard(lgN) = card
    lgField(lgN) = fieldName
    lgReason(lgN) = reason
    lgValue(lgN) = value
    lgN = lgN + 1
End Sub

Private Sub FlushImportLog()
    Dim k As Long
    For k = 0 To lgN - 1
        LogAdd "#" & lgCard(k) & " " & CardLabel(lgCard(k)), lgField(k), lgReason(k), lgValue(k)
    Next k
End Sub

Private Sub AddEntry(ByVal family As String, ByVal value As String)
    If eCount > UBound(eCard) Then
        ReDim Preserve eCard(0 To eCount * 2)
        ReDim Preserve eFam(0 To eCount * 2)
        ReDim Preserve eVal(0 To eCount * 2)
    End If
    eCard(eCount) = cCount
    eFam(eCount) = family
    eVal(eCount) = value
    eCount = eCount + 1
End Sub

' Splits the file into logical lines: folded lines (leading space or tab) and
' quoted-printable soft breaks (trailing "=") are joined with the previous line.
Private Function UnfoldLines(ByVal text As String, ByRef n As Long) As String()
    Dim phys() As String, res() As String
    Dim i As Long, p As String, first As String
    Dim joined As Boolean

    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    phys = Split(text, vbLf)
    ReDim res(0 To UBound(phys))
    n = 0
    For i = 0 To UBound(phys)
        p = phys(i)
        first = Left$(p, 1)
        joined = False
        If n > 0 Then
            If IsQpContinued(res(n - 1)) Then
                res(n - 1) = Left$(res(n - 1), Len(res(n - 1)) - 1) & p
                joined = True
            ElseIf first = " " Or first = vbTab Then
                ' Skip the payload of embedded binaries; appending it line by line is very slow
                If IsBinaryLine(res(n - 1)) Then
                    If Right$(res(n - 1), 1) = ":" Then res(n - 1) = res(n - 1) & "x"
                Else
                    res(n - 1) = res(n - 1) & Mid$(p, 2)
                End If
                joined = True
            End If
        End If
        If Not joined And Len(p) > 0 Then
            res(n) = p
            n = n + 1
        End If
    Next i
    UnfoldLines = res
End Function

Private Function IsQpContinued(ByVal s As String) As Boolean
    IsQpContinued = (Right$(s, 1) = "=") And (InStr(1, s, "QUOTED-PRINTABLE", vbTextCompare) > 0)
End Function

Private Function IsBinaryLine(ByVal s As String) As Boolean
    Dim head As String
    head = UCase$(Left$(s, 8))
    IsBinaryLine = (Left$(head, 5) = "PHOTO") Or (Left$(head, 4) = "LOGO") Or (Left$(head, 5) = "SOUND") Or (Left$(head, 3) = "KEY")
End Function

Private Function IsBinaryProperty(ByVal propName As String) As Boolean
    IsBinaryProperty = (propName = "PHOTO") Or (propName = "LOGO") Or (propName = "SOUND") Or (propName = "KEY")
End Function

' Parses all BEGIN:VCARD ... END:VCARD blocks into entries
Private Sub ParseVCards(ByVal text As String)
    Dim lines() As String
    Dim n As Long, i As Long, stepLines As Long, startEntries As Long, startLog As Long
    Dim inCard As Boolean

    lines = UnfoldLines(text, n)
    stepLines = n \ 100
    If stepLines < 1 Then stepLines = 1
    For i = 0 To n - 1
        Select Case UCase$(Trim$(lines(i)))
            Case "BEGIN:VCARD"
                cCount = cCount + 1
                If cCount > UBound(cFn) Then ReDim Preserve cFn(0 To cCount * 2)
                cFn(cCount) = vbNullString
                startEntries = eCount
                startLog = lgN
                inCard = True
            Case "END:VCARD"
                ' Cards that hold nothing storable (e.g. only a photo) are dropped
                If inCard Then
                    If eCount = startEntries And Len(cFn(cCount)) = 0 Then
                        cCount = cCount - 1
                        lgN = startLog
                    End If
                End If
                inCard = False
            Case Else
                If inCard Then ApplyProperty lines(i)
        End Select
        If i Mod stepLines = 0 Then ProgressUpdate i / n
    Next i
    ProgressUpdate 1
End Sub

' Reads one logical vCard line ("[group.]NAME;params:value") and records what it contains
Private Sub ApplyProperty(ByVal line As String)
    Dim head() As String, parts() As String
    Dim colonPos As Long, dot As Long, eq As Long, k As Long
    Dim propName As String, token As String, key As String, pv As String
    Dim types As String, charset As String, rawValue As String, v As String
    Dim isQp As Boolean
    Dim d As Date

    colonPos = FindUnquoted(line, ":")
    If colonPos < 2 Then Exit Sub
    rawValue = Mid$(line, colonPos + 1)
    head = SplitUnquoted(Left$(line, colonPos - 1), ";")

    propName = Trim$(head(0))
    dot = InStr(propName, ".")
    If dot > 0 Then propName = Mid$(propName, dot + 1)
    propName = UCase$(propName)

    ' Parameters: TYPE=a,b / TYPE=a;TYPE=b / bare "CELL" (vCard 2.1)
    For k = 1 To UBound(head)
        token = Trim$(head(k))
        eq = InStr(token, "=")
        If eq > 0 Then
            key = UCase$(Left$(token, eq - 1))
            pv = Replace(Mid$(token, eq + 1), """", vbNullString)
        Else
            key = "TYPE"
            pv = token
        End If
        If key = "TYPE" Then types = types & "," & UCase$(pv) & ","
        If key = "CHARSET" Then charset = pv
        If UCase$(pv) = "QUOTED-PRINTABLE" Then isQp = True
    Next k

    If isQp Then rawValue = DecodeQuotedPrintable(rawValue, charset)
    If Len(Trim$(rawValue)) = 0 Then Exit Sub

    Select Case propName
        Case "N"
            parts = SplitUnescaped(rawValue, ";")
            AddEntryText "N_FAMILY", PartAt(parts, 0)
            AddEntryText "N_GIVEN", PartAt(parts, 1)
            AddEntryText "N_MIDDLE", PartAt(parts, 2)
            AddEntryText "N_PREFIX", PartAt(parts, 3)
            AddEntryText "N_SUFFIX", PartAt(parts, 4)

        Case "FN"
            cFn(cCount) = Trim$(Unesc(rawValue))

        Case "BDAY"
            If TryParseBirthday(rawValue, d) Then
                AddEntry "BDAY", Format$(d, "yyyy-mm-dd")
            Else
                AddSkipped "BDAY (invalid or no year)"
                LogLater cCount, "BDAY", Txt(S_R_BAD_BIRTHDAY), Left$(rawValue, 200)
            End If

        Case "TEL"
            v = Trim$(Unesc(rawValue))
            If LCase$(Left$(v, 4)) = "tel:" Then v = Mid$(v, 5)
            AddEntryText TelFamily(types), FormatPhone(v)

        Case "EMAIL"
            If HasType(types, "WORK") Then
                AddEntryText "EMAIL|WORK", Trim$(Unesc(rawValue))
            ElseIf HasType(types, "OTHER") Then
                AddEntryText "EMAIL|OTHER", Trim$(Unesc(rawValue))
            Else
                AddEntryText "EMAIL|HOME", Trim$(Unesc(rawValue))
            End If

        Case "ADR"
            v = JoinParts(SplitUnescaped(rawValue, ";"), ", ")
            If HasType(types, "WORK") Then
                AddEntryText "ADR|WORK", v
            ElseIf HasType(types, "OTHER") Then
                AddEntryText "ADR|OTHER", v
            Else
                AddEntryText "ADR|HOME", v
            End If

        Case "ORG"
            AddEntryText "ORG", JoinParts(SplitUnescaped(rawValue, ";"), ", ")

        Case "TITLE", "ROLE", "NICKNAME", "NOTE"
            AddEntryText propName, Trim$(Unesc(rawValue))

        Case "URL"
            AddEntryText "URL", Trim$(Unesc(rawValue))

        Case "CATEGORIES"
            AddEntryText "CATEGORIES", JoinParts(SplitUnescaped(rawValue, ","), ", ")

        Case "VERSION", "PRODID", "REV", "UID", "CALSCALE", "SOURCE", _
             "X-ABLABEL", "X-ABADR", "X-ABSHOWAS", "X-ABUID"
            ' Metadata without contact data

        Case Else
            If IsBinaryProperty(propName) Then
                AddSkipped propName
                LogLater cCount, propName, Txt(S_R_BINARY), vbNullString
            Else
                AddEntryText "X|" & propName, Trim$(Unesc(rawValue))
            End If
    End Select
End Sub

Private Sub AddEntryText(ByVal family As String, ByVal value As String)
    If Len(value) > 0 Then AddEntry family, value
End Sub

' Import family of a phone number; untyped numbers follow the "Untyped phone numbers" setting
Private Function TelFamily(ByVal types As String) As String
    If HasType(types, "FAX") Then
        If HasType(types, "HOME") Then
            TelFamily = "TEL|HOMEFAX"
        Else
            TelFamily = "TEL|WORKFAX"
        End If
    ElseIf HasType(types, "CELL") Then
        TelFamily = "TEL|CELL"
    ElseIf HasType(types, "HOME") Then
        TelFamily = "TEL|HOME"
    ElseIf HasType(types, "WORK") Then
        TelFamily = "TEL|WORK"
    ElseIf HasType(types, "OTHER") Then
        TelFamily = "TEL|OTHER"
    Else
        Select Case LCase$(SettingText("untyped_phone"))
            Case "home": TelFamily = "TEL|HOME"
            Case "work": TelFamily = "TEL|WORK"
            Case Else: TelFamily = "TEL|CELL"
        End Select
    End If
End Function

Private Function FamilyKind(ByVal family As String) As Long
    If Left$(family, 4) = "TEL|" Or Left$(family, 6) = "EMAIL|" Or Left$(family, 4) = "ADR|" Or family = "URL" Then
        FamilyKind = KIND_SLOT
    ElseIf family = "NOTE" Or family = "CATEGORIES" Or Left$(family, 2) = "X|" Then
        FamilyKind = KIND_JOIN
    Else
        FamilyKind = KIND_SINGLE
    End If
End Function

Private Function JoinSeparator(ByVal family As String) As String
    Select Case family
        Case "NOTE": JoinSeparator = vbLf
        Case "CATEGORIES": JoinSeparator = ", "
        Case Else: JoinSeparator = "; "
    End Select
End Function

' Number of the cache entry for a family, built from the Fields sheet on first use
Private Function FamilyEntry(ByVal family As String) As Long
    Dim i As Long
    For i = 1 To fcN
        If fcName(i) = family Then
            FamilyEntry = i
            Exit Function
        End If
    Next i
    If fcN >= UBound(fcName) Then
        ReDim Preserve fcName(1 To fcN * 2)
        ReDim Preserve fcSlots(1 To fcN * 2)
        ReDim Preserve fcUsed(1 To fcN * 2)
    End If
    fcN = fcN + 1
    fcName(fcN) = family
    fcSlots(fcN) = FamilyIndexes(family)
    FamilyEntry = fcN
End Function

Private Function SlotCount(ByVal fe As Long) As Long
    Dim s() As Long
    s = fcSlots(fe)
    SlotCount = s(0)
End Function

Private Function SlotAt(ByVal fe As Long, ByVal pos As Long) As Long
    Dim s() As Long
    s = fcSlots(fe)
    SlotAt = s(pos)
End Function

Private Sub AppendSlot(ByVal fe As Long, ByVal fieldNo As Long)
    Dim s() As Long, n As Long
    s = fcSlots(fe)
    n = s(0) + 1
    ReDim Preserve s(0 To n)
    s(n) = fieldNo
    s(0) = n
    fcSlots(fe) = s
End Sub

' Property, parameters and headers of a field number (existing field or one that will be added)
Private Function FieldKeyOf(ByVal f As Long) As String
    If f <= gFldCount Then FieldKeyOf = gFldKey(f) Else FieldKeyOf = nvKey(f - gFldCount)
End Function

Private Function FieldPropOf(ByVal f As Long) As String
    If f <= gFldCount Then FieldPropOf = gFldProp(f) Else FieldPropOf = nvProp(f - gFldCount)
End Function

Private Function FieldParamsOf(ByVal f As Long) As String
    If f <= gFldCount Then FieldParamsOf = gFldParams(f) Else FieldParamsOf = nvParams(f - gFldCount)
End Function

Private Function FieldHdrENOf(ByVal f As Long) As String
    If f <= gFldCount Then FieldHdrENOf = gFldHdrEN(f) Else FieldHdrENOf = nvHdrEN(f - gFldCount)
End Function

Private Function FieldHdrTROf(ByVal f As Long) As String
    If f <= gFldCount Then FieldHdrTROf = gFldHdrTR(f) Else FieldHdrTROf = nvHdrTR(f - gFldCount)
End Function

Private Function KeyInUse(ByVal key As String) As Boolean
    Dim j As Long
    If FieldIndexByKey(key) > 0 Then
        KeyInUse = True
        Exit Function
    End If
    For j = 1 To nvCount
        If nvKey(j) = key Then
            KeyInUse = True
            Exit Function
        End If
    Next j
End Function

' Registers a column that will be created when the import is confirmed; returns its field number, 0 at the limit
Private Function NewField(ByVal key As String, ByVal prop As String, ByVal params As String, _
                          ByVal family As String, ByVal hdrEN As String, ByVal hdrTR As String) As Long
    If nvCount >= MAX_NEW_FIELDS Then Exit Function
    Do While KeyInUse(key)
        key = IncrementTrailing(key)
    Loop
    nvCount = nvCount + 1
    nvKey(nvCount) = key
    nvProp(nvCount) = prop
    nvParams(nvCount) = params
    nvFamily(nvCount) = family
    nvHdrEN(nvCount) = hdrEN
    nvHdrTR(nvCount) = hdrTR
    NewField = gFldCount + nvCount
End Function

' Field of a family that has a column now or will get one on import (deleted columns are recreated)
Private Function FieldIdx(ByVal family As String) As Long
    Dim f As Long
    For f = 1 To gFldCount
        If gFldFamily(f) = family Then
            If gFldCol(f) > 0 Then
                FieldIdx = f
                Exit Function
            End If
            If f <= UBound(mRestore) Then
                If mRestore(f) Then
                    FieldIdx = f
                    Exit Function
                End If
            End If
        End If
    Next f
End Function

' Next field of a family that has no column yet: first a field of the Fields table whose column was deleted
' (its column is recreated), otherwise a new column cloned from the last slot or built from the property name.
' Returns 0 when nothing can be added.
Private Function NextSlotField(ByVal fe As Long, ByVal fam As String) As Long
    Dim f As Long, slots As Long, last As Long, key As String, prop As String

    For f = 1 To gFldCount
        If gFldFamily(f) = fam And gFldCol(f) = 0 And Not mRestore(f) Then
            mRestore(f) = True
            mRestoreCount = mRestoreCount + 1
            AppendSlot fe, f
            NextSlotField = f
            Exit Function
        End If
    Next f

    slots = SlotCount(fe)
    If slots > 0 Then
        last = SlotAt(fe, slots)
        NextSlotField = NewField(IncrementTrailing(FieldKeyOf(last)), FieldPropOf(last), FieldParamsOf(last), fam, _
                                 IncrementTrailing(FieldHdrENOf(last)), IncrementTrailing(FieldHdrTROf(last)))
    ElseIf Left$(fam, 2) = "X|" Then
        prop = Mid$(fam, 3)
        key = SafeKey(prop)
        If Left$(key, 2) <> "X_" Then key = "X_" & key
        NextSlotField = NewField(key, prop, vbNullString, fam, prop, prop)
    End If
    If NextSlotField > 0 Then AppendSlot fe, NextSlotField
End Function

' Assigns the parsed entries to fields, in column order per family; extra values and unknown properties
' get new columns when "Add missing columns" is on, otherwise they are skipped and reported.
Private Sub ResolveEntries(ByVal addMissing As Boolean)
    Dim e As Long, c As Long, curCard As Long
    Dim fam As String, v As String, label As String, prop As String, key As String
    Dim fe As Long, slots As Long, pos As Long, Target As Long, kind As Long
    Dim gi As Long, mi As Long, fi As Long, oi As Long, ni As Long
    Dim composed As String, fnv As String, org As String

    ReDim mOut(1 To cCount, 1 To gFldCount + MAX_NEW_FIELDS)
    ReDim mRestore(1 To gFldCount)
    mFieldTotal = 0
    curCard = 0
    For e = 0 To eCount - 1
        If eCard(e) <> curCard Then
            curCard = eCard(e)
            For c = 1 To fcN
                fcUsed(c) = 0
            Next c
        End If
        fam = eFam(e)
        v = eVal(e)
        fe = FamilyEntry(fam)
        slots = SlotCount(fe)
        kind = FamilyKind(fam)
        Target = 0
        label = Replace(fam, "|", ";")

        ' Position of the slot this value needs: the next free one for slot families, the first one otherwise
        pos = 1
        If kind = KIND_SLOT Then pos = fcUsed(fe) + 1

        If pos <= slots Then
            Target = SlotAt(fe, pos)
            If kind = KIND_SINGLE Then
                If Not IsEmpty(mOut(eCard(e), Target)) Then
                    Target = 0
                    AddSkipped label & " (extra)"
                    LogLater eCard(e), label, Txt(S_R_NO_FREE), v
                End If
            End If
        ElseIf addMissing Then
            Target = NextSlotField(fe, fam)
            If Target = 0 Then
                If Left$(fam, 2) = "X|" Then label = Mid$(fam, 3)
                If nvCount >= MAX_NEW_FIELDS Then
                    AddSkipped label & " (limit)"
                    LogLater eCard(e), label, Txt(S_R_LIMIT), v
                Else
                    AddSkipped label
                    LogLater eCard(e), label, Txt(S_R_NO_COLUMN), v
                End If
            End If
        ElseIf slots > 0 Then
            AddSkipped label & " (extra)"
            LogLater eCard(e), label, Txt(S_R_NO_FREE), v
        Else
            If Left$(fam, 2) = "X|" Then label = Mid$(fam, 3)
            AddSkipped label
            LogLater eCard(e), label, Txt(S_R_NO_COLUMN), v
        End If
        If Target > 0 And kind = KIND_SLOT Then fcUsed(fe) = pos

        If Target > 0 Then
            If kind = KIND_JOIN And Not IsEmpty(mOut(eCard(e), Target)) Then
                mOut(eCard(e), Target) = mOut(eCard(e), Target) & JoinSeparator(fam) & v
            Else
                mOut(eCard(e), Target) = v
            End If
        End If
    Next e

    ' Without N, FN becomes the first name unless it only repeats the organization; a display name that
    ' differs from the name parts is kept in the FN column when there is one.
    mFieldTotal = gFldCount + nvCount
    gi = FieldIdx("N_GIVEN")
    mi = FieldIdx("N_MIDDLE")
    fi = FieldIdx("N_FAMILY")
    oi = FieldIdx("ORG")
    ni = FieldIdx("FN")
    For c = 1 To cCount
        composed = JoinWords(CellOf(c, gi), CellOf(c, mi), CellOf(c, fi))
        fnv = cFn(c)
        org = CellOf(c, oi)
        If Len(fnv) > 0 Then
            If Len(composed) = 0 Then
                If StrComp(fnv, org, vbTextCompare) <> 0 Then
                    If gi > 0 Then
                        mOut(c, gi) = fnv
                    ElseIf ni > 0 Then
                        mOut(c, ni) = fnv
                    End If
                End If
            ElseIf StrComp(fnv, composed, vbTextCompare) <> 0 Then
                If ni > 0 Then mOut(c, ni) = fnv
            End If
        End If
    Next c
End Sub

Private Function CellOf(ByVal card As Long, ByVal f As Long) As String
    If f = 0 Then Exit Function
    If IsEmpty(mOut(card, f)) Then Exit Function
    CellOf = CStr(mOut(card, f))
End Function

' Name of a parsed contact for the preview and log: full name, else organization, else the first value
Private Function CardLabel(ByVal card As Long) As String
    Dim s As String
    Dim f As Long, gi As Long, mi As Long, fi As Long, oi As Long

    If card < 1 Or card > cCount Then Exit Function
    gi = FieldIdx("N_GIVEN")
    mi = FieldIdx("N_MIDDLE")
    fi = FieldIdx("N_FAMILY")
    oi = FieldIdx("ORG")
    s = JoinWords(CellOf(card, gi), CellOf(card, mi), CellOf(card, fi))
    If Len(s) = 0 Then s = CellOf(card, oi)
    If Len(s) = 0 Then
        For f = 1 To mFieldTotal
            If Len(CellOf(card, f)) > 0 Then
                s = CellOf(card, f)
                Exit For
            End If
        Next f
    End If
    CardLabel = s
End Function

' Header of a field in the selected language
Private Function FieldHdrOfLang(ByVal f As Long) As String
    If LangNo() = 2 Then
        FieldHdrOfLang = FieldHdrTROf(f)
    Else
        FieldHdrOfLang = FieldHdrENOf(f)
    End If
End Function

' Appends "+ header" to the message unless the list limit is reached; returns the new number of listed lines
Private Function AddColumnLine(ByRef msg As String, ByVal shown As Long, ByVal header As String) As Long
    AddColumnLine = shown + 1
    If shown = MAX_LISTED_FIELDS Then
        msg = msg & vbCrLf & "..."
    ElseIf shown < MAX_LISTED_FIELDS Then
        msg = msg & vbCrLf & "+ " & header
    End If
End Function

Private Function ImportSummary(ByVal encodingUsed As String) As String
    Dim msg As String
    Dim k As Long, shown As Long

    msg = cCount & " " & Txt(S_IMP_FOUND) & vbCrLf & Txt(S_IMP_ENCODING) & ": " & UCase$(encodingUsed) & vbCrLf & vbCrLf
    msg = msg & Txt(S_IMP_PREVIEW) & ":" & vbCrLf
    shown = cCount
    If shown > PREVIEW_COUNT Then shown = PREVIEW_COUNT
    For k = 1 To shown
        msg = msg & "- " & CardLabel(k) & vbCrLf
    Next k

    If mFormatPhones And (mPhoneFormatted > 0 Or mPhoneKept > 0) Then
        msg = msg & vbCrLf & Txt(S_IMP_PHONES) & ": " & mPhoneFormatted & " " & Txt(S_IMP_PHONE_FMT) & ", " & mPhoneKept & " " & Txt(S_IMP_PHONE_KEPT) & vbCrLf
        For k = 1 To mPhoneSamples
            msg = msg & "  " & mPhoneBefore(k) & " -> " & mPhoneAfter(k) & vbCrLf
        Next k
    End If

    If UsedAnsiFallback(encodingUsed) Then msg = msg & vbCrLf & Txt(S_IMP_ANSI_HINT) & vbCrLf

    If nvCount > 0 Or mRestoreCount > 0 Then
        msg = msg & vbCrLf & Txt(S_IMP_NEWCOLS) & vbCrLf
        shown = 0
        For k = 1 To gFldCount
            If mRestore(k) Then shown = AddColumnLine(msg, shown, FieldHdrOfLang(k))
        Next k
        For k = 1 To nvCount
            shown = AddColumnLine(msg, shown, FieldHdrOfLang(gFldCount + k))
        Next k
        msg = msg & vbCrLf
    End If

    If mSkipN > 0 Then
        msg = msg & vbCrLf & Txt(S_IMP_SKIP_WARN) & vbCrLf
        For k = 0 To mSkipN - 1
            If k >= MAX_LISTED_FIELDS Then
                msg = msg & vbCrLf & "..."
                Exit For
            End If
            msg = msg & vbCrLf & "- " & mSkipKeys(k) & " (" & mSkipCounts(k) & ")"
        Next k
        msg = msg & vbCrLf
    End If
    ImportSummary = msg & vbCrLf & Txt(S_IMP_CONTINUE)
End Function

' Creates the planned columns, then appends all contacts below the last used row; returns the number written
Private Function WriteContacts(ByVal ws As Worksheet, ByRef startRow As Long) As Long
    Dim out As Variant
    Dim baseCount As Long, maxCol As Long, bdayCol As Long, lastRow As Long
    Dim bdayIdx As Long, j As Long
    Dim rng As Range

    baseCount = gFldCount
    If mRestoreCount > 0 Then
        For j = 1 To baseCount
            If mRestore(j) Then RestoreFieldColumn gFldKey(j)
        Next j
        LoadFields
    End If
    For j = 1 To nvCount
        AddFieldRow nvKey(j), nvProp(j), nvParams(j), nvFamily(j), nvHdrEN(j), nvHdrTR(j)
    Next j
    If nvCount > 0 Or mRestoreCount > 0 Then RefreshAutoFilter

    maxCol = LastFieldCol()
    bdayIdx = FieldIndexByFamily("BDAY")
    If bdayIdx > 0 Then bdayCol = gFldCol(bdayIdx)
    out = BuildBlock(maxCol, bdayIdx)

    lastRow = FindLastRow(ws, HEADER_ROW + 1, maxCol)
    If lastRow < HEADER_ROW + 1 Then
        startRow = HEADER_ROW + 1
    Else
        startRow = lastRow + 1
    End If

    ' Text format keeps leading zeros and stops text such as "=..." or "+90..." from being evaluated
    Set rng = ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + cCount - 1, maxCol))
    rng.NumberFormat = "@"
    If bdayCol > 0 Then rng.Columns(bdayCol).NumberFormat = "m/d/yyyy"
    rng.value = out
    If SettingYes("highlight") Then rng.Interior.Color = RGB(226, 239, 218)
    WriteContacts = cCount
End Function

' Contacts as a sheet block: every field value goes to the column of its field (birthdays become dates)
Private Function BuildBlock(ByVal maxCol As Long, ByVal bdayIdx As Long) As Variant
    Dim out() As Variant
    Dim c As Long, f As Long, col As Long
    Dim v As Variant, parts() As String

    ReDim out(1 To cCount, 1 To maxCol)
    For c = 1 To cCount
        For f = 1 To mFieldTotal
            col = gFldCol(f)
            If col > 0 Then
                v = mOut(c, f)
                If Not IsEmpty(v) Then
                    If f = bdayIdx Then
                        parts = Split(CStr(v), "-")
                        out(c, col) = DateSerial(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
                    Else
                        out(c, col) = v
                    End If
                End If
            End If
        Next f
    Next c
    BuildBlock = out
End Function
