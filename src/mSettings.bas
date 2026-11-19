Attribute VB_Name = "mSettings"
' Mesut Akcan
' https://github.com/mesutakcan/Excel-VCF-Converter
' youtube.com/mesutakcan
Option Explicit

' Row of the Data sheet that holds the column headers; contacts start on the next row
Public Const HEADER_ROW As Long = 2

' Row indexes in the lang_data table (Strings sheet)
Public Const S_TITLE As Long = 26
Public Const S_BTN_EXPORT As Long = 27
Public Const S_NO_DATA As Long = 28
Public Const S_FILE_TYPE As Long = 29
Public Const S_SAVE_PROMPT As Long = 30
Public Const S_FILE_EXISTS As Long = 31
Public Const S_OVERWRITE As Long = 32
Public Const S_EXPORTED_TO As Long = 33
Public Const S_NO_NAME As Long = 34
Public Const S_ERROR As Long = 35
Public Const S_ROW_NUMBERS As Long = 36
Public Const S_BAD_BIRTHDAY As Long = 37
Public Const S_FILTERED As Long = 38
Public Const S_BTN_IMPORT As Long = 39
Public Const S_OPEN_PROMPT As Long = 40
Public Const S_IMP_NO_CONTACTS As Long = 41
Public Const S_IMP_SKIP_WARN As Long = 42
Public Const S_IMP_CONTINUE As Long = 43
Public Const S_IMP_DONE As Long = 44
Public Const S_IMP_ENCODING As Long = 45
Public Const S_IMP_ANSI_HINT As Long = 46
Public Const S_IMP_PREVIEW As Long = 47
Public Const S_IMP_FOUND As Long = 48
Public Const S_BTN_DEDUP As Long = 49
Public Const S_BTN_CLEAR As Long = 50
Public Const S_DEDUP_CONFIRM As Long = 80
Public Const S_DEDUP_DONE As Long = 81
Public Const S_DEDUP_NONE As Long = 82
Public Const S_CLEAR_FILTER As Long = 83
Public Const S_CLEAR_CONFIRM As Long = 84
Public Const S_CLEAR_DONE As Long = 85
Public Const S_NO_SELECTION As Long = 86
Public Const S_IMP_NEWCOLS As Long = 87
Public Const S_LOG_OPERATION As Long = 88
Public Const S_LOG_CONTACT As Long = 89
Public Const S_LOG_FIELD As Long = 90
Public Const S_LOG_REASON As Long = 91
Public Const S_LOG_VALUE As Long = 92
Public Const S_R_NO_NAME As Long = 93
Public Const S_R_BAD_BIRTHDAY As Long = 94
Public Const S_R_NO_COLUMN As Long = 95
Public Const S_R_NO_FREE As Long = 96
Public Const S_OP_IMPORT As Long = 97
Public Const S_OP_EXPORT As Long = 98
Public Const S_LOG_WRITTEN As Long = 99
Public Const S_ST_READING As Long = 100
Public Const S_ST_PARSING As Long = 101
Public Const S_ST_WRITING As Long = 102
Public Const S_R_LIMIT As Long = 103
Public Const S_R_BINARY As Long = 107
Public Const S_BTN_SETTINGS As Long = 51
Public Const S_IMP_PHONES As Long = 136
Public Const S_IMP_PHONE_FMT As Long = 137
Public Const S_IMP_PHONE_KEPT As Long = 138
Public Const S_R_PHONE As Long = 139
Public Const S_BTN_MERGE As Long = 145
Public Const S_MRG_COL_GROUP As Long = 146
Public Const S_MRG_COL_STATUS As Long = 147
Public Const S_MRG_COL_ROW As Long = 148
Public Const S_MRG_COL_NAME As Long = 149
Public Const S_MRG_COL_PHONES As Long = 150
Public Const S_MRG_COL_GROUPS As Long = 151
Public Const S_MRG_COL_ACTION As Long = 152
Public Const S_MRG_AUTO As Long = 153
Public Const S_MRG_REVIEW As Long = 154
Public Const S_MRG_KEEP As Long = 155
Public Const S_MRG_REMOVE As Long = 156
Public Const S_MRG_UNCHANGED As Long = 157
Public Const S_MRG_RECORDS As Long = 158
Public Const S_MRG_GROUPS_AUTO As Long = 159
Public Const S_MRG_ROWS_GONE As Long = 160
Public Const S_MRG_REVIEW_GROUPS As Long = 161
Public Const S_MRG_SEE_SHEET As Long = 162
Public Const S_MRG_ANALYSIS_ONLY As Long = 163
Public Const S_MRG_NOTHING As Long = 164
Public Const S_BTN_MERGE_APPLY As Long = 171
Public Const S_BTN_ABOUT As Long = 108

Public Const S_SET_MMETHOD As Long = 172
Public Const S_SET_MMETHOD_DESC As Long = 173
Public Const S_MRG_COL_COPIES As Long = 174
Public Const S_MRG_APPLY_Q As Long = 175
Public Const S_MRG_IDENT_OF As Long = 176
Public Const S_MRG_BACKUP_NOTE As Long = 177
Public Const S_MRG_DONE As Long = 178
Public Const S_MRG_ROWS_REMOVED As Long = 179
Public Const S_MRG_BACKUP_WRITTEN As Long = 180
Public Const S_MRG_COLS_ADDED As Long = 181
Public Const S_MRG_UNPLACED As Long = 182
Public Const S_MRG_BK_RUN As Long = 183
Public Const S_MRG_BK_KEPT As Long = 184
Public Const S_MRG_BK_MAIN As Long = 185
Public Const S_MRG_BK_MERGED As Long = 186
Public Const S_MRG_BK_COPY As Long = 187
Public Const S_MRG_LOG_INTO As Long = 188
Public Const S_MRG_LOG_COPY As Long = 189
Public Const S_MRG_LOG_DIFF As Long = 190
Public Const S_OP_MERGE As Long = 191

' Returns the text at row idx of the lang_data table in the selected language
Public Function Txt(ByVal idx As Long) As String
    Dim tbl As Range, lang As Long
    Set tbl = ThisWorkbook.names("lang_data").RefersToRange
    lang = LangNo()
    If lang > tbl.Columns.count Then lang = 1
    Txt = CStr(tbl.Cells(idx, lang).value)
End Function

' Column of lang_data for the language chosen on the Settings sheet (the lang_no name points at a helper cell there)
Public Function LangNo() As Long
    Dim v As Variant
    On Error Resume Next
    v = ThisWorkbook.names("lang_no").RefersToRange.value
    On Error GoTo 0
    LangNo = 1
    If IsNumeric(v) Then
        If CLng(v) >= 1 Then LangNo = CLng(v)
    End If
End Function

' The Data sheet is found through the data_head name, so it does not depend on the tab name or code name
Public Function DataSheet() As Worksheet
    Set DataSheet = ThisWorkbook.names("data_head").RefersToRange.Worksheet
End Function

' Input cell of a setting (the hidden name set_<key> points at it); Nothing when the name is missing
Private Function SettingCell(ByVal key As String) As Range
    On Error Resume Next
    Set SettingCell = ThisWorkbook.names("set_" & key).RefersToRange
End Function

' Name of the choice list behind a setting ("" for free text); the list table is the name opt_<list>_tbl
Private Function ListName(ByVal key As String) As String
    Select Case key
        Case "lang": ListName = "lang"
        Case "import_encoding": ListName = "enc"
        Case "ansi_codepage": ListName = "cp"
        Case "add_missing", "export_bom", "fold_lines", "highlight", "write_log": ListName = "yesno"
        Case "fn_format": ListName = "fn"
        Case "untyped_phone": ListName = "untyped"
        Case "phone_format": ListName = "phone"
        Case "export_scope": ListName = "scope"
        Case "dup_match": ListName = "dup"
        Case "merge_level": ListName = "merge"
        Case "merge_method": ListName = "mmethod"
    End Select
End Function

' The drop-downs show the choices in the selected language and the cell keeps that text; this maps the text of
' either language back to the English choice that the code works with. Unknown text is returned unchanged.
Public Function CanonicalFromTable(ByRef tbl As Variant, ByVal text As String) As String
    Dim r As Long
    For r = LBound(tbl, 1) To UBound(tbl, 1)
        If StrComp(CStr(tbl(r, 1)), text, vbTextCompare) = 0 Or StrComp(CStr(tbl(r, 2)), text, vbTextCompare) = 0 Then
            CanonicalFromTable = CStr(tbl(r, 1))
            Exit Function
        End If
    Next r
    CanonicalFromTable = text
End Function

Private Function CanonicalOption(ByVal list As String, ByVal text As String) As String
    If list = "merge" Then
        If StrComp(text, "Özdeþ", vbTextCompare) = 0 Or StrComp(text, "Ozdes", vbTextCompare) = 0 Or StrComp(text, "Exact", vbTextCompare) = 0 Then
            CanonicalOption = "Exact"
            Exit Function
        ElseIf StrComp(text, "Kapsamlý", vbTextCompare) = 0 Or StrComp(text, "Kapsamli", vbTextCompare) = 0 Or StrComp(text, "Extended", vbTextCompare) = 0 Then
            CanonicalOption = "Extended"
            Exit Function
        End If
    End If
    Dim tbl As Variant
    On Error Resume Next
    tbl = ThisWorkbook.names("opt_" & list & "_tbl").RefersToRange.value
    On Error GoTo 0
    If IsArray(tbl) Then
        CanonicalOption = CanonicalFromTable(tbl, text)
    Else
        CanonicalOption = text
    End If
End Function

' Text of the choice (English value) in the given language column of the list table; unknown text is returned unchanged
Public Function LocalizedFromTable(ByRef tbl As Variant, ByVal canonical As String, ByVal lang As Long) As String
    Dim r As Long, col As Long
    col = lang
    If col < 1 Or col > UBound(tbl, 2) Then col = 1
    For r = LBound(tbl, 1) To UBound(tbl, 1)
        If StrComp(CStr(tbl(r, 1)), canonical, vbTextCompare) = 0 Then
            LocalizedFromTable = CStr(tbl(r, col))
            Exit Function
        End If
    Next r
    LocalizedFromTable = canonical
End Function

Private Function LocalizedOption(ByVal list As String, ByVal canonical As String, ByVal lang As Long) As String
    If list = "merge" Then
        If StrComp(canonical, "Exact", vbTextCompare) = 0 Then
            If lang = 2 Then
                LocalizedOption = "Özdeþ"
            Else
                LocalizedOption = "Exact"
            End If
            Exit Function
        ElseIf StrComp(canonical, "Extended", vbTextCompare) = 0 Then
            If lang = 2 Then
                LocalizedOption = "Kapsamlý"
            Else
                LocalizedOption = "Extended"
            End If
            Exit Function
        End If
    End If
    Dim tbl As Variant
    On Error Resume Next
    tbl = ThisWorkbook.names("opt_" & list & "_tbl").RefersToRange.value
    On Error GoTo 0
    If IsArray(tbl) Then
        LocalizedOption = LocalizedFromTable(tbl, canonical, lang)
    Else
        LocalizedOption = canonical
    End If
End Function

' Choice of a list setting in the selected language, whatever language the cell currently holds
Public Function OptionLabel(ByVal key As String) As String
    OptionLabel = LocalizedOption(ListName(key), SettingText(key), LangNo())
End Function

' Value of a setting as the code uses it (English choice); empty or missing cells use the default
Public Function SettingText(ByVal key As String) As String
    Dim c As Range, res As String, list As String

    Set c = SettingCell(key)
    If Not c Is Nothing Then
        If Not IsError(c.value) Then res = Trim$(CStr(c.value))
    End If
    list = ListName(key)
    If Len(res) > 0 And Len(list) > 0 Then res = CanonicalOption(list, res)
    If Len(res) = 0 Then res = DefaultSetting(key)
    SettingText = res
End Function

' Text of a setting as shown on the Settings sheet (selected language)
Public Function SettingDisplay(ByVal key As String) As String
    Dim c As Range, res As String
    Set c = SettingCell(key)
    If Not c Is Nothing Then
        If Not IsError(c.value) Then res = Trim$(CStr(c.value))
    End If
    If Len(res) = 0 Then res = DefaultSetting(key)
    SettingDisplay = res
End Function

' The national prefix may legitimately be empty, so only a missing setting cell falls back to the default
Public Function TrunkPrefix() As String
    Dim c As Range
    Set c = SettingCell("trunk_prefix")
    If c Is Nothing Then
        TrunkPrefix = DefaultSetting("trunk_prefix")
    ElseIf IsError(c.value) Then
        TrunkPrefix = DefaultSetting("trunk_prefix")
    Else
        TrunkPrefix = Trim$(CStr(c.value))
    End If
End Function

Public Function SettingYes(ByVal key As String) As Boolean
    SettingYes = (LCase$(SettingText(key)) = "yes")
End Function

' True when phone numbers are formatted in the given phase ("import" or "export")
Public Function PhoneFormatEnabled(ByVal phase As String) As Boolean
    Select Case LCase$(SettingText("phone_format"))
        Case "both": PhoneFormatEnabled = True
        Case "on import": PhoneFormatEnabled = (phase = "import")
        Case "on export": PhoneFormatEnabled = (phase = "export")
    End Select
End Function

Private Function DefaultSetting(ByVal key As String) As String
    Select Case key
        Case "import_encoding": DefaultSetting = "Auto"
        Case "ansi_codepage": DefaultSetting = "windows-1254 (Turkish)"
        Case "add_missing": DefaultSetting = "Yes"
        Case "export_bom": DefaultSetting = "No"
        Case "fold_lines": DefaultSetting = "No"
        Case "fn_format": DefaultSetting = "First Middle Last"
        Case "untyped_phone": DefaultSetting = "Mobile"
        Case "phone_format": DefaultSetting = "Off"
        Case "country_code": DefaultSetting = "90"
        Case "trunk_prefix": DefaultSetting = "0"
        Case "export_scope": DefaultSetting = "All rows"
        Case "highlight": DefaultSetting = "Yes"
        Case "write_log": DefaultSetting = "Yes"
        Case "dup_match": DefaultSetting = "All columns"
        Case "merge_level": DefaultSetting = "Normal"
        Case "merge_method": DefaultSetting = "Combine all"
    End Select
End Function

' Charset name of the ANSI code page setting, e.g. "windows-1254 (Turkish)" -> "windows-1254"
Public Function AnsiCharset() As String
    AnsiCharset = Split(SettingText("ansi_codepage"), " ")(0)
End Function

' Sets the button captions for the selected language (called when the workbook opens)
Public Sub ApplyLanguage()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = DataSheet()
    ws.Buttons("button1").text = Txt(S_BTN_IMPORT)
    ws.Buttons("button2").text = Txt(S_BTN_EXPORT)
    ws.Buttons("button3").text = Txt(S_BTN_DEDUP)
    ws.Buttons("button4").text = Txt(S_BTN_CLEAR)
    ws.Buttons("button5").text = Txt(S_BTN_SETTINGS)
    ws.Buttons("button6").text = Txt(S_BTN_MERGE)
    ws.Buttons("button7").text = Txt(S_BTN_MERGE_APPLY)
    ws.Buttons("button8").text = Txt(S_BTN_ABOUT)
End Sub

' Called from ThisWorkbook when a cell changes. When the language setting changes, the choices held by the other
' list settings are translated to the new language and the button captions are refreshed.
Public Sub OnSettingsChanged(ByVal Target As Range)
    Dim hit As Range, c As Range
    Dim keys As Variant
    Dim k As Long, list As String, canonical As String

    On Error Resume Next
    Set hit = Intersect(Target, SettingCell("lang"))
    On Error GoTo 0
    If hit Is Nothing Then Exit Sub

    ' The language number is a formula on the sheet; make sure it is current before reading it
    Application.Calculate
    Application.EnableEvents = False
    On Error GoTo done
    keys = Array("import_encoding", "ansi_codepage", "add_missing", "export_bom", "fold_lines", "fn_format", _
                 "untyped_phone", "phone_format", "export_scope", "highlight", "write_log", "dup_match", "merge_level", "merge_method")
    For k = 0 To UBound(keys)
        list = ListName(CStr(keys(k)))
        Set c = SettingCell(CStr(keys(k)))
        If Not c Is Nothing Then
            canonical = CanonicalOption(list, Trim$(CStr(c.value)))
            If Len(canonical) > 0 Then c.value = LocalizedOption(list, canonical, LangNo())
        End If
    Next k
    ApplyLanguage
done:
    Application.EnableEvents = True
End Sub

' Goes to the Settings sheet
Public Sub Show_Settings()
    Dim c As Range
    Set c = SettingCell("lang")
    If c Is Nothing Then Exit Sub
    c.Worksheet.Activate
    c.Select
End Sub

Sub button8_Click()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("About")
    If Not ws Is Nothing Then
        Application.GoTo ws.Range("C3")
    End If
End Sub

' Settings sayfasýndaki Benzerlik Düzeyi ve Birleþtirme Yöntemi tablolarýný ve adlandýrýlmýþ aralýklarýný yapýlandýrýr
Public Sub Setup_Merge_Options()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Settings")
    If ws Is Nothing Then Exit Sub
    
    ' Benzerlik düzeyi tablosu (H30:I33)
    ws.Cells(30, 8).value = "Exact"
    ws.Cells(30, 9).value = "Özdeþ"
    ws.Cells(31, 8).value = "Safe"
    ws.Cells(31, 9).value = "Güvenli"
    ws.Cells(32, 8).value = "Normal"
    ws.Cells(32, 9).value = "Normal"
    ws.Cells(33, 8).value = "Extended"
    ws.Cells(33, 9).value = "Kapsamlý"
    
    ' Birleþtirme yöntemi tablosu (H34:I36)
    ws.Cells(34, 8).value = "Keep main only"
    ws.Cells(34, 9).value = "Yalnýzca asýl kayýt"
    ws.Cells(35, 8).value = "Fill empty fields"
    ws.Cells(35, 9).value = "Boþlarý doldur"
    ws.Cells(36, 8).value = "Combine all"
    ws.Cells(36, 9).value = "Hepsini birleþtir"
    
    ' Tanýmlý adlarý güncelle
    ThisWorkbook.names.Add Name:="opt_merge_tbl", RefersTo:="=Settings!$H$30:$I$33"
    ThisWorkbook.names.Add Name:="opt_mmethod_tbl", RefersTo:="=Settings!$H$34:$I$36"
    
    ' Açýklama hücresini güncelle (C18)
    ws.Range("C18").value = "Özdeþ: yalnýzca birebir ayný ad ve ortak telefon/e-posta. Güvenli: ayný veya birbirini kapsayan adlar. Normal: ayrýca baþ harf veya tek harf yazým farký. Kapsamlý: sýralý sözcük eþleþmesi ve geniþletilmiþ baðlar."
    
    MsgBox "Ayarlar tablosu güncellendi. 'Özdeþ' (Exact) ve 'Kapsamlý' (Extended) seçenekleri eklendi.", vbInformation
End Sub

Public Sub Setup_Extended_Option()
    Setup_Merge_Options
End Sub
