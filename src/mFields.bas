Attribute VB_Name = "mFields"
Option Explicit

' Field table: every Data column is described by a row on the Fields sheet (key, vCard property,
' parameters, import family and headers). The column of a field is found through the hidden name
' f_<key> that points at its header cell, so columns can be moved, inserted or deleted freely.

Public gFldCount As Long
Public gFldKey() As String
Public gFldProp() As String
Public gFldParams() As String
Public gFldFamily() As String
Public gFldHdrEN() As String
Public gFldHdrTR() As String
Public gFldCol() As Long

Public Function FieldsSheet() As Worksheet
    Set FieldsSheet = ThisWorkbook.names("fields_keys").RefersToRange.Worksheet
End Function

' Reads the Fields sheet and resolves the Data column of each field (0 when the column was deleted)
Public Sub LoadFields()
    Dim ws As Worksheet, data As Variant
    Dim lastRow As Long, r As Long, n As Long

    Set ws = FieldsSheet()
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    gFldCount = 0
    If lastRow < 2 Then Exit Sub

    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 6)).value
    ReDim gFldKey(1 To lastRow - 1)
    ReDim gFldProp(1 To lastRow - 1)
    ReDim gFldParams(1 To lastRow - 1)
    ReDim gFldFamily(1 To lastRow - 1)
    ReDim gFldHdrEN(1 To lastRow - 1)
    ReDim gFldHdrTR(1 To lastRow - 1)
    ReDim gFldCol(1 To lastRow - 1)
    For r = 1 To lastRow - 1
        If Len(Trim$(CStr(data(r, 1)))) > 0 Then
            n = n + 1
            gFldKey(n) = Trim$(CStr(data(r, 1)))
            gFldProp(n) = UCase$(Trim$(CStr(data(r, 2))))
            gFldParams(n) = Trim$(CStr(data(r, 3)))
            gFldFamily(n) = Trim$(CStr(data(r, 4)))
            gFldHdrEN(n) = Trim$(CStr(data(r, 5)))
            gFldHdrTR(n) = Trim$(CStr(data(r, 6)))
            gFldCol(n) = ColumnOfKey(gFldKey(n))
        End If
    Next r
    gFldCount = n
End Sub

Private Function ColumnOfKey(ByVal key As String) As Long
    On Error Resume Next
    ColumnOfKey = ThisWorkbook.names("f_" & key).RefersToRange.Column
End Function

' Highest Data column used by a field
Public Function LastFieldCol() As Long
    Dim i As Long, res As Long
    For i = 1 To gFldCount
        If gFldCol(i) > res Then res = gFldCol(i)
    Next i
    LastFieldCol = res
End Function

' Index of the first field with the given import family that has a column, 0 if none
Public Function FieldIndexByFamily(ByVal family As String) As Long
    Dim i As Long
    For i = 1 To gFldCount
        If gFldFamily(i) = family And gFldCol(i) > 0 Then
            FieldIndexByFamily = i
            Exit Function
        End If
    Next i
End Function

' Indexes of all fields of a family that have a column, in table order (empty array if none)
Public Function FamilyIndexes(ByVal family As String) As Long()
    Dim res() As Long
    Dim i As Long, n As Long
    ReDim res(0 To gFldCount)
    For i = 1 To gFldCount
        If gFldFamily(i) = family And gFldCol(i) > 0 Then
            n = n + 1
            res(n) = i
        End If
    Next i
    res(0) = n
    FamilyIndexes = res
End Function

' Adds the trailing number of a key or header: "Mobile Phone3" -> "Mobile Phone4", "TEL_CELL_3" -> "TEL_CELL_4"
Public Function IncrementTrailing(ByVal s As String) As String
    Dim k As Long, digits As String
    k = Len(s)
    Do While k >= 1
        If Mid$(s, k, 1) < "0" Or Mid$(s, k, 1) > "9" Then Exit Do
        digits = Mid$(s, k, 1) & digits
        k = k - 1
    Loop
    If Len(digits) = 0 Then
        IncrementTrailing = s & "2"
    Else
        IncrementTrailing = Left$(s, k) & CStr(CLng(digits) + 1)
    End If
End Function

' Key for a custom vCard property: letters and digits only, other characters become underscores
Public Function SafeKey(ByVal prop As String) As String
    Dim k As Long, ch As String, res As String
    prop = UCase$(prop)
    For k = 1 To Len(prop)
        ch = Mid$(prop, k, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            res = res & ch
        Else
            res = res & "_"
        End If
    Next k
    SafeKey = res
End Function

' Appends a field to the Fields sheet and creates its Data column; returns the field index
Public Function AddFieldRow(ByVal key As String, ByVal prop As String, ByVal params As String, _
                            ByVal family As String, ByVal hdrEN As String, ByVal hdrTR As String) As Long
    Dim ws As Worksheet, r As Long

    Set ws = FieldsSheet()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).NumberFormat = "@"
    ws.Cells(r, 1).value = key
    ws.Cells(r, 2).value = prop
    ws.Cells(r, 3).value = params
    ws.Cells(r, 4).value = family
    ws.Cells(r, 5).value = hdrEN
    ws.Cells(r, 6).value = hdrTR
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Borders.LineStyle = xlContinuous

    CreateDataColumn key
    LoadFields
    AddFieldRow = FieldIndexByKey(key)
End Function

Public Function FieldIndexByKey(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To gFldCount
        If gFldKey(i) = key Then
            FieldIndexByKey = i
            Exit Function
        End If
    Next i
End Function

' Recreates the Data column of a field whose column was deleted (its f_<key> name is broken)
Public Sub RestoreFieldColumn(ByVal key As String)
    Dim idx As Long
    idx = FieldIndexByKey(key)
    ' gFldCol is updated at once: the next restored column is placed after this one
    If idx > 0 Then gFldCol(idx) = CreateDataColumn(key)
End Sub

' Creates the header cell (formula from the Fields sheet), the f_<key> name and the column formatting; returns the column
Private Function CreateDataColumn(ByVal key As String) As Long
    Dim ws As Worksheet, newCol As Long, srcCol As Long

    Set ws = DataSheet()
    srcCol = LastFieldCol()
    newCol = srcCol + 1

    ' The header format is copied from the last field column; with no column left the header stays plain
    If srcCol > 0 Then
        ws.Cells(HEADER_ROW, srcCol).Copy
        ws.Cells(HEADER_ROW, newCol).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If

    ws.Cells(HEADER_ROW, newCol).Formula = "=INDEX(fields_hdr,MATCH(""" & key & """,fields_keys,0),lang_no)"
    ThisWorkbook.names.Add Name:="f_" & key, RefersTo:="='" & ws.Name & "'!" & ws.Cells(HEADER_ROW, newCol).Address, Visible:=False
    ws.Columns(newCol).ColumnWidth = 14
    ws.Columns(newCol).NumberFormat = "@"
    ws.Cells(HEADER_ROW, newCol).NumberFormat = "General"
    CreateDataColumn = newCol
End Function

' Re-applies the AutoFilter so it covers the header cells of newly added columns
Public Sub RefreshAutoFilter()
    Dim ws As Worksheet
    Set ws = DataSheet()
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, LastFieldCol())).AutoFilter
End Sub
