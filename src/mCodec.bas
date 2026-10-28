Attribute VB_Name = "mCodec"
' Mesut Akcan
' https://github.com/mesutakcan/Excel-VCF-Converter
' youtube.com/mesutakcan

Option Explicit

' Text helpers shared by export and import: vCard escaping, parsing helpers, encodings and file I/O

' Cell value as trimmed text; error and empty cells give "", non-breaking spaces count as spaces
Public Function CellText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    s = Replace(CStr(v), ChrW$(160), " ")
    CellText = Trim$(s)
End Function

Public Function JoinWords(ByVal a As String, ByVal b As String, ByVal c As String) As String
    Dim res As String
    If Len(a) > 0 Then res = a
    If Len(b) > 0 Then
        If Len(res) > 0 Then res = res & " "
        res = res & b
    End If
    If Len(c) > 0 Then
        If Len(res) > 0 Then res = res & " "
        res = res & c
    End If
    JoinWords = res
End Function

' Escapes characters that are special in vCard text values
Public Function Esc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, ";", "\;")
    s = Replace(s, ",", "\,")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    Esc = s
End Function

' Escapes each category but keeps the comma as the list separator
Public Function EscCategories(ByVal s As String) As String
    Dim parts() As String, item As String, res As String
    Dim k As Long
    parts = Split(s, ",")
    For k = LBound(parts) To UBound(parts)
        item = Trim$(parts(k))
        If Len(item) > 0 Then
            If Len(res) > 0 Then res = res & ","
            res = res & Esc(item)
        End If
    Next k
    EscCategories = res
End Function

' Reverses Esc: \n becomes a line break, any other escaped character is kept literally
Public Function Unesc(ByVal s As String) As String
    Dim res As String, ch As String, nxt As String
    Dim k As Long

    If InStr(s, "\") = 0 Then
        Unesc = s
        Exit Function
    End If
    k = 1
    Do While k <= Len(s)
        ch = Mid$(s, k, 1)
        If ch = "\" And k < Len(s) Then
            nxt = Mid$(s, k + 1, 1)
            If nxt = "n" Or nxt = "N" Then
                res = res & vbLf
            Else
                res = res & nxt
            End If
            k = k + 2
        Else
            res = res & ch
            k = k + 1
        End If
    Loop
    Unesc = res
End Function

' Splits on sep, ignoring separators escaped with a backslash; parts stay escaped
Public Function SplitUnescaped(ByVal s As String, ByVal sep As String) As String()
    Dim res() As String
    Dim n As Long, k As Long
    Dim ch As String, cur As String

    ReDim res(0 To Len(s))
    k = 1
    Do While k <= Len(s)
        ch = Mid$(s, k, 1)
        If ch = "\" And k < Len(s) Then
            cur = cur & ch & Mid$(s, k + 1, 1)
            k = k + 2
        ElseIf ch = sep Then
            res(n) = cur
            n = n + 1
            cur = vbNullString
            k = k + 1
        Else
            cur = cur & ch
            k = k + 1
        End If
    Loop
    res(n) = cur
    ReDim Preserve res(0 To n)
    SplitUnescaped = res
End Function

' Splits on sep outside double quotes (parameter lists such as TYPE="a,b")
Public Function SplitUnquoted(ByVal s As String, ByVal sep As String) As String()
    Dim res() As String
    Dim n As Long, k As Long
    Dim ch As String, cur As String
    Dim inQuote As Boolean

    ReDim res(0 To Len(s))
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If ch = """" Then inQuote = Not inQuote
        If ch = sep And Not inQuote Then
            res(n) = cur
            n = n + 1
            cur = vbNullString
        Else
            cur = cur & ch
        End If
    Next k
    res(n) = cur
    ReDim Preserve res(0 To n)
    SplitUnquoted = res
End Function

' Position of the first ch outside double quotes, 0 if none
Public Function FindUnquoted(ByVal s As String, ByVal ch As String) As Long
    Dim k As Long, c As String
    Dim inQuote As Boolean
    For k = 1 To Len(s)
        c = Mid$(s, k, 1)
        If c = """" Then
            inQuote = Not inQuote
        ElseIf c = ch And Not inQuote Then
            FindUnquoted = k
            Exit Function
        End If
    Next k
End Function

' Unescaped, trimmed part idx of a split value; "" when the part does not exist
Public Function PartAt(ByRef parts() As String, ByVal idx As Long) As String
    If idx <= UBound(parts) Then PartAt = Trim$(Unesc(parts(idx)))
End Function

Public Function JoinParts(ByRef parts() As String, ByVal sep As String) As String
    Dim res As String, token As String
    Dim k As Long
    For k = 0 To UBound(parts)
        token = PartAt(parts, k)
        If Len(token) > 0 Then
            If Len(res) > 0 Then res = res & sep
            res = res & token
        End If
    Next k
    JoinParts = res
End Function

' True when the parameter string (",CELL,VOICE,") contains the type t
Public Function HasType(ByVal types As String, ByVal t As String) As Boolean
    HasType = InStr(types, "," & t & ",") > 0
End Function

Public Function IsHex(ByVal s As String) As Boolean
    Dim k As Long
    For k = 1 To Len(s)
        If InStr("0123456789ABCDEFabcdef", Mid$(s, k, 1)) = 0 Then Exit Function
    Next k
    IsHex = True
End Function

Public Function IsDigits(ByVal s As String) As Boolean
    Dim k As Long, c As String
    For k = 1 To Len(s)
        c = Mid$(s, k, 1)
        If c < "0" Or c > "9" Then Exit Function
    Next k
    IsDigits = Len(s) > 0
End Function

' Accepts real dates and date-like text; rejects times and years before 1900
Public Function TryFormatBirthday(ByVal v As Variant, ByRef result As String) As Boolean
    Dim d As Date
    Select Case VarType(v)
        Case vbDate
            d = v
        Case vbString
            If Not IsDate(v) Then Exit Function
            d = CDate(v)
        Case Else
            Exit Function
    End Select
    If Year(d) < 1900 Then Exit Function
    result = Format$(d, "yyyy-mm-dd")
    TryFormatBirthday = True
End Function

' Accepts yyyy-mm-dd, yyyymmdd and date-time forms; rejects values without a year
Public Function TryParseBirthday(ByVal s As String, ByRef d As Date) As Boolean
    Dim digits As String
    Dim y As Long, m As Long, dd As Long, t As Long

    t = InStr(s, "T")
    If t > 0 Then s = Left$(s, t - 1)
    s = Trim$(s)
    digits = Replace(s, "-", vbNullString)
    If Len(digits) <> 8 Or Not IsDigits(digits) Then Exit Function
    y = CLng(Left$(digits, 4))
    m = CLng(Mid$(digits, 5, 2))
    dd = CLng(Mid$(digits, 7, 2))
    If y < 1900 Or m < 1 Or m > 12 Or dd < 1 Or dd > 31 Then Exit Function
    d = DateSerial(y, m, dd)
    If Day(d) <> dd Then Exit Function
    TryParseBirthday = True
End Function

' Converts a phone number to the international format (+<country code><number>) when the form is unambiguous:
' numbers that already start with + or 00, or national numbers that start with the national prefix.
' Letters, extensions, short codes, unknown country and bare national numbers are left unchanged (returns False).
Public Function TryNormalizePhone(ByVal s As String, ByVal countryCode As String, ByVal trunk As String, ByRef result As String) As Boolean
    Dim k As Long, ch As String, digits As String, intl As String
    Dim hasPlus As Boolean

    result = s
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function

    For k = 1 To Len(s)
        If InStr("0123456789 +-()./", Mid$(s, k, 1)) = 0 Then Exit Function
    Next k
    If InStr(2, s, "+") > 0 Then Exit Function

    hasPlus = (Left$(s, 1) = "+")
    ' "+44 (0)20 ..." writes the national prefix in brackets after the country code
    If hasPlus Then s = Replace(s, "(0)", vbNullString)
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch
    Next k
    If Len(digits) = 0 Then Exit Function

    If hasPlus Then
        ' A national prefix typed after the default country code ("+90 0532 ...") is dropped; Italy keeps its 0
        If Len(countryCode) > 0 And Len(trunk) > 0 And countryCode <> "39" Then
            If Left$(digits, Len(countryCode)) = countryCode Then
                If Mid$(digits, Len(countryCode) + 1, Len(trunk)) = trunk And Len(digits) > Len(countryCode) + Len(trunk) Then
                    digits = countryCode & Mid$(digits, Len(countryCode) + Len(trunk) + 1)
                End If
            End If
        End If
        intl = digits
    ElseIf Left$(digits, 2) = "00" Then
        intl = Mid$(digits, 3)
    Else
        If Len(countryCode) = 0 Then Exit Function
        If Len(trunk) > 0 And Left$(digits, Len(trunk)) = trunk And Len(digits) > Len(trunk) Then
            If countryCode = "39" Then
                intl = countryCode & digits
            Else
                intl = countryCode & Mid$(digits, Len(trunk) + 1)
            End If
        ElseIf countryCode = "1" And Len(digits) = 10 Then
            intl = countryCode & digits
        Else
            Exit Function
        End If
    End If

    If Len(intl) < 7 Or Len(intl) > 15 Then Exit Function
    result = "+" & intl
    TryNormalizePhone = True
End Function

' Normalized number, or the original text when it is not recognized
Public Function NormalizePhone(ByVal s As String, ByVal countryCode As String, ByVal trunk As String) As String
    Dim res As String
    If TryNormalizePhone(s, countryCode, trunk, res) Then
        NormalizePhone = res
    Else
        NormalizePhone = s
    End If
End Function

' Folds a vCard line at 75 octets (UTF-8) with CRLF + space, never splitting a character
Public Function FoldLine(ByVal s As String) As String
    Dim res As String, ch As String
    Dim k As Long, code As Long, size As Long, used As Long, limit As Long, unit As Long

    limit = 75
    k = 1
    Do While k <= Len(s)
        code = AscW(Mid$(s, k, 1))
        If code < 0 Then code = code + 65536
        unit = 1
        If code < &H80 Then
            size = 1
        ElseIf code < &H800 Then
            size = 2
        ElseIf code >= &HD800 And code <= &HDBFF And k < Len(s) Then
            size = 4
            unit = 2
        Else
            size = 3
        End If
        If used + size > limit Then
            res = res & vbCrLf & " "
            used = 1
            limit = 75
        End If
        res = res & Mid$(s, k, unit)
        used = used + size
        k = k + unit
    Loop
    FoldLine = res
End Function

' Decodes a quoted-printable value (=XX bytes) using the property's charset
Public Function DecodeQuotedPrintable(ByVal s As String, ByVal charset As String) As String
    Dim bytes() As Byte
    Dim n As Long, k As Long, code As Long
    Dim hex As String

    If InStr(s, "=") = 0 Then
        DecodeQuotedPrintable = s
        Exit Function
    End If
    ReDim bytes(0 To Len(s))
    k = 1
    Do While k <= Len(s)
        hex = vbNullString
        If Mid$(s, k, 1) = "=" And k + 2 <= Len(s) Then hex = Mid$(s, k + 1, 2)
        If Len(hex) = 2 And IsHex(hex) Then
            bytes(n) = CByte("&H" & hex)
            k = k + 3
        Else
            code = AscW(Mid$(s, k, 1))
            If code < 0 Or code > 255 Then code = 63
            bytes(n) = CByte(code)
            k = k + 1
        End If
        n = n + 1
    Loop
    ReDim Preserve bytes(0 To n - 1)
    DecodeQuotedPrintable = BytesToText(bytes, charset)
End Function

' Decodes raw bytes with the given charset (UTF-8 when empty or unknown)
Public Function BytesToText(ByRef bytes() As Byte, ByVal charset As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    On Error Resume Next
    stream.charset = IIf(Len(charset) > 0, charset, "utf-8")
    If Err.Number <> 0 Then
        Err.Clear
        stream.charset = "utf-8"
    End If
    On Error GoTo 0
    BytesToText = stream.ReadText
    stream.Close
End Function

' Charset of a file's bytes: BOM first, then a strict UTF-8 check; anything else uses the ANSI code page
Public Function DetectCharset(ByRef bytes() As Byte, ByVal AnsiCharset As String) As String
    Dim n As Long
    n = UBound(bytes) + 1
    DetectCharset = AnsiCharset
    If n >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            DetectCharset = "utf-8"
            Exit Function
        End If
    End If
    If n >= 2 Then
        If bytes(0) = &HFF And bytes(1) = &HFE Then
            DetectCharset = "unicode"
            Exit Function
        End If
        If bytes(0) = &HFE And bytes(1) = &HFF Then
            DetectCharset = "unicodeFFFE"
            Exit Function
        End If
    End If
    If IsValidUtf8(bytes) Then DetectCharset = "utf-8"
End Function

' Strict UTF-8 check (no overlong forms, surrogates or values above U+10FFFF); ASCII-only data counts as valid
Public Function IsValidUtf8(ByRef bytes() As Byte) As Boolean
    Dim i As Long, k As Long, last As Long, need As Long
    Dim b As Long, second As Long

    last = UBound(bytes)
    i = 0
    Do While i <= last
        b = bytes(i)
        If b < &H80 Then
            i = i + 1
        Else
            If b >= &HC2 And b <= &HDF Then
                need = 1
            ElseIf b >= &HE0 And b <= &HEF Then
                need = 2
            ElseIf b >= &HF0 And b <= &HF4 Then
                need = 3
            Else
                Exit Function
            End If
            If i + need > last Then Exit Function
            second = bytes(i + 1)
            If b = &HE0 And second < &HA0 Then Exit Function
            If b = &HED And second > &H9F Then Exit Function
            If b = &HF0 And second < &H90 Then Exit Function
            If b = &HF4 And second > &H8F Then Exit Function
            For k = 1 To need
                If (bytes(i + k) And &HC0) <> &H80 Then Exit Function
            Next k
            i = i + need + 1
        End If
    Loop
    IsValidUtf8 = True
End Function

' Reads a file and decodes it. mode is Auto, UTF-8 or ANSI; encodingUsed returns the charset that was applied.
Public Function ReadTextFile(ByVal path As String, ByVal mode As String, ByVal AnsiCharset As String, ByRef encodingUsed As String) As String
    Dim stream As Object, raw As Variant
    Dim bytes() As Byte
    Dim s As String

    encodingUsed = "utf-8"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.LoadFromFile path
    If stream.size = 0 Then
        stream.Close
        Exit Function
    End If
    raw = stream.Read
    stream.Close

    bytes = raw
    Select Case LCase$(mode)
        Case "utf-8"
            encodingUsed = "utf-8"
        Case "ansi"
            encodingUsed = AnsiCharset
        Case Else
            encodingUsed = DetectCharset(bytes, AnsiCharset)
    End Select
    s = BytesToText(bytes, encodingUsed)
    If Left$(s, 1) = ChrW$(&HFEFF) Then s = Mid$(s, 2)
    ReadTextFile = s
End Function

' Saves text as UTF-8; the BOM that the stream writes is removed unless bom is True
Public Sub SaveText(ByVal path As String, ByVal text As String, ByVal bom As Boolean)
    Const adTypeBinary As Long = 1
    Const adTypeText As Long = 2
    Const adSaveCreateOverWrite As Long = 2
    Dim textStream As Object, binStream As Object

    Set textStream = CreateObject("ADODB.Stream")
    textStream.Type = adTypeText
    textStream.charset = "utf-8"
    textStream.Open
    textStream.WriteText text
    textStream.Position = 0
    textStream.Type = adTypeBinary
    If Not bom Then textStream.Position = 3

    Set binStream = CreateObject("ADODB.Stream")
    binStream.Type = adTypeBinary
    binStream.Open
    textStream.CopyTo binStream
    binStream.SaveToFile path, adSaveCreateOverWrite
    binStream.Close
    textStream.Close
End Sub
