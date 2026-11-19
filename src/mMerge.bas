Attribute VB_Name = "mMerge"
' Mesut Akcan
' https://github.com/mesutakcan/Excel-VCF-Converter
' youtube.com/mesutakcan
Option Explicit

' Finds records that describe the same person. Records are linked by shared strong keys (mobile/landline numbers
' and e-mail addresses); their names, written in any split, must be compatible. Rows that are identical to another
' row are treated as copies of it and are collapsed before the records are grouped.
' Part 1: matching core (no sheet access). Part 2: plan of the Data sheet and the analysis report on the Merge sheet.
' The analysis never changes the Data sheet; mMergeApply carries out the plan.

Public Const LEVEL_EXACT As Long = 0
Public Const LEVEL_SAFE As Long = 1
Public Const LEVEL_NORMAL As Long = 2
Public Const LEVEL_EXTENDED As Long = 3

Public Const REL_EQUAL As Long = 0
Public Const REL_SUBSET As Long = 1
Public Const REL_INITIAL As Long = 2
Public Const REL_TYPO As Long = 3
Public Const REL_CONFLICT As Long = 9

Public Const ST_SINGLE As Long = 0
Public Const ST_AUTO As Long = 1
Public Const ST_REVIEW As Long = 2

' Honorifics and kinship titles that are ignored when comparing names
Private Const STOP_WORDS As String = "|dr|prof|doc|av|sn|sayin|bey|hanim|mr|mrs|ms|abi|abim|abla|ablam|kuzen|kardes|kardesim|"
' A key shared by more than this many records is too common to link them (e.g. a company switchboard)
Private Const MAX_KEY_GROUP As Long = 50

' Result of ClusterRows (index = record number 1..n)
Public gRoot() As Long
Public gStatus() As Long
Public gRevA() As Long
Public gRevB() As Long
Public gRevN As Long

Private mParent() As Long
Private mStart() As Long
Private mCount() As Long
Private mList() As Long

Public Const PLAN_OK As Long = 0
Public Const PLAN_FILTERED As Long = 1
Public Const PLAN_NO_DATA As Long = 2

' Plan made by BuildPlan. Records are numbered 1..gN from the first row below the header.
' gIsCopy(i): row i is identical to the earlier row gRepOf(i); gPrimary(root): main record of a cluster that is merged
Public gData As Variant
Public gN As Long
Public gFirstRow As Long
Public gLastCol As Long
Public gBlank() As Boolean
Public gIsCopy() As Boolean
Public gRepOf() As Long
Public gPrimary() As Long
Public gRecords As Long
Public gAutoGroups As Long
Public gCopyRows As Long
Public gRowsGone As Long

' ---------------------------------------------------------------- text and name tokens

Private Function FoldChar(ByVal code As Long) As String
    Select Case code
        Case 48 To 57: FoldChar = Chr$(code)
        Case 65 To 90: FoldChar = Chr$(code + 32)
        Case 97 To 122: FoldChar = Chr$(code)
        Case 199, 231: FoldChar = "c"
        Case 286, 287: FoldChar = "g"
        Case 304, 305: FoldChar = "i"
        Case 214, 246: FoldChar = "o"
        Case 350, 351: FoldChar = "s"
        Case 220, 252: FoldChar = "u"
        Case 192 To 197, 224 To 229: FoldChar = "a"
        Case 200 To 203, 232 To 235: FoldChar = "e"
        Case 204 To 207, 236 To 239: FoldChar = "i"
        Case 209, 241: FoldChar = "n"
        Case 210 To 213, 242 To 245: FoldChar = "o"
        Case 217 To 219, 249 To 251: FoldChar = "u"
        Case 221, 253, 255: FoldChar = "y"
    End Select
End Function

' Lower case ASCII letters and digits, Turkish and Latin-1 letters reduced to ASCII, everything else becomes one space.
' The case and accent rules are written out so the result does not depend on the Windows language.
Public Function FoldText(ByVal s As String) As String
    Dim res As String, ch As String
    Dim k As Long, code As Long
    Dim lastSpace As Boolean

    lastSpace = True
    For k = 1 To Len(s)
        code = AscW(Mid$(s, k, 1))
        If code < 0 Then code = code + 65536
        ch = FoldChar(code)
        If Len(ch) = 0 Then
            If Not lastSpace Then
                res = res & " "
                lastSpace = True
            End If
        Else
            res = res & ch
            lastSpace = False
        End If
    Next k
    If Right$(res, 1) = " " Then res = Left$(res, Len(res) - 1)
    FoldText = res
End Function

' Name words as "|word1|word2|" (folded, unique, honorifics removed); "" when there are none
Public Function NameTokens(ByVal text As String) As String
    Dim parts() As String
    Dim res As String, t As String
    Dim k As Long

    text = FoldText(text)
    If Len(text) = 0 Then Exit Function
    parts = Split(text, " ")
    res = "|"
    For k = 0 To UBound(parts)
        t = parts(k)
        If Len(t) > 0 Then
            If InStr(STOP_WORDS, "|" & t & "|") = 0 Then
                If InStr(res, "|" & t & "|") = 0 Then res = res & t & "|"
            End If
        End If
    Next k
    If res <> "|" Then
        NameTokens = res
    Else
        NameTokens = "|" & Replace(Trim$(text), " ", "|") & "|"
    End If
End Function

Private Function TokenArray(ByVal tokens As String) As String()
    If Len(tokens) < 3 Then
        TokenArray = Split(vbNullString, "|")
    Else
        TokenArray = Split(Mid$(tokens, 2, Len(tokens) - 2), "|")
    End If
End Function

' True when every word of arr is in the token string
Private Function AllIn(ByRef arr() As String, ByVal tokens As String) As Boolean
    Dim k As Long
    For k = 0 To UBound(arr)
        If Len(arr(k)) > 0 Then
            If InStr(tokens, "|" & arr(k) & "|") = 0 Then Exit Function
        End If
    Next k
    AllIn = True
End Function

' Words of arr that are not in the token string, as "|w1|w2|"
Private Function MissingFrom(ByRef arr() As String, ByVal tokens As String) As String
    Dim k As Long, res As String
    For k = 0 To UBound(arr)
        If Len(arr(k)) > 0 Then
            If InStr(tokens, "|" & arr(k) & "|") = 0 Then res = res & arr(k) & "|"
        End If
    Next k
    If Len(res) > 0 Then MissingFrom = "|" & res
End Function

' One edit apart (insert, delete or replace one letter); both words need at least 4 letters
Private Function Lev1(ByVal a As String, ByVal b As String) As Boolean
    Dim i As Long
    If Len(a) < 4 Or Len(b) < 4 Then Exit Function
    If Abs(Len(a) - Len(b)) > 1 Then Exit Function
    If a = b Then
        Lev1 = True
        Exit Function
    End If
    i = 0
    Do While i < Len(a) And i < Len(b)
        If Mid$(a, i + 1, 1) <> Mid$(b, i + 1, 1) Then Exit Do
        i = i + 1
    Loop
    Lev1 = (Mid$(a, i + 2) = Mid$(b, i + 1)) Or (Mid$(a, i + 1) = Mid$(b, i + 2)) Or (Mid$(a, i + 2) = Mid$(b, i + 2))
End Function

' Every word in x is a single letter that starts some word in y
Private Function InitialsMatch(ByRef x() As String, ByVal y As String) As Boolean
    Dim k As Long, ya() As String, j As Long, found As Boolean
    ya = TokenArray(y)
    For k = 0 To UBound(x)
        If Len(x(k)) <> 1 Then Exit Function
        found = False
        For j = 0 To UBound(ya)
            If Left$(ya(j), 1) = x(k) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then Exit Function
    Next k
    InitialsMatch = True
End Function

' Every word in x has a one-letter-different partner in y
Private Function TyposMatch(ByRef x() As String, ByVal y As String) As Boolean
    Dim k As Long, j As Long, ya() As String, found As Boolean
    ya = TokenArray(y)
    For k = 0 To UBound(x)
        found = False
        For j = 0 To UBound(ya)
            If Lev1(x(k), ya(j)) Then
                found = True
                Exit For
            End If
        Next j
        If Not found Then Exit Function
    Next k
    TyposMatch = True
End Function

' How two names relate: equal words, one name inside the other (also when one is empty),
' initials against full words, one-letter typos, or conflict
Public Function NameRelation(ByVal a As String, ByVal b As String) As Long
    Dim ta() As String, tb() As String, ma() As String, mb() As String
    Dim missA As String, missB As String
    Dim aInB As Boolean, bInA As Boolean

    ta = TokenArray(a)
    tb = TokenArray(b)
    aInB = AllIn(ta, b)
    bInA = AllIn(tb, a)
    If aInB And bInA Then
        NameRelation = REL_EQUAL
        Exit Function
    End If
    If aInB Or bInA Then
        NameRelation = REL_SUBSET
        Exit Function
    End If

    missA = MissingFrom(ta, b)
    missB = MissingFrom(tb, a)
    ma = TokenArray(missA)
    mb = TokenArray(missB)
    If InitialsMatch(ma, missB) And InitialsMatch(mb, missA) Then
        NameRelation = REL_INITIAL
    ElseIf TyposMatch(ma, missB) And TyposMatch(mb, missA) Then
        NameRelation = REL_TYPO
    Else
        NameRelation = REL_CONFLICT
    End If
End Function

' Safe: equal or contained names. Normal also accepts initials and one-letter typos. Exact: strictly equal names only.
Public Function RelationCompatible(ByVal rel As Long, ByVal level As Long) As Boolean
    If level = LEVEL_EXACT Then
        RelationCompatible = (rel = REL_EQUAL)
    ElseIf rel = REL_EQUAL Or rel = REL_SUBSET Then
        RelationCompatible = True
    ElseIf (level = LEVEL_NORMAL Or level = LEVEL_EXTENDED) And (rel = REL_INITIAL Or rel = REL_TYPO) Then
        RelationCompatible = True
    End If
End Function

' ---------------------------------------------------------------- keys

Public Function LowerAscii(ByVal s As String) As String
    Dim k As Long, code As Long, res As String
    For k = 1 To Len(s)
        code = AscW(Mid$(s, k, 1))
        If code >= 65 And code <= 90 Then
            res = res & Chr$(code + 32)
        Else
            res = res & Mid$(s, k, 1)
        End If
    Next k
    LowerAscii = res
End Function

' Toll-free and call-center numbers (444, 0850, 0800) are shared by many unrelated records
Private Function IsServiceNumber(ByVal digits As String) As Boolean
    Dim n As String
    n = digits
    If Left$(n, 2) = "90" And Len(n) > 10 Then n = Mid$(n, 3)
    If Left$(n, 1) = "0" Then n = Mid$(n, 2)
    If Left$(n, 3) = "444" Or Left$(n, 3) = "850" Or Left$(n, 3) = "800" Then
        IsServiceNumber = True
    ElseIf Len(n) = 10 And Mid$(n, 4, 3) = "444" Then
        IsServiceNumber = True
    End If
End Function

' Digits of a phone number with the country part ignored (last 10 digits).
' Service numbers get an "s" key: they may link only records whose names are exactly the same.
Public Function NumberKey(ByVal s As String) As String
    Dim k As Long, ch As String, d As String
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next k
    If Len(d) < 7 Then Exit Function
    If Len(d) >= 10 Then
        d = Right$(d, 10)
    End If
    If IsServiceNumber(d) Then
        NumberKey = "s" & d
    Else
        NumberKey = "p" & d
    End If
End Function

Public Function EmailKey(ByVal s As String) As String
    Dim e As String
    e = LowerAscii(Trim$(s))
    If InStr(e, "@") > 1 Then EmailKey = "e" & e
End Function

' Adds a key to a record's key list (keys separated by Chr(1), each key once)
Public Function AddKey(ByVal keys As String, ByVal key As String) As String
    AddKey = keys
    If Len(key) = 0 Then Exit Function
    If InStr(Chr$(1) & keys & Chr$(1), Chr$(1) & key & Chr$(1)) > 0 Then Exit Function
    If Len(keys) = 0 Then
        AddKey = key
    Else
        AddKey = keys & Chr$(1) & key
    End If
End Function

Public Function CountShared(ByVal ka As String, ByVal kb As String) As Long
    Dim a() As String, k As Long, n As Long
    If Len(ka) = 0 Or Len(kb) = 0 Then Exit Function
    a = Split(ka, Chr$(1))
    kb = Chr$(1) & kb & Chr$(1)
    For k = 0 To UBound(a)
        If InStr(kb, Chr$(1) & a(k) & Chr$(1)) > 0 Then n = n + 1
    Next k
    CountShared = n
End Function

' Counts shared phone ("p") and email ("e") keys between two records
Public Function CountSharedContactKeys(ByVal ka As String, ByVal kb As String) As Long
    Dim a() As String, k As Long, n As Long
    If Len(ka) = 0 Or Len(kb) = 0 Then Exit Function
    a = Split(ka, Chr$(1))
    kb = Chr$(1) & kb & Chr$(1)
    For k = 0 To UBound(a)
        If Left$(a(k), 1) = "p" Or Left$(a(k), 1) = "e" Then
            If InStr(kb, Chr$(1) & a(k) & Chr$(1)) > 0 Then n = n + 1
        End If
    Next k
    CountSharedContactKeys = n
End Function

' Shared call-centre/service-number keys.  These are deliberately kept separate
' from ordinary phone keys because many unrelated contacts may share one.
Private Function CountSharedServiceKeys(ByVal ka As String, ByVal kb As String) As Long
    Dim a() As String, k As Long, n As Long
    If Len(ka) = 0 Or Len(kb) = 0 Then Exit Function
    a = Split(ka, Chr$(1))
    kb = Chr$(1) & kb & Chr$(1)
    For k = 0 To UBound(a)
        If Left$(a(k), 1) = "s" Then
            If InStr(kb, Chr$(1) & a(k) & Chr$(1)) > 0 Then n = n + 1
        End If
    Next k
    CountSharedServiceKeys = n
End Function

' Alphabetically sorted name key for detecting same-name candidates (requires at least 2 distinct words)
Public Function SortedNameKey(ByVal tokens As String) As String
    Dim parts() As String, i As Long, j As Long, t As String
    If Len(tokens) < 3 Then Exit Function
    parts = TokenArray(tokens)
    If UBound(parts) < 1 Then Exit Function
    For i = 0 To UBound(parts) - 1
        For j = i + 1 To UBound(parts)
            If StrComp(parts(i), parts(j), vbBinaryCompare) > 0 Then
                t = parts(i)
                parts(i) = parts(j)
                parts(j) = t
            End If
        Next j
    Next i
    SortedNameKey = "n:" & Join(parts, "_")
End Function

' ---------------------------------------------------------------- clustering

Private Function FindRoot(ByVal x As Long) As Long
    Do While mParent(x) <> x
        mParent(x) = mParent(mParent(x))
        x = mParent(x)
    Loop
    FindRoot = x
End Function

Private Sub UnionRows(ByVal a As Long, ByVal b As Long)
    Dim ra As Long, rb As Long
    ra = FindRoot(a)
    rb = FindRoot(b)
    If ra <> rb Then mParent(ra) = rb
End Sub

' Two records may be grouped for auto-merge: MUST share at least one phone/email AND have compatible names.
' Records sharing only a name key (different people like two "Ali AK") return False and are routed to the review list.
Private Function PairGood(ByRef tok() As String, ByRef keys() As String, ByVal i As Long, ByVal j As Long, ByVal level As Long) As Boolean
    Dim rel As Long, sharedContacts As Long

    sharedContacts = CountSharedContactKeys(keys(i), keys(j))
    rel = NameRelation(tok(i), tok(j))
    ' A normal phone or e-mail key is a strong link.  A shared service number is
    ' allowed only when the complete normalized names are equal.
    If sharedContacts = 0 Then
        If CountSharedServiceKeys(keys(i), keys(j)) = 0 Or rel <> REL_EQUAL Then Exit Function
    End If

    If RelationCompatible(rel, level) Then
        PairGood = True
    ElseIf level = LEVEL_EXTENDED Then
        ' If names explicitly conflict (e.g. different persons sharing an office switchboard), do not auto-group!
        If rel <> REL_CONFLICT Then
            PairGood = (sharedContacts >= 2)
        End If
    End If
End Function

' Array growth helpers using native ReDim Preserve
Private Sub GrowDoubles(ByRef arr() As Double, ByVal newMax As Long)
    ReDim Preserve arr(1 To newMax)
End Sub

Public Sub GrowLongs(ByRef arr() As Long, ByVal newMax As Long)
    ReDim Preserve arr(1 To newMax)
End Sub

Private Sub ShellSortKeys(ByRef k() As String, ByRef r() As Long, ByVal n As Long)
    Dim gap As Long, i As Long, j As Long
    Dim tk As String, tr As Long
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            tk = k(i)
            tr = r(i)
            j = i
            Do While j > gap
                If StrComp(k(j - gap), tk, vbBinaryCompare) <= 0 Then Exit Do
                k(j) = k(j - gap)
                r(j) = r(j - gap)
                j = j - gap
            Loop
            k(j) = tk
            r(j) = tr
        Next i
        gap = gap \ 2
    Loop
End Sub

Private Sub ShellSortDoubles(ByRef v() As Double, ByVal n As Long)
    Dim gap As Long, i As Long, j As Long
    Dim t As Double
    gap = n \ 2
    Do While gap > 0
        For i = gap + 1 To n
            t = v(i)
            j = i
            Do While j > gap
                If v(j - gap) <= t Then Exit Do
                v(j) = v(j - gap)
                j = j - gap
            Loop
            v(j) = t
        Next i
        gap = gap \ 2
    Loop
End Sub

' VBA's And/Or operators evaluate both operands.  Keep the lower-bound check
' separate so that testing the first candidate never attempts pairs(0).
Private Function IsNewPair(ByRef pairs() As Double, ByVal index As Long) As Boolean
    If index = 1 Then
        IsNewPair = True
    Else
        IsNewPair = (pairs(index) <> pairs(index - 1))
    End If
End Function

Private Function CanMergeClusters(ByRef tok() As String, ByRef keys() As String, ByVal rootA As Long, ByVal rootB As Long, ByRef head() As Long, ByRef nxt() As Long, ByVal level As Long) As Boolean
    Dim cA As Long, cB As Long, cand As Long, t As Long
    Dim allOk As Boolean, namesOk As Boolean, hubOk As Boolean
    Dim candRoots(1 To 2) As Long, rIdx As Long, targetIdx As Long, candRoot As Long, targetRoot As Long

    ' 1. Check all-pairs clique
    allOk = True
    cA = head(rootA)
    Do While cA > 0
        cB = head(rootB)
        Do While cB > 0
            If Not PairGood(tok, keys, cA, cB, level) Then
                allOk = False
                Exit Do
            End If
            cB = nxt(cB)
        Loop
        If Not allOk Then Exit Do
        cA = nxt(cA)
    Loop
    If allOk Then
        CanMergeClusters = True
        Exit Function
    End If

    ' In EXACT mode, only all-pairs cliques are permitted (no hubs bridging different names/numbers)
    If level = LEVEL_EXACT Then
        CanMergeClusters = False
        Exit Function
    End If

    ' 2. Check Hub Consistency: All names between rootA and rootB must be mutually compatible
    namesOk = True
    cA = head(rootA)
    Do While cA > 0
        cB = head(rootB)
        Do While cB > 0
            If Not RelationCompatible(NameRelation(tok(cA), tok(cB)), level) Then
                namesOk = False
                Exit Do
            End If
            cB = nxt(cB)
        Loop
        If Not namesOk Then Exit Do
        cA = nxt(cA)
    Loop
    If Not namesOk Then
        CanMergeClusters = False
        Exit Function
    End If

    ' Test if any member in rootA or rootB is a Hub connecting to all members in both clusters
    candRoots(1) = rootA
    candRoots(2) = rootB
    For rIdx = 1 To 2
        candRoot = candRoots(rIdx)
        cand = head(candRoot)
        Do While cand > 0
            hubOk = True
            For targetIdx = 1 To 2
                targetRoot = candRoots(targetIdx)
                t = head(targetRoot)
                Do While t > 0
                    If t <> cand Then
                        If Not PairGood(tok, keys, cand, t, level) Then
                            hubOk = False
                            Exit Do
                        End If
                    End If
                    t = nxt(t)
                Loop
                If Not hubOk Then Exit For
            Next targetIdx
            If hubOk Then
                CanMergeClusters = True
                Exit Function
            End If
            cand = nxt(cand)
        Loop
    Next rIdx

    CanMergeClusters = False
End Function

Private Sub UnionRoots(ByVal rootA As Long, ByVal rootB As Long, ByRef head() As Long, ByRef tail() As Long, ByRef nxt() As Long, ByRef cSize() As Long)
    mParent(rootB) = rootA
    nxt(tail(rootA)) = head(rootB)
    tail(rootA) = tail(rootB)
    cSize(rootA) = cSize(rootA) + cSize(rootB)
End Sub

Private Sub RecordReviewLink(ByVal i As Long, ByVal j As Long)
    gRevN = gRevN + 1
    If gRevN > UBound(gRevA) Then
        GrowLongs gRevA, UBound(gRevA) * 2
        GrowLongs gRevB, UBound(gRevB) * 2
    End If
    gRevA(gRevN) = i
    gRevB(gRevN) = j
End Sub

' Groups records 1..n into clusters. tok(i) holds the name words of record i, keys(i) its keys.
' Results: gRoot(i) = cluster of record i, gStatus(root) = ST_SINGLE / ST_AUTO / ST_REVIEW and the pairs that
' share a key but were not grouped (gRevA/gRevB). Returns the number of clusters with more than one record.
Public Function ClusterRows(ByRef tok() As String, ByRef keys() As String, ByVal n As Long, ByVal level As Long) As Long
    Dim eKey() As String, eRow() As Long
    Dim pairs() As Double
    Dim parts() As String
    Dim m As Long, i As Long, j As Long, k As Long, g0 As Long, g1 As Long, np As Long
    Dim a As Long, b As Long, rootA As Long, rootB As Long, big As Long
    Dim head() As Long, tail() As Long, nxt() As Long, cSize() As Long

    ReDim mParent(1 To n)
    ReDim head(1 To n)
    ReDim tail(1 To n)
    ReDim nxt(1 To n)
    ReDim cSize(1 To n)
    For i = 1 To n
        mParent(i) = i
        head(i) = i
        tail(i) = i
        nxt(i) = 0
        cSize(i) = 1
    Next i
    ReDim gRoot(1 To n)
    ReDim gStatus(1 To n)
    ReDim gRevA(1 To 16)
    ReDim gRevB(1 To 16)
    gRevN = 0
    If n < 2 Then
        For i = 1 To n
            gRoot(i) = i
            gStatus(i) = ST_SINGLE
        Next i
        Exit Function
    End If

    ' one entry per (key, record)
    For i = 1 To n
        If Len(keys(i)) > 0 Then m = m + UBound(Split(keys(i), Chr$(1))) + 1
    Next i
    If m > 0 Then
        ReDim eKey(1 To m)
        ReDim eRow(1 To m)
        m = 0
        For i = 1 To n
            If Len(keys(i)) > 0 Then
                parts = Split(keys(i), Chr$(1))
                For k = 0 To UBound(parts)
                    m = m + 1
                    eKey(m) = parts(k)
                    eRow(m) = i
                Next k
            End If
        Next i
        ShellSortKeys eKey, eRow, m

        ' candidate pairs: records that share a key
        ReDim pairs(1 To 64)
        g0 = 1
        Do While g0 <= m
            g1 = g0
            Do While g1 < m
                If eKey(g1 + 1) <> eKey(g0) Then Exit Do
                g1 = g1 + 1
            Loop
            If g1 - g0 + 1 <= MAX_KEY_GROUP Then
                For a = g0 To g1 - 1
                    For b = a + 1 To g1
                        i = eRow(a)
                        j = eRow(b)
                        If i > j Then
                            k = i
                            i = j
                            j = k
                        End If
                        np = np + 1
                        If np > UBound(pairs) Then GrowDoubles pairs, UBound(pairs) * 2
                        pairs(np) = CDbl(i) * (n + 1) + j
                    Next b
                Next a
            End If
            g0 = g1 + 1
        Loop

        If np > 0 Then
            ShellSortDoubles pairs, np
            
            ' Pass 1: exact same name matches (highest priority)
            For k = 1 To np
                If IsNewPair(pairs, k) Then
                    i = CLng(Int(pairs(k) / (n + 1)))
                    j = CLng(pairs(k) - CDbl(i) * (n + 1))
                    If PairGood(tok, keys, i, j, level) Then
                        If NameRelation(tok(i), tok(j)) = REL_EQUAL Then
                            rootA = FindRoot(i)
                            rootB = FindRoot(j)
                            If rootA <> rootB Then
                                If CanMergeClusters(tok, keys, rootA, rootB, head, nxt, level) Then
                                    UnionRoots rootA, rootB, head, tail, nxt, cSize
                                End If
                            End If
                        End If
                    End If
                End If
                If k Mod 100 = 0 Or k = np Then ProgressUpdate 0.4 + 0.1 * k / np
            Next k

            ' Pass 2: other compatible pairs (if not LEVEL_EXACT)
            If level <> LEVEL_EXACT Then
                For k = 1 To np
                    If IsNewPair(pairs, k) Then
                        i = CLng(Int(pairs(k) / (n + 1)))
                        j = CLng(pairs(k) - CDbl(i) * (n + 1))
                        If PairGood(tok, keys, i, j, level) Then
                            If NameRelation(tok(i), tok(j)) <> REL_EQUAL Then
                                rootA = FindRoot(i)
                                rootB = FindRoot(j)
                                If rootA <> rootB Then
                                    If CanMergeClusters(tok, keys, rootA, rootB, head, nxt, level) Then
                                        UnionRoots rootA, rootB, head, tail, nxt, cSize
                                    Else
                                        RecordReviewLink i, j
                                    End If
                                End If
                            End If
                        Else
                            RecordReviewLink i, j
                        End If
                    End If
                    If k Mod 100 = 0 Or k = np Then ProgressUpdate 0.5 + 0.1 * k / np
                Next k
            End If
        End If
    End If

    ' cluster membership and status assignment
    For i = 1 To n
        gRoot(i) = FindRoot(i)
    Next i

    big = 0
    For i = 1 To n
        If gRoot(i) = i Then
            If cSize(i) > 1 Then
                gStatus(i) = ST_AUTO
                big = big + 1
            Else
                gStatus(i) = ST_SINGLE
            End If
        End If
    Next i
    For i = 1 To n
        gStatus(i) = gStatus(gRoot(i))
    Next i

    ClusterRows = big
End Function

' ================================================================ part 2: analysis report

Private Function IsLowerCode(ByVal code As Long) As Boolean
    IsLowerCode = (code >= 97 And code <= 122) Or (code >= 223 And code <= 246) Or (code >= 248 And code <= 255) _
                  Or code = 287 Or code = 305 Or code = 351
End Function

Private Function IsUpperCode(ByVal code As Long) As Boolean
    IsUpperCode = (code >= 65 And code <= 90) Or (code >= 192 And code <= 214) Or (code >= 216 And code <= 222) _
                  Or code = 286 Or code = 304 Or code = 350
End Function

' Text written entirely in capitals (at least two letters), e.g. "ABDULLAH"
Public Function IsAllUpper(ByVal s As String) As Boolean
    Dim k As Long, code As Long, upper As Long
    For k = 1 To Len(s)
        code = AscW(Mid$(s, k, 1))
        If code < 0 Then code = code + 65536
        If IsLowerCode(code) Then Exit Function
        If IsUpperCode(code) Then upper = upper + 1
    Next k
    IsAllUpper = (upper >= 2)
End Function

' Higher is better: name split into given/middle/family, family name present, not written in capitals, more filled cells
Public Function NameScore(ByVal given As String, ByVal middle As String, ByVal family As String, ByVal filled As Long) As Double
    Dim score As Double
    If Len(given) > 0 Then score = score + 10
    If Len(middle) > 0 Then score = score + 10
    If Len(family) > 0 Then score = score + 15
    If Len(given) > 0 And IsAllUpper(given) Then score = score - 6
    If Len(middle) > 0 And IsAllUpper(middle) Then score = score - 6
    If Len(family) > 0 And IsAllUpper(family) Then score = score - 6
    If filled > 40 Then filled = 40
    NameScore = score + filled / 10
End Function

Public Function FieldCell(ByRef data As Variant, ByVal r As Long, ByVal f As Long) As Variant
    If f > 0 Then
        If gFldCol(f) > 0 Then FieldCell = data(r, gFldCol(f))
    End If
End Function

' Name words and keys of every record; gBlank(i) marks completely empty rows
Private Sub PrepareRows(ByRef tok() As String, ByRef keys() As String, ByVal level As Long)
    Dim gi As Long, mi As Long, fi As Long, oi As Long
    Dim i As Long, f As Long, c As Long
    Dim t As String, k As String, nk As String

    gi = FieldIndexByFamily("N_GIVEN")
    mi = FieldIndexByFamily("N_MIDDLE")
    fi = FieldIndexByFamily("N_FAMILY")
    oi = FieldIndexByFamily("ORG")
    ReDim tok(1 To gN)
    ReDim keys(1 To gN)
    ReDim gBlank(1 To gN)
    For i = 1 To gN
        gBlank(i) = True
        For c = 1 To gLastCol
            If Len(CellText(gData(i, c))) > 0 Then
                gBlank(i) = False
                Exit For
            End If
        Next c
        If Not gBlank(i) Then
            t = NameTokens(CellText(FieldCell(gData, i, gi)) & " " & CellText(FieldCell(gData, i, mi)) & " " & CellText(FieldCell(gData, i, fi)))
            If Len(t) = 0 Then t = NameTokens(CellText(FieldCell(gData, i, oi)))
            tok(i) = t
            k = vbNullString
            For f = 1 To gFldCount
                If gFldCol(f) > 0 Then
                    If gFldProp(f) = "TEL" Then
                        k = AddKey(k, NumberKey(CellText(gData(i, gFldCol(f)))))
                    ElseIf gFldProp(f) = "EMAIL" Then
                        k = AddKey(k, EmailKey(CellText(gData(i, gFldCol(f)))))
                    End If
                End If
            Next f
            If level = LEVEL_EXTENDED Then
                nk = SortedNameKey(t)
                If Len(nk) > 0 Then k = AddKey(k, nk)
            End If
            keys(i) = k
        End If
        If i Mod 100 = 0 Or i = gN Then ProgressUpdate 0.05 + 0.2 * i / gN
    Next i
End Sub

' Rows with the same text in every column (ignoring case and outer spaces) are copies of the first such row.
' Copies get no name words and no keys, so they stay out of the grouping and join the group of their first row later.
Private Sub MarkIdentical(ByRef tok() As String, ByRef keys() As String)
    Dim sig() As String, idx() As Long
    Dim i As Long, c As Long, m As Long, g0 As Long, g1 As Long, k As Long, rep As Long
    Dim s As String

    ReDim gIsCopy(1 To gN)
    ReDim gRepOf(1 To gN)
    For i = 1 To gN
        If Not gBlank(i) Then m = m + 1
    Next i
    If m < 2 Then Exit Sub

    ReDim sig(1 To m)
    ReDim idx(1 To m)
    m = 0
    For i = 1 To gN
        If Not gBlank(i) Then
            m = m + 1
            s = vbNullString
            For c = 1 To gLastCol
                s = s & Chr$(1) & CellText(gData(i, c))
            Next c
            sig(m) = LowerAscii(s)
            idx(m) = i
        End If
        If i Mod 100 = 0 Or i = gN Then ProgressUpdate 0.25 + 0.15 * i / gN
    Next i
    ShellSortKeys sig, idx, m

    g0 = 1
    Do While g0 <= m
        g1 = g0
        Do While g1 < m
            If sig(g1 + 1) <> sig(g0) Then Exit Do
            g1 = g1 + 1
        Loop
        If g1 > g0 Then
            rep = idx(g0)
            For k = g0 + 1 To g1
                If idx(k) < rep Then rep = idx(k)
            Next k
            For k = g0 To g1
                If idx(k) <> rep Then
                    gIsCopy(idx(k)) = True
                    gRepOf(idx(k)) = rep
                    tok(idx(k)) = vbNullString
                    keys(idx(k)) = vbNullString
                End If
            Next k
        End If
        g0 = g1 + 1
    Loop
End Sub

' Members of every cluster, grouped by cluster root (mStart/mCount/mList)
Private Sub BuildMembers(ByVal n As Long)
    Dim i As Long, k As Long, rootId As Long
    Dim pos() As Long
    ReDim mCount(1 To n)
    ReDim mStart(1 To n)
    ReDim mList(1 To n)
    ReDim pos(1 To n)
    For i = 1 To n
        mCount(gRoot(i)) = mCount(gRoot(i)) + 1
    Next i
    k = 1
    For i = 1 To n
        mStart(i) = k
        pos(i) = k
        k = k + mCount(i)
    Next i
    For i = 1 To n
        rootId = gRoot(i)
        mList(pos(rootId)) = i
        pos(rootId) = pos(rootId) + 1
    Next i
End Sub

Public Function MemberCount(ByVal root As Long) As Long
    MemberCount = mCount(root)
End Function

' k-th record (1..MemberCount) of the cluster, in row order
Public Function MemberAt(ByVal root As Long, ByVal k As Long) As Long
    MemberAt = mList(mStart(root) + k - 1)
End Function

' Reads the Data sheet and groups its records (the sheet is not changed). Returns PLAN_OK, PLAN_FILTERED or PLAN_NO_DATA.
Public Function BuildPlan() As Long
    Dim ws As Worksheet
    Dim tok() As String, keys() As String, copiesIn() As Long, rows() As Long
    Dim lastRow As Long, level As Long, i As Long, k As Long, nr As Long, distinct As Long

    LoadFields
    Set ws = DataSheet()
    If Not ActiveSheet Is ws Then ws.Activate
    If ws.FilterMode Then
        BuildPlan = PLAN_FILTERED
        Exit Function
    End If

    gFirstRow = HEADER_ROW + 1
    gLastCol = LastFieldCol()
    lastRow = FindLastRow(ws, gFirstRow, gLastCol)
    If lastRow < gFirstRow Then
        BuildPlan = PLAN_NO_DATA
        Exit Function
    End If
    gN = lastRow - gFirstRow + 1
    gData = ws.Range(ws.Cells(gFirstRow, 1), ws.Cells(lastRow, gLastCol)).value

    Select Case LCase$(SettingText("merge_level"))
        Case "exact", "ozdes", "özdeþ": level = LEVEL_EXACT
        Case "safe", "guvenli", "güvenli": level = LEVEL_SAFE
        Case "extended", "kapsamli", "kapsamlý": level = LEVEL_EXTENDED
        Case Else: level = LEVEL_NORMAL
    End Select
    PrepareRows tok, keys, level
    ProgressUpdate 0.25
    MarkIdentical tok, keys
    ProgressUpdate 0.4
    ClusterRows tok, keys, gN, level
    ProgressUpdate 0.6
    For i = 1 To gN
        If gIsCopy(i) Then gRoot(i) = gRoot(gRepOf(i))
    Next i
    BuildMembers gN

    ReDim copiesIn(1 To gN)
    ReDim gPrimary(1 To gN)
    gRecords = 0
    gCopyRows = 0
    For i = 1 To gN
        If Not gBlank(i) Then gRecords = gRecords + 1
        If gIsCopy(i) Then
            gCopyRows = gCopyRows + 1
            copiesIn(gRoot(i)) = copiesIn(gRoot(i)) + 1
        End If
    Next i

    gAutoGroups = 0
    gRowsGone = gCopyRows
    For i = 1 To gN
        If gRoot(i) = i Then
            If mCount(i) > 1 And gStatus(i) = ST_SINGLE Then gStatus(i) = ST_AUTO
            If gStatus(i) = ST_AUTO Then
                distinct = mCount(i) - copiesIn(i)
                If distinct > 0 Then
                    ReDim rows(1 To distinct)
                    nr = 0
                    For k = 1 To mCount(i)
                        If Not gIsCopy(mList(mStart(i) + k - 1)) Then
                            nr = nr + 1
                            rows(nr) = mList(mStart(i) + k - 1)
                        End If
                    Next k
                    gPrimary(i) = PickPrimary(gData, rows, nr, gLastCol)
                    If distinct > 1 Then
                        gAutoGroups = gAutoGroups + 1
                        gRowsGone = gRowsGone + distinct - 1
                    End If
                Else
                    ' Tüm küme kopyalardan oluþuyorsa (orijinal kayýt baþka kümede veya tek baþýna)
                    gPrimary(i) = mList(mStart(i))
                End If
            End If
        End If
    Next i
    BuildPlan = PLAN_OK
End Function

' Record of the cluster with the best name layout (ties: the first one); rows are record numbers
Private Function PickPrimary(ByRef data As Variant, ByRef rows() As Long, ByVal nRows As Long, ByVal lastCol As Long) As Long
    Dim gi As Long, mi As Long, fi As Long
    Dim k As Long, c As Long, filled As Long
    Dim best As Double, sc As Double

    gi = FieldIndexByFamily("N_GIVEN")
    mi = FieldIndexByFamily("N_MIDDLE")
    fi = FieldIndexByFamily("N_FAMILY")
    best = -1000
    For k = 1 To nRows
        filled = 0
        For c = 1 To lastCol
            If Len(CellText(data(rows(k), c))) > 0 Then filled = filled + 1
        Next c
        sc = NameScore(CellText(FieldCell(data, rows(k), gi)), CellText(FieldCell(data, rows(k), mi)), CellText(FieldCell(data, rows(k), fi)), filled)
        If sc > best Then
            best = sc
            PickPrimary = rows(k)
        End If
    Next k
End Function

' Digits of a phone number, country part ignored; text without digits is compared as folded text
Public Function PhoneIdentity(ByVal v As String) As String
    Dim k As Long, ch As String, d As String
    For k = 1 To Len(v)
        ch = Mid$(v, k, 1)
        If ch >= "0" And ch <= "9" Then d = d & ch
    Next k
    If Len(d) = 0 Then
        PhoneIdentity = FoldText(v)
    ElseIf Len(d) >= 10 Then
        PhoneIdentity = Right$(d, 10)
    Else
        PhoneIdentity = d
    End If
End Function

Public Function SheetByName(ByVal nm As String, ByVal createIt As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing And createIt Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = nm
    End If
    Set SheetByName = ws
End Function

Public Function RecordName(ByRef data As Variant, ByVal i As Long) As String
    Dim s As String
    s = JoinWords(CellText(FieldCell(data, i, FieldIndexByFamily("N_GIVEN"))), CellText(FieldCell(data, i, FieldIndexByFamily("N_MIDDLE"))), _
                  CellText(FieldCell(data, i, FieldIndexByFamily("N_FAMILY"))))
    If Len(s) = 0 Then s = CellText(FieldCell(data, i, FieldIndexByFamily("ORG")))
    RecordName = s
End Function

' Distinct phone numbers of a record as written in the sheet, comma separated
Public Function RecordPhones(ByRef data As Variant, ByVal i As Long) As String
    Dim f As Long, v As String, ids As String, id As String, res As String
    ids = "|"
    For f = 1 To gFldCount
        If gFldCol(f) > 0 And gFldProp(f) = "TEL" Then
            v = CellText(data(i, gFldCol(f)))
            If Len(v) > 0 Then
                id = PhoneIdentity(v)
                If InStr(ids, "|" & id & "|") = 0 Then
                    ids = ids & id & "|"
                    If Len(res) > 0 Then res = res & ", "
                    res = res & v
                End If
            End If
        End If
    Next f
    RecordPhones = res
End Function

Private Function RecordLabels(ByRef data As Variant, ByVal i As Long) As String
    Dim f As Long
    For f = 1 To gFldCount
        If gFldCol(f) > 0 And gFldProp(f) = "CATEGORIES" Then
            RecordLabels = CellText(data(i, gFldCol(f)))
            Exit Function
        End If
    Next f
End Function

Private Function FindComp(ByRef p() As Long, ByVal x As Long) As Long
    Do While p(x) <> x
        p(x) = p(p(x))
        x = p(x)
    Loop
    FindComp = x
End Function

' Writes the analysis to the Merge sheet: one line per record of every group that is the same person or needs a look.
' Rows that are copies of another row are not listed; their row numbers are shown on the line of the row they copy.
' Clusters that are linked by a review pair are shown as one group. Returns the number of groups that need review.
Public Function WriteReport() As Long
    Dim comp() As Long, interesting() As Boolean, gid() As Long, gcount() As Long, gstart() As Long, gpos() As Long
    Dim order() As Long, reviewGroup() As Boolean, copyText() As String
    Dim out() As Variant
    Dim ws As Worksheet
    Dim n As Long, i As Long, k As Long, ra As Long, rb As Long, g As Long, ng As Long, lines As Long, cp As Long
    Dim line As Long, reviewGroups As Long, root As Long

    n = gN
    If n < 1 Then Exit Function
    ReDim comp(1 To n)
    ReDim interesting(1 To n)
    ReDim gid(1 To n)
    For i = 1 To n
        comp(i) = i
    Next i

    ' Review links only group components that are NOT already in an auto-merge cluster
    For k = 1 To gRevN
        ra = gRoot(gRevA(k))
        rb = gRoot(gRevB(k))
        If ra >= 1 And ra <= n And rb >= 1 And rb <= n Then
            If gStatus(ra) <> ST_AUTO Or gStatus(rb) <> ST_AUTO Then
                ra = FindComp(comp, ra)
                rb = FindComp(comp, rb)
                If ra <> rb Then comp(ra) = rb
            End If
        End If
    Next k

    ' Mark roots that appear in review links or multi-member clusters as interesting
    For k = 1 To gRevN
        ra = gRoot(gRevA(k))
        rb = gRoot(gRevB(k))
        If ra >= 1 And ra <= n Then If gStatus(ra) <> ST_AUTO Then interesting(ra) = True
        If rb >= 1 And rb <= n Then If gStatus(rb) <> ST_AUTO Then interesting(rb) = True
    Next k
    For i = 1 To n
        If gRoot(i) = i And mCount(i) > 1 Then interesting(i) = True
    Next i

    ' Single pass: assign group IDs and count lines
    ' AUTO clusters first (each is its own group), then REVIEW components
    For i = 1 To n
        If RowShown(i) Then
            root = gRoot(i)
            If gStatus(root) = ST_AUTO And mCount(root) > 1 Then
                If gid(root) = 0 Then
                    ng = ng + 1
                    gid(root) = ng
                End If
                lines = lines + 1
            End If
        End If
    Next i
    For i = 1 To n
        If RowShown(i) Then
            root = gRoot(i)
            If gStatus(root) <> ST_AUTO And interesting(root) Then
                cp = FindComp(comp, root)
                If cp >= 1 And cp <= n Then
                    If gid(cp) = 0 Then
                        ng = ng + 1
                        gid(cp) = ng
                    End If
                    lines = lines + 1
                End If
            End If
        End If
    Next i

    If lines = 0 Then
        WriteReport = 0
        Exit Function
    End If

    ' Which groups need review (contain a non-AUTO cluster or review pair)
    ReDim reviewGroup(0 To ng)
    For i = 1 To n
        If RowShown(i) Then
            root = gRoot(i)
            If gStatus(root) <> ST_AUTO And interesting(root) Then
                cp = FindComp(comp, root)
                If cp >= 1 And cp <= n Then
                    If gid(cp) > 0 Then reviewGroup(gid(cp)) = True
                End If
            End If
        End If
    Next i
    For g = 1 To ng
        If reviewGroup(g) Then reviewGroups = reviewGroups + 1
    Next g

    ' Count members per group, then compute start positions
    ReDim gcount(0 To ng)
    For i = 1 To n
        If RowShown(i) Then
            root = gRoot(i)
            If gStatus(root) = ST_AUTO And mCount(root) > 1 Then
                g = gid(root)
                If g >= 1 And g <= ng Then gcount(g) = gcount(g) + 1
            ElseIf gStatus(root) <> ST_AUTO And interesting(root) Then
                cp = FindComp(comp, root)
                If cp >= 1 And cp <= n Then
                    g = gid(cp)
                    If g >= 1 And g <= ng Then gcount(g) = gcount(g) + 1
                End If
            End If
        End If
    Next i
    ReDim gstart(0 To ng)
    ReDim gpos(0 To ng)
    k = 1
    For g = 1 To ng
        gstart(g) = k
        gpos(g) = k
        k = k + gcount(g)
    Next g
    ' k-1 = sum of gcount = lines (they must match; +1 guard avoids zero-size array)
    ReDim order(1 To k)
    For i = 1 To n
        If RowShown(i) Then
            root = gRoot(i)
            If gStatus(root) = ST_AUTO And mCount(root) > 1 Then
                g = gid(root)
                If g >= 1 And g <= ng Then
                    If gpos(g) <= UBound(order) Then
                        order(gpos(g)) = i
                        gpos(g) = gpos(g) + 1
                    End If
                End If
            ElseIf gStatus(root) <> ST_AUTO And interesting(root) Then
                cp = FindComp(comp, root)
                If cp >= 1 And cp <= n Then
                    g = gid(cp)
                    If g >= 1 And g <= ng Then
                        If gpos(g) <= UBound(order) Then
                            order(gpos(g)) = i
                            gpos(g) = gpos(g) + 1
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ReDim copyText(1 To n)
    For i = 1 To n
        If gIsCopy(i) Then
            If Len(copyText(gRepOf(i))) > 0 Then copyText(gRepOf(i)) = copyText(gRepOf(i)) & ", "
            copyText(gRepOf(i)) = copyText(gRepOf(i)) & CStr(gFirstRow + i - 1)
        End If
    Next i

    ReDim out(1 To lines + 1, 1 To 8)
    out(1, 1) = Txt(S_MRG_COL_GROUP)
    out(1, 2) = Txt(S_MRG_COL_STATUS)
    out(1, 3) = Txt(S_MRG_COL_ROW)
    out(1, 4) = Txt(S_MRG_COL_NAME)
    out(1, 5) = Txt(S_MRG_COL_PHONES)
    out(1, 6) = Txt(S_MRG_COL_GROUPS)
    out(1, 7) = Txt(S_MRG_COL_ACTION)
    out(1, 8) = Txt(S_MRG_COL_COPIES)
    For line = 1 To lines
        i = order(line)
        root = gRoot(i)
        If gStatus(root) = ST_AUTO And mCount(root) > 1 Then
            g = gid(root)
            out(line + 1, 2) = Txt(S_MRG_AUTO)
            If gPrimary(root) = i Then
                out(line + 1, 7) = Txt(S_MRG_KEEP)
            Else
                out(line + 1, 7) = Txt(S_MRG_REMOVE)
            End If
        Else
            cp = FindComp(comp, root)
            g = gid(cp)
            out(line + 1, 2) = Txt(S_MRG_REVIEW)
            out(line + 1, 7) = Txt(S_MRG_UNCHANGED)
        End If
        out(line + 1, 1) = g
        out(line + 1, 3) = gFirstRow + i - 1
        out(line + 1, 4) = RecordName(gData, i)
        out(line + 1, 5) = RecordPhones(gData, i)
        out(line + 1, 6) = RecordLabels(gData, i)
        out(line + 1, 8) = copyText(i)
    Next line

    Set ws = SheetByName("Merge", True)
    ws.Cells.Clear
    ws.Range(ws.Cells(1, 1), ws.Cells(lines + 1, 8)).NumberFormat = "@"
    ws.Range(ws.Cells(1, 1), ws.Cells(lines + 1, 8)).value = out
    ws.rows(1).Font.Bold = True
    ws.Columns("A:H").AutoFit
    WriteReport = reviewGroups
End Function

' Record that is listed on the report: not an empty row and not a copy of another row
Private Function RowShown(ByVal i As Long) As Boolean
    RowShown = Not (gBlank(i) Or gIsCopy(i))
End Function

' Counts of the last plan as text lines
Public Function PlanSummary(ByVal reviewGroups As Long) As String
    PlanSummary = gRecords & " " & Txt(S_MRG_RECORDS) & vbCrLf & _
                  gAutoGroups & " " & Txt(S_MRG_GROUPS_AUTO) & vbCrLf & _
                  gRowsGone & " " & Txt(S_MRG_ROWS_GONE) & " (" & gCopyRows & " " & Txt(S_MRG_IDENT_OF) & ")" & vbCrLf & _
                  reviewGroups & " " & Txt(S_MRG_REVIEW_GROUPS)
End Function

' Looks for records of the same person and writes the analysis to the Merge sheet. The Data sheet is not changed.
Public Sub Merge_Contacts()
    Dim rep As Worksheet
    Dim plan As Long, reviewGroups As Long
    Dim msg As String, errText As String, stage As String

    On Error GoTo ErrHandler

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

    stage = "report"
    StatusText Txt(S_ST_WRITING)
    ProgressUpdate 0.75
    reviewGroups = WriteReport()
    ProgressUpdate 1#
    ProgressEnd

    msg = PlanSummary(reviewGroups)
    ' WriteReport may create the Merge sheet, which Excel activates.  The
    ' analysis is launched from Data, so always leave the user on Data.
    DataSheet().Activate
    If gAutoGroups = 0 And gCopyRows = 0 And reviewGroups = 0 Then
        MsgBox msg & vbCrLf & vbCrLf & Txt(S_MRG_NOTHING), vbInformation
        Exit Sub
    End If
    MsgBox msg & vbCrLf & vbCrLf & Txt(S_MRG_SEE_SHEET) & vbCrLf & Txt(S_MRG_ANALYSIS_ONLY), vbInformation
    Exit Sub

ErrHandler:
    errText = Err.Description
    ProgressEnd
    On Error Resume Next
    DataSheet().Activate
    MsgBox Txt(S_ERROR) & ":" & vbCrLf & errText & " [" & stage & "]", vbExclamation
End Sub
