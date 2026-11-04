Attribute VB_Name = "AccountFinder"
Option Explicit
' ---------------------------------------------------------------------------
' Account Finder for Excel (desktop)
' Searches every sheet for AWS account IDs or client names and shows the
' client, team, all of the client's accounts and every related row.
'
' Setup: import this module, run SetupFinder once, save as .xlsm.
' Use:   type in Finder!C3, press Enter (with the optional sheet code) or Ctrl+Shift+F.
' ---------------------------------------------------------------------------

Private Const FINDER As String = "Finder"
Private Const INPUT_CELL As String = "C3"
Private Const OUT_ROW As Long = 7
Private Const OUT_ROWS As Long = 4000

' one entry per data sheet
Private nS As Long
Private shWs() As Worksheet
Private shData() As Variant      ' 2-D string array of cell text
Private shR0() As Long, shC0() As Long, shRows() As Long, shCols() As Long
Private shClientCol() As Long
Private shHeaders() As Variant   ' 1-D array of header names
Private shClient() As Variant    ' 1-D array: client for each row
Private shHow() As Variant       ' 1-D array: how the client was identified
Private shIsData() As Variant    ' 1-D Boolean array: True for data rows
Private idx As Object            ' account ID -> Collection of "sheet|row|col"

Private rxWork As Object, rxHdr As Object, rxClient As Object, rxClientEx As Object
Private rxAcct As Object, rxIdEx As Object, rxProfile As Object, rxNum As Object
Private rxSkip As Object, rxTitle As Object, rxLabel As Object

Private outWs As Worksheet, outR As Long

' ============================== public macros ==============================

Public Sub SetupFinder()
    Dim ws As Worksheet
    Set ws = EnsureFinder()
    BuildLayout ws
    On Error Resume Next
    Application.MacroOptions Macro:="FindAccount", Description:="Search all sheets for an account ID or client", ShortcutKey:="F"
    On Error GoTo 0
    ws.Activate
    ws.Range(INPUT_CELL).Select
    MsgBox "Finder sheet is ready." & vbCrLf & vbCrLf & _
           "Type an account ID or client name in " & INPUT_CELL & " and press Ctrl+Shift+F." & vbCrLf & _
           "Remember to save the workbook as .xlsm.", vbInformation, "Account Finder"
End Sub

Public Sub FindAccount()
    Dim q As String, ids As Collection, it As Variant
    Set outWs = EnsureFinder()
    q = Trim$(CellStr(outWs.Range(INPUT_CELL).Value))
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ClearOut
    If Len(q) = 0 Then
        Emit "", "Type an account ID, several IDs, or a client name in " & INPUT_CELL & ".", "src"
    Else
        AnalyzeAll
        Set ids = IdsIn(q)
        If ids.Count > 0 Then
            For Each it In ids
                ShowId CStr(it)
            Next
        Else
            ShowName q
        End If
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    outWs.Activate
    outWs.Range(INPUT_CELL).Select
End Sub

Public Sub ClearFinder()
    Set outWs = EnsureFinder()
    Application.EnableEvents = False
    outWs.Range(INPUT_CELL).ClearContents
    ClearOut
    Application.EnableEvents = True
    outWs.Range(INPUT_CELL).Select
End Sub

' ============================== reading sheets =============================

Private Sub InitRx()
    If Not rxHdr Is Nothing Then Exit Sub
    Set rxWork = NewRx("x")
    Set rxHdr = NewRx("(^|\b)(s\.?\s?no|sr\.?|#|no\.?|client|customer|company|pod|lead|pm|project manager|manager|engineer|primary|secondary|delivery|account|acct|id|name|alias|owner|email|contact|poc|region|status|tier|type|environment|payer|support|plan|notes?|remarks|date|team|ou)(\b|$)")
    Set rxClient = NewRx("client|customer|company|organi[sz]ation|end.?user|tenant|business")
    Set rxClientEx = NewRx("\bid\b|number|#|email|contact|poc|manager|owner|lead|engineer")
    Set rxAcct = NewRx("account.?name|acct.?name|alias|environment|^name$|account$")
    Set rxIdEx = NewRx("\bid\b|number|#")
    Set rxProfile = NewRx("pod|lead|\bpm\b|project manager|delivery|manager|engineer|primary|secondary|owner|\bam\b|\btam\b|\bcsm\b|tier|status|plan|support|region|segment|contact|poc|email")
    Set rxNum = NewRx("^[\d\s.,\-/:%$]+$")
    Set rxSkip = NewRx("^(total|sum|note|notes|n/a|tbd|-)\b")
    Set rxTitle = NewRx("\b(accounts|list|details|summary|report|tracker|sheet)\b")
    Set rxLabel = NewRx("^(client|customer|company)\s*(name)?\s*[:\-]\s*")
End Sub

Private Function NewRx(ByVal pattern As String) As Object
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = pattern
    r.IgnoreCase = True
    r.Global = True
    Set NewRx = r
End Function

Private Sub AnalyzeAll()
    Dim ws As Worksheet, n As Long
    InitRx
    Set idx = CreateObject("Scripting.Dictionary")
    n = ThisWorkbook.Worksheets.Count
    ReDim shWs(1 To n): ReDim shData(1 To n): ReDim shR0(1 To n): ReDim shC0(1 To n)
    ReDim shRows(1 To n): ReDim shCols(1 To n): ReDim shClientCol(1 To n)
    ReDim shHeaders(1 To n): ReDim shClient(1 To n): ReDim shHow(1 To n): ReDim shIsData(1 To n)
    nS = 0
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> FINDER Then
            If Application.WorksheetFunction.CountA(ws.UsedRange) > 0 Then
                nS = nS + 1
                AnalyzeSheet nS, ws
            End If
        End If
    Next
End Sub

' Works out the header row, client heading rows and which client each row belongs to.
Private Sub AnalyzeSheet(ByVal k As Long, ws As Worksheet)
    Dim ur As Range, v As Variant, nR As Long, nC As Long, r As Long, c As Long
    Dim s() As String, hdr() As String, cli() As String, how() As String, isD() As Boolean
    Dim best As Long, bestScore As Long, bestKw As Long, sc As Long, kw As Long, txt As Long
    Dim hasId As Boolean, clientCol As Long, grp As String, grpRow As Long, lastCli As String
    Dim ne As Long, firstVal As String, hdrHits As Long, lbl As String, isTitle As Boolean
    Dim ids As Collection, it As Variant

    Set ur = ws.UsedRange
    If ur.Cells.Count = 1 Then
        ReDim v(1 To 1, 1 To 1): v(1, 1) = ur.Value
    Else
        v = ur.Value
    End If
    nR = UBound(v, 1): nC = UBound(v, 2)
    ReDim s(1 To nR, 1 To nC)
    For r = 1 To nR
        For c = 1 To nC
            s(r, c) = CellStr(v(r, c))
        Next c
    Next r

    ' header = early row that reads most like column labels and holds no account ID
    For r = 1 To IIf(nR < 25, nR, 25)
        hasId = False: kw = 0: txt = 0
        For c = 1 To nC
            If Len(s(r, c)) > 0 Then
                If IdsIn(s(r, c)).Count > 0 Then hasId = True: Exit For
                If Not rxNum.Test(s(r, c)) Then
                    txt = txt + 1
                    If Len(s(r, c)) <= 40 Then If rxHdr.Test(s(r, c)) Then kw = kw + 1
                End If
            End If
        Next c
        If Not hasId Then
            sc = kw * 3 + txt
            If sc > bestScore Then bestScore = sc: best = r: bestKw = kw
        End If
    Next r
    If Not (bestKw >= 1 Or bestScore >= 2) Then best = 0

    ReDim hdr(1 To nC)
    For c = 1 To nC
        If best > 0 Then hdr(c) = s(best, c)
        If Len(hdr(c)) = 0 Then hdr(c) = "Column " & ColLetter(ur.Column + c - 1)
    Next c
    For c = 1 To nC
        If Left$(hdr(c), 7) <> "Column " Then
            If rxClient.Test(hdr(c)) And Not rxClientEx.Test(hdr(c)) Then clientCol = c: Exit For
        End If
    Next c

    ReDim cli(1 To nR): ReDim how(1 To nR): ReDim isD(1 To nR)
    For r = 1 To nR
        ne = 0: hdrHits = 0: hasId = False: firstVal = ""
        For c = 1 To nC
            If Len(s(r, c)) > 0 Then
                ne = ne + 1
                If Len(firstVal) = 0 Then firstVal = s(r, c)
                If best > 0 Then If LCase$(s(r, c)) = LCase$(hdr(c)) Then hdrHits = hdrHits + 1
                If Not hasId Then If IdsIn(s(r, c)).Count > 0 Then hasId = True
            End If
        Next c

        If ne = 0 Or r = best Or (hdrHits >= 2 And hdrHits >= ne - 1) Then
            ' blank row, the header row, or a repeated header row
        ElseIf Not hasId And ne = 1 And nC >= 2 And Not rxNum.Test(firstVal) And Len(firstVal) <= 90 And Not rxSkip.Test(firstVal) Then
            ' a row holding only a name = client heading for the rows below it
            lbl = CleanLabel(firstVal)
            isTitle = False
            If best > 0 And r < best Then
                If InStr(Norm(lbl), Norm(ws.Name)) > 0 Or InStr(Norm(ws.Name), Norm(lbl)) > 0 Or rxTitle.Test(lbl) Then isTitle = True
            End If
            If Not isTitle Then grp = lbl: grpRow = ur.Row + r - 1: lastCli = ""
        ElseIf r > best Then
            isD(r) = True
            If clientCol > 0 Then
                If Len(s(r, clientCol)) > 0 Then
                    lastCli = s(r, clientCol): cli(r) = lastCli: how(r) = """" & hdr(clientCol) & """ column"
                ElseIf Len(grp) > 0 Then
                    cli(r) = grp: how(r) = "client heading in row " & grpRow
                ElseIf Len(lastCli) > 0 Then
                    cli(r) = lastCli: how(r) = """" & hdr(clientCol) & """ column (from the row above)"
                End If
            ElseIf Len(grp) > 0 Then
                cli(r) = grp: how(r) = "client heading in row " & grpRow
            End If
            For c = 1 To nC
                If Len(s(r, c)) >= 9 Then
                    Set ids = IdsIn(s(r, c))
                    For Each it In ids
                        If Not idx.Exists(it) Then idx.Add it, New Collection
                        AddHit idx(it), k & "|" & r & "|" & c
                    Next
                End If
            Next c
        End If
    Next r

    Set shWs(k) = ws
    shData(k) = s: shR0(k) = ur.Row: shC0(k) = ur.Column: shRows(k) = nR: shCols(k) = nC
    shClientCol(k) = clientCol: shHeaders(k) = hdr: shClient(k) = cli: shHow(k) = how: shIsData(k) = isD
End Sub

Private Sub AddHit(col As Collection, ByVal v As String)
    On Error Resume Next
    col.Add v, Left$(v, InStrRev(v, "|") - 1)   ' one hit per sheet row
    On Error GoTo 0
End Sub

' ================================ helpers ==================================

Private Function CellStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            CellStr = Format$(v, "yyyy-mm-dd")
        Case vbDouble, vbCurrency, vbLong, vbInteger, vbSingle, vbDecimal
            If v = Int(v) And Abs(v) < 1E+15 Then CellStr = Format$(v, "0") Else CellStr = CStr(v)
        Case Else
            CellStr = Trim$(CStr(v))
    End Select
End Function

' All 12-digit AWS account IDs in a piece of text (handles 1234-5678-9012 and lost leading zeros)
Private Function IdsIn(ByVal s As String) As Collection
    Dim out As New Collection, m As Object, t As String
    InitRx
    If Len(s) >= 9 Then
        rxWork.Pattern = "(\d{4})[-\s](\d{4})[-\s](\d{4})"
        t = rxWork.Replace(s, "$1$2$3")
        rxWork.Pattern = "(^|\D)(\d{12})(?!\d)"
        On Error Resume Next
        For Each m In rxWork.Execute(t)
            out.Add CStr(m.SubMatches(1)), CStr(m.SubMatches(1))
        Next
        On Error GoTo 0
        If out.Count = 0 Then
            t = Trim$(s)
            If t Like "#########" Or t Like "##########" Or t Like "###########" Then out.Add Right$("000" & t, 12)
        End If
    End If
    Set IdsIn = out
End Function

Private Function Norm(ByVal s As String) As String
    InitRx
    s = LCase$(Replace(s, "&", "and"))
    rxWork.Pattern = "\b(pvt|private|ltd|limited|llc|inc|corp|co|plc|gmbh)\b"
    s = rxWork.Replace(s, "")
    rxWork.Pattern = "[^a-z0-9]"
    Norm = rxWork.Replace(s, "")
End Function

Private Function CleanLabel(ByVal s As String) As String
    s = rxLabel.Replace(s, "")
    Do While Len(s) > 0 And (Right$(s, 1) = ":" Or Right$(s, 1) = " ")
        s = Left$(s, Len(s) - 1)
    Loop
    CleanLabel = Trim$(s)
End Function

Private Function ColLetter(ByVal n As Long) As String
    ColLetter = Split(ThisWorkbook.Worksheets(1).Cells(1, n).Address(True, False), "$")(0)
End Function

Private Function RowNum(ByVal k As Long, ByVal r As Long) As Long
    RowNum = shR0(k) + r - 1
End Function

Private Function AcctName(ByVal k As Long, ByVal r As Long, ByVal skipC As Long) As String
    Dim c As Long, h As String
    For c = 1 To shCols(k)
        h = shHeaders(k)(c)
        If c <> skipC And c <> shClientCol(k) Then
            If rxAcct.Test(h) And Not rxIdEx.Test(h) Then AcctName = shData(k)(r, c): Exit Function
        End If
    Next c
End Function

Private Function RowId(ByVal k As Long, ByVal r As Long, ByRef col As Long) As String
    Dim c As Long, ids As Collection
    For c = 1 To shCols(k)
        If Len(shData(k)(r, c)) >= 9 Then
            Set ids = IdsIn(shData(k)(r, c))
            If ids.Count > 0 Then RowId = ids(1): col = c: Exit Function
        End If
    Next c
End Function

Private Function Summary(ByVal k As Long, ByVal r As Long) As String
    Dim c As Long, t As String
    For c = 1 To shCols(k)
        If Len(shData(k)(r, c)) > 0 Then
            If Len(t) > 0 Then t = t & "   ·   "
            t = t & shHeaders(k)(c) & ": " & shData(k)(r, c)
        End If
    Next c
    Summary = t
End Function

' Every data row that belongs to or mentions the client
Private Function Related(ByVal client As String) As Collection
    Dim out As New Collection, n As String, m As String, k As Long, r As Long, c As Long, hit As Boolean
    n = Norm(client)
    If Len(n) >= 2 Then
        For k = 1 To nS
            For r = 1 To shRows(k)
                If shIsData(k)(r) Then
                    hit = (Norm(shClient(k)(r)) = n)
                    If Not hit Then
                        For c = 1 To shCols(k)
                            If Len(shData(k)(r, c)) > 0 Then
                                m = Norm(shData(k)(r, c))
                                If m = n Then hit = True: Exit For
                                If Len(n) >= 5 And Len(m) <= Len(n) + 30 Then If InStr(m, n) > 0 Then hit = True: Exit For
                            End If
                        Next c
                    End If
                    If hit Then out.Add k & "|" & r
                End If
            Next r
        Next k
    End If
    Set Related = out
End Function

' ================================ results ==================================

Private Sub ShowId(ByVal id As String)
    Dim hits As Collection, h As Variant, p() As String, k As Long, r As Long, c As Long
    Dim votes As Object, excl As Object, a As Variant, key As Variant, nm As String
    Dim client As String, src As String, others As String, bestN As Long, rel As Collection, title As String

    If Not idx.Exists(id) Then
        Emit "Account " & id, "Not found in any sheet", "h"
        Emit "", "It may be a new account that hasn't been added to the workbook yet.", "src"
        Emit
        Exit Sub
    End If
    Set hits = idx(id)
    Set votes = CreateObject("Scripting.Dictionary")
    Set excl = CreateObject("Scripting.Dictionary")
    For Each h In hits
        p = Split(h, "|"): k = CLng(p(0)): r = CLng(p(1))
        excl(k & "|" & r) = True
        nm = shClient(k)(r)
        If Len(nm) > 0 Then
            If votes.Exists(Norm(nm)) Then a = votes(Norm(nm)) Else a = Array(nm, 0, "")
            a(1) = a(1) + 1
            a(2) = a(2) & IIf(Len(a(2)) > 0, "; ", "") & shWs(k).Name & " (" & shHow(k)(r) & ")"
            votes(Norm(nm)) = a
        End If
    Next
    For Each key In votes.Keys
        a = votes(key)
        If a(1) > bestN Then
            If Len(client) > 0 Then others = others & IIf(Len(others) > 0, "; ", "") & client
            bestN = a(1): client = a(0): src = a(2)
        Else
            others = others & IIf(Len(others) > 0, "; ", "") & a(0)
        End If
    Next

    Emit "Account " & id, "Found in " & hits.Count & " row(s)", "h"
    If Len(client) > 0 Then
        Emit "CLIENT", client, "client"
        Emit "", "Identified from " & src & IIf(Len(others) > 0, "   |   Also labelled as: " & others, ""), "src"
        Emit
        Set rel = Related(client)
        ShowProfile rel
    Else
        Emit "CLIENT", "Not identified", "client"
        Emit "", "No client column or client heading row was found above this account. Check the details below.", "src"
        Emit
    End If

    For Each h In hits
        p = Split(h, "|"): k = CLng(p(0)): r = CLng(p(1)): c = CLng(p(2))
        title = AcctName(k, r, c)
        If Len(title) = 0 Then title = "Account"
        Emit title, "", "sub"
        LinkCell outR - 1, 3, k, r, c, "Open in " & shWs(k).Name & ", row " & RowNum(k, r)
        RowDetails k, r, c
        Emit
    Next

    If Len(client) > 0 Then
        ShowAccounts client, rel, hits, id
        ShowOthers rel, excl
    End If
    Emit
End Sub

Private Sub ShowName(ByVal q As String)
    Dim cands As Object, k As Long, r As Long, c As Long, nm As String, nq As String
    Dim names() As String, i As Long, j As Long, tmp As String, n As Long, chosen As Object
    Dim rel As Collection, cnt As Long, hit As Boolean, it As Variant

    Set cands = CreateObject("Scripting.Dictionary")
    For k = 1 To nS
        For r = 1 To shRows(k)
            nm = shClient(k)(r)
            If Len(nm) > 0 Then
                If InStr(1, nm, q, vbTextCompare) > 0 Then
                    If Not cands.Exists(Norm(nm)) Then cands.Add Norm(nm), nm
                End If
            End If
        Next r
    Next k

    Set chosen = CreateObject("Scripting.Dictionary")
    If cands.Count > 0 Then
        ' exact match first, then shortest names; show up to 3 clients
        nq = Norm(q): n = cands.Count
        ReDim names(1 To n)
        i = 0
        For Each it In cands.Items
            i = i + 1: names(i) = it
        Next
        For i = 1 To n - 1
            For j = i + 1 To n
                If Rank(names(j), nq) < Rank(names(i), nq) Then tmp = names(i): names(i) = names(j): names(j) = tmp
            Next j
        Next i
        For i = 1 To IIf(n < 3, n, 3)
            chosen(Norm(names(i))) = True
            Emit "CLIENT", names(i), "client"
            Set rel = Related(names(i))
            Emit "", "Found in " & rel.Count & " row(s) across the workbook", "src"
            Emit
            ShowProfile rel
            ShowAccounts names(i), rel, Nothing, ""
            ShowOthers rel, Nothing
            Emit
        Next i
    End If

    ' rows that mention the text but belong to another client, e.g. a pod lead or engineer
    For k = 1 To nS
        For r = 1 To shRows(k)
            If shIsData(k)(r) And Not chosen.Exists(Norm(shClient(k)(r))) Then
                hit = False
                For c = 1 To shCols(k)
                    If InStr(1, shData(k)(r, c), q, vbTextCompare) > 0 Then hit = True: Exit For
                Next c
                If hit Then
                    If cnt = 0 Then Emit IIf(chosen.Count > 0, "Other rows that mention """ & q & """ (for example as a pod lead, PM or engineer)", "Rows that mention """ & q & """"), "", "sec"
                    cnt = cnt + 1
                    Emit "", IIf(Len(shClient(k)(r)) > 0, "[" & shClient(k)(r) & "]   ", "") & Summary(k, r)
                    LinkCell outR - 1, 2, k, r, 1, shWs(k).Name & ", row " & RowNum(k, r)
                    If cnt >= 60 Then Emit "", "Showing the first 60 rows.", "src": Exit Sub
                End If
            End If
        Next r
    Next k
    If cnt = 0 And chosen.Count = 0 Then Emit "", "Nothing in the workbook matches """ & q & """.", "src"
End Sub

Private Function Rank(ByVal nm As String, ByVal nq As String) As Long
    Rank = Len(nm) + IIf(Norm(nm) = nq, 0, 1000)
End Function

' Pod, pod lead, PM, engineers and similar, gathered from every row of the client
Private Sub ShowProfile(rel As Collection)
    Dim facts As Object, it As Variant, p() As String, k As Long, r As Long, c As Long
    Dim h As String, v As String, key As Variant, shown As Long
    Set facts = CreateObject("Scripting.Dictionary")
    facts.CompareMode = vbTextCompare
    For Each it In rel
        p = Split(it, "|"): k = CLng(p(0)): r = CLng(p(1))
        For c = 1 To shCols(k)
            v = shData(k)(r, c)
            If Len(v) > 0 And c <> shClientCol(k) Then
                h = shHeaders(k)(c)
                If rxProfile.Test(h) Then If IdsIn(v).Count = 0 Then AddFact facts, h, v
            End If
        Next c
        If InStr(1, shWs(k).Name, "pod", vbTextCompare) > 0 And Left$(shHow(k)(r), 14) = "client heading" Then
            AddFact facts, "Pod", Trim$(Replace(Replace(shWs(k).Name, "Accounts", "", , , vbTextCompare), "Account", "", , , vbTextCompare))
        End If
    Next
    If facts.Count = 0 Then Exit Sub
    Emit "Team & details", "", "sec"
    For Each key In facts.Keys
        Emit CStr(key), Join(facts(key).Keys, ", "), "fact"
        shown = shown + 1
        If shown >= 14 Then Exit For
    Next
    Emit
End Sub

Private Sub AddFact(facts As Object, ByVal h As String, ByVal v As String)
    Dim d As Object
    If Not facts.Exists(h) Then
        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare
        facts.Add h, d
    End If
    If Not facts(h).Exists(v) Then facts(h).Add v, True
End Sub

Private Sub ShowAccounts(ByVal client As String, rel As Collection, hits As Collection, ByVal curId As String)
    Dim acc As Object, it As Variant, a As Variant, key As Variant
    Set acc = CreateObject("Scripting.Dictionary")
    For Each it In rel
        AddAcct acc, CStr(it)
    Next
    If Not hits Is Nothing Then
        For Each it In hits
            AddAcct acc, Left$(it, InStrRev(it, "|") - 1)
        Next
    End If
    Emit "All accounts for " & client & " (" & acc.Count & ")", "", "sec"
    If acc.Count = 0 Then
        Emit "", "No account IDs were found for this client. The pod sheets may spell the name differently.", "src"
        Emit
        Exit Sub
    End If
    Emit "Account name", "Account ID", "th", "Listed in"
    For Each key In acc.Keys
        a = acc(key)
        Emit IIf(Len(a(0)) > 0, a(0), "-"), CStr(key), IIf(CStr(key) = curId, "me", ""), ""
        LinkCell outR - 1, 4, CLng(a(2)), CLng(a(3)), CLng(a(4)), a(1)
    Next
    Emit
End Sub

Private Sub AddAcct(acc As Object, ByVal kr As String)
    Dim p() As String, k As Long, r As Long, col As Long, id As String, a As Variant
    p = Split(kr, "|"): k = CLng(p(0)): r = CLng(p(1))
    id = RowId(k, r, col)
    If Len(id) = 0 Then Exit Sub
    If Not acc.Exists(id) Then
        acc.Add id, Array(AcctName(k, r, col), shWs(k).Name, k, r, col)
    Else
        a = acc(id)
        If Len(a(0)) = 0 Then a(0) = AcctName(k, r, col)
        If InStr(1, ", " & a(1) & ",", ", " & shWs(k).Name & ",", vbTextCompare) = 0 Then a(1) = a(1) & ", " & shWs(k).Name
        acc(id) = a
    End If
End Sub

Private Sub ShowOthers(rel As Collection, excl As Object)
    Dim it As Variant, p() As String, k As Long, r As Long, cnt As Long
    For Each it In rel
        If excl Is Nothing Then
        ElseIf excl.Exists(it) Then
            GoTo nextRow
        End If
        p = Split(it, "|"): k = CLng(p(0)): r = CLng(p(1))
        If cnt = 0 Then Emit "Client rows across the workbook", "", "sec"
        cnt = cnt + 1
        Emit "", Summary(k, r)
        LinkCell outR - 1, 2, k, r, 1, shWs(k).Name & ", row " & RowNum(k, r)
        If cnt >= 80 Then Emit "", "Showing the first 80 rows.", "src": Exit For
nextRow:
    Next
    If cnt > 0 Then Emit
End Sub

Private Sub RowDetails(ByVal k As Long, ByVal r As Long, ByVal hitC As Long)
    Dim c As Long
    For c = 1 To shCols(k)
        If Len(shData(k)(r, c)) > 0 Then Emit shHeaders(k)(c), shData(k)(r, c), IIf(c = hitC, "hit", "")
    Next c
End Sub

' ================================ output ===================================

Private Sub Emit(Optional ByVal a As String = "", Optional ByVal b As String = "", Optional ByVal style As String = "", Optional ByVal d As String = "")
    With outWs
        .Cells(outR, 2).Value = a
        .Cells(outR, 3).Value = b
        If Len(d) > 0 Then .Cells(outR, 4).Value = d
        Select Case style
            Case "h"
                .Cells(outR, 2).Font.Bold = True: .Cells(outR, 2).Font.Size = 14
                .Cells(outR, 3).Font.Color = RGB(111, 104, 90)
            Case "client"
                With .Range(.Cells(outR, 2), .Cells(outR, 4))
                    .Interior.Color = RGB(28, 26, 21)
                    .Font.Color = RGB(255, 253, 248)
                End With
                .Cells(outR, 2).Font.Color = RGB(247, 160, 109)
                .Cells(outR, 3).Font.Size = 18: .Cells(outR, 3).Font.Bold = True
                .Rows(outR).RowHeight = 30
            Case "src"
                .Cells(outR, 3).Font.Italic = True: .Cells(outR, 3).Font.Color = RGB(111, 104, 90)
            Case "sec"
                .Cells(outR, 2).Font.Bold = True: .Cells(outR, 2).Font.Size = 12
                .Cells(outR, 2).Font.Color = RGB(228, 87, 15)
                .Range(.Cells(outR, 2), .Cells(outR, 4)).Borders(xlEdgeBottom).Color = RGB(228, 87, 15)
            Case "sub"
                .Cells(outR, 2).Font.Bold = True: .Cells(outR, 2).Font.Size = 12
                .Range(.Cells(outR, 2), .Cells(outR, 4)).Interior.Color = RGB(253, 230, 214)
            Case "fact"
                .Cells(outR, 2).Font.Color = RGB(111, 104, 90)
                .Cells(outR, 3).Font.Bold = True
            Case "th"
                With .Range(.Cells(outR, 2), .Cells(outR, 4))
                    .Font.Bold = True
                    .Interior.Color = RGB(248, 244, 234)
                End With
            Case "hit"
                .Cells(outR, 3).Interior.Color = RGB(255, 242, 168)
                .Cells(outR, 3).Font.Bold = True
            Case "me"
                .Range(.Cells(outR, 2), .Cells(outR, 4)).Interior.Color = RGB(255, 242, 168)
        End Select
    End With
    outR = outR + 1
End Sub

Private Sub LinkCell(ByVal row As Long, ByVal col As Long, ByVal k As Long, ByVal r As Long, ByVal c As Long, ByVal text As String)
    Dim target As String
    target = "'" & Replace(shWs(k).Name, "'", "''") & "'!" & shWs(k).Cells(RowNum(k, r), shC0(k) + c - 1).Address(False, False)
    outWs.Hyperlinks.Add Anchor:=outWs.Cells(row, col), Address:="", SubAddress:=target, TextToDisplay:=text
End Sub

Private Sub ClearOut()
    With outWs.Range(outWs.Cells(OUT_ROW, 1), outWs.Cells(OUT_ROW + OUT_ROWS, 6))
        .Hyperlinks.Delete
        .Clear
        .RowHeight = outWs.StandardHeight
    End With
    With outWs.Range(outWs.Cells(OUT_ROW, 2), outWs.Cells(OUT_ROW + OUT_ROWS, 4))
        .NumberFormat = "@"            ' keeps 12-digit IDs as text
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    outR = OUT_ROW
End Sub

Private Function EnsureFinder() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FINDER)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = FINDER
        BuildLayout ws
    End If
    Set EnsureFinder = ws
End Function

Private Sub BuildLayout(ws As Worksheet)
    Application.EnableEvents = False
    With ws
        .Range("A1:F6").Clear
        .Columns("A").ColumnWidth = 2
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 75
        .Columns("D").ColumnWidth = 34
        With .Range("B1")
            .Value = "Account Finder"
            .Font.Size = 22: .Font.Bold = True: .Font.Color = RGB(28, 26, 21)
        End With
        With .Range("B2")
            .Value = "Type an account ID, several IDs or a client name in the orange box, then press Enter or Ctrl+Shift+F. To paste a whole alert, paste it into the formula bar."
            .Font.Italic = True: .Font.Color = RGB(111, 104, 90)
        End With
        With .Range("B3")
            .Value = "Search"
            .Font.Bold = True: .Font.Size = 12: .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        End With
        With .Range(INPUT_CELL)
            .NumberFormat = "@"
            .Interior.Color = RGB(253, 230, 214)
            .Font.Name = "Consolas": .Font.Size = 14
            .Borders.Color = RGB(228, 87, 15)
            .Borders.Weight = xlMedium
            .VerticalAlignment = xlCenter
        End With
        .Rows(3).RowHeight = 28
    End With
    On Error Resume Next
    ws.Shapes("btnFindAccount").Delete
    ws.Shapes("btnClearFinder").Delete
    On Error GoTo 0
    AddButton ws, "Find  (Ctrl+Shift+F)", "FindAccount", ws.Range("D3").Left + 6, ws.Range("D3").Top + 2, 150, RGB(228, 87, 15)
    AddButton ws, "Clear", "ClearFinder", ws.Range("D3").Left + 162, ws.Range("D3").Top + 2, 60, RGB(111, 104, 90)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    Application.EnableEvents = True
End Sub

Private Sub AddButton(ws As Worksheet, ByVal cap As String, ByVal macroName As String, ByVal l As Double, ByVal t As Double, ByVal w As Double, ByVal fillColor As Long)
    Dim s As Shape
    Set s = ws.Shapes.AddShape(5, l, t, w, 24)   ' 5 = rounded rectangle
    s.Name = "btn" & macroName
    s.OnAction = macroName
    s.Fill.ForeColor.RGB = fillColor
    s.Line.Visible = 0
    With s.TextFrame2
        .VerticalAnchor = 3                        ' middle
        .TextRange.Text = cap
        .TextRange.Font.Size = 11
        .TextRange.Font.Bold = -1
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = 2   ' centre
    End With
End Sub

' ---------------------------------------------------------------------------
' OPTIONAL: search automatically when you press Enter in C3.
' Right-click the Finder sheet tab > View Code, and paste these 3 lines there:
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     If Not Intersect(Target, Me.Range("C3")) Is Nothing Then FindAccount
' End Sub
' ---------------------------------------------------------------------------
