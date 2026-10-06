Attribute VB_Name = "modImportApp"
Option Explicit

' =====================================================================
' Imports the CSV exported from the app (Account > Export to Excel)
' into the monthly journal sheets "1".."12". Run: ImportAppEntries
' - writes only the INPUT columns (A..J, L, N, O). Columns K and M stay formulas.
' - an entry number that already exists in the workbook is skipped (no duplicates)
' - an account name that is not in the chart (row 7 of sheet "1") is skipped and listed
' - result is listed in the sheet "Import_Report"
' =====================================================================

Private Const FIRST_ROW As Long = 9
Private Const SCAN_ROWS As Long = 1100

Private rpt As Worksheet
Private rptRow As Long

Public Sub ImportAppEntries()
    Dim fn As Variant, txt As String, recs As Collection
    Dim hdr As Object, existing As Object, nextRow As Object, capRow As Object, accts As Object
    Dim ws As Worksheet, nm As Variant, rec As Variant
    Dim i As Long, c As Long, r As Long
    Dim noS As String, dS As String, dr As String, cr As String, sh As String
    Dim d As Date, amt As Double, qty As Double, dsc As Double, prc As Double
    Dim imported As Long, dup As Long, skipped As Long
    Dim calcMode As Long

    fn = Application.GetOpenFilename("CSV (*.csv),*.csv", , U("0627 062E 062A 0627 0631 0020 0645 0644 0641 0020 0627 0644 0642 064A 0648 062F 0020 0028 0043 0053 0056 0029 0020 0627 0644 0644 064A 0020 0637 0644 0639 0020 0645 0646 0020 0627 0644 062A 0637 0628 064A 0642"))
    If VarType(fn) = vbBoolean Then Exit Sub

    On Error GoTo Fail
    txt = ReadUtf8(CStr(fn))
    Set recs = ParseCsv(txt)
    If recs.Count < 2 Then MsgBox U("0627 0644 0645 0644 0641 0020 0645 0627 0020 0641 064A 0647 0648 0634 0020 0642 064A 0648 062F"), vbExclamation: Exit Sub

    Set hdr = CreateObject("Scripting.Dictionary")
    rec = recs(1)
    For i = 0 To UBound(rec)
        hdr(LCase(Trim(CStr(rec(i))))) = i
    Next i
    For Each nm In Array("no", "date", "debit_acct", "credit_acct", "amount")
        If Not hdr.Exists(nm) Then
            MsgBox U("0627 0644 0645 0644 0641 0020 0645 0634 0020 0628 0627 0644 0634 0643 0644 0020 0627 0644 0645 0637 0644 0648 0628 002E 0020 0639 0645 0648 062F 0020 0646 0627 0642 0635 003A 0020") & nm, vbExclamation
            Exit Sub
        End If
    Next nm

    Set accts = CreateObject("Scripting.Dictionary")
    For c = 16 To 60 Step 2
        If Len(Trim(Txt(ThisWorkbook.Worksheets("1").Cells(7, c).Value))) > 0 Then accts(Trim(Txt(ThisWorkbook.Worksheets("1").Cells(7, c).Value))) = 1
    Next c

    Set existing = CreateObject("Scripting.Dictionary")
    Set nextRow = CreateObject("Scripting.Dictionary")
    Set capRow = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Worksheets
        If IsNumeric(ws.Name) Then
            If CLng(ws.Name) >= 0 And CLng(ws.Name) <= 13 Then ScanSheet ws, existing, nextRow, capRow
        End If
    Next ws

    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    PrepareReport

    For i = 2 To recs.Count
        rec = recs(i)
        noS = Trim(Fld(rec, hdr, "no"))
        dS = Trim(Fld(rec, hdr, "date"))
        dr = Trim(Fld(rec, hdr, "debit_acct"))
        cr = Trim(Fld(rec, hdr, "credit_acct"))
        amt = Val(Fld(rec, hdr, "amount"))
        If noS = "" And dS = "" And dr = "" Then GoTo NextRec

        If Fld(rec, hdr, "opening") = "1" Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0631 0635 064A 062F 0020 0627 0641 062A 062A 0627 062D 064A 0020 0028 0634 064A 062A 0020 0030 0029")
            skipped = skipped + 1: GoTo NextRec
        End If
        If Not IsNumeric(noS) Or Len(noS) = 0 Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0631 0642 0645 0020 0627 0644 0642 064A 062F 0020 063A 0644 0637")
            skipped = skipped + 1: GoTo NextRec
        End If
        If Len(dS) < 10 Or Not IsNumeric(Left(dS, 4)) Or Not IsNumeric(Mid(dS, 6, 2)) Or Not IsNumeric(Mid(dS, 9, 2)) Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0627 0644 062A 0627 0631 064A 062E 0020 063A 0644 0637")
            skipped = skipped + 1: GoTo NextRec
        End If
        d = DateSerial(CInt(Left(dS, 4)), CInt(Mid(dS, 6, 2)), CInt(Mid(dS, 9, 2)))
        noS = CStr(CLng(noS))

        If existing.Exists(noS) Then
            AddLine U("0645 0643 0631 0631"), noS, dS, U("0645 0648 062C 0648 062F 0020 0628 0627 0644 0641 0639 0644 0020 0641 064A 0020 0634 064A 062A 0020") & existing(noS)
            dup = dup + 1: GoTo NextRec
        End If
        If Not accts.Exists(dr) Or Not accts.Exists(cr) Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("062D 0633 0627 0628 0020 0645 0634 0020 0645 0648 062C 0648 062F 0020 0641 064A 0020 0627 0644 062F 0644 064A 0644 003A 0020") & dr & " / " & cr
            skipped = skipped + 1: GoTo NextRec
        End If
        sh = CStr(Month(d))
        If Not SheetExists(sh) Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0645 0641 064A 0634 0020 0634 064A 062A 0020 0644 0644 0634 0647 0631 0020 062F 0647")
            skipped = skipped + 1: GoTo NextRec
        End If
        If Not nextRow.Exists(sh) Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0627 0644 0634 064A 062A 0020 0645 0634 0020 0645 062A 062C 0647 0632")
            skipped = skipped + 1: GoTo NextRec
        End If
        r = nextRow(sh)
        If r > capRow(sh) Then
            AddLine U("062A 062E 0637 0651 064A"), noS, dS, U("0634 064A 062A 0020 0627 0644 0634 0647 0631 0020 0645 0645 062A 0644 0626")
            skipped = skipped + 1: GoTo NextRec
        End If

        Set ws = ThisWorkbook.Worksheets(sh)
        qty = Val(Fld(rec, hdr, "qty"))
        dsc = Val(Fld(rec, hdr, "disc"))
        prc = Val(Fld(rec, hdr, "price"))
        If Len(Fld(rec, hdr, "party_debit")) > 0 Then ws.Cells(r, 1).Value = Fld(rec, hdr, "party_debit")
        If Len(Fld(rec, hdr, "party_credit")) > 0 Then ws.Cells(r, 2).Value = Fld(rec, hdr, "party_credit")
        ws.Cells(r, 3).Value = d
        ws.Cells(r, 4).Value = CLng(noS)
        ws.Cells(r, 5).Value = Fld(rec, hdr, "stmt")
        ws.Cells(r, 6).Value = dr
        ws.Cells(r, 7).Value = cr
        If Len(Fld(rec, hdr, "product")) > 0 Then ws.Cells(r, 8).Value = Fld(rec, hdr, "product")
        If qty <> 0 Then ws.Cells(r, 9).Value = qty
        If dsc <> 0 Then ws.Cells(r, 10).Value = dsc
        If prc <> 0 Then ws.Cells(r, 12).Value = prc
        ws.Cells(r, 14).Value = amt
        ws.Cells(r, 15).Value = amt

        existing(noS) = sh
        nextRow(sh) = r + 1
        imported = imported + 1
        AddLine U("062A 0645"), noS, dS, U("0634 064A 062A 0020") & sh & U("060C 0020 0635 0641 0020") & r
NextRec:
    Next i

    Application.Calculation = calcMode
    Application.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    rpt.Columns("A:D").AutoFit
    rpt.Columns("D").ColumnWidth = 60

    MsgBox U("062A 0645 0020 0627 0633 062A 064A 0631 0627 062F 0020") & imported & U("0020 0642 064A 062F 002E") & vbCrLf & U("0645 0643 0631 0631 0020 0028 0645 0648 062C 0648 062F 0020 0642 0628 0644 0020 0643 062F 0647 0029 003A 0020") & dup & vbCrLf & _
           U("0627 062A 062E 0637 0651 0649 0020 0644 0623 0633 0628 0627 0628 0020 062A 0627 0646 064A 0629 003A 0020") & skipped & vbCrLf & U("0627 0644 062A 0641 0627 0635 064A 0644 0020 0641 064A 0020 0634 064A 062A 0020 0049 006D 0070 006F 0072 0074 005F 0052 0065 0070 006F 0072 0074"), vbInformation + vbMsgBoxRtlReading + vbMsgBoxRight
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox U("062D 0635 0644 0020 062E 0637 0623 003A 0020") & Err.Description, vbCritical
End Sub

' ---------------------------- helpers ----------------------------

Private Sub ScanSheet(ws As Worksheet, existing As Object, nextRow As Object, capRow As Object)
    Dim a As Variant, i As Long, k As String, lastUsed As Long
    a = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + SCAN_ROWS, 15)).Value
    lastUsed = FIRST_ROW - 1
    For i = 1 To UBound(a, 1)
        If Len(Txt(a(i, 6))) > 0 Or Len(Txt(a(i, 7))) > 0 Then
            k = Replace(Trim(Txt(a(i, 4))), "*", "")
            If Len(k) > 0 Then existing(k) = ws.Name
            lastUsed = FIRST_ROW + i - 1
        End If
    Next i
    nextRow(ws.Name) = lastUsed + 1
    capRow(ws.Name) = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
End Sub

Private Function SheetExists(nm As String) As Boolean
    Dim s As Worksheet
    For Each s In ThisWorkbook.Worksheets
        If s.Name = nm Then SheetExists = True: Exit Function
    Next s
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Then Txt = "": Exit Function
    If IsNull(v) Then Txt = "": Exit Function
    Txt = CStr(v)
End Function

Private Function Fld(rec As Variant, hdr As Object, nm As String) As String
    If Not hdr.Exists(nm) Then Fld = "": Exit Function
    If hdr(nm) > UBound(rec) Then Fld = "": Exit Function
    Fld = CStr(rec(hdr(nm)))
End Function

Private Function ReadUtf8(path As String) As String
    Dim st As Object, s As String
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    s = st.ReadText
    st.Close
    If Len(s) > 0 Then
        If AscW(Left(s, 1)) = &HFEFF Or AscW(Left(s, 1)) = -257 Then s = Mid(s, 2)
    End If
    ReadUtf8 = s
End Function

Private Function ParseCsv(ByVal s As String) As Collection
    Dim res As New Collection
    Dim f(0 To 60) As String, nF As Long
    Dim i As Long, ch As String, cur As String, inQ As Boolean, n As Long

    n = Len(s)
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If i < n And Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQ = True
                Case ","
                    If nF < 60 Then f(nF) = cur: nF = nF + 1
                    cur = ""
                Case vbCr
                    ' ignored
                Case vbLf
                    AddRec res, f, nF, cur
                    nF = 0: cur = ""
                Case Else
                    cur = cur & ch
            End Select
        End If
        i = i + 1
    Loop
    If nF > 0 Or Len(cur) > 0 Then AddRec res, f, nF, cur
    Set ParseCsv = res
End Function

Private Sub AddRec(res As Collection, f() As String, nF As Long, cur As String)
    Dim v() As Variant, k As Long
    If nF = 0 And Len(cur) = 0 Then Exit Sub
    ReDim v(0 To nF)
    For k = 0 To nF - 1
        v(k) = f(k)
    Next k
    v(nF) = cur
    res.Add v
End Sub

Private Sub PrepareReport()
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Import_Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set rpt = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    rpt.Name = "Import_Report"
    rpt.DisplayRightToLeft = True
    rpt.Range("A1:D1").Value = Array(U("0627 0644 062D 0627 0644 0629"), U("0631 0642 0645 0020 0627 0644 0642 064A 062F"), U("0627 0644 062A 0627 0631 064A 062E"), U("0645 0644 0627 062D 0638 0629"))
    rpt.Range("A1:D1").Font.Bold = True
    rptRow = 1
End Sub

Private Sub AddLine(st As String, no As String, dt As String, note As String)
    rptRow = rptRow + 1
    rpt.Cells(rptRow, 1).Value = st
    rpt.Cells(rptRow, 2).Value = "'" & no
    rpt.Cells(rptRow, 3).Value = "'" & dt
    rpt.Cells(rptRow, 4).Value = note
End Sub

' Builds an Arabic string from hex character codes (safe on any code page)
Private Function U(ByVal h As String) As String
    Dim p() As String, i As Long, s As String
    p = Split(h, " ")
    For i = 0 To UBound(p)
        s = s & ChrW(CLng("&H" & p(i)))
    Next i
    U = s
End Function
