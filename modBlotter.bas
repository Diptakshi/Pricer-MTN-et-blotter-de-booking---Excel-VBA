Attribute VB_Name = "modBlotter"
Option Explicit
' ---------------------------------------------------------------------------
' Blotter of MTN issues: add the trade priced on the Pricer sheet, run the
' pre-booking controls, export the validated trades to a CSV for booking.
' ---------------------------------------------------------------------------

Private Const HDR_ROW As Long = 4
Private Const MAX_GAP_BP As Double = 5          ' pricing tolerance vs par coupon

Private Function LastRow(ws As Worksheet) As Long
    Dim r As Long
    r = HDR_ROW
    Do While Len(CStr(ws.Cells(r + 1, 1).Value)) > 0
        r = r + 1
    Loop
    LastRow = r
End Function

Public Sub AjouterAuBlotter()
    Dim p As Worksheet, b As Worksheet, r As Long
    Set p = Worksheets("Pricer")
    Set b = Worksheets("Blotter")
    r = LastRow(b) + 1
    b.Cells(r, 1).Value = "MTN-" & Format(Now, "yyyymmdd") & "-" & Format(r - HDR_ROW, "000")
    b.Cells(r, 2).Value = Date
    b.Cells(r, 3).Value = p.Range("B4").Value          ' issuer
    b.Cells(r, 4).Value = p.Range("B5").Value          ' currency
    b.Cells(r, 5).Value = p.Range("B6").Value          ' nominal
    b.Cells(r, 6).Value = p.Range("B7").Value          ' maturity
    b.Cells(r, 7).Value = p.Range("B8").Value          ' coupon
    b.Cells(r, 8).Value = p.Range("B12").Value         ' par coupon
    b.Cells(r, 9).Value = p.Range("B14").Value         ' gap, bp
    b.Cells(r, 10).Value = p.Range("B15").Value        ' all-in, bp
    b.Cells(r, 11).Value = "À valider"
    b.Cells(r, 12).Value = ""
    b.Range("B2").Value = "Trade ajouté : " & b.Cells(r, 1).Value
End Sub

Public Sub ValiderBlotter()
    Dim b As Worksheet, r As Long, last As Long, msg As String, i As Long
    Dim nOK As Long, nKO As Long, ccy As String
    Set b = Worksheets("Blotter")
    last = LastRow(b)
    For r = HDR_ROW + 1 To last
        msg = ""
        For i = HDR_ROW + 1 To r - 1                      ' duplicate trade id
            If b.Cells(i, 1).Value = b.Cells(r, 1).Value Then msg = msg & "ID en double ; "
        Next i
        ccy = UCase(CStr(b.Cells(r, 4).Value))
        If InStr("|EUR|USD|GBP|CHF|JPY|", "|" & ccy & "|") = 0 Or Len(ccy) <> 3 Then msg = msg & "Devise invalide ; "
        If Not IsNumeric(b.Cells(r, 5).Value) Then
            msg = msg & "Nominal manquant ; "
        ElseIf b.Cells(r, 5).Value <= 0 Then
            msg = msg & "Nominal <= 0 ; "
        End If
        If Not IsNumeric(b.Cells(r, 6).Value) Then
            msg = msg & "Maturité manquante ; "
        ElseIf b.Cells(r, 6).Value < 1 Or b.Cells(r, 6).Value > 30 Or b.Cells(r, 6).Value <> Int(b.Cells(r, 6).Value) Then
            msg = msg & "Maturité hors 1-30 ans ; "
        End If
        If Not IsNumeric(b.Cells(r, 7).Value) Or Len(CStr(b.Cells(r, 7).Value)) = 0 Then
            msg = msg & "Coupon manquant ; "
        ElseIf b.Cells(r, 7).Value < 0 Or b.Cells(r, 7).Value > 15 Then
            msg = msg & "Coupon hors bornes ; "
        End If
        If IsNumeric(b.Cells(r, 9).Value) And Len(CStr(b.Cells(r, 9).Value)) > 0 Then
            If Abs(b.Cells(r, 9).Value) > MAX_GAP_BP Then msg = msg & "Écart au pair > " & MAX_GAP_BP & " bp ; "
        End If
        If Len(msg) = 0 Then
            b.Cells(r, 11).Value = "Validé": b.Cells(r, 12).Value = "": nOK = nOK + 1
        Else
            b.Cells(r, 11).Value = "Erreur": b.Cells(r, 12).Value = Left(msg, Len(msg) - 3): nKO = nKO + 1
        End If
    Next r
    b.Range("B2").Value = "Contrôle du " & Format(Now, "dd/mm/yyyy hh:mm") & " : " & nOK & " validé(s), " & nKO & " en erreur"
End Sub

Public Sub ExporterBooking()
    ' validated trades only, semicolon-separated for a French Excel / booking upload
    Dim b As Worksheet, r As Long, f As Integer, path As String, sep As String, n As Long
    Set b = Worksheets("Blotter")
    sep = "\"
    If InStr(ThisWorkbook.Path, "/") > 0 Then sep = "/"
    path = ThisWorkbook.Path & sep & "booking_mtn_" & Format(Now, "yyyymmdd") & "-" & Format(Now, "hhmm") & ".csv"
    f = FreeFile
    Open path For Output As #f
    Print #f, "TradeID;TradeDate;Issuer;Currency;Notional;MaturityYears;CouponPct;AllInBp"
    For r = HDR_ROW + 1 To LastRow(b)
        If b.Cells(r, 11).Value = "Validé" Then
            Print #f, b.Cells(r, 1).Value & ";" & Format(b.Cells(r, 2).Value, "yyyy-mm-dd") & ";" & _
                b.Cells(r, 3).Value & ";" & b.Cells(r, 4).Value & ";" & b.Cells(r, 5).Value & ";" & _
                b.Cells(r, 6).Value & ";" & Format(b.Cells(r, 7).Value, "0.0000") & ";" & Format(b.Cells(r, 10).Value, "0.0")
            n = n + 1
        End If
    Next r
    Close #f
    b.Range("B2").Value = n & " trade(s) exporté(s) : " & path
End Sub
