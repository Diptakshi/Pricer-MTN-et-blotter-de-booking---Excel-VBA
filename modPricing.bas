Attribute VB_Name = "modPricing"
Option Explicit
' ---------------------------------------------------------------------------
' MTN pricing - independent re-implementation of the workbook formulas.
' Reads the Market sheet, interpolates par swap rates and funding spreads,
' bootstraps annual discount factors, and prices fixed MTNs and FRNs on the
' issuer curve (swap + funding spread). Used as worksheet functions:
'   =MTN_COUPON_PAIR(maturity)            par coupon, %
'   =MTN_PRIX(coupon, maturity)           price of a fixed MTN, % of nominal
'   =MTN_MARGE_FRN(maturity)              FRN margin over Euribor at par, bp
'   =MTN_TOUT_COMPRIS(coupon, maturity)   all-in cost swapped to Euribor, bp
'   =MTN_DV01(nominal, coupon, maturity)  EUR per bp, by bump and reprice
' Diptakshi Sen - personal project, 2026
' ---------------------------------------------------------------------------

Private Const FIRST_ROW As Long = 7
Private Const MAX_T As Long = 30

Private Sub LoadMarket(ten() As Double, par() As Double, spr() As Double, n As Long)
    Dim ws As Worksheet, r As Long
    Set ws = Worksheets("Market")
    n = 0
    r = FIRST_ROW
    Do While Len(CStr(ws.Cells(r, 1).Value)) > 0
        n = n + 1
        r = r + 1
    Loop
    ReDim ten(1 To n): ReDim par(1 To n): ReDim spr(1 To n)
    For r = 1 To n
        ten(r) = CDbl(ws.Cells(FIRST_ROW + r - 1, 1).Value)
        par(r) = CDbl(ws.Cells(FIRST_ROW + r - 1, 2).Value) / 100#
        spr(r) = CDbl(ws.Cells(FIRST_ROW + r - 1, 3).Value) / 10000#
    Next r
End Sub

Private Function Interp(x() As Double, y() As Double, n As Long, t As Double) As Double
    ' linear interpolation, flat beyond the first and last tenors
    Dim i As Long
    If t <= x(1) Then Interp = y(1): Exit Function
    If t >= x(n) Then Interp = y(n): Exit Function
    For i = 1 To n - 1
        If t >= x(i) And t <= x(i + 1) Then
            Interp = y(i) + (t - x(i)) / (x(i + 1) - x(i)) * (y(i + 1) - y(i))
            Exit Function
        End If
    Next i
End Function

Private Sub BuildCurve(shiftBp As Double, dfr() As Double, fwd() As Double)
    ' annual par-swap bootstrap; shiftBp moves every par rate in parallel
    Dim ten() As Double, par() As Double, spr() As Double, n As Long
    Dim df(0 To MAX_T) As Double, t As Long, s As Double, annuity As Double
    LoadMarket ten, par, spr, n
    ReDim dfr(1 To MAX_T): ReDim fwd(1 To MAX_T)
    df(0) = 1#
    annuity = 0#
    For t = 1 To MAX_T
        s = Interp(ten, par, n, CDbl(t)) + shiftBp / 10000#
        df(t) = (1# - s * annuity) / (1# + s)
        annuity = annuity + df(t)
        fwd(t) = df(t - 1) / df(t) - 1#
        dfr(t) = df(t) * Exp(-Interp(ten, spr, n, CDbl(t)) * t)
    Next t
End Sub

Private Function CheckMaturity(maturity As Variant) As Long
    If Not IsNumeric(maturity) Then Err.Raise 5
    If maturity < 1 Or maturity > MAX_T Or maturity <> Int(maturity) Then Err.Raise 5
    CheckMaturity = CLng(maturity)
End Function

Private Function PriceOnCurve(coupon As Double, m As Long, dfr() As Double) As Double
    Dim t As Long, pv As Double
    For t = 1 To m
        pv = pv + coupon / 100# * dfr(t)
    Next t
    PriceOnCurve = (pv + dfr(m)) * 100#
End Function

Public Function MTN_COUPON_PAIR(maturity As Variant) As Variant
    Application.Volatile
    Dim dfr() As Double, fwd() As Double, m As Long, t As Long, ann As Double
    On Error GoTo bad
    m = CheckMaturity(maturity)
    BuildCurve 0, dfr, fwd
    For t = 1 To m: ann = ann + dfr(t): Next t
    MTN_COUPON_PAIR = (1# - dfr(m)) / ann * 100#
    Exit Function
bad:
    MTN_COUPON_PAIR = CVErr(xlErrValue)
End Function

Public Function MTN_PRIX(coupon As Variant, maturity As Variant) As Variant
    Application.Volatile
    Dim dfr() As Double, fwd() As Double, m As Long
    On Error GoTo bad
    m = CheckMaturity(maturity)
    BuildCurve 0, dfr, fwd
    MTN_PRIX = PriceOnCurve(CDbl(coupon), m, dfr)
    Exit Function
bad:
    MTN_PRIX = CVErr(xlErrValue)
End Function

Public Function MTN_MARGE_FRN(maturity As Variant) As Variant
    ' margin m over Euribor such that sum (fwd_t + m) DFr_t + DFr_T = 1
    Application.Volatile
    Dim dfr() As Double, fwd() As Double, m As Long, t As Long, ann As Double, flt As Double
    On Error GoTo bad
    m = CheckMaturity(maturity)
    BuildCurve 0, dfr, fwd
    For t = 1 To m
        ann = ann + dfr(t)
        flt = flt + fwd(t) * dfr(t)
    Next t
    MTN_MARGE_FRN = (1# - dfr(m) - flt) / ann * 10000#
    Exit Function
bad:
    MTN_MARGE_FRN = CVErr(xlErrValue)
End Function

Public Function MTN_TOUT_COMPRIS(coupon As Variant, maturity As Variant) As Variant
    ' fixed coupon swapped to Euribor: Euribor + (coupon - par swap rate)
    Application.Volatile
    Dim ten() As Double, par() As Double, spr() As Double, n As Long, m As Long
    On Error GoTo bad
    m = CheckMaturity(maturity)
    LoadMarket ten, par, spr, n
    MTN_TOUT_COMPRIS = (CDbl(coupon) / 100# - Interp(ten, par, n, CDbl(m))) * 10000#
    Exit Function
bad:
    MTN_TOUT_COMPRIS = CVErr(xlErrValue)
End Function

Public Function MTN_DV01(nominal As Variant, coupon As Variant, maturity As Variant) As Variant
    ' re-bootstrap the curve with every par rate +/- 1 bp and reprice
    Application.Volatile
    Dim dfr() As Double, fwd() As Double, m As Long, up As Double, dn As Double
    On Error GoTo bad
    m = CheckMaturity(maturity)
    BuildCurve 1, dfr, fwd
    up = PriceOnCurve(CDbl(coupon), m, dfr)
    BuildCurve -1, dfr, fwd
    dn = PriceOnCurve(CDbl(coupon), m, dfr)
    MTN_DV01 = (dn - up) / 2# / 100# * CDbl(nominal)
    Exit Function
bad:
    MTN_DV01 = CVErr(xlErrValue)
End Function
