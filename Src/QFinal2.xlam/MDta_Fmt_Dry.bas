Attribute VB_Name = "MDta_Fmt_Dry"

Function DryFmt(A, Optional MaxColWdt% = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean) As String()
If Sz(A) = 0 Then Exit Function
Dim A1(): A1 = ZCvAllCellToStr(A, ShwZer) ' Convert each cell in Dry-A into string
Dim W%(): W = DryWdtAy(A1, MaxColWdt)
Dim Hdr$: Hdr = WdtAyHdrLin(W)
Dim O$()
    If BrkColIx >= 0 Then
        O = ZInsBrk(A1, BrkColIx, Hdr, W)
    Else
        For Each Dr In A1
            PushI O, DrFmt(Dr, W)
        Next
    End If
    Push O, Hdr
DryFmt = O
End Function

Private Function ZCvAllCellToStr(Dry, ShwZer As Boolean) As Variant()
Dim Dr
For Each Dr In Dry
   Push ZCvAllCellToStr, ZToStr1(Dr, ShwZer)
Next
DryCvCellToStr = O
End Function

Private Function ZToStr1(Dr, ShwZer As Boolean) As String()
Dim I
For Each I In Dr
    PushI ZToStr1, ZToStr2(I, ShwZer)
Next
End Function
Private Function ZToStr2$(V, Optional ShwZer As Boolean) ' Convert V into a string in a cell
'CellStr is a string can be displayed in a cell
Select Case True
Case IsNumeric(V)
    If V = 0 Then
        If ShwZer Then
            ZToStr2 = "0"
        End If
    Else
        ZToStr2 = V
    End If
Case IsEmp(V):
Case IsArray(V)
    Dim N&: N = Sz(V)
    If N = 0 Then
        ZToStr2 = "*[0]"
    Else
        ZToStr2 = "*[" & N & "]" & V(0)
    End If
Case IsObject(V): ZToStr2 = TypeName(V)
Case Else:        DyrFmt1b = V
End Select
End Function

Private Function ZInsBrk(Dry, BrkColIx%, Hdr$, W%()) As String()
Dim Dr, DrIx&, IsBrk As Boolean
Push ZInsBrk, Hdr
For Each Dr In Dry
    IsBrk = ZIsBrk(Dry, DrIx, BrkColIx)
    If IsBrk Then Push O, Hdr
    Push ZInsBrk, DrFmt(Dr, W)
    DrIx = DrIx + 1
Next
End Function

Function ZIsBrk(Dry, DrIx&, BrkColIx%) As Boolean
If Sz(Dry) = 0 Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
ZIsBrk = True
End Function
