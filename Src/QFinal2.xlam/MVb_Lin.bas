Attribute VB_Name = "MVb_Lin"
Option Explicit
Function LinCnt&(Lines)
LinCnt = SubStrCnt(Lines, vbCrLf) + 1
End Function

Function LinHasDDRmk(A$) As Boolean
LinHasDDRmk = HasSubStr(A, "--")
End Function

Function LinHasLikItm(A, Lik$, Itm$) As Boolean
Dim L$, I$
AyAsg LinTT(A), L, I
If Not Lik Like L Then Exit Function
LinHasLikItm = I = Itm
End Function

Function LinIsSngTerm(A) As Boolean
LinIsSngTerm = InStr(Trim(A), " ") = 0
End Function

Function LinIsDDLin(A) As Boolean
LinIsDDLin = FstTwoChr(LTrim(A)) = "--"
End Function

Function LinIsDotLin(A) As Boolean
LinIsDotLin = FstChr(A) = "."
End Function

Function LinIsInT1Ay(A, T1Ay$()) As Boolean
LinIsInT1Ay = AyHas(T1Ay, LinT1(A))
End Function

Function LinPfx$(A, ParamArray PfxAp())
Dim Av(): Av = PfxAp
Dim X
For Each X In Av
    If HasPfx(A, X) Then LinPfx = X: Exit Function
Next
End Function

Function LinPfxErMsg$(Lin, Pfx$)
If HasPfx(Lin, Pfx) Then Exit Function
LinPfxErMsg = FmtQQ("First Char must be [?]", Pfx)
End Function
