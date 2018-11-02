Attribute VB_Name = "MVb_Is_Var"
Option Explicit

Function IsAv(A) As Boolean
IsAv = VarType(A) = vbArray + vbVariant
End Function

Function IsAyDic(A As Dictionary) As Boolean
If Not IsSy(A.Keys) Then Exit Function
If Not IsAyOfAy(A.Items) Then Exit Function
IsAyDic = True
End Function

Function IsAyOfAy(A) As Boolean
If Not IsAv(A) Then Exit Function
Dim X
For Each X In AyNz(A)
    If Not IsArray(X) Then Exit Function
Next
IsAyOfAy = True
End Function

Function IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Function

Function IsByt(A) As Boolean
IsByt = VarType(A) = vbByte
End Function
Function IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Function

Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsDte(A) As Boolean
IsDte = VarType(A) = vbDate
End Function

Function IsNe(A, B) As Boolean
IsNe = Not IsEq(A, B)
End Function

Function IsEq(A, B) As Boolean
If VarType(A) <> VarType(B) Then Exit Function
Select Case True
Case IsArray(A): IsEq = IsEqAy(A, B)
Case IsObject(A): IsEq = ObjPtr(A) = ObjPtr(B)
Case Else: IsEq = A = B
End Select
End Function

Function IsEqAy(A, B) As Boolean
Const CSub = "IsEqAy"
If VarType(A) <> VarType(B) Then Exit Function
If Not IsArray(A) Then Er CSub, "[A] is not array", A
If Not IsEqSzAy(A, B) Then Exit Function
Dim J&, X
For Each X In AyNz(A)
    If Not IsEq(X, B(J)) Then Exit Function
    J = J + 1
Next
IsEqAy = True
End Function

Function IsEqDic(A As Dictionary, B As Dictionary) As Boolean
Dim AK(): AK = A.Keys
If Not IsEqAy(AK, B.Keys) Then Exit Function
Dim K
For Each K In AK
    If Not IsEq(A(K), B(K)) Then Exit Function
Next
IsEqDic = True
End Function

Function IsEqSzAy(A, B) As Boolean
IsEqSzAy = Sz(A) = Sz(B)
End Function

Function IsEqVarTy(A, B) As Boolean
IsEqVarTy = VarType(A) = VarType(B)
End Function

Function IsFb(A) As Boolean
IsFb = LCase(FfnExt(A)) = ".accdb"
End Function

Function IsFfnExist(A) As Boolean
IsFfnExist = Fso.FileExists(A)
End Function

Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = AyHas(Av, V)
End Function

Function IsInLikAy(A, LikAy$()) As Boolean
If A = "" Then Exit Function
Dim Lik
For Each Lik In AyNz(LikAy)
    If A Like Lik Then IsInLikAy = True: Exit Function
Next
End Function

Function IsIntAy(A) As Boolean
IsIntAy = VarType(A) = vbArray + vbInteger
End Function

Function IsItr(A) As Boolean
IsItr = TypeName(A) = "Collection"
End Function

Function IsLetter(A$) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsLines(A) As Boolean
IsLines = True
If HasSubStr(A, vbCr) Then Exit Function
If HasSubStr(A, vbLf) Then Exit Function
IsLines = False
End Function

Function IsLinesAy(A) As Boolean
If Not IsSy(A) Then Exit Function
Dim L
For Each L In AyNz(A)
    If IsLines(CStr(L)) Then IsLinesAy = True: Exit Function
Next
End Function

Function IsLng(A) As Boolean
IsLng = VarType(A) = vbLong
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsMdAllRemarked(A As CodeModule) As Boolean
Dim J%, L$
For J = 1 To A.CountOfLines
    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsMdAllRemarked = True
End Function

Function IsNeedQuote(A$) As Boolean
IsNeedQuote = True
If HasSubStr(A, " ") Then Exit Function
If HasSubStr(A, "#") Then Exit Function
If HasSubStr(A, ".") Then Exit Function
IsNeedQuote = False
End Function

Function IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A$) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function

Function IsNmSel(A$, B As WhNm) As Boolean
If IsNothing(B) Then IsNmSel = True: Exit Function
IsNmSel = IsNmSelReExl(A, B.Re, B.ExlAy)
End Function

Function IsNmSelExl(A$, ExlLikAy$()) As Boolean
IsNmSelExl = Not IsInLikAy(A, ExlLikAy)
End Function

Function IsNmSelRe(A$, Re As RegExp) As Boolean
If A = "" Then Exit Function
If IsNothing(Re) Then IsNmSelRe = True: Exit Function
IsNmSelRe = Re.Test(A)
End Function

Function IsNmSelReExl(A$, Re As RegExp, ExlLikAy$()) As Boolean
If Not IsNmSelRe(A, Re) Then Exit Function
If Not IsNmSelExl(A, ExlLikAy) Then Exit Function
IsNmSelReExl = True
End Function

Function IsNoLinMd(A As CodeModule) As Boolean
IsNoLinMd = A.CountOfLines = 0
End Function

Function IsNonBlankStr(V) As Boolean
If Not IsStr(V) Then Exit Function
IsNonBlankStr = V <> ""
End Function

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function

Function IsObjAy(A) As Boolean
IsObjAy = VarType(A) = vbArray + vbObject
End Function

Function IsPrim(A) As Boolean
Select Case VarType(A)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function

Function IsPun(A$) As Boolean
If IsLetter(A) Then Exit Function
If IsDigit(A) Then Exit Function
If A = "_" Then Exit Function
IsPun = True
End Function

Function IsQuoted(A, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If FstChr(A) <> Q1 Then Exit Function
IsQuoted = LasChr(A) = Q2
End Function

Function IsRemarked(Cxt$()) As Boolean
If Sz(Cxt) = 0 Then Exit Function
If Not HasPfx(Cxt(0), "Stop '") Then Exit Function
Dim L
For Each L In Cxt
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRemarked = True
End Function

Function IsSngQRmk(A$) As Boolean
IsSngQRmk = FstChr(LTrim(A)) = "'"
End Function

Function IsSngQuoted(A$) As Boolean
IsSngQuoted = IsQuoted(A, "'")
End Function

Function IsSomething(A) As Boolean
IsSomething = Not IsNothing(A)
End Function

Function IsSqBktNeed(A) As Boolean
If IsSqBktQuoted(A) Then Exit Function
Select Case True
Case HasSpc(A), HasDot(A), HasHyphen(A), HasPound(A): IsSqBktNeed = True
End Select
End Function


Function IsMod(A As CodeModule) As Boolean
IsMod = A.Parent.Type = vbext_ct_StdModule
End Function

Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function

Function IsSy(A) As Boolean
IsSy = IsStrAy(A)
End Function

Function IsStrAy(A) As Boolean
IsStrAy = VarType(A) = vbArray + vbString
End Function

Function IsStrDic(A) As Boolean
Dim D As Dictionary, I
If Not IsDic(A) Then Exit Function
Set D = A
For Each I In D.Keys
    If Not IsStr(D(I)) Then Exit Function
Next
IsStrDic = True
End Function

Function IsSyDic(A) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
Set D = A
For Each I In D.Keys
    V = D(I)
    If Not IsSy(V) Then Exit Function
Next
IsSyDic = True
End Function

Function IsTgl(A) As Boolean
IsTgl = TypeName(A) = "ToggleButton"
End Function


Function IsVbl(A$) As Boolean
Select Case True
Case Not IsStr(A)
Case HasSubStr(A, vbCr)
Case HasSubStr(A, vbLf)
Case Else: IsVbl = True
End Select
End Function

Function IsVblAy(VblAy$()) As Boolean
If Sz(VblAy) = 0 Then IsVblAy = True: Exit Function
Dim Vbl
For Each Vbl In VblAy
    If Not IsVbl(CStr(Vbl)) Then Exit Function
Next
IsVblAy = True
End Function

Function IsVbTyNum(A As VbVarType) As Boolean
Select Case A
Case vbInteger, vbLong, vbDouble, vbSingle, vbDouble: IsVbTyNum = True: Exit Function
End Select
End Function

Function IsVdtLyDicStr(LyDicStr$) As Boolean
If Left(LyDicStr, 3) <> "***" Then Exit Function
Dim I, K$(), Key$
For Each I In SplitCrLf(LyDicStr$)
   If Left(I, 3) = "***" Then
       Key = Mid(I, 4)
       If AyHas(K, Key) Then Exit Function
       Push K, Key
   End If
Next
IsVdtLyDicStr = True
End Function

Private Sub Z_IsVdtLyDicStr()
Ass IsVdtLyDicStr(RplVBar("***ksdf|***ksdf1")) = True
Ass IsVdtLyDicStr(RplVBar("***ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(RplVBar("**ksdf|***ksdf")) = False
Ass IsVdtLyDicStr(RplVBar("***")) = True
Ass IsVdtLyDicStr("**") = False
End Sub

Function IsVdtVbl(Vbl$) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Function
If HasSubStr(Vbl, vbLf) Then Exit Function
IsVdtVbl = True
End Function

Function IsWhiteChr(A) As Boolean
Select Case Left(A, 1)
Case " ", vbCr, vbLf, vbTab: IsWhiteChr = True
End Select
End Function

Private Sub ZZ_IsStrAy()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass IsStrAy(A) = True
Ass IsStrAy(B) = True
Ass IsStrAy(C) = False
Ass IsStrAy(D) = False
End Sub

Private Sub ZIsSy()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass IsSy(A) = True
Ass IsSy(B) = True
Ass IsSy(C) = False
Ass IsSy(D) = False
End Sub

