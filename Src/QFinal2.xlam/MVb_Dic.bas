Attribute VB_Name = "MVb_Dic"
Option Explicit
Sub AssDicHasKy(A As Dictionary, Ky)
Dim K
For Each K In Ky
   If Not A.Exists(K) Then Debug.Print K: Stop
Next
End Sub

Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function

Function DicKySy(A As Dictionary, Ky$()) As String()
Dim K
For Each K In AyNz(Ky)
    PushI DicKySy, A(K)
Next
End Function
Function DicAddAy(A As Dictionary, Dy() As Dictionary) As Dictionary
Set DicAddAy = DicClone(A)
Dim J%
For J = 0 To UB(Dy)
   PushDic DicAddAy, Dy(J)
Next
End Function

Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set DicAddKeyPfx = O
End Function

Function DicAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicAllKeyIsNm = True
End Function

Function DicAllKeyIsStr(A As Dictionary) As Boolean
DicAllKeyIsStr = AyIsAllStr(A.Keys)
End Function

Function DicAllValIsStr(A As Dictionary) As Boolean
DicAllValIsStr = AyIsAllStr(A.Items)
End Function

Function DicAyMge(A() As Dictionary) As Dictionary
'Assume there is no duplicated key in each of the dic in A()
Dim I
For Each I In AyNz(A)
    PushDic DicAyMge, CvDic(I)
Next
End Function

Function DicAyAdd(A() As Dictionary) As Dictionary
Dim O As New Dictionary, D
For Each D In A
    PushDic O, CvDic(D)
Next
Set DicAyAdd = O
End Function

Function DicAyDr(DicAy, K) As Variant()
Dim U%: U = UB(DicAy)
Dim O()
ReDim O(U + 1)
Dim I, Dic As Dictionary, J%
J = 1
O(0) = K
For Each I In DicAy
   Set Dic = I
   If Dic.Exists(K) Then O(J) = Dic(K)
   J = J + 1
Next
DicAyDr = O
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set DicClone = O
End Function


Function DicDr(A As Dictionary, Ky$()) As Variant()
Dim O(), I, J&
ReDim O(UB(Ky))
For Each I In Ky
    If A.Exists(I) Then
        O(J) = A(I)
    End If
    J = J + 1
Next
DicDr = O
End Function


Function DicHasAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicHasAllKeyIsNm = True
End Function

Function DicHasAllValIsStr(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsStr(A(K)) Then Exit Function
Next
DicHasAllValIsStr = True
End Function

Function DicHasBlankKey(A As Dictionary) As Boolean
If A.Count = 0 Then Exit Function
Dim K
For Each K In A.Keys
   If Trim(K) = "" Then DicHasBlankKey = True: Exit Function
Next
End Function

Function DicHasKeySsl(A As Dictionary, KeySsl) As Boolean
DicHasKeySsl = A.Exists(SslSy(KeySsl))
End Function

Function DicHasKy(A As Dictionary, Ky) As Boolean
Ass IsArray(Ky)
If Sz(Ky) = 0 Then Stop
Dim K
For Each K In Ky
   If Not A.Exists(K) Then
       Debug.Print FmtQQ("Dix.HasKy: Key(?) is missing", K)
       Exit Function
   End If
Next
DicHasKy = True
End Function

Function DicHasStrKy(A As Dictionary) As Boolean
DicHasStrKy = ItrPredAllTrue(A.Keys, "IsStr")
End Function

Function DicHasStrKy1(A As Dictionary) As Boolean
Dim I
For Each I In A.Keys
    If Not IsStr(I) Then Exit Function
Next
DicHasStrKy1 = True
End Function

Function DicIntersect(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Then GoTo X
If B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicIntersect = O
End Function

Function DicIsEmp(A As Dictionary) As Boolean
DicIsEmp = A.Count = 0
End Function

Function DicIsEq(A As Dictionary, B As Dictionary) As Boolean
If A.Count <> B.Count Then Exit Function
If A.Count = 0 Then Exit Function
Dim K1, K2
K1 = AyQSrt(A.Keys)
K2 = AyQSrt(B.Keys)
If IsEqAy(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> A(K) Then Exit Function
Next
DicIsEq = True
End Function

Function DicKeySy(A As Dictionary) As String()
DicKeySy = AySy(A.Keys)
End Function

Function DicLines(A As Dictionary) As String
DicLines = JnCrLf(DicFmt(A))
End Function

Function DicLines_Dic(A$, Optional JnSep$ = vbCrLf) As Dictionary
Set DicLines_Dic = DicLyDic(SplitLines(A), JnSep)
End Function

Function DicLyDic(A$(), Optional JnSep$ = vbCrLf) As Dictionary
Dim O As New Dictionary
Dim A1$(): A1 = AyRmvEmpEleAtEnd(A)
If Sz(A) = 0 Then Set DicLyDic = O: Exit Function
Dim I, T1$, Rst$
For Each I In A
    LinAsgTRst I, T1, Rst
    If O.Exists(T1) Then
        If FstChr(Rst) = "~" Then Rst = RplFstChr(Rst, " ")
        O(T1) = O(T1) & JnSep & Rst
    Else
        O.Add T1, Rst
    End If
 Next
Set DicLyDic = O
End Function
Function CvDic(A) As Dictionary
Set CvDic = A
End Function

Function DicLblLy(A As Dictionary, Lbl$) As String()
PushI DicLblLy, Lbl
PushI DicLblLy, vbTab & "Count=" & A.Count
PushIAy DicLblLy, AyAddPfx(DicFmt(A, InclValTy:=True), vbTab)
End Function
Function DicLy1(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim Key: Key = A.Keys
Dim O$(): O = AyAlignL(Key)
Dim J&
For J = 0 To UB(Key)
   O(J) = O(J) & " " & A(Key(J))
Next
DicLy1 = O
End Function

Function DicLy2(A As Dictionary) As String()
Dim O$(), K
If A.Count = 0 Then Exit Function
For Each K In A.Keys
    Push O, DicLy2__1(CStr(K), A(K))
Next
DicLy2 = O
End Function

Function DicLy2__1(K$, Lines$) As String()
Dim O$(), J&
Dim Ly$()
    Ly = SplitCrLf(Lines)
For J = 0 To UB(Ly)
    Dim Lin$
        Lin = Ly(J)
        If FstChr(Lin) = " " Then Lin = "~" & RmvFstChr(Lin)
    Push O, K & " " & Lin
Next
DicLy2__1 = O
End Function

Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function

Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function

Function DicMge(A As Dictionary, PfxSsl$, ParamArray DicAp()) As Dictionary
Dim Av(): Av = DicAp
Dim Ny$()
   Ny = SslSy(PfxSsl)
   Ny = AyAddSfx(Ny, "@")
If Sz(Av) <> Sz(Ny) Then Stop
Dim Dy() As Dictionary
Dim D As Dictionary
   Dim J%
   For J = 0 To UB(Ny)
       Set D = Av(J)
       Push Dy, DicAddKeyPfx(A, Ny(J))
   Next
Set DicMge = DicAddAy(A, Dy)
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function

Function DicTy(A As Dictionary) As Dictionary
Set DicTy = DicMap(A, "TyNm")
End Function

Sub DicTyBrw(A As Dictionary)
DicBrw DicTy(A)
End Sub


Function DiczTRLy(TermRestLy$()) As Dictionary
Dim I, L$, K$, O As New Dictionary
If Sz(TermRestLy) > 0 Then
    For Each I In TermRestLy
        L = I
        K = ShfTerm(L)
        O.Add K, L
    Next
End If
Set DiczTRLy = O
End Function

Sub DicAddOrUpd(A As Dictionary, K$, V, Sep$)
If A.Exists(K) Then
    A(K) = A(K) & Sep & V
Else
    A.Add K, V
End If
End Sub

Function DicVal(A As Dictionary, K)
If A.Exists(K) Then Asg A(K), DicVal
End Function

Function DicAyKy(A() As Dictionary) As Variant()
Dim I
For Each I In AyNz(A)
   PushNoDupAy DicAyKy, CvDic(I).Keys
Next
End Function

Function DicByDry(DicDry) As Dictionary
Dim O As New Dictionary
If Sz(DicDry) <> 0 Then
   Dim Dr
   For Each Dr In DicDry
       O.Add Dr(0), Dr(1)
   Next
End If
Set DicByDry = O
End Function

Function DicDrsFny(InclDicValTy As Boolean) As String()
DicDrsFny = SplitSpc("Key Val"): If InclDicValTy Then PushI DicDrsFny, "ValTy"
End Function

Function DicHasK(A As Dictionary, K$) As Boolean
DicHasK = A.Exists(K)
End Function

Function DicHasKeyLvs(A As Dictionary, KeyLvs) As Boolean
DicHasKeyLvs = DicHasKy(A, SslSy(KeyLvs))
End Function

Function DicKVLy(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, W%, Ky
Ky = A.Keys
W = AyWdt(Ky)
For Each K In Ky
   Push O, AlignL(K, W) & " = " & A(K)
Next
DicKVLy = O
End Function

Function DicSelIntoAy(A As Dictionary, Ky$()) As Variant()
Dim O()
Dim U&: U = UB(Ky)
ReDim O(U)
Dim J&
For J = 0 To U
   If Not A.Exists(Ky(J)) Then Stop
   O(J) = A(Ky(J))
Next
DicSelIntoAy = O
End Function

Function DicSelIntoSy(A As Dictionary, Ky$()) As String()
DicSelIntoSy = AySy(DicSelIntoAy(A, Ky))
End Function

Function DicStrKy(A As Dictionary) As String()
DicStrKy = AySy(A.Keys)
End Function

Function DicVblDic(A$, Optional JnSep$ = vbCrLf) As Dictionary
Set DicVblDic = DicLyDic(SplitVBar(A), JnSep)
End Function

Function LyDic(A, Optional JnSep$ = vbCrLf) As Dictionary
Dim X, K$, O As New Dictionary
For Each X In AyNz(A)
    K = ShfTerm(X)
    If O.Exists(K) Then
        O(K) = O(K) & vbCrLf & X
    Else
        O.Add K, X
    End If
Next
Set LyDic = O
End Function

Private Sub Z_DicMaxValSz()
Dim D As Dictionary, M%
'Set D = PjMthDic(CurPj)
M = DicMaxValSz(D)
Stop
End Sub
