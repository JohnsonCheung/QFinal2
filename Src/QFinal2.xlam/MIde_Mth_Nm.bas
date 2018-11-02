Attribute VB_Name = "MIde_Mth_Nm"
Option Explicit
Sub Z()
Z_LinMthNm
Z_SrcMthNy
End Sub

Function CurMthDNm$()
Dim M$: M = CurMthNm
If M = "" Then Exit Function
CurMthDNm = CurMdDNm & "." & M
End Function

Function MthNy(Optional A As WhPjMth) As String()
MthNy = VbeMthNy(CurVbe, A)
End Function


Private Sub Z_SrcMthNy()
Brw SrcMthNy(CurSrc)
End Sub

Function MdMthDDNy(A As CodeModule) As String()
MdMthDDNy = SrcMthDDNy(MdSrc(A))
End Function

Function MthNm$(A As Mth)
MthNm = A.Nm
End Function

Function LinMthNm$(A)
LinMthNm = MthNmBrkNm(LinMthNmBrk(A))
End Function

Function MdDftMthNm$(Optional A As CodeModule, Optional MthNm$)
If MthNm = "" Then
   MdDftMthNm = MdCurMthNm(DftMd(A))
Else
   MdDftMthNm = A
End If
End Function

Function MthNmMd(A$) As CodeModule '
Dim O As CodeModule
Set O = CurMd
If MdHasMth(O, A) Then Set MthNmMd = O: Exit Function
Dim N$
N = MthFul(A)
If N = "" Then
    Debug.Print FmtQQ("Mth[?] not found in any Pj")
    Exit Function
End If
Set MthNmMd = Md(N)
End Function

Private Sub Z_LinMthNm()
GoTo ZZ
Dim A$
A = "Function LinMthNm$(A)": Ept = "LinMthNm.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = LinMthNm(A)
    C
    Return
ZZ:
    Dim O$(), L, P, M
    For Each P In VbePjAy(CurVbe)
        For Each M In PjMdAy(CvPj(P))
            For Each L In MdBdyLy(CvMd(M))
                PushNonBlankStr O, LinMthNm(CStr(L))
            Next
        Next
    Next
    Brw O
End Sub

Sub Z_MdDftMthNm()
Dim I, Md As CodeModule
For Each I In PjMdAy(CurPj)
   MdShw CvMd(I)
   Debug.Print MdNm(Md), MdDftMthNm(Md)
Next
End Sub

Function DftMthNm$(MthNm0$)
If MthNm0 = "" Then
    DftMthNm = CurMthNm
    Exit Function
End If
DftMthNm = MthNm0
End Function

Function DDNmMth(MthDDNm$) As Mth
Dim M As CodeModule
Dim Nm$
Dim Ny$(): Ny = Split(MthDDNm, ".")
Select Case Sz(Ny)
Case 1: Nm = Ny(0): Set M = CurMd
Case 2: Nm = Ny(1): Set M = Md(Ny(0))
Case 3: Nm = Ny(2): Set M = PjMd(Pj(Ny(0)), Ny(1))
Case Else: Stop
End Select
Set DDNmMth = Mth(M, Nm)
End Function

Function LinPrpNm$(A)
Dim L$
L = RmvMdy(A)
If ShfMthKd(L) <> "Property" Then Exit Function
LinPrpNm = TakNm(L)
End Function

Function MthDDNyWh(A$(), B As WhMth) As String()
If IsNothing(B) Then
    MthDDNyWh = A
    Exit Function
End If
Dim N
For Each N In AyNz(A)
    If IsMthDDNmSel(N, B) Then
        PushI MthDDNyWh, N
    End If
Next
End Function

Function MthDotNTM$(MthDot$)
'MthDot is a string with last 3 seg as Mdy.ShtTy.Nm
'MthNTM is a string with last 3 seg as Nm:ShtTy.Mdy
Dim Ay$(), Nm$, ShtTy$, Mdy$
Ay = SplitDot(MthDot)
'AyAsg AyPop(Ay), Ay, Nm
'AyAsg AyPop(Ay), Ay, ShtTy
'AyAsg AyPop(Ay), Ay, Mdy
Push Ay, FmtQQ("?:?.?", Nm, ShtTy, Mdy)
MthDotNTM = JnDot(Ay)
End Function


Function MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthDNm_Nm = Nm
End Function
