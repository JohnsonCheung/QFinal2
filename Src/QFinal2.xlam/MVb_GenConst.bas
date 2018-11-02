Attribute VB_Name = "MVb_GenConst"
Option Explicit
Private A_Pj$
Private A_Md$
Private A_Nm$
Private A_IsPub As Boolean
Sub Z_ConstEdt()
ConstEdt "ZZ_A"
End Sub
Sub ConstEdt(ConstNm$, Optional IsPub As Boolean)
ZSetA ConstNm, IsPub
Dim Ft$
    Ft = ZFt
StrWrt ZValFmSrc, Ft, True
FtBrw Ft
End Sub
Sub Z_ConstUpdSrc()
ConstUpdSrc "ZZ_A"
End Sub

Sub ConstUpdSrc(ConstNm$, Optional IsPub As Boolean)
ZSetA ConstNm, IsPub
Dim Md As CodeModule
    Set Md = ZMd
ZRmvMth Md
ZAddMth Md
End Sub

Private Function ZMd() As CodeModule
Set ZMd = Application.Vbe.VBProjects(A_Pj).VBComponents(A_Md).CodeModule
End Function

Private Function ZFm&(Ly$())
Dim L, O&
O = 1
For Each L In Ly
    If ZIsHit(L) Then
        ZFm = O
        Exit Function
    End If
    O = O + 1
Next
End Function
Private Sub Z_ZIsHit()
Dim Lin$
'----
A_Nm = "AA":  Lin = "Private Function AA$()": Ept = True: GoSub Tst
A_Nm = "AAA": Lin = "Private Function AA$()": Ept = False: GoSub Tst
A_Nm = "AA":  Lin = "Function AA$()":         Ept = True: GoSub Tst
A_Nm = "AAA": Lin = "Function AA$()":         Ept = False: GoSub Tst
Exit Sub
Tst:
    Act = ZIsHit(Lin)
    C
    Return
End Sub
Private Function ZIsHit(Lin) As Boolean
Dim L$: L = RmvPfx(Lin, "Private ")
If Not ShfPfxSpc(L, "Function") Then Exit Function
ZIsHit = HasPfx(L, A_Nm & "$")
End Function
Private Function ZCnt&(Ly$(), Fm&)
Dim I&, O&
O = 2
For I = Fm To UB(Ly)
    If HasPfx(Ly(I), "End Function") Then ZCnt = O: Exit Function
    O = O + 1
Next
Stop 'Impossible
End Function
Private Sub Z_ZRmvMth()
Dim Md As CodeModule
'----
Set Md = Application.Vbe.VBProjects("QVb").VBComponents("M_GenConst").CodeModule
A_Nm = "ZZ_A"
GoSub MakeMth
GoSub Tst
Exit Sub
MakeMth:
    Dim Lines$, L1$, L2$, L3$
    L1 = "Private Function " & A_Nm & "$()" & vbCrLf
    L2 = "'...." & vbCrLf
    L3 = "End Function"
    Lines = vbCrLf & L1 & L2 & L3
    Md.InsertLines Md.CountOfLines + 1, Lines
    Return
Tst:
    ZRmvMth Md
    Return
End Sub
Private Sub ZRmvMth(A As CodeModule)
Dim IFm&, ICnt&
    Dim Ly$()
    Ly = SplitCrLf(A.Lines(1, A.CountOfLines))
    IFm = ZFm(Ly):        If IFm = 0 Then Exit Sub
    ICnt = ZCnt(Ly, IFm): If ICnt = 0 Then Exit Sub
A.DeleteLines IFm, ICnt
End Sub
Private Sub ZAddMth(A As CodeModule)
Dim MthLines$
    Dim ConstVal$
    ConstVal = ZValFmFt
    If ConstVal = "" Then Debug.Print "ConstUpdSrc: ZValFmFt Ft[" & ZFt & "] is blank": Exit Sub
    MthLines = ConstValMthLInes(ConstVal, A_Nm, A_IsPub)
A.InsertLines A.CountOfLines + 1, MthLines
End Sub
Private Function ZDftPjNm$()
ZDftPjNm = ZDftMd.Parent.Collection.Parent.Name
End Function

Private Function ZDftMdNm$()
ZDftMdNm = ZDftMd.Parent.Name
End Function
Private Function ZDftMd() As CodeModule
Set ZDftMd = Application.Vbe.ActiveCodePane.CodeModule
End Function
Private Sub ZSetA(ConstNm$, IsPub As Boolean)
Dim Ay$()
    Ay = SplitDot(ConstNm)
Select Case Sz(Ay)
Case 1: A_Pj = ZDftPjNm: A_Md = ZDftMdNm:   A_Nm = ConstNm
Case 2: A_Pj = ZDftPjNm: A_Md = Ay(0):      A_Nm = Ay(1)
Case 3: A_Pj = Ay(0):    A_Md = Ay(1):      A_Nm = Ay(2)
Case Else: Stop
End Select
A_IsPub = IsPub
If ZMd.Parent.Name <> A_Md Then Stop
If ZMd.Parent.Collection.Parent.Name <> A_Pj Then Stop
End Sub
Private Function ZFt$()
ZFt = ZPth & JnDot(ApSy(A_Pj, A_Md, A_Nm, "txt"))
End Function
Private Function ZPth$()
Static X$
If X = "" Then X = TmpHom & "GenConst\": PthEns X
ZPth = X
End Function
Sub Z_ZValFmFt()
ZSetA "ZZ_A", True
Debug.Print ZValFmFt
End Sub
Private Function ZValFmFt$()
ZValFmFt = FtMayLines(ZFt)
End Function
Private Function ZDotNm$()
ZDotNm = ApJnDot(A_Pj, A_Md, A_Nm)
End Function
Private Sub Z_ZValFmSrc()
Dim ConstNm$, IsPub As Boolean
'----
ConstNm = "ZZ_A"
IsPub = True
GoSub Tst
Tst:
    ZSetA ConstNm, IsPub
    Act = ZValFmSrc
    C
    Return
End Sub
Private Function ZValFmSrc$()
On Error Resume Next
ZValFmSrc = Run(ZDotNm)
End Function

Sub AA()
Z_ConstUpdSrc
End Sub
Private Function ZZ_A$()
ZZ_A = "JOhnson lskdfj klsdjf lksdj fklsdjf skldf" & _
vbCrLf & "sdkljf lksjdf " & _
vbCrLf & ""
End Function

