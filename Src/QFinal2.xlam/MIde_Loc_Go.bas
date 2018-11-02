Attribute VB_Name = "MIde_Loc_Go"
Option Explicit
Const CMod$ = "MIde_Loc_Go."
Sub MdGoLno(A As CodeModule, Lno&)
'MdGoRRCC A, RRCC(Lno, Lno, 1, 1)
End Sub

Sub MdLCCGo(A As CodeModule, B As LCC)
Const CSub$ = CMod & "MdPos"
MdGo A
'If IsNothing(LCC) Then Msg CSub, "Given LCC is nothing": Exit Sub
With B
    A.CodePane.TopLine = .Lno
    A.CodePane.SetSelection .Lno, .C1, .Lno, .C2
End With
SendKeys "^{F4}"
End Sub
Sub MdMthGo(A As CodeModule, MthNm$)
'MdGoRRCC A, MdMthRRCC(A, MthNm)
End Sub

Function MthLCC(A As Mth) As LCC
Dim L%, C As LCC
Dim M As CodeModule
Set M = A.Md
For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
    C = LinMthLCC(M.Lines(L, 1), A.Nm, L)
    'If Not IsNothing(C) Then
    '    Set MthLCC = C
        Exit Function
    'End If
Next
End Function


Sub MthGo(A As Mth)
'MdGoLCC A.Md, MthLCC(A)
End Sub

Sub MdGo1(A As CodeModule, B As VbeLoc)
'If IsEmpRRCC(RRCC) Then Debug.Print FmtQQ("Given RRCC_ is empty"): Exit Sub
MdShw A
'If IsRRCCOutSidMd(RRCC, A) Then
'    With RRCC
'        Debug.Print FmtQQ("MdGoRg: Given ? is outside given Md[?]-(MaxR ?)(MaxR1C ?)(MaxR2C ?)", RRCC_Str(RRCC), MdNm(A), MdNLin(A), Len(A.Lines(.R1, 1)), Len(A.Lines(.R2, 1)))
'    End With
    Exit Sub
'End If
'With RRCC
'    A.CodePane.SetSelection .R1, .C1, .R2, .C2
'End With
End Sub

Sub MdGoTy(A As CodeModule, TyNm$)
'MdGoRRCC A, MdTyRRCC(A, TyNm)
End Sub



Sub MdGo(A As CodeModule)
MdShw A
BrwObjWin.Visible = True
'WinApKeep MdWin(A), BrwObjWin
ClrImmWin
TileV
End Sub
