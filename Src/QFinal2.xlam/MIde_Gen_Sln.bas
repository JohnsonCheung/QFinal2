Attribute VB_Name = "MIde_Gen_Sln"
Option Explicit
Private Sub Z_ZPjNy()
D ZPjNy
End Sub

Sub GenSln()
Stop
Dim N
ZOupPthClr
For Each N In ZPjNy
    ZGenPj N
Next
End Sub
Private Sub ZOupPthClr()
PthClr ZOupPth
If Sz(PthFnAy(ZOupPth)) > 0 Then Stop
End Sub
Private Sub ZOupPthBrw()
PthBrw ZOupPth
End Sub
Private Function ZOupPth$()
ZOupPth = TmpPth("QFinalSln")
End Function
Private Function ZCrtPj(PjNm) As VBProject
Set ZCrtPj = FxaPj(ZOupPth & PjNm & ".xlam")
End Function

Private Sub ZGenPj(PjNm)
Dim M, ToPj As VBProject
Set ToPj = ZCrtPj(PjNm)
For Each M In ZPjMdNy(PjNm)
    MdCpy ZSrcMd(M), ToPj
Next
PjSav ToPj
End Sub
Private Function ZSrcMd(MdNm) As CodeModule
Set ZSrcMd = ZSrcPj.VBComponents(MdNm).CodeModule
End Function
Private Function ZPjNy() As String()
Dim N
For Each N In PjModNy(ZSrcPj)
    PushNoDupNonBlankStr ZPjNy, ZMdNmPjNm(N)
Next
End Function
Private Function ZMdNmPjNm__Len%(MdNm)
Dim J%, C$
For J = 5 To Len(MdNm)
    C = Mid(MdNm, J, 1)
    If AscIsUCase(Asc(C)) Or C = "_" Then
        ZMdNmPjNm__Len = J - 4
        Exit Function
    End If
Next
ZMdNmPjNm__Len = J - 4
End Function
Private Function ZMdNmPjNm$(MdNm)
If Left(MdNm, 3) <> "Lib" Then Exit Function
Dim L%
    L = ZMdNmPjNm__Len(MdNm)
    
ZMdNmPjNm = "Q" & Mid(MdNm, 4, L)
End Function
Private Sub Z_ZPjMdNy()
D ZPjMdNy("QVb")
End Sub
Private Function ZPjMdNy(PjNm) As String()
Dim N
For Each N In PjModNy(ZSrcPj)
    If ZMdNmPjNm(N) = PjNm Then
        PushNoDup ZPjMdNy, N
    End If
Next
For Each N In PjClsNy(ZSrcPj)
    If Not ZClsPjDic.Exists(N) Then Stop
    If ZClsPjDic(N) = PjNm Then
        PushNoDup ZPjMdNy, N
    End If
Next
End Function
Private Function ZClsPjLy() As String()
Dim O$()
PushI O, "Blk       QTp"
PushI O, "DCRslt    QVb"
PushI O, "Drs       QDta"
PushI O, "Ds        QDta"
PushI O, "Dt        QDta"
PushI O, "FmCnt     QVb"
PushI O, "FTIx      QVb"
PushI O, "FTNo      QVb"
PushI O, "Gp        QTp"
PushI O, "LCC       QVb"
PushI O, "LnkCol    QDao"
PushI O, "LnoCnt    QVb"
PushI O, "Lnx       QTp"
PushI O, "VbeLoc    QIde"
PushI O, "Mth       QIde"
PushI O, "RRCC      QIde"
PushI O, "S1S2      QVb"
PushI O, "TblImpSpec    QDao"
PushI O, "WhMd      QIde"
PushI O, "WhMdMth   QIde"
PushI O, "WhMth     QIde"
PushI O, "WhNm      QIde"
PushI O, "WhPjMth   QIde"
PushI O, "P123      QVb"
PushI O, "SwBrk     QTp"
PushI O, "Sql_Shared    QTp"
ZClsPjLy = O
End Function
Private Function ZClsPjDic() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = LyDic(ZClsPjLy)
Set ZClsPjDic = X
End Function
Private Function ZSrcPj() As VBProject
Static X As VBProject
If IsNothing(X) Then Set X = Pj("QFinal1")
Set ZSrcPj = X
End Function
