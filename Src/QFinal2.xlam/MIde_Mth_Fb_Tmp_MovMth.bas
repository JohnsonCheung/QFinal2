Attribute VB_Name = "MIde_Mth_Fb_Tmp_MovMth"
Option Explicit
Sub Z()
Z_M
End Sub
Private Sub Z_M()
Dim Ly$()
Push Ly, "2  ^Dr    DtaDr  AAAMod"
Const TarMdNm$ = "AAARunModule"
MdClr Md(TarMdNm)
MdLinesApp Md(TarMdNm), SrcCd(Ly)
End Sub

Private Function SrcCd$(FmMd_MthNmXXX_Patn_ToMd_FmMd$())
Dim O$()
    Dim MthXXX$, Patn$, ToMd$, FmMd$
    Dim L
    For Each L In FmMd_MthNmXXX_Patn_ToMd_FmMd
        LinAsg3TRst CStr(L), MthXXX, Patn, ToMd, FmMd
        PushI O, SrcCdMovMthXXX(MthXXX, Patn, FmMd, ToMd) '<=======
    Next
    PushI O, SrcCdMovMth(FmMd_MthNmXXX_Patn_ToMd_FmMd)    '<===
SrcCd = JnCrLf(O)
End Function

Private Function SrcCdMovMth$(Ly$())
Dim O$()
PushI O, "Sub Mov()"
Dim I
For Each I In AyTakT1(Ly)
    PushI O, "ZZMov_" & I
Next
PushI O, "Debug.Print ""MthCnt-In-Md-AAAMd:""" & MdMthCnt(Md("AAAMod"))
PushI O, "End Sub"
SrcCdMovMth = JnCrLf(O)
End Function

Private Function SrcCdMovMthXXX$(MthXXX$, Patn$, FmMd$, ToMd$)
Dim O$()
PushI O, "Private Sub ZZMov_" & MthXXX
PushI O, FmtQQ("PjEnsMod CurPj,""?""", ToMd)
PushI O, SrcCdBdy(Patn, FmMd, ToMd)
PushI O, "End Sub"
SrcCdMovMthXXX = JnCrLf(O)
End Function

Private Function SrcCdBdy$(Patn$, FmMd$, ToMd$)
Dim N, O$(), J%, Ny$()
Ny = MthNmBrkAyNy(MthNmBrkAyWh(MdMthBrkAy(Md(FmMd)), WhMth(Nm:=WhNm(Patn))))
For Each N In AyNz(Ny)
    PushI O, SrcCdLin(CStr(N), FmMd, ToMd)
Next
SrcCdBdy = JnCrLf(AyFmt(O, "MthMov"))
End Function

Private Function SrcCdLin$(Nm$, FmMd$, ToMd$)
SrcCdLin = FmtQQ("A=""?"":    MthMov Mth(Md(""" & FmMd & """),A),Md(""?"")", Nm, ToMd)
End Function

Sub DmpCnt()
Dim M As CodeModule
Set M = Md("AAAMod")
Debug.Print "AAAMod-Prv-MthCnt"; MdPrvMthCnt(M)
Debug.Print "AAAMod-Tot-MthCnt"; MdMthCnt(M)
Debug.Print "AAAMod-Tot-MthPfxCnt"; Sz(MdMthPfxAy(M))
End Sub

Function MdPrvMthCnt%(A As CodeModule)
Dim O%, L, B As WhMth
Set B = WhMth("Prv")
For Each L In AyNz(MdSrc(A))
    If LinIsMthWh(CStr(L), B) Then O = O + 1
Next
MdPrvMthCnt = O
End Function

Function SrcPrvMthCnt%(A$())
Dim O%, L
For Each L In AyNz(A)
    If LinIsMthWh(CStr(L), WhMth("Prv")) Then O = O + 1
Next
SrcPrvMthCnt = O
End Function
