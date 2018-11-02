Attribute VB_Name = "MIde__Action"
Option Explicit
Sub AddMod(Nm$)
PjAddMod CurPj, Nm
End Sub

Sub CrtResTbl()
DbCrtResTbl CurrentDb
End Sub

Sub AddCls(Nm$)
PjAddCmp CurPj, Nm, vbext_ComponentType.vbext_ct_ClassModule
End Sub

Sub AddFun(FunNm$)
'Des: Add Empty-Fun-Mth to CurMd
MdLinesApp CurMd, FmtQQ("Function ?()|End Function", FunNm)
MdMthGo CurMd, FunNm
End Sub

Sub AddSub(SubNm$)
MdLinesApp CurMd, FmtQQ("Sub ?()|End Sub", SubNm)
MdMthGo CurMd, SubNm
End Sub

Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
    A = PjCmpNy(CurPj, WhMd("Std", WhNm(Patn, Exl)))
    A = AySrt(A)
    A = AyAddPfx(A, "ShwMbr """)
D A
End Sub

Sub LisMdMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$)
Dim Ny$(), M As WhMth
Set M = WhMth(WhMdy, WhKd, WhNm(MthPatn, MthExl))
Ny = MthDDNyWh(MdMthDDNy(CurMd), M)
D AyAddPfx(Ny, CurPjNm & ".")
End Sub

Sub LisPj()
Dim A$()
    A = VbePjNy(CurVbe)
    D AyAddPfx(A, "ShwPj """)
D A
End Sub


Function PjMdLisDt(A As VBProject, Optional B As WhMd) As Dt
Stop
End Function

Sub PjMdLisDtBrw(A As VBProject, Optional B As WhMd)
DtBrw PjMdLisDt(A, B)
End Sub

Sub PjMdLisDtDmp(A As VBProject, Optional B As WhMd)
DtDmp PjMdLisDt(A, B)
End Sub
