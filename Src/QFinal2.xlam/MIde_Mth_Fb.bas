Attribute VB_Name = "MIde_Mth_Fb"
Option Explicit
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"
Public Const MthFb$ = "C:\Users\User\Desktop\Vba-Lib-1\Mth.accdb"
Public Const WrkFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthWrk.accdb"
Sub EnsMthTbl()
Dim A As Drs
Set A = CurPjFfnAyMthFullDrs
DrsRplDbt A, MthDb, "Mth"
End Sub

Sub EnsMthFb()
MthFbEns MthFb
End Sub

Function MthFbEns(A$) As Database
FbEns A
Dim Db As Database
Set Db = FbDb(A)

Dim B As StruBase
Set B.F = StruFld( _
    "Nm  Nm Md Pj", _
    "T50 MchStr", _
    "T10 MthPfx", _
    "Txt PjFfn Prm Ret LinRmk", _
    "T3  Ty Mdy", _
    "T4  MdTy", _
    "Lng Lno", _
    "Mem Lines TopRmk")
Const MthCache$ = "MthCache PjFfn Md Nm Ty | Mdy Prm Ret LinRmk TopRmk Lines Lno Pj PjDte MdTy"
Const MthLoc$ = "MthLoc Nm | MthMchTy MchStr ToMdNm"
Const MthPfxMd$ = "MthPfxMd MthPfx | MdNm"
Dim Mth$: Mth = "Mth " & RmvT1(MthCache)

DbStruEns Db, MthCache, B
DbStruEns Db, MthPfxMd, B
DbStruEns Db, MthLoc, B
DbStruEns Db, Mth, B
Set MthFbEns = Db
End Function

Function MthDb() As Database
Static A As Database, B As Boolean
If Not B Then
    B = True
    Set A = FbDb(MthFb)
End If
Set MthDb = A
End Function

Sub BrwMthFb()
FbBrw MthFb
End Sub

