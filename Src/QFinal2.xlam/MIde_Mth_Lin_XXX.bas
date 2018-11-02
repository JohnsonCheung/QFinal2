Attribute VB_Name = "MIde_Mth_Lin_XXX"
Option Explicit
Function LinMdy$(A)
LinMdy = AyFstEqV(MdyAy, LinT1(A))
End Function

Private Sub Z_MthLinRetTy()
Dim MthLin$
Dim A As MthPmTy:
MthLin = "Function MthPm(MthPmStr$) As MthPm"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPm"
Ass A.IsAy = False
Ass A.TyChr = ""

MthLin = "Function MthPm(MthPmStr$) As MthPm()"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = "MthPm"
Ass A.IsAy = True
Ass A.TyChr = ""

MthLin = "Function MthPm$(MthPmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = "$"

MthLin = "Function MthPm(MthPmStr$)"
A = MthLinRetTy(MthLin)
Ass A.TyAsNm = ""
Ass A.IsAy = False
Ass A.TyChr = ""
End Sub

Function LinMthKd$(A)
LinMthKd = TakMthKd(RmvMdy(A))
End Function

Function LinMthShtTy$(A)
LinMthShtTy = MthShtTy(LinMthTy(A))
End Function

Function LinMthTy$(A)
LinMthTy = TakPfxAySpc(RmvMdy(A), MthTyAy)
End Function



Private Sub Z_LinMthTy()
Dim O$(), L
For Each L In Src("Fct")
    Push O, LinMthShtTy(L) & "." & L
Next
Brw O
End Sub



Private Sub Z_LinMthKd()
Dim A$
Ept = "Property": A = "Private Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = LinMthKd(A)
    C
    Return
End Sub
Function MthLinRetTy(MthLin$) As MthPmTy
If Not HasSubStr(MthLin, "(") Then Exit Function
If Not HasSubStr(MthLin, ")") Then Exit Function
Dim TC$: TC = LasChr(TakBefBkt(MthLin))
With MthLinRetTy
    If IsTyChr(TC) Then .TyChr = TC: Exit Function
    Dim Aft$: Aft = TakAftBkt(MthLin)
        If Aft = "" Then Exit Function
        If Not HasPfx(Aft, " As ") Then Stop
        Aft = RmvPfx(Aft, " As ")
        If HasSfx(Aft, "()") Then
            .IsAy = True
            Aft = RmvSfx(Aft, "()")
        End If
        .TyAsNm = Aft
        Exit Function
End With
End Function
