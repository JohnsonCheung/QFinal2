Attribute VB_Name = "MSql_Upd"
Option Explicit
Private X As New Sql_Shared
Sub Z()
Z_UpdSqlFmt
End Sub
Function UpdSqlFmt$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSqlFmt = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVy(): GoSub X_Fny1_Dr1_SkVy
    Upd = "Update [" & T & "]"
    Set_ = SetSqp(Fny1, Dr1)
    Wh = X.WhFnyEqAy(Sk, SkVy)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = VarAySqlQuote(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = AyIx(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not AyHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Private Sub Z_UpdSqlFmt()
Dim T$, Sk$(), Fny$(), Dr
T = "A"
Sk = LinTermAy("X Y")
Fny = LinTermAy("X Y A B C")
Dr = Array(1, 2, 3, 4, 5)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    A = 3 ," & _
vbCrLf & "    B = 4 ," & _
vbCrLf & "    C = 5 " & _
vbCrLf & "  Where" & _
vbCrLf & "    X = 1 And" & _
vbCrLf & "    Y = 2 "
GoSub Tst

T = "A"
Sk = LinTermAy("[A 1] B CD")
Fny = LinTermAy("X Y B Z CD [A 1]")
Dr = Array(1, 2, 3, 4, "XX", #1/2/2018 12:34:00 PM#)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    X = 1 ," & _
vbCrLf & "    Y = 2 ," & _
vbCrLf & "    Z = 4 " & _
vbCrLf & "  Where" & _
vbCrLf & "    [A 1] = #2018-01-02 12:34:00# And" & _
vbCrLf & "    B     = 3                     And" & _
vbCrLf & "    CD    = 'XX'                  "
GoSub Tst
Exit Sub
Tst:
    Act = UpdSql(T, Sk, Fny, Dr)
    C
    Return
End Sub

Function QAddCol$(T, Fny0, F As Drs, E As Dictionary)
Dim O$(), Fld
For Each Fld In CvNy(Fny0)
'    PushI O, Fld & " " & FldSqlTy(Fld, F, E)
Next
QAddCol = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function

Function CrtPkSql$(T)
CrtPkSql = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function

Function CrtSkSql$(T, Sk0)
CrtSkSql = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(AyQuoteSqBktIfNeed(CvNy(Sk0))))
End Function

Function CrtTblSql$(T, FldList$)
CrtTblSql = FmtQQ("Create Table [?] (?)", T, FldList)
End Function

Function DrpFldSql$(T, F)
DrpFldSql = FmtQQ("Alter Table [?] drop column [?]", T, F)
End Function

Function DrpTblSql$(T)
DrpTblSql = "Drop Table [" & T & "]"
End Function

Function DrsCrtTblSql$(A As Drs, T)
Dim F, J%, Dry(), O$()
Dry = A.Dry
For Each F In A.Fny
    PushI O, F & " " & DryColSqlTy(Dry, J)
    J = J + 1
Next
DrsCrtTblSql = CrtTblSql(T, JnComma(O))
End Function

Function InsDrSql$(T, Fny0, Dr)
InsDrSql = FmtQQ("Insert into [?] (?) values(?)", T, JnComma(CvNy(Fny0)), JnComma(VarAySqlQuote(Dr)))
End Function

Function InsSql$(T, Fny$(), Dr)
Dim A$, B$
A = JnComma(Fny)
B = JnComma(AyMapSy(Dr, "VarSqlQuote"))
InsSql = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function

Function StruSkSql$(Stru, T)
Dim Sk$()
Sk = SslSy(RmvT1(Replace(TakBef(Stru, "|"), "*", T)))
If Sz(Sk) = 0 Then Exit Function
StruSkSql = CrtSkSql(T, Sk)
End Function

Function SelTblSql$(T)
SelTblSql = "Select * from [" & T & "]"
End Function

Function SelTblWhSql$(T, WhBExpr$)
SelTblWhSql = SelTblSql(T) & X.Wh(WhBExpr)
End Function

Function SelTFSql$(T, F)
SelTFSql = FmtQQ("Select [?] from [?]", F, T)
End Function

Function UpdSql$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSql = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVy(): GoSub X_Fny1_Dr1_SkVy
    Upd = "Update [" & T & "]"
    Set_ = SetSqp(Fny1, Dr1)
    Wh = X.WhFnyEqAy(Sk, SkVy)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = VarAySqlQuote(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = AyIx(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not AyHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function



Function SetSqp$(Fny$(), Vy())
Dim A$: GoSub X_A
SetSqp = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = AySqBktQuoteIfNeed(Fny)
    Dim R$(): R = VarAySqlQuote(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function
Function FnyAlignQuote(Fny$()) As String()
FnyAlignQuote = AyAlignL(AySqBktQuoteIfNeed(Fny))
End Function

