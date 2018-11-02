Attribute VB_Name = "MSql_Sel"
Option Explicit
Private X As New Sql_Shared

Function QSel_FF_Fm$(FF, Fm, Optional WhBExpr$, Optional IsDis As Boolean)
QSel_FF_Fm = X.Sel(FF, IsDis) & X.Fm(Fm) & X.Wh(WhBExpr)
End Function
Function QSel_FF_Fm_WhF_InAy$(FF, Fm, WhF, InAy, Optional IsDis As Boolean)
Dim W$
W = X.FldInAy(WhF, InAy)
QSel_FF_Fm_WhF_InAy = QSel_FF_Fm(FF, Fm, W, IsDis)
End Function

Function SelDis_FF_Fm$(FF, T)
SelDis_FF_Fm = QSel_FF_Fm(FF, T, IsDis:=True)
End Function

Function QSel_FF_EDic_Fm$(FF, EDic As Dictionary, Fm, Optional IsDis As Boolean)
'SelFFExprDicSqp = "Select" & vbCrLf & FFExprDicAsLines(FF, ExprDic)
End Function
Function SelDis_FF_EDic_Fm$(FF, EDic As Dictionary, Fm)
SelDis_FF_EDic_Fm = QSel_FF_EDic_Fm(FF, EDic, Fm, IsDis:=True)
End Function

Function QSel_XX_Into_Fm$(XX$, Into, Fm, Optional WhBExpr$)
QSel_XX_Into_Fm = X.SelX(X) & X.Into(Into) & X.Fm(Fm) & X.Wh(WhBExpr)
End Function

Function QSel_FF_Into_Fm$(FF, Into, Fm, Optional WhBExpr$)
QSel_FF_Into_Fm = QSel_XX_Into_Fm(X.FFJnComma(FF), Into, Fm, WhBExpr)
End Function

Function QSel_FF_Fm_WhFny_Ay$(FF, Fm, Fny$(), Ay)
QSel_FF_Fm_WhFny_Ay = QSel_FF_Fm(FF, Fm, X.WhFnyEqAy(Fny, Ay))
End Function

Function QSel_Fny_Ey_Into_Fm$(Fny$(), Ey$(), Into, Fm, Optional WhBExpr$)
Dim XX$: XX = X.FnyEyAsLines(Fny, Ey)
QSel_Fny_Ey_Into_Fm = QSel_XX_Into_Fm(XX, Into, Fm, WhBExpr)
End Function

Function FF_EDic_T$(FF, T, EDic As Dictionary)
'SelTFExprSql = QSel_FF_EDic(FF, EDic) & FmSqp(T)
End Function

Function QSel_F_T$(F, T, Optional WhBExpr$)
QSel_F_T = FmtQQ("Select [?] from [?]?", T, F, X.Wh(WhBExpr))
End Function

Function QSel_Fm$(Fm, Optional WhBExpr$)
QSel_Fm = "Select *" & X.Fm(Fm) & X.Wh(WhBExpr)
End Function
