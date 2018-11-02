Attribute VB_Name = "MIde_Z_Pj_Dte"
Option Explicit
Function FbPjDte(A) As Date
Static Y As New Access.Application
Y.OpenCurrentDatabase A
Y.Visible = False
FbPjDte = AcsPjDte(Y)
Y.CloseCurrentDatabase
End Function

Function PjFfnPjDte(PjFfn) As Date
Select Case True
Case IsFxa(PjFfn): PjFfnPjDte = FileDateTime(PjFfn)
Case IsFb(PjFfn): PjFfnPjDte = FbPjDte(PjFfn)
Case Else: Stop
End Select
End Function

Function AcsPjDte(A As Access.Application)
Dim O As Date
Dim M As Date
M = ItrMaxPrp(A.CurrentProject.AllForms, "DateModified")
O = Max(O, M)
O = Max(O, ItrMaxPrp(A.CurrentProject.AllModules, "DateModified"))
O = Max(O, ItrMaxPrp(A.CurrentProject.AllReports, "DateModified"))
AcsPjDte = O
End Function
