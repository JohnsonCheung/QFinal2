Attribute VB_Name = "MAdo_Fb_Fbq"
Option Explicit
Const SampleFb_DutyPrepare$ = ""
Function FbqAdoDrs(A$, Q$) As Drs
Set FbqAdoDrs = ARsDrs(FbqARs(A, Q))
End Function

Private Sub Z_FbqAdoDrs()
Const Fb$ = SampleFb_DutyPrepare
Const Q$ = "Select * from Permit"
DrsBrw FbqAdoDrs(Fb, Q)
End Sub

Function FbqARs(A$, Q$) As ADODB.Recordset
Set FbqARs = FbCn(A).Execute(Q)
End Function

Sub FbqARun(A$, Q$)
FbCn(A).Execute Q
End Sub

Sub FbtDrp(A$, T$)
'DbtDrp FbDb(A), T
End Sub

Private Sub Z_FbqARun()
Const Fb$ = SampleFb_DutyPrepare
Const Q$ = "Select * into [#a] from Permit"
FbtDrp Fb, "#a"
FbqARun Fb, Q
End Sub
