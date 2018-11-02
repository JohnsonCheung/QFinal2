Attribute VB_Name = "MDao_Schm_Stru_Chk"
Option Explicit
Const CMod = "DaoSchmChk."
Function StruChk(Stru, F As Drs, E As Dictionary) As String() ' Chk may return Sz=0, But Er always Sz>0
Dim Fny$(), Fld, O$()
Fny = StruFny(Stru)
For Each Fld In Fny
    PushIAy O, FldChk(Fld, F, E)
Next
If Sz(0) > 0 Then
    StruChk = AyAddAp("", O)
End If
End Function

Private Function FldChk(Fld, F As Drs, E As Dictionary) As String()
Const CSub$ = CMod & "FldChk"
If IsStdFld(Fld) Then Exit Function
Dim Ele$
Ele = FldChk1$(Fld, F): If Ele = "" Then FldChk = Er1(Fld, Ele): Exit Function
If IsStdEle(Ele) Then Exit Function
If IsNothing(E) Then FldChk = Er2(Fld)
If Not E.Exists(Ele) Then FldChk = Er3(Fld, Ele)
End Function

Private Function FldChk1$(Fld, F As Drs) ' Return Ele$
Dim Dr
For Each Dr In AyNz(F.Dry)
    If Fld Like Dr(1) Then FldChk1 = Dr(0): Exit Function
Next
End Function

Private Function Er1(Fld, Ele) As String()
Const CSub$ = CMod & "Er1"
Const Msg$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
PushI Er1, FunMsgLin(CSub, Msg, Fld, Ele)
End Function

Private Function Er2(Fld) As String()
Const CSub$ = CMod & "Er2"
Const Msg$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
PushI Er2, FunMsgLin(CSub, Msg, Msg, Fld)
End Function

Private Function Er3(Fld, Ele) As String()
Const Msg1$ = "[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]"
PushI Er3, MsgLin(Msg1, Fld, Ele)
End Function
