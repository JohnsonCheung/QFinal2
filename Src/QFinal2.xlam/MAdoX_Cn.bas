Attribute VB_Name = "MAdoX_Cn"
Option Explicit

Function FxCn(A) As ADODB.Connection
Set FxCn = AdoCnStrCn(FxAdoCnStr(A))
End Function

Function FbCn(A) As ADODB.Connection
Set FbCn = AdoCnStrCn(FbAdoCnStr(A))
End Function
