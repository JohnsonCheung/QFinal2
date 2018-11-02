Attribute VB_Name = "MVb__Tst"
Option Explicit
Public Act, Ept, Dbg As Boolean, Trc As Boolean
Sub C()
If Not IsEq(Act, Ept) Then
    'ShwDbg
    D "=========================="
    D "Act"
    D Act
    D "---------------"
    D "Ept"
    D Ept
    Stop
End If
End Sub

Function TstResFdr$(Fdr$)
Dim O$
    O = TstResPth & Fdr & "\"
    PthEns O
TstResFdr = O
End Function

Sub TstResFdrBrw(Fdr$)
PthBrw TstResFdr(Fdr)
End Sub

Function TstResPth$()
Dim Pth$
Stop '
TstResPth = PthEns(Pth & "TstRes\")
End Function

Sub TstResPthBrw()
PthBrw TstResPth
End Sub
