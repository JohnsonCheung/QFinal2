VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LnoCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Lno&, A_Cnt&
Property Get Init(Lno, Cnt) As LnoCnt
A_Lno = Lno
A_Cnt = Cnt
Set Init = Me
End Property
Property Get Lno&()
Lno = A_Lno
End Property
Property Get Cnt&()
Cnt = A_Cnt
End Property
