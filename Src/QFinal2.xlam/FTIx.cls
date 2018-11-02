VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Fm&, A_To&
Property Get FmIx&()
FmIx = A_Fm
End Property
Property Get ToIx&()
ToIx = A_To
End Property
Friend Property Get Init(FmIx, ToIx) As FTIx
If FmIx < 0 Then Stop
If ToIx < 0 Then Stop
If FmIx > ToIx - 1 Then Stop
A_Fm = FmIx
A_To = ToIx
Set Init = Me
End Property
