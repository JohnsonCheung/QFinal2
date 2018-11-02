VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FTNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public FmNo&, ToNo& ' From-Number, To-Number.  Number is starting from 1

Friend Property Get Init(FmNo%, ToNo%) As FTNo
If FmNo <= 0 Then Stop
If ToNo <= 0 Then Stop
If FmNo - 1 > ToNo Then Stop
Me.FmNo = FmNo
Me.ToNo = ToNo
Set Init = Me
End Property
