Attribute VB_Name = "MTp___EnmAndTy"
Option Explicit
Enum eTpPosTy
    ePosRCC = 1
    ePosRR = 2
    ePosR = 3
End Enum
Type TpPos
    Ty As eTpPosTy
    R1 As Integer
    R2 As Integer
    C1 As Integer
    C2 As Integer
End Type
