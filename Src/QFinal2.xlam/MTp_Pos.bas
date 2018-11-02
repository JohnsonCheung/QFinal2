Attribute VB_Name = "MTp_Pos"
Option Explicit
Function TpPos_FmtStr$(A As TpPos)
Dim O$
With A
    Select Case .Ty
    Case ePosRCC
        O = FmtQQ("RCC(? ? ?) ", .R1, .C1, .C2)
    Case ePosRR
        O = FmtQQ("RR(? ?) ", .R1, .R2)
    Case ePosR
        O = FmtQQ("R(?)", .R1)
    Case Else
        'Er "TpPos_FmtStr", "Invalid {TpPos}", A.Ty
    End Select
End With
End Function
