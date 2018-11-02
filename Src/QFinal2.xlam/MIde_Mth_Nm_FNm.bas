Attribute VB_Name = "MIde_Mth_Nm_FNm"
Option Explicit
Sub FunFNm_BrkAsg(A$, OFunNm$, OPjNm$, OMdNm$)
With Brk(A, ":")
    OFunNm = .S1
    With Brk(.S2, ".")
        OPjNm = .S1
        OMdNm = .S2
    End With
End With
End Sub

Function FunFNm_MdDNm$(A)
FunFNm_MdDNm = Brk(A, ":").S2
End Function
