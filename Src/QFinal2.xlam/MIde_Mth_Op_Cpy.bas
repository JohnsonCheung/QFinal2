Attribute VB_Name = "MIde_Mth_Op_Cpy"
Option Explicit
Function MthCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean) As Boolean
Const CSub$ = "MthCpy"
If MdHasMth(ToMd, A.Nm) Then
    If Not IsSilent Then
'        FunMsgLinDmp CSub, "[FmMth] is Found in [ToMd]", MthDNm(A), MdNm(ToMd)
    End If
    MthCpy = True
    Exit Function
End If
If ObjPtr(A.Md) = ObjPtr(ToMd) Then
    If Not IsSilent Then
'        FunMsgLinDmp CSub, "[FmMth] module cannot be same as [ToMd]", MthDNm(A), MdNm(ToMd)
    End If
    MthCpy = True
    Exit Function
End If
MdLinesApp ToMd, vbCrLf & MthLines(A)
If Not IsSilent Then
'    FunMsgLinDmp CSub, "[FmMth] is copied [ToMd]", MthDNm(A), MdNm(ToMd)
End If
End Function
