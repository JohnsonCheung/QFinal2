Attribute VB_Name = "MVb__Var"
Option Explicit

Function VarCsv$(A)
Select Case True
Case IsStr(A): VarCsv = """" & A & """"
Case IsDte(A): VarCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case Else: VarCsv = IIf(IsNull(A), "", A)
End Select
End Function
