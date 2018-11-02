Attribute VB_Name = "MVb__Ty"
Option Explicit
Function TyNm$(A)
TyNm = TypeName(A)
End Function

Function VbTySqlTy$(A As VbVarType, Optional IsMem As Boolean)
Select Case A
Case vbEmpty: VbTySqlTy = "Text(255)"
Case vbBoolean: VbTySqlTy = "YesNo"
Case vbByte: VbTySqlTy = "Byte"
Case vbInteger: VbTySqlTy = "Short"
Case vbLong: VbTySqlTy = "Long"
Case vbDouble: VbTySqlTy = "Double"
Case vbSingle: VbTySqlTy = "Single"
Case vbCurrency: VbTySqlTy = "Currency"
Case vbDate: VbTySqlTy = "Date"
Case vbString: VbTySqlTy = IIf(IsMem, "Memo", "Text(255)")
Case Else: Stop
End Select
End Function
