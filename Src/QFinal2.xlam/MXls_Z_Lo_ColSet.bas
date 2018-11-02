Attribute VB_Name = "MXls_Z_Lo_ColSet"
Option Explicit
Sub LcSetFml(A As ListObject, C, Fml$)
LcRg(A, C).Formula = Fml
End Sub

Sub LcSetCor(A As ListObject, C, Cor&)
LcRg(A, C).Interior.Color = Cor
End Sub

Sub LcSetFmt(A As ListObject, C, Fmt$)
LcRg(A, C).Formula = Fmt
End Sub

Sub LcSetLvl(A As ListObject, C, Optional Lvl As Byte = 2)
LcRg(A, C).EntireColumn.OutlineLevel = Lvl
End Sub

Sub LcSetWdt(A As ListObject, C, W%)
LcRg(A, C).EntireCumn.ColumnWidth = W
End Sub

Sub LcSetAlign(A As ListObject, C, B As XlHAlign)
LcRg(A, C).HorizontalAlignment = B
End Sub

Private Function LcRg(A As ListObject, C) As Range
Set LcRg = A.ListColumns(C).DataBodyRange
End Function

Sub LcSetBdrLeft(A As ListObject, C)
RgBdrL LcRg(A, C)
End Sub

Sub LcSetBdrRight(A As ListObject, C)
RgBdrL LcRg(A, C)
End Sub

Sub LcSetTot(A As ListObject, C, B As XlTotalsCalculation)
A.ListColumns(C).TotalsCalculation = B
End Sub
