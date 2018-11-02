Attribute VB_Name = "MXls_Z_Qt"
Option Explicit
Function QtFbtStr$(A As QueryTable)
If IsNothing(A) Then Exit Function
Dim Ty As XlCmdType, Tbl$, CnStr$
With A
    Ty = .CommandType
    If Ty <> xlCmdTable Then Exit Function
    Tbl = .CommandText
    CnStr = .Connection
End With
QtFbtStr = FmtQQ("[?].[?]", CnStr_DtaSrc(CnStr), Tbl)
End Function
