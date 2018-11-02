Attribute VB_Name = "MDao___New"
Option Explicit
Public Q$

Function LnkCol(Nm$, Ty As DAO.DataTypeEnum, Extnm$) As LnkCol
Dim O As New LnkCol
Set LnkCol = O.Init(Nm, Ty, Extnm)
End Function
