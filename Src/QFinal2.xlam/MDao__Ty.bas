Attribute VB_Name = "MDao__Ty"
Option Explicit
Function DaoShtTyStrTy(A$) As DAO.DataTypeEnum
Dim O$
Select Case A
Case "Lgc": O = DAO.DataTypeEnum.dbBoolean
Case "Dbl": O = DAO.DataTypeEnum.dbDouble
Case "Txt": O = DAO.DataTypeEnum.dbText
Case "Dte": O = DAO.DataTypeEnum.dbDate
Case "Byt": O = DAO.DataTypeEnum.dbByte
Case "Int": O = DAO.DataTypeEnum.dbInteger
Case "Lng": O = DAO.DataTypeEnum.dbLong
Case "Dec": O = DAO.DataTypeEnum.dbDecimal
Case "Cur": O = DAO.DataTypeEnum.dbCurrency
Case "Sng": O = DAO.DataTypeEnum.dbSingle
Case Else: Stop
End Select
DaoShtTyStrTy = O
End Function

Function VarDaoTy(A) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case VarType(A)
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbString: O = dbText
Case VbVarType.vbDate: O = dbDate
Case Else: Stop
End Select
VarDaoTy = O
End Function



Function DaoTyShtStr$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbByte: O = "Byt"
Case DAO.DataTypeEnum.dbLong: O = "Lng"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbDate: O = "Dte"
Case DAO.DataTypeEnum.dbText: O = "Txt"
Case DAO.DataTypeEnum.dbBoolean: O = "Yes"
Case DAO.DataTypeEnum.dbDouble: O = "Dbl"
Case DAO.DataTypeEnum.dbCurrency: O = "Cur"
Case DAO.DataTypeEnum.dbMemo: O = "Mem"
Case DAO.DataTypeEnum.dbAttachment: O = "Att"
Case DAO.DataTypeEnum.dbSingle: O = "Sng"
Case DAO.DataTypeEnum.dbDecimal: O = "Dec"
Case Else: O = "?" & A & "?"
End Select
DaoTyShtStr = O
End Function

Function DaoTySim(A As DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case A
Case _
   DAO.DataTypeEnum.dbBigInt, _
   DAO.DataTypeEnum.dbByte, _
   DAO.DataTypeEnum.dbCurrency, _
   DAO.DataTypeEnum.dbDecimal, _
   DAO.DataTypeEnum.dbDouble, _
   DAO.DataTypeEnum.dbFloat, _
   DAO.DataTypeEnum.dbInteger, _
   DAO.DataTypeEnum.dbLong, _
   DAO.DataTypeEnum.dbNumeric, _
   DAO.DataTypeEnum.dbSingle
   O = eNbr
Case _
   DAO.DataTypeEnum.dbChar, _
   DAO.DataTypeEnum.dbGUID, _
   DAO.DataTypeEnum.dbMemo, _
   DAO.DataTypeEnum.dbText
   O = eTxt
Case _
   DAO.DataTypeEnum.dbBoolean
   O = eLgc
Case _
   DAO.DataTypeEnum.dbDate, _
   DAO.DataTypeEnum.dbTimeStamp, _
   DAO.DataTypeEnum.dbTime
   O = eDte
Case Else
   O = eOth
End Select
DaoTySim = O
End Function

Function DaoTySqlTy$(A As DataTypeEnum, Optional Sz%, Optional Precious%)
Stop '
End Function

Function DaoTyStr$(T As DAO.DataTypeEnum)
Dim O$
Select Case T
Case DAO.DataTypeEnum.dbBoolean: O = "Boolean"
Case DAO.DataTypeEnum.dbDouble: O = "Double"
Case DAO.DataTypeEnum.dbText: O = "Text"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbByte: O = "Byte"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbLong: O = "Long"
Case DAO.DataTypeEnum.dbDouble: O = "Doubld"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbDecimal: O = "Decimal"
Case DAO.DataTypeEnum.dbCurrency: O = "Currency"
Case DAO.DataTypeEnum.dbSingle: O = "Single"
Case DAO.DataTypeEnum.dbAttachment: O = "Attachment"
Case DAO.DataTypeEnum.dbMemo: O = "Memo"
Case DAO.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case DAO.DataTypeEnum.dbBinary: O = "Binary"
Case DAO.DataTypeEnum.dbGUID: O = "GUID"
Case Else: Stop
End Select
DaoTyStr = O
End Function

Function SCShtTy_TyAy(A) As DAO.DataTypeEnum()
SCShtTy_TyAy = AyMapInto(SplitSC(A), "DaoShtTyStrTy", SCShtTy_TyAy)
End Function

