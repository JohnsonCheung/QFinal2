Attribute VB_Name = "MDao_Z_Db_DbInf"
Option Explicit
Sub DbInfBrw(A As Database)
AyBrw DsFmt(DbInfDs(A), 2000, DtBrkLinMapStr:="TblFld:Tbl")
End Sub


Function DbInfDs(A As Database) As Ds
Dim O As Ds
DsAddDt O, XLnk(A)
DsAddDt O, XTbl(A)
DsAddDt O, XTblF(A)
DsAddDt O, XPrp(A)
DsAddDt O, XFld(A, Tny)
O.DsNm = A.Name
DbInfDs = O
End Function

Private Sub Z_DbInfBrw()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
'DbInf(SampleDb_DutyPrepare).Brw
End Sub

Private Function XTbl(A As Database) As Dt
Dim TT, T$, Dry()
For Each TT In DbTny(A)
    T = TT
    Push Dry, Array(T, DbtRecCnt(A, T), DbtDes(A, T), DbtStru(A, T))
Next
Set XTbl = Dt("DbTbl", "Tbl RecCnt Des Stru", Dry)
End Function

Private Function XLnk(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
Set XLnk = Dt("DbLnk", "Tbl Connect", Dry)
End Function

Private Function XPrp(A As Database) As Dt
Dim Dry()
Set XPrp = Dt("DbPrp", "Prp Ty Val", Dry)
End Function
Private Function XFld(A As Database, Tny$()) As Dt
Dim Dry(), T
For Each T In Tny
Next
Set XFld = Dt("DbFld", "Tbl Fld Pk Ty Sz Dft Req Des", Dry)
End Function

Private Function XTblF(A As Database) As Dt
Dim Dry()
Dim T
For Each T In DbTny(A)
    PushAy Dry, XTblF1(A, T)
Next
Set XTblF = Dt("TblFld", "Tbl Fld", Dry)
End Function

Private Function XTblF1(A As Database, T) As Variant()
Dim O(), F, Dr(), Fny$()
Fny = DbtFny(A, T)
If Sz(Fny) = 0 Then Exit Function
Dim SeqNo%
SeqNo = 0
For Each F In Fny
    Erase Dr
    Push Dr, T
    Push Dr, SeqNo: SeqNo = SeqNo + 1
    PushAy Dr, DbtFFdInfDr(A, T, CStr(F))
    Push O, Dr
Next
XTblF1 = O
End Function
