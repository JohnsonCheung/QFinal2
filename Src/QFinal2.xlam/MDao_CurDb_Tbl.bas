Attribute VB_Name = "MDao_CurDb_Tbl"
Option Explicit
Private Sub ZZ_TblFny()
AyDmp TblFny(">KE24")
End Sub

Private Sub ZZ_TblSq()
Dim A()
A = TblSq("@Oup")
Stop
End Sub

Function TblIsSys(T) As Boolean
TblIsSys = DbtIsSys(CurDb, T)
End Function

Sub TblAddPfx(T, Pfx$)
DbtAddPfx CurDb, T, Pfx
End Sub

Property Get TblAttDes$()
TblAttDes = DbtDes(CurDb, "Att")
End Property

Property Let TblAttDes(Des$)
DbtDes(CurDb, "Att") = Des
End Property

Sub TblBrw(T)
DoCmd.OpenTable T
End Sub

Sub TblCls(T)
DoCmd.Close acTable, T
End Sub

Sub TblCls_1(T)
DoCmd.Close acTable, T
End Sub

Property Let TblDes(T, Des$)
DbtDes(CurDb, T) = Des
End Property

Property Get TblDes$(T)
TblDes = DbtDes(CurDb, T)
End Property

Sub TblDrp(T)
DbtDrp CurDb, T
End Sub

Function TblErAyzCol(T, ColNy$(), DtaTyAy() As DAO.DataTypeEnum, Optional AddTblLinMsg As Boolean) As String()
Dim Fny$(), F, Fny1$(), Fny2$()
Fny = TblFny(T)
For Each F In ColNy
    If AyHas(Fny, F) Then
        Push F, Fny1
    Else
        Push F, Fny2
    End If
Next
Dim O$()
If Sz(Fny2) > 0 Then
    Dim J%
    For J = 0 To UB(ColNy)
        If AyHas(Fny2, ColNy(J)) Then
            If TfTy(T, ColNy(J)) <> DtaTyAy(J) Then
                Push O, "Column [?] has unexpected DataType[?].  It is expected to be [?]"
            End If
        End If
    Next
End If
If AddTblLinMsg Then
    Push O, ""
    
End If
End Function

Function TblFny(A) As String()
TblFny = DbtFny(CurDb, A)
End Function

Function TblHasFld(T As TableDef, F) As Boolean
TblHasFld = FdsHasFld(T.Fields, F)
End Function

Function TblImpSpec(T, LnkSpec$, Optional WhBExpr$) As TblImpSpec
Dim O As New TblImpSpec
Set TblImpSpec = O.Init(T, LnkSpec$, WhBExpr)
End Function

Function TblIsExist(T) As Boolean
TblIsExist = DbHasTbl(CurDb, T)
End Function

Function TkfV(T, K&, F$)
TkfV = DbtkfV(CurDb, T, K, F)
End Function

Function TblLnkFb(T, Fb$, Optional Fbt$) As String()
TblLnkFb = DbtLnkFb(CurDb, T, Fb$, Fbt)
End Function

Function TblNCol&(T)
TblNCol = DbtNCol(CurDb, T)
End Function

Function TblNm_LoNm$(T)
TblNm_LoNm = "T_" & RmvFstNonLetter(T)
End Function

Function TblNRec&(A)
TblNRec = SqlLng(FmtQQ("Select Count(*) from [?]", A))
End Function

Function TblNRow&(T, Optional WhBExpr$)
TblNRow = DbtNRec(CurDb, T, WhBExpr)
End Function

Sub TblOpn(T)
DoCmd.OpenTable T
End Sub

Function TblPk(T) As String()
TblPk = DbtPk(CurDb, T)
End Function

Sub TblPrm_Cpy_Fm_C_or_N(IsDev As Boolean)
CurDb.Execute "Delete * from Prm"
If IsDev Then
    CurDb.Execute "insert into Prm select * from Prm_C"
Else
    CurDb.Execute "insert into Prm select * from Prm_N"
End If
End Sub

Function TblRs(T) As DAO.Recordset
Set TblRs = DbtRs(CurDb, T)
End Function

Function TblScly(T) As String()
TblScly = DbtScly(CurDb, T)
End Function

Function TblSk(T) As String()
'Sk is secondary key.  Same name as the table and is unique
'Thw if there is key with same name as T, but not primary key.
'This is done as DbtSIdx
TblSk = DbtSk(CurDb, T)
End Function

Function TblSq(T) As Variant()
TblSq = DbtSq(CurDb, T)
End Function

Function TblSrc$(T)
TblSrc = CurDb.TableDefs(T).SourceTableName
End Function

Function TblStru$(T)
TblStru = DbtStru(CurDb, T)
End Function

Sub TblStruDmp(T)
D TblStru(T)
End Sub
