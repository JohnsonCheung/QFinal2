Attribute VB_Name = "MSql_Ins"
Option Explicit
Function QIns_T_FF_DrAp(T, Fny0, ParamArray DrAp()) As String()
Dim Dr, Av(), Fny$()
Fny = CvNy(Fny0)
Av = DrAp
For Each Dr In Av
    PushI QIns_T_FF_DrAp, InsDrSql(T, Fny, Dr)
Next
End Function

Private Sub ZZ_DtInsDbSqy()
'Tmp1Tbl_Ens
Stop
Dim Db As Database: Set Db = SampleDb_DutyPrepare
Dim Dt As Dt: 'Dt = DbtDt(Db, "Tmp1")
Dim O$(): 'O = 'DtInsDbSqy(Db, Dt)
Stop
End Sub
