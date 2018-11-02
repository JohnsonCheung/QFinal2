Attribute VB_Name = "MDao_DML_SngFldSkTbl_Operation"
Option Explicit
Sub Z()
Z_AyInsDbt
End Sub

Private Sub Z_AyInsDbt()
Dim Db As Database
Set Db = TmpDb
Db.Execute "Create Table [Tbl-AA] ([Fld-AA] Int)"
Db.Execute CrtSkSql("Tbl-AA", "Fld-AA")
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(1)"
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(3)"
Db.Execute "Insert Into [Tbl-AA] ([Fld-AA]) Values(5)"
AyInsDbt Array(1, 2, 3, 4, 5, 6, 7), Db, "Tbl-AA"
Act = DbtfIntAy(Db, "Tbl-AA", "Fld-AA")
Ept = ApIntAy(1, 2, 3, 4, 5, 6, 7)
C
End Sub

Function DbtSngSkFld$(A As Database, T$)
Const CSub$ = "DbtSngSkFld"
Dim Sk$(): Sk = DbtSk(A, T)
If Sz(Sk) = 0 Then DbtSngSkFld = Sk(0): Exit Function
Er CSub, "Given [Db]-[T] has [Sk] of [Sz]<>1", DbNm(A), T, Sk, Sz(Sk)
End Function



Sub AyDltDbt(A, Db As Database, T) _
'Delete Db-T record for those record's Sk not in Ay-A, _
'Assume T has single-fld-sk
Dim Q$, Sk$, ExcessAy
Dim X
If Sz(X) > 0 Then
    DbSqyRun Db, DltInAySqy(T, Sk, ExcessAy, Q)
End If
End Sub
Private Function ExcessEle(A, Db As Database, T) _
'Return Sub-Ay from Ay-A for those element not in Db-T sk _
'Asume, Db-T has single-fld-sk
ExcessEle = AyCln(A)
Dim X, Dic As Dictionary
For Each X In A
    If Not Dic.Exists(X) Then
        PushI ExcessEle, X
    End If
Next
End Function
Sub AyInsDbt(A, Db As Database, T$) _
'Insert Ay-A into Db-T _
'Assume T has single-fld-sk and can be inserted by just giving such Sk-value
Const CSub$ = "AyInsDbt"
Dim SkFld$
    Dim Sk$(): Sk = DbtSk(Db, T)
    If Sz(Sk) <> 1 Then ErWh CSub, "Dbt does not have Sk of sz=1", "Db T Sk-Sz Sk", DbNm(Db), T, Sz(Sk), Sk
    SkFld = Sk(0)

ZCrtTmpTbl:
    Dim TmpTbl$
    TmpTbl = TmpNm
    DbtfCrtTmpTbl Db, T, SkFld, TmpTbl

ZInsAyToTmpTbl:
    Dim X
    With DbtRs(Db, TmpTbl)
        For Each X In AyNz(A)
            .AddNew
            .Fields(0).Value = X
            .Update
        Next
    End With

ZInsIntoTarTbl:
    Q = FmtQQ("Insert Into [?] select x.[?] from [?] x left join [?] a on x.[?] = a.[?] where a.[?] is null", _
        T, SkFld, TmpTbl, T, SkFld, SkFld, SkFld)
    Db.Execute Q
    
ZDltTmpTbl:
    DbtDrp Db, TmpTbl
End Sub

