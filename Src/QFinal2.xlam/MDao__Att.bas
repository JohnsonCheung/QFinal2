Attribute VB_Name = "MDao__Att"
Option Explicit
Private Type AttRs
    TblRs As DAO.Recordset
    AttRs As DAO.Recordset
End Type

Sub AttClr(A)
DbAttClr CurDb, A
End Sub

Sub AttDrp(Att)
DbDrpAtt CurDb, Att
End Sub

Function AttExp$(A$, ToFfn$)
'Exporting the only file in Att
AttExp = DbAttExp(CurDb, A, ToFfn)
Debug.Print "-----"
Debug.Print "AttExp"
Debug.Print "Att   : "; A
Debug.Print "ToFfn : "; ToFfn
Debug.Print "Att is: Export to ToFfn"
End Function

Function AttExpFfn$(A$, AttFn$, ToFfn$)
AttExpFfn = DbAttExpFfn(CurDb, A, AttFn, ToFfn)
End Function

Function AttFfn$(A)
'Return Fst-Ffn-of-Att-A
AttFfn = RsMovFst(AttRs(A).AttRs)!Filename
End Function

Function AttFilCnt%(A)
AttFilCnt = DbAttFilCnt(CurDb, A)
End Function

Function AttFnAy(A) As String()
AttFnAy = DbAttFnAy(CurDb, "AA")
End Function

Function AttFny() As String()
AttFny = ItrNy(DbFstAttRs(CurDb).AttRs.Fields)
End Function

Function AttFstFn$(A)
AttFstFn = DbAttFstFn(CurDb, A)
End Function

Function AttHasOnlyOneFile(A$) As Boolean
AttHasOnlyOneFile = DbAttHasOnlyOneFile(CurDb, A)
End Function

Sub AttImp(A$, FmFfn$)
DbAttImp CurDb, A, FmFfn
End Sub

Function AttIsOld(A$, Ffn$) As Boolean
Dim T1 As Date, T2 As Date
T1 = AttTim(A)
T2 = FfnTim(Ffn)
Dim M$
M = FmtQQ("[Att] is ? in comparing with [file] using [Att-Tim] & [file-Tim].  Is Att [Older]?  (@AttIsOld)", IIf(T1 < T2, "older", "newer"), "?")
Msg "AttIsOld", M, A, Ffn, T1, T2, T1 < T2
AttIsOld = T1 < T2
End Function

Function AttLines$(A)
AttLines = DbAttLines(CurDb, A)
End Function

Function AttNy() As String()
AttNy = CurDbAttNy
End Function

Function AttRs(A) As AttRs
AttRs = DbAttRs(CurDb, A)
End Function

Function AttRsAttNm$(A As AttRs)
AttRsAttNm = A.TblRs!AttNm
End Function

Function AttRsExp$(A As AttRs, ToFfn)
'Export the only File in {AttRs} {ToFfn}
Dim Fn$, Ext$, T$, F2 As DAO.Field2
With A.AttRs
    If FfnExt(CStr(!Filename)) <> FfnExt(ToFfn) Then Stop
    Set F2 = .Fields("FileData")
End With
F2.SaveToFile ToFfn
AttRsExp = ToFfn
End Function

Function AttRsFilCnt%(A As AttRs)
AttRsFilCnt = RsNRec(A.AttRs)
End Function

Function AttRsFstFn$(A As AttRs)
With A.AttRs
    If .EOF Then
        If .BOF Then
            Msg "AttRsFstFn", "[AttNm] has no attachment files", AttRsAttNm(A)
            Exit Function
        End If
    End If
    .MoveFirst
    AttRsFstFn = !Filename
End With
End Function

Sub AttRsImp(A As AttRs, Ffn$)
Const CSub$ = "AttRsImp"
Dim F2 As Field2
Dim S&, T$
S = FfnSz(Ffn)
T = FfnDTim(Ffn)
Msg CSub, "[Att] is going to import [Ffn] with [Sz] and [Tim]", FldVal(A.TblRs!AttNm), Ffn, S, T
With A
    .TblRs.Edit
    With .AttRs
        If RsHasFldV(A.AttRs, "FileName", FfnFn(Ffn)) Then
            MsgDmp "Ffn is found in Att and it is replaced"
            .Edit
        Else
            MsgDmp "Ffn is not found in Att and it is imported"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TblRs.Fields!FilTim = FfnTim(Ffn)
    .TblRs.Fields!FilSz = FfnSz(Ffn)
    .TblRs.Update
End With
End Sub

Function AttRsLines$(A As AttRs)
Dim F As DAO.Field2, N%, Fn$
N = AttRsFilCnt(A)
If N <> 1 Then
    Msg "AttRsLines", "The [AttNm] should have one 1 attachment, but now [n-attachments]", AttRsAttNm(A), N
    Exit Function
End If
Fn = FfnExt(AttRsFstFn(A))
If Fn <> ".txt" Then
    Msg "AttRsLines", "The [AttNm] has [Att-Fn] not being [.txt].  Cannot return Lines", AttRsAttNm(A), Fn
    Exit Function
End If
AttRsLines = Fld2Lines(A.AttRs!FileData)
End Function

Function AttSz(A) As Date
AttSz = TsfV("Att", A, "FilSz")
End Function

Function AttTim(A$) As Date
AttTim = TfkV("Att", "FilTim", A)
End Function

Sub AttyDrp(Atty0)
DbDrpAtty CurDb, Atty0
End Sub

Function CurDbAttNy() As String()
CurDbAttNy = DbAttNy(CurDb)
End Function

Function DbAttRs(A As Database, Att) As AttRs
With DbAttRs
    Set .TblRs = A.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TblRs.EOF Then
        A.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TblRs = A.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .AttRs = .TblRs.Fields(0).Value
End With
End Function

Sub DbAttClr(A As Database, Att)
RsClr DbAttRs(A, Att).AttRs
End Sub

Function DbAttExp$(A As Database, Att, ToFfn)
'Exporting the first File in Att.
'If no file in att, error
'If any, export and return the
Dim N%
N = DbAttFilCnt(A, Att)
If N <> 1 Then
    Er "DbAttExp", "[Att] in [Db] has [FilCnt] which should be one.|Not export to [ToFfn].  (@DbAttExp)", _
        Att, A.Name, N, ToFfn
End If
DbAttExp = AttRsExp(DbAttRs(A, Att), ToFfn)
Msg "DbAttExp", "[Att] is exported [ToFfn] from [Db]", Att, ToFfn, DbNm(A)
End Function


Function DbAttFilCnt%(A As Database, Att)
'DbAttFilCnt = DbAttRs(A, Att).AttRs.RecordCount
DbAttFilCnt = AttRsFilCnt(DbAttRs(A, Att))
End Function

Function DbAttFstFn(A As Database, Att)
DbAttFstFn = AttRsFstFn(DbAttRs(A, Att))
End Function

Function DbAttHasOnlyOneFile(A As Database, Att$) As Boolean
Debug.Print "DbAttHasOnlyFile: " & DbAttRs(A, Att).AttRs.RecordCount
DbAttHasOnlyOneFile = DbAttRs(A, Att).AttRs.RecordCount = 1
End Function

Sub DbAttImp(A As Database, Att$, FmFfn$)
AttRsImp DbAttRs(A, Att), FmFfn
End Sub

Function DbAttLines$(A As Database, Att)
DbAttLines = AttRsLines(DbAttRs(A, Att))
End Function

Function DbAttExpFfn$(A As Database, Att$, AttFn$, ToFfn$)
Dim F2 As Field2, O$(), AttRs As AttRs
If FfnExt(AttFn) <> FfnExt(ToFfn) Then
    Stop
End If
If FfnIsExist(ToFfn) Then Stop
AttRs = DbAttRs(A, Att)
With AttRs
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !Filename = AttFn Then
                Set F2 = !FileData
                F2.SaveToFile ToFfn
                DbAttExpFfn = ToFfn
                Exit Function
            End If
            .MoveNext
        Wend
        Push O, "Database          : " & A.Name
        Push O, "AttKey            : " & Att
        Push O, "Missing-AttFn     : " & AttFn
        Push O, "AttKey-File-Count : " & AttRs.AttRs.RecordCount
        PushAy O, AyAddPfx(RsSy(AttRs.AttRs, "FileName"), "Fn in AttKey      : ")
        Push O, "Att-Table in Database has AttKey, but no Fn-of-Ffn"
        AyBrw O
        Stop
        Exit Function
    End With
End With
If IsNothing(F2) Then Stop
F2.SaveToFile ToFfn
DbAttExpFfn = ToFfn
End Function

Function DbAttFnAy(A As Database, Att$) As String()
Dim T As DAO.Recordset ' AttTblRs
Dim F As DAO.Recordset ' AttFldRs
Set T = DbAttTblRs(A, Att)
If T.EOF And T.BOF Then Exit Function
Set F = T.Fields("Att").Value
DbAttFnAy = RsSy(F, "FileName")
End Function

Function DbAttNy(A As Database) As String()
Q = "Select AttNm from Att order by AttNm": DbAttNy = RsSy(A.OpenRecordset(Q))
End Function

Function DbAttTblRs(A As Database, AttNm$) As DAO.Recordset
Set DbAttTblRs = A.OpenRecordset(FmtQQ("Select * from Att where AttNm='?'", AttNm))
End Function

Sub DbClrAtt(A As Database, Att$)
RsClr DbAttRs(A, Att).AttRs
End Sub

Sub DbDrpAtt(A As Database, Att)
A.Execute FmtQQ("Delete * from Att where AttNm='?'", Att)
End Sub

Sub DbDrpAtty(A As Database, Atty0)
AyDoPX CvNy(Atty0), "DbDrpAtt", A
End Sub

Function DbFstAttRs(A As Database) As AttRs
With DbFstAttRs
    Set .TblRs = A.TableDefs("Att").OpenRecordset
    Set .AttRs = .TblRs.Fields("Att").Value
End With
End Function

Function DbResAttFld(A As Database, ResNm) As DAO.Field2
Stop '
End Function

Private Sub Z_AttFnAy()
FbCurDb SampleFb_ShpRate
D AttFnAy("AA")
ClsCurDb
End Sub

Sub ZZ_AttImp()
Dim T$
T = TmpFt
StrWrt "sdfdf", T
AttImp "AA", T
Kill T
'T = TmpFt
'AttExpFfn "AA", T
'FtBrw T
End Sub

Sub ZZ_DbAttExpFfn()
Dim T$
T = TmpFx
DbAttExpFfn CurDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert FfnIsExist(T)
Kill T
End Sub

