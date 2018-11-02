Attribute VB_Name = "MDao__Ccm"
Option Explicit
Const CMod$ = "MDao__Ccm"
Private Sub Z_DbLnkCcm()
Dim CurDb As Database, IsLcl As Boolean
Set CurDb = FbDb(SampleFb_ShpRate)
IsLcl = True
GoSub Tst
Exit Sub
Tst:
    DbLnkCcm CurDb, IsLcl
    Return
End Sub
Sub DbLnkCcm(CurDb As Database, IsLcl As Boolean)
'Ccm stands for Space-[C]ir[c]umflex-accent
'CcmTbl is ^xxx table in CurDb (pgm-database),
'          which should be same stru as N:\..._Data.accdb @ xxx
'          and   data should be copied from N:\..._Data.accdb for development purpose
'At the same time, in CurDb, there will be xxx as linked table either
'  1. In production, linking to N:\..._Data.accdb @ xxx
'  2. In development, linking to CurDb @ ^xxx
'Notes:
'  The TarFb (N:\..._Data.accdb) of each CcmTbl may be diff
'      They are stored in Description of CcmTbl manual, it is edited manually during development.
'  those xxx table in CurDb will be used in the program.
'  and ^xxx is create manually in development and should be deployed to N:\..._Data.accdb
'  assume CurDb always have some ^xxx, otherwise throw
'This Sub is to re-link the xxx in given [CurDb] to
'  1. [CurDb] if [TarFb] is not given
'  2. [TarFb] if [TarFb] is given.
Const CSub = CMod$ & "DbLnkCcm"
Dim T$()  ' All ^xxx
    T = ZCcmTny(CurDb)
    If Sz(T) = 0 Then Er CSub, "No ^xxx table in [CurDb]", CurDb.Name 'Assume always
ZAss CurDb, T, IsLcl ' Chk if all T after rmv ^ is in TarFb
ZLnk CurDb, T, IsLcl
End Sub
Private Sub ZAss(CurDb As Database, CcmTny$(), IsLcl As Boolean)
Const CSub$ = CMod & ".ZAss"
If Not IsLcl Then ZAss2 CurDb, CcmTny: Exit Sub ' Asserting for TarFb is stored in CcmTny's description

'Asserting for TarFb = CurDb
Dim Miss$(): Miss = ZAss1(CurDb, CcmTny)
If Sz(Miss) = 0 Then Exit Sub
Er CSub, "[Some-missing-Tar-Tbl] in [CurDb] cannot be found according to given [CcmTny] in [CurDb]", Miss, CurDb.Name, CcmTny, CurDb.Name
End Sub
Private Function ZAss1(CurDb As Database, CcmTny$()) As String()
Dim N1$(): N1 = DbTny(CurDb)
Dim N2$(): N2 = AyRmvFstChr(CcmTny)
ZAss1 = AyMinus(N2, N1)
End Function

Private Sub ZAss2(CurDb As Database, CcmTny$())
'Throw if any Corresponding-Table in TarFb is not found
Dim O$(), T
For Each T In CcmTny
    PushIAy O, ZAss3(CurDb, T)
Next
AyBrwThw O
End Sub
Private Function ZAss3(CurDb As Database, CcmTbl) As String()
Dim TarFb$
    TarFb = DbtDes(CurDb, CcmTbl)
Select Case True
Case TarFb = "":            ZAss3 = MsgLy("[CcmTbl] in [CurDb] should have 'Des' which is TarFb, but this TarFb is blank", CcmTbl, CurDb.Name)
Case FfnNotExist(TarFb):    ZAss3 = MsgLy("[CcmTbl] in [CurDb] should have [Des] which is TarFb, but this TarFb does not exist", CcmTbl, CurDb.Name, TarFb)
Case Not FbHasTbl(TarFb, RmvFstChr(CcmTbl)):
    ZAss3 = MsgLy("[CcmTbl] in [CurDb] should have [Des] which is TarFb, but this TarFb does not exist [Tbl-RmvFstChr(CcmTbl)]", CcmTbl, CurDb.Name, TarFb, RmvFstChr(CcmTbl))
End Select
End Function

Private Sub ZLnk(CurDb As Database, CcmTny$(), IsLcl As Boolean)
Dim CcmTbl, TarFb$
TarFb = CurDb.Name
For Each CcmTbl In CcmTny
    If FstChr(CcmTbl) <> "^" Then Stop
    DbttLnkFb CurDb, RmvFstChr(CcmTbl), TarFb, CcmTbl
Next
End Sub
Private Function ZCcmTny(CurDb As Database) As String()
ZCcmTny = AyWhPfx(DbTny(CurDb), "^")
End Function

Private Sub Z_ZCcmTny()
Dim CurDb As Database
'
Set CurDb = FbDb(SampleFb_ShpRate)
Ept = SslSy("^CurYM ^IniRate ^IniRateH ^InvD ^InvH ^YM ^YMGR ^YMGRnoIR ^YMOH ^YMRate")
GoSub Tst
Exit Sub
Tst:
    Act = ZCcmTny(CurDb)
    C
    Return
End Sub
