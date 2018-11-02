Attribute VB_Name = "MDao_Lg"
Option Explicit
Private XSchm$()
Private X_W As Database
Private X_L As Database
Private X_Sess&
Private X_Msg&
Private X_Lg&
Private O$() ' Used by PthEntAyR
Sub Z()
Z_Lg
End Sub

Sub CurLgLis(Optional Sep$ = " ", Optional Top% = 50)
D CurLgLy(Sep, Top)
End Sub

Function CurLgLy(Optional Sep$ = " ", Optional Top% = 50) As String()
CurLgLy = RsLy(CurLgRs(Top), Sep)
End Function
Private Function RsLy(A As DAO.Database, Sep$) As String()

End Function
Function CurLgRs(Optional Top% = 50) As DAO.Recordset
Set CurLgRs = L.OpenRecordset(FmtQQ("Select Top ? x.*,Fun,MsgTxt from Lg x left join Msg a on x.Msg=a.Msg order by Sess desc,Lg", Top))
End Function

Sub CurSessLis(Optional Sep$ = " ", Optional Top% = 50)
D CurSessLy(Sep, Top)
End Sub

Function CurSessLy(Optional Sep$, Optional Top% = 50) As String()
CurSessLy = RsLy(CurSessRs(Top), Sep)
End Function

Function CurSessRs(Optional Top% = 50) As DAO.Recordset
Set CurSessRs = L.OpenRecordset(FmtQQ("Select Top ? * from sess order by Sess desc", Top))
End Function
Private Function DbqV(A As Database, Q)
Stop
End Function
Private Function CvSess&(A&)
If A > 0 Then CvSess = A: Exit Function
CvSess = DbqV(L, "select Max(Sess) from Sess")
End Function
Private Sub EnsMsg(Fun$, MsgTxt$)
With L.TableDefs("Msg").OpenRecordset
    .Index = "Msg"
    .Seek "=", Fun, MsgTxt
    If .NoMatch Then
        .AddNew
        !Fun = Fun
        !MsgTxt = MsgTxt
        X_Msg = !Msg
        .Update
    Else
        X_Msg = !Msg
    End If
End With
End Sub


Private Sub EnsSess()
If X_Sess > 0 Then Exit Sub
With L.TableDefs("Sess").OpenRecordset
    .AddNew
    X_Sess = !Sess
    .Update
    .Close
End With
End Sub

Private Function L() As Database
On Error GoTo X
If IsNothing(X_L) Then
    LgOpn
End If
Set L = X_L
Exit Function
X:
Dim Er$, ErNo%
ErNo = Err.Number
Er = Err.Description
If ErNo = 3024 Then
    'LgSchmImp
    LgCrt_v1
    LgOpn
    Set L = X_L
    Exit Function
End If
NyLyDmp "Err Er#", Er, ErNo
Stop
End Function

Sub Lg(Fun$, MsgTxt$, ParamArray Ap())
EnsSess
EnsMsg Fun, MsgTxt
WrtLg Fun, MsgTxt
Dim Av(): Av = Ap
If Sz(Av) = 0 Then Exit Sub
Dim J%, V
With L.TableDefs("LgV").OpenRecordset
    For Each V In Av
        .AddNew
        !Lines = VarLines(V)
        .Update
    Next
    .Close
End With
End Sub

Private Sub RsAsg(A As DAO.Recordset, ParamArray OAp())

End Sub

Sub LgAsg(A&, OSess&, ODTim$, OFun$, OMsgTxt$)
Dim Q$
Q = FmtQQ("select Fun,MsgTxt,Sess,x.CrtTim from Lg x inner join Msg a on x.Msg=a.Msg where Lg=?", A)
Dim D As Date
RsAsg L.OpenRecordset(Q), OFun, OMsgTxt, OSess, D
ODTim = DteDTim(D)
End Sub

Sub LgBeg()
Lg ".", "Beg"
End Sub

Sub LgBrw()
FtBrw LgFt
End Sub

Sub LgCls()
On Error GoTo Er
X_L.Close
Er:
Set X_L = Nothing
End Sub
Private Sub FbCrt(A)
Stop '
End Sub
Private Function FbDb(A) As Database
Stop '
End Function
Private Sub TdAddId(A As DAO.TableDef)
End Sub
Sub LgCrt()
'FbCrt LgFb
'Dim Db As Database, T As DAO.TableDef
'Set Db = FbDb(LgFb)
''
'Set T = New DAO.TableDef
'T.Name = "Sess"
'TdAddId T
'TdAddStamp T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "Msg"
'TdAddId T
'TdAddTxtFld T, "Fun"
'TdAddTxtFld T, "MsgTxt"
'TdAddStamp T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "Lg"
'TdAddId T
'TdAddLngFld T, "Sess"
'TdAddLngFld T, "Msg"
'TdAddStamp T, "Dte"
'Db.TableDefs.Append T
''
'Set T = New DAO.TableDef
'T.Name = "LgV"
'TdAddId T
'TdAddLngFld T, "Lg"
'TdAddLngTxt T, "Val"
'Db.TableDefs.Append T
'
'DbttCrtPk Db, "Sess Msg Lg LgV"
'DbtCrtSk Db, "Msg", "Fun MsgTxt"
End Sub

Sub LgCrt_v1()
Dim Fb$
Fb = LgFb
If FfnIsExist(Fb) Then Exit Sub
'DbCrtSchm FbCrt(Fb), LgSchmLines
End Sub

Function LgDb() As Database
Set LgDb = L
End Function

Sub LgDbBrw()
'Acs.OpenCurrentDatabase LgFb
'AcsVis Acs
End Sub

Sub LgEnd()
Lg ".", "End"
End Sub

Function LgFb$()
LgFb = LgPth & LgFn
End Function

Function LgFn$()
LgFn = "Lg.accdb"
End Function

Function LgFt$()
Stop '
End Function
Private Sub X(A$)
PushI XSchm, A
End Sub
Function LgSchm() As String()
If Sz(XSchm) = 0 Then
X "E Mem | Mem Req AlwZLen"
X "E Txt | Txt Req"
X "E Crt | Dte Req Dft=Now"
X "E Dte | Dte"
X "E Amt | Cur"
X "F Amt * | *Amt"
X "F Crt * | CrtDte"
X "F Dte * | *Dte"
X "F Txt * | Fun * Txt"
X "F Mem * | Lines"
X "T Sess | * CrtDte"
X "T Msg  | * Fun *Txt | CrtDte"
X "T Lg   | * Sess Msg CrtDte"
X "T LgV  | * Lg Lines"
X "D . Fun | Function name that call the log"
X "D . Fun | Function name that call the log"
X "D . Msg | it will a new record when Lg-function is first time using the Fun+MsgTxt"
X "D . Msg | ..."
End If
LgSchm = XSchm
End Function

Sub LgKill()
LgCls
If FfnIsExist(LgFb) Then Kill LgFb: Exit Sub
Debug.Print "LgFb-[" & LgFb & "] not exist"
End Sub

Function LgLinesAy(A&) As Variant()
Dim Q$
Q = FmtQQ("Select Lines from LgV where Lg = ? order by LgV", A)
'LgLinesAy = RsAy(L.OpenRecordset(Q))
End Function

Sub LgLis(Optional Sep$ = " ", Optional Top% = 50)
CurLgLis Sep, Top
End Sub

Function LgLy(A&) As String()
Dim Fun$, MsgTxt$, DTim$, Sess&, Sfx$
LgAsg A, Sess, DTim, Fun, MsgTxt
Sfx = FmtQQ(" @? Sess(?) Lg(?)", DTim, Sess, A)
LgLy = FunMsgLy(Fun & Sfx, MsgTxt, LgLinesAy(A))
End Function

Private Sub LgOpn()
Set X_L = FbDb(LgFb)
End Sub

Function LgPth$()
Static Y$
'If Y = "" Then Y = PgmPth & "Log\": PthEns Y
LgPth = Y
End Function

Function Lik(A, K) As Boolean
Lik = A Like K
End Function


Sub LnoCnt_Dmp(A As LnoCnt)
Debug.Print LnoCnt_Str(A)
End Sub

Function LnoCnt_Str$(A As LnoCnt)
LnoCnt_Str = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Function LTrimWhite$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhiteChr(Mid(A, J, 1)) Then Exit For
    Next
LTrimWhite = Left(A, J)
End Function

Sub SessBrw(Optional A&)
AyBrw SessLy(CvSess(A))
End Sub

Function SessLgAy(A&) As Long()
Dim Q$
Q = FmtQQ("select Lg from Lg where Sess=? order by Lg", A)
'SessLgAy = DbqLngAy(L, Q)
End Function

Sub SessLis(Optional Sep$ = " ", Optional Top% = 50)
CurSessLis Sep, Top
End Sub

Function SessLy(Optional A&) As String()
Dim LgAy&()
LgAy = SessLgAy(A)
SessLy = AyOfAy_Ay(AyMap(LgAy, "LgLy"))
End Function

Function SessNLg%(A&)
SessNLg = DbqV(L, "Select Count(*) from Lg where Sess=" & A)
End Function

Private Sub WrtLg(Fun$, MsgTxt$)
With L.TableDefs("Lg").OpenRecordset
    .AddNew
    !Sess = X_Sess
    !Msg = X_Msg
    X_Lg = !Lg
    .Update
End With
End Sub

Private Sub Z_Lg()
LgKill
Debug.Assert Dir(LgFb) = ""
LgBeg
Debug.Assert Dir(LgFb) = LgFn
End Sub
