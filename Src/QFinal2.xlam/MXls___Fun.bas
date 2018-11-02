Attribute VB_Name = "MXls___Fun"
Option Explicit

Sub BdrSetLin(A As Border)
With A
    .LineStyle = XlLineStyle.xlContinuous
    .Weight = xlMedium
    .Color = ColorConstants.vbBlack
End With
End Sub
Private Sub ZZ_AyabWs()
Dim A, B
A = SslSy("A B C D E")
B = SslSy("1 2 3 4 5")
WsVis AyabWs(A, B)
Stop
End Sub
Function FmlNy(A$) As String()
FmlNy = MacroNy(A, Bkt:="[]")
End Function


Function AyRgH(A, At As Range) As Range
Set AyRgH = SqRg(AySqH(A), At)
End Function

Function AyRgV(A, At As Range) As Range
Set AyRgV = SqRg(AySqV(A), At)
End Function

Function AyWs(A, Optional WsNm$) As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm)
SqRg AySqV(A), WsA1(O)
Set AyWs = O
End Function


Function AyabWs(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional LoNm$ = "AyAB") As Worksheet
Dim N&, AtA1 As Range, R As Range
N = Sz(A)
If N <> Sz(B) Then Stop
Set AtA1 = NewA1

AyRgH Array(N1, N2), AtA1
AyRgV A, AtA1.Range("A2")
AyRgV B, AtA1.Range("B2")
RgLo RgRCRC(AtA1, 1, 1, N + 1, 2)
Set AyabWs = AtA1.Parent
End Function

Function MaxCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(CurXls.Version = "16.0", 16384, 255)
End If
MaxCol = C
End Function

Function MaxRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(CurXls.Version = "16.0", 1048576, 65535)
End If
MaxRow = R
End Function

Sub AyPutCol(A, At As Range)
Dim Sq()
Sq = AySqV(A)
RgReSz(At, Sq).Value = Sq
End Sub

Sub AyPutLoCol(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'AyDmp LoFny(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
AyPutCol A, At
End Sub

Sub AyPutRow(A, At As Range)
Dim Sq()
Sq = AySqH(A)
RgReSz(At, Sq).Value = Sq
End Sub

Function DicWs(A As Workbook, Optional InclDicValTy As Boolean) As Worksheet
Set DicWs = DrsWs(DicDrs(A, InclDicValTy))
End Function

Function DicWb(A As Dictionary) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicAllKeyIsNm(A)
Ass DicAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook
Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set DicWb = O
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = LinesSqV(A(K))
Next
X: Set DicWb = O
End Function

Function N_SqH(N%) As Variant()
Dim O(), J%
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = J
Next
N_SqH = O
End Function

Function N_SqV(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = J
Next
N_SqV = O
End Function

Function N_ZerFill$(N, NDig%)
N_ZerFill = Format(N, String(NDig, "0"))
End Function


Function DicWsVis(A As Dictionary) As Worksheet
Dim O As Worksheet
   Set O = DicWs(A)
   WsVis O
Set DicWsVis = O
End Function


Function S1S2AyWs(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set S1S2AyWs = SqWs(S1S2AySq(A, Nm1, Nm2))
End Function


Function DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtWs = O
If Vis Then WsVis O
End Function
Sub DbOupWb(A As Database, Wb As Workbook, OupNmSsl$)
'OupNm is used for Table-Name-@*, WsCdNm-Ws*, LoNm-Tbl*
Dim Ay$(), OupNm
Ay = SslSy(OupNmSsl)
WbVdtOupNy Wb, Ay
Dim T$
For Each OupNm In Ay
    T = "@" & OupNm
    DbtOupWb A, T, Wb, OupNm
Next
End Sub

Function TblWs(T, Optional WsNm$ = "Data") As Worksheet
Set TblWs = LoWs(SqLo(TblSq(T), NewA1(WsNm)))
End Function

Function FbOupTblWb(A$) As Workbook
Dim O As Workbook
Set O = NewWb
AyDoABX FbOupTny(A), "WbAddWc", O, A
ItrDo O.Connections, "WcAddWs"
WbRfh O, A
Set FbOupTblWb = O
End Function


Sub FbRplWbLo(Fb$, A As Workbook)
Dim I, Lo As ListObject, Db As Database
Set Db = FbDb(Fb)
For Each I In WbOupLoAy(A)
    Set Lo = I
    DbtRplLo Db, "@" & Mid(Lo.Name, 3), Lo
Next
Db.Close
Set Db = Nothing
End Sub

Private Sub ZZ_FbOupTblWb()
Dim W As Workbook
'Set W = FbOupTblWb(WFb)
WbVis W
Stop
W.Close False
Set W = Nothing
End Sub



Sub FbWrtFx_zForExpOupTb(A$, Fx$)
FbOupTblWb(A).SaveAs Fx
End Sub


Function DbtAtLo(A As Database, T, At As Range, Optional UseWc As Boolean) As ListObject
Dim N$, Q As QueryTable
N = TblNm_LoNm(T)
If UseWc Then
    Set Q = RgWs(At).ListObjects.Add(SourceType:=0, Source:=FbAdoCnStr(A.Name), Destination:=At).QueryTable
    With Q
        .CommandType = xlCmdTable
        .CommandText = T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = T
        .Refresh BackgroundQuery:=False
    End With
    Exit Function
End If
Set DbtAtLo = LoSetNm(RgLo(DbtRg(A, T, At)), N)
End Function


Function DbtLo(A As Database, T, At As Range) As ListObject
Set DbtLo = LoSetNm(SqLo(DbtSq(A, T), At), TblNm_LoNm(T))
End Function

Sub DbtOupWb(A As Database, T, Wb As Workbook, OupNm)
'OupNm is used for WsCdNm-Ws*, LoNm-Tbl*
Dim Ws As Worksheet
Set Ws = WbWsCd(Wb, "WsO" & OupNm)
DbtPutWs A, T, Ws
End Sub


Function DbtRg(A As Database, T, At As Range) As Range
Set DbtRg = SqRg(DbtSq(A, T), At)
End Function

Function DbtRgByCn(A As Database, T, At As Range, Optional LoNm0$) As ListObject
If FstChr(T) <> "@" Then Stop
Dim LoNm$, Lo As ListObject
If LoNm0 = "" Then
    LoNm = "Tbl" & RmvFstChr(T)
Else
    LoNm = LoNm0
End If
Dim AtA1 As Range, CnStr, Ws As Worksheet
Set AtA1 = RgRC(At, 1, 1)
Set Ws = RgWs(At)
With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
        , _
        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
        , _
        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
        , _
        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
        , _
        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=" _
        , _
        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=AtA1).QueryTable '<---- At
        .CommandType = xlCmdTable
        .CommandText = Array(T) '<-----  T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = LoNm '<------------ LoNm
        .Refresh BackgroundQuery:=False
    End With

End Function


Function DbtPutFx(A As Database, T, Fx$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = FxWb(Fx)
Set Ws = WbWs(O, WsNm)
WsClrLo Ws
Stop ' LoNm need handle?
DbtPutWs A, T, WbWs(O, WsNm)
Set DbtPutFx = O
End Function

Sub DbtPutLo(A As Database, T, Lo As ListObject)
Dim Sq(), Drs As Drs, Rs As DAO.Recordset
Set Rs = DbtRs(A, T)
If Not IsEqAy(RsFny(Rs), LoFny(Lo)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    AyDmp RsFny(Rs)
    Debug.Print "--"
    Debug.Print "Lo"
    Debug.Print "--"
    AyDmp LoFny(Lo)
    Stop
End If
Sq = SqAddSngQuote(RsSq(Rs))
LoMin Lo
SqRg Sq, Lo.DataBodyRange
End Sub

Sub DbtPutWs(A As Database, T, Ws As Worksheet)
'Assume the WsCdNm is WsXXX and there will only 1 Lo with Name TblXXX
'Else stop
Dim Lo As ListObject
Set Lo = WsFstLo(Ws)

If Not HasPfx(Ws.CodeName, "WsO") Then Stop
If Ws.ListObjects.Count <> 1 Then Stop
If Mid(Lo.Name, 4) <> Mid(Ws.CodeName, 4) Then Stop
DbtPutLo A, T, Lo
End Sub


Function DbtRplLo(A As Database, T, Lo As ListObject, Optional ReSeqSpec$) As ListObject
Set DbtRplLo = SqRplLo(DbtReSeqSq(A, T, ReSeqSpec), Lo)
End Function

Function DbttWb(A As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
Set DbttWb = WbAddDbtt(O, A, TT, UseWc)
WbWs(O, "Sheet1").Delete
End Function

Sub DbttWrtFx(A As Database, TT, Fx$)
DbttWb(A, TT).SaveAs Fx
End Sub
Sub DbttFx(A As Database, Tny0, Fx$)
DbttWb(A, Tny0).SaveAs Fx
End Sub
Function FxDftWsNy(Fx$, WsNy0) As String()
Dim O$(): O = CvSy(WsNy0)
If Sz(O) = 0 Then
    FxDftWsNy = FxWsNy(Fx)
Else
    FxDftWsNy = O
End If
End Function

Sub TTFx(TT$, Fx$)
DbttFx CurDb, TT, Fx
End Sub

Sub LcSetTotLnk(A As ListColumn)
Dim R1 As Range, R2 As Range, R As Range, Ws As Worksheet
Set R = A.DataBodyRange
Set Ws = RgWs(R)
Set R1 = RgRC(R, 0, 1)
Set R2 = RgRC(R, R.Rows.Count + 1, 1)
Ws.Hyperlinks.Add Anchor:=R1, Address:="", SubAddress:=R2.Address
Ws.Hyperlinks.Add Anchor:=R2, Address:="", SubAddress:=R1.Address
R1.Font.ThemeColor = xlThemeColorDark1
End Sub

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs()
AyRgV Ly, WsA1(O)
Set LyWs = O
End Function
Private Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Function

