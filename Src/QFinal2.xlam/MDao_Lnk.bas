Attribute VB_Name = "MDao_Lnk"
Option Explicit

Function ColLnkExpFny(A$()) As String()
ColLnkExpFny = AyTakT3(A)
End Function

Sub DbDrpLnkTbl(A As Database)
DbttDrp A, DbLnkTny(A)
End Sub

Function DbLnkSpecImp(A As Database, LnkSpec$()) As String()
Dim O$(), J%, T$(), L$(), W$(), U%
LnkSpecAyAsg LnkSpec, T, L, W
U = UB(LnkSpec)
For J = 0 To U
    PushAy O, DbtChkCol(A, T(J), L(J))
Next
If Sz(O) > 0 Then DbLnkSpecImp = O: Exit Function
For J = 0 To U
    DbtImpMap A, T(J), L(J), W(J)
Next
DbLnkSpecImp = O
End Function

Function FilLin_Msg$(A$)
Dim FilNm$, Ffn$, L$
Ffn = A
FilNm = ShfT(Ffn)
If FfnIsExist(Ffn) Then Exit Function
FilLin_Msg = FmtQQ("[?] file not found [?]", FilNm, Ffn)
End Function


Function DbLnkVbly(A As Database) As String()
DbLnkVbly = AyMapPXSy(DbTny(A), "DbtLnkVbl", A)
End Function

Sub DrpLnkTbl()
CurDbDrpLnkTbl
End Sub


Function LinLnkCol(A$) As LnkCol
Dim Nm$, ShtTy$, Extnm$, Ty As DAO.DataTypeEnum
LinAsg2TRst A, Nm, ShtTy, Extnm
Extnm = RmvSqBkt(Extnm)
Ty = DaoShtTyStrTy(ShtTy)
Set LinLnkCol = LnkCol(Nm, Ty, IIf(Extnm = "", Nm, Extnm))
End Function

Function LnkColAy_ExtNy(A() As LnkCol) As String()
LnkColAy_ExtNy = OyPrpSy(A, "Extnm")
End Function

Function LnkColAy_Ny(A() As LnkCol) As String()
LnkColAy_Ny = OyPrpSy(A, "Nm")
End Function

Function LnkColIsEq(A As LnkCol, B As LnkCol) As Boolean
With A
    If .Extnm <> B.Extnm Then Exit Function
    If .Ty <> B.Ty Then Exit Function
    If .Nm <> B.Nm Then Exit Function
End With
LnkColIsEq = True
End Function

Function LnkColStr_LnkColAy(A) As LnkCol()
Dim Emp() As LnkCol, Ay$()
Ay = SplitVBar(A): If Sz(Ay) = 0 Then Stop
LnkColStr_LnkColAy = AyMapInto(Ay, "LinLnkCol", Emp)
End Function

Function LnkColStr_Ly(A$) As String()
Dim A1$(), A2$(), Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
A1 = LnkColAy_Ny(Ay)
A2 = AyAlignL(AyQuoteSqBkt(LnkColAy_ExtNy(Ay)))
Dim J%, O$()
For J = 0 To UB(A1)
    Push O, A2(J) & "  " & A1(J)
Next
LnkColStr_Ly = O
End Function

Sub LnkEdt()
SpnmEdt "Lnk"
End Sub

Sub LnkExp()
SpnmExp "Lnk"
End Sub

Private Function LnkFt$()
LnkFt = SpnmFt("Lnk")
End Function

Sub LnkImp()
SpnmImp "Lnk"
End Sub

Sub LnkSpecAyAsg(A$(), OTny$(), OLnkColStrAy$(), OWhBExprAy$())
Dim U%, J%
U = UB(A)
ReDim OTny(U)
ReDim OLnkColStrAy(U)
ReDim OWhBExprAy(U)
For J = 0 To U
    LSpecAsg A(J), OTny(J), OLnkColStrAy(J), OWhBExprAy(J)
Next
End Sub

Function LSpecLnkColStr$(A)
Dim L$
LSpecAsg A, , L
LSpecLnkColStr = L
End Function

Private Function NewLnkSpec(LnkSpec$) As LnkSpec
Dim Cln$():   Cln = LyCln(SplitCrLf(LnkSpec))
Dim AFx() As LnkAFil
Dim AFb() As LnkAFil
Dim ASw() As LnkASw

Dim FmFx() As LnkFmFil
Dim FmFb() As LnkFmFil
Dim FmIp() As String
Dim FmSw() As LnkFmSw
Dim FmWh() As LnkFmWh
Dim FmStu() As LnkFmStu

Dim IpFx() As LnkIpFil
Dim IpFb() As LnkIpFil
Dim IpS1() As String
Dim IpWs() As LnkIpWs

Dim StEle() As LnkStEle
Dim StExt() As LnkStExt
Dim StFld() As LnkStFld
    
    FmIp = SslSy(AyWhRmvTT(Cln, "FmIp", "|")(0))
'    FmFx = NewFmFil(AyWhRmvTT(Cln, "IpFx", "|"))
'    FmSw = NewFmSw(AyWhRmvT1(Cln, "IpSw"))
'    FmFb = NewFmFil(AyWhRmvTT(Cln, "IpFb", "|"))
'    FmWh = NewFmWh(AyWhRmvT1(Cln, "FmWh"))
    IpS1 = AyWhRmvTT(Cln, "IpS1", "|")
'    IpWs = NewIpWs(AyWhRmvTT(Cln, "IpWs", "|"))
Stop

With NewLnkSpec
    .AFx = AFx
    .AFb = AFb
    .ASw = ASw
    .FmFx = FmFx
    .FmFb = FmFb
    .FmIp = FmIp
    .FmSw = FmSw
    .FmStu = FmStu
    .FmWh = FmWh
    .IpFx = IpFx
    .IpFb = IpFb
    .IpS1 = IpS1
    .IpWs = IpWs
    .StEle = StEle
    .StExt = StExt
    .StFld = StFld
End With
End Function

Sub TTLnkFb(TT$, Fb$, Optional Fbtt)
DbttLnkFb CurDb, TT, Fb$, Fbtt
End Sub


Sub ZZ_LinLnkCol()
Dim A$, Act As LnkCol, Exp As LnkCol
A = "AA Txt XX"
Exp = LnkCol("AA", dbText, "AA")
GoSub Tst
Exit Sub
Tst:
Act = LinLnkCol(A)
Debug.Assert LnkColIsEq(Act, Exp)
Return
End Sub

Function DbInfLnkDt(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Set DbInfLnkDt = Dt("Lnk", CvNy("Tbl Connect"), Dry)
End Function

Function LnkColStr_ImpSql$(A$, T, Optional WhBExpr$)
Dim Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
LnkColStr_ImpSql = LnkColAy_ImpSql(Ay, T, WhBExpr)
End Function

Function LnkColAy_ImpSql$(A() As LnkCol, T, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "T must have first char = '>'"
    Stop
End If
Dim Ny$(), ExtNy$(), J%, O$(), S$, N$(), E$()
Ny = LnkColAy_Ny(A)
ExtNy = LnkColAy_ExtNy(A)
N = AyAlignL(Ny)
E = AyAlignL(AyQuoteSqBkt(ExtNy))
Erase O
For J = 0 To UB(Ny)
    If ExtNy(J) = Ny(J) Then
        Push O, FmtQQ("     ?    ?", Space(Len(E(J))), N(J))
    Else
        Push O, FmtQQ("     ? As ?", E(J), N(J))
    End If
Next
S = Join(O, "," & vbCrLf)
LnkColAy_ImpSql = FmtQQ("Select |?| Into [#I?]| From [?] |?", S, RmvFstChr(T), T, X.Wh(WhBExpr))
End Function

Function ColLnk_ImpSql$(A$(), Fm)
'data ColLnk = F T E
Dim Into$, Ny$(), Ey$()
If FstChr(Fm) <> ">" Then Stop
Into = "#I" & Mid(Fm, 2)
Ny = AyTakT1(A)
Ey = AyMapSy(A, "RmvTT")
'ColLnk_ImpSql = SelNyEyIntoFmSql$(Fm, Into, Ny, Ey)
End Function
