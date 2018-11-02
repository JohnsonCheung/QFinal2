Attribute VB_Name = "MIde_Mth_Fb_Tmp_Tbl"
Option Explicit
Const CMod$ = "IdeMthFbTmpTbl."
Private X_W As Database
Const IMthMch$ = "MthMch"
Const OMthMd$ = "MthMd"
Const YFny$ = "MthNm MdNm MchTyStr MchStr"
Sub Z()
Z_ZZMthNy
End Sub

Function AAAModDic() As Dictionary
' Return a dic of Key=MdNm and Val=MdLines from MdMthNmDic("AAAMod")
' Use #MthMd : MthNm MdNm
Dim O As Dictionary
    Set O = DbqSyDic(W, "Select MdNm,MthNm from [#MthMd]")
    Dim K, MthNy$(), MthNmDic As Dictionary
    Set MthNmDic = MdMthNmDic(PjMd(Pj("QFinal"), "AAAMod"))
    For Each K In O.Keys
        If IsNull(K) Then Stop
        MthNy = CvSy(O(K)) ' The value of the dic is MthNy
        O(K) = DicKyJnVal(MthNmDic, MthNy) ' return a MdLines from MthNmDic using MthNy to look MthNmDic
    Next
Set AAAModDic = O
End Function

Sub RfhTmpMthMd()
RfhTmpMthNy
Const T$ = "#MthMd"
DbtDrp W, T
W.Execute "Create Table [#MthMd] (MthNm Text Not Null, MdNm Text(31),MthMchStr Text)"
'W.Execute CrtSkSql(T, "MthNm")
AyInsDbt MthNy, W, T ' MthNy is from #MthNy
XUpd
End Sub

Private Sub BrwTmpMthNy()
DbtBrw W, "#MthNy"
End Sub

Private Sub BrwTmpMthMd()
DbqBrw W, DbtFFSql(W, "#MthMd", "MdNm MthNm")
End Sub

Sub XUpd()
Const CSub$ = CMod & "XUpd"
Dim Rs As DAO.Recordset, XDic As Dictionary
Set XDic = DbqDic(W, "Select MthMchStr,ToMdNm from MthMch order by Seq Desc,Ty")
Set Rs = DbqRs(W, "Select MthNm,MthMchStr,MdNm from [#MthMd] where IIf(IsNull(MthMchStr),'',MthMchStr)=''")
While Not Rs.EOF
    DrUpdRs XDr(Rs.Fields("MthNm").Value, XDic), Rs
    Rs.MoveNext
Wend
Dim A%, B%, C%
A = DbtNRec(W, "#MthNy")
B = DbtNRec(W, "#MthMd")
C = DbtNRec(W, "#MthMd", "MdNm='AAAMod'")
C = DbqV(W, "Select count(*) from [#MthMd] where MdNm='AAAMod'")
Debug.Print CSub, "A: #MthNy-Cnt "; A
Debug.Print CSub, "B: #MthMd-Cnt "; B
Debug.Print CSub, "C: #MthMd-Wh-MdNm=AAAMod-Cnt "; C
DbqBrw MthDb, "select * from [#MthMd] where MthMchStr='' order by MthNm"
End Sub

Private Function XDr(MthNm, XDic As Dictionary) As Variant()
With StrDicMch(MthNm, XDic)
    If .Patn = "" Then
        XDr = Array(MthNm, "", "AAAMod")
    Else
        XDr = Array(MthNm, .Patn, .Rslt)
    End If
End With
End Function

Private Function MthNy() As String()
MthNy = DbtSy(W, "#MthNy")
End Function

Sub RfhTmpMthNy()
Const T$ = "#MthNy"
DbtDrp W, T
W.Execute "Create Table [#MthNy] (MthNm Text)"
'W.Execute CrtSkSql(T, "MthNm")
AyInsDbt MdMthNy(Md("AAAMod")), W, T
End Sub

Private Sub WDrpTmpMthNy()
DbtDrp W, "#MthNy"
End Sub

Private Function ZZMd() As CodeModule
Set ZZMd = Md("AAAMod")
End Function

Private Sub Z_ZZMthNy()
Brw ZZMthNy
End Sub

Private Function ZZMthNy() As String()
ZZMthNy = WMdMthNy(ZZMd)
End Function

Private Function WMdMthNy(A As CodeModule) As String()
WMthNy_EnsCache A
WMdMthNy = DbtfSy(W, "#MthNy", "MthNm")
End Function

Private Sub WMthNy_EnsCache(A As CodeModule)
Const T$ = "#MthNy"
If DbHasTbl(W, T) Then Exit Sub
W.Execute "Create Table [#MthNy] (MthNm Text(255) Not Null)"
'W.Execute CrtSkSql(T, "MthNm")
AyInsDbt MdMthNy(A), W, T
End Sub

Private Function WMchDic() As Dictionary
Static X As Dictionary
If IsNothing(X) Then Set X = DbqDic(W, "Select MthMchStr,ToMdNm from MthMch order by Seq,MthMchStr")
Set WMchDic = X
End Function

Private Function W() As Database
If IsNothing(X_W) Then Set X_W = MthDb
Set W = X_W
End Function
