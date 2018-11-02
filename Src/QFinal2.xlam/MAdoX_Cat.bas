Attribute VB_Name = "MAdoX_Cat"
Option Explicit
Private Function CatHasTbl(A As Catalog, T) As Boolean
CatHasTbl = ItrHasNm(A.Tables, T)
End Function

Private Function CatTny(A As Catalog) As String()
CatTny = ItrNy(A.Tables)
End Function

Private Function FbCat(A) As Catalog
Set FbCat = CnCat(FbCn(A))
End Function

Private Function CnCat(A As ADODB.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CnCat = O
End Function

Private Function FxCat(A) As Catalog
Set FxCat = CnCat(FxCn(A))
End Function

Private Function CatTblFny(A As ADOX.Table) As String()
CatTblFny = ItrNy(A.Columns)
End Function

Sub Z_FbTny()
AyDmp FbTny(SampleFb_Duty_Dta)
End Sub
Sub Z_FxWny()
AyDmp FxWny(SampleFx_KE24)
End Sub

Function FbHasTbl(A, T) As Boolean
FbHasTbl = AyHas(FbTny(A), T)
End Function
Function FxHasWs(A, W) As Boolean
FxHasWs = AyHas(FxWsNy(A), W)
End Function
Function FbTny(A) As String()
'FbTny = DbTny(FbDb(A))
'FbTny = CvSy(AyWhPredXPNot(CatTny(FbCat(A)), "HasPfx", "MSys"))
FbTny = AyWhExl(CatTny(FbCat(A)), "MSys* f_*_Data")
End Function

Function FxWsNy(A) As String()
FxWsNy = FxWny(A)
End Function

Function FxWny(A) As String()
Dim T$()
T = CatTny(FxCat(A))
FxWny = CvSy(AyWhDist(AyTakBefOrAll(AyRmvSngQuote(T), "$")))
End Function
Private Sub Z_FxwFny()
Dim W
For Each W In FxWny(SampleFx_KE24)
    D W & "<====================="
    D FxwFny(SampleFx_KE24, W)
Next
End Sub
Private Function FbtCatTbl(A, T) As ADOX.Table
Set FbtCatTbl = FbCat(A).Tables(T)
End Function
Function FbtFny(A, T) As String()
FbtFny = CatTblFny(FbtCatTbl(A, T))
End Function
Private Function FxwCatTbl(A, W) As ADOX.Table
Set FxwCatTbl = FxCat(A).Tables(QuoteSng(W) & "$")
End Function
Function FxwFny(A, W) As String()
FxwFny = CatTblFny(FxwCatTbl(A, W))
End Function
