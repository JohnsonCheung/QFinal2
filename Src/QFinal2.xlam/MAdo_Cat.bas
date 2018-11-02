Attribute VB_Name = "MAdo_Cat"
Option Explicit
Function CvCatTbl(A) As ADOX.Table
Set CvCatTbl = A
End Function

Function CatHasTbl(A As Catalog, T) As Boolean
CatHasTbl = ItrHasNm(A.Tables, T)
End Function

Function CatTny(A As Catalog) As String()
CatTny = ItrNy(A.Tables)
End Function

Function FbCat(A) As Catalog
Set FbCat = CnCat(FbCn(A))
End Function

Function CnCat(A As ADODB.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CnCat = O
End Function

Function FxCat(A) As Catalog
Set FxCat = CnCat(FxCn(A))
End Function

