Attribute VB_Name = "MDao_Fb"
Option Explicit

Function FbCrt(A) As Database
Set FbCrt = DAO.DBEngine.CreateDatabase(A, dbLangGeneral)
End Function
Private Sub Z_FbBrw()
FbBrw SampleFb_Duty_Dta
End Sub
Sub FbBrw(A)
Static Acs As New Access.Application
Acs.OpenCurrentDatabase A
Acs.Visible = True
End Sub
Function FbDaoCn(A) As DAO.Connection
Set FbDaoCn = DBEngine.OpenConnection(A)
End Function

Function FbDb(A) As Database
Set FbDb = DAO.DBEngine.OpenDatabase(A)
End Function

Sub FbEns(A)
If Not FfnIsExist(A) Then FbCrt A
End Sub

Function FbOupTny(A) As String()
FbOupTny = AyWhPfx(FbTny(A), "@")
End Function

Function FbtFny(A$, T$) As String()
FbtFny = RsFny(DbqRs(FbDb(A), SelTblSql(T)))
End Function

Sub FbtStrAsg(FbtStr$, OFb$, OT$)
If FbtStr = "" Then
    OFb = ""
    OT = ""
    Exit Sub
End If
BrkAsg FbtStr, "].[", OFb, OT
If FstChr(OFb) <> "[" Then Stop
If LasChr(OT) <> "]" Then Stop
OFb = RmvFstChr(OFb)
OT = RmvLasChr(OT)
End Sub

Sub FbtDrp(A$, T$)
DbtDrp FbDb(A), T
End Sub
Sub ZZ_FbHasTbl()
Ass FbHasTbl(SampleFb_Duty_Dta, "SkuB")
End Sub


Sub ZZ_FbOupTny()
Dim Fb$
D FbOupTny(Fb)
End Sub

Sub ZZ_FbTny()
AyDmp FbTny(SampleFb_Duty_Dta)
End Sub
