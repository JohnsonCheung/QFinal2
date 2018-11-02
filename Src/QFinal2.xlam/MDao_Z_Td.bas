Attribute VB_Name = "MDao_Z_Td"
Option Explicit
Sub TdAddId(A As DAO.TableDef)
A.Fields.Append NewFd_zId(A.Name)
End Sub

Sub TdAddLngFld(A As DAO.TableDef, FF0)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append NewFd(F, dbLong)
Next
End Sub

Sub TdAddLngTxt(A As DAO.TableDef, FF0)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append NewFd(F, dbMemo)
Next
End Sub

Sub TdAddStamp(A As DAO.TableDef, F$)
A.Fields.Append NewFd(F, DAO.dbDate, Dft:="Now")
End Sub

Sub TdAddTxtFld(A As DAO.TableDef, FF0, Optional Sz As Byte = 255)
Dim F
For Each F In CvNy(FF0)
    A.Fields.Append NewFd(F, dbText, Sz)
Next
End Sub

Function TdFdScly(A As DAO.TableDef) As String()
Dim N$
N = A.Name & ";"
TdFdScly = AyAddPfx(ItrMapSy(A.Fields, "FdScl"), N)
End Function

Function TdScl$(A As DAO.TableDef)
TdScl = ApScl(A.Name, AddLbl(A.OpenRecordset.RecordCount, "NRec"), AddLbl(A.DateCreated, "CrtDte"), AddLbl(A.LastUpdated, "UpdDte"))
End Function

Function TdScly(A As DAO.TableDef) As String()
TdScly = AyAdd(Sy(TdScl(A)), TdFdScly(A))
End Function

Function TdScly_AddPfx(A) As String()
Dim O$(), U&, J&, X
U = UB(A)
If U = -1 Then Exit Function
ReDim O(U)
For Each X In AyNz(A)
    O(J) = IIf(J = 0, "Td;", "Fd;") & X
    J = J + 1
Next
TdScly_AddPfx = O
End Function
