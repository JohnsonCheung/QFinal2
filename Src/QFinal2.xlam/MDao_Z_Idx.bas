Attribute VB_Name = "MDao_Z_Idx"
Option Explicit
Function IdxFny(A As DAO.Index) As String()
If IsNothing(A) Then Exit Function
IdxFny = ItrNy(A.Fields)
End Function

Function IdxIsSk(A As DAO.Index, T) As Boolean
If A.Name <> T Then Exit Function
IdxIsSk = A.Unique
End Function
Function IsUniqIdx(A As DAO.Index) As Boolean
If IsNothing(A) Then Exit Function
IsUniqIdx = A.Unique
End Function
