Attribute VB_Name = "MDao__Dic"
Option Explicit
Function DbtfDic(A As Database, T, F) As Dictionary
Set DbtfDic = RsDic(DbqRs(A, QSel_FF_Fm(F, T, IsDis:=True)))
End Function
Function DbqSyDic(A As Database, Q) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Sy
Set DbqSyDic = RsSyDic(DbqRs(A, Q))
End Function
Function DbqAyDic(A As Database, Q) As Dictionary
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Ay
Set DbqAyDic = RsAyDic(DbqRs(A, Q))
End Function
Function RsAyDic(A As DAO.Recordset) As Dictionary
Set RsAyDic = RsAyDicInto(A, EmpAy)
End Function
Function RsSyDic(A As DAO.Recordset) As Dictionary
Set RsSyDic = RsAyDicInto(A, EmpSy)
End Function

Function RsAyDicInto(A As DAO.Recordset, OInto) As Dictionary _
'Return a dictionary of Ay with Fst-Fld as Key and Snd-Fld as Ay
Dim O As New Dictionary, K, V, Ay
With A
    While Not .EOF
        K = .Fields(0).Value
        If IsNull(K) Then Stop
        V = .Fields(1).Value
        If Not O.Exists(K) Then
            O.Add K, AyCln(OInto)
        End If
        Ay = O(K)
        PushI Ay, V
        O(K) = Ay
        .MoveNext
    Wend
End With
Set RsAyDicInto = O
End Function
Function DbtSkDic(A As Database, T$) As Dictionary
Set DbtSkDic = DbtfDic(A, T, DbtSngSkFld(A, T))
End Function
Function RsDic(A As DAO.Recordset, Optional Sep$ = vbCrLf & vbCrLf) As Dictionary _
'Return a Dic from Fst-2-Fld-of-Rs-A with Fst as key and Snd as val _
'If Rs-A has only one fld, the Snd val will all be Empty _
'Assume Fst-Fld is unique, else throw error
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(0).Value
    V = A.Fields(1).Value
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, V
    End If
    A.MoveNext
Wend
Set RsDic = O
End Function

Function DbqDic(A As Database, Q, Optional Sep$ = vbCrLf & vbCrLf) As Dictionary
Set DbqDic = RsDic(DbqRs(A, Q))
End Function

Function RsIxDic(A As DAO.Recordset) As Dictionary
Dim O As New Dictionary, V
While Not A.EOF
    V = A.Fields(0).Value
    If Not O.Exists(V) Then
        O.Add V, O.Count
    End If
    A.MoveNext
Wend
Set RsIxDic = O
End Function

