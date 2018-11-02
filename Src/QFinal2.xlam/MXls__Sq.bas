Attribute VB_Name = "MXls__Sq"
Option Explicit

Function SqA1(A, Optional WsNm$ = "Data") As Range
Set SqA1 = SqRg(A, NewA1(WsNm))
End Function

Function SqAddSngQuote(A)
Dim NC%, C%, R&, O
O = A
NC = UBound(A, 2)
For R = 1 To UBound(A, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
SqAddSngQuote = O
End Function

Sub SqBrw(A)
DryBrw SqDry(A)
End Sub

Function SqCol(A, C%) As Variant()
Dim O()
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C)
Next
SqCol = O
End Function

Function SqColInto(A, C%, OInto) As String()
Dim O
O = OInto
Erase O
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C%)
Next
SqColInto = O
End Function

Function SqColSy(A, C%) As String()
SqColSy = SqColInto(A, C, EmpSy)
End Function

Function SqDr(A, R&, Optional CnoAy) As Variant()
Dim mCnoAy%()
   Dim J%
   If IsMissing(CnoAy) Then
       ReDim mCnoAy(UBound(A, 2) - 1)
       For J = 0 To UB(mCnoAy)
           mCnoAy(J) = J + 1
       Next
   Else
       mCnoAy = CnoAy
   End If
Dim UCol%
   UCol = UB(mCnoAy)
Dim O()
   ReDim O(UCol)
   Dim C%
   For J = 0 To UCol
       C = mCnoAy(J)
       O(J) = A(R, C)
   Next
SqDr = O
End Function

Function SqDry(A) As Variant()
If Not IsArray(A) Then
    SqDry = Array(Array(A))
    Exit Function
End If
Dim R&
For R = 1 To UBound(A, 1)
    PushI SqDry, SqRowDr(A, R)
Next
End Function

Function SqFstColSy(A) As String()
SqFstColSy = SqColSy(A, 1)
End Function

Function SqInsDr(A, Dr, Optional Row& = 1)
Dim O(), C%, R&, NC%, NR&
NC = SqNCol(A)
NR = SqNRow(A)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = A(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = A(R, C)
    Next
Next
SqInsDr = O
End Function

Function SqIsEmp(Sq) As Boolean
SqIsEmp = True
On Error GoTo X
Dim A
If UBound(Sq, 1) < 0 Then Exit Function
If UBound(Sq, 2) < 0 Then Exit Function
SqIsEmp = False
Exit Function
X:
End Function

Function NewSq(R&, C&) As Variant()
ReDim NewSq(1 To R, 1 To C)
End Function

Function SqIsEq(A, B) As Boolean
Dim NR&, NC&
NR = UBound(A, 1)
NC = UBound(A, 2)
If NR <> UBound(B, 1) Then Exit Function
If NC <> UBound(B, 2) Then Exit Function
Dim R&, C&
For R = 1 To NR
    For C = 1 To NC
        If A(R, C) <> B(R, C) Then
            Exit Function
        End If
    Next
Next
SqIsEq = True
End Function

Function SqNCol&(A)
On Error Resume Next
SqNCol = UBound(A, 2)
End Function

Function SqNewA1(A, Optional WsNm$ = "Data") As Range
Dim A1 As Range
Set A1 = NewA1(WsNm)
Set SqNewA1 = SqRg(A, A1)
End Function

Function SqNewLo(A, Optional WsNm$ = "Data") As ListObject
Dim R As Range
Set R = SqNewA1(A, WsNm)
Set SqNewLo = RgLo(R)
End Function

Function SqNewWs(A) As Worksheet
Set SqNewWs = LoWs(SqNewLo(A))
End Function

Function SqNRow&(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Function

Sub SqRowDrSet(A, R, Dr)
Dim C%, V
For Each V In Dr
    C = C + 1
    A(R, C) = V
Next
End Sub


Function SqRowSel(A, R&, ColIxAy&()) As Variant()
Dim Ix
For Each Ix In ColIxAy
    Push SqRowSel, A(R, Ix + 1)
Next
End Function

Function SqRplLo(A, Lo As ListObject) As ListObject
Dim LoNm$, At As Range
LoNm = Lo.Name
Set At = Lo.Range
Lo.Delete
Set SqRplLo = LoSetNm(RgLo(SqRg(A, At)), LoNm)
End Function


Function SqSel(A, ColIxAy&()) As Variant()
Dim R&
For R = 1 To SqNRow(A)
    Push SqSel, SqRowSel(A, R, ColIxAy)
Next
End Function

Sub SqSetRow(OSq, R&, Dr)
Dim J%
For J = 0 To UB(Dr)
    OSq(R, J + 1) = Dr(J)
Next
End Sub

Function SqSyV(A) As String()
SqSyV = SqColSy(A, 1)
End Function

Function SqTranspose(A) As Variant()
Dim NR&, NC&
NR = SqNRow(A): If NR = 0 Then Exit Function
NC = SqNCol(A): If NC = 0 Then Exit Function
Dim O(), J&, I&
ReDim O(1 To NC, 1 To NR)
For J = 1 To NR
    For I = 1 To NC
        O(I, J) = A(J, I)
    Next
Next
SqTranspose = O
End Function

Function SqWs(A, Optional WsNm$) As Worksheet
Set SqWs = LoWs(SqLo(A, NewA1(WsNm)))
End Function
