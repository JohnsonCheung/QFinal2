Attribute VB_Name = "MDao_Z_Rs"
Option Explicit
Private Sub ZZ_RsAsg()
Dim Y As Byte, M As Byte
RsAsg TblRs("YM"), Y, M
Stop
End Sub

Function RsNoRec(A As DAO.Recordset) As Boolean
RsNoRec = Not RsAny(A)
End Function

Function RsAny(A As DAO.Recordset) As Boolean
If A.EOF Then Exit Function
If A.BOF Then Exit Function
RsAny = True
End Function

Sub RsAsg(A As DAO.Recordset, ParamArray OAp())
Dim F As DAO.Field, J%, U%
Dim Av(): Av = OAp
U = UB(Av)
For Each F In A.Fields
    OAp(J) = F.Value
    If J = U Then Exit Sub
    J = J + 1
Next
End Sub

Function RsAy(A As DAO.Recordset, Optional F0) As Variant()
RsAy = RsAyInto(A, RsAy, F0)
End Function

Function RsAyInto(A As DAO.Recordset, OInto, Optional F0)
Dim O: O = OInto: Erase O
Dim F
F = DftF0(F0)
With A
    If .EOF Then RsAyInto = O: Exit Function
    .MoveFirst
    While Not .EOF
        Push O, .Fields(F).Value
        .MoveNext
    Wend
End With
RsAyInto = O
End Function

Sub RsBrw(A As DAO.Recordset)
DrsBrw RsDrs(A)
End Sub

Sub RsBrw_zSingleRec(A As DAO.Recordset)
AyBrw RsLy_zSingleRec(A)
End Sub

Sub RsClr(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub

Function RsCsv$(A As DAO.Recordset)
RsCsv = FdsCsv(A.Fields)
End Function

Function RsCsvLy(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    For Each F In A.Fields
        Dr(I) = VarCsv(F.Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLy = O
End Function

Function RsCsvLyByFny0(A As DAO.Recordset, Fny0) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = CvNy(Fny0)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = VarCsv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLyByFny0 = O
End Function

Sub RsDmp(A As Recordset)
AyDmp RsCsvLy(A)
A.MoveFirst
End Sub

Sub RsDmpByFny0(A As Recordset, Fny0)
AyDmp RsCsvLyByFny0(A, Fny0)
A.MoveFirst
End Sub

Function RsDr(A As DAO.Recordset) As Variant()
RsDr = FdsDr(A.Fields)
End Function

Function RsKyDr(A As DAO.Recordset, Ky0) As Variant()
RsKyDr = FdsKyDr(A.Fields, Ky0)
End Function

Function RsDrs(A As DAO.Recordset) As Drs
Set RsDrs = Drs(RsFny(A), RsDry(A))
End Function

Function RsDry(A As DAO.Recordset) As Variant()
If Not RsAny(A) Then Exit Function
With A
    .MoveFirst
    While Not .EOF
        PushI RsDry, FdsDr(.Fields)
        .MoveNext
    Wend
    .MoveFirst
End With
End Function

Function RsFny(A As DAO.Recordset) As String()
RsFny = ItrNy(A.Fields)
End Function

Function RsHasFldV(A As DAO.Recordset, F$, V) As Boolean
With A
    If .BOF Then
        If .EOF Then Exit Function
    End If
    .MoveFirst
    While Not .EOF
        If .Fields(F) = V Then RsHasFldV = True: Exit Function
        .MoveNext
    Wend
End With
End Function

Function RsIntAy(A As DAO.Recordset, Optional F) As Integer()
RsIntAy = RsAyInto(A, RsIntAy)
End Function

Function RsIsBrk(A As DAO.Recordset, GpKy$(), LasVy()) As Boolean
RsIsBrk = Not IsEqAy(RsKyDr(A, GpKy), LasVy)
End Function

Function RsLin$(A As DAO.Recordset, Optional Sep$ = " ")
RsLin = Join(RsDr(A), Sep)
End Function

Function RsLngAy(A As DAO.Recordset, Optional FldNm$) As Long()
RsLngAy = RsAyInto(A, FldNm, RsLngAy)
End Function

Function RsLy(A As DAO.Recordset, Optional Sep$ = " ") As String()
Dim O$()
With A
    Push O, Join(RsFny(A), Sep)
    While Not .EOF
        Push O, RsLin(A, Sep)
        .MoveNext
    Wend
End With
RsLy = O
End Function

Function RsLy_zSingleRec(A As DAO.Recordset)
RsLy_zSingleRec = NyAvLy(RsFny(A), RsDr(A), 0)
End Function

Function RsMovFst(A As DAO.Recordset) As DAO.Recordset
A.MoveFirst
Set RsMovFst = A
End Function

Function RsNRec&(A As DAO.Recordset)
Dim O&
With A
    .MoveFirst
    While Not .EOF
        O = O + 1
        .MoveNext
    Wend
    .MoveFirst
End With
RsNRec = O
End Function


Sub RsPutSq(A As DAO.Recordset, Sq, R&, Optional NoTxtSngQ As Boolean)
FdsPutSq A.Fields, Sq, R, NoTxtSngQ
End Sub

Function RsSq(A As DAO.Recordset) As Variant()
RsSq = DrySq(RsDry(A))
End Function

Function RsStrCol(A As DAO.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
End With
RsStrCol = O
End Function

Function RsStru$(A As DAO.Recordset)
Dim O$(), F As DAO.Field2
For Each F In A.Fields
    PushI O, FdStr(F)
Next
RsStru = JnCrLf(O)
End Function
Function Nz(A, B)
Nz = IIf(IsNull(A), B, A)
End Function
Function RsInto(A As Recordset, F, OInto)
RsInto = AyCln(OInto)
While Not A.EOF
    PushI RsInto, Nz(A(F).Value, Empty)
    A.MoveNext
Wend
End Function

Function RsSy(A As Recordset, Optional F = 0) As String()
RsSy = RsInto(A, F, EmpSy)
End Function

Function RsV(A As DAO.Recordset)
If RsAny(A) Then RsV = A.Fields(0).Value
End Function

Property Let RsFldVal(A As DAO.Recordset, F, V)
With A
    .Edit
    .Fields(F).Value = V
    .Update
End With
End Property

Property Get RsFldVal(A As DAO.Recordset, F)
With A
    If .EOF Then Exit Property
    If .BOF Then Exit Property
    RsFldVal = .Fields(F).Value
End With
End Property

Function RsVy(A As DAO.Recordset, Optional Ky0) As Variant()
RsVy = FdsVy(A.Fields, Ky0)
End Function

Function RsXTSz$(A As DAO.Recordset)
If A.Fields(0).Type <> DAO.dbDate Then Stop
If A.Fields(1).Type <> DAO.dbLong Then Stop
If RsNoRec(A) Then Exit Function
RsXTSz = DteDTim(A.Fields(0).Value) & "." & A.Fields(1).Value
End Function
Sub ApAddRs(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
DrInsRs Dr, Rs
End Sub


Sub DrInsRs(A, Rs As DAO.Recordset)
Rs.AddNew
DrSetRs A, Rs
Rs.Update
End Sub

Sub ApUpdRs(Rs As DAO.Recordset, ParamArray Ap())
Dim Dr(): Dr = Ap
DrUpdRs Dr, Rs
End Sub

Sub DrSetRs(Dr, Rs As DAO.Recordset)
Dim V, J%
For Each V In Dr
    If IsEmpty(V) Then
        Rs(J).Value = Rs(J).DefaultValue
    Else
        Rs(J).Value = V
    End If
    J = J + 1
Next
End Sub


Sub DrUpdRs(A(), Rs As DAO.Recordset)
If Sz(A) = 0 Then Exit Sub
Rs.Edit
DrSetRs A, Rs
Rs.Update
End Sub
