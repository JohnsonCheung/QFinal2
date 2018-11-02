Attribute VB_Name = "MIde_Ens_PrpEr"
Option Explicit

Private Sub MdPrpEnsExitPrpLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsExitPrpLin"
Dim L&
L = MdPrpInsExitPrpLno(A, PrpLno)
If L = 0 Then Exit Sub
A.InsertLines L, "Exit Property"
Msg CSub, "Exit Property is inserted [at]", L
End Sub

Private Sub MdPrpEnsLblXLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsLblXLin"
Dim E$, L%, ActLblXLin$, EndPrpLno&
E = MdPrpLblXLin(A, PrpLno)
L = MdPrpLblXLno(A, PrpLno)
If L <> 0 Then
    ActLblXLin = A.Lines(L, 1)
End If
If E <> ActLblXLin Then
    If L = 0 Then
        EndPrpLno = MdPrpEndPrpLno(A, PrpLno)
        If EndPrpLno = 0 Then Stop
        A.InsertLines EndPrpLno, E
        If Trc Then Msg CSub, "Inserted [at] with [line]", EndPrpLno, E
    Else
        A.ReplaceLine L, E
        If Trc Then Msg CSub, "Replaced [at] with [line]", L, E
    End If
End If
End Sub
Private Sub MdPrpEnsOnEr(A As CodeModule, PrpLno&)
If HasSubStr(A.Lines(PrpLno, 1), "End Property") Then
    Exit Sub
End If
MdPrpEnsLblXLin A, PrpLno
MdPrpEnsExitPrpLin A, PrpLno
MdPrpEnsOnErLin A, PrpLno
End Sub

Private Sub MdPrpEnsOnErLin(A As CodeModule, PrpLno&)
Const CSub$ = "MdPrpEnsOnErLin"
Dim L&
L = MdPrpOnErLno(A, PrpLno)
If L <> 0 Then Exit Sub
A.InsertLines PrpLno + 1, "On Error Goto X"
If Trc Then Msg CSub, "Exit Property is inserted [at]", L
End Sub

Function MdPrpExitPrpLno&(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then Exit Function
Dim J%, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "Exit Property") Then MdPrpExitPrpLno = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Function MdPrpInsExitPrpLno&(A As CodeModule, PrpLno)
If MdPrpExitPrpLno(A, PrpLno) <> 0 Then Exit Function
Dim L%
L = MdPrpLblXLno(A, PrpLno)
If L = 0 Then Stop
MdPrpInsExitPrpLno = L
End Function

Function MdPrpLblXLin$(A As CodeModule, PrpLno)
Dim Nm$, Lin$
Lin = A.Lines(PrpLno, 1)
Nm = LinPrpNm(Lin)
If Nm = "" Then Stop
MdPrpLblXLin = FmtQQ("X: Debug.Print ""?.?.PrpEr...[""; Err.Description; ""]""", MdNm(A), Nm)
End Function

Function MdPrpLblXLno&(A As CodeModule, PrpLno)
Dim J&, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "X: Debug.Print") Then MdPrpLblXLno = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop
End Function

Function MdPrpLnoAy(A As CodeModule) As Long()
Dim O&(), Lno&
For Lno = 1 To A.CountOfLines
    If LinIsPrp(A.Lines(Lno, 1)) Then
        Push O, Lno
    End If
Next
MdPrpLnoAy = O
End Function

Function MdPrpLy(A As CodeModule) As String()
Dim O$(), Lno
For Lno = 0 To AyNz(MdPrpLnoAy(A))
    Push O, A.Lines(Lno, 1)
Next
MdPrpLy = O
End Function

Function MdPrpNy(A As CodeModule) As String()
Dim O$(), Lno
For Each Lno In AyNz(MdPrpLnoAy(A))
    PushNoDup O, LinPrpNm(A.Lines(Lno, 1))
Next
MdPrpNy = O
End Function

Function MdPrpOnErLno&(A As CodeModule, PrpLno)
Dim J%, L$
For J = PrpLno + 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If HasPfx(L, "On Error Goto X") Then MdPrpOnErLno = J: Exit Function
    If HasPfx(L, "End Property") Then Exit Function
Next
Stop '
End Function

Function MdPrpEndPrpLno&(A As CodeModule, PrpLno)
If HasSfx(A.Lines(PrpLno, 1), "End Property") Then MdPrpEndPrpLno = PrpLno: Exit Function
Dim J%
For J = PrpLno + 1 To A.CountOfLines
    If HasPfx(A.Lines(J, 1), "End Property") Then MdPrpEndPrpLno = J: Exit Function
Next
Stop
End Function
Sub MdEnsPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&()
L = MdPrpLnoAy(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpEnsOnEr A, L(J)
Next
End Sub



Private Sub MdPrpRmvOnEr(A As CodeModule, PrpLno&)
MdLno_Rmv A, MdPrpExitPrpLno(A, PrpLno)
MdLno_Rmv A, MdPrpOnErLno(A, PrpLno)
MdLno_Rmv A, MdPrpLblXLno(A, PrpLno)
End Sub



Function MdPrpPrpNm$(A As CodeModule, PrpLno)
MdPrpPrpNm = LinPrpNm(A.Lines(PrpLno, 1))
End Function

Sub MdRmvPrpOnEr(A As CodeModule)
If A.Parent.Type <> vbext_ct_ClassModule Then Exit Sub
Dim J%, L&()
L = MdPrpLnoAy(A)
If Not AyIsSrt(L) Then Stop
For J = UB(L) To 0 Step -1
    MdPrpRmvOnEr A, L(J)
Next
End Sub

Sub RmvPrpOnEr()
MdRmvPrpOnEr CurMd
End Sub

Sub RmvSchmPrpOnEr()
MdRmvPrpOnEr Md("Schm")
MdRmvPrpOnEr Md("SchmT")
MdRmvPrpOnEr Md("SchmF")
End Sub
