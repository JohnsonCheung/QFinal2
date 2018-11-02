Attribute VB_Name = "MXls__Dta"
Option Explicit
Function DryWs(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm)
DryRg Dry, WsA1(O)
Set DryWs = O
End Function
Function DrsWs(A As Drs, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm)
DrsLo A, WsA1(O)
Set DrsWs = O
End Function

Function DrsAt(A As Drs, At As Range) As Range
Set DrsAt = SqRg(DrsSq(A), At)
End Function

Function DrsWsBrw(A As Drs) As Worksheet
Set DrsWsBrw = WsVis(DrsWs(A))
End Function

Function DrsWsVis(A As Drs) As Worksheet
Set DrsWsVis = WsVis(DrsWs(A))
End Function

Function SqRg(A, At As Range) As Range
Dim O As Range
Set O = RgReSz(At, A)
O.MergeCells = False
O.Value = A
Set SqRg = O
End Function

Function DsWb(A As Ds) As Workbook
Dim O As Workbook
Set O = NewWb
With WbFstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Dim I
For Each I In AyNz(A.DtAy)
    WbAddDt O, CvDt(I)
Next
Set DsWb = O
End Function

Private Sub Z_DsWb()
Dim Wb As Workbook
Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub

Function DsWs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NewWs
WsA1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Dim I
For Each I In AyNz(A.DtAy)
    Set At = DtAtNxt(CvDt(I), At, J)
Next
Set DsWs = O
End Function


Function DrsRg(A As Drs, At As Range) As Range
Set DrsRg = SqRg(DrsSq(A), At)
End Function
Function DtAt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = DrsFmt(DtDrs(A))
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAt = At
End Function

Function DtAtNxt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = DrsFmt(DtDrs(A))
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAtNxt = At
End Function

Function DtLo(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set DtLo = DrsLo(DtDrs(A), R)
RgRC(R, 0, 1).Value = A.DtNm
End Function


Function DtPutWb(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(Wb, A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtPutWb = O
End Function


Sub Z_ItrPrpDrs()
DrsBrw ItrPrpDrs(DbtFds(SampleDb_DutyPrepare, "Permit"), "Name Type Required")
'DrsBrw ItrPrpDrs(Application.VBE.VBProjects, "Name Type")
End Sub

Function DrsLo(A As Drs, At As Range) As ListObject
Set DrsLo = RgLo(SqRg(DrsSq(A), At))
End Function

Function DrsLoFmt(A As Drs, At As Range, LoFmtr$()) As ListObject
Set DrsLoFmt = LoFmt(DrsLo(A, At), LoFmtr)
End Function


Function DrsNewWs(A As Drs) As Worksheet
Set DrsNewWs = SqNewWs(DrsSq(A))
End Function

Function DryRg(A, At As Range) As Range
Set DryRg = SqRg(DrySq(A), At)
End Function


Function DryNewWs(A) As Worksheet
Set DryNewWs = SqNewWs(DrySq(A))
End Function

Private Sub ZZ_DsWb()
Dim Wb As Workbook
Stop '
'Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub



Sub ZZ_DsWs()
WsVis DsWs(SampleDs)
End Sub
