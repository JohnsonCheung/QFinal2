Attribute VB_Name = "MDta_Sel"
Option Explicit
Function DrSel(A, IxAy) As Variant()
Dim Ix
For Each Ix In IxAy
    PushI DrSel, A(Ix)
Next
End Function
Function DrySel(A, IxAy) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    PushI DrySel, DrSel(Dr, IxAy)
Next
End Function
Function DrySelIxAp(A, ParamArray IxAp()) As Variant()
Dim IxAy(): IxAy = IxAp
DrySelIxAp = DrySel(A, IxAy)
End Function



Function DrsSel(A As Drs, Fny0) As Drs
Dim Fny$(): Fny = CvNy(Fny0)
If IsEqAy(A.Fny, Fny) Then Set DrsSel = A: Exit Function
AyBrwEr AyHasAyChk(A.Fny, Fny)

Dim Ix&(): Ix = AyIxAy(A.Fny, Fny)
Dim Dry(), Dr
For Each Dr In A.Dry
    PushI Dry, DrSel(Dr, Ix)
Next
Set DrsSel = Drs(Fny, Dry)
End Function

Private Sub Z_DrsSel()
'DrsBrw DrsSel(Vmd.MthDrs, "MthNm Mdy Ty MdNm")
'DrsBrw Vmd.MthDrs
End Sub

Function DtSel(A As Dt, ColNy0) As Dt
Dim ReOrdFny$(): ReOrdFny = CvNy(ColNy0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DtSel = Dt(A.DtNm, OFny, ODry)
End Function
