Attribute VB_Name = "MDta__ObjPrp"
Option Explicit

Function ItrPrpDrs(A, PrpNy0) As Drs
Dim P$(): P = CvNy(PrpNy0)
Dim Dry()
    Dim I
    For Each I In A
        PushI Dry, ObjPrpDr(I, P)
    Next
Set ItrPrpDrs = Drs(P, Dry)
End Function

Function OyPrpDrs(A, PrpNy0) As Drs
Set OyPrpDrs = Drs(CvNy(PrpNy0), OyPrpDry(A, PrpNy0))
End Function

Function OyPrpDry(A, PrpNy0) As Variant()
Dim O(), U%, I
Dim PrpNy$()
PrpNy = CvNy(PrpNy0)
For Each I In A
    Push O, ObjDr(I, PrpNy)
Next
OyPrpDry = O
End Function

Sub Z_ItrPrpDrs()
DrsBrw ItrPrpDrs(Excel.Application.AddIns, "Name FullName CLSId Installed")
End Sub
