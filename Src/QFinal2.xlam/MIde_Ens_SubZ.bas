Attribute VB_Name = "MIde_Ens_SubZ"
Option Explicit

Sub EnsPjZ()
ZPjEns CurPj
PjEnsZPrv CurPj
End Sub

Sub EnsZ()
ZMdEns CurMd
MdEnsZPrv CurMd
End Sub

Private Function ZLinesAct(A As CodeModule)
ZLinesAct = MdMthLines(A, "Z")
End Function

Private Function ZLinesEpt1$(ZMthNy$())
If Sz(ZMthNy) = 0 Then Exit Function
Dim O$()
Push O, "Sub Z()"
PushAy O, AySrt(ZMthNy)
Push O, "End Sub"
ZLinesEpt1 = JnCrLf(O)
End Function

Private Function ZLinesEpt(A As CodeModule)
ZLinesEpt = ZLinesEpt1(MdMthNy(A, WhMth(Nm:=WhNm("^Z_"))))
End Function

Private Sub ZMdEns(A As CodeModule)
Dim Ept$
Ept = ZLinesEpt(A)
If ZLinesAct(A) = Ept Then Exit Sub
MdMthRmv A, "Z"
MdLinesApp A, Ept
End Sub

Private Sub ZPjEns(A As VBProject)
Dim I
For Each I In PjMdAy(A)
    Debug.Print MdNm(CvMd(I))
    ZMdEns CvMd(I)
Next
End Sub
