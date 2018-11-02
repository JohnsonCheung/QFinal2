Attribute VB_Name = "MIde_Z_Vbe_Cur"
Option Explicit

Function CurVbe() As Vbe
Set CurVbe = Application.Vbe
End Function

Sub CurVbeExp()
VbeExp CurVbe
End Sub

Function CurVbeHasBar(Nm$) As Boolean
CurVbeHasBar = VbeHasBar(CurVbe, Nm)
End Function

Function CurVbeHasPjFfn(PjFfn) As Boolean
CurVbeHasPjFfn = VbeHasPjFfn(CurVbe, PjFfn)
End Function

Function CurVbeMthDot(Optional MthRe As RegExp, Optional MthExlAy$, Optional WhMdyAy, Optional WhMthKd0$, Optional PjRe As RegExp, Optional PjExlAy$, Optional MdRe As RegExp, Optional MdExlAy$)
Stop '
'CurVbeMthDot = VbeMthDot(CurVbe, MthPatn, MthExlAy, WhMdyA, WhMthKd0, PjPatn, PjExlAy, MdPatn, MdExlAy)
End Function

Function CurVbeMthNy(Optional A As WhPjMth) As String()
CurVbeMthNy = VbeMthNy(CurVbe, A)
End Function

Function CurVbeMthWb() As Workbook
Set CurVbeMthWb = VbeMthWb(CurVbe)
End Function

Function CurVbeMthWs() As Worksheet
Set CurVbeMthWs = VbeMthWs(CurVbe)
End Function

Function CurVbePj(A$) As VBProject
Set CurVbePj = CurVbe.VBProjects(A)
End Function

Function CurVbePjFfnPj(PjFfn) As VBProject
Set CurVbePjFfnPj = VbePjFfnPj(CurVbe, PjFfn)
End Function

Sub CurVbePjMdFmtBrw()
Brw VbePjMdFmt(CurVbe)
End Sub

Sub CurVbeSav()
VbeSav CurVbe
End Sub

Function CurVbeSrc() As String()
CurVbeSrc = VbeSrc(CurVbe)
End Function
