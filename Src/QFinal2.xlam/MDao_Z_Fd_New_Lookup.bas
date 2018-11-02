Attribute VB_Name = "MDao_Z_Fd_New_Lookup"
Option Explicit

Function LookupFd(Fld, Tbl, F As Drs, E As Dictionary) As DAO.Field2
Const CSub$ = "LookupFd"
Dim O As DAO.Field2, Ele$
Set O = StdFldFd(Fld, Tbl): If IsSomething(O) Then Set LookupFd = O: Exit Function
Ele = LookupFd1(Fld, F): If Ele = "" Then Er CSub, "[Tbl]-[Fld] is not found given [Fld-Drs]", Tbl, Fld, DrsFmt(F)
Set O = StdEleFd(Ele, Fld): If IsSomething(O) Then Set LookupFd = O:   Exit Function
If IsNothing(E) Then Er CSub, "[Tbl]-[Fld] cannot of get a Fd by [Fld-Drs] and [Ele-Dic]", Tbl, Fld, DrsFmt(F), EleDicFmt(E)
If Not E.Exists(Ele) Then Er CSub, ""
Set LookupFd = FdClone(E(Ele), Fld)
End Function

Private Function LookupFd1$(Fld, F As Drs) ' Return Ele$
Dim Dr
For Each Dr In AyNz(F.Dry)
    If Fld Like Dr(1) Then LookupFd1 = Dr(0): Exit Function
Next
End Function

