Attribute VB_Name = "MVb_Fs_Pth"
Option Explicit
Private O$() ' Used in PthPushEntAyR
Private Sub ZZ_PthEntAy()
Dim A$(): A = PthEntAy("C:\users\user\documents\", IsRecursive:=True)
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Private Sub ZZ_PthFxAy()
Dim A$()
A = PthFxAy(CurDir)
AyDmp A
End Sub

Private Sub ZZ_PthRmvEmpSubDir()
PthRmvEmpSubDir TmpPth
End Sub
Function Cd$(Optional A$)
If A = "" Then
    Cd = PthEnsSfx(CurDir)
    Exit Function
End If
ChDir A
Cd = PthEnsSfx(A)
End Function


Sub ZZ_PthSel()
MsgBox FfnSel("C:\")
End Sub

Sub PthBrw(A$)
Shell FmtQQ("Explorer ""?""", A), vbMaximizedFocus
End Sub

Sub PthClr(A$)
FfnAyDltIfExist PthFfnAy(A)
End Sub

Sub PthClrFil(A$)
If Not PthIsExist(A) Then Exit Sub
Dim F
For Each F In AyNz(PthFfnAy(A))
   FfnDlt F
Next
End Sub

Function PthEns$(A$)
If Not Fso.FolderExists(A) Then MkDir A
PthEns = A
End Function

Function PthEnsSfx$(A$)
PthEnsSfx = EnsSfx(A, "\")
End Function

Function PthEntAy(A$, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
If Not IsRecursive Then
    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(A, FilSpec, Atr))
    Exit Function
End If
Erase O
PthPushEntAyR A
PthEntAy = O
Erase O
End Function

Private Sub Z_PthEntAy()
Dim A$(): A = PthEntAy("C:\users\user\documents\", IsRecursive:=True)
Debug.Print Sz(A)
Stop
AyDmp A
End Sub

Function PthFdr$(A$)
PthFdr = TakAftRev(RmvLasChr(A), "\")
End Function

Function PthFfnAy(A$, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Function

Function PthFnAy(A$, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
Ass PthIsExist(A)
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
       Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
Ass PthHasPthSfx(A)
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Function

Function PthFxAy(A$) As String()
Dim O$(), B$
If Right(A, 1) <> "\" Then Stop
B = Dir(A & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If FfnExt(B) = ".xls" Then
        Push O, A & B
    End If
    B = Dir
Wend
PthFxAy = O
End Function

Function PthHasFil(A$) As Boolean
Ass PthHasPthSfx(A)
If Not PthIsExist(A) Then Exit Function
PthHasFil = (Dir(A & "*.*") <> "")
End Function

Function PthHasPthSfx(A$) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Function

Function PthHasSubDir(A$) As Boolean
If Not PthIsExist(A) Then Exit Function
Ass PthHasPthSfx(A)
Dim P$: P = Dir(A & "*.*", vbDirectory)
Dir
PthHasSubDir = Dir <> ""
End Function

Function PthIsEmp(A$) As Boolean
If PthHasFil(A) Then Exit Function
If PthHasSubDir(A) Then Exit Function
PthIsEmp = True
End Function

Function PthIsExist(A$) As Boolean
PthIsExist = Fso.FolderExists(A)
End Function

Sub PthMovFilUp(A$)
Dim I, Tar$
Tar$ = PthUp(A)
For Each I In AyNz(PthFnAy(A))
    FfnMov CStr(I), Tar
Next
End Sub

Private Sub PthPushEntAyR(A$)
'Debug.Print "PthPUshEntAyR:" & A
Dim P$(): P = PthSubPthAy(A)
If Sz(P) = 0 Then Exit Sub
If Sz(O) Mod 1000 = 0 Then Debug.Print "PthPushEntAyR: (Each 1000): " & A
PushAy O, P
PushAy O, PthFfnAy(A)
Dim PP
For Each PP In P
    PthPushEntAyR CStr(PP)
Next
End Sub

Sub PthRmvEmpSubDir(A$)
Dim I
For Each I In AyNz(PthSubPthAy(A))
   PthRmvIfEmp CStr(I)
Next
End Sub

Sub PthRmvIfEmp(A$)
If Not PthIsExist(A) Then Exit Sub
If PthIsEmp(A) Then Exit Sub
RmDir A
End Sub

Function PthSel$(A$, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = IIf(IsNull(A), "", A)
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = PthEnsSfx(.SelectedItems(1))
    End If
End With
End Function

Function PthSubFdrAy(A$, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
'PthSubFdrAy = ItrNy(Fso.GetFolder(A).SubFolders, Spec)
Ass PthIsExist(A)
Ass PthHasPthSfx(A)
Dim O$(), M$, X&, XX&
X = Atr Or vbDirectory
M = Dir(A & Spec, vbDirectory)
While M <> ""
    If InStr(M, "?") > 0 Then
        Debug.Print "PthSubFdrAy: Skip -> [" & M & "]"
        GoTo Nxt
    End If
    XX = GetAttr(A & M)
    If M = "." Then GoTo Nxt
    If M = ".." Then GoTo Nxt
    If XX And X Then
        Push O, M
    End If
Nxt:
    M = Dir
Wend
PthSubFdrAy = O
End Function

Function PthSubPthAy(A$, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthSubPthAy = AyAddPfxSfx(PthSubFdrAy(A, Spec, Atr), A, "\")
End Function

Function PthUp$(A, Optional Up% = 1)
Dim O$, J%
O = A
For J = 1 To Up
    O = PthUpOne(O)
Next
PthUp = O
End Function

Function PthUpOne$(A$)
PthUpOne = TakBefOrAllRev(RmvSfx(A, "\"), "\") & "\"
End Function
