Attribute VB_Name = "MVb_Fs_Ffn"
Option Explicit
Function CutExt$(A)
CutExt = FfnCutExt(A)
End Function

Function FfnAddFnPfx$(A$, Pfx$)
FfnAddFnPfx = FfnPth(A) & Pfx & FfnFn(A)
End Function

Function FfnNotExistChk(FfnAy0) As String()
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    If Not FfnIsExist(CStr(I)) Then
        PushI O, I
    End If
Next
If Sz(O) = 0 Then Exit Function
FfnNotExistChk = MsgAp_Ly("[File(s)] not found", O)
End Function
Function FfnNotExistMsg(A, Optional FilKind$ = "File") As String()
Dim F$
F = FmtQQ("[?] not found in [folder]", FilKind)
FfnNotExistMsg = MsgLy(F, FfnFn(A), FfnPth(A))
End Function

Function CvFfnAy(FfnAy0) As String()
Select Case True
Case IsStr(FfnAy0):   CvFfnAy = ApSy(FfnAy0)
Case IsArray(FfnAy0): CvFfnAy = AySy(FfnAy0)
Case Else: Stop
End Select
End Function

Function FfnAddFnSfx$(A$, Sfx$)
FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
End Function

Function FfnAlreadyLdMsgLy(A$, FilKind$, LdTim$) As String()
Dim Sz&, Tim$, Ld$, Msg$
Sz = FfnSz(A)
Tim = FfnDTim(A)
Msg = FmtQQ("[?] file of [time] and [size] is already loaded [at].", FilKind)
FfnAlreadyLdMsgLy = MsgLy(Msg, A, Tim, Sz, LdTim)
End Function

Sub FfnAsgXTSz(A$, OTim As Date, OSz&)
If Not FfnIsExist(A) Then
    OTim = 0
    OSz = 0
    Exit Sub
End If
OTim = FfnTim(A)
OSz = FfnSz(A)
End Sub

Sub FfnAssExist(A$)
If AyBrwEr(FfnNotFndChk(A)) Then Stop
End Sub

Sub FfnAyDltIfExist(A)
AyDo A, "FfnDltIfExist"
End Sub

Sub FfnCpy(A, ToFfn$, Optional OvrWrt As Boolean)
If OvrWrt Then FfnDlt ToFfn
FileSystem.FileCopy A, ToFfn
End Sub

Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

Function FfnCpyToPthIfDif(A, Pth$) As Boolean
Const M_Sam$ = "File is same the one in Path."
Const M_Copied$ = "File is copied to Path."
Const M_NotFnd$ = "File not found, cannot copy to Path."
Dim B$, Msg$
Select Case True
Case FfnIsExist(A)
    B = Pth & FfnFn(A)
    Select Case True
    Case FfnIsSam(B, A)
        Msg = M_Sam: GoSub Prt
    Case Else
        Fso.CopyFile A, B, True
        Msg = M_Copied: GoSub Prt
    End Select
Case Else
    Msg = M_NotFnd: GoSub Prt
    FfnCpyToPthIfDif = A
End Select
Exit Function
Prt:
    Debug.Print FmtQQ("FfnCpyToPthIfDif: ? Path=[?] File=[?]", Msg, Pth, A)
    Return
End Function

Function FfnAyCpyToPthIfDif(FfnAy0, Pth$) As String()
PthEns Pth
Dim I, O$()
For Each I In CvFfnAy(FfnAy0)
    PushNonBlankStr O, FfnCpyToPthIfDif(I, Pth)
Next
If Sz(O) > 0 Then
    PushMsgUnderLinDbl O, "Above files not found"
    FfnAyCpyToPthIfDif = O
End If
End Function

Function FfnRmvExt$(A)
FfnRmvExt = FfnCutExt(A)
End Function

Function FfnCutExt$(A)
Dim B$, C$, P%
B = FfnFn(A)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
FfnCutExt = FfnPth(A) & C
End Function

Sub FfnDlt(A)
On Error GoTo X
Kill A
Exit Sub
X: Debug.Print FmtQQ("FfnDlt: Kill(?) Er(?)", A, Err.Description)
'    RaiseErr
End Sub

Sub FfnDltIfExist(A)
If FfnIsExist(A) Then FfnDlt A
End Sub

Function FfnDTim$(A)
If FfnIsExist(A) Then
    FfnDTim = DteDTim(FileDateTime(A))
End If
End Function

Function FfnExt$(A)
Dim B$, P%
B = FfnFn(A)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
FfnExt = Mid(B, P)
End Function

Function FfnFdr$(A)
FfnFdr = PthFdr(FfnPth(A))
End Function

Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnCutExt(FfnFn(A))
End Function

Function FfnIsExist(A) As Boolean
FfnIsExist = Fso.FileExists(A)
End Function

Function FfnNotExist(A) As Boolean
FfnNotExist = Not Fso.FileExists(A)
End Function

Function FfnIsSam(A, B) As Boolean
If FfnTim(A) <> FfnTim(B) Then Exit Function
If FfnSz(A) <> FfnSz(B) Then Exit Function
FfnIsSam = True
End Function

Function FfnIsSamMsg(A, B, Sz&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Sz
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
FfnIsSamMsg = O
End Function

Sub FfnMov(A$, ToFfn$)
Fso.MoveFile A, ToFfn
End Sub

Function FfnNotFndChk(A$) As String()
If FfnIsExist(A) Then Exit Function
FfnNotFndChk = MsgLy("[File] not exist", A)
End Function

Function FfnNxt$(A$)
If FfnIsExist(A) Then FfnNxt = A: Exit Function
Dim J%, O$
For J = 1 To 999
    O = FfnAddFnSfx(A, "(" & Format(J, "000") & ")")
    If Not FfnIsExist(O) Then FfnNxt = O: Exit Function
Next
Stop
End Function

Function FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
FfnPth = Left(A, P)
End Function

Function FfnRplExt$(A, NewExt)
FfnRplExt = FfnRmvExt(A) & NewExt
End Function

Function FfnSel$(A, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With Application.FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = A
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function

Function FfnStampDr(A) As Variant()
FfnStampDr = Array(A, FfnSz(A), FfnTim(A), Now)
End Function

Function FfnSz&(A)
If Not FfnIsExist(A) Then FfnSz = -1: Exit Function
FfnSz = FileLen(A)
End Function

Function FfnTim(A) As Date
If FfnIsExist(A) Then FfnTim = FileDateTime(A)
End Function

Function FfnTimSzStr$(A)
If FfnIsExist(A) Then FfnTimSzStr = FfnDTim(A) & "." & FfnSz(A)
End Function
