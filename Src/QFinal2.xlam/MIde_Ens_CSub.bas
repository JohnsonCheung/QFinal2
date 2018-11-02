Attribute VB_Name = "MIde_Ens_CSub"
Option Explicit
Private Type CSubBrk
   NeedDlt As Boolean
   NeedIns As Boolean
   OldLno As Long
   OldCSub As String
   NewLno As Long
   NewCSub As String
   MdNm As String
   MthNm As String
End Type


Private Function MdCSubDrs(A As CodeModule) As Drs
Set MdCSubDrs = Drs(CSubBrkDryFny, MdCSubDry(A))
End Function

Private Function MdCSubDry(A As CodeModule) As Variant()
Dim Mth, Dry()
Dim Ny$(): Ny = MdMthNy(A)
If Sz(Ny) = 0 Then Exit Function
For Each Mth In Ny
   Push Dry, CSubBrk_Dr(MdMthCSubBrk(A, CStr(Mth)))
Next
MdCSubDry = Dry
End Function

Private Sub CSubBrkDmp(A As CSubBrk)
D CSubBrkLin(A)
End Sub

Private Sub Z_MdMthCSubBrk()
CSubBrkDmp MdMthCSubBrk(CurMd, "MdMthCSubBrk")
Stop
End Sub
Function CSubBrkLin$(A As CSubBrk)
End Function


Private Function CSubBrkDryFny() As String()
CSubBrkDryFny = SslSy("MdNm MthNm NeedDlt NeedIns NewLno NewCSub OldLno OldCSub")
End Function

Private Function CSubBrk_Dr(A As CSubBrk) As Variant()
With A
   CSubBrk_Dr = Array(.MdNm, .MthNm, .NeedDlt, .NeedIns, .NewLno, .NewCSub, .OldLno, .OldCSub)
End With
End Function

Private Function CSubBrkDic(A As CSubBrk) As Dictionary
With A
'   CSubBrk_Drec.Dr = CSubBrk_Dr(A)
'   CSubBrk_Drec.Fny = CSubBrkDryFny
End With
End Function

Function CSubBrk_Str$(A As CSubBrk)
With A
   CSubBrk_Str = JnTab(Array(IIf(.OldCSub = .NewCSub, "*NoChg", "*Upd"), A.MdNm, AlignL(A.MthNm, 25), "CSub=[" & .NewCSub & "]"))
End With
End Function

Sub MdEnsCSub(A As CodeModule)
Dim M
For Each M In AyNz(MdMthNy(A))
   MdMthEnsCSub A, CvMd(M)
Next
End Sub



Sub MdBrwCSubDrs(A As CodeModule)
DrsBrw MdCSubDrs(A)
End Sub

Private Function MdMthCSubBrk(A As CodeModule, MthNm) As CSubBrk
Const CSub$ = "MdMthCSubBrk"
Dim MLy$()
Dim MLno&
   MLno = MdMthLno(A, MthNm)
   MLy = MdMthBdyLy(A, MthNm)

Dim IsUsingCSub As Boolean '-> NewAt
   IsUsingCSub = False
   If HasSubStrAy(Join(MLy), ApSy("Er CSub,", "Debug.Print CSub", "(CSub,")) Then
       IsUsingCSub = True
   End If

Dim OldCSubIx%
   Dim J%
   OldCSubIx = -1
   For J = 0 To UB(MLy)
       If HasPfx(MLy(J), "Const CSub") Then
           OldCSubIx = J
       End If
   Next

Dim OOldLno&
   OOldLno = IIf( _
       OldCSubIx >= 0, _
       MLno + OldCSubIx, _
       0)

Dim OOldCSub$
   If OldCSubIx >= 0 Then
       OOldCSub = MLy(OldCSubIx)
   Else
       OOldCSub = ""
   End If

Dim ONewLno&
   If IsUsingCSub Then
       Dim Fnd As Boolean
       For J = 0 To UB(MLy)
           If LasChr(MLy(J)) <> "_" Then
               Fnd = True
               ONewLno = MLno + J + 1
               Exit For
           End If
       Next
       If Not Fnd Then Er CSub, "{MthLy} has all lines with _ as sfx with is impossible", MLy
   Else
       ONewLno = 0
   End If

Dim ONewCSub$

   If ONewLno > 0 Then
       ONewCSub = FmtQQ("Const CSub$ = ""?""", A)
   Else
       ONewCSub = ""
   End If

Dim O As CSubBrk
   Dim HasOldCSub As Boolean
   Dim HasNewCSub As Boolean
   Dim IsDiff As Boolean
       HasOldCSub = OOldCSub <> ""
       HasNewCSub = ONewCSub <> ""
       IsDiff = OOldCSub <> ONewCSub
   With O
       .NeedDlt = IsDiff And HasOldCSub
       .NeedIns = IsDiff And HasNewCSub
       .NewCSub = ONewCSub
       .NewLno = ONewLno
       .OldCSub = OOldCSub
       .OldLno = OOldLno
       .MdNm = MdNm(A)
       .MthNm = MthNm
   End With
MdMthCSubBrk = O
End Function


Sub MdMthDmpCSubBrk(A As CodeModule, MthNm$)
Const CSub$ = "DmpCSubBrk"
Debug.Print CSubBrk_Str(MdMthCSubBrk(A, MthNm))
End Sub

Sub MdMthEnsCSub(A As CodeModule, MthNm$)
Const CSub$ = "MdMthEnsCSub"
Dim B As CSubBrk
    B = MdMthCSubBrk(A, MthNm)
With B
   If .NeedDlt Then
       A.DeleteLines .OldLno         '<==
   End If
   If .NeedIns Then
       A.InsertLines .NewLno, .NewCSub
   End If
End With
Debug.Print CSubBrk_Str(B)
End Sub

Function PjCSubDt(A As VBProject) As Dt
Dim I, Md As CodeModule
Dim Dry()
For Each I In PjMdAy(A)
   Set Md = I
   PushAy Dry, MdCSubDry(Md)
Next
PjCSubDt = Dt("Pj-CSub", CSubBrkDryFny, Dry)
End Function
Sub PjEnsCSub(A As VBProject)
Dim I, Md As CodeModule
For Each I In PjMdAy(A)
   Set Md = I
   MdEnsCSub Md
Next
End Sub

