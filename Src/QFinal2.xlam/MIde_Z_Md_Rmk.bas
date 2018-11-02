Attribute VB_Name = "MIde_Z_Md_Rmk"
Option Explicit
Sub RmkMd()
MdRmk CurMd
End Sub
Sub UnRmkMd()
MdUnRmk CurMd
End Sub
Sub RmkAllMd()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In CurPjMdAy
    If Md.Name <> "LibIdeRmkMd" Then
        If MdRmk(CvMd(I)) Then
            NRmk = NRmk + 1
        Else
            Skip = Skip + 1
        End If
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Sub UnRmkAllMd()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In CurPjMdAy
    Set Md = I
    If MdUnRmk(Md) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Private Function MdRmk(A As CodeModule) As Boolean
Debug.Print "Rmk " & A.Parent.Name,
If MdIsAllRemarked(A) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To A.CountOfLines
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
MdRmk = True
End Function

Private Function MdUnRmk(A As CodeModule) As Boolean
Debug.Print "UnRmk " & A.Parent.Name,
If Not MdIsAllRemarked(A) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
MdUnRmk = True
End Function
Private Function MdIsAllRemarked(A As CodeModule) As Boolean
Dim J%, L$
For J = 1 To A.CountOfLines
    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Function
Next
MdIsAllRemarked = True
End Function
