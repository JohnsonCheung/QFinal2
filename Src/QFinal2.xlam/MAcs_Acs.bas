Attribute VB_Name = "MAcs_Acs"
Option Explicit
Sub AcsSavRec(A As Access.Application)
A.DoCmd.RunCommand acCmdSaveRecord
End Sub

Sub AcsOpn(A As Access.Application, Fb$)
Select Case True
Case IsNothing(A.CurrentDb)
    A.OpenCurrentDatabase Fb
Case A.CurrentDb.Name = Fb
Case Else
    A.CurrentDb.Close
    A.OpenCurrentDatabase Fb
End Select
End Sub

Function AcsOpnFb(A$) As Access.Application
Acs.CloseCurrentDatabase
Acs.OpenCurrentDatabase A
Set AcsOpnFb = Acs
End Function

Sub AcsQuit(A As Access.Application)
AcsClsDb A
A.Quit
Set A = Nothing
End Sub

Sub AcsVis(A As Access.Application)
If Not A.Visible Then A.Visible = True
End Sub

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function
Function FbAcs(A, Optional Vis As Boolean) As Access.Application
Dim O As New Access.Application
O.OpenCurrentDatabase A
O.Visible = Vis
Set FbAcs = O
End Function

Function NewAcs(Optional Hid As Boolean) As Access.Application
Dim O As New Access.Application
If Not Hid Then O.Visible = True
Set NewAcs = O
End Function


Sub AcsCls(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub AcsClsDb(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Function Acs() As Access.Application
Static X As Boolean, Y As Access.Application
On Error GoTo X
If X Then
    Set Y = New Access.Application
    X = True
End If
If Y.Application.Name = "Microsoft Access" Then
    Set Acs = Y
    Exit Function
End If
X:
    Set Y = New Access.Application
    Debug.Print "Acs: New Acs instance is crreated."
Set Acs = Y
End Function
