Attribute VB_Name = "MVb_Str_Mch"
Option Explicit
Type PatnRslt
    Patn As String
    Rslt As Variant
End Type
Sub Z()
Z_StrMchPatn
End Sub
Function StrDicMch(A, PatnRsltDic As Dictionary) As PatnRslt
Dim Patn
For Each Patn In PatnRsltDic
    If StrMchPatn(A, Patn) Then
        With StrDicMch
            .Rslt = PatnRsltDic(Patn)
            .Patn = Patn
        End With
        Exit Function
    End If
Next
End Function
Function StrDicMap(A, PatnRsltDic As Dictionary)
With StrDicMch(A, PatnRsltDic)
    If .Patn = "" Then Exit Function
    Asg .Rslt, StrDicMap
End With
End Function

Private Sub Z_StrMchPatn()
Dim A$, Patn$
Ept = True: A = "AA": Patn = "AA": GoSub Tst
Ept = True: A = "AA": Patn = "^AA$": GoSub Tst
Exit Sub
Tst:
    Act = StrMchPatn(A, Patn)
    C
    Return
End Sub

Function StrMchPatn(A, Patn) As Boolean
Static Re As New RegExp
Re.Pattern = Patn
StrMchPatn = Re.Test(A)
End Function
