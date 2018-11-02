Attribute VB_Name = "MDao_Lnk_Spec"
Option Explicit

Function LSClnInpy(A) As String()
LSClnInpy = SslSy(RmvT1(AyFstT1(A, "A-Inp")))
End Function

Function LSLines$()
LSLines = SpnmLines("Lnk")
End Function

Sub LSpecAsg(A, Optional OTblNm$, Optional OLnkColStr$, Optional OWhBExpr$)
Dim Ay$()
Ay = AyTrim(SplitVBar(A))
OTblNm = AyShf(Ay)
If LinT1(AyLasEle(Ay)) = "Where" Then
    OWhBExpr = RmvT1(Pop(Ay))
Else
    OWhBExpr = ""
End If
OLnkColStr = JnVBar(Ay)
End Sub

Sub LSpecDmp(A)
Debug.Print RplVBar(A)
End Sub

Function LSpecLy(A) As String()
Const L2Spec$ = ">GLAnp |" & _
    "Whs    Txt Plant |" & _
    "Loc    Txt [Storage Location]|" & _
    "Sku    Txt Material |" & _
    "PstDte Txt [Posting Date] |" & _
    "MovTy  Txt [Movement Type]|" & _
    "Qty    Txt Quantity|" & _
    "BchNo  Txt Batch |" & _
    "Where Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*'"
End Function

