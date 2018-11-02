Attribute VB_Name = "MDao_Schm_Asg"
Option Explicit

Sub SchmAsg(Schm$, OEr$(), OStruAy$(), OStruBase As StruBase)
Dim Ly$()
Ly = Split(Schm, vbCrLf)
OStruAy = AyWhT1SelRst(Ly, "Tbl")
OStruBase = StruBase1(Ly)
OEr = SchmLyEr(Ly)
End Sub

Private Function StruBase1(Ly$()) As StruBase
With StruBase1
    Set .E = SchmEleDic(Ly)
    Set .F = SchmFldDrs(Ly)
    Set .TDes = SchmTDesDic(Ly)
    Set .FDes = SchmFDesDic(Ly)
    Set .TFDes = SchmTFDesDrs(Ly)
End With
End Function

Private Function SchmLyEr(SchmLy$()) As String()
Dim E1$(), E2$(), E3$(), E4$(), E5$()
PushIAy SchmLyEr, E1
PushIAy SchmLyEr, E2
PushIAy SchmLyEr, E3
PushIAy SchmLyEr, E4
PushIAy SchmLyEr, E5
End Function

Private Function SchmEleDic(Ly$()) As Dictionary
Dim E, Ele$, EleStr$
Set SchmEleDic = New Dictionary
For Each E In AyNz(AyWhT1SelRst(Ly, "Ele"))
    Ele = ShfTerm(E)
    SchmEleDic.Add Ele, EleStrFd(E)
Next
End Function

Private Function SchmFDesDic(Ly$()) As Dictionary
Set SchmFDesDic = LyDic(AyWhT1SelRst(Ly, "FDes"))
End Function

Private Function SchmFldDrs(Ly$()) As Drs
Dim L, Ele$, Dry(), FldLik, F
For Each F In AyNz(AyWhT1SelRst(Ly, "Fld"))
    Ele = ShfTerm(F)
    For Each FldLik In SslSy(F)
        PushI Dry, Array(Ele, FldLik)
    Next
Next
Set SchmFldDrs = Drs("Ele FldLik", Dry)
End Function

Private Function SchmTDesDic(Ly$()) As Dictionary
Set SchmTDesDic = LyDic(AyWhT1SelRst(Ly, "TDes"))
End Function

Private Function SchmTFDesDrs(Ly$()) As Drs
Dim Dry(), L
For Each L In Ly
    PushI Dry, Lin3TRst(L)
Next
Set SchmTFDesDrs = Drs("T F Des", Dry)
End Function

