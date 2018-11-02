Attribute VB_Name = "LibApp_Fun"
Option Explicit
Sub Z()
Z_AppFbAy
End Sub
Function DtaDb() As Database
Set DtaDb = DAO.DBEngine.OpenDatabase(DtaFb)
End Function

Sub BrwDtaFb()
'Acs.OpenCurrentDatabase DtaFb
End Sub

Function DtaFn$()
DtaFn = Apn & "_Data.accdb"
End Function

Function IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not PthIsExist(ProdPth)
End If
IsDev = Y
End Function

Sub CrtDtaFb()
If IsDev Then Exit Sub
If FfnIsExist(DtaFb) Then Exit Sub
FbCrt DtaFb
Dim Src, Tar$, TarFb$
TarFb = DtaFb
Stop
'For Each Src In CcmTny
    Tar = Mid(Src, 2)
    Application.DoCmd.CopyObject TarFb, Tar, acTable, Src
    Debug.Print MsgLin("CrtDtaFb: Cpy [Src] to [Tar]", Src, Tar)
'Next
End Sub

Sub Doc()
'#BNmMIS is Method-B-Nm-of-Missing.
'           Missing means the method is found in FmPj, but not ToPj
'#FmDicB is MthDic-of-MthBNm-zz-MthLines.   It comes from FmPj
'#ToDicA is MthDic-of-MthANm-zz-MthLinesAy. It comes from ToPj
'#ToDicAB is ToDicA and FmDicB
'#ANm is method-a-name, NNN or NNN:YYY
'        If the method is Sub|Fun, just MthNm
'        If the method is Prp    ,      MthNm:MthShtTy
'        It is from ToPj (#ToA)
'        One MthANm will have one or more MthLines
'#BNm is method-b-name, MMM.NNN or MMM.NNN:YYY
'        MdNm.MthNm[:MthShtTy]
'        It is from FmPj (#BFm)
'        One MthBNm will have only one MthLines
'#Missing is for each MthBNm found in FmDicB, but its MthNm is not found in any-method-name-in-ToDicA
'#Dif is for each MthBNm found in FmDicB and also its MthANm is found in ToDicA
'        and the MthB's MthLines is dif and any of the MthA's MthLines
'       (Note.MthANm will have one or more MthLines (due to in differmodule))
End Sub

Sub DocUOM()
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    AC_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        AC_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub

Function DtaFb$()
DtaFb = AppHom & DtaFn
End Function



Private Sub Z_AppFbAy()
Dim F
For Each F In AppFbAy
If Not IsFfnExist(F) Then Stop
Next
End Sub
Function AppDtaHom$()
AppDtaHom = PthUp(TmpHom)
End Function

Function AppDtaPth$()
AppDtaPth = PthEns(AppDtaHom & Apn & "\")
End Function

Sub AppExp()
PthClr SrcPth
SpecExp
AppExpMd
AppExpFrm
AppExpStru
End Sub

Sub AppExpFrm()
Dim Nm$, P$, I
P = SrcPth
For Each I In AppFrmNy
    Nm = I
    SaveAsText acForm, Nm, P & Nm & ".Frm.Txt"
Next
End Sub

Sub AppExpMd()
Dim MdNm$, I, P$
P = SrcPth
For Each I In AppMdNy
    MdNm = I
    SaveAsText acModule, MdNm, P & MdNm & ".bas"
Next
End Sub

Sub AppExpStru()
StrWrt Stru, SrcPth & "Stru.txt"
End Sub

Function AppFbAy() As String()
Push AppFbAy, AppJJFb
Push AppFbAy, AppStkShpCstFb
Push AppFbAy, AppStkShpRateFb
Push AppFbAy, AppTaxExpCmpFb
Push AppFbAy, AppTaxRateAlertFb
End Function

Function AppFrmNy() As String()
AppFrmNy = ItrNy(CodeProject.AllForms)
End Function

Function AppMdNy() As String()
AppMdNy = ItrNy(CodeProject.AllModules)
End Function

Function AppPushAppFcmd$()
AppPushAppFcmd = WPth & "PushApp.Cmd"
End Function

Function AppRoot$()
Stop '
End Function

Function IsProd() As Boolean
IsProd = Not IsDev
End Function



Function AutoExec()
'D "AutoExec:"
'D "-Before LnkCcm: CnSy--------------------------"
'D CnSy
'D "-Before LnkCcm: Srcy--------------------------"
'D Srcy
'
SpecEnsTbl

DbLnkCcm CurDb, IsDev
'D "-After LnkCcm: CnSy--------------------------"
'D CnSy
'D "-After LnkCcm: Srcy--------------------------"
'D Srcy
End Function
