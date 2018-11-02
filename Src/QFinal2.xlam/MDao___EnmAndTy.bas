Attribute VB_Name = "MDao___EnmAndTy"
Option Explicit
Type LnkAFil:  Ffn As String:    Kd As String:                      End Type
Type LnkASw:   Nm As String:     Bool As Boolean:                   End Type
Type LnkFmStu: Stu As String:    Inp() As String:                   End Type
Type LnkFmSw:: SwNm As String:   TF As Boolean:    Inp() As String: End Type
Type LnkFmWh:  Inp As String:    WhBExpr As String:                 End Type
Type LnkFmFil: Inp As String:    WhBExpr As String:                 End Type
Type LnkIpFil: Fil As String:    Inp As String:                     End Type
Type LnkIpWs:  FxKd As String:   WsNm As String:   Inp As String:   End Type
Type LnkStEle: Ele As String:    Stu As String:    Fny() As String: End Type
Type LnkStExt: LikInp As String: F As String:      Ext As String:   End Type
Type LnkStFld: Stu As String:    Fny() As String:                   End Type
Type LnkSpec
    AFb() As LnkAFil
    AFx() As LnkAFil
    ASw() As LnkASw
    FmFb() As LnkFmFil
    FmFx() As LnkFmFil
    FmIp() As String
    FmStu() As LnkFmStu
    FmSw() As LnkFmSw
    FmWh() As LnkFmWh
    IpFb() As LnkIpFil
    IpFx() As LnkIpFil
    IpS1() As String
    IpWs() As LnkIpWs
    StEle() As LnkStEle
    StExt() As LnkStExt
    StFld() As LnkStFld
End Type
Type FR: Er() As String: OkFilKind() As String: End Type ' FilRslt
Type Wr: Er() As String: OkWny() As String:     End Type ' WnyRslt
Type TR: Er() As String: OkTny() As String:     End Type ' TnyRslt
Type Cr: Er() As String:                        End Type ' ColRslt
Type XlsLnkInf
    IsXlsLnk As Boolean
    Fx As String
    WsNm As String
End Type
'============
Type StruBase
    F As Drs    'Ele FldLik
    E As Dictionary 'Ele->Fd
    TDes As Dictionary
    FDes As Dictionary
    TFDes As Drs
End Type
Type FDes: F As String: Des As String: End Type
Type TDes: T As String: Des As String: End Type

Type T: Lno As Integer: T As String: Fny() As String: Sk() As String:     End Type
Type F: Lno As Integer: E As String: LikT As String:  LikFny() As String: End Type
Type DD: Lno As Integer: T As String: F As String:     Des As String:      End Type
Type E
    Lno As Integer
    E As String
    Ty As DAO.DataTypeEnum
    Req As Boolean
    ZLen As Boolean
    TxtSz As Byte
    Expr As String
    VRul As String
    Dft As String
    VTxt As String
End Type
