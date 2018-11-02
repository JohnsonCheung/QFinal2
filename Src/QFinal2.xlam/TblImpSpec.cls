VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TblImpSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Tbl$, LnkColVbl$, WhBExpr$

Friend Property Get Init(Tbl, LnkColVbl, Optional WhBExpr)
Me.Tbl = Tbl
Me.LnkColVbl = LnkColVbl
Me.WhBExpr = WhBExpr
Set Init = Me
End Property
