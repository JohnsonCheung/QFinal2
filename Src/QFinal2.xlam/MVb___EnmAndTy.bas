Attribute VB_Name = "MVb___EnmAndTy"
Option Explicit
Enum eBoolOp
    eOpEQ = 1
    eOpNE = 2
    eOpAND = 3
    eOpOR = 4
End Enum
Enum eEqNeOp
    eOpEQ = eBoolOp.eOpEQ
    eOpNE = eBoolOp.eOpNE
End Enum
Enum eAndOrOp
    eOpAND = eBoolOp.eOpAND
    eOpOR = eBoolOp.eOpOR
End Enum
