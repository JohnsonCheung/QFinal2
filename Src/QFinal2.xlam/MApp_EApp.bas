Attribute VB_Name = "MApp_EApp"
Option Explicit
Function EAppFbDic() As Dictionary
Const A$ = "N:\SAPAccessReports\"
Set EAppFbDic = New Dictionary
With EAppFbDic
        .Add "Duty", A & "DutyPrepay\.accdb"
       .Add "SkHld", A & "StkHld\.accdb"
     .Add "ShpRate", A & "DutyPrepay\StockShipRate_Data.accdb"
      .Add "ShpCst", A & "StockShipCost\.accdb"
      .Add "TaxCmp", A & "TaxExpCmp\.accdb"
    .Add "TaxAlert", A & "TaxRateAlert\.accdb"
End With
End Function
