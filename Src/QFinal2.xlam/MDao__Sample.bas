Attribute VB_Name = "MDao__Sample"
Option Explicit

Function SampleDb_Duty_Dta() As Database
Set SampleDb_Duty_Dta = DBEngine.OpenDatabase(SampleFb_Duty_Dta)
End Function
