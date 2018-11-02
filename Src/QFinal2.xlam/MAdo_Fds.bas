Attribute VB_Name = "MAdo_Fds"
Option Explicit
Function AFdsDr(A As ADODB.Fields) As Variant()
Dim F As ADODB.Field
For Each F In A
   PushI AFdsDr, F.Value
Next
End Function

Function AFdsFny(A As ADODB.Fields) As String()
AFdsFny = ItrNy(A)
End Function
