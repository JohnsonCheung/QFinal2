Attribute VB_Name = "MDta_Sample"
Option Explicit
Function SampleDr1() As Variant()
SampleDr1 = Array(1, 2, 3)
End Function

Function SampleDr2() As Variant()
SampleDr2 = Array(2, 3, 4)
End Function

Function SampleDr3() As Variant()
SampleDr3 = Array(3, 4, 5)
End Function

Function SampleDr4() As Variant()
SampleDr4 = Array(43, 44, 45)
End Function

Function SampleDr5() As Variant()
SampleDr5 = Array(53, 54, 55)
End Function

Function SampleDr6() As Variant()
SampleDr6 = Array(63, 64, 65)
End Function

Function SampleDrs1() As Drs
Set SampleDrs1 = Drs("A B C", SampleDry1)
End Function

Function SampleDrs2() As Drs
Set SampleDrs2 = Drs("A B C", SampleDry2)
End Function

Function SampleDrs() As Drs
Set SampleDrs = Drs("A B C D E G H I J K", SampleDry)
End Function

Function SampleDrsFny() As String()
SampleDrsFny = SslSy("A B C D E F G")
End Function

Function SampleDry1() As Variant()
SampleDry1 = Array(SampleDr1, SampleDr2, SampleDr3)
End Function

Function SampleDry2() As Variant()
SampleDry2 = Array(SampleDr3, SampleDr4, SampleDr5)
End Function

Function SampleDry() As Variant()
PushI SampleDry, Array("A", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "A"))
PushI SampleDry, Array("B", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "B"))
PushI SampleDry, Array("C", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "C"))
PushI SampleDry, Array("D", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "D"))
PushI SampleDry, Array("E", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "E"))
PushI SampleDry, Array("F", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "F"))
PushI SampleDry, Array("G", True, CByte(8), 1, 2&, 3#, 4!, 5@, Now, String(300, "G"))
End Function

Function SampleDs() As Ds
Set SampleDs = Ds(ApDtAy(SampleDt1, SampleDt2), "SampleDs")
End Function

Function SampleDt1() As Dt
Set SampleDt1 = Dt("SampleDt1", "A B C", SampleDry1)
End Function

Function SampleDt2() As Dt
Set SampleDt2 = Dt("SampleDt2", "A B C", SampleDry2)
End Function
