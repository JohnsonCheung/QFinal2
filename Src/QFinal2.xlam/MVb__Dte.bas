Attribute VB_Name = "MVb__Dte"
Option Explicit
Function CurM() As Byte
CurM = Month(Now)
End Function

Function M_NxtM(M As Byte) As Byte
M_NxtM = IIf(M = 12, 1, M + 1)
End Function

Function M_PrvM(M As Byte) As Byte
M_PrvM = IIf(M = 1, 12, M - 1)
End Function

Function NowDTim$()
NowDTim = DteDTim(Now)
End Function

Function NowStr$()
NowStr = Format(Now, "YYYY-MM-DD HH:MM:SS")
End Function

Function DteDTim$(Dte)
If Not IsDate(Dte) Then Exit Function
DteDTim = Format(Dte, "YYYY-MM-DD HH:MM:SS")
End Function

Function DteFstDayOfMth(A As Date) As Date
DteFstDayOfMth = DateSerial(Year(A), Month(A), 1)
End Function

Function DteFstDteOfMth(A As Date) As Date
DteFstDteOfMth = DateSerial(Year(A), Month(A), 1)
End Function

Function DteIsVdt(A$) As Boolean
On Error Resume Next
DteIsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function DteLasDayOfMth(A As Date) As Date
DteLasDayOfMth = DtePrvDay(DteFstDteOfMth(DteNxtMth(A)))
End Function

Function DteNxtMth(A As Date) As Date
DteNxtMth = DateTime.DateAdd("M", 1, A)
End Function

Function DtePrvDay(A As Date) As Date
DtePrvDay = DateAdd("D", -1, A)
End Function

Function DteYYMM$(A As Date)
DteYYMM = Right(Year(A), 2) & Format(Month(A), "00")
End Function

Function YYMM_FstDte(A) As Date
YYMM_FstDte = DateSerial(Left(A, 2), Mid(A, 3, 2), 1)
End Function

Function YYYYMMDD_IsVdt(A) As Boolean
On Error Resume Next
YYYYMMDD_IsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function

Function YM_FstDte(Y As Byte, M As Byte) As Date
YM_FstDte = DateSerial(2000 + Y, M, 1)
End Function

Function YM_LasDte(Y As Byte, M As Byte) As Date
YM_LasDte = DteNxtMth(YM_FstDte(Y, M))
End Function

Function YM_YofNxtM(Y As Byte, M As Byte) As Byte
YM_YofNxtM = IIf(M = 12, Y + 1, Y)
End Function

Function YM_YofPrvM(Y As Byte, M As Byte) As Byte
YM_YofPrvM = IIf(M = 1, Y - 1, Y)
End Function

Function CurY() As Byte
CurY = CurYY - 2000
End Function

Function CurYY%()
CurYY = Year(Now)
End Function
