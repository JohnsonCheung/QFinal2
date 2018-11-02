Attribute VB_Name = "MVb__Re"
Option Explicit
Private Sub ZZ_ReMatch()
Dim A As MatchCollection
Dim R  As RegExp: Set R = Re("m[ae]n")
Set A = R.Execute("alskdflfmensdklf")
Stop
End Sub

Private Sub ZZ_ReRpl()
Dim R As RegExp: Set R = Re("(.+)(m[ae]n)(.+)")
Dim Act$: Act = R.Replace("a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub
