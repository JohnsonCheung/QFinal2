Attribute VB_Name = "MIde_Mth_Fb_Gen"
Option Explicit

Private Function W() As Database
Set W = MthDb
End Function

Sub GenCrtMdDic()
DbtDrp W, "MdDic"
DbtCrtFldLisTbl W, "DistMth", "MdDic", "ToMd", "LinesLis", vbCrLf & vbCrLf, True, "Lines"
End Sub
Sub UpdMthLoc()
DbtDrp W, "#A"
Q = "Select x.Nm into [#A] from DistMth x left join MthLoc a on x.Nm=a.Nm where IsNull(a.Nm)": W.Execute Q
Q = "Insert into MthLoc (Nm) Select Nm from [#A]": W.Execute Q
End Sub

Sub GenCrtDistMth()
DbttDrp W, "DistMth #A #B"
Q = "Select Distinct Nm,Count(*) as LinesIdCnt Into DistMth from DistLines group by Nm": W.Execute Q
Q = "Alter Table DistMth Add Column LinesIdLis Text(255), LinesLis Memo, ToMd Text(50)": W.Execute Q
DbtCrtFldLisTbl W, "DistLines", "#A", "Nm", "LinesId", " ", True
DbtCrtFldLisTbl W, "DistLines", "#B", "Nm", "Lines", vbCrLf & vbCrLf, True
Q = "Update DistMth x inner join [#A] a on x.Nm = a.Nm set x.LinesIdLis = a.LinesIdLis": W.Execute Q
Q = "Update DistMth x inner join [#B] a on x.Nm = a.Nm set x.LinesLis = a.LinesLis": W.Execute Q
Q = "Update DistMth x inner join MthLoc a on x.Nm = a.Nm set x.ToMd = IIf(a.ToMd='','AAMod',a.ToMd)": W.Execute Q
End Sub
