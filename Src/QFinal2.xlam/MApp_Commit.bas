Attribute VB_Name = "MApp_Commit"
Option Explicit
Sub XX()
Dim O$()
Push O, "echo ""# QLib"" >> README.md"
Push O, "git Init"
Push O, "git add README.md"
Push O, "git commit -m ""first commit"""
Push O, "git remote add origin https://github.com/JohnsonCheung/QLib.git"
Push O, "git push -u origin master"

End Sub
Sub XX1()
'git remote add origin https://github.com/JohnsonCheung/QLib.git
'git push -u origin master
End Sub
Sub Export()
CurVbeExp
End Sub

Sub AppCommit(Optional Msg$ = "Commit")
AppExp
FcmdRunMax BldCommitFcmd, Msg
End Sub

Function AppCommitFcmd$()
AppCommitFcmd = WPth & "Commit.Cmd"
End Function

Sub Commit(Optional Msg$ = "Commit")
AppCommit Msg
End Sub

Private Function BldCommitFcmd$()
Dim O$(), Cd$, GitAdd$, GitCommit$, GitPush, T$
Cd = FmtQQ("Cd ""?""", SrcPth)
GitAdd = "git add -A"
GitCommit = "git commit --message=%1%"
Push O, Cd
Push O, GitAdd
Push O, GitCommit
Push O, "Pause"
T = AppCommitFcmd
AyWrt O, T
BldCommitFcmd = T
End Function

Sub FcommitBrw()
FtBrw BldCommitFcmd
End Sub

Sub AppPush()
FcmdRunMax BldPushAppFcmd
End Sub

Private Function BldPushAppFcmd$()
Dim O$(), Cd$, GitPush, T
Cd = FmtQQ("Cd ""?""", SrcPth)
GitPush = "git push -u https://johnsoncheung@github.com/johnsoncheung/StockShipRate.git master"
Push O, Cd
Push O, GitPush
Push O, "Pause"
T = AppPushAppFcmd
AyWrt O, T
BldPushAppFcmd = T
End Function
