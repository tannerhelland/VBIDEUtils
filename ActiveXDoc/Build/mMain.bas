Attribute VB_Name = "mMain"
Option Explicit

Public Sub Main()
Dim sCmd As String
    sCmd = Command
    frmDocHelp.CommandLine = sCmd
    frmDocHelp.Show
End Sub
