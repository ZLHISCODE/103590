Attribute VB_Name = "mdlMain"
Option Explicit
Public gblnTest As Boolean

Sub Main()
    gblnTest = False
    frmFtpMain.Show
    
    If gblnTest Then
        frmFtpMain.Test
    End If
End Sub


Public Function FormatPath(ByVal strPath As String) As String
    FormatPath = Mid(strPath, 1, 2) & Replace(strPath, "\\", "\", 3)
End Function
