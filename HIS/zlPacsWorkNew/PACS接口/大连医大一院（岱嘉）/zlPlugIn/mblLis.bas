Attribute VB_Name = "mblLis"
Public glngCount As Long
Public gblnInit As Boolean
Public gblnDebug As Boolean

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOW As Long = 5
Public gFso As New FileSystemObject


Public Sub WriteLog(ByVal strOutput As String)
    '------------------------------------------------------
    '--  功能:根据调试标志,写日志到当前目录
    '------------------------------------------------------
    
    '以下变量用于记录调用接口的入参
    Dim strDate As String
    Dim strFileName As String
    Dim objStream As TextStream
    
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    strFileName = App.Path & "\zlPlugIn_" & Format(Date, "yyyyMMdd") & ".LOG"
    If Not gFso.FileExists(strFileName) Then Call gFso.CreateTextFile(strFileName)
    Set objStream = gFso.OpenTextFile(strFileName, ForAppending)
    objStream.WriteLine (strDate & ":" & strOutput)
    objStream.Close
    Set objStream = Nothing
End Sub

 

