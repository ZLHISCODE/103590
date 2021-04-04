VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Private mlngFileNum As Long
Private mstrOpenFile As String


Public Function OpenLog(Optional ByVal strFileName As String = "") As Boolean
'打开日志文件
    Dim strCurFileName As String
    
    strCurFileName = strFileName
    
    If Trim(strCurFileName) = "" Then
        strCurFileName = App.Path & "\" & "ZL9PACSWORK(" & Format(Now, "yyyymmdd") & ").txt"
    End If
    
    If mstrOpenFile <> strCurFileName Then
        If mlngFileNum <> 0 Then Call CloseLog
        
        mlngFileNum = FreeFile()
        Open strCurFileName For Append As mlngFileNum
        
        mstrOpenFile = strCurFileName
    End If
    
    OpenLog = IIf(mlngFileNum = 0, False, True)
End Function

Public Sub WriteLog(ByVal strLog As String, Optional ByVal blnUseTimeStamp As Boolean = True, Optional ByVal blnUseMsec As Boolean = True)
'写入日志内容
    Print #mlngFileNum, IIf(blnUseTimeStamp, Now & IIf(blnUseMsec, "." & (Timer() * 1000 Mod 1000), "") & "：", "") & strLog
End Sub


Public Sub CloseLog(Optional ByVal blnCRLF As Boolean = False)
'关闭日志文件
    If mlngFileNum <> 0 Then
        If blnCRLF Then
            Print #mlngFileNum, ""
            Print #mlngFileNum, ""
        End If
        
        Close #mlngFileNum
        
        mstrOpenFile = ""
        mlngFileNum = 0
    End If
End Sub

Private Sub Class_Terminate()
    Call CloseLog(True)
End Sub
