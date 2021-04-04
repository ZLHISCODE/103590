VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFtpGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit




Public Function FtpTagInstance(ByVal strIP As String, _
    ByVal strUser As String, _
    ByVal strPwd As String, _
    ByVal strVirtualPath As String, Optional ByVal lngPort As Long) As TFtpConTag

    FtpTagInstance = FtpCreateTag(strIP, strUser, strPwd, strVirtualPath, lngPort)
End Function


Public Function FtpConnectTest(ByRef ftpTag As TFtpConTag) As Boolean
'ftp 连接测试
    FtpConnectTest = FtpIsValid(ftpTag)
End Function

'删除文件
Public Function FtpDelete(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, _
    Optional ByVal blnIsAutoHint As Boolean = True) As Long
    FtpDelete = FtpDeleteFile(ftpTag, strFtpFile, blnIsAutoDiscon, blnIsAutoHint)
End Function

'下载文件
Public Function FtpDownload(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, _
    Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult

    FtpDownload = FtpDownloadFile(ftpTag, strFtpFile, strLocalFile, blnIsAutoDiscon, blnIsAutoHint)
    
End Function

'上传文件
Public Function FtpUpload(ByRef ftpTag As TFtpConTag, ByVal strFtpFile As String, ByVal strLocalFile As String, _
    Optional ByVal blnIsAutoDiscon As Boolean = True, _
    Optional ByVal blnIsAutoHint As Boolean = True) As FtpResult
    
    FtpUpload = FtpUploadFile(ftpTag, strFtpFile, strLocalFile, blnIsAutoDiscon)
    
End Function





Private Sub Class_Initialize()
    Call FtpParInit
End Sub

Private Sub Class_Terminate()
    Call FtpFree
End Sub
