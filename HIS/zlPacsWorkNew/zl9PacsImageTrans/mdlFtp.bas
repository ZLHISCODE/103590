Attribute VB_Name = "mdlFtp"
Option Explicit

Public Type TFtpConTag
    Ip As String             'IP
    Port As Long             '¶Ë¿Ú
    User As String           'User
    pwd As String            'ÃÜÂë
    
    VirtualPath As String    'ÐéÄâÄ¿Â¼
    
    ShareDir As String
    ShareUser As String
    SharePwd As String
End Type

Public Function FtpTagInstance(ByVal strIP As String, _
    ByVal strUser As String, _
    ByVal strPwd As String, _
    ByVal strVirtualPath As String, Optional ByVal lngPort As Long) As TFtpConTag

    FtpTagInstance = FtpCreateTag(strIP, strUser, strPwd, strVirtualPath, lngPort)
End Function


Public Function FtpCreateTag(ByVal strIP As String, _
    ByVal strUser As String, _
    ByVal strPwd As String, _
    ByVal strVirtualPath As String, _
    Optional ByVal lngPort As Long, _
    Optional ByVal strShareDir As String, _
    Optional ByVal strShareUser As String, _
    Optional ByVal strSharePwd As String) As TFtpConTag

    FtpCreateTag.Ip = ""
    If strIP = "" Then Exit Function
    
    FtpCreateTag.Ip = strIP
    FtpCreateTag.Port = lngPort
    FtpCreateTag.User = strUser
    FtpCreateTag.pwd = strPwd
    FtpCreateTag.VirtualPath = strVirtualPath
    FtpCreateTag.ShareDir = strShareDir
    FtpCreateTag.ShareUser = strShareUser
    FtpCreateTag.SharePwd = strSharePwd
End Function
