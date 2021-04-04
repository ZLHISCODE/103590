Attribute VB_Name = "mdlNetResouce"
Option Explicit
Private Const RESOURCE_CONNECTED       As Long = &H1&
Private Const RESOURCEUSAGE_CONNECTABLE       As Long = &H1&
Private Const NO_ERROR = 0
Private Const RESOURCE_ENUM_ALL       As Long = &HFFFF

Private Type NETRESOURCE
        dwScope   As Long
        dwType   As Long
        dwDisplayType   As Long
        dwUsage   As Long
        pLocalName   As Long
        pRemoteName   As Long
        pComment   As Long
        pProvider   As Long
End Type
Private Type NETRESOURCE_REAL
        dwScope   As Long
        dwType   As Long
        dwDisplayType   As Long
        dwUsage   As Long
        sLocalName   As String
        sRemoteName   As String
        sComment   As String
        sProvider   As String
End Type
Private Declare Function WNetOpenEnum Lib "mpr.dll" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, lpNetResource As Any, lplngEnum As Long) As Long
Private Declare Function WNetEnumResource Lib "mpr.dll" Alias "WNetEnumResourceA" (ByVal lngEnum As Long, lpcCount As Long, lpBuffer As NETRESOURCE, lpBufferSize As Long) As Long
Private Declare Function WNetCloseEnum Lib "mpr.dll" (ByVal lngEnum As Long) As Long
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (lpTo As Any, lpFrom As Any, ByVal lLen As Long)
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Public Function zlNetCancelConnected(Optional strIp As String = "", Optional strComputerName As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：终止磁盘网终资源连接
    '返回：终止成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-04-28 16:04:34
    '说明：只要有一个连接没结成功,则也返回false,否则返回true
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngReturn     As Long, lngEnum    As Long, lngCount    As Long
    Dim lngMin     As Long, lngLength    As Long, l    As Long
    Dim lngBufferSize     As Long, lngLastIndex    As Long
    Dim uNetApi(0 To 256) As NETRESOURCE
    Dim uNet()     As NETRESOURCE_REAL
    Dim blnReturn As Boolean
    
    lngReturn = WNetOpenEnum(RESOURCE_CONNECTED, 0, RESOURCEUSAGE_CONNECTABLE, ByVal 0&, lngEnum)
    If lngReturn = NO_ERROR Then
            lngCount = RESOURCE_ENUM_ALL
            lngBufferSize = UBound(uNetApi) * Len(uNetApi(0)) / 2
            lngReturn = WNetEnumResource(lngEnum, lngCount, uNetApi(0), lngBufferSize)
            If lngCount > 0 Then
                    ReDim Preserve uNet(0 To lngMin + lngCount - 1) As NETRESOURCE_REAL
                    For l = 0 To lngCount - 1
                            'Each   Resource   will   appear   here   as   uNet(i)
                            uNet(lngMin + l).dwScope = uNetApi(l).dwScope
                            uNet(lngMin + l).dwType = uNetApi(l).dwType
                            uNet(lngMin + l).dwDisplayType = uNetApi(l).dwDisplayType
                            uNet(lngMin + l).dwUsage = uNetApi(l).dwUsage
                            If uNetApi(l).pLocalName Then
                                lngLength = lstrlen(uNetApi(l).pLocalName)
                                uNet(lngMin + l).sLocalName = Space$(lngLength)
                                CopyMem ByVal uNet(lngMin + l).sLocalName, ByVal uNetApi(l).pLocalName, lngLength
                            End If
                            If uNetApi(l).pRemoteName Then
                                    lngLength = lstrlen(uNetApi(l).pRemoteName)
                                    uNet(lngMin + l).sRemoteName = Space$(lngLength)
                                    CopyMem ByVal uNet(lngMin + l).sRemoteName, ByVal uNetApi(l).pRemoteName, lngLength
                            End If
                            If uNetApi(l).pComment Then
                                    lngLength = lstrlen(uNetApi(l).pComment)
                                    uNet(lngMin + l).sComment = Space$(lngLength)
                                    CopyMem ByVal uNet(lngMin + l).sComment, ByVal uNetApi(l).pComment, lngLength
                            End If
                            If uNetApi(l).pProvider Then
                                    lngLength = lstrlen(uNetApi(l).pProvider)
                                    uNet(lngMin + l).sProvider = Space$(lngLength)
                                    CopyMem ByVal uNet(lngMin + l).sProvider, ByVal uNetApi(l).pProvider, lngLength
                            End If
                    Next l
            Else
                zlNetCancelConnected = True
                Exit Function
            End If
            If lngEnum Then
                l = WNetCloseEnum(lngEnum)
            End If
    End If
    
    '结束连接
    blnReturn = True
  
    For l = 0 To UBound(uNet)
'        WriteTxtLog "开始并结束检查共享网终连接:" & uNet(l).sRemoteName
        If uNet(l).sRemoteName Like "\\" & strIp & "\*" Or uNet(l).sRemoteName Like "\\" & strIp & "\*" Or (strIp = "" And strComputerName = "") Then
            If CancelNetServer(IIf(uNet(l).sLocalName = "", uNet(l).sRemoteName, uNet(l).sLocalName)) = False Then
                blnReturn = False
            End If
        End If
    Next
    zlNetCancelConnected = blnReturn
End Function
