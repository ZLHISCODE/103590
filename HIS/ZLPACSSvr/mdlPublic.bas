Attribute VB_Name = "mdlPublic"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'---------------------------------------------------------------
'-注册表 API 声明...
'---------------------------------------------------------------
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Public Const REG_SZ = 1                         ' Unicode空终结字符串
Public Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Public Const REG_DWORD = 4                      ' 32-bit 数字
' 注册表创建类型值...
Public Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留
' 注册表关键字安全选项...
Public Const READ_CONTROL = &H20000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Public Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Public Const KEY_EXECUTE = KEY_READ
Public Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
' 注册表关键字根类型...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
' 返回值...
Public Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'---------------------------------------------------------
'窗体相关
'---------------------------------------------------------
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1


'判断数组是否为空
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

'读取网卡的多个IP
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const WSADescription_Len = 256
Private Const WSASYS_Status_Len = 128

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Private Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Integer, lpWSAData As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'定义DICOM服务相关的结构体和数组
Public Type AEconnection        '记录连接信息，类似于DICOM控件中DicomConnection的作用
    Association As Long         '记录当前连接的id
    ServiceAE As String                '被呼叫的AE名称
    DeviceIP As String                '设备IP地址
    TimeStamp As String         '时间戳，记录连接建立的时间
    Deleted As Boolean          '删除标记，是否被删除
End Type
Public AEconnections() As AEconnection  '存储连接信息的数组

Public Type Service
    DeviceIP As String          '记录设备的IP地址
    DeviceAE As String          '记录设备的AE名称
    DevicePort As String        '记录设备的端口
    DeviceName As String        '记录设备名称
    ServiceAE As String         '记录PACS服务的AE名称
    ServicePort As String       '记录PACS服务的端口号
    SOP As String               '记录服务功能
    Modality As String          '记录设备的影像类别
    Started  As Boolean         '记录当前服务是否成功启动
End Type
Public Services() As Service    '存储应用于当前IP地址的DICOM服务对

Public Type AEPara              '记录各个服务的简单参数
    AE As String                '记录被呼叫的AE名称
    IP As String                '记录设备IP地址
    ParaName As String          '参数名称
    ParaValue As String         '参数值
End Type
Public AEParas() As AEPara      '存储应用于当前IP地址的参数


Public Type FTPDevice           '记录FTP存储设备
    No As String                '存储设备号
    IP As String                'IP地址
    User As String              '用户名
    Password As String          '密码
    FTPDir As String            'FTP目录
End Type
Public FTPDevices() As FTPDevice        '存储应用于当前IP的FTP存储设备

Public gstrLocalIP As String             '存储本机IP地址
Private iNet As New clsFtp      '作为公共参数的目的是，以后修改成FTP设备号不改变的时候，不再重连FTP

Public Const ATTR_检查日期 As String = "8:20"
Public Const ATTR_检查时间 As String = "8:30"
Public Const ATTR_采集日期 As String = "8:22"
Public Const ATTR_采集时间 As String = "8:32"
Public Const ATTR_图像日期 As String = "8:23"
Public Const ATTR_图像时间 As String = "8:33"


Public Const ATTR_影像类别 As String = "8:60"
Public Const ATTR_检查设备 As String = "8:1090"
Public Const ATTR_长宽比 As String = "28:34"
Public Const ATTR_序列号 As String = "20:11"
Public Const ATTR_图像号 As String = "20:13"
Public Const ATTR_图像类型 As String = "8:8"

Public Const ATTR_层厚 As String = "18:50"
Public Const ATTR_图像位置病人 As String = "20:32"
Public Const ATTR_图像方向病人 As String = "20:37"
Public Const ATTR_参考帧UID As String = "20:52"
Public Const ATTR_切片位置 As String = "20:1041"
Public Const ATTR_行数 As String = "28:10"
Public Const ATTR_列数 As String = "28:11"
Public Const ATTR_像素距离 As String = "28:30"

Public Const TS_JPEG无损压缩 As String = "1.2.840.10008.1.2.4.70"
Public Const TS_RLE行程压缩 As String = "1.2.840.10008.1.2.5"
Public Const TS_JPEG2000无损压缩 As String = "1.2.840.10008.1.2.4.90"

Public gcnAccess As New ADODB.connection, strBeginDate As String
Public gstrAccessPath As String         'Access数据库的文件路径和文件名前缀，无“.mdb”
Public gstrAccessName As String         'Access数据库的文件路径和文件名

Public gstrSQL As String


Public Function funGetFTPDevice(strDeviceNO As String, strIP As String, strUser As String, strPsw As String, strFTPDir As String) As Boolean
    Dim i As Integer
    
    For i = 1 To UBound(FTPDevices)
        If FTPDevices(i).No = strDeviceNO Then
            strIP = FTPDevices(i).IP
            strUser = FTPDevices(i).User
            strPsw = FTPDevices(i).Password
            strFTPDir = FTPDevices(i).FTPDir
            Exit For
        End If
    Next i
    If i <= UBound(FTPDevices) Then
        funGetFTPDevice = True
    Else
        funGetFTPDevice = False
    End If
End Function

Public Function funGetQRParas(strServiceAE As String, strDeviceIP As String, blnCGet As Boolean, _
    intPatientIDMatch As Integer)
    Dim i As Integer
    
    '读取基本参数
    intPatientIDMatch = 0
    blnCGet = False
    
    For i = 1 To UBound(AEParas)
        If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
            Select Case AEParas(i).ParaName
            Case ZLPACS_QR允许CGET
                blnCGet = AEParas(i).ParaValue
            Case ZLPACS_QR病人ID匹配
                intPatientIDMatch = AEParas(i).ParaValue
            End Select
        End If
    Next i
    funGetQRParas = True
End Function

Public Function funGetAEMWLParas(strServiceAE As String, strDeviceIP As String, intFilterModality As Integer, _
        intDayInterval As Integer, blnUseForceResult As Boolean, intBodypartType As Integer, _
        strBodypartSplitter As String, intResultFilter As Integer) As Boolean
    Dim i As Integer
    
    '初始化参数
    intDayInterval = 3
    intFilterModality = 0
    intBodypartType = 0
    strBodypartSplitter = ""
    intResultFilter = 0
    
    '读取基本参数
    If SafeArrayGetDim(AEParas) <> 0 Then
        '记录处理日志
        Call WriteProcessLog("funGetAEMWLParas", "读取Worklist的参数", "UBound(AEParas) = " & UBound(AEParas) & " ,UCase(strServiceAE) = " & UCase(strServiceAE) & " , strDeviceIP = " & strDeviceIP)
    
        For i = 1 To UBound(AEParas)
            If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
                Select Case AEParas(i).ParaName
                Case ZLPACS_MWL过滤方式
                    intFilterModality = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL检索天数
                    intDayInterval = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL用强制结果
                    blnUseForceResult = AEParas(i).ParaValue
                Case ZLPACS_MWL多部位方式
                    intBodypartType = Val(AEParas(i).ParaValue)
                Case ZLPACS_MWL多部位分隔符
                    strBodypartSplitter = AEParas(i).ParaValue
                Case ZLPACS_MWL查询结束条件
                    intResultFilter = Val(AEParas(i).ParaValue)
                End Select
            End If
        Next i
    End If
    
    funGetAEMWLParas = True
End Function
    
Private Function GetAEconnection(ByVal Association As Long, ByRef strServiceAE As String, ByRef strDeviceIP As String) As Boolean
    
    Dim i As Integer
    '查找服务AE和IP
    For i = 1 To UBound(AEconnections)
        If AEconnections(i).Association = Association Then
            strServiceAE = AEconnections(i).ServiceAE
            strDeviceIP = AEconnections(i).DeviceIP
            Exit For
        End If
    Next i
    
    If i <= UBound(AEconnections) Then
        GetAEconnection = True
    Else
        GetAEconnection = False
    End If
End Function

Private Function GetFilmStor(ByVal iService As Long, ByRef strServiceAE As String, ByRef strDeviceIP As String) As Boolean
    
    On Error GoTo err
    strServiceAE = Services(iService).ServiceAE
    strDeviceIP = Services(iService).DeviceIP
    
    GetFilmStor = True
    Exit Function
err:
    GetFilmStor = False
End Function


Public Function funGetAEStoreParas(ByVal Association As String, ByVal Modality As String, ByRef strIPAddress As String, ByRef blnSplitSeriesUID As Boolean, ByRef intImageMatchItem As Integer, _
    ByRef intDBMatchItem As Integer, ByRef blnMatchStudyUID As Boolean, ByRef strStoreDeviceNo As String, ByRef intEncode As Integer, _
    ByRef strAutoRoute As String, ByRef intFilterModality As Integer, ByRef strAutoRouteCompression, ByRef strAutoRouteDir, ByRef strPhysician) As Boolean
    
'    '服务参数设置
    Dim i As Integer
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strServiceAE As String      'PACS服务的AE名称
    Dim strDeviceIP As String       '设备IP地址
    Dim blnRet As Boolean
    
    blnRet = GetAEconnection(Val(Association), strServiceAE, strDeviceIP)
    
    '找不到对应的服务AE，记录查找失败，然后从Services 中找到一个影像类别相同的设备，读取这个设备的参数
    If blnRet = False Then
        WriteLog 41, vbObjectError + 1, "通过Association，找不到对应的服务AE,Association = " & Association & vbCrLf _
                & " UBound(AEconnections) = " & UBound(AEconnections) & " 影像类别 =" & Modality
                
        For i = 1 To UBound(Services)
            If UCase(Services(i).Modality) = UCase(Modality) And Services(i).Started = True Then
                strServiceAE = Services(i).ServiceAE
                strDeviceIP = Services(i).DeviceIP
                WriteLog 42, vbObjectError + 1, "根据影像类别查找到该图像对应的服务AE和设备IP，ServiceAE = " & strServiceAE & vbCrLf _
                    & " DeviceIP = " & strDeviceIP
                Exit For
            End If
        Next i
        If strServiceAE = "" Or strDeviceIP = "" Then
            WriteLog 43, vbObjectError + 1, "错误，找不到该图像对应的服务AE，图像无法保存。"
            funGetAEStoreParas = False
            Exit Function
        End If
    End If
    
    '返回设备IP地址
    strIPAddress = strDeviceIP
    
    '初始化参数
    blnSplitSeriesUID = False
    blnMatchStudyUID = True
    strStoreDeviceNo = ""
    intEncode = 0
    intImageMatchItem = 0
    intDBMatchItem = 0
    strAutoRoute = ""
    strAutoRouteCompression = ""
    strAutoRouteDir = ""
    intFilterModality = 0
    strPhysician = ""
    
    '读取基本参数
    If SafeArrayGetDim(AEParas) <> 0 Then
        For i = 1 To UBound(AEParas)
            If UCase(AEParas(i).AE) = UCase(strServiceAE) And AEParas(i).IP = strDeviceIP Then
                Select Case AEParas(i).ParaName
                Case ZLPACS_按图像类型拆分序列
                    blnSplitSeriesUID = AEParas(i).ParaValue
                Case ZLPACS_存储设备号
                    strStoreDeviceNo = AEParas(i).ParaValue
                Case ZLPACS_启用检查UID匹配
                    blnMatchStudyUID = AEParas(i).ParaValue
                Case ZLPACS_压缩方式
                    If AEParas(i).ParaValue = "JPEG无损压缩" Then
                        intEncode = 0
                    ElseIf AEParas(i).ParaValue = "RLE压缩" Then
                        intEncode = 1
                    ElseIf AEParas(i).ParaValue = "JPEG2000压缩" Then
                        intEncode = 2
                    Else    '不压缩
                        intEncode = 3
                    End If
                Case ZLPACS_数据库匹配项
                    intDBMatchItem = Val(AEParas(i).ParaValue)
                Case ZLPACS_图像匹配项
                    intImageMatchItem = Val(AEParas(i).ParaValue)
                Case ZLPACS_自动路由
                    strAutoRoute = AEParas(i).ParaValue
                Case ZLPACS_自动路由压缩方式
                    strAutoRouteCompression = AEParas(i).ParaValue
                Case ZLPACS_自动路由目录结构
                    strAutoRouteDir = AEParas(i).ParaValue
                Case ZLPACS_存储过滤方式
                    intFilterModality = Val(AEParas(i).ParaValue)
                Case ZLPACS_提取检查技师
                    If InStr(AEParas(i).ParaValue, ":") > 0 Then
                        strPhysician = AEParas(i).ParaValue
                    End If
                End Select
            End If
        Next i
    End If
    
    '如果没有定义存储设备号，则使用数据库中第一个存储设备
    If strStoreDeviceNo = "" Then
        strSQL = "Select 设备号,设备名 From 影像设备目录 Where 类型=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取AE服务基本参数", CLng(1))
        
        If rsTmp.EOF Then
            WriteLog 4, vbObjectError + 1, "未定义影像存储设备，请到影像设备目录中设置！"
            funGetAEStoreParas = False
            Exit Function
        Else
            strStoreDeviceNo = rsTmp(0)
        End If
    End If
    
    funGetAEStoreParas = True
End Function

Private Function funGetStudyUID(ByVal strOldStudyUID As String) As String
'-----------------------------------------------------------------------------
'功能:查询数据库，判断当前图像的检查UID是否已经存在于正常表和临时表中，
'     如果存在，则在检查UID后面增加后缀，不存在则直接返回输入的检查UID
'修改人:黄捷
'修改日期:2007-1-27
'-----------------------------------------------------------------------------
    Dim rsMatch As New ADODB.Recordset
    
    funGetStudyUID = strOldStudyUID
    gstrSQL = "select 检查UID from 影像检查记录 where 检查UID = [1]" & _
              " Union All Select 检查UID from 影像临时记录 where 检查UID = [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strOldStudyUID)
    If Not rsMatch.EOF Then
        '创建一个新的检查UID
        gstrSQL = "Select 影像检查UID序号_ID.Nextval From Dual"
        Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存")
        If Len(strOldStudyUID) <= 55 Then
            funGetStudyUID = strOldStudyUID & ".A" & rsMatch(0)
        Else
            funGetStudyUID = Left(strOldStudyUID, 55) & ".A" & rsMatch(0)
        End If
    End If
End Function

Public Function WriteToURL(ByRef ftpNet As clsFtp, ByVal SrcFileName As String, ByVal DestFileName As String) As Long
'功能：将本地文件保存到远程网络上
    Dim objFileSystem As New Scripting.FileSystemObject
    
    WriteToURL = 0  '正确
    
    '创建远程目路
    WriteToURL = ftpNet.FuncFtpMkDir("/", objFileSystem.GetParentFolderName(DestFileName))
    
    Call WriteProcessLog("WriteToURL", " 创建远程目录", "创建远程目录，结果为： " & IIf(WriteToURL = 1, "失败", "成功") & "，远程目录为： " & objFileSystem.GetParentFolderName(DestFileName))
    
    '目录创建成功再上传图像
    If WriteToURL = 1 Then Exit Function
    WriteToURL = ftpNet.FuncUploadFile(objFileSystem.GetParentFolderName(DestFileName), SrcFileName, objFileSystem.GetFileName(DestFileName))
    
    Call WriteProcessLog("WriteToURL", "上传图像", "上传图像结果为：" & IIf(WriteToURL = 0, "成功", IIf(WriteToURL = 1, "连接失败", "上传失败")) & "， 文件名为：" & SrcFileName)
    
End Function

Public Function GetImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String) As String
'-----------------------------------------------------------------------------
'功能:提取DICOM属性集中的指定属性值,根据VM判断值的维度，使用“\”把各个维度连接成一个串
'参数： objAttr ----属性集合
'       AttrName ----要查找的属性名称
'返回值：属性的内容
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    Dim i As Integer
    
    GetImageAttribute = ""
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM = 1 Then
            GetImageAttribute = NVL(objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(1))
        Else
            For i = 1 To objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).VM
                GetImageAttribute = GetImageAttribute & "\" & objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).ValueByIndex(i)
            Next i
        End If
    End If
End Function

Public Sub DeleteImageAttribute(objAttr As DicomAttributes, ByVal AttrName As String)
'-----------------------------------------------------------------------------
'功能:删除DICOM属性集中的指定属性值
'-----------------------------------------------------------------------------
    Dim AttrTag() As String
    
    AttrTag = Split(AttrName, ":")
    If objAttr("&h" & AttrTag(0), "&h" & AttrTag(1)).Exists Then
        Call objAttr.Remove("&h" & AttrTag(0), "&h" & AttrTag(1))
    End If
End Sub

Public Sub ResizeRegion(ByVal ImageCount As Integer, ByVal RegionWidth As Long, _
    ByVal RegionHeight As Long, Rows As Integer, Cols As Integer, _
    Optional ByVal MaxRows As Integer = 0, Optional ByVal MaxCols As Integer = 0)
'功能：计算DicomViewer的行列数
    Dim iCols As Integer, iRows As Integer
    
    iCols = CInt(Sqr(ImageCount * RegionWidth / RegionHeight))
    iRows = CInt(Sqr(ImageCount * RegionHeight / RegionWidth))
    
    Do While iRows * iCols < ImageCount
        If RegionWidth / RegionHeight > 1 Then
            iCols = iCols + 1
        Else
            iRows = iRows + 1
        End If
    Loop
    If MaxRows > 0 And iRows > MaxRows Then
        iRows = MaxRows
        iCols = CInt(ImageCount / iRows)
        If iRows * iCols < ImageCount Then iCols = iCols + 1
    End If
    If MaxCols > 0 And iCols > MaxCols Then
        iCols = MaxCols
        iRows = CInt(ImageCount / iCols)
        If iRows * iCols < ImageCount Then iRows = iRows + 1
    End If
    If MaxRows > 0 And iRows > MaxRows Then iRows = MaxRows
    
    Rows = iRows: Cols = iCols
End Sub

Public Function ImageExist(Images As DicomImages, SeekImage As DicomImage) As Boolean
    Dim curImage As DicomImage
    
    ImageExist = False
    For Each curImage In Images
        If curImage.instanceUID = SeekImage.instanceUID Then ImageExist = True: Exit For
    Next
End Function

Private Sub WriteRecord(ByVal ImageType As String, ByVal strCheckNo As String, ByVal CheckDev As String, _
    ByVal PatientName As String, ByVal EnglishName As String, ByVal Sex As String, Age As Integer, _
    ByVal CheckUID As String, ByVal SeriesUID As String, ByVal ifTmp As Boolean)
'-----------------------------------------------------------------------------
'功能:保存影像接收序列，保存到本地Access的数据库文件中
'参数： ImageType ----影像类别
'       strCheckNo ----图像中的匹配ID，可能是PatientID，PatientName，AccessionNumber
'       CheckDev ----检查设备
'       PatientName ----姓名
'       EnglishName ----英文名
'       Sex ----性别
'       Age ----年龄
'       CheckUID ----检查UID
'       SeriesUID ----序列UID
'       ifTmp ----是否临时记录
'返回值：直接插入“影像接收序列”表
'-----------------------------------------------------------------------------
    
    Dim rsTmp As ADODB.Recordset, strSQL As String
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Select id from 影像接收序列 Where 序列UID='" & SeriesUID & "' And 接收时间>cDate('" & _
        strBeginDate & "')"
    Set rsTmp = gcnAccess.Execute(strSQL)
    If rsTmp.EOF Then
        strSQL = "Insert Into 影像接收序列(影像类别,检查号,检查设备,姓名,英文名,性别,年龄,影像数,序列UID,检查UID,对应检查,接收时间)" & _
            " Values('" & ImageType & "'," & IIf(strCheckNo = "0", "Null", "'" & strCheckNo & "'") & ",'" & CheckDev & "','" & _
            funDeleteSpcialChar(PatientName) & "','" & funDeleteSpcialChar(EnglishName) & "','" & Sex & "'," & IIf(Age = -1, "Null", Age) & ",1,'" & _
            SeriesUID & "','" & CheckUID & "'," & CStr(Not ifTmp) & ",cDate('" & _
            Date & " " & Time() & "'))"
    Else
        strSQL = "Update 影像接收序列 Set 影像数=影像数+1 Where 序列UID='" & SeriesUID & "' And 接收时间>cDate('" & _
        strBeginDate & "')"
    End If
    gcnAccess.Execute strSQL
End Sub

Public Sub WriteLog(ByVal ErrorType As Integer, ErrorNum As Long, ErrorDesc As String)
'-----------------------------------------------------------------------------
'功能:填写错误日志
'参数： ErrorType ----错误类型代码，保存图像错误100，WORKLIST和QR错误200，FTP错误300,funSplitSeriesUID错误1001
'       ErrorNum ----错误号
'       ErrorDesc ----错误描述
'返回值：无
'-----------------------------------------------------------------------------
    Dim strSQL As String
    On Error Resume Next
    If gcnAccess.State = adStateClosed Then Exit Sub
    
    strSQL = "Insert Into 错误日志(产生时间,错误类型,错误号,错误信息) " & _
        "Values(cDate('" & Date & " " & Time() & "')," & ErrorType & "," & ErrorNum & ",'" & Replace(ErrorDesc, "'", "''") & "')"
    
    gcnAccess.Execute strSQL
End Sub

Private Function funcAutoRouting(img As DicomImage, BufferDir As String, dtReceived As String, _
    strStudyUID As String, iEncode As Integer, strAutoRoute As String, strAutoRouteCompression As String, _
    strAutoRouteDir As String) As Long
'-----------------------------------------------------------------------------
'功能:自动路由，把图像发送到指定的地方
'参数： img ----需要发送的图像
'       BufferDir---本地缓存路径
'       dtReceived---接收日期，作为图像路径的一部分
'       strStudyUID---检查UID，作为图像路径的一部分，对于手工关联的图像，路经不一定是图像中的检查UID，所以需要从外部传入
'       iEncode---压缩方式
'       strAutoRoute---路由目的地集合，使用“|”分隔各个存储设备号
'       strAutoRouteCompression---自动路由的压缩方法集合，使用“|”分隔各个压缩方式，0--按照当前方式压缩，1--不压缩
'       strAutoRouteDir---自动路由的目录结构集合，使用“|”分隔各个目录结构，0--检查级别目录（默认），1--序列级别目录（3D）
'返回值：无
'-----------------------------------------------------------------------------
    Dim i As Integer            '用于循环的变量
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strDirURL As String         'FTP主机的目录
    Dim strHost As String, strUser As String, strPwd As String
    Dim strRouteDest() As String    '记录自动路由目的地的设备号
    Dim strRouteCompression() As String     '记录自动路由的压缩方式
    Dim strRouteDir() As String     '记录自动路由的目录结构
    Dim thisNet As New clsFtp       'FTP连接
    Dim intCurRouteCompression As Integer
    Dim intCurRouteDir As Integer
    Dim strUploadDir As String      '保存到FTP中的目录名称
    
    If strAutoRoute = "" Then Exit Function
    
    On Error GoTo ProcError
    
    '获取自动路由规则
    strRouteDest = Split(strAutoRoute, "|")
    strRouteCompression = Split(strAutoRouteCompression, "|")
    strRouteDir = Split(strAutoRouteDir, "|")
    '如果自动路由的设备数量和参数数量不一致，则记录错误日志作为提醒
    If UBound(strRouteDest) <> UBound(strRouteCompression) Or UBound(strRouteDest) <> UBound(strRouteDir) Then
        Call WriteLog(201, 100, "图像的检查UID为 " & strStudyUID & " 。自动路由的设备数量和参数数量不一致，可能导致自动路由无法正确完成，请到“影像设备目录”中进行设置。")
    End If
    
    '对比存储规则，不匹配则退出
    For i = 0 To UBound(strRouteDest)
        '从数据库中查找对应的存储设备IP地址和用户名密码
        strSQL = "Select IP地址,FTP用户名,FTP密码,'/'||Decode(Ftp目录,Null,'',Ftp目录||'/') As DirUrl From 影像设备目录 Where 设备号=  [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", strRouteDest(i))
        If rsTmp.EOF Then
            err.Raise vbObjectError + 1, "PACS图像保存", "自动路由 设备号 " & strRouteDest(i) & " 设置错误！"
        End If
        
        strHost = rsTmp!IP地址
        strUser = rsTmp!FTP用户名
        strPwd = NVL(rsTmp!FTP密码)
        strDirURL = rsTmp!DirUrl
        
        '读取自动路由参数
        intCurRouteCompression = 0
        intCurRouteDir = 0
        On Error Resume Next
        intCurRouteCompression = Val(strRouteCompression(i))
        intCurRouteDir = Val(strRouteDir(i))
        
        On Error GoTo ProcError
        '保存图像到指定URL
        If intCurRouteCompression = 1 Then  '不压缩
            img.WriteFile BufferDir & img.instanceUID, True
        Else
            Select Case iEncode
                Case 0
                    img.WriteFile BufferDir & img.instanceUID, True, TS_JPEG无损压缩
                Case 1
                    img.WriteFile BufferDir & img.instanceUID, True, TS_RLE行程压缩
                Case 2
                    img.WriteFile BufferDir & img.instanceUID, True, TS_JPEG2000无损压缩
                Case 3
                    img.WriteFile BufferDir & img.instanceUID, True
            End Select
        End If
        
        '初始Ftp对象,FTP 连接成功，则上传图像
        thisNet.FuncFtpDisConnect
        If thisNet.FuncFtpConnect(strHost, strUser, strPwd) <> 0 Then
            '创建目录成功，则上传图像
            If intCurRouteDir = 1 Then      '序列级别的目录（3D）
                strUploadDir = strDirURL & dtReceived & "/" & strStudyUID & "/" & img.SeriesUID
            Else            '检查级别的目录（默认）
                strUploadDir = strDirURL & dtReceived & "/" & strStudyUID
            End If
            If thisNet.FuncFtpMkDir("/", strUploadDir) <> 1 Then
                Call thisNet.FuncUploadFile(strUploadDir, BufferDir & img.instanceUID, img.instanceUID)
            End If
        End If
        Kill BufferDir & img.instanceUID
    Next
    
    thisNet.FuncFtpDisConnect
    Exit Function
ProcError:
    Call WriteLog(2, err.Number, err.Description)
    thisNet.FuncFtpDisConnect
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''WorkList部分程序''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub AddResultItem(DataSet As DicomDataSet, request As DicomDataSet, group As Long, element As Long, value As Variant)
    '只发送需要的项目
    If request.Attributes(group, element).Exists Then
        If IsNull(value) Then
            value = ""
        End If
        DataSet.Attributes.Add group, element, value
    End If
End Sub

Public Sub AddLinkedDateTimeCondition(ByRef query As String, datecondition As DicomAttribute, timecondition As DicomAttribute, dbname As String)
    Dim startdatetime As Date, enddatetime As Date
    If datecondition.Exists And timecondition.Exists Then
        startdatetime = datecondition.DateTimeFrom("1/1/1800") + timecondition.DateTimeFrom("0")
        enddatetime = datecondition.DateTimeTo("1/1/2999") + timecondition.DateTimeTo("0.9999")
        AddSingleDateCondition query, startdatetime, ">=", dbname
        AddSingleDateCondition query, enddatetime, "<=", dbname
        
    Else
        AddDateCondition query, datecondition, "DateValue(" & dbname & ")"
        AddDateCondition query, timecondition, "TimeValue(" & dbname & ")"
    End If
End Sub

Private Sub AddSingleDateCondition(ByRef query As String, Condition As Date, operator As String, dbname As String)
    ' all date formating goes through here to make it easy to change for different databases or locales
    query = query & " AND " & dbname & operator & "to_Date('" & Condition & "', 'yyyy-mm-dd hh24:mi:ss')"
End Sub

Public Sub AddDateCondition(ByRef query As String, Condition As DicomAttribute, dbname As String)
'-----------------------------------------------------------------------------
'功能:添加日期查询条件,不为*才添加条件
'参数： query ----查询的SQL
'       Condition ---- 条件所在的数据集
'       dbname ----数据库字段名称
'返回值：无，直接修改query
'-----------------------------------------------------------------------------
    If Condition.Exists Then
        If Condition.value <> "" And Condition.value <> "*" Then
            AddSingleDateCondition query, Condition.DateTimeFrom("1/1/1900") & " " & "00:00:01", ">=", dbname
            AddSingleDateCondition query, Condition.DateTimeTo("1/1/2900") & " " & "23:59:59", "<=", dbname
        End If
    End If
End Sub

Public Sub AddDateTimeCondition(ByRef query As String, ConditionDate As DicomAttribute, ConditionTime As DicomAttribute, dbname As String)
'-----------------------------------------------------------------------------
'功能:添加日期时间查询条件,不为*才添加条件，只有时间没有日期的不添加
'参数： query ----查询的SQL
'       ConditionDate ---- 日期条件所在的数据集
'       ConditionTime ---- 时间条件所在的数据集
'       dbname ----数据库字段名称
'返回值：无，直接修改query
'-----------------------------------------------------------------------------
    '
    If ConditionDate.Exists And ConditionTime.Exists Then
        If ConditionTime.value <> "" And ConditionTime.value <> "*" Then
            If ConditionDate.value <> "" And ConditionDate.value <> "*" Then
                '日期时间条件都有
                AddSingleDateCondition query, ConditionDate.DateTimeFrom("1/1/1900") & " " & ConditionTime.DateTimeFrom("0:0:1"), ">=", dbname
                AddSingleDateCondition query, ConditionDate.DateTimeTo("1/1/2900") & " " & ConditionTime.DateTimeTo("23:59:59"), "<=", dbname
            Else
                '有时间，没有日期条件，查询不了，不处理
            End If
        Else
            '时间条件为空，则只增加日期条件
            AddDateCondition query, ConditionDate, dbname
        End If
    End If
End Sub

Public Sub AddIDCondition(ByRef query As String, Condition As DicomAttribute, dbID As String, dbSendNum As String, Optional ByVal blnAndConnect As Boolean = True)
'添加ID查询条件,不为*才添加条件
'参数说明：
'           query---查询的条件串
'           Condition ----请求的查询条件DicomAttribute，如果值中格式是“数据1_数据2”,则使用dbID，dbSendNum两个条件组织语句
'           dbID ----数据库中的ID字段名
'           dbSendNum ----数据库中的发送号字段名
'           blnAndConnect ----是And 或者Or条件，True---And条件，False---Or条件

    Dim strAdviceID As String, strSendNum As String
    Dim strID As String
    If Condition.Exists And Not IsNull(Condition.value) Then
        strID = Condition.value
        If strID <> "*" Then        '不为*才添加条件
            strAdviceID = Split(strID, "_")(0)
            AddStringCondition query, Val(strAdviceID), dbID, blnAndConnect
            If InStr(strID, "_") > 0 And Len(Trim(dbSendNum)) > 0 Then
                strSendNum = Split(strID, "_")(1)
                AddStringCondition query, Val(strSendNum), dbSendNum, blnAndConnect
            End If
        End If
    End If
End Sub

Public Sub AddCondition(ByRef query As String, Condition As DicomAttribute, dbname As String, Optional ByVal blnAndConnect As Boolean = True)
'添加综合查询条件,不为*才添加条件

    Dim values As Variant
    Dim i As Integer
    
    '判断条件是否存在且不为空
    If Condition.Exists And Not IsNull(Condition.value) Then
        If Condition.Multiple Then
            query = query & IIf(blnAndConnect = True, " AND ", " OR ") & " (FALSE "
            values = Condition.value
            For i = 1 To UBound(values, 1)
                If values(i) <> "*" Then
                    query = query & "OR " & dbname & "='" & values(i) & "'"
                End If
            Next
            query = query & ")"
        Else
            AddStringCondition query, Condition.value, dbname, blnAndConnect
        End If
    End If
End Sub

Public Sub AddStringCondition(ByRef query As String, Condition As String, dbname As String, Optional ByVal blnAndConnect As Boolean = True)
'添加字符串查询条件,不为*才添加条件
    If Condition <> "" And Condition <> "*" And Condition <> "*=*" Then
        If InStr(Condition, "*") Then
            query = query & IIf(blnAndConnect, " AND (", " OR (") & dbname & " like '" & StarToPercent(Condition) & "')"
        Else
            query = query & IIf(blnAndConnect, " AND (", " OR (") & dbname & "= '" & Condition & "')"
        End If
    End If
End Sub

Private Function StarToPercent(s As String) As String
    Dim z As Integer
    While InStr(s, "*")
       z = InStr(s, "*")
       s = Left(s, z - 1) & "%" & Mid(s, z + 1)
    Wend
    StarToPercent = s
End Function

Public Function NewResultItem(request As DicomDataSet) As DicomDataSet
    Dim d As DicomDataSet, a As DicomAttribute
    Set d = New DicomDataSet
    For Each a In request.Attributes
        d.Attributes.Add a.group, a.element, a.value
    Next
    Set NewResultItem = d
End Function

Public Sub AddCountItem(DataSet As DicomDataSet, request As DicomDataSet, group As Long, element As Long, _
                SourceName As String, SourceValue As String, TargetName As String)
'-----------------------------------------------------------------------------
'功能:  根据传入的请求，查询对应级别的序列数量、或者图像数量，在Query/Retrieve中使用，
'       这种查询的速度很慢，尽可能不使用,现在只使用了查询图像数量的部分
'参数： DataSet ----返回的数据集
'       request ----要查找的数据请求
'       group ----要查找的请求的组号
'       element ----要查找的请求的元素号
'       SourceName ----查找的源级别，包括：PATIENTID，StudyUID，SERIESUID，其实就是数据值所对应的数据项
'       SourceValue ----查找的数据值
'       TargetName ----要返回的数据级别，包括：STUDYUID，SERIESUID，INSTANCEUID
'返回值：无，直接往DataSet填写返回的内容
'-----------------------------------------------------------------------------
    Dim rsTemp As Recordset
    Dim strSQL As String
    
    '如果请求中没有这个项目，则不进行查询，直接退出
    If Not request.Attributes(group, element).Exists Then Exit Sub
    
    If UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "STUDYUID" Then
        strSQL = "select count(*) as count from " _
                & "(select c.姓名 from 影像检查记录 c , " _
                & "(select a.病人id,b.医嘱id,b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " _
                & "where a.病人id=[1] AND A.相关ID IS NULL and a.id=b.医嘱id) d " _
                & "where c.医嘱id = d.医嘱id and c.发送号 = d.发送号)"
    ElseIf UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "SERIESUID" Then
        strSQL = "select count(*) as count from " _
                & "(select e.序列uid from 影像检查记录 c , 影像检查序列 e , " _
                & "(select a.病人id,b.医嘱id,b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " _
                & "where a.病人id=[1] AND A.相关ID IS NULL and a.id=b.医嘱id) d " _
                & "where c.医嘱id = d.医嘱id and c.发送号 = d.发送号 and c.检查uid = e.检查uid)"
    ElseIf UCase(SourceName) = "PATIENTID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = " select count(*) as count from " _
                & "(select f.图像uid from 影像检查记录 c , 影像检查序列 e , 影像检查图象 f , " _
                & "(select a.病人id,b.医嘱id,b.发送号 from 病人医嘱记录 a,病人医嘱发送 b " _
                & "where a.病人id=[1] AND A.相关ID IS NULL and a.id=b.医嘱id) d " _
                & "Where c.医嘱id = d.医嘱id And c.发送号 = d.发送号 " _
                & "and c.检查uid = e.检查uid and e.序列uid = f.序列uid) "
    ElseIf UCase(SourceName) = "STUDYUID" And UCase(TargetName) = "SERIESUID" Then
        strSQL = " select count(*) as count from " _
                & "(select b.序列uid from 影像检查记录 a , 影像检查序列 b " _
                & "where a.检查uid = [1] and a.检查uid = b.检查uid) "
    ElseIf UCase(SourceName) = "STUDYUID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = " select count(*) as count from " _
                & "(select d.图像uid from 影像检查图象 d , " _
                & "(select b.序列uid from 影像检查记录 a , 影像检查序列 b " _
                & "where a.检查uid =[1] and a.检查uid = b.检查uid) c " _
                & "where d.序列uid = c.序列uid)"
    ElseIf UCase(SourceName) = "SERIESUID" And UCase(TargetName) = "INSTANCEUID" Then
        strSQL = "select count(*) as count from " _
                & "(select b.图像uid from 影像检查序列 a , 影像检查图象 b " _
                & "where a.序列uid = [1] and a.序列uid = b.序列uid)"
    End If
    If UCase(SourceName) = "PATIENTID" Then
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询返回记录的数量", CLng(SourceValue))
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询返回记录的数量", SourceValue)
    End If
    DataSet.Attributes.Add group, element, rsTemp!Count
End Sub

Private Function funcGetSeriesUID(strOldSeriesUID As String, strImageType As String) As String
'-----------------------------------------------------------------------------
'功能:根据现有序列UID查询，返回用影像类型拆分后的新序列UID
'修改人:黄捷
'修改日期:2007-4-18
'-----------------------------------------------------------------------------
    Dim rsMatch As New ADODB.Recordset
    Dim intMax As Integer
    Dim intCur As Integer
    Dim blnMatch As Boolean
    
    funcGetSeriesUID = strOldSeriesUID
    gstrSQL = "select 0 as 临时,序列UID,序列描述 from 影像检查序列 where 序列UID like  [1]" & _
              " Union All Select 1 as 临时,序列UID,序列描述 from 影像临时序列 where 序列UID like [1]"
    Set rsMatch = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strOldSeriesUID & "%")
    
    While Not rsMatch.EOF
        If rsMatch("序列UID") = strOldSeriesUID Then
            intCur = 0
        Else
            intCur = Val(Right(rsMatch("序列UID"), Len(rsMatch("序列UID")) - InStrRev(rsMatch("序列UID"), ".")))
        End If
        
        If intMax < intCur Then intMax = intCur
        If rsMatch("序列描述") = strImageType Then
            funcGetSeriesUID = rsMatch("序列UID")
            blnMatch = True
            rsMatch.MoveLast
        End If
        rsMatch.MoveNext
    Wend
    
    If blnMatch = False Then
        '创建新的UID
        funcGetSeriesUID = strOldSeriesUID & "." & intMax + 1
    End If
End Function


Public Sub SaveImages(Images As DicomImages, ByVal BufferDir As String)
'功能：保存图像
    Dim curImage As DicomImage
    Dim i As Integer, iCount As Integer     '保存的图像数
    Dim rsTmp As New ADODB.Recordset
    Dim blnTmp As Boolean                   '是否被保存成临时记录
    Dim dtReceived As String
    
    Dim PatientName As String, EnglishName As String, Sex As String, Age As Integer
    
    Dim strBirth As String
    Dim strAdviceID As String   '图像中的匹配ID，ID可能超长，因此用String，可能是PatientID，PatientName，AccessionNumber，统称为医嘱ID
    
    Dim lngSeriesNo As Long
    Dim lngImageNo As Long
    Dim strStudyDateTime As String  '存储到数据库中的“接收时间”，图象中的检查日期和时间
    Dim strStudyUID As String       '存储本次保存图像时使用的检查UID
    Dim strSeriesUID As String      '存储本次保存图像时使用的序列UID
    
    Dim strSeriesDesp As String     '序列描述
    Dim strSQLbak As String
    '服务参数设置
    Dim blnSplitSeriesUID As Boolean    '根据图像类型拆分序列UID
    Dim intImageMatchItem As Integer    '图像匹配项 ，0--PatientID，1--AccessionNumber，2--PatientName
    Dim intDBMatchItem As Integer       '数据库匹配项， 0--检查号匹配;1--病人标识匹配;2--检查标识匹配
    Dim blnMatchStudyUID As Boolean     '启用检查UID匹配
    Dim strStoreDeviceNo As String      '存储设备号
    Dim intEncode As Integer            '压缩方式
    Dim strOldStoreDeviceNo As String   '保存上一个图像的FTP设备号
    Dim strAutoRoute As String          '保存自动路由目的地集合，使用“|”分隔各个存储设备号
    Dim strAutoRouteCompression As String '保存自动路由的压缩方法集合，使用“|”分隔各个压缩方式，0--按照当前方式压缩，1--不压缩
    Dim strAutoRouteDir As String       '保存自动路由的目录结构集合，使用“|”分隔各个目录结构，0--检查级别目录（默认），1--序列级别目录（3D）
    Dim intFilterModality As Integer    '过滤方式 0--按影像类别过滤，1--按IP地址过滤
    Dim strPhysician As String          '提取检查技师
    Dim strPhysicianName As String      '图像中检查技师字段的内容
    'FTP存储参数
    Dim strFTPDir As String
    '临时使用的FTP存储参数
    Dim strNewDeviceID As String
        
    'AE连接参数
    Dim strServiceAE As String
    Dim strDeviceIP As String
    
    Dim lngResult As Long           '保存FTP操作返回的错误
    Dim blnNewStudy As Boolean      '记录是否新的检查
    
    Dim blnInDBTrans As Boolean     '记录是否在数据库事务之中
    Dim arrSQL() As Variant         '记录需要执行的存储过程的数组
    Dim strModality As String       '记录图像的影像类别
    Dim str检查设备 As String       '记录图像中的检查设备，如果匹配成功，则是数据库中的检查设备字段的内容
    
    On Error GoTo DBError
    
    iCount = 0
    For Each curImage In Images
        '先检查这个图像是否已经存在数据库中了
        gstrSQL = "Select 图像UID From 影像检查图象 Where 图像UID= [1] " & _
            " Union All Select 图像UID From 影像临时图象 Where 图像UID= [1] "
        strSQLbak = gstrSQL
        strSQLbak = Replace(strSQLbak, "影像检查图象", "H影像检查图象")
        gstrSQL = gstrSQL & " Union ALL " & strSQLbak
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS保存图像", curImage.instanceUID)
        
        Call WriteProcessLog("SaveImages", "检查图像是否已经存在数据库中", "检查结果是：数据库中" & IIf(rsTmp.EOF = True, "没有", "有") & "这个图像。")
        
        '是新图像，则保存图像,否则处理下一个图像
        If rsTmp.EOF Then
            '记录原来的存储设备号，并且读取当前图像对应的存储参数
            strOldStoreDeviceNo = strStoreDeviceNo
            strModality = GetImageAttribute(curImage.Attributes, ATTR_影像类别)
            str检查设备 = GetImageAttribute(curImage.Attributes, ATTR_检查设备)
            
            Call WriteProcessLog("SaveImages", "读取当前图像的存储基本参数", "影像类别 = " & strModality & "，检查设备 = " & str检查设备)
            
            '读取当前图像的存储基本参数和自动匹配参数
            If funGetAEStoreParas(curImage.Tag, strModality, strDeviceIP, blnSplitSeriesUID, intImageMatchItem, intDBMatchItem, blnMatchStudyUID, _
                strStoreDeviceNo, intEncode, strAutoRoute, intFilterModality, strAutoRouteCompression, strAutoRouteDir, strPhysician) = True Then
                
                '确定是新图像，而且有对应接收存储参数，准备保存图像，首先处理图像文件，然后保存成FTP文件
                '处理影像属性
                DeleteImageAttribute curImage.Attributes, ATTR_长宽比 '删除该属性
                '读取图像信息,优先读取采集时间和日期，此日期不存在，再读取检查日期
                '原因是柯达的设备，通过Worklist提取信息时，会把“报到时间”填写到检查日期，
                '如果提前报到，会导致检查日期跟采集日期相差几天，影响后续时间查询
                dtReceived = Format(GetImageAttribute(curImage.Attributes, ATTR_采集日期), "yyyyMMdd")  '根据图像中的采集日期给dtReceived赋初值
                strStudyDateTime = Format(GetImageAttribute(curImage.Attributes, ATTR_采集日期), "yyyy-MM-dd") & _
                    " " & Format(GetImageAttribute(curImage.Attributes, ATTR_采集时间), "HH:MM")
                If dtReceived = "" Then
                    '读取检查时间和日期
                    dtReceived = Format(GetImageAttribute(curImage.Attributes, ATTR_检查日期), "yyyyMMdd")  '根据图像中的检查日期给dtReceived赋初值
                    strStudyDateTime = Format(GetImageAttribute(curImage.Attributes, ATTR_检查日期), "yyyy-MM-dd") & _
                        " " & Format(GetImageAttribute(curImage.Attributes, ATTR_检查时间), "HH:MM")
                End If
                
                strStudyUID = curImage.StudyUID             '根据图像内的检查UID给strStudyUID赋初值
                PatientName = Replace(curImage.Name, "'", "’")
                EnglishName = Replace(curImage.Name, "'", "’")
                Sex = Replace(curImage.Sex, "'", "’")
                
                '如果是多帧图像，则创建新的序列UID
                strSeriesUID = curImage.SeriesUID
                If curImage.FrameCount > 1 Then
                    strSeriesUID = funcGetSeriesUID(strSeriesUID, "MultiFrame")
                End If
                strSeriesDesp = Replace(curImage.SeriesDescription, "'", "＇")
                '提取图像中的主匹配ID
                strAdviceID = funGetMatchIDInImg(curImage, intImageMatchItem, intDBMatchItem)
                '根据图像类型拆分序列UID
                If blnSplitSeriesUID = True Then
                    If funSplitSeriesUID(curImage, strSeriesUID, strSeriesDesp) <> 0 Then
                        err.Raise vbObjectError + 4, "funSplitSeriesUID", "根据类型拆分序列UID错误,出现错误的图像是：" & curImage.Name
                    End If
                End If
                
                '判断当前图像存储设备号是否改变，如果改变，则重新提取FTP存储设备参数并重新连接FTP
                If strStoreDeviceNo <> strOldStoreDeviceNo Then
                    '重新连接FTP,lngResult=0成功；lngResult=1连接失败；lngResult=2无法获取用户名和密码
                    lngResult = funReConnectFTP(strStoreDeviceNo, iNet, strFTPDir, 1)
                    If lngResult = 1 Then
                        err.Raise vbObjectError + 1, "PACS图像保存", "FTP 连接失败！"
                    ElseIf lngResult = 2 Then
                        err.Raise vbObjectError + 1, "PACS图像保存", "FTP 无法获取FTP目录的用户名和密码！"
                    End If
                End If
                
                '查询是否有已经匹配成功的记录
                lngResult = funIsPreMatched(blnMatchStudyUID, intDBMatchItem, strStudyUID, strAdviceID, strDeviceIP, _
                                 strSeriesUID, strModality, dtReceived, intFilterModality, strNewDeviceID, strStoreDeviceNo, _
                                 blnTmp, str检查设备, PatientName, EnglishName, Age, Sex, strStudyDateTime)
                If lngResult = 0 Then   '匹配成功
                    blnNewStudy = False '匹配成功，则不是新的检查
                    '如果设备号改变，则重新连接FTP
                    If strNewDeviceID <> strStoreDeviceNo Then
                        strStoreDeviceNo = strNewDeviceID
                        lngResult = funReConnectFTP(strStoreDeviceNo, iNet, strFTPDir, 2)
                        If lngResult = 1 Then
                            err.Raise vbObjectError + 1, "PACS图像保存", "FTP 连接失败！"
                        ElseIf lngResult = 2 Then
                            err.Raise vbObjectError + 1, "PACS图像保存", "FTP 无法获取FTP目录的用户名和密码！"
                        End If
                    End If
                Else    '匹配不成功
                    If blnMatchStudyUID = False Then  '查询检查UID是否重复，若重复则创建新的检查UID
                        strStudyUID = funGetStudyUID(strStudyUID)
                    End If
                    blnNewStudy = True
                End If
                
                '保存FTP图像文件到缓存目录
                lngResult = funUploadImage(curImage, iNet, intEncode, BufferDir, strFTPDir, strStudyUID, dtReceived)
                If lngResult = 1 Then
                    err.Raise vbObjectError + 2, "PACS图像保存", "FTP 第" & Val(curImage.BorderWidth) & "次存储失败！" _
                        & " 病人姓名：" & curImage.Name & " 图像UID ： " & curImage.instanceUID _
                        & " 检查设备： " & str检查设备
                ElseIf lngResult = 2 Then
                    err.Raise vbObjectError + 3, "PACS图像保存", "图像被放弃，FTP 第" & Val(curImage.BorderWidth) & "次存储失败！" _
                        & " 病人姓名：" & curImage.Name & " 图像UID ： " & curImage.instanceUID _
                        & " 检查设备： " & str检查设备 & " 图像副本保存为：" & BufferDir & "图像UID.bak"
                ElseIf lngResult = 3 Then
                    err.Raise vbObjectError, "上传错误", "funUploadImage 上传图像出现错误。"
                End If
                
                '准备开始组织保存图像的存储过程数组
                arrSQL = Array()
                
                '如果没有预先匹配成功的记录，则说明这个图像是某个检查的第一个图像，查找这个检查并且做匹配
                '如果查找不到这个检查，则说明匹配不成功，图像会被保存成临时检查中一个记录
                If blnNewStudy = True Then      '没有已经匹配成功的记录，则按病人ID或英文名查找
                    Select Case intDBMatchItem
                        Case 0 '检查号匹配
                            gstrSQL = "Select Distinct A.姓名,A.英文名,A.性别,A.年龄,A.检查设备,A.医嘱ID,A.发送号,B.首次时间,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.首次时间) as tInterval,b.执行过程 " & _
                                " From 影像检查记录 A,病人医嘱发送 B,影像设备目录 C " & _
                                " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And A.检查设备 =C.设备号 And b.执行状态=3 And b.执行过程>=2 " & _
                                " And " & IIf(intFilterModality = 0, " UPPER(C.影像类别)=[3] ", " C.IP地址=[2] ") & " And A.检查号= [1] And A.检查UID Is Null Order By tInterval"
                        Case 1 '病人标识匹配
                            gstrSQL = "Select Distinct A.姓名,A.英文名,A.性别,A.年龄,A.检查设备,A.医嘱ID,A.发送号,B.首次时间,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.首次时间) as tInterval,b.执行过程 " & _
                                " From 影像检查记录 A,病人医嘱发送 B,病人医嘱记录 C,病人信息 D,影像设备目录 E " & _
                                " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And A.检查设备 =E.设备号 And C.相关ID IS NULL And A.医嘱ID=C.ID And C.病人ID=D.病人ID" & _
                                " And " & IIf(intFilterModality = 0, " UPPER(E.影像类别)=[3] ", " E.IP地址=[2] ") & " And b.执行状态=3 And b.执行过程>=2 " & _
                                " And ((D.住院号=[1] AND C.病人来源=2) OR (D.门诊号= [1] AND C.病人来源<>2)) And A.检查UID Is Null Order By tInterval"
                        Case 2 '检查标识匹配
                            gstrSQL = "Select Distinct A.姓名,A.英文名,A.性别,A.年龄,A.检查设备,A.医嘱ID,A.发送号,B.首次时间,abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.首次时间) as tInterval,b.执行过程 " & _
                                " From 影像检查记录 A,病人医嘱发送 B,病人医嘱记录 C" & _
                                " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID And C.相关ID IS NULL And b.执行状态=3 And b.执行过程>=2 " & _
                                " And A.医嘱ID= [1] And A.检查UID Is Null Order By tInterval"
                    End Select
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strAdviceID, strDeviceIP, UCase(strModality))
                        
                    '查找到有匹配的记录，则与HIS填写的检查记录对应
                    If rsTmp.EOF = False Then
                        '记录当前的检查设备
                        str检查设备 = NVL(rsTmp("检查设备"))
                        PatientName = NVL(rsTmp("姓名"))
                        EnglishName = NVL(rsTmp("英文名"))
                        Age = Val(NVL(rsTmp("年龄"), 0))
                        Sex = NVL(rsTmp("性别"))
                        
                        '设置匹配记录
                        gstrSQL = "ZL_影像检查记录_SET(" & rsTmp("医嘱ID") & "," & rsTmp("发送号") & ",'" & _
                            strStudyUID & "',null," & _
                            "to_Date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strStoreDeviceNo & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = gstrSQL
                        
                        '设为执行完成
                        '先判断当前"病人医嘱发送"中的"执行过程"是否小于3,如果是,才需要修改执行过程
                        If rsTmp!执行过程 < 3 Then
                            gstrSQL = "ZL_影像检查_STATE(" & rsTmp("医嘱ID") & "," & rsTmp("发送号") & ",3)"
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = gstrSQL
                        End If
                        
                        '记录检查技师
                        If strPhysician <> "" Then
                            strPhysicianName = GetImageAttribute(curImage.Attributes, strPhysician)
                            If strPhysicianName <> "" Then
                                gstrSQL = "Zl_影像技师执行('" & strPhysicianName & "'," & rsTmp("医嘱ID") & ",2)"
                                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                                arrSQL(UBound(arrSQL)) = gstrSQL
                            End If
                        End If
                        blnTmp = False
                    Else        '没有找到匹配的记录，则插入临时检查记录
                        '计算和填充年龄
                        If IsDate(curImage.DateOfBirthAsDate) Then
                            If curImage.DateOfBirthAsDate <> "0:00:00" Then
                                strBirth = Format(curImage.DateOfBirthAsDate, "YYYY-MM-DD")
                            Else
                                strBirth = ""
                            End If
                            
                            If curImage.Attributes(&H10, &H1010).Exists And Not IsNull(curImage.Attributes(&H10, &H1010)) Then
                                Age = Val(curImage.Attributes(&H10, &H1010).value)
                            Else
                                If strBirth = "" Then
                                    Age = 0
                                Else
                                    Age = CStr(Year(Date) - Year(strBirth))
                                End If
                            End If
                        Else
                            Age = 0: strBirth = ""
                        End If
                        '填充其他必要字段
                        PatientName = Replace(curImage.Name, "'", "’")
                        EnglishName = Replace(curImage.Name, "'", "’")
                        Sex = Replace(curImage.Sex, "'", "’")
                        
                        '如果用来匹配的图像中的号码为-1 ， 则从图像的PatientID中重新读取一遍号码。
                        If strAdviceID = "-1" Then
                            '如果是按照检查号匹配，直接提取patientID，是病人标识或医嘱ID才提取数值
                            strAdviceID = IIf(intDBMatchItem = 0, curImage.PatientID, getStrNumber(curImage.PatientID))
                        End If
                        gstrSQL = "ZL_影像临时检查_INSERT('" & strModality & "','" & strAdviceID & "','" & _
                            PatientName & "','" & EnglishName & "','" & Sex & "','" & Age & "'," & _
                            IIf(Len(strBirth) = 0, "Null", "to_Date('" & strBirth & "','YYYY-MM-DD')") & ",Null,Null,'" & _
                            str检查设备 & "','" & strStudyUID & "'," & _
                            "to_Date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS'),'" & strStoreDeviceNo & "')"
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = gstrSQL
                        blnTmp = True
                    End If
                End If
                
                '判断是否需要插入新的序列
                gstrSQL = "Select 序列UID From " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                    " Where 序列UID= [1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strSeriesUID)
                
                If rsTmp.EOF Then
                    '插入新的检查序列
                    lngSeriesNo = IIf(GetImageAttribute(curImage.Attributes, ATTR_序列号) = "", -1, GetImageAttribute(curImage.Attributes, ATTR_序列号))
                    If lngSeriesNo <> -1 Then
                        gstrSQL = "select 序列号 from " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                            " where 检查UID=[1] AND 序列号 =[2]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strStudyUID, lngSeriesNo)
                        
                        If Not rsTmp.EOF Then
                            gstrSQL = "select max(序列号) from " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                                " where 检查UID=[1] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strStudyUID)
                            If Not rsTmp.EOF Then lngSeriesNo = NVL(rsTmp(0), 0) + 1
                        End If
                    Else
                        gstrSQL = "select max(序列号) from " & IIf(blnTmp, "影像临时序列", "影像检查序列") & _
                            " where 检查UID=[1] "
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strStudyUID)
                        If rsTmp.EOF = False Then
                            lngSeriesNo = NVL(rsTmp(0), 0) + 1
                        Else
                            lngSeriesNo = 1
                        End If
                    End If
                    '插入新的序列
                    gstrSQL = "ZL_影像序列_INSERT('" & strStudyUID & "','" & strSeriesUID & "','" & _
                        strSeriesDesp & "'," & IIf(blnTmp, 1, 0) & "," & lngSeriesNo & ")"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                
                '处理可能重复的图像号
                lngImageNo = IIf(GetImageAttribute(curImage.Attributes, ATTR_图像号) = "", -1, GetImageAttribute(curImage.Attributes, ATTR_图像号))
                If lngImageNo <> -1 Then
                    gstrSQL = "select 图像号 from " & IIf(blnTmp, "影像临时图象", "影像检查图象") & _
                        " where 序列UID = [1] and 图像号 = [2]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strSeriesUID, lngImageNo)
                    
                    If rsTmp.EOF = False Then
                        gstrSQL = "select max(图像号) from " & IIf(blnTmp, "影像临时图象", "影像检查图象") & _
                            " where 序列UID=[1]"
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strSeriesUID)
                        lngImageNo = NVL(rsTmp(0), 0) + 1
                    End If
                Else
                    gstrSQL = "select max(图像号) from " & IIf(blnTmp, "影像临时图象", "影像检查图象") & _
                        " where 序列UID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "PACS图像保存", strSeriesUID)
                    If rsTmp.EOF = False Then
                        lngImageNo = NVL(rsTmp(0), 0) + 1
                    Else
                        lngImageNo = 1
                    End If
                End If
                '插入新的图像
                gstrSQL = "ZL_影像图象_INSERT('" & curImage.instanceUID & "','" & strSeriesUID & "','" _
                    & strSeriesDesp & "'," & IIf(blnTmp, 1, 0) & "," & lngImageNo & "," _
                    & "to_Date('" & Format(GetDateAttribute(curImage.Attributes, ATTR_采集日期, 1) & " " & GetDateAttribute(curImage.Attributes, ATTR_采集时间, 2), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')," _
                    & "to_Date('" & Format(GetDateAttribute(curImage.Attributes, ATTR_图像日期, 1) & " " & GetDateAttribute(curImage.Attributes, ATTR_图像时间, 2), "yyyy-MM-dd HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS'),'" _
                    & GetImageAttribute(curImage.Attributes, ATTR_层厚) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_图像位置病人) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_图像方向病人) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_参考帧UID) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_切片位置) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_行数) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_列数) & "','" _
                    & GetImageAttribute(curImage.Attributes, ATTR_像素距离) & "'," _
                    & IIf(curImage.FrameCount = 1, 0, 1) & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                
                '把B超图像保存成报告图像
                If UCase(strModality) = "US" Then
                    gstrSQL = "ZL_影像检查报告_ADD('" & strStudyUID & "','" & curImage.instanceUID & ".jpg')"
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
                
                '启动数据库事务来保存图像
                gcnOracle.BeginTrans
                blnInDBTrans = True
                For i = 0 To UBound(arrSQL)
                    Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存图像")
                Next i
                gcnOracle.CommitTrans
                blnInDBTrans = False
                
                
                '保存本地日志，保存影像接收序列
                WriteRecord strModality, strAdviceID, str检查设备, PatientName, EnglishName, Sex, Age, strStudyUID, strSeriesUID, blnTmp
                
                '自动路由
                '--------------------------还没有处理
                Call funcAutoRouting(curImage, BufferDir, dtReceived, strStudyUID, intEncode, strAutoRoute, strAutoRouteCompression, strAutoRouteDir)
            Else        'funGetAEStoreParas的结束
                '读取不到存储的参数，则记录错误日志,并处理下一个图像
                '匹配参数未知，不是系统允许的服务所发送的图像，不保存
                Call GetAEconnection(Val(curImage.Tag), strServiceAE, strDeviceIP)
                WriteLog 3, vbObjectError + 1, "从 IP= " & strDeviceIP & " 发送给 AE= " & strServiceAE & " 的图像，已经接收到，但是本次连接不是系统所允许的服务对，图像无法保存。"
                If strDeviceIP = "" Or strServiceAE = "" Then
                    '查找服务AE和IP
                    For i = 1 To UBound(AEconnections)
                        WriteLog 200, 201, " Association = " & Val(curImage.Tag) & " i = " & i & " UBound(AEconnections) = " & UBound(AEconnections) & vbCrLf _
                            & " AEconnections(i).Association = " & AEconnections(i).Association & " AEconnections(i).ServiceAE = " & AEconnections(i).ServiceAE & vbCrLf _
                            & " AEconnections(i).DeviceIP = " & AEconnections(i).DeviceIP & " AEconnections(i).TimeStamp  = " & AEconnections(i).TimeStamp
                    Next i
                End If
            End If
        Else    'end of 检查图像是否在数据库中
            '图像已经保存在数据库中的某个表，不是新图像，则记录错误日志，并处理下一个图像
            WriteLog 3, vbObjectError + 1, "影像：" & curImage.instanceUID & "已存在！"
        End If
        iCount = iCount + 1
        If iCount >= 20 Then Exit For
    Next
    
    For i = 1 To iCount
        Images.Remove 1
    Next
    iNet.FuncFtpDisConnect
    Exit Sub
    
DBError:
    
    If blnInDBTrans = True Then
        gcnOracle.RollbackTrans
    End If
    
    Dim lngerrNumber As Long
    
    lngerrNumber = err.Number   '先记录错误号，后续的写日志等函数，内部又启用了错误处理，会改变错误号
    
    '先记录错误日志，再处理其他
    Call WriteLog(4, err.Number, "保存图像时出现错误，错误描述为：" & err.Description)
    
    '处理特定错误
    If lngerrNumber = vbObjectError + 2 Then  '第X次上传失败
        For i = 1 To iCount
            Images.Remove 1
        Next
    ElseIf lngerrNumber = vbObjectError + 3 Or lngerrNumber = vbObjectError + 4 Then  '3上传失败次数达到极限，4是funSplitSeriesUID， 放弃图像，并将图像保存到本地
        For i = 1 To iCount + 1
            If i = iCount + 1 Then
                '这是一个将要被放弃的图像，保存到本机缓存中
                Images(1).WriteFile BufferDir & Images(1).instanceUID & ".bak", True
            End If
            Images.Remove 1
        Next
    End If
    
    iNet.FuncFtpDisConnect
End Sub

Public Sub subSaveAssociation(connection As DicomConnection)
    Dim lngCount  As Long

    '增加连接数组
    ReDim Preserve AEconnections(UBound(AEconnections) + 1) As AEconnection
    lngCount = UBound(AEconnections)

    AEconnections(lngCount).ServiceAE = connection.CalledAET
    AEconnections(lngCount).Association = connection.Association
    AEconnections(lngCount).DeviceIP = connection.RemoteIP
    AEconnections(lngCount).TimeStamp = Now
    AEconnections(lngCount).Deleted = False
End Sub

Public Function GetDateAttribute(objAttr As DicomAttributes, ByVal AttrName As String, iType As Integer) As String
'-----------------------------------------------------------------------------
'功能:提取日期类型的属性值，如果出现空值，则自动使用当前日期
'参数： objAttr ----属性集合
'       AttrName ----要查找的属性名称
'       iType ----类型 1--日期；2--时间
'返回值：属性的内容
'-----------------------------------------------------------------------------
    Dim strDateValue As String
    
    strDateValue = GetImageAttribute(objAttr, AttrName)
    
    If iType = 1 Then   '日期
        If strDateValue = "" Or IsDate(strDateValue) = False Then
            strDateValue = Date
        End If
        strDateValue = Format(strDateValue, "yyyy-mm-dd")
    Else    '时间
        If strDateValue = "" Or IsDate(strDateValue) = False Then
            strDateValue = Time
        End If
        strDateValue = Format(strDateValue, "hh:mm:ss")
    End If
    
    GetDateAttribute = strDateValue
End Function

Private Function funReConnectFTP(strStoreDeviceNo As String, ByRef ftpNet As clsFtp, strFTPDir As String, intType As Integer) As Long
'-----------------------------------------------------------------------------
'功能:根据输入的参数，重新连接FTP
'参数： strStoreDeviceNo ----FTP连接的设备号
'       ftpNet ---- FTP连接
'       strFTPDir ----返回的FTP目录
'       intType ----读取连接参数的方法 1--从FTPDevices数组中读取；2--从数据库中查询
'返回值：0--成功；1--连接失败；2--获取用户名和密码失败
'-----------------------------------------------------------------------------
    Dim strIP As String
    Dim strUser As String
    Dim strPassWord As String
    Dim blnRet As Boolean
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngResult As Long
    
    On Error GoTo err
    
    '读取当前图像的存储设备
    If intType = 1 Then     '从FTPDevices数组中读取
        blnRet = funGetFTPDevice(strStoreDeviceNo, strIP, strUser, strPassWord, strFTPDir)
    Else        '从数据库中查询
        strSQL = "select IP地址,FTP目录,FTP用户名,FTP密码 from 影像设备目录  Where 设备号  = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取存储设备", strStoreDeviceNo)
        If rsTemp.RecordCount = 1 Then
            strIP = NVL(rsTemp("IP地址"))
            strUser = NVL(rsTemp("FTP用户名"))
            strPassWord = NVL(rsTemp("FTP密码"))
            strFTPDir = NVL(rsTemp("FTP目录")) & "/"
            blnRet = True
        End If
    End If
    
    '重新连接FTP
    If blnRet = True Then
        ftpNet.FuncFtpDisConnect
        lngResult = ftpNet.FuncFtpConnect(strIP, strUser, strPassWord)
        If lngResult = 0 Then
            'FTP连接错误
            funReConnectFTP = 1
            WriteLog 300, vbObjectError + 1, "funReConnectFTP FTP连接错误，该图像无法保存，设备号 = " & strStoreDeviceNo
        End If
    Else
        '根据设备号，无法获取FTP目录的用户名和密码
        funReConnectFTP = 2
        WriteLog 301, vbObjectError + 1, "funReConnectFTP 无法获取FTP目录的用户名和密码，该图像无法保存，设备号 = " & strStoreDeviceNo
    End If

    Exit Function
err:
    Call WriteLog(302, err.Number, "funReConnectFTP: " & err.Description)
    ftpNet.FuncFtpDisConnect
End Function

Private Function funSplitSeriesUID(ByRef img As DicomImage, ByRef strSeriesUID As String, ByRef strSeriesDesp As String) As Long
'-----------------------------------------------------------------------------
'功能:根据图像类型拆分序列UID
'参数： img ----需要拆分的图像
'       strSeriesUID ---- 返回的序列UID
'       strSeriesDesp ----返回的序列描述
'返回值：0--成功；1--失败
'-----------------------------------------------------------------------------
    Dim strImageType As String      '图像类别，LOCALIZER,AXIAL
    Dim vImageType() As String      '图像类别
    
    On Error GoTo err
    
    '读取图像类别
    strImageType = GetImageAttribute(img.Attributes, ATTR_图像类型)
    vImageType = Split(strImageType, "\")
    strImageType = vImageType(3)
    '根据图像类型拆分序列
    strSeriesUID = funcGetSeriesUID(strSeriesUID, strImageType)
    strSeriesDesp = strImageType
    img.SeriesUID = strSeriesUID
    
    Exit Function
err:
    Call WriteLog(1001, err.Number, "funSplitSeriesUID: " & err.Description)
    funSplitSeriesUID = 1
End Function

Private Function funUploadImage(ByRef img As DicomImage, ByRef ftpNet As clsFtp, ByVal intEncode As Integer, _
    ByVal strBufferDir As String, ByVal strFTPDir As String, ByVal strStudyUID As String, ByVal strDtReceived As String) As Long
'-----------------------------------------------------------------------------
'功能:保存图像到FTP中
'参数： img ----需要保存的图像
'       ftpNet ---- FTP连接
'       intEncode ---- 压缩方式
'       strBufferDir ---- 本地缓存路径
'       strFTPDir ---- FTP的存储目录
'       strStudyUID ---- 检查UID
'       strDtReceived --- 接收日期，FTP目录的组成部分
'返回值：0--成功；1--第X次尝试上传失败；2--上传失败次数达到极限，放弃图像；3--其他错误
'-----------------------------------------------------------------------------
    Dim blnNoCompress As Boolean    '记录当前图像是否不需要压缩
    Dim lngResult As Long           '记录返回值

    On Error GoTo err
    
    '首先判断图像是否属于不能压缩的，比如Philips的3D重建效果图就不能压缩，压缩后图像会变成黑白
    blnNoCompress = False
    If Not IsNull(img.Attributes(&H28, &H2)) And img.Attributes(&H28, &H2).Exists _
        And Not IsNull(img.Attributes(&H28, &H4)) And img.Attributes(&H28, &H4).Exists _
        And Not IsNull(img.Attributes(&H28, &H6)) And img.Attributes(&H28, &H6).Exists Then
        
        If img.Attributes(&H28, &H2).value = 3 And img.Attributes(&H28, &H4).value = "RGB" _
            And img.Attributes(&H28, &H6).value = 1 Then
            
            blnNoCompress = True
        End If
    End If
    
    '判断图像是否由PNMS设备产生的CT图，如果是，则增加空白处的像素定义
    If Not IsNull(img.Attributes(&H8, &H70)) And img.Attributes(&H8, &H70).Exists Then
        If UCase(img.Attributes(&H8, &H70)) = "PNMS" Then
            If UCase(img.Attributes(&H8, &H60)) = "CT" Then
                img.Attributes.Add &H28, &H120, -2000
            End If
        End If
    End If
    
    Call WriteProcessLog("funUploadImage", "上传图像到FTP", "图像压缩=" & blnNoCompress & "， 图像压缩方式=" & intEncode)
    
    If blnNoCompress = True Then
        img.WriteFile strBufferDir & img.instanceUID, True
    Else
        Select Case intEncode
            Case 0
                img.WriteFile strBufferDir & img.instanceUID, True, TS_JPEG无损压缩
            Case 1
                img.WriteFile strBufferDir & img.instanceUID, True, TS_RLE行程压缩
            Case 2
                img.WriteFile strBufferDir & img.instanceUID, True, TS_JPEG2000无损压缩
            Case 3
                img.WriteFile strBufferDir & img.instanceUID, True
        End Select
    End If
    
    Call WriteProcessLog("funUploadImage", "保存本地缓存文件", "本地缓存文件名为：" & strBufferDir & img.instanceUID)
    
    '上传FTP图像文件
    lngResult = WriteToURL(ftpNet, strBufferDir & img.instanceUID, strFTPDir & "/" & _
        strDtReceived & "/" & strStudyUID & "/" & img.instanceUID)
    
    '如果上传失败，则进行对应的处理，使用BorderWidth来暂时保存图像被尝试上传的次数
    '尝试上传10次都失败，则放弃保存图像
    If lngResult <> 0 Then
        If NVL(img.BorderWidth, 0) = 0 Then
            img.BorderWidth = 1
        Else
            img.BorderWidth = img.BorderWidth + 1
        End If
        If img.BorderWidth < 10 Then
            funUploadImage = 1
            
            'FTP 第 img.BorderWidth 次存储失败！删除临时图像
            Kill strBufferDir & img.instanceUID
            Exit Function
        Else
            funUploadImage = 2
            
            '图像被放弃，FTP 第 img.BorderWidth 次存储失败！删除临时图像
            Kill strBufferDir & img.instanceUID
            Exit Function
        End If
    End If
    
    '针对通过DICOM方式接收B超图的情况，自动把B超图像保存成报告图像
    If UCase(GetImageAttribute(img.Attributes, ATTR_影像类别)) = "US" Then
        img.FileExport strBufferDir & img.instanceUID & ".jpg", "JPG", 80
        WriteToURL ftpNet, strBufferDir & img.instanceUID & ".jpg", strFTPDir & "/" & _
            strDtReceived & "/" & strStudyUID & "/" & img.instanceUID & ".jpg"
    End If
    
    '删除临时图像
    Kill strBufferDir & img.instanceUID
    Exit Function
err:
    Call WriteLog(1001, err.Number, "funUploadImage: " & err.Description)
    funUploadImage = 3
End Function

Private Function funGetMatchIDInImg(img As DicomImage, intImgMatchItem As Integer, intDatabaseMatchItem As Integer) As String
'-----------------------------------------------------------------------------
'功能:根据条件，提取图像中的匹配ID
'参数： img ----需要匹配的图像
'       intImgMatchItem ---- 匹配的项目，0--PatientID，1--AccessionNumber，2--PatientName
'       intDatabaseMatchItem ---数据库匹配项目， 0--检查号匹配;1--病人标识匹配;2--检查标识匹配
'返回值：匹配ID，如果图像中条件为空，返回-1
'-----------------------------------------------------------------------------
    Dim aPatientID() As String
    Dim strTemp As String
    
    On Error GoTo err

    If intDatabaseMatchItem = 0 Then    '检查号匹配，支持字符型，不需要做分隔，直接提取整个检查号的内容
        Select Case intImgMatchItem
            Case 0 'Patient ID
                funGetMatchIDInImg = NVL(img.PatientID)
            Case 1 'Accession Number
                funGetMatchIDInImg = NVL(img.Attributes(&H8, &H50).value)
            Case 2 'Patient Name
                funGetMatchIDInImg = NVL(img.Name)
        End Select
        
        '去掉匹配号码前面多余的0
        While Left(funGetMatchIDInImg, 1) = "0"
            funGetMatchIDInImg = Mid(funGetMatchIDInImg, 2)
        Wend
        
        Exit Function
    End If
    
    '病人ID，医嘱ID匹配，需要从图像中提取数值型的ID出来匹配
    Select Case intImgMatchItem
        Case 0 'Patient ID
            aPatientID = Split(Replace(NVL(img.PatientID), "-", "_"), "_")
        Case 1 'Accession Number
            aPatientID = Split(Replace(NVL(img.Attributes(&H8, &H50).value), "-", "_"), "_")
        Case 2 'Patient Name
            aPatientID = Split(Replace(NVL(img.Name), "-", "_"), "_")
    End Select
    
    If UBound(aPatientID) >= 0 Then
        If UBound(aPatientID) > 0 Then
            strTemp = aPatientID(1)
        Else
            strTemp = aPatientID(0)
        End If
        
        '处理数值和字符串混合在一起的情况
        strTemp = getStrNumber(strTemp)
        funGetMatchIDInImg = strTemp
    Else
        funGetMatchIDInImg = -1
    End If
    Exit Function
err:
    funGetMatchIDInImg = 0
End Function

Private Function funIsPreMatched(ByVal blnMatchStudyUID As Boolean, ByVal intDBMatchItem As Integer, ByRef strStudyUID As String, _
    ByVal strAdviceID As String, ByVal strDeviceIP As String, ByVal strSeriesUID As String, ByVal strModality As String, _
    ByRef dtReceived As String, ByVal intFilterModality As Integer, ByRef strNewDeviceID As String, _
    ByVal strStoreDeviceNo As String, ByRef blnTmp As Boolean, ByRef str检查设备 As String, ByRef strPatientName As String, _
    ByRef strEnglishName As String, ByRef intAge As Integer, ByRef strSex As String, ByVal strStudyDateTime As String) As Long
'-----------------------------------------------------------------------------
'功能:判断是否已经有匹配成功的记录
'参数： blnMatchStudyUID ----是否匹配检查UID
'       intDBMatchItem ---- 匹配的数据库项目，0--检查号匹配，1--病人标识匹配，2--检查标识匹配
'       strStudyUID ---- [IN][OUT]检查UID，查询后如果查到，会修改检查UID
'       strAdviceID ---- 图像中的匹配ID，有三种情况：PatientID，PatientName，AccessionNumber，统称为医嘱ID
'       strDeviceIP ---- 存储设备IP
'       strSeriesUID ---- 图像序列UID
'       strModality ---- 影像类别
'       dtReceived ---[OUT] 如果预匹配成功，则返回数据库中“影像检查记录”的“接收日期”
'       intFilterModality ---- 是否按照影像类别过滤
'       strNewDeviceID ---- 查询到的新存储设备ID
'       strStoreDeviceNo ---- 原来的存储设备号
'       blnTmp ---- 是否匹配成临时记录
'       str检查设备 ---- [IN][OUT]图像中的检查设备，如果匹配成功，则修改成数据库中的检查设备
'       strPatientName ----[OUT] 如果匹配成功，返回数据库中的中文名
'       strEnglishName ----[OUT] 如果匹配成功，返回数据库中的英文名
'       intAge ----[OUT] 如果匹配成功，返回数据库中的年龄
'       strSex ----[OUT] 如果匹配成功，返回数据库中的性别
'       strStudyDateTime ---- 用来做预匹配查询的时间参数，当不按照检查UID匹配，使用的是可能重复的检查号，标识号时，使用此参数对时间进行比对，查找最近时间的检查记录。
'返回值：0-匹配成功，1-无匹配记录
'-----------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    If blnMatchStudyUID Then    '按照检查UID匹配
        Select Case intDBMatchItem
            Case 0, 1 '检查号匹配,'病人标识(门诊号，住院号)匹配 都直接判断检查UID是否相同
                strSQL = "Select 0 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 A" & _
                    " Where A.检查UID= [1]"
            Case 2 '检查标识（医嘱ID）匹配
                strSQL = "Select 0 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 Where 检查UID= [1]" & _
                    " Or (医嘱ID= [2] And 检查UID Is Not Null)"
        End Select
        strSQL = strSQL & " Union All Select 1 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像临时记录 Where 检查UID= [1]"
        
        strSQL = strSQL & " Union All Select 0 As 临时,C.检查UID,C.接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 C, 影像检查序列 D Where C.检查UID = D.检查UID And D.序列UID = [4] " & _
            " Union All Select 1 As 临时,E.检查UID,E.接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像临时记录 E, 影像临时序列 F Where E.检查UID = F.检查UID And F.序列UID = [4]"
    Else    '不按照检查UID匹配
        Select Case intDBMatchItem
            Case 0 '检查号匹配
                strSQL = "Select 0 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 A,病人医嘱发送 B,影像设备目录 E " & _
                    " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 AND A.检查设备 =E.设备号 AND (B.执行状态=3 " & _
                    " And B.执行过程>2 And " & _
                    IIf(intFilterModality = 0, " UPPER(E.影像类别)=[5] ", " E.IP地址=[3] ") & " And A.检查号= [2] And A.检查UID Is Not Null" & _
                    " And abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.首次时间) = (Select min(abs(to_date('" & _
                    strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-D.首次时间)) from 影像检查记录 C, 病人医嘱发送 D,影像设备目录 F Where C.医嘱ID=D.医嘱ID" & _
                    " And C.发送号=D.发送号 AND C.检查设备 =F.设备号 AND (D.执行状态=3 And D.执行过程>=2 And " & _
                    IIf(intFilterModality = 0, " UPPER(F.影像类别)=[5] ", " F.IP地址=[3] ") & " And C.检查号= [2])))"
            Case 1 '病人标识（门诊号，住院号）匹配
                strSQL = "Select 0 As 临时,检查UID,接收日期,位置一,位置二,检查设备,a.姓名,a.英文名,a.性别,a.年龄 From 影像检查记录 A,病人医嘱发送 B,病人医嘱记录 C,病人信息 D,影像设备目录 I " & _
                    " Where A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And A.医嘱ID=C.ID And C.病人ID=D.病人ID And A.检查设备 =I.设备号 " & _
                    " And B.执行状态=3 And B.执行过程>2 And " & _
                    IIf(intFilterModality = 0, " UPPER(I.影像类别)=[5] ", " I.IP地址=[3] ") & _
                    " And ((D.住院号=[2] AND C.病人来源=2) OR (D.门诊号= [2] AND C.病人来源<>2))" & _
                    " And A.检查UID Is Not Null  AND C.相关ID IS NULL  " & _
                    " And abs(to_date('" & strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-B.首次时间) = (Select min(abs(to_date('" & _
                    strStudyDateTime & "','YYYY-MM-DD HH24:MI:SS')-F.首次时间)) from 影像检查记录 E,病人医嘱发送 F,病人医嘱记录 G,病人信息 H,影像设备目录 J " & _
                    " Where E.医嘱ID=F.医嘱ID And E.发送号=F.发送号 And E.医嘱ID=G.ID And G.病人ID=H.病人ID AND E.检查设备 =J.设备号 AND G.相关ID IS NULL " & _
                    " And F.执行状态=3 And F.执行过程>=2 And " & _
                    IIf(intFilterModality = 0, " UPPER(J.影像类别)=[5] ", " J.IP地址=[3] ") & _
                    " And ((H.住院号=[2] AND G.病人来源=2) OR (H.门诊号= [2] AND G.病人来源<>2)))"
                    
            Case 2 '检查标识（医嘱ID）匹配
                strSQL = "Select 0 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 A Where  A.医嘱ID= [2] And A.检查UID Is Not Null"
        End Select
        strSQL = strSQL & " Union All Select 0 As 临时,C.检查UID,C.接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像检查记录 C, 影像检查序列 D Where C.检查UID = D.检查UID And D.序列UID = [4] " & _
            " Union All Select 1 As 临时,E.检查UID,E.接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像临时记录 E, 影像临时序列 F Where E.检查UID = F.检查UID And F.序列UID = [4]"
        '当查询的 lngAdviceID 号码不为-1时，再查询临时表看是否有匹配成功的。
        If strAdviceID <> "-1" Then
            strSQL = strSQL & " Union All Select 1 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像临时记录 Where 检查号= [2] and UPPER(影像类别) =[5] and 检查UID= [1] "
        Else
            '如果匹配ID是-1，说明匹配ID为空，此时查询临时表时，应该按照检查UID来查询，因为后面保存这种图像时，重新提取了一次匹配ID
            strSQL = strSQL & " Union All Select 1 As 临时,检查UID,接收日期,位置一,位置二,检查设备,姓名,英文名,性别,年龄 From 影像临时记录 Where 检查UID= [1]"
        End If
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "PACS图像保存", strStudyUID, strAdviceID, strDeviceIP, strSeriesUID, strModality)
               
    '记录SQL日志
    WriteProcessLog "funIsPreMatched", "判断是否已经有匹配成功的记录", "查询SQL为： " & vbCrLf & Replace(strSQL, "'", "‘") & vbCrLf _
                    & "参数为：1=" & strStudyUID & " ,2=" & strAdviceID & " ,3=" & strDeviceIP & " ,4=" & strSeriesUID & " ,5=" & strModality
    
    '如果匹配成功，则记录匹配医嘱的检查UID、接收时间、是否临时记录、检查设备等
    If rsTemp.EOF = False Then
        strStudyUID = rsTemp("检查UID")
        dtReceived = Format(rsTemp("接收日期"), "yyyyMMdd")
        blnTmp = IIf(rsTemp("临时") = 1, True, False)    '序列和图像是否放入临时记录中
        str检查设备 = NVL(rsTemp("检查设备"))
        strPatientName = NVL(rsTemp("姓名"))
        strEnglishName = NVL(rsTemp("英文名"))
        intAge = Val(NVL(rsTemp("年龄"), 0))
        strSex = NVL(rsTemp("性别"))
        
        '判断该图像所在的纪录中，存储设备是否等于当前设置的存储设备
        If NVL(rsTemp("位置一")) <> "" Then
            strNewDeviceID = NVL(rsTemp("位置一"))
        ElseIf NVL(rsTemp("位置二")) <> "" Then
            strNewDeviceID = NVL(rsTemp("位置二"))
        Else    '位置一和位置二都没有存储设备号
            '记录错误日志，然后使用当前设置的存储设备号
            WriteLog 11, 100, "从病人的影像检查记录中无法找到存储设备，使用网关设置的存储设备保存图像。" & " 病人姓名：" & strPatientName
            strNewDeviceID = strStoreDeviceNo
        End If
        funIsPreMatched = 0
    Else
        funIsPreMatched = 1
    End If
    Exit Function
err:
    Call WriteLog(1002, err.Number, "funIsPreMatched: " & err.Description)
    funIsPreMatched = 1
End Function

Public Function funcGetLocalIP() As String
'返回当前计算机的IP地址串，用逗号分隔
    Dim hostname As String * 256
    Dim hostent_addr As Long
    Dim host As HOSTENT
    Dim hostip_addr As Long
    Dim temp_ip_address() As Byte
    Dim i As Integer
    Dim ip_address As String
    Dim strLocalIPs As String

    '启动Socket
    Call SocketsInitialize

    If gethostname(hostname, 256) = SOCKET_ERROR Then
        MsgBox "Windows Sockets error " & Str(WSAGetLastError())
        Exit Function
    Else
        hostname = Trim$(hostname)
    End If

    hostent_addr = gethostbyname(hostname)

    If hostent_addr = 0 Then
        MsgBox "Winsock.dll is not responding."
        Exit Function
    End If

    RtlMoveMemory host, hostent_addr, LenB(host)
    RtlMoveMemory hostip_addr, host.hAddrList, 4

    ''''''''''''''''get all of the IP address if machine is  multi-homed

    Do
        ReDim temp_ip_address(1 To host.hLength)
        RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

        For i = 1 To host.hLength
            ip_address = ip_address & temp_ip_address(i) & "."
        Next
        ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)

        strLocalIPs = IIf(strLocalIPs = "", ip_address, strLocalIPs & "," & ip_address)

        ip_address = ""
        host.hAddrList = host.hAddrList + LenB(host.hAddrList)
        RtlMoveMemory hostip_addr, host.hAddrList, 4
     Loop While (hostip_addr <> 0)

    '清除Socket
    Call SocketsCleanup
    
    funcGetLocalIP = strLocalIPs
End Function


Private Sub SocketsInitialize()
    Dim WSAD As WSADATA
    Dim iReturn As Integer
    Dim sLowByte As String, sHighByte As String, sMsg As String

    iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

    If iReturn <> 0 Then
        MsgBox "Winsock.dll is not responding."
        End
    End If

    If lobyte(WSAD.wversion) < WS_VERSION_MAJOR Or (lobyte(WSAD.wversion) = _
        WS_VERSION_MAJOR And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then

        sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
        sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
        sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
        sMsg = sMsg & " is not supported by winsock.dll "
        MsgBox sMsg
        End
    End If

    ''''''''''''''''iMaxSockets is not used in winsock 2. So the following check is only
    ''''''''''''''''necessary for winsock 1. If winsock 2 is requested,
    ''''''''''''''''the following check can be skipped.

    If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
        sMsg = "This application requires a minimum of "
        sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) & " supported sockets."
        MsgBox sMsg
        End
    End If
End Sub

Private Sub SocketsCleanup()
Dim lReturn As Long

    lReturn = WSACleanup()

    If lReturn <> 0 Then
        MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred in Cleanup "
        End
    End If
End Sub

Private Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Private Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Private Function getStrNumber(ByVal strNumber As String) As String
    '处理数值和字符串混合在一起的情况，不能直接使用VAL，
    '因为号码可能超过16位，使用VAL会自动转换成科学记数法，而且最后几位还四舍五入
    
    On Error GoTo err
    If IsNumeric(strNumber) Then
        getStrNumber = strNumber
    Else
        If Val(strNumber) > 999999999999999# Then
            getStrNumber = 0
        Else
            getStrNumber = Val(strNumber)
        End If
    End If
    Exit Function
err:
      getStrNumber = 0
End Function

Public Sub WriteProcessLog(logSubName As String, logTitle As String, logDesc As String)
'功能： 记录操作日志
'参数： logSubName -- 执行操作的过程名称
'       logTitle  --  操作标题
'       logDesc --    日志描述

    Dim strSQL As String
    
    On Error Resume Next
    If gblnProcessLog Then
        If gcnAccess.State = adStateClosed Then Exit Sub
        
        strSQL = "Insert into DICOM通讯日志 (通讯时间,通讯函数,记录标题,记录内容) " & _
            "Values( cDate('" & Date & " " & Time() & "'),'" & logSubName & "','" & logTitle & _
            "','" & logDesc & "')"
        gcnAccess.Execute strSQL
    End If
End Sub

Public Sub subNewLogFile()
'功能： 创建新的日志文件
'参数： 无
    
    Dim strDate As String
    
    On Error GoTo err
    
    '创建当前的日期时间标记
    strDate = Date & "-" & Hour(Time) & "-" & Minute(Time) & "-" & Second(Time)
    
    '复制日志文件之前，先关闭日志文件
    If gcnAccess.State <> adStateClosed Then gcnAccess.Close
    FileCopy gstrAccessName, gstrAccessPath & "-" & strDate & ".mdb"
        
    '重新连接数据库
    gcnAccess.Open
    '清空当前日志中的内容
    gcnAccess.Execute "delete from DICOM通讯日志"
    gcnAccess.Execute "delete from 错误日志"
    gcnAccess.Execute "delete from 影像接收序列"
    
    
    '压缩数据库文件
    gcnAccess.Close
    DBEngine.CompactDatabase gstrAccessName, gstrAccessPath & "-zip.mdb"
    Kill gstrAccessName
    FileCopy gstrAccessPath & "-zip.mdb", gstrAccessName
    Kill gstrAccessPath & "-zip.mdb"
    gcnAccess.Open
    
    Exit Sub
err:
    Call WriteLog(1013, err.Number, "创建新日志出现错误，错误描述是：" & err.Description)
End Sub

Private Function funDeleteSpcialChar(strString As String) As String
'删除不可见的特殊字符，ascii码在0到32之间的特殊字符,32为空格
    Dim i As Integer
    Dim strText As String
    
    On Error GoTo err
    
    strText = strString
    For i = 0 To 32
        strText = Replace(strText, Chr(i), "")
    Next i
    
    funDeleteSpcialChar = strText
    Exit Function
err:
    funDeleteSpcialChar = strString
    
End Function

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function

Public Function DynamicGet(ByVal strclass As String, ByVal strCaption As String) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicGet = GetObject("", strclass)
   
    If err <> 0 Then
        MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, "提示"
        Set DynamicGet = Nothing
    End If
    err.Clear
End Function
