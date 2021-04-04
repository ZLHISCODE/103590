VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServiceHelper"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type
 
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

Public Type TransferFileMsg
    strAdviceId As String
    
    strName As String
    strSex As String
    strAge As String
    
    ftpInfo As TFtpDeviceInf
    bakFtpInfo As TFtpDeviceInf
    
    strLocalDir As String
    strFileName As String
    strSubDir As String
    strMediaType As String
End Type

Private mstrServerCachePath As String

Public Function SendDataToservice(ByVal strDataTag As String, ByVal intCommandIdentify As Integer, ByVal strDataFrom As String, fileMsg As TransferFileMsg) As Boolean
    Dim objFile As New FileSystemObject
    Dim pc As New clsPipeClient
    Dim pd As New clsPipeData
    Dim blnResult As Boolean
    
    With pd
        .DataName = fileMsg.strName
        .DataRundown = .DataName & "的图像"
        .DataTag = strDataTag
        
        Select Case intCommandIdentify
            Case LoadCommand.COMMAND_CAPIMG_UPLOAD
                .CommandIdentify = mdlBusinessConst.COMMAND_CAPIMG_UPLOAD
            Case LoadCommand.COMMAND_RPTIMG_DOWNLOAD
                .CommandIdentify = mdlBusinessConst.COMMAND_RPTIMG_DOWNLOAD
        End Select
        
        .DataFrom = strDataFrom
        
        Call pd.SetData(mdlBusinessConst.ADVICEID, fileMsg.strAdviceId)
        Call pd.SetData(mdlBusinessConst.Name, fileMsg.strName)
        Call pd.SetData(mdlBusinessConst.SEX, fileMsg.strSex)
        Call pd.SetData(mdlBusinessConst.AGE, fileMsg.strAge)
        
        Call pd.SetData(mdlBusinessConst.FTPDEVICE, fileMsg.ftpInfo.strDeviceId)
        Call pd.SetData(mdlBusinessConst.FTPADDR, fileMsg.ftpInfo.strFTPIP)
        Call pd.SetData(mdlBusinessConst.FTPUSER, fileMsg.ftpInfo.strFTPUser)
        Call pd.SetData(mdlBusinessConst.FTPPWD, getDecryptionPassW(fileMsg.ftpInfo.strFTPPwd))
        Call pd.SetData(mdlBusinessConst.FTPDIR, fileMsg.ftpInfo.strFtpDir)
        Call pd.SetData(mdlBusinessConst.SDDIR, fileMsg.ftpInfo.strSDDir)
        Call pd.SetData(mdlBusinessConst.SDUSER, fileMsg.ftpInfo.strSDUser)
        Call pd.SetData(mdlBusinessConst.SDPSWD, fileMsg.ftpInfo.strSDPswd)

        Call pd.SetData(mdlBusinessConst.BAKFTPDEVICE, fileMsg.bakFtpInfo.strDeviceId)
        Call pd.SetData(mdlBusinessConst.BAKFTPADDR, fileMsg.bakFtpInfo.strFTPIP)
        Call pd.SetData(mdlBusinessConst.BAKFTPUSER, fileMsg.bakFtpInfo.strFTPUser)
        Call pd.SetData(mdlBusinessConst.BAKFTPPWD, getDecryptionPassW(fileMsg.bakFtpInfo.strFTPPwd))
        Call pd.SetData(mdlBusinessConst.BAKFTPDIR, fileMsg.bakFtpInfo.strFtpDir)
        Call pd.SetData(mdlBusinessConst.BAKSDDIR, fileMsg.bakFtpInfo.strSDDir)
        Call pd.SetData(mdlBusinessConst.BAKSDUSER, fileMsg.bakFtpInfo.strSDUser)
        Call pd.SetData(mdlBusinessConst.BAKSDPSWD, fileMsg.bakFtpInfo.strSDPswd)

        Call pd.SetData(mdlBusinessConst.LOCALDIR, objFile.GetParentFolderName(fileMsg.strLocalDir) & "\")
        Call pd.SetData(mdlBusinessConst.FileName, fileMsg.strFileName)
        Call pd.SetData(mdlBusinessConst.SUBDIR, objFile.GetParentFolderName(fileMsg.strSubDir) & "\")
        Call pd.SetData(mdlBusinessConst.MEDIATYPE, fileMsg.strMediaType)
        
        blnResult = pc.SendData(pd)
        
        If Not blnResult Then
            blnResult = SaveFaildData(pd)
        End If
    End With
    
    Set pd = Nothing
    Set pc = Nothing
    
    SendDataToservice = blnResult
End Function

Private Function SaveFaildData(pd As clsPipeData) As Boolean
    Dim strServicePath As String
    Dim strFileTag As String
    Dim strFileName As String
    Dim strFilePath As String
    Dim objFile As New FileSystemObject

On Error GoTo errHandle

    strServicePath = GetServiceCachePath
    
    If Trim(strServicePath) <> "" Then
        strFileTag = GetNewGUID
        strFileName = "【VB】" & pd.DataName & "[" & strFileTag & "].XML"
        strFilePath = strServicePath & "\" & strFileName
        
        Call pd.SaveToXmlFile(strFilePath)
        objFile.GetFile(strFilePath).Attributes = Hidden
        
        SaveFaildData = True
    Else
        SaveFaildData = False
    End If
    
    Exit Function
errHandle:
    SaveFaildData = False
End Function

Private Function GetNewGUID() As String '生成GUID
    Dim uGUID As GUID
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long

    lLen = 40
 
    bGUID = String(lLen, 0)

    CoCreateGuid uGUID '把结构转换为一个可显示的字符串
 
    RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)

    sGUID = bGUID

    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1

    GetNewGUID = Replace(Replace(Replace(Left$(sGUID, RetVal), "-", ""), "{", ""), "}", "")
End Function

Public Function GetServiceStatus() As SERVICE_STATE
'获取PACS_SERVICE_NAME服务状态
    Dim objService As New clsService
    Dim enmServiceStatus As SERVICE_STATE

On Error GoTo errHandle
    objService.Name = PACS_SERVICE_NAME
    enmServiceStatus = objService.GetServiceStatus
    
    If enmServiceStatus = SERVICE_NOTEXIST Then
        '服务未安装
BUGEX "服务未安装打开失败，可能是未安装"
    ElseIf enmServiceStatus <> SERVICE_RUNNING Then
        '服务未正常启动，尝试启动
        objService.StartNTService
        enmServiceStatus = objService.GetServiceStatus
    Else
        '服务已启动
    End If
    
    Set objService = Nothing

    GetServiceStatus = enmServiceStatus
    
    Exit Function
errHandle:
    GetServiceStatus = SERVICE_NOTEXIST
End Function

Private Function GetServiceCachePath() As String
On Error GoTo errHandle
    Dim strPath As String
    Dim strConfig As String
    Dim strPathValue As String
    
    If Trim(mstrServerCachePath) <> "" Then
        GetServiceCachePath = mstrServerCachePath
        Exit Function
    End If
    
    strPath = CreateObject("WScript.Shell").RegRead("HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\" & PACS_SERVICE_NAME & "\ImagePath")
    
    mstrServerCachePath = Replace(Replace(strPath, PACS_SERVICE_NAME & ".exe", ""), """", "")
    
    If Dir(mstrServerCachePath, vbDirectory) = "" Then
        mstrServerCachePath = ""
        GetServiceCachePath = ""
        Exit Function
    End If
    
    '解析服务配置文件中的路径
    strConfig = OpenServerConfig(mstrServerCachePath + "\" + PACS_SERVICE_NAME & ".exe.config")
    mstrServerCachePath = mstrServerCachePath & "FileCache"
    If Trim(strConfig) = "" Then
        GetServiceCachePath = mstrServerCachePath
        Exit Function
    End If
    
    strPathValue = Mid(strConfig, InStr(strConfig, "文件缓存目录") + 7, 512)
    strPathValue = Mid(strPathValue, 1, InStr(strPathValue, "/>") - 1)
    
    strPathValue = Mid(strPathValue, InStr(strPathValue, """") + 1, 255)
    strPathValue = Mid(strPathValue, 1, InStr(strPathValue, """") - 1)
    
    If Trim(strPathValue) = "" Then
        GetServiceCachePath = mstrServerCachePath
        Exit Function
    End If
    
    If Dir(strPathValue, vbDirectory) = "" Then
        GetServiceCachePath = mstrServerCachePath
        Exit Function
    End If
    
    If strPathValue <> mstrServerCachePath Then
        mstrServerCachePath = strPathValue
    End If
    
    GetServiceCachePath = mstrServerCachePath
    
    Exit Function
errHandle:
    mstrServerCachePath = ""
    GetServiceCachePath = ""
End Function

'打开服务配置文件
Private Function OpenServerConfig(ByVal strConfigFile As String) As String
    Dim textBytes() As Byte
    Dim objConvert As New clsPipeClient
    
    OpenServerConfig = ""
    
    Open strConfigFile For Binary As #1
        'ReDim textBytes(LOF(1) - 2)
        
        ReDim textBytes(LOF(1))
        Get #1, 3, textBytes()
    
    Close #1
    
    OpenServerConfig = objConvert.Unicode8Decode(textBytes())
End Function


Private Sub Class_Initialize()
    mstrServerCachePath = ""
End Sub
