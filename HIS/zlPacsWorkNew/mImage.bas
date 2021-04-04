Attribute VB_Name = "mImage"
Option Explicit

Private Const GW_CHILD = 5
Private Const GW_HWNDNEXT = 2

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_HIDDEN = &H2


Public Const IMG_LAB_CHECKBOX_TAG = "CHECKBOX"
Public Const IMG_LAB_HINT_TAG = "HINT"
Public Const IMG_LAB_ORDER_TAG = "ORDER"
Public Const IMG_LAB_ERRORINFO_TAG = "ERRORINFO"
Public Const IMG_LAB_ERRORSTATE_TAG = "ERRORSTATE"

Public Const IMG_BACK_BORDER_COLOR = &HE0E0E0

Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As String, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long


Private mlastDicomInfo As TDicomBaseInfo
Private mlineFtpInfo As TFtpDeviceInf
Private mbackFtpInfo As TFtpDeviceInf


Public Function IsExistsBGServer() As String
'检查后台服务程序文件是否存在
    Dim strServerFile As String
    Dim strBgExe As String
    
    IsExistsBGServer = ""
    
    strBgExe = "ZL9PACSIMGTRANS"
    
    strServerFile = FormatFilePath(SysRootPath & "\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
        
    strServerFile = FormatFilePath(SysRootPath & "\Apply\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
        
    strServerFile = FormatFilePath(SysRootPath & "\PUBLIC\" & strBgExe & ".exe")
    If Trim(Dir(strServerFile, vbHidden)) <> "" Then
        IsExistsBGServer = strServerFile
        Exit Function
    End If
End Function

Public Function GetBgImgInfo(dcmInfo As TDicomBaseInfo, _
    lineFtpInfo As TFtpDeviceInf, backFtpInfo As TFtpDeviceInf, _
    Optional ByVal blnIsUpload As Boolean = True) As clsBgImgInfo
    
    Dim objBgImgInfo As clsBgImgInfo
    
    
    Set objBgImgInfo = New clsBgImgInfo
    
    objBgImgInfo.Key = dcmInfo.strInstanceUID
    objBgImgInfo.Filename = dcmInfo.strInstanceUID
    objBgImgInfo.FilePath = GetStudyImgPath(dcmInfo)  '图像所在本地存储路径
    objBgImgInfo.StudyUID = dcmInfo.strStudyUID
    
    objBgImgInfo.AdviceId = dcmInfo.lngAdviceId
    objBgImgInfo.PatientName = dcmInfo.strName
    objBgImgInfo.SeriesNoTag = dcmInfo.lngSeriesNo
    
    If blnIsUpload Then
        objBgImgInfo.ImgCommand = icUpLoad
        
        If dcmInfo.lngMediaTag = 0 Then objBgImgInfo.JpgConvert = True
    Else
        objBgImgInfo.ImgCommand = icReadly
    End If
    
    
    objBgImgInfo.IsBackGround = True '从参数读取是否后台处理图像上传
    
    Select Case dcmInfo.lngMediaTag
        Case ImgTag
            objBgImgInfo.Format = ifDcm
        Case VIDEOTAG
            objBgImgInfo.Format = ifAvi
        Case AUDIOTAG
            objBgImgInfo.Format = ifWav
    End Select
    
    With lineFtpInfo
        objBgImgInfo.FtpIp = .strFtpIp
        objBgImgInfo.FtpUser = .strFTPUser
        objBgImgInfo.FtpPwd = .strFTPPwd
        objBgImgInfo.FtpVirtualPath = .strFtpVirtualURL
        objBgImgInfo.FtpFile = dcmInfo.strInstanceUID
    End With

    If Len(backFtpInfo.strDeviceId) > 0 Then
    With backFtpInfo
        objBgImgInfo.BakIp = .strFtpIp
        objBgImgInfo.BakUser = .strFTPUser
        objBgImgInfo.BakPwd = .strFTPPwd
        objBgImgInfo.BakVirtualPath = .strFtpVirtualURL
    End With
    End If
    
    Set GetBgImgInfo = objBgImgInfo
End Function

Public Function GetStudyImgPath(ByRef dcmInfo As TDicomBaseInfo) As String
'获取检查图像路径
    Dim strPath As String
    
    strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "\" & dcmInfo.strStudyUID & "\")
    
    If DirExists(strPath) = False Then MkLocalDir strPath
    
    GetStudyImgPath = strPath
End Function


Public Function GetTempImgPath(Optional ByVal blnAutoCreate As Boolean = True) As String
'获取临时图像路径
    Dim strPath As String
    
    strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\")
    
    If DirExists(strPath) = False And blnAutoCreate Then MkLocalDir strPath
    
    GetTempImgPath = strPath
End Function

Public Function GetCachePath(ByVal strFmtDate As String, Optional ByVal strMark As String = "") As String
'获取缓存路径
    GetCachePath = FormatFilePath(SysRootPath & "\Apply\TmpAfterImage\" & strFmtDate & "\" & IIf(Len(strMark) <= 0, "", strMark & "\"))
End Function


Public Function GetLineFtpInfo(ByVal strLineDeviceNo As String, ByVal blnMoved As Boolean, ByRef dcmInfo As TDicomBaseInfo, ByRef strErr As String) As TFtpDeviceInf
'获取新的存储设备信息，如果设备存储信息不存在，则需要进行增加

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim curDate As Date
 
On Error GoTo errhandle
    strErr = ""
    
    If mlineFtpInfo.strIID = dcmInfo.strStudyUID Then
        With mlineFtpInfo
            GetLineFtpInfo.strDeviceId = .strDeviceId
            GetLineFtpInfo.strFtpDir = .strFtpDir
            GetLineFtpInfo.strFtpIp = .strFtpIp
            GetLineFtpInfo.strFTPPwd = .strFTPPwd
            GetLineFtpInfo.strFTPUser = .strFTPUser
            GetLineFtpInfo.strFtpVirtualURL = .strFtpVirtualURL
            GetLineFtpInfo.strIID = .strIID
        End With
    Else
        strSQL = "Select D.IP地址 As Host, D.FTP用户名 As FtpUser,D.FTP密码 As FtpPwd, Decode(C.位置一,Null,C.位置二,C.位置一) as 位置,C.接收日期," & _
            "'/'|| D.Ftp目录 ||'/' As Root, Decode(C.接收日期, Null,'',to_Char(C.接收日期,'YYYYMMDD') || '/') || C.检查UID As URL " & _
            " From 影像检查记录 C,影像设备目录 D " & _
            " Where Decode(C.位置一,Null,C.位置二,C.位置一)=D.设备号(+)" & _
            " And C.检查UID= [1]"
        If blnMoved Then strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查存储设备", dcmInfo.strStudyUID)
    
        blnIsGetNewDevice = False
    
        If rsData.RecordCount <= 0 Then
            blnIsGetNewDevice = True
        Else
            '如果执行到这里，说明是执行图像关联,需要判断当前检查的存储设备是否有效，如果无效需生成新的存储设备
            If Trim(rsData!接收日期) = "" Or nvl(rsData!位置) = "" Then
                blnIsGetNewDevice = True
            Else
                GetLineFtpInfo.strDeviceId = nvl(rsData!位置)
                GetLineFtpInfo.strFtpIp = nvl(rsData!Host)
                GetLineFtpInfo.strFtpDir = nvl(rsData!Root)
                GetLineFtpInfo.strFTPUser = nvl(rsData!FtpUser)
                GetLineFtpInfo.strFTPPwd = nvl(rsData!FtpPwd)
                GetLineFtpInfo.strFtpVirtualURL = GetLineFtpInfo.strFtpDir & nvl(rsData!Url)
            End If
        End If
    
        If blnIsGetNewDevice Then
            If Val(strLineDeviceNo) <= 0 Then
                strErr = "未找到图像存储设备,请确认对应存储设备已进行了设置。"
                Exit Function
            End If
     
            strSQL = "Select 设备号,设备名,'/'|| Ftp目录 || '/' As Root,FTP用户名,FTP密码,IP地址 " & _
                        " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
    
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询采集在线存储设备", strLineDeviceNo)
    
            '如果存储设备停用，则直接退出
            If rsTemp.RecordCount <= 0 Then
                strErr = "未找到在线存储设备,请确认设备号为 [" & strLineDeviceNo & "] 的设备是否启用。"
                Exit Function
            End If
    
            GetLineFtpInfo.strDeviceId = strLineDeviceNo
            GetLineFtpInfo.strFtpIp = nvl(rsTemp("IP地址"))
            GetLineFtpInfo.strFTPUser = nvl(rsTemp("FTP用户名"))
            GetLineFtpInfo.strFTPPwd = nvl(rsTemp("FTP密码"))
            GetLineFtpInfo.strFtpDir = nvl(rsTemp("Root"))
    
            GetLineFtpInfo.strFtpVirtualURL = GetLineFtpInfo.strFtpDir & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "/" & dcmInfo.strStudyUID
     
        End If
    End If
    
    mlineFtpInfo = GetLineFtpInfo
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function ResetStorageDevice(ByVal lngAdviceId As Long, ByRef objImgInf As clsBgImgInfo, ByVal blnMoved As Boolean) As String
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    ResetStorageDevice = ""
    
    strSQL = " Select A.检查UID,to_Char(A.接收日期,'YYYYMMDD') As 接收日期, Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/') ||A.检查UID||'/' As URL," & _
            " B.设备号 as 设备号1, B.设备名 As 设备名1, B.FTP用户名 As User1,B.FTP密码 As Pwd1, B.IP地址 As Host1, " & _
                    " decode(B.Ftp目录, null, '/', '/'||B.Ftp目录||'/') As Root1,B.共享目录 as 共享目录1,B.共享目录用户名 as 共享目录用户名1,B.共享目录密码 as 共享目录密码1 " & _
            " From  影像检查记录 A,影像设备目录 B " & _
            " Where A.医嘱ID=[1] And Nvl(A.位置一, 位置二) = B.设备号(+) "
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图像存储", lngAdviceId)
            
    If rsData.RecordCount <= 0 Then
        ResetStorageDevice = "未找到检查对应的图像存储设备，请检查数据是否正确。"
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        objImgInf.DeviceNo = Val(nvl(rsData!设备号1))
        objImgInf.FtpIp = nvl(rsData!Host1)
        objImgInf.FtpUser = nvl(rsData!User1)
        objImgInf.FtpPwd = nvl(rsData!Pwd1)
        objImgInf.FtpVirtualPath = nvl(rsData!Root1) & nvl(rsData!Url)
    End If
    
    objImgInf.AdviceId = lngAdviceId
    objImgInf.StudyUID = nvl(rsData!检查UID)
    objImgInf.RecFmtDate = nvl(rsData!接收日期)
    objImgInf.FilePath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & nvl(rsData!接收日期) & "\" & nvl(rsData!检查UID) & "\")
 
Exit Function
errhandle:
    ResetStorageDevice = err.Description
End Function

Public Function GetNewStorageDevice(ByVal lngAdviceId As Long, _
    ByVal strStudyUID As String, ByVal strRecFmtDate As String, _
    ByRef objImgInf As clsBgImgInfo) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDeviceNo As String
    Dim strRoot As String
    
    GetNewStorageDevice = ""
    '查询医技工作站中，检查所对应的存储设备
    strSQL = "select d.参数值 " & _
                " from 医技执行房间 a, 病人医嘱发送 b, 影像DICOM服务对 c, 影像DICOM服务参数 d " & _
                " Where a.科室ID = b.执行部门id And a.执行间 = b.执行间 And a.检查设备 = c.设备号 " & _
                " and c.服务功能='图像接收' and c.服务ID=d.服务ID and d.参数名称='存储设备' and b.医嘱id=[1]"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then
        GetNewStorageDevice = "未找到图像存储设备,请确认当前检查所用设备是否在影像设备目录的服务配置中配置了图像存储。"
        Exit Function
    End If
    
    strDeviceNo = nvl(rsTemp!参数值)


    strSQL = "Select 设备号,设备名,'/'||Decode(Ftp目录,Null,'',Ftp目录 || '/') As Root,FTP用户名,FTP密码,IP地址 " & _
                " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"
                
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, strDeviceNo)
    
    '如果存储设备停用，则直接退出
    If rsTemp.RecordCount <= 0 Then
        GetNewStorageDevice = "未找到存储设备,请确认设备号为 [" & strDeviceNo & "] 的设备是否启用。"
        Exit Function
    End If
    
    
    strRoot = nvl(rsTemp("Root"))
    
    objImgInf.DeviceNo = Val(strDeviceNo)
    objImgInf.AdviceId = lngAdviceId
    objImgInf.StudyUID = strStudyUID
    objImgInf.RecFmtDate = strRecFmtDate
    
    objImgInf.FtpIp = nvl(rsTemp("IP地址"))
    objImgInf.FtpUser = nvl(rsTemp("FTP用户名"))
    objImgInf.FtpPwd = nvl(rsTemp("FTP密码"))
    objImgInf.FtpVirtualPath = IIf(strRoot = "/", "//", strRoot) & strRecFmtDate & "/" & strStudyUID & "/"
    
    
    objImgInf.FilePath = FormatFilePath(SysRootPath & "\Apply\TmpImage\" & strRecFmtDate & "\" & strStudyUID & "\")

End Function

Public Function GetBackFtpInfo(ByVal strBackDeviceNo As String, ByRef dcmInfo As TDicomBaseInfo, ByRef strErr As String) As TFtpDeviceInf
'获取新的存储设备信息，如果设备存储信息不存在，则需要进行增加

    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnIsGetNewDevice As Boolean
    Dim curDate As Date
 
On Error GoTo errhandle
    strErr = ""
    
    If mbackFtpInfo.strIID = dcmInfo.strStudyUID Then
        With mbackFtpInfo
            GetBackFtpInfo.strDeviceId = .strDeviceId
            GetBackFtpInfo.strFtpDir = .strFtpDir
            GetBackFtpInfo.strFtpIp = .strFtpIp
            GetBackFtpInfo.strFTPPwd = .strFTPPwd
            GetBackFtpInfo.strFTPUser = .strFTPUser
            GetBackFtpInfo.strFtpVirtualURL = .strFtpVirtualURL
            GetBackFtpInfo.strIID = .strIID
        End With
    Else
        If Len(strBackDeviceNo) <= 0 Then Exit Function
    
        strSQL = "Select 设备号,设备名,'/'|| Ftp目录 || '/' As Root,FTP用户名,FTP密码,IP地址 " & _
                    " From 影像设备目录 Where 类型=1 and 设备号=[1] and NVL(状态,0)=1"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询采集备份存储设备", strBackDeviceNo)

        '如果存储设备停用，则直接退出
        If rsTemp.RecordCount <= 0 Then
            strErr = "未找到备份存储设备,请确认设备号为 [" & strBackDeviceNo & "] 的设备是否启用。"
            Exit Function
        End If

        GetBackFtpInfo.strDeviceId = strBackDeviceNo
        GetBackFtpInfo.strFtpIp = nvl(rsTemp("IP地址"))
        GetBackFtpInfo.strFTPUser = nvl(rsTemp("FTP用户名"))
        GetBackFtpInfo.strFTPPwd = nvl(rsTemp("FTP密码"))
        GetBackFtpInfo.strFtpDir = nvl(rsTemp("Root"))

        GetBackFtpInfo.strFtpVirtualURL = GetBackFtpInfo.strFtpDir & Format(dcmInfo.strReceiveFullTime, "yyyymmdd") & "/" & dcmInfo.strStudyUID
      
    End If
    
    mbackFtpInfo = GetBackFtpInfo
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetDicomAge(ByVal dtBirth As String, Optional ByVal strAge As String = "") As String
'获取Dicom年龄格式
    Dim dtStart As Date
    Dim lngDays As Long
    Dim lngAge As Long
    
    GetDicomAge = ""
    
    If Len(dtBirth) > 0 Then
        dtStart = CDate(Format(dtBirth, "yyyy-mm-dd"))
        lngDays = DateDiff("d", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
        
        '转换为岁,月,周,天
        Select Case True
            Case lngDays > 365 * 3 '3岁
                '岁
                GetDicomAge = DateDiff("yyyy", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "Y"
            Case lngDays > 30 * 3 '3月
                '月
                GetDicomAge = DateDiff("m", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "M"
            Case lngDays > 7 * 4 '一月
                '周
                GetDicomAge = DateDiff("ww", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "W"
            Case Else
                '天
                GetDicomAge = DateDiff("d", CDate(Format(dtBirth, "yyyy-mm-dd")), zlDatabase.Currentdate)
                GetDicomAge = Format(GetDicomAge, "000") & "D"
        End Select
        
        Exit Function
    End If
    
    If Len(strAge) > 0 Then
        '根据录入的年龄转换为dicom格式的年龄形式
        lngAge = Val(strAge)
        
        Select Case True
            Case (InStr(strAge, "岁") > 0), (InStr(UCase(strAge), "Y") > 0):
                GetDicomAge = Format(lngAge, "000") & "Y"
            Case (InStr(strAge, "月") > 0), (InStr(UCase(strAge), "M") > 0):
                GetDicomAge = Format(lngAge, "000") & "M"
            Case (InStr(strAge, "周") > 0), (InStr(UCase(strAge), "W") > 0):
                GetDicomAge = Format(lngAge, "000") & "W"
            Case Else
                GetDicomAge = Format(lngAge, "000") & "D"
        End Select
    End If
    
End Function


Public Function GetDicomBaseInfoEx(ByVal lngAdviceId As Long, dcmImg As DicomImage, Optional ByRef strDeviceNo As String) As TDicomBaseInfo
    Dim objValue As Variant
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    
    If dcmImg Is Nothing Then Exit Function
     
    If mlastDicomInfo.lngAdviceId = lngAdviceId Then
        With mlastDicomInfo
            GetDicomBaseInfoEx.lngAdviceId = .lngAdviceId
            GetDicomBaseInfoEx.lngSendNo = .lngSendNo
            GetDicomBaseInfoEx.lngID = .lngID
            GetDicomBaseInfoEx.strAge = .strAge
            GetDicomBaseInfoEx.strBirthDate = .strBirthDate
            
            GetDicomBaseInfoEx.strInstitution = .strInstitution
            GetDicomBaseInfoEx.strModality = .strModality
            GetDicomBaseInfoEx.strName = .strName
            GetDicomBaseInfoEx.strSex = .strSex
            GetDicomBaseInfoEx.strReceiveFullTime = IIf(Len(.strReceiveFullTime) > 0, .strReceiveFullTime, zlDatabase.Currentdate)
            
            GetDicomBaseInfoEx.strStudyUID = .strStudyUID
            GetDicomBaseInfoEx.strSeriesUID = .strSeriesUID
            GetDicomBaseInfoEx.strDeviceNo = .strDeviceNo
             
        End With
    Else
        strSQL = "select b.病人ID,a.发送号,a.影像类别,a.检查设备,nvl(a.位置一, a.位置二) as 存储位置, a.姓名,a.性别,a.出生日期,a.年龄,a.检查UID,a.接收日期 " & _
                " From 影像检查记录 a, 病人医嘱记录 b " & _
                " Where a.医嘱ID=b.Id and a.医嘱ID=[1]"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查信息", lngAdviceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
    
        GetDicomBaseInfoEx.lngAdviceId = lngAdviceId
        GetDicomBaseInfoEx.lngSendNo = Val(nvl(rsData!发送号))
        GetDicomBaseInfoEx.lngID = Val(nvl(rsData!病人ID))
        GetDicomBaseInfoEx.strAge = GetDicomAge(nvl(rsData!出生日期), nvl(rsData!年龄))
        GetDicomBaseInfoEx.strBirthDate = Format(nvl(rsData!出生日期), "yyyymmdd")
        GetDicomBaseInfoEx.strInstitution = RegInstitution
        GetDicomBaseInfoEx.strModality = nvl(rsData!影像类别)
        GetDicomBaseInfoEx.strName = nvl(rsData!姓名)
        GetDicomBaseInfoEx.strSex = Decode(nvl(rsData!性别), "男", "M", "女", "F", "O")
        GetDicomBaseInfoEx.strReceiveFullTime = nvl(rsData!接收日期, zlDatabase.Currentdate) ', "yyyymmdd")
        GetDicomBaseInfoEx.strStudyUID = nvl(rsData!检查UID)
        GetDicomBaseInfoEx.strDeviceNo = nvl(rsData!存储位置)
    End If
    
    objValue = dcmImg.Attributes(&H10, &H20).value      '病人ID
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.lngID = Val(objValue)
     
    objValue = dcmImg.Attributes(&H10, &H1010).value      '病人年龄
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strAge = objValue
     
    objValue = dcmImg.Attributes(&H10, &H30).value      '病人出生日期
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strBirthDate = dcmImg.DateOfBirthAsDate
     
    objValue = dcmImg.Attributes(&H8, &H80).value      '单位机构
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strInstitution = objValue
     
    objValue = dcmImg.Attributes(&H8, &H60).value      '影像类别
    If Not IsNull(objValue) Then GetDicomBaseInfoEx.strInstitution = objValue
      
    If Len(dcmImg.StudyUID) > 0 Then GetDicomBaseInfoEx.strStudyUID = dcmImg.StudyUID     '检查UID
    If Len(dcmImg.SeriesUID) > 0 Then GetDicomBaseInfoEx.strSeriesUID = dcmImg.SeriesUID  '序列UID
    If Len(dcmImg.InstanceUID) > 0 Then GetDicomBaseInfoEx.strInstanceUID = dcmImg.InstanceUID    '实例UID
    
    GetDicomBaseInfoEx.lngSeriesNo = Val(dcmImg.Attributes(&H20, &H11).value)  '序列号
    GetDicomBaseInfoEx.lngImgNo = Val(dcmImg.Attributes(&H20, &H13).value)     '图像号
    
    mlastDicomInfo = GetDicomBaseInfoEx
End Function


Public Function GetDicomBaseInfo(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As TDicomBaseInfo
'获取Dicom基本信息
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    If mlastDicomInfo.lngAdviceId = lngAdviceId Then
        With mlastDicomInfo
            GetDicomBaseInfo.lngAdviceId = .lngAdviceId
            GetDicomBaseInfo.lngSendNo = .lngSendNo
            GetDicomBaseInfo.lngID = .lngID
            GetDicomBaseInfo.strAge = .strAge
            GetDicomBaseInfo.strBirthDate = .strBirthDate
            
            GetDicomBaseInfo.strInstitution = .strInstitution
            GetDicomBaseInfo.strModality = .strModality
            GetDicomBaseInfo.strName = .strName
            GetDicomBaseInfo.strSex = .strSex
            GetDicomBaseInfo.strReceiveFullTime = .strReceiveFullTime
            
            GetDicomBaseInfo.strStudyUID = .strStudyUID
            GetDicomBaseInfo.strSeriesUID = .strSeriesUID
            GetDicomBaseInfo.strInstanceUID = CreateUID
            
            GetDicomBaseInfo.lngSeriesNo = .lngSeriesNo
            
            GetDicomBaseInfo.lngImgNo = .lngImgNo + 1
        End With
    Else
        strSQL = "select b.病人ID,a.发送号,a.影像类别,a.检查设备,a.姓名,a.性别,a.出生日期,a.年龄,a.检查UID,a.接收日期 " & _
                " From 影像检查记录 a, 病人医嘱记录 b " & _
                " Where a.医嘱ID=b.Id and a.医嘱ID=[1]"
                
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查信息", lngAdviceId)
        
        If rsData.RecordCount <= 0 Then Exit Function
    
        GetDicomBaseInfo.lngAdviceId = lngAdviceId
        GetDicomBaseInfo.lngSendNo = Val(nvl(rsData!发送号))
        GetDicomBaseInfo.lngID = Val(nvl(rsData!病人ID))
        GetDicomBaseInfo.strAge = GetDicomAge(nvl(rsData!出生日期), nvl(rsData!年龄))
        GetDicomBaseInfo.strBirthDate = Format(nvl(rsData!出生日期), "yyyymmdd")
        GetDicomBaseInfo.strInstanceUID = CreateUID
        GetDicomBaseInfo.strInstitution = RegInstitution
        GetDicomBaseInfo.strModality = nvl(rsData!影像类别)
        GetDicomBaseInfo.strName = nvl(rsData!姓名)
        GetDicomBaseInfo.strSex = Decode(nvl(rsData!性别), "男", "M", "女", "F", "O")
        GetDicomBaseInfo.strReceiveFullTime = nvl(rsData!接收日期, zlDatabase.Currentdate) ', "yyyymmdd")
        
        GetDicomBaseInfo.strStudyUID = nvl(rsData!检查UID)
        GetDicomBaseInfo.lngSeriesNo = 1
        GetDicomBaseInfo.lngImgNo = 1
        
        If Len(GetDicomBaseInfo.strStudyUID) > 0 Then   'lngImgNo=0表示第一次采集图像
            '获取序列UID和图像号
            strSQL = "Select 序列UID,序列号 From 影像检查序列 Where 检查UID=[1] and 序列描述 is Null order by 序列号"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询序列信息", GetDicomBaseInfo.strStudyUID)
            
            If rsData.RecordCount > 0 Then
                GetDicomBaseInfo.strSeriesUID = nvl(rsData!序列UID)
                GetDicomBaseInfo.lngSeriesNo = Val(nvl(rsData!序列号))
                
                strSQL = "select max(nvl(图像号, 0)) as 图像号 From 影像检查图象 Where 序列UID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询图像号", GetDicomBaseInfo.strSeriesUID)
                
                If rsData.RecordCount > 0 Then
                    GetDicomBaseInfo.lngImgNo = Val(nvl(rsData!图像号)) + 1
                End If
                
            Else
                GetDicomBaseInfo.strSeriesUID = CreateUID
            End If
        Else
            GetDicomBaseInfo.strStudyUID = CreateUID     '获取新的检查UID
            GetDicomBaseInfo.strSeriesUID = CreateUID    '获取新的序列UID
        End If
    End If
    
    mlastDicomInfo = GetDicomBaseInfo
End Function


Public Sub WriteDicomPara(img As DicomImage, dicomInfo As TDicomBaseInfo, _
    Optional blnIsAfterCapture As Boolean = False)
'------------------------------------------------
'功能：给输入的图像填写DICOM文件头信息
'参数：img－－输入的DICOM文件,lngAdviceID－－医嘱ID
'返回：无，直接文件头信息写入img的文件头
'------------------------------------------------
    Dim curDate As Date

    curDate = zlDatabase.Currentdate
    
    If blnIsAfterCapture Then
        img.Attributes.Add &H10, &H10, ""                           'Name 姓名
        img.Attributes.Add &H10, &H20, ""                           'Patient ID 病人ID
        img.Attributes.Add &H10, &H30, ""                           'BirthDate 生日
        img.Attributes.Add &H10, &H40, ""                           'Sex 性别
        img.Attributes.Add &H10, &H1010, ""                         'Age 年龄
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
        img.Attributes.Add &H20, &H10, ""                           'Study ID 检查ID
        img.Attributes.Add &H8, &H60, dicomInfo.strModality         'Modality 影像类别
        img.Attributes.Add &H20, &H11, "1"                          'Series Number 序列号
        img.Attributes.Add &H20, &H13, "1"                          'ImageNumber 图像号
    Else
        img.Attributes.Add &H10, &H10, dicomInfo.strName            'Name 姓名
        img.Attributes.Add &H10, &H20, dicomInfo.lngID              'Patient ID 病人ID
        img.Attributes.Add &H10, &H30, dicomInfo.strBirthDate       'BirthDate 生日
        img.Attributes.Add &H10, &H40, dicomInfo.strSex             'Sex 性别
        img.Attributes.Add &H10, &H1010, dicomInfo.strAge           'Age 年龄
        img.Attributes.Add &H10, &H4000, ""                         'Patient Comment 病人注释
        img.Attributes.Add &H8, &H60, dicomInfo.strModality         'Modality 影像类别
        
        img.StudyUID = dicomInfo.strStudyUID                        ' &H20, &H10 Study ID 检查ID
        img.SeriesUID = dicomInfo.strSeriesUID                      ' 序列UID
        img.InstanceUID = dicomInfo.strInstanceUID                  '图像实例UID
        
        img.Attributes.Add &H20, &H11, dicomInfo.lngSeriesNo        'Series Number 序列号
        img.Attributes.Add &H20, &H13, dicomInfo.lngImgNo           'ImageNumber 图像号
    End If
    
    img.Attributes.Add &H8, &H8, ""                                 'ImageType  空
    img.Attributes.Add &H8, &H16, "1.2.840.10008.5.1.4.1.1.7"       'SOP Class  UID，二次捕捉
    img.Attributes.Add &H8, &H20, Format(curDate, "yyyy-mm-dd")     'Study Date 检查日期
    img.Attributes.Add &H8, &H21, Format(curDate, "yyyy-mm-dd")     'Series Date 序列日期
    img.Attributes.Add &H8, &H22, Format(curDate, "yyyy-mm-dd")     'Acquisition Date 采集日期
    img.Attributes.Add &H8, &H23, Format(curDate, "yyyy-mm-dd")     'Image Date   图像日期
    img.Attributes.Add &H8, &H30, Format(curDate, "HH24:MI:SS")     'Study Time   检查时间
    img.Attributes.Add &H8, &H31, Format(curDate, "HH24:MI:SS")     'Series Time  序列时间
    img.Attributes.Add &H8, &H32, Format(curDate, "HH24:MI:SS")     'Acquisition Time  采集时间
    img.Attributes.Add &H8, &H33, Format(curDate, "HH24:MI:SS")     'Image Time  图像时间
    img.Attributes.Add &H8, &H50, ""                            'Accession Number 空
    img.Attributes.Add &H8, &H70, "ZLSOFT"                      'Manufacturer 厂商
    img.Attributes.Add &H8, &H80, RegInstitution               'Institution Name 单位名称
    img.Attributes.Add &H8, &H90, ""                            'Referring Physician's Name 空
    img.Attributes.Add &H8, &H1030, ""                          'Study Description 检查描述 空

    img.Attributes.Add &H20, &H20, ""                           'Orientation 空
    
End Sub

Public Sub SaveImageInfo(ByRef dcmInfo As TDicomBaseInfo, ByRef ftpInfo As TFtpDeviceInf)
'保存采集图像
    Dim arySql() As String
    Dim strSQL As String
    Dim blnInTrans As Boolean
    
    Dim rsData As ADODB.Recordset
    Dim blnHasStudy As Boolean
    Dim blnHasSeries As Boolean
    
    Dim i As Long
    
On Error GoTo errhandle

    ReDim arySql(0)
    
    strSQL = "select 1 from 影像检查记录 where 检查UID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查信息", dcmInfo.strStudyUID)
    
    blnHasStudy = IIf(rsData.RecordCount <= 0, False, True)
    If blnHasStudy Then
        '判断是否有对应的序列信息
        strSQL = "select 1 from 影像检查序列 where 序列UID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询检查信息", dcmInfo.strSeriesUID)
        
        blnHasSeries = IIf(rsData.RecordCount <= 0, False, True)
    Else
        blnHasSeries = False
    End If
    
    If blnHasStudy = False Then
        '首次采集图像,需要写入采集基本信息和存储设备信息...
        strSQL = "ZL_影像检查记录_SET(" & dcmInfo.lngAdviceId & "," & dcmInfo.lngSendNo & ",'" & _
                                        dcmInfo.strStudyUID & "',null," & _
                                        "to_Date('" & Format(dcmInfo.strReceiveFullTime, "yyyy-mm-dd hh:mm") & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                                        ftpInfo.strDeviceId & "')"
                                        
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = strSQL
        
        dcmInfo.lngImgNo = 1
    End If
    
    If blnHasSeries = False Then
        strSQL = "ZL_影像序列_INSERT('" & dcmInfo.strStudyUID & "','" & dcmInfo.strSeriesUID & "','" & dcmInfo.strSeriesDes & "',0)"
        
        ReDim Preserve arySql(UBound(arySql) + 1)
        arySql(UBound(arySql)) = strSQL
    End If
    
    If dcmInfo.lngMediaTag = 0 Then
        strSQL = "ZL_影像图象_INSERT('" & dcmInfo.strInstanceUID & "','" & dcmInfo.strSeriesUID & "',NULL,0, null, sysdate)"
    Else
        strSQL = "ZL_影像图象_INSERT('" & dcmInfo.strInstanceUID & "','" & dcmInfo.strSeriesUID & "',Null,0" & _
        ",null,sysdate,null,null,null,null,null,null,null,null,null," & dcmInfo.lngMediaTag & ",'" & dcmInfo.strMediaEncode & "'," & dcmInfo.lngMediaLen & ")"
    End If
    
    ReDim Preserve arySql(UBound(arySql) + 1)
    arySql(UBound(arySql)) = strSQL
        
    gcnOracle.BeginTrans        '----------保存媒体，图像，视频，音频
    blnInTrans = True
    
    For i = 1 To UBound(arySql)
        Call zlDatabase.ExecuteProcedure(CStr(arySql(i)), "批量执行采集媒体保存")
    Next i
    
    gcnOracle.CommitTrans
Exit Sub
errhandle:
    If blnInTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function IsValidFile(ByVal strFile As String, Optional ByVal lngSize As Long = 0)
    Dim objFileSystem As New FileSystemObject
    Dim lngFileSize As Long
    
On Error GoTo errhandle
    IsValidFile = False
    
    If Trim(Dir(strFile, 7)) = "" Then Exit Function
    
    lngFileSize = lngSize
    If lngFileSize <= 0 Then lngFileSize = 1000
    
    If objFileSystem.GetFile(strFile).Size < lngFileSize Then Exit Function
    
    IsValidFile = True
    
    Set objFileSystem = Nothing
Exit Function
errhandle:
    IsValidFile = False
    Set objFileSystem = Nothing
End Function

Private Function IsFileLocked(ByVal strFileName As String) As Boolean
   Dim iFn As Integer
   Dim blnRetVal As Boolean
   
On Error GoTo E_HandleFA
    blnRetVal = True
    
    If (Len(Dir$(strFileName, 7)) > 0) Then
       iFn = FreeFile
       
       Open strFileName For Binary Lock Read Write As #iFn
       Close iFn
       
       blnRetVal = False
       
       blnRetVal = IsFileOpen(strFileName)
    Else
        '文件不存在，则返回未锁定状态
        blnRetVal = False
    End If
   
E_HandleFA:
   IsFileLocked = blnRetVal
End Function


Private Function IsFileOpen(ByVal pFile As String) As Boolean
    Dim ret As Long
    
    ret = CreateFile(pFile, GENERIC_READ Or GENERIC_WRITE, 0&, vbNullString, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    
    IsFileOpen = (ret = INVALID_HANDLE_VALUE)
    
    CloseHandle ret
End Function

Private Function WaitReadDcm(dImgs As DicomImages, ByVal strFile As String, _
    ByRef strError As String) As DicomImage
On Error Resume Next
    Dim i As Long
    Dim blnUseUrl As Boolean
    
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    Set WaitReadDcm = Nothing
    
    While True
        err.Clear
        dImgs.Clear
        
        If blnUseUrl Then
            'readurl不支持空格
            Set WaitReadDcm = dImgs.ReadURL(strFile)
        Else
            Set WaitReadDcm = dImgs.ReadFile(strFile)
        End If
        
        If err.Description = "" Then Exit Function
        
        i = i + 1
         
        If i > 100 Then
            strError = err.Description
            Exit Function
        End If
        
        Call Sleep(10)
    Wend
End Function

Private Function GetFileSize(ByVal strFile As String) As Long
'获取文件大小
On Error GoTo errhandle
    GetFileSize = FileLen(strFile)
Exit Function
errhandle:
    GetFileSize = 0
End Function

Public Function ReadDicomFile(ByVal strFile As String, ByRef strError As String, _
    Optional ByVal blnIsDcmFormat As Boolean = False) As DicomImage
On Error Resume Next
    Dim dImgs As New DicomImages
        
    Dim curImage As DicomImage
    Dim blnUseUrl As Boolean
    Dim strFileTime As String
    Dim strCopyFileName As String
    Dim lngSize As Long
    
    strError = ""
    blnUseUrl = IIf(InStr(strFile, " ") <= 0, True, False)
    
    '返回占位
    If strFile = "NULL" Then
        Set curImage = dImgs.AddNew
    
        dImgs.Clear
        Set dImgs = Nothing
        
        Set ReadDicomFile = curImage
        
        Exit Function
    End If
    
    If blnUseUrl Then
        'readurl不支持空格
        Set curImage = dImgs.ReadURL(strFile)
    Else
        Set curImage = dImgs.ReadFile(strFile)
    End If
    
    If err.Number = 0 Then
        If Not curImage Is Nothing Then
            If Len(curImage.InstanceUID) > 0 Then
                dImgs.Clear
                Set dImgs = Nothing
                
                Set ReadDicomFile = curImage
                Exit Function
            End If
        End If
    End If
    

    '2098错误一种是文件不是dicom文件，另一种是存在共享访问错误
    If InStr(err.Description, "sharing violation") > 0 Then
        
        lngSize = GetFileSize(strFile)
        
        strFileTime = Format(Now, "MMDD") & GetTickCount
        strCopyFileName = strFile & "_copy_vdat_" & strFileTime
        
        Call FileCopy(strFile, strCopyFileName)
        
        err.Clear
        
        If IsValidFile(strCopyFileName, lngSize) = False Then
            '文件复制失败，尝试重新复制
            If WaitCopy(strFile, strCopyFileName, strError, lngSize) = False Then
                '文件复制失败
                dImgs.Clear
                Set dImgs = Nothing
                 
                Set ReadDicomFile = Nothing
                Exit Function
            End If
        End If
    
        If blnUseUrl Then
            'readurl不支持空格
            Set curImage = dImgs.ReadURL(strCopyFileName)
        Else
            Set curImage = dImgs.ReadFile(strCopyFileName)
        End If
        
        If curImage Is Nothing Or err.Number <> 0 Then
            err.Clear
            Set curImage = WaitReadDcm(dImgs, strCopyFileName, strError)
        End If
        
        If err.Number = 0 Then
            Call Kill(strCopyFileName)
            '使用ReadFile方式读取的文件进行删除时，可能会产生异常
            err.Clear
        Else
            Call Kill(strCopyFileName)
        End If
    Else
        If blnIsDcmFormat = False Then
            err.Clear
            Set curImage = dImgs.AddNew
            Call curImage.FileImport(strFile, "JPG")
            
            If err.Number <> 0 Then
                err.Clear
                'not a JPG file
                Call curImage.FileImport(strFile, "BMP")
            End If
            
            If err.Number <> 0 Then
                '文件打开异常时，删除添加的item项
                Call dImgs.Remove(dImgs.Count)
            End If
        Else
            '指定读取dicom文件是，需要进行如下容错处理
            If err.Number <> 0 Or curImage Is Nothing Then
                Set curImage = WaitReadDcm(dImgs, strFile, strError)
                
                If Not curImage Is Nothing Then err.Clear
            End If
        End If
    End If
    
    dImgs.Clear
    
    Set dImgs = Nothing
    Set ReadDicomFile = Nothing
    
    If err.Number = 0 Then
        If curImage Is Nothing Then
            strError = "文件格式读取错误"
            Exit Function
        End If
        
        If Len(curImage.InstanceUID) <= 0 Then
            strError = "DICOM格式文件错误,未获取到实例UID"
            Exit Function
        End If
        
        Set ReadDicomFile = curImage
    Else
        strError = err.Description
    End If
    
End Function

Private Function WaitCopy(ByVal strSourceFile As String, ByVal strTargetFile As String, _
    ByRef strError As String, Optional ByVal lngSize As Long = 0) As Boolean
    Dim i As Long
On Error Resume Next
    WaitCopy = False
    
    i = 0
    While True
    
        If IsFileLocked(strTargetFile) = False Then
            Call FileCopy(strSourceFile, strTargetFile)
        
            If IsValidFile(strTargetFile, lngSize) <> "" Then
                WaitCopy = True
                Exit Function
            End If
        End If
        
        i = i + 1
        
        If i > 300 Then
            strError = err.Description
            Exit Function
        End If
        Sleep 10
    Wend
    
End Function


Public Function HasProcess(ByVal strAppTitle As String)
    Dim lngDeskTopHandle As Long
    Dim lngHand As Long
    Dim strName As String * 255
    Dim strCurAppName As String
    
On Error GoTo errhandle
    lngDeskTopHandle = GetDesktopWindow()
    lngHand = GetWindow(lngDeskTopHandle, GW_CHILD)
    
    Do While lngHand <> 0
       GetWindowText lngHand, strName, Len(strName)
       lngHand = GetWindow(lngHand, GW_HWNDNEXT)
       If Left$(strName, 1) <> vbNullChar Then
          strCurAppName = Left$(strName, InStr(1, strName, vbNullChar) - 1)
          If UCase(strCurAppName) = strAppTitle Then
              HasProcess = True
              Exit Function
          End If
       End If
    Loop
     
    HasProcess = False
Exit Function
errhandle:
    HasProcess = False
End Function


Public Function FormatFilePath(ByVal strFilePath As String) As String
'格式化文件路径
    FormatFilePath = Replace(strFilePath, "\\", "\")
End Function


Public Function GetImgCmdPath(Optional ByVal blnIsFailed = False) As String
    Dim strPath As String
    
    If blnIsFailed Then
        strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\TransCmd\Failed\")
    Else
        strPath = FormatFilePath(SysRootPath & "\Apply\TmpImage\TransCmd\")
    End If
    
    If DirExists(strPath) = False Then
        Call MkLocalDir(strPath)
    End If
    
    GetImgCmdPath = strPath
End Function


Public Function GetImgCmdFile(objImgInfo As clsBgImgInfo) As String
    Dim strPath As String
    
    strPath = GetImgCmdPath
    
    GetImgCmdFile = FormatFilePath(strPath & objImgInfo.Key)
End Function

Public Function GetImgCmdFailed(objImgInfo As clsBgImgInfo) As String
    Dim strPath As String
    
    strPath = GetImgCmdPath(True)
    
    GetImgCmdFailed = FormatFilePath(strPath & objImgInfo.Key)
End Function


Public Sub SetFileHide(ByVal strFile As String)
    Dim dwAtrr   As Long
    
    '先获取原来的文件属性
    dwAtrr = GetFileAttributes(strFile)
    '加上隐藏属性
    dwAtrr = dwAtrr Or FILE_ATTRIBUTE_HIDDEN
    '去除隐藏属性
    'dwAtrr = dwAtrr And Not FILE_ATTRIBUTE_HIDDEN
    '设置新的文件属性
    Call SetFileAttributes(strFile, dwAtrr)
End Sub


Private Sub HideFile(ByVal strFile As String)
    Dim oFileSystem As New FileSystemObject
    Dim oFile As File
    
    Set oFile = oFileSystem.GetFile(strFile)
    
    oFile.Attributes = 2
    
    Set oFile = Nothing
    Set oFileSystem = Nothing
End Sub


Public Function TransCmd(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
    Dim blnStartState As Boolean
On Error GoTo errhandle
    TransCmd = False
    strError = ""
    
    '恢复状态属性
    imgInfo.Redo = 0
    imgInfo.ErrorInfo = ""
    imgInfo.StartTime = Now
    imgInfo.EndTime = 0
    
    TransCmd = CreateCmdFileEx(imgInfo, strCmdFile, strError)
    If TransCmd = False Then Exit Function
    
    imgInfo.LoadState = lsSent
 
    TransCmd = True
Exit Function
errhandle:
    strError = err.Description
End Function

Private Function CreateCmdFile(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
'创建数据交换命令文件
'数据交换通过Ini文件进行，只有隐藏后的文件才允许被读取，文件名以key命名，文件目录为TransCmd
    Dim strFile As String
    Dim objIni As New clsIniFile
    
On Error GoTo errhandle
    CreateCmdFile = False
    
    strFile = strCmdFile
    
    If Trim(Dir(strFile, 7)) <> "" Then
        '命令已经存在则直接退出
        CreateCmdFile = True
        Exit Function
    End If
    
    Call objIni.SetIniFile(strFile)
    
    With imgInfo
        objIni.WriteValue "BASEINFO", "KEY", .Key
        objIni.WriteValue "BASEINFO", "FILENAME", .Filename
        objIni.WriteValue "BASEINFO", "FILEPATH", .FilePath
        objIni.WriteValue "BASEINFO", "FORMAT", .Format
        objIni.WriteValue "BASEINFO", "PATIENTNAME", .PatientName
        objIni.WriteValue "BASEINFO", "ADVICEID", .AdviceId
        objIni.WriteValue "BASEINFO", "ADVICEDES", .AdviceDes
        
        objIni.WriteValue "FTPINFO", "FTPIP", .FtpIp
        objIni.WriteValue "FTPINFO", "FTPPORT", .FtpPort
        objIni.WriteValue "FTPINFO", "FTPUSER", .FtpUser
        objIni.WriteValue "FTPINFO", "FTPPWD", .FtpPwd
        objIni.WriteValue "FTPINFO", "FTPVIRTUALPATH", .FtpVirtualPath
        objIni.WriteValue "FTPINFO", "FTPSHDIR", .FtpShareDir
        objIni.WriteValue "FTPINFO", "FTPSHUSER", .FtpShareUser
        objIni.WriteValue "FTPINFO", "FTPSHPWD", .FtpSharePwd
        objIni.WriteValue "FTPINFO", "FTPFILE", .FtpFile
         
        objIni.WriteValue "OTHERINFO", "IMGCOMMAND", .ImgCommand
        objIni.WriteValue "OTHERINFO", "STARTTIME", Now
        objIni.WriteValue "OTHERINFO", "ENDTIME", 0
        objIni.WriteValue "OTHERINFO", "REDO", 0
        
        objIni.WriteValue "OTHERINFO", "ISCOMPRESS", CStr(.IsCompress)
        objIni.WriteValue "OTHERINFO", "JPGCONVERT", CStr(.JpgConvert)
    End With
    
    Set objIni = Nothing
    
    '隐藏文件
'    HideFile strFile   '该方法会造成进程处理卡顿
    Call SetFileHide(strFile)
     
    CreateCmdFile = True
Exit Function
errhandle:
    CreateCmdFile = False
    strError = err.Description
End Function


Private Function CreateCmdFileEx(imgInfo As clsBgImgInfo, ByVal strCmdFile As String, ByRef strError As String) As Boolean
'创建数据交换命令文件
'数据交换通过Ini文件进行，只有隐藏后的文件才允许被读取，文件名以key命名，文件目录为TransCmd
    Dim strFile As String
'    Dim objIni As New clsIniFile
    Dim strIniContext As String
    
On Error GoTo errhandle
    CreateCmdFileEx = False
    
    strFile = strCmdFile
    
    If Trim(Dir(strFile, 7)) <> "" Then
        '命令已经存在则直接退出
        CreateCmdFileEx = True
        Exit Function
    End If
    
'    Call objIni.SetIniFile(strFile)
    
    With imgInfo
        strIniContext = "[BASEINFO]" & vbCrLf & _
                                    "KEY=" & .Key & vbCrLf & _
                                    "FILENAME=" & .Filename & vbCrLf & _
                                    "FILEPATH=" & .FilePath & vbCrLf & _
                                    "FORMAT=" & .Format & vbCrLf & _
                                    "PATIENTNAME=" & .PatientName & vbCrLf & _
                                    "ADVICEID=" & .AdviceId & vbCrLf & _
                                    "ADVICEDES=" & .AdviceDes & vbCrLf & _
                                    "[FTPINFO]" & vbCrLf & _
                                    "FTPIP=" & .FtpIp & vbCrLf & _
                                    "FTPPORT=" & .FtpPort & vbCrLf & _
                                    "FTPUSER=" & .FtpUser & vbCrLf & _
                                    "FTPPWD=" & .FtpPwd & vbCrLf & _
                                    "FTPVIRTUALPATH=" & .FtpVirtualPath & vbCrLf & _
                                    "FTPSHDIR=" & .FtpShareDir & vbCrLf & _
                                    "FTPSHUSER=" & .FtpShareUser & vbCrLf & _
                                    "FTPSHPWD=" & .FtpSharePwd & vbCrLf & _
                                    "FTPFILE=" & .FtpFile & vbCrLf & _
                                    "[OTHERINFO]" & vbCrLf & _
                                    "IMGCOMMAND=" & .ImgCommand & vbCrLf & _
                                    "STARTTIME=" & Now & vbCrLf & _
                                    "ENDTIME=0" & vbCrLf & _
                                    "REDO=0" & vbCrLf & _
                                    "ISCOMPRESS=" & CStr(.IsCompress) & vbCrLf & _
                                    "JPGCONVERT=" & CStr(.JpgConvert)
    End With
    
    Call WritTextFile(strFile, strIniContext)
'    Set objIni = Nothing
    
    '隐藏文件
'    HideFile strFile   '该方法会造成进程处理卡顿
    Call SetFileHide(strFile)
     
    CreateCmdFileEx = True
Exit Function
errhandle:
    CreateCmdFileEx = False
    strError = err.Description
End Function

Public Sub DrawErrorText(objImg As DicomImage, ByVal strError As String)
'绘制错误文本
    Dim i As Long
    Dim objLabInfo As DicomLabel
    
    Set objLabInfo = Nothing
    
    For i = 1 To objImg.Labels.Count
        If objImg.Labels(i).tag = IMG_LAB_ERRORINFO_TAG Then
            Set objLabInfo = objImg.Labels(i)
        End If
    Next
    
    If objLabInfo Is Nothing Then
        Set objLabInfo = objImg.Labels.AddNew
        objLabInfo.tag = IMG_LAB_ERRORINFO_TAG
    End If
    
    'Text*********************************************
    objLabInfo.LabelType = doLabelText
    objLabInfo.Margin = 0
    objLabInfo.FontSize = 10
    objLabInfo.AutoSize = True
    
    objLabInfo.Text = "●" & strError
    objLabInfo.ForeColour = vbYellow ' vbRed
'    objLabInfo.BackColour = vbYellow
    
    objLabInfo.Transparent = True
    objLabInfo.ScaleWithCell = False
     
    objLabInfo.Left = 40
    objLabInfo.Top = 2
'    objLabInfo.Width = 1000
'    objLabInfo.Height = 1000
    
    objLabInfo.Visible = True
    
    Call objImg.Refresh(False)
End Sub

Public Sub DrawErrorInfo(objImg As DicomImage, objImgInfo As clsBgImgInfo, _
    Optional ByVal blnIsClear As Boolean = False)
'绘制错误信息
    Dim i As Long
    Dim objLabInfo As DicomLabel
    Dim objLabState As DicomLabel
    Dim lngLabIndex As Long
    Dim lngStateIndex As Long
    Dim lngBackIndex As Long
    Dim strErrorHint As String
    
    Set objLabInfo = Nothing
    Set objLabState = Nothing
    
    For i = 1 To objImg.Labels.Count
        If objImg.Labels(i).tag = IMG_LAB_ERRORINFO_TAG Then
            Set objLabInfo = objImg.Labels(i)
            lngLabIndex = i
        End If
        
        If objImg.Labels(i).tag = IMG_LAB_ERRORSTATE_TAG Then
            Set objLabState = objImg.Labels(i)
            lngStateIndex = i
        End If
    Next
    
    If blnIsClear Then
        If Not objLabInfo Is Nothing Then
            Call objImg.Labels.Remove(lngLabIndex)
        End If
        
        If Not objLabState Is Nothing Then
            Call objImg.Labels.Remove(lngLabIndex)
        End If
        
        If Not objLabInfo Is Nothing Or Not objLabState Is Nothing Then 'Or Not objLabBack Is Nothing
            Call objImg.Refresh(False)
        End If
        
        Exit Sub
    End If
    
    If objLabInfo Is Nothing Then
        Set objLabInfo = objImg.Labels.AddNew
        objLabInfo.tag = IMG_LAB_ERRORINFO_TAG
    End If
    
    If objImgInfo.Redo > 0 Then
        strErrorHint = "[第" & objImgInfo.Redo & "次" & IIf(objImgInfo.ImgCommand = icUpLoad, "上传", "下载") & "]" & vbCrLf & objImgInfo.ErrorInfo
    Else
        strErrorHint = objImgInfo.ErrorInfo
    End If
    
    '如果文本内容相同，则不需要刷新
    If objLabInfo.Text = strErrorHint Then
        If (objImgInfo.LoadState = lsError) And Not (objLabState Is Nothing) Then Exit Sub
        If (objImgInfo.LoadState = lsRedo) Then Exit Sub
    End If
    
    If objImgInfo.LoadState = lsError Then
        If objLabState Is Nothing Then
            Set objLabState = objImg.Labels.AddNew
            objLabState.tag = IMG_LAB_ERRORSTATE_TAG
        End If
        
        objLabState.LabelType = doLabelText
        objLabState.FontSize = 10
'        objLabState.FontName = "黑体"
        objLabState.Font.Bold = True
        objLabState.AutoSize = False
       
        objLabState.Text = "!!" '"×"
        objLabState.Font.Bold = True
        objLabState.ForeColour = vbRed
        objLabState.BackColour = vbWhite
        objLabState.Shadow = doShadowAll
    
        objLabState.Transparent = False
        objLabState.ScaleWithCell = False
        objLabState.ImageTied = False
         
        objLabState.Left = 0
        
        
        objLabState.Top = 20 + (Len(objImgInfo.DrawHint) + 1) * 20 ' 21
       
       objLabState.Visible = True
    End If
    
    'Text*********************************************
    objLabInfo.LabelType = doLabelText
    objLabInfo.Margin = 0
    objLabInfo.FontSize = 10
    objLabInfo.AutoSize = True
    
    objLabInfo.Text = strErrorHint
    objLabInfo.ForeColour = vbRed
    objLabInfo.BackColour = vbYellow
    
    objLabInfo.Transparent = False
    objLabInfo.ScaleWithCell = False
     
    objLabInfo.Left = 40
    objLabInfo.Top = 1
    
    objLabInfo.Visible = True
    
    Call objImg.Refresh(False)
End Sub



Public Sub DrawBorder(objDcmImg As DicomImage, ByVal lngSelColorStyle As ColorConstants, _
    Optional ByVal blnIsSel As Boolean = False)
    Dim lngColor As OLE_COLOR
    
On Error GoTo errhandle
    lngColor = IMG_BACK_BORDER_COLOR
    If blnIsSel Then lngColor = lngSelColorStyle
 
    objDcmImg.BorderStyle = 0
    objDcmImg.BorderWidth = 1
    objDcmImg.BorderColour = lngColor
    
Exit Sub
errhandle:

End Sub


Public Sub DrawImgOrder(objDcmImg As DicomImage)
'绘制图像序号
    Dim objImgInfo As clsBgImgInfo
    Dim objLabOrder As DicomLabel
    
    Set objImgInfo = objDcmImg.tag
    If objImgInfo Is Nothing Then Exit Sub
    
    Set objLabOrder = objDcmImg.Labels.AddNew
    
    With objLabOrder
        .LabelType = doLabelText
        .tag = IMG_LAB_ORDER_TAG
    
        .FontSize = 9
        .AutoSize = False
        .Shadow = doShadowAll
 
        If objImgInfo.ImageOrder <= 0 Then
            .Text = "**"
        Else
            .Text = IIf(nvl(objImgInfo.SeriesNoTag) <= 1, "", objImgInfo.SeriesNoTag & "-") & IIf(objImgInfo.ImageOrder < 10, "0", "") & objImgInfo.ImageOrder
        End If
        
        .Font.Bold = True
        .ForeColour = &H40C0&
        .BackColour = vbWhite

        .Transparent = False
        .ScaleWithCell = False
     
        .Left = 0
        .Top = 21  '20 + 1 * 20
       
       
        .Visible = True
    End With
    
End Sub


Public Sub DrawCheckBox(objDcmImg As DicomImage, ByVal lngSelColorStyle As ColorConstants, _
    Optional ByVal blnIsSel As Boolean = False)
    Dim lSelect As DicomLabel
    Dim i As Long
    
On Error GoTo errhandle
    
    For i = 1 To objDcmImg.Labels.Count
        If objDcmImg.Labels(i).tag = IMG_LAB_CHECKBOX_TAG Then
            Set lSelect = objDcmImg.Labels(i)
            Exit For
        End If
    Next
 
    If lSelect Is Nothing Then
        Set lSelect = objDcmImg.Labels.AddNew
    Else
        If lSelect.Transparent = Not blnIsSel Then Exit Sub
    End If

    With lSelect
        .LabelType = doLabelRectangle            '矩形
        .Width = 18
        .Height = 18
        .Margin = 4
        .Left = 1
        .Top = 1
        .LineWidth = 2
        
'        .ForeColour = vbYellow
'        .BackColour = vbRed
        .ForeColour = CLng(&HC0C0C0)
        .BackColour = lngSelColorStyle 'CLng(&H8000000F)
        
        .Transparent = Not blnIsSel
        .ScaleWithCell = False
'        .ImageTied = False
    
        .tag = IMG_LAB_CHECKBOX_TAG
        
        .Visible = True
    End With
    
    Call objDcmImg.Refresh(False)
Exit Sub
errhandle:

End Sub


Public Sub DrawHints(objDcmImg As DicomImage)
    Dim i As Long
    Dim strHint As String
    Dim strChar As String
    
On Error GoTo errhandle
    
    strHint = objDcmImg.tag.DrawHint
    For i = 1 To Len(strHint)
        strChar = Mid(strHint, i, 1)
        Call DrawHint(objDcmImg, strChar)
    Next

Exit Sub
errhandle:


End Sub

Public Sub DrawHint(objDcmImg As DicomImage, ByVal strChar As String, _
    Optional ByVal blnIsClear As Boolean = False)
    Dim i As Long
    Dim objLabHint As DicomLabel
    Dim lngLabIndex As Long
    Dim lngHintIndex As Long
    
On Error GoTo errhandle
    
    Set objLabHint = Nothing
    lngHintIndex = 0
    
    For i = 1 To objDcmImg.Labels.Count
        If objDcmImg.Labels(i).tag = IMG_LAB_HINT_TAG Then
            lngHintIndex = lngHintIndex + 1
            If objDcmImg.Labels(i).Text = strChar Then
                Set objLabHint = objDcmImg.Labels(i)
                lngLabIndex = i
                
                Exit For
            End If
        End If
    Next
    
    If blnIsClear Then
        If objLabHint Is Nothing Then Exit Sub
        
            Call objDcmImg.Labels.Remove(lngLabIndex)
        
            '需要判断是否调用refresh
        Exit Sub
    End If
    
    If Not objLabHint Is Nothing Then Exit Sub  '已经绘制则退出
    
    Set objLabHint = objDcmImg.Labels.AddNew
    
    objLabHint.LabelType = doLabelText
    objLabHint.tag = IMG_LAB_HINT_TAG
    
    objLabHint.FontSize = 10
    objLabHint.AutoSize = False
    objLabHint.Shadow = doShadowAll
 
    objLabHint.Text = strChar
    objLabHint.Font.Bold = True
    objLabHint.ForeColour = vbBlue
    objLabHint.BackColour = vbWhite

    objLabHint.Transparent = False
    objLabHint.ScaleWithCell = False
     
    objLabHint.Left = 0
    objLabHint.Top = 20 + (lngHintIndex + 1) * 20
       
       
    objLabHint.Visible = True
Exit Sub
errhandle:

End Sub


Public Sub DrawMarks(img As DicomImage, thisMarks As clsPicMarks, ByVal dblMarkZoom As Double)
'------------------------------------------------
'功能：显示标注，支持数字编号，箭头，圆形，文字标注
'参数：
'返回：无
'------------------------------------------------
    Dim i As Integer
    Dim oneLabel As DicomLabel

    On Error GoTo err

    img.Labels.Clear
    'thisMarks(i).类型定义 '0-文本,1-线条,2,折线,3-矩形,4-多边形,5-圆(椭圆), 6-顺序编号，7-箭头（PACS中增加）
    For i = 1 To thisMarks.Count
        With thisMarks(i)
            If thisMarks(i).类型 = 0 Then       '文本
                img.Labels.Add GetNewLabel(doLabelText, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.Font.Bold = True
                oneLabel.Text = .内容
            ElseIf thisMarks(i).类型 = 5 Then   '椭圆
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, (.X2 - .X1) * dblMarkZoom, (.Y2 - .Y1) * dblMarkZoom)
            ElseIf thisMarks(i).类型 = 6 Then   '顺序编号
                '圆形背景色
                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom - 8, .Y1 * dblMarkZoom - 8, 17, 17)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.XOR = False
                oneLabel.BackColour = IIf(.填充色 = 0, vbYellow, .填充色)
                oneLabel.Transparent = False
'                oneLabel.tag = m_LabelTag_Back

                '圆形框
'                img.Labels.Add GetNewLabel(doLabelEllipse, .X1 * dblMarkZoom - 8, .Y1 * dblMarkZoom - 8, 17, 17)
'                Set oneLabel = img.Labels(img.Labels.Count)
'                oneLabel.XOR = False
                oneLabel.ForeColour = vbBlack
'                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Circle
'                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)


                '圆形编号数字
                img.Labels.Add GetNewLabel(doLabelText, .X1 * dblMarkZoom - 8 + 2, .Y1 * dblMarkZoom - 8 + 2, 0, 0)
                Set oneLabel = img.Labels(img.Labels.Count)
                oneLabel.ForeColour = vbBlack
                oneLabel.XOR = False
                oneLabel.Transparent = True
                oneLabel.tag = m_LabelTag_Number
                oneLabel.FontSize = 8
                oneLabel.FontName = "Arial Bold"
                oneLabel.AutoSize = True
                oneLabel.Text = .内容
                If Val(.内容) < 10 Then  '10以下的数字，需要微调一下位置，数字才能出现在圆圈的正中间
                    oneLabel.Left = oneLabel.Left + 3
                End If

                oneLabel.TagObject = img.Labels(img.Labels.Count - 1)
                img.Labels(img.Labels.Count - 1).TagObject = oneLabel  'TagObject形成闭环  'img.Labels(img.Labels.Count - 2).TagObject = oneLabel

            ElseIf thisMarks(i).类型 = 7 Then   '箭头
                img.Labels.Add GetNewLabel(doLabelArrow, .X1 * dblMarkZoom, .Y1 * dblMarkZoom, (.X2 - .X1) * dblMarkZoom, (.Y2 - .Y1) * dblMarkZoom)
            End If
        End With
    Next i
    
    Call img.Refresh(False)
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume Next
End Sub
