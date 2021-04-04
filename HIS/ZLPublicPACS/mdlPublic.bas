Attribute VB_Name = "mdlPublic"
Option Explicit

Public gobjComLib As Object    'zl9ComLib.clsComLib
Public gcnOracle As ADODB.Connection
Public gcnOledb As ADODB.Connection
Public gstrPrivs As String
Public gstrSysName  As String
Public gstrDBUser As String
Public gstrSQL As String
Private mclsUnzip As Object
Public gobjPacsCore As Object   'PACS观片对象

Public Const VIEW_ALLREPORT = "全院影像查询"

Public Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Public Declare Function LoadImage Lib "SLInterCOM.dll" (ByVal hWnd As Long, ByVal pType As String, ByVal pStuNO As String, ByVal pParam1 As String, ByVal pParam2 As String, ByVal pParam3 As String) As Long
Public Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type NETRESOURCE ' 网络资源
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Public Const RESOURCETYPE_ANY = &H0

Public Type tFtpInfo
    FtpDir As String
    FtpIP As String
    FtpPswd As String
    FTPUser As String
    DiviceId As String
    
    SubDir As String
    DestMainDir As String
End Type

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO

Public Function GetColNum(listTemp As Object, strHead As String) As Integer
    Dim i As Integer
    Select Case UCase(TypeName(listTemp))
        Case UCase("ReportControl")
            For i = 0 To listTemp.Columns.Count - 1
                If listTemp.Columns.Column(i).Caption = strHead Then GetColNum = listTemp.Columns.Column(i).ItemIndex: Exit Function
            Next
        Case UCase("ListView")
            For i = 1 To listTemp.ColumnHeaders.Count
                If listTemp.ColumnHeaders(i).Text = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("MSHFlexGrid") '以下类型待增，尚未用到
        Case UCase("BillEdit")
        Case UCase("VSFlexGrid")
            For i = 0 To listTemp.Cols - 1
                If listTemp.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
            Next
        Case UCase("BillEdit")
        Case UCase("DataGrid")
    End Select
End Function

Public Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function

Function GetFileContent(ByVal strFileName As String) As String
'读取本地文件内容
    Dim i As Integer, strContent As String, bty() As Byte
    
    If Dir(strFileName) = "" Then Exit Function
    
    i = FreeFile
    
    ReDim bty(FileLen(strFileName) - 1)
    
    Open strFileName For Binary As #i
    Get #i, , bty
    Close #i
    strContent = StrConv(bty, vbUnicode)
    
    GetFileContent = strContent
End Function

Public Sub MkLocalDir(ByVal strDir As String)
'功能：创建本地目录
    Dim objFile As New Scripting.FileSystemObject
    Dim aNestDirs() As String, i As Integer
    Dim strPath As String
    On Error Resume Next
    
    '读取全部需要创建的目录信息
    ReDim Preserve aNestDirs(0)
    aNestDirs(0) = strDir
    
    strPath = objFile.GetParentFolderName(strDir)
    Do While Len(strPath) > 0
        ReDim Preserve aNestDirs(UBound(aNestDirs) + 1)
        aNestDirs(UBound(aNestDirs)) = strPath
        strPath = objFile.GetParentFolderName(strPath)
    Loop
    '创建全部目录
    For i = UBound(aNestDirs) To 0 Step -1
        MkDir aNestDirs(i)
    Next
End Sub

Public Sub ClearCacheFolder(ByVal strCacheFolder As String)
'功能：当指定目录的大小达到一定百分比时，清空该目录
    Dim objFile As New Scripting.FileSystemObject
    Dim objCurFolder As Scripting.Folder, objCurFile As Scripting.File, objFiles As Scripting.Files
    Dim strDriver As String
    
    On Error Resume Next
    strDriver = objFile.GetDriveName(strCacheFolder)
    Set objCurFolder = objFile.GetFolder(strCacheFolder)
    If objCurFolder.Size / objFile.GetDrive(strDriver).FreeSpace > 0.2 Then
        objCurFolder.Delete True
    End If
End Sub

Public Function funcConnectShardDir(strShareRemoteDir As String, strUserName As String, strPassWord As String) As Long
    '创建网络资源
    Dim NetR As NETRESOURCE
    Dim lngResult As Long
    
    NetR.dwType = RESOURCETYPE_ANY
    NetR.lpLocalName = vbNullString
    NetR.lpRemoteName = strShareRemoteDir
    NetR.lpProvider = vbNullString
    lngResult = WNetAddConnection2(NetR, strPassWord, strUserName, 0)
    
    If lngResult <> 0 Then
        MsgBox "网络连接失败，请检查网络设置是否正确！"
    End If
    funcConnectShardDir = lngResult
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetAdviceID(ByVal lngReportID As Long) As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    strSql = "select 医嘱ID from 病人医嘱报告 where 病历ID =[1]"
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", lngReportID)
    
    If rsData.RecordCount > 0 Then GetAdviceID = Val(Nvl(rsData!医嘱ID))
End Function

Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = IIf(IsNull(rsTmp!用户名), "", rsTmp!用户名)
        GetUserInfo = True
    End If
End Function

Public Function View3DImage(ByVal lng医嘱ID As Long, frmParent As Object) As Long
    Dim blnCanViewImage As Boolean  '该医嘱的报告还没有完成(没有正式签名或完成执行)时，是否可以观片
    Dim lngResut As Long
    Dim str3DViewType As String
    Dim intImageLocation As Long    '图像位置，有三种情况：1、旧版PACS；2、旧版RIS+新版PACS；3、新版RIS+PACS
    
    On Error GoTo DBError
    
    If getImageLocation(lng医嘱ID, intImageLocation, blnCanViewImage) = False Then Exit Function
    
    str3DViewType = gobjComLib.zlDatabase.GetPara("XW3D观片类型", 100, 1288, "Study3D")
    If Trim(str3DViewType) = "" Then str3DViewType = "Study3D"
    
    lngResut = LoadImage(0, str3DViewType, CStr(lng医嘱ID), "", "", "")
    
    If lngResut = -121 Then
        MsgBox "调用参数错误", vbInformation, gstrSysName
    ElseIf lngResut = -122 Or lngResut = -102 Then
        MsgBox lngResut & ":未正确安装PACS及接口文件", vbInformation, gstrSysName
    ElseIf lngResut = -108 Then
        MsgBox lngResut & ":网络连接错误", vbInformation, gstrSysName
    ElseIf lngResut = -104 Then
        MsgBox lngResut & ":数据库错误", vbInformation, gstrSysName
    ElseIf lngResut = -101 Then
        MsgBox lngResut & ":其他错误", vbInformation, gstrSysName
    End If
    
    View3DImage = lngResut
    
    Exit Function
DBError:
    lngResut = -1
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ViewStaticImage(ByVal lng医嘱ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
'功能：调用观片站
    Dim intImageLocation As Long
    Dim blnCanViewImage As Boolean  '该医嘱的报告还没有完成(没有正式签名或完成执行)时，是否可以观片
    
    On Error GoTo DBError
    
    If getImageLocation(lng医嘱ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    '图像在新网数据库，则调用新网的WEB浏览
    If intImageLocation = 1 Then
        Call XWWebViewerStaticOpen(lng医嘱ID)
    End If
    
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ViewPatientImage(ByVal lng医嘱ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据医嘱ID，查找到患者的所有病人ID，打开专业版PACS观片
'参数：lng医嘱ID--病人医嘱ID编号
'       frmParent -- 父窗体
'       blnMoved -- 是否转储
'返回：无
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intImageLocation As Long
    Dim blnCanViewImage As Boolean  '该医嘱的报告还没有完成(没有正式签名或完成执行)时，是否可以观片
    
    On Error GoTo DBError
    
    If getImageLocation(lng医嘱ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    '图像在新网数据库，则调用新网的WEB浏览
    If intImageLocation = 1 Then
        Call XWWebViewerPatientOpen(lng医嘱ID)
    End If
    
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Sub ViewImage(ByVal lng医嘱ID As Long, frmParent As Object, Optional ByVal blnMoved As Boolean = False, Optional ByVal strPrivs As String = "")
'功能：调用观片站
    Dim strFtpHost As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    Dim intImageLocation As Long
    Dim lng报告ID As Long
    Dim blnCanViewImage As Boolean  '该医嘱的报告还没有完成(没有正式签名或完成执行)时，是否可以观片
    
    On Error GoTo DBError
    
    If getImageLocation(lng医嘱ID, intImageLocation, blnCanViewImage, blnMoved) = False Then Exit Sub
    
    '图像在新网数据库，则调用新网的WEB浏览
    If intImageLocation = 1 Or intImageLocation = 2 Then
        Call XWWebViewerOpen(lng医嘱ID)
        
        If intImageLocation = 2 Then
            Call XWDownLoadImage(lng医嘱ID)
        End If
        
        Exit Sub
    End If
    
    
    '先判断是否存在图像，没有图像则提示并退出
    strSql = "Select A.检查UID,Count(B.序列UID) as 序列总数 From 影像检查记录 A,影像检查序列 B Where A.检查UID=B.检查UID And A.医嘱ID=[1] Group by A.检查UID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "观片处理", lng医嘱ID)
    If rsTmp.EOF Then
        MsgBox "没有可用于观片的报告图像。", vbInformation, gstrSysName
        Exit Sub
    End If

    strFtpHost = ""
    
    '查找需要打开的所有图象信息
    strSql = "Select /*+RULE*/ D.IP地址 As Host1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'\')" & _
        "||C.检查UID||'\' As Path,E.IP地址 As Host2,e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And C.医嘱ID=[1] "
        
    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取共享目录信息", lng医嘱ID)
    
    If rsTmp.RecordCount > 0 Then
        '创建本地的缓存目录，需要在调用观片站之前先创建这个目录，观片站中只是下载，不创建本地缓存目录
        MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        ClearCacheFolder App.Path & "\TmpImage\"
        
        '读取FTP参数，包括用户名，密码，IP地址等
        If rsTmp("设备号1") <> "" Then
            strFtpHost = rsTmp("Host1")
            strSDPath = Nvl(rsTmp("共享目录1"))
            strSDUser = Nvl(rsTmp("共享目录用户名1"))
            strSDPwd = Nvl(rsTmp("共享目录密码1"))
        ElseIf Nvl(rsTmp("设备号2")) <> "" Then
            strFtpHost = rsTmp("Host2")
            strSDPath = Nvl(rsTmp("共享目录2"))
            strSDUser = Nvl(rsTmp("共享目录用户名2"))
            strSDPwd = Nvl(rsTmp("共享目录密码2"))
        End If
        
        '判断共享目录是否已经连接，如果没有连接，则进行连接
        On Error Resume Next
        If strSDPath <> "" Then
            Call funcConnectShardDir("\\" & strFtpHost & "\" & strSDPath, strSDUser, strSDPwd)
        End If
        
        If gobjPacsCore Is Nothing Then
            Set gobjPacsCore = CreateObject("zl9PacsCore.clsViewer")
        End If
        gobjPacsCore.CallOpenViewer "", lng医嘱ID, frmParent, gcnOracle, blnMoved, False
        
    End If

    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function CheckEPRReport(ByVal lng医嘱ID As Long, Optional lng报告ID As Long, Optional blnBySign As Boolean, Optional ByVal int执行状态 As Integer = -999) As Integer
'功能：检查对应项目的报告填写情况
'参数：lng医嘱ID=可见行的医嘱ID
'      lng报告ID=可以传入，主要用于返回报告病历ID
'      int执行状态=用于检验完成时，传入综合的执行状态
'参数：blnBySign=报告是否完成通过签名级别判断(用于医技工作站)
'返回：0-报告还没有填写
'      1-报告已填写完成(已签名,包括修订后签名,或已执行完成)
'      2-报告未填写完成(未签名,或修订后未签名,且未执行完成)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo ErrH
    
    '检查报告是否已书写
    If lng报告ID = 0 Then
        strSql = "Select 病历ID From 病人医嘱报告 Where 医嘱ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng医嘱ID)
        If Not rsTmp.EOF Then lng报告ID = rsTmp!病历id
    End If
    If lng报告ID = 0 Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
        '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
        strSql = _
            " Select 2 as 排序,医嘱ID,执行过程,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
            " Union ALL" & _
            " Select 排序,医嘱ID,执行过程,Decode([2],-999,执行状态,[2]) as 执行状态,发送时间" & _
            " From (" & _
                " Select 1 as 排序,B.医嘱ID,B.执行过程,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
                " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                    " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
                " Order by A.序号" & _
            " ) Where Rownum=1" & _
            " Order by 排序,发送时间 Desc"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng医嘱ID, int执行状态)
        If Nvl(rsTmp!执行过程, 0) >= 5 Or Nvl(rsTmp!执行状态, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        '通过签名版本判断报告完成的方式
        strSql = "Select B.文件ID,Max(B.开始版) as 签名版本 From 电子病历内容 B Where B.文件ID=[1] And B.对象类型=8 Group by B.文件ID"
        strSql = "Select B.完成时间,B.最后版本,C.签名版本 From 电子病历记录 B,(" & strSql & ") C Where B.ID=[1] And B.ID=C.文件ID(+)"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng报告ID)
            
        '(签名后不能直接修改，除非修订；因此签名后最后版本应与签名版本一致)
        If IsNull(rsTmp!完成时间) Or Nvl(rsTmp!最后版本, 0) <> Nvl(rsTmp!签名版本, 0) Then
            '如果医嘱本身已经执行,即使没有签名或不符也视同完成
            strSql = _
                " Select 2 as 排序,医嘱ID,执行状态,发送时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
                " Union ALL" & _
                " Select 排序,医嘱ID,Decode([2],-999,执行状态,[2]) as 执行状态,发送时间" & _
                " From (" & _
                    " Select 1 as 排序,B.医嘱ID,B.执行状态,B.发送时间 From 病人医嘱记录 A,病人医嘱发送 B" & _
                    " Where A.ID=B.医嘱ID And A.相关ID=(" & _
                        " Select A.ID From 病人医嘱记录 A,诊疗项目目录 B Where A.ID=[1] And A.诊疗项目ID=B.ID And A.诊疗类别='E' And B.操作类型='6')" & _
                    " Order by A.序号" & _
                " ) Where Rownum=1" & _
                " Order by 排序,发送时间 Desc"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "CheckEPRReport", lng医嘱ID, int执行状态)
            If Nvl(rsTmp!执行状态, 0) = 1 Then
                CheckEPRReport = 1
            Else
                CheckEPRReport = 2
            End If
        Else
            CheckEPRReport = 1
        End If
    End If
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function

Private Function XWDownLoadImage(lngOrderID As Long) As Long
''--------------------------------------------
''功能： 从云平台下载图像
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------

    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strStudyUID As String
    
    On Error GoTo err
    strSql = "select 检查UID from 影像检查记录 where 医嘱ID = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取检查UID", lngOrderID)
    If rsTemp.EOF = True Then Exit Function
    
    strStudyUID = Nvl(rsTemp!检查UID, "")
    
    '调用新网存储过程“P_OEM_DOWNLOADIMG_RIS”，从云平台下载图像
    strSql = "P_OEM_DOWNLOADIMG_RIS@XWPacs('" & strStudyUID & "')"
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
    
    Call MsgBox("该患者影像图片已经上传云端，需从云端下载，" & vbCrLf & vbCrLf & "此过程需要一定时间，" & vbCrLf & vbCrLf & "如没看到图片说明正在下载，请稍等。", vbOKOnly, "提示信息")
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then Resume
    XWDownLoadImage = 1
End Function

Private Function XWWebViewerOpen(lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开新网的WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XWWEB观片地址", 100, 1288, "")
    
    If strPath <> "" Then
        strPath = Replace(strPath, "[@STU_NO]", lngOrderID)

        '兼容64位的操作系统，XW WEB观片不支持64位IE，所以要使用32位的IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerOpen = 0
    Else
        MsgBox "XWWEB观片地址为空，请先设置好WEB服务器。", vbOKOnly, "提示信息"
        XWWebViewerOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerStaticOpen(lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开新网的WEB Viewer
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XW关键图像地址", 100, 1288, "")
    
    If strPath <> "" Then
        strPath = Replace(strPath, "[@STU_NO]", lngOrderID)
        
        '兼容64位的操作系统，XW WEB观片不支持64位IE，所以要使用32位的IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerStaticOpen = 0
    Else
        MsgBox "XW关键图像地址址为空，请先设置好关键图像地址。", vbOKOnly, "提示信息"
        XWWebViewerStaticOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function XWWebViewerPatientOpen(lngOrderID As Long) As Long
''--------------------------------------------
''功能： 打开新网的WEB Viewer，根据病人ID，显示检查列表后观片
'           lngOrderID -- 医嘱ID
''返回：0-成功;1-出错
''--------------------------------------------
    Dim strPath As String
    Dim strURL As String
    Dim strPatientIDs As String     '传给专业版PACS的病人ID串，格式是“'编号1','编号2'”
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strPath = gobjComLib.zlDatabase.GetPara("XWWeb检查列表观片地址", 100, 1288, "")
    
    If strPath <> "" Then
    
        '根据医嘱ID，提取病人ID串
        strSql = "select 病人ID from 病人信息  where 身份证号=(select 身份证号 from 病人信息 a,病人医嘱记录 b where a.病人id =b.病人id and b.id=[1]) "
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取病人ID串", lngOrderID)
        If rsTemp.EOF = True Then
            strSql = "select 病人ID from 病人医嘱记录  where id=[1] "
            Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "提取病人ID串", lngOrderID)
            If rsTemp.EOF = True Then
                MsgBox "根据医嘱ID " & lngOrderID & " 提取不到病人ID。", vbOKOnly, "zlPublicPACS观片提示"
                Exit Function
            End If
        End If
        
        While rsTemp.EOF = False
            strPatientIDs = strPatientIDs & ",'" & rsTemp!病人ID & "'"
            rsTemp.MoveNext
        Wend
        strPatientIDs = Mid(strPatientIDs, 2)
        
        strPath = Replace(strPath, "[@PAT_NOs]", strPatientIDs)
        
        '兼容64位的操作系统，XW WEB观片不支持64位IE，所以要使用32位的IE
        If Dir("C:\Program Files (x86)\Internet Explorer", vbDirectory) = "" Then
            strURL = "C:\Program Files\Internet Explorer\iexplore.exe " & strPath
        Else
            strURL = "C:\Program Files (x86)\Internet Explorer\iexplore.exe " & strPath
        End If
        
        Shell strURL, vbMaximizedFocus
        XWWebViewerPatientOpen = 0
    Else
        MsgBox "XWWeb检查列表观片地址为空，请先设置观片地址。", vbOKOnly, "提示信息"
        XWWebViewerPatientOpen = 1
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub BlobToFile(fld As ADODB.Field, Filename As String, Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesleft As Long, bytes As Long
    Dim tmp() As Byte
    
    If (fld.Attributes And adFldLong) = 0 Then
        err.Raise 1001, , "field doesn't support the GetChunk method."
    End If
    
    If Dir$(Filename) <> "" Then Kill Filename
    
    fnum = FreeFile
    Open Filename For Binary As fnum
    bytesleft = fld.ActualSize
    Do While bytesleft
        bytes = bytesleft
        If bytes > ChunkSize Then bytes = ChunkSize
        tmp = fld.GetChunk(bytes)
        Put #fnum, , tmp
        bytesleft = bytesleft - bytes
    Loop
    
    Close #fnum
End Sub

Public Function InitOledbConn(Optional ByVal blnUseAlone As Boolean = False) As Boolean
    Dim objRegister As Object
    Dim strError As String

On Error GoTo err

    If blnUseAlone Then
        Set objRegister = GetObject("", "zlRegisterAlone.clsRegister")
    Else
        Set objRegister = GetObject("", "zlRegister.clsRegister")
    End If

    Set gcnOledb = objRegister.ReGetConnection(1, strError)

    InitOledbConn = True
    Exit Function
err:
    InitOledbConn = False

    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetRecordset(ByVal strSql As String) As ADODB.Recordset
On Error GoTo ErrHand
    If gcnOledb Is Nothing Then
        Call InitOledbConn
    End If

    Set GetRecordset = New ADODB.Recordset

    If gcnOledb Is Nothing Then Exit Function

    If GetRecordset.State = adStateOpen Then GetRecordset.Close
    '打开
    GetRecordset.Open strSql, gcnOledb, adOpenKeyset, adLockOptimistic
     
'    Set GetRecordset = gobjComLib.zlDatabase.OpenSQLRecordByArray(strSql, "判断是否有影像图片", Null, 1)
 
    Exit Function
ErrHand:
    If err <> 0 Then
        MsgBox "发生错误：" & err.Description, vbInformation, "系统信息"
    End If
End Function

Public Sub BUGEX(ByVal strDebug As String, Optional ByVal blnIsForce As Boolean = False)
    OutputDebugString Format(Now, "mmddhhmmss") & " |-> " & strDebug
End Sub

Public Function HasImage(lngOrderID As Long) As Boolean
''--------------------------------------------
''功能： 判断该检查是否有图像
'           lngAdviceID -- 医嘱ID
''返回：True-有图像；False-无图像
''--------------------------------------------
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intImageLocation As Integer

    On Error GoTo err
    
    '判断该检查是否有图像，有三种情况：1、旧版PACS；2、旧版RIS+新版PACS；3、新版RIS+PACS
    
    '先查询图像是否在旧版PACS
    strSql = "Select 检查UID,图像位置 From 影像检查记录 Where 医嘱ID =[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "判断是否有影像图片", lngOrderID)
    
    intImageLocation = 0
    If rsTemp.RecordCount > 0 Then
        '图像属于情况1或2
        If Nvl(rsTemp!图像位置, 0) = 0 Then
            '图像在旧版PACS中
            '如果有 检查UID 的记录说明数据库中有图像，则返回True，反之返回false
            HasImage = IIf(Nvl(rsTemp!检查UID, 0) <> 0, True, False)
        Else
            '图像在旧版RIS+新版PACS中
            intImageLocation = 1
        End If
    Else
        '启用255参数，则使用了影像信息系统专业版，图像在新版RIS+PACS中
        If Val(gobjComLib.zlDatabase.GetPara(255, 100)) = 1 Then
            intImageLocation = 1
        End If
    End If
    
    If intImageLocation = 1 Then
        '图像在新版PACS中,根据 执行过程>=3 判断是否有图像
        strSql = "SELECT 医嘱ID from 病人医嘱发送  where 执行过程>=3 and 医嘱ID =[1]"
        Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "判断是否有影像图片", lngOrderID)
        
        If rsTemp.EOF Then
            HasImage = False
        Else
            HasImage = True
        End If
    End If
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isUseXWInterface(strSubName As String) As Boolean
''--------------------------------------------
''功能： 判断是否使用新网RIS
'           strSubName -- 调用的程序名称
''返回：True-使用；False-不使用
''--------------------------------------------
    Dim strUseXWInterface As String
    
    On Error GoTo err
    
    strUseXWInterface = gobjComLib.zlDatabase.GetPara(255, 100)
    
    BUGEX strSubName & ": strUseXWInterface = " & strUseXWInterface
    
    '获取是否启用影像信息系统接口
    isUseXWInterface = Val(strUseXWInterface) = 1
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function getImageLocation(ByVal lng医嘱ID As Long, ByRef intImageLocation As Long, ByRef blnCanViewImage As Boolean, _
    Optional ByVal blnMoved As Boolean = False) As Boolean
''--------------------------------------------
''功能： 判断图像位置，是否可以未审核报告查看图像
''参数：    lng医嘱ID -- 医嘱ID
'           intImageLocation -- 图像位置，有三种情况：1、旧版PACS intImageLocation=0；
'                   2、旧版RIS+新版PACS intImageLocation=1或2(上传到云存储)；3、新版RIS+PACS intImageLocation=1
'           blnCanViewImage -- 该医嘱的报告还没有完成(没有正式签名或完成执行)时，是否可以观片
''返回：True-成功；False-失败
''--------------------------------------------
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim lng执行科室ID As Long
    Dim blnIsGreen As Boolean
    Dim blnIsUrgent As Boolean
    
    On Error GoTo err
    
    lng执行科室ID = 0

    '查询图像位置,以及执行科室ID
    strSql = "Select a.图像位置, a.执行科室id, a.绿色通道, b.紧急标志 From 影像检查记录 a, 病人医嘱记录 b Where a.医嘱id = b.Id And a.医嘱id =[1]"
    
    If blnMoved Then
        strSql = Replace(strSql, "影像检查记录", "H影像检查记录")
        strSql = Replace(strSql, "病人医嘱记录", "H病人医嘱记录")
    End If
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "查询图像所在的位置", lng医嘱ID)
    
    If rsTmp.RecordCount <> 0 Then
        intImageLocation = Nvl(rsTmp!图像位置, 0)
        lng执行科室ID = Val(Nvl(rsTmp!执行科室ID, 0))
        blnIsGreen = IIf(Val(Nvl(rsTmp!绿色通道, 0)) = 1, True, False)
        blnIsUrgent = IIf(Val(Nvl(rsTmp!紧急标志, 0)) = 1, True, False)
    Else
        intImageLocation = 1
    End If
    
    If lng执行科室ID > 0 Then
        '图像存在位置1或2
        strSql = "Select 参数值 from 影像流程参数 where 科室ID = [1] and 参数名=[2]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "读取参数", lng执行科室ID, "采图后医生站即可观片")
        If rsTmp.RecordCount > 0 Then blnCanViewImage = Val(Nvl(rsTmp!参数值, 0)) = 1
    Else
        '图像存在位置3，或者医嘱ID输入错误
        blnCanViewImage = isUseXWInterface("getImageLocation")
    End If
    
    '获取报告状态
    strSql = "Select 执行过程 from 病人医嘱发送 where 医嘱id= " & lng医嘱ID
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "getImageLocation", lng医嘱ID)
    
    If rsTmp.RecordCount > 0 Then
        If blnCanViewImage Then
            '如果报告未完成，并且勾选了参数“采图后医生站即可观片”，单独有“审核前观片”权限时才可进行观片
            '急诊或绿色通道病人不考虑审核前观片权限
            If Nvl(rsTmp!执行过程, 0) < 5 Then
                If InStr(gstrPrivs, "审核前观片") <= 0 And Not (blnIsGreen And blnIsUrgent) Then
                    MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，在没有审核前观片权限时不能查看图像！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            '没有勾选参数“采图后医生站即可观片”时，报告完成后才可进行观片
            If Nvl(rsTmp!执行过程, 0) < 5 Then
                MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，现在不能查看图像！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    getImageLocation = True
    
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function
