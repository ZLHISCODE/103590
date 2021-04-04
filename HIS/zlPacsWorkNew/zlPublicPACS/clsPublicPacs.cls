VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPublicPacs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InitInterface(ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String) As Boolean
'对接口所需数据进行初始化，成功返回TRUE
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    Set gcnOracle = cnOracle
    
    If gcnOracle Is Nothing Then Exit Function
    
    If gobjComLib Is Nothing Then
        Set gobjComLib = CreateObject("zl9ComLib.clsComLib")

        Call gobjComLib.InitCommon(gcnOracle)
        Call gobjComLib.SetDbUser(UCase(strDBUser))
        Call gobjComLib.RegCheck
    End If
    
    InitInterface = True
End Function

Public Function GetReportImage(ByVal lngAdviceId As Long, ByVal strTmpImgDir As String) As Collection
'功能:将检查医嘱所对应报告中的所有图像，保存到指定目录,获取PACS报告中的所有图像
'lngAdviceId:医嘱id
'strTmpImgFolder:图像缓存目录
'objImgFileName:报告图像文件名集合
    Dim objFSO As New Scripting.FileSystemObject
    Dim strSQL As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strTmpImgFile As String
    Dim objImgFileName As New Collection
    Dim objFtp As New clsFtp
    Dim strFtpDir As String
    Dim strIP As String
    Dim strFtpPswd As String
    Dim strFTPUser As String
    Dim strSubDir As String
     
    If lngAdviceId <= 0 Then Exit Function
    
    '创建指定目目录
    If Not objFSO.FolderExists(strTmpImgDir) Then MkLocalDir strTmpImgDir
    
    strSQL = "Select D.设备号,D.FTP用户名,D.FTP密码,D.IP地址,'/'||D.Ftp目录||'/' As Root, " & _
             "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')||C.检查UID As URL,d.设备号 as 设备号 " & _
             "From 影像检查记录 C,影像设备目录 D " & _
             "Where  C.位置一=D.设备号(+) and C.医嘱ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取FTP信息", lngAdviceId)
    
    If rsTemp.EOF Then
        MsgBox "没有报告图像。", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        If Trim(Nvl(rsTemp("设备号"))) = "" Then Exit Function
        
        strFtpDir = Nvl(rsTemp("Root"))
        strIP = Nvl(rsTemp("IP地址"))
        strFtpPswd = Nvl(rsTemp("FTP密码"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strSubDir = Nvl(rsTemp("URL"))
    End If
    
    '连接FTP
    If objFtp.FuncFtpConnect(strIP, strFTPUser, strFtpPswd) = 0 Then
        MsgBox "FTP不能正常连接，请检查网络设置。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    strSQL = "select c.对象属性 " & _
             "from 病人医嘱报告 a,电子病历内容 b,电子病历内容 c " & _
             "Where a.医嘱ID =[1] And a.病历id = b.文件id And b.Id = c.父id " & _
             "and b.对象类型 = 3 And Substr(b.对象属性, Instr(b.对象属性, ';', 1, 18) + 1, 1) = '2' " & _
             "and Nvl(c.终止版, 0) = 0 "

    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "对象属性", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function

    Do While Not rsData.EOF
        strTmpImgFile = Split(rsData!对象属性, ";")(11)
        
        If Trim(strTmpImgFile) <> "" Then
            '将报告图像文件加入集合
            objImgFileName.Add strTmpImgFile
            
            '文件存在则删除
            If objFSO.FileExists(strTmpImgDir & "\" & strTmpImgFile) Then objFSO.DeleteFile (strTmpImgDir & "\" & strTmpImgFile)
            
            '通过FTP下载报告图像到指定目录
            Call objFtp.FuncDownloadFile(strFtpDir & strSubDir, strTmpImgDir & "\" & strTmpImgFile, strTmpImgFile)
        End If
        
        rsData.MoveNext
    Loop
    
    Set GetReportImage = objImgFileName
    
    '关闭FTP连接
    objFtp.FuncFtpDisConnect
End Function

Public Function GetReportList(ByVal lngPatId As Long, ByVal lngPageId As Long) As ADODB.Recordset
'功能:用于显示报告清单，内容预览
'lngPatID:病人id
'lngPageID:主页id
'返回ADO记录集，包含报告ID，名称，完成人，完成时间
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
       
    If lngPatId <= 0 Then Exit Function
    
    strSQL = "select B.病历ID AS 报告ID,A.医嘱内容 AS 名称,C.保存人 AS 完成人,C.完成时间 " & _
             "from 病人医嘱记录 A,病人医嘱报告 B,电子病历记录 C,影像检查记录 D " & _
             "Where A.病人ID = [1] And A.主页ID = [2] And A.Id = B.医嘱ID And B.病历ID = C.Id " & _
             "And A.诊疗类别='D' and 相关ID is null And C.完成时间 Is Not NULL " & _
             "And A.ID=D.医嘱ID(+) and A.医嘱期效=1 And A.医嘱状态 IN (3,5,6,7,8) " & _
             "order by C.完成时间 desc "
    
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "显示报告清单", lngPatId, lngPageId)
    
    Set GetReportList = rsData
End Function

Public Function GetReportFormHandle(ByVal lngReportID As Long) As Long
'功能:用于显示报告内容，内容预览
'lngReportID:报告id
'返回窗口句柄
    Dim objfrmReport As New frmReport
    
    GetReportFormHandle = objfrmReport.zlInitModule(lngReportID)
End Function

Public Sub ShowImage(ByVal lngAdviceId As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'功能:根据传入的医嘱，打开对应的影像观片站进行观片
'lngAdviceId: 医嘱ID
'frmParent:父窗体
'blnMoved:是否被转移过
    Dim objPacsCore As Object
    Dim strFtpHost As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strSDPath As String
    Dim strSDUser As String
    Dim strSDPwd As String
    Dim intImageLocation As Long
    Dim objFSO As New Scripting.FileSystemObject
    Dim lng报告ID As Long
    
    On Error GoTo DBError
    If lngAdviceId <= 0 Then Exit Sub
    
    strSQL = "Select 病历ID From 病人医嘱报告 Where 医嘱ID=[1]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lngAdviceId)
    If Not rsTmp.EOF Then lng报告ID = rsTmp!病历id
    
    '先判断是否可以继续操作
    If CheckEPRReport(lngAdviceId, IIf(lng报告ID = 0, -1, lng报告ID)) <> 1 Then
        MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，现在不能查看图像！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '先判断是否存在图像，没有图像则提示并退出
    strSQL = "Select A.检查UID,Count(B.序列UID) as 序列总数 From 影像检查记录 A,影像检查序列 B Where A.检查UID=B.检查UID And A.医嘱ID=[1] Group by A.检查UID"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "观片处理", lngAdviceId)
    If rsTmp.EOF Then
        MsgBox "没有可用于观片的报告图像。", vbInformation, gstrSysName
        Exit Sub
    End If

    '查找需要打开的所有图象信息
    strSQL = "Select /*+RULE*/ D.IP地址 As Host1,d.设备号 as 设备号1," & _
        "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'\')" & _
        "||C.检查UID||'\' As Path,E.IP地址 As Host2,e.设备号 as 设备号2, " & _
        "D.共享目录 AS 共享目录1, E.共享目录 AS 共享目录2,D.共享目录用户名 as 共享目录用户名1, " & _
        "E.共享目录用户名 AS 共享目录用户名2,D.共享目录密码 AS 共享目录密码1,E.共享目录密码 AS 共享目录密码2 " & _
        "From 影像检查记录 C,影像设备目录 D,影像设备目录 E " & _
        "Where C.位置一=D.设备号(+) And C.位置二=E.设备号(+) And C.医嘱ID=[1] "
        
    '如果有转储标志，则读取转储的历史表
    If blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    End If
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "获取共享目录信息", lngAdviceId)
    
    If rsTmp.RecordCount > 0 Then
        '创建本地的缓存目录，需要在调用观片站之前先创建这个目录，观片站中只是下载，不创建本地缓存目录
        If Not objFSO.FolderExists(App.Path & "\TmpImage\" & rsTmp("Path")) Then
            MkLocalDir App.Path & "\TmpImage\" & rsTmp("Path")
        End If
        
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
        
        Set objPacsCore = CreateObject("zl9PacsCore.clsViewer")
        objPacsCore.CallOpenViewer "", lngAdviceId, frmParent, gcnOracle, blnMoved, False
    End If

    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
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
    Dim strSQL As String
    
    On Error GoTo errH
    
    '检查报告是否已书写
    If lng报告ID = 0 Then
        strSQL = "Select 病历ID From 病人医嘱报告 Where 医嘱ID=[1]"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID)
        If Not rsTmp.EOF Then lng报告ID = rsTmp!病历id
    End If
    If lng报告ID = 0 Then
        CheckEPRReport = 0: Exit Function
    End If
    
    If Not blnBySign Then
        '检查报告执行过程(5-审核;6-报告完成)和状态(1-完成)
        '检验报告是关联到采集方式上面的，但采集方式可能为叮嘱未产生发送记录
        strSQL = _
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
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID, int执行状态)
        If Nvl(rsTmp!执行过程, 0) >= 5 Or Nvl(rsTmp!执行状态, 0) = 1 Then
            CheckEPRReport = 1
        Else
            CheckEPRReport = 2
        End If
    Else
        '通过签名版本判断报告完成的方式
        strSQL = "Select B.文件ID,Max(B.开始版) as 签名版本 From 电子病历内容 B Where B.文件ID=[1] And B.对象类型=8 Group by B.文件ID"
        strSQL = "Select B.完成时间,B.最后版本,C.签名版本 From 电子病历记录 B,(" & strSQL & ") C Where B.ID=[1] And B.ID=C.文件ID(+)"
        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng报告ID)
            
        '(签名后不能直接修改，除非修订；因此签名后最后版本应与签名版本一致)
        If IsNull(rsTmp!完成时间) Or Nvl(rsTmp!最后版本, 0) <> Nvl(rsTmp!签名版本, 0) Then
            '如果医嘱本身已经执行,即使没有签名或不符也视同完成
            strSQL = _
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
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "CheckEPRReport", lng医嘱ID, int执行状态)
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
errH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Function
