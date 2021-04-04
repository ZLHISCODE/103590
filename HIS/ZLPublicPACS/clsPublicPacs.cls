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

Private mobjReport As Object  'zlRichEPR.cDockReport  '电子病历报告对象
Private mblnIsInit As Boolean   '判断此接口是否进行初始化操作
Private mfrmReport As frmReport '报告文档编辑器窗口

Public Function InitInterface(ByVal cnOracle As ADODB.Connection, ByVal strDBUser As String, Optional ByVal blnUseAlone As Boolean = False) As Boolean
'对接口所需数据进行初始化，成功返回TRUE
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "gstrSysName", "")
    
    Set gcnOracle = cnOracle
    
    If gcnOracle Is Nothing Then Exit Function
    
    If gcnOledb Is Nothing Then
        Call InitOledbConn(blnUseAlone)
    End If
    
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9ComLib.clsComLib")
    End If
    
    gstrPrivs = ";" & gobjComLib.GetPrivFunc(100, 9004) & ";"
        
    InitInterface = True
    mblnIsInit = True
End Function

Public Function zlHasImage(lngOrderID As Long) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：传递一个医嘱ID给接口方法,根据这个医嘱ID进行判断是否有影像图片
'参数：lngOrderID--病人医嘱ID编号
'返回：true--该病人有影像图片  false--该病人没有影像图片
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    On Error GoTo errHandle
    
    zlHasImage = mdlPublic.HasImage(lngOrderID)
    
    Exit Function
errHandle:
    If gobjComLib.ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function GetReportImage(ByVal lngAdviceID As Long, ByVal strDocReportId As String, _
    ByVal strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'功能:将检查医嘱所对应报告中的所有图像，保存到指定目录,获取PACS报告中的所有图像
'lngAdviceId:医嘱id
'strDocReportId:PACS智能报告编辑器文档ID，如果不是PACS智能报告编辑器书写的报告，则不传递该参数
'strTmpImgFolder:图像缓存目录
'objImgFileName:报告图像文件名集合

'说明：此过程目前只有体检调用，一个医嘱对应一份报告
    Dim objFSO As New Scripting.FileSystemObject
    Dim strSql As String
    Dim rsData As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim strTmpImgFile As String
    Dim objImgFileName As New Collection
    Dim objFtp As New clsFtp
    Dim strFtpDir As String
    Dim strIp As String
    Dim strFtpPswd As String
    Dim strFTPUser As String
    Dim strSubDir As String
    Dim lngPatFrom As Long
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
     
    If lngAdviceID <= 0 Then Exit Function
    
    '创建指定目目录
    If Not objFSO.FolderExists(strTmpImgDir) Then MkLocalDir strTmpImgDir
  
    If Trim(strDocReportId) <> "" Then '说明是报告文档编辑器的报告
        Set GetReportImage = frmReport.GetReportImage(lngAdviceID, strDocReportId, strTmpImgDir, strMsg)
        Exit Function
    End If
    
    strSql = "Select D.设备号,D.FTP用户名,D.FTP密码,D.IP地址,'/'||D.Ftp目录||'/' As Root, " & _
             "Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')||C.检查UID As URL,d.设备号 as 设备号 " & _
             "From 影像检查记录 C,影像设备目录 D " & _
             "Where  C.位置一=D.设备号(+) and C.医嘱ID=[1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", lngAdviceID)
    
    If rsTemp.EOF Then
        strMsg = "此报告没有图像。"
        Exit Function
    End If
    
    If rsTemp.RecordCount > 0 Then
        If Trim(Nvl(rsTemp("设备号"))) = "" Then Exit Function
        
        strFtpDir = Nvl(rsTemp("Root"))
        strIp = Nvl(rsTemp("IP地址"))
        strFtpPswd = Nvl(rsTemp("FTP密码"))
        strFTPUser = Nvl(rsTemp("FTP用户名"))
        strSubDir = Nvl(rsTemp("URL"))
    End If
    
    strSql = "select c.对象属性 " & _
             "from 病人医嘱报告 a,电子病历内容 b,电子病历内容 c " & _
             "Where a.医嘱ID =[1] And a.病历id = b.文件id And b.Id = c.父id " & _
             "and b.对象类型 = 3 And Substr(b.对象属性, Instr(b.对象属性, ';', 1, 18) + 1, 1) = '2' " & _
             "and Nvl(c.终止版, 0) = 0 "

    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "对象属性", lngAdviceID)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    '连接FTP
    If objFtp.FuncFtpConnect(strIp, strFTPUser, strFtpPswd) = 0 Then
        strMsg = "FTP不能正常连接，请检查网络设置。"
        Exit Function
    End If

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
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
       
    If lngPatId <= 0 Then Exit Function
    
    strSql = "select B.病历ID AS 报告ID,A.医嘱内容 AS 名称,C.保存人 AS 完成人,C.完成时间 " & _
             "from 病人医嘱记录 A,病人医嘱报告 B,电子病历记录 C,影像检查记录 D " & _
             "Where A.病人ID = [1] And A.主页ID = [2] And A.Id = B.医嘱ID And B.病历ID = C.Id " & _
             "And A.诊疗类别='D' and 相关ID is null And C.完成时间 Is Not NULL " & _
             "And A.ID=D.医嘱ID and A.医嘱期效=1 And A.医嘱状态 IN (3,5,6,7,8) " & _
             "order by C.完成时间 desc "
    
    Set rsData = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "显示报告清单", lngPatId, lngPageId)
    
    Set GetReportList = rsData
End Function

Public Function GetReportFormHandle(ByVal lngReportID As Long) As Long
'功能:用于显示报告内容，内容预览
'lngReportID:报告id
'返回窗口句柄
    Dim lngAdviceID As Long
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If

    Set mobjReport = CreateObject("zlRichEPR.cDockReport")

    If mobjReport Is Nothing Then Exit Function

    lngAdviceID = GetAdviceID(lngReportID)
    mobjReport.zlRefresh 0, 0
    mobjReport.zlRefresh lngAdviceID, 0

    GetReportFormHandle = mobjReport.zlGetForm().hWnd
End Function

Public Sub ShowImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'功能:根据传入的医嘱，打开对应的影像观片站进行观片
'lngAdviceId: 医嘱ID
'frmParent:父窗体
'blnMoved:是否被转移过
    Call mdlPublic.ViewImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub ShowPatientImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'功能:根据传入的医嘱ID，查找到对应的病人ID，打开专业版PACS的检查列表观片
'lngAdviceId: 医嘱ID
'frmParent:父窗体
'blnMoved:是否被转移过
    Call mdlPublic.ViewPatientImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Sub ShowStaticImage(ByVal lngAdviceID As Long, ByVal frmParent As Object, Optional ByVal blnMoved As Boolean = False)
'功能:查看关键图像
'lngAdviceId: 医嘱ID
'frmParent:父窗体
'blnMoved:是否被转移过
    Call mdlPublic.ViewStaticImage(lngAdviceID, frmParent, blnMoved)
    Exit Sub
DBError:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComLib.SaveErrLog
End Sub

Public Function Show3DImage(ByVal lng医嘱ID As Long, frmParent As Object) As Long
'功能：3D观片
'0   成功
'-121    调用参数错误
'-122    未正确安装PACS及接口文件
'-102    未正确安装PACS及接口文件
'-103    网络连接错误
'-104    数据库错误
'-101    其他错误
    Show3DImage = mdlPublic.View3DImage(lng医嘱ID, frmParent)
End Function


Public Function GetEditerType(ByVal lngDeptId As Long) As Integer
'参数：检查医嘱对应的执行科室ID
'功能：获取当前使用的编辑器模式
'返回：0--电子病历编辑器, 1--PACS报告编译器，2--报告文档编辑器
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Function
    End If
    
    GetEditerType = 0
    
    On Error GoTo ErrH
    '获取当前编辑器模式
    strSql = "Select 参数值 From 影像流程参数 Where 科室ID=[1] And 参数名=[2]"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "GetReportContent", lngDeptId, "报告编辑器")
    
    If rsTmp.RecordCount > 0 Then GetEditerType = Val(Nvl(rsTmp!参数值))
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
    GetEditerType = 0
End Function

Public Function zlDocGetList(ByVal lngPatId As Long, Optional ByVal lngPageId As Long, Optional ByVal strRegNo As String) As Recordset
'返回Pacs文档编辑器书写的报告列表
'lngPatID:病人ID
'lngPageID:主页ID
'strRegNo:挂号单
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Set zlDocGetList = mfrmReport.zlDocGetList(lngPatId, lngPageId, strRegNo)
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetList = Nothing
End Function

Public Function zlDocGetListWithAdvice(ByVal strAdviceId As String) As Recordset
'返回Pacs文档编辑器书写的报告列表
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Set zlDocGetListWithAdvice = mfrmReport.zlDocGetListWithAdvice(strAdviceId)
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetListWithAdvice = Nothing
End Function

Public Function zlDocGetForm(Optional blnShow As Boolean = False, Optional blnRIS As Boolean = False, Optional blnAutoView As Boolean = True) As Object
'功能：返回Pacs文档编辑器的报告查看窗口对象
'参数：     blnShow -- 是否显示左边的过滤条件
'           blnRIS --  是否RIS，是RIS则不显示退出等按钮
'           blnAutoView -- 是否自动查阅
    
On Error GoTo ErrH
    If mfrmReport Is Nothing Then
        Set mfrmReport = New frmReport
        
        '设置是否显示过滤条件和退出按钮，这两句必须要放在BorderStyle前面，确保FormLoad时blnShow参数是正确的
        '调用BorderStyle时会触发FormLoad
        mfrmReport.mblnShow = blnShow
        mfrmReport.mblnRIS = blnRIS
        mfrmReport.mblnAutoView = blnAutoView
        
        mfrmReport.BorderStyle = 0
        mfrmReport.Caption = mfrmReport.Caption
    End If
    
    Set zlDocGetForm = mfrmReport
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetForm = Nothing
End Function

Public Sub zlDocRefresh(ByVal strReportId As String)
'刷新检查报告id对应的报告内容
On Error GoTo ErrH
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.RefreshReportContent(strReportId)
    
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Sub zlDocViewStateUpdate(ByVal blnState As Boolean, ByVal lngAdviceID As Long)
'更新pacs文档编辑器报告的查阅状态,如果blnState=True，表示对应的所有报告为已阅，blnState=False,表示为未阅

On Error GoTo ErrH
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.zlDocViewStateUpdate(blnState, lngAdviceID)
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub


Public Sub zlDocShowReport(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional ByVal blnAutoView As Boolean = True, Optional objParent As Object, Optional blnShowModal As Boolean = False)
'功能；当使用报告文档编辑器时，根据医嘱ID获取报告内容并显示
'objParent:父窗体
'lngAdviceId:医嘱ID
'blnAutoView:是否自动查阅，True--是
    On Error GoTo ErrH
    
    If mblnIsInit = False Then
        Call ShowErrInfo
        Exit Sub
    End If
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.ShowMe(lngAdviceID, strReportId, blnAutoView, objParent, blnShowModal)
    
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Sub PrintReport(ByVal strDocId As String, Optional ByVal strPrinterName As String = "", Optional ByVal blnIsPreview As Boolean = False)
'功能：打印报告
'参数说明:
'strDocId---报告ID
'strPrinterName---打印机名称，为空时弹出打印设置框
'blnIsPreview---为True时显示预览窗口
    On Error GoTo ErrH
    
    Call frmPrintReport.PrintReport(strDocId, strPrinterName, blnIsPreview)
    Exit Sub
    
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Call gobjComLib.SaveErrLog
End Sub

Public Function zlDocRefreshForm(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional objParent As Object) As Boolean
    '功能： 刷新报告文档编辑器窗口，使用在RIS接口中，在RIS中打开PACS的历史报告和图像
    '参数： lngAdviceId -- 医嘱ID
    '       strReportId -- 报告ID，可空
    '       objParent -- 父窗体
    On Error GoTo err
    
    If mfrmReport Is Nothing Then Set mfrmReport = New frmReport
    Call mfrmReport.RefreshForm(lngAdviceID, strReportId, objParent)
    
    zlDocRefreshForm = True
    Exit Function
err:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Function

Private Sub ShowErrInfo()
    MsgBox "接口还未初始化，不能进行相应操作！", vbInformation, "系统信息"
End Sub

Private Sub Class_Initialize()
    mblnIsInit = False
End Sub

Private Sub Class_Terminate()
    mblnIsInit = False
    
    If Not mfrmReport Is Nothing Then Set mfrmReport = Nothing
    If Not mobjReport Is Nothing Then Set mobjReport = Nothing
End Sub
