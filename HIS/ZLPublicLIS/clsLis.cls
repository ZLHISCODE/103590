VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mobjDrawReport  As Object
Private mobjPackage     As Object

Public Sub Init(ByVal cnOracle As ADODB.Connection)
'功能：接口初始化
'参数：cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的主程序系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      病历调用 lngModual=1070，无int场合参数
    Dim objFso  As FileSystemObject
    Dim rsTmp   As ADODB.Recordset
    
    If Not gblnInited Then
        Set gcnOracle = cnOracle
        Call GetUserInfo
        gblnInited = True
        gstrSysName = "检验报告打印"
        gstrHospital = gobjComLib.zlRegInfo("单位名称")
        InitObjLis
        Set objFso = New FileSystemObject
        gstrFilePath = App.Path & "\LisReportImge"
        
        gstrSignPath = gstrFilePath & "\SignImage"
        
        If Not objFso.FolderExists(gstrFilePath) Then
            objFso.CreateFolder (gstrFilePath)
        End If
        If Not objFso.FolderExists(gstrSignPath) Then
            objFso.CreateFolder (gstrSignPath)
        End If
        Set objFso = Nothing
        gbln显示图片 = True
        Call DeleteImge
        Set mobjPackage = New clsLisPackage
        '判断是否为新版LIS，根据检验报告记录与检验标本记录中的核收时间进行判断
        Set rsTmp = mobjPackage.Get_LisVersion()
        If Not ChkRsState(rsTmp) Then
            If CDate(rsTmp!核收时间1) > CDate(rsTmp!核收时间2) Then
                gblnNewLis = True
                glngSys = 2500
                glngModual = 2500
            Else
                gblnNewLis = False
                glngSys = 100
                glngModual = 1208
            End If
        Else
            MsgBox "判断LIS版本出错。", vbInformation, gstrSysName
            Exit Sub
        End If
'        gblnNewLis = False
'        glngSys = 100
'        glngModual = 1208
    End If
End Sub

Public Function GetForm() As Object
'功能：根据索引值返回对象。工作站窗体加载时调用
On Error GoTo ErrH:
   Set GetForm = New frmListReportView
'frmListReportView.Show vbModal
    Exit Function
ErrH:
    MsgBox err.Description
End Function

Public Sub PrintLisReport(ByRef objForm As Object, ByVal lngPatiID As Long, ByVal str挂号单 As String, ByVal lng主页ID As Long, ByVal bytType As Integer, _
                          bytMode As Byte, ByVal blnDoctor As Boolean, ByVal blnPrintAll As Boolean)
'功能：刷新内部界面
'参数：strName 当前窗体名称
'      lngPatiID 病人ID,str挂号单 挂号单号门诊病人才有否则为空，lng主页ID 住院病人才有否则为0
'      str挂号单 挂号单，门诊病人传入
'      lng主页ID 主页ID，住院病人传入
'      bytType 1-选项卡，2-按钮打印
'    Dim bytMode         As Byte        1-预览，2-打印
    Dim strSql As String
    Dim strPage As String
    Dim blnState As Boolean
    Dim rsParameter As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim str挂号天数 As String
    Dim str申请开始时间 As String
    Dim str申请结束时间 As String
    Dim bln是否急诊 As Boolean
    Dim lngPageCount As Long

    On Error GoTo ErrH:
    
    gblnPrintAll = blnPrintAll
    gblnDoctor = blnDoctor
    gbtyModel = bytType
    If objForm Is Nothing Then Exit Sub
    If lngPatiID = 0 Then    'Exit Sub
        Call objForm.HideControl(lngPatiID, blnState)
        Exit Sub
    Else
        If gbtyModel = 1 Then objForm.Show

        If Not mobjDrawReport Is Nothing Then Set mobjDrawReport = Nothing    'mobjDrawReport.DestoryDrawReport

        Set mobjDrawReport = New clsDrawLisReport

        If Not (mobjDrawReport Is Nothing) Then
            blnState = mobjDrawReport.InitData("zlPublicLIS.clsLisReportPublic", "公共模块\" & "zl9PrintMode" & "\Default")
        End If


        '获取报告申请的开始与结束时间
        Set rsTmp = mobjPackage.Get_RequestTime(lngPatiID, lng主页ID, str挂号单)

        str挂号天数 = gobjDatabase.GetPara("挂号有效天数", 100)
        If Not ChkRsState(rsTmp) Then
            bln是否急诊 = (rsTmp!急诊 = 1)
            If bln是否急诊 And lng主页ID = 0 Then
                str申请结束时间 = Format(rsTmp!申请结束时间, "yyyy-MM-dd HH:mm:ss")
                str申请开始时间 = Format(CDate(Format(rsTmp!申请结束时间, "YYYY-MM-DD")) - Mid(str挂号天数, 2, 1) & " 00:00:00", "yyyy-MM-dd HH:mm:ss")
            ElseIf (Not bln是否急诊) And lng主页ID = 0 Then
                str申请结束时间 = Format(rsTmp!申请结束时间, "yyyy-MM-dd HH:mm:ss")
                str申请开始时间 = Format(CDate(Format(rsTmp!申请结束时间, "YYYY-MM-DD")) - Mid(str挂号天数, 1, 1) & " 00:00:00", "yyyy-MM-dd HH:mm:ss")
            Else
                str申请开始时间 = Format(rsTmp!申请开始时间, "yyyy-MM-dd HH:mm:ss")
                str申请结束时间 = Format(rsTmp!申请结束时间, "yyyy-MM-dd HH:mm:ss")
            End If
        Else
            Exit Sub
        End If

        Set rsParameter = mobjPackage.CreateParameter()

        Call mobjPackage.SetParameter(rsParameter, "报告属性", "")
        Call mobjPackage.SetParameter(rsParameter, "病人ID", lngPatiID)
        Call mobjPackage.SetParameter(rsParameter, "挂号单号", str挂号单)
        Call mobjPackage.SetParameter(rsParameter, "主页ID", lng主页ID)
        Call mobjPackage.SetParameter(rsParameter, "申请开始时间", str申请开始时间)
        Call mobjPackage.SetParameter(rsParameter, "申请结束时间", str申请结束时间)
        ''        Call mobjPackage.SetParameter(rsParameter, "标本顺序", mobjPackage.Get_LisSampleSort())
        ''        Call mobjPackage.SetParameter(rsParameter, "类别顺序", mobjPackage.Get_LisTypeSort())
        ''        Call mobjPackage.SetParameter(rsParameter, "项目顺序", mobjPackage.Get_LisNameSort())

        '字体名称=宋体'字体大小=9'字体粗体=False'字体斜体=False|标题1字体名称=宋体'标题1字体大小=18'标题1字体粗体=FALSE'标题1字体斜体=FALSE

        mobjDrawReport.BeginGetData

        '第二步：页眉页脚
        '------------------------------------------------------------------------------------------------------------------
        '    If mmobjDrawReport.GetPageHeadData("zlPlugIn.clsElePersonPublic", rsParameter) = False Then GoTo EndHand
        blnState = mobjDrawReport.GetPageFootData("zlPublicLIS.clsLisReportPublic", rsParameter)

        '第四步：再输出内容(根据团体报告的组件组成循环调用子组件来完成数据组织)
        '------------------------------------------------------------------------------------------------------------------
        '开始获取组件输出数据
        blnState = mobjDrawReport.GetElementData("zlPublicLIS.clsLisResultItem", rsParameter, False, 0)

        Call mobjDrawReport.EndGetData
        If blnState Then
            blnState = mobjDrawReport.ExportReport(objForm, bytMode, "检验报告打印", "", "", strPage)

            '写入打印记录
            If Not grsSample Is Nothing Then
                grsSample.Filter = ""
                If grsSample.RecordCount > 0 Then grsSample.MoveFirst
                Do While Not grsSample.EOF
                    If gblnNewLis Then
                        strSql = "Zl_检验报告打印_Edit(1," & grsSample("标本ID") & "," & IIf(blnDoctor, 1, 0) & ")"
                        Call gobjPublicHisCommLis.ComExecuteOtherProc(1, strSql, "打印标本")
                    Else
                        strSql = "Zl_检验标本记录_标本质控(" & grsSample("标本ID") & ",'',1)"
                         Call gobjDatabase.ExecuteProcedure(strSql, gstrSysName)
                    End If
                    grsSample.MoveNext
                Loop
            End If
        End If
        lngPageCount = objForm.GetPageCount()
        If lngPageCount <= 0 Then blnState = False
        Call objForm.HideControl(lngPatiID, blnState)
    End If
    Exit Sub
ErrH:
    MsgBox err.Description, vbCritical, "检验报告打印"
    '    err.Clear
End Sub

Private Sub Class_Initialize()
    If gobjComLib Is Nothing Then
        Set gobjComLib = GetObject("", "zl9Comlib.clsComlib")
        Set gobjCommFun = GetObject("", "zl9Comlib.clsCommfun")
        Set gobjControl = GetObject("", "zl9Comlib.clsControl")
        Set gobjDatabase = GetObject("", "zl9Comlib.clsDatabase")
        Set gobjPrintMode = GetObject("", "zl9PrintMode.zlPrintMethod")
        Set gobjSystem = GetObject("", "zl9Comlib.clsSystem")
    End If
    gstrSysName = GetSetting(appName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
End Sub

Private Sub Class_Terminate()
    If Not mobjDrawReport Is Nothing Then mobjDrawReport.DestoryDrawReport
    '退出程序时，删除图片目录下的所有图片
    Call DeleteImge
    Set mobjPackage = Nothing
    gblnInited = False
'    Set gobjComLib = Nothing
End Sub
