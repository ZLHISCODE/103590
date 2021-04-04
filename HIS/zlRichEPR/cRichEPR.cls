VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichEPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'################################################################################################################
'##模 块 名：czlRichEPR.cls
'##创 建 人：吴庆伟
'##日    期：2005年8月11日
'##修 改 人：
'##日    期：
'##描    述：系统入口类
'################################################################################################################

Option Explicit

'################################################################################################################
'## 功能：  部件初始化
'##
'## 说明：  所有公共变量应在本部件的模块中完成定义
'################################################################################################################
Public Sub InitRichEPR(cnMain As ADODB.Connection, frmParent As Object, ByVal lngSys As Long, ByVal blnShowInTaskBar As Boolean)
    If gfrmPublic Is Nothing Then Set gfrmPublic = New frmPublic
    
    Set gfrmParent = frmParent
    Set gcnOracle = cnMain
    'Call InitCommon(gcnOracle)
        
    glngSys = lngSys
    gstrDbOwner = GetDbOwner(glngSys)
    gblnShowInTaskBar = blnShowInTaskBar
    gstrProductName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("产品名称"), Default:="中联")
    gstrAviPath = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrAviPath"), Default:="")
    gstrSysName = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrSysName"), Default:="")
    gstrVersion = GetSetting(AppName:="ZLSOFT", Section:="注册信息", Key:=UCase("gstrVersion"), Default:="")
    gstrMatch = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = "0", "%", "")
    gstrPrivsEpr = ";" & GetPrivFunc(glngSys, 1070) & ";"
    
    Call GetUserInfo
    Call InitPreDefinedKeys
    
    On Error Resume Next
    If Not frmParent Is Nothing Then
        Set gobjEmr = frmParent.mobjEMR  '新病历需要部件
        If Err.Number <> 0 Then
            Set gobjEmr = Nothing
            Err.Clear
        End If
    End If
      
    If gobjPlugIn Is Nothing Then
        On Error Resume Next
        Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not gobjPlugIn Is Nothing Then
            Call gobjPlugIn.Initialize(gcnOracle, glngSys, 1070)
        End If
        Err.Clear: On Error GoTo 0
    End If
End Sub

'################################################################################################################
'## 功能：  外部直接模块调用
'## 参数：  lngSys      :系统编号
'##         lngModul    :模块编号
'##         frmParent   :父窗体
'################################################################################################################
Public Sub zlExecuteCode(ByVal lngSys As Long, ByVal lngModul As Long, ByVal frmParent As Variant, Optional strPrivs As String)
    glngSys = lngSys: glngModul = lngModul
    gstrDbOwner = GetDbOwner(glngSys)
    
    If strPrivs = "" Then
        gstrPrivs = GetPrivFunc(glngSys, glngModul)
    Else
        gstrPrivs = strPrivs
    End If
    
    Select Case lngModul
    Case 1071 '病历标记图形管理
        If IsObject(frmParent) Then
            frmMarkMapMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmMarkMapMan.hWnd, frmParent
        End If
    Case 1072 '护理记录项目管理
        If IsObject(frmParent) Then
            frmTendItemMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmTendItemMan.hWnd, frmParent
        End If
    Case 1073 '病历预制提纲管理
        If IsObject(frmParent) Then
            frmSentenceMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmSentenceMan.hWnd, frmParent
        End If
    Case 1074 '病历文件管理
        If IsObject(frmParent) Then
            frmEPRFileMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmEPRFileMan.hWnd, frmParent
        End If
    Case 1075 '病历范文管理
        If IsObject(frmParent) Then
            frmEPRModelMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmEPRModelMan.hWnd, frmParent
        End If
    Case 1076   '诊疗单据设置
        If IsObject(frmParent) Then
            frmEPRBillMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmEPRBillMan.hWnd, frmParent
        End If
    Case 1950   '移动护士站基础设置
        If IsObject(frmParent) Then
            frmTendItemTransfusion.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmTendItemTransfusion.hWnd, frmParent
        End If
    Case 1273 '病人病历检索
        If IsObject(frmParent) Then
            Dim objForm As Object
            Set objForm = frmParent '直接赋值不兼容
            frmEPRSearchMan.ShowSearchClinic objForm, glngDeptId
        Else
            'ZLBH融合调用
            frmEPRSearchMan.ShowSearchClinic Nothing, glngDeptId
        End If
    Case 1274 '疾病申报管理
        If IsObject(frmParent) Then
            frmDiseaseReportMan.Show , frmParent
        Else
            'ZLBH融合调用
            zlCommFun.ShowChildWindow frmDiseaseReportMan.hWnd, frmParent
        End If
    End Select
End Sub

'################################################################################################################
'## 功能：  关闭所有窗体
'################################################################################################################
Public Function CloseWindows() As Boolean
    '--------------------------------------
    '功能:关闭所有子窗口
    '--------------------------------------
    Dim frmThis As Form
    Err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
    Set gfrmPublic = Nothing
    CloseWindows = (Forms.Count = 0)
End Function

Public Sub PrintOrPreviewDoc(frmParent As Object, eDocType As EPRDocTypeEnum, lngRecId As Long, _
    Optional blnPrint As Boolean, Optional blnFixLastVersion As Boolean = False)
'正式打印或预览已经保存的指定文档
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select 编辑方式 From 电子病历记录 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, App.Title, lngRecId)
    If rsTemp.EOF Then Exit Sub
    
    If rsTemp!编辑方式 = 0 Then
        Dim frmP As frmPrintPreview, blnOrigMode As Boolean, strAsk As String
        If lngRecId = 0 Then Exit Sub
        
        If blnFixLastVersion Then
            blnOrigMode = False
        Else
            strAsk = "可以按清洁格式或原始格式" & IIf(blnPrint, "打印", "预览") & "："
            strAsk = strAsk & vbCrLf & "    最终格式：不包含修改痕迹的清洁格式"
            strAsk = strAsk & vbCrLf & "    原始格式：包含修改痕迹的草稿格式"
            strAsk = strAsk & vbCrLf & "按“最终格式”模式" & IIf(blnPrint, "打印", "预览") & "吗？"
            
            Select Case MsgBox(strAsk, vbYesNoCancel + vbQuestion, gstrSysName)
            Case vbYes
                blnOrigMode = False
            Case vbNo
                blnOrigMode = True
            Case Else
                Exit Sub
            End Select
        End If
        
        Set frmP = New frmPrintPreview
        frmP.DoMultiDocPreview frmParent, eDocType, , , , , lngRecId, blnPrint, blnOrigMode
        Unload frmP
        Set frmP = Nothing
    Else
        Dim objTableEpr As New cTableEPR
        objTableEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTableEpr.InitOpenEPR frmParent, TabEm_修改, TabET_单病历审核, lngRecId, False, 2
        objTableEpr.zlPrintDoc frmParent, Not blnPrint
        Set objTableEpr = Nothing
    End If
End Sub

Public Sub ViewDocument(frmParent As Object, lngRecId As Long, Optional blnCanPrint As Boolean)
'打开查阅文档窗口
Dim rsTemp As ADODB.Recordset
    gstrSQL = "Select 编辑方式 From 电子病历记录 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, App.Title, lngRecId)
    If rsTemp.EOF Then Exit Sub
    
    If rsTemp!编辑方式 = 0 Then
        Dim fViewDoc As New frmEPRView
        fViewDoc.ShowMe frmParent, lngRecId, , blnCanPrint
    Else
        Dim objTableEpr As New cTableEPR
        objTableEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTableEpr.InitOpenEPR frmParent, TabEm_修改, TabET_单病历审核, lngRecId, True, 2
        Set objTableEpr = Nothing
    End If
End Sub


Public Function GetDocumentText(lngRecId As Long) As String
'功能：根据病人病历id获取病历文件的文本内容。
'参数：lngRecId-电子病历文件ID
    Dim frmContent As frmDockEPRContent, strTmp As String
    
    strTmp = Clipboard.GetText
    
    Set frmContent = New frmDockEPRContent
    Call frmContent.zlRefresh(lngRecId, "", False)
    frmContent.edtThis.SelectAll
    frmContent.edtThis.Copy
    GetDocumentText = Clipboard.GetText
        
    Clipboard.SetText strTmp
End Function
Public Sub EditDocument(ByVal lngModul As Long, ByVal frmParent As Object, ByVal lngDeptId As Long, ByVal strFileID As String)
'lngModul 哪个模块调用
'frmParent 父窗口
'lngDeptID 操作科室,并非病人科室,也非操作员科室,比如在非缺省科室书写非本科会诊记录
'strFileID 文件ID,预留扩展用string
Dim rsTemp As New ADODB.Recordset, strEprPrivs As String, lngFileID As Long
Dim objDoc As New cEPRDocument, objTabEpr As Object
    On Error GoTo errHand
    lngFileID = CLng(strFileID)
    gstrSQL = "Select 病人id, 主页id, 病人来源, 婴儿, 病历种类, 科室id, 文件id, 创建人, 保存人, 归档人, 完成时间, Nvl(编辑方式, 0) 编辑方式" & vbNewLine & _
                "From 电子病历记录" & vbNewLine & _
                "Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历信息", lngFileID)
    If rsTemp.EOF Then Exit Sub
            
    Select Case rsTemp!病历种类
        Case cpr护理病历
            strEprPrivs = GetPrivFunc(glngSys, 1255)
        Case cpr住院病历, cpr知情文件, cpr诊断文书, cpr诊疗报告
            strEprPrivs = GetPrivFunc(glngSys, 1251)
        Case cpr门诊病历
            strEprPrivs = GetPrivFunc(glngSys, 1250)
    End Select
    
    If InStr(strEprPrivs, IIf(rsTemp!病历种类 = cpr护理病历, "护理病历书写", "病历书写")) = 0 Then ViewDocument frmParent, lngFileID: Exit Sub
    If NVL(rsTemp!完成时间) = "" And lngDeptId <> rsTemp!科室ID And rsTemp!保存人 <> gstrUserName Then ViewDocument frmParent, lngFileID: Exit Sub '非本科病历不允许修改(书写后,科室合并,人员科室从属变更应允许修改)
    If NVL(rsTemp!完成时间) = "" And InStr(1, strEprPrivs, IIf(rsTemp!病历种类 = cpr护理病历, "他人护理病历", "他人病历")) = 0 And rsTemp!创建人 <> gstrUserName Then ViewDocument frmParent, lngFileID: Exit Sub '未签名，无他人病历权限，非本人创建不允许修改
    If NVL(rsTemp!完成时间) <> "" And InStr(1, strEprPrivs, IIf(rsTemp!病历种类 = cpr护理病历, "护理病历审阅", "病历审阅")) = 0 Then ViewDocument frmParent, lngFileID: Exit Sub  '已签名，无他人病历权限，非最后本人编辑不允许修改
'    If NVL(rsTemp!归档人, "") <> "" Then ViewDocument frmParent, lngFileID: Exit Sub '已归档不允许编辑
    
    If rsTemp!编辑方式 = 0 Then '全文式编辑
        objDoc.InitEPRDoc cprEM_修改, IIf(rsTemp!完成时间 = "", cprET_单病历编辑, cprET_单病历审核), lngFileID, NVL(rsTemp!病人来源, 2), rsTemp!病人ID, rsTemp!主页ID, NVL(rsTemp!婴儿, 0), lngDeptId
        objDoc.ShowEPREditor frmParent
    Else '表格病历
        Set objTabEpr = DynamicCreate("zlTableEPR.cTableEPR", "表格病历")
        objTabEpr.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        objTabEpr.InitOpenEPR frmParent, cprEM_修改, IIf(rsTemp!完成时间 = "", cprET_单病历编辑, cprET_单病历审核), lngFileID, True, 0, NVL(rsTemp!病人来源, 2), rsTemp!病人ID, rsTemp!主页ID, NVL(rsTemp!婴儿, 0), lngDeptId, , strEprPrivs, False, False
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload frmPublic
    If Not gfrmPublic Is Nothing Then Unload gfrmPublic
    Set gfrmPublic = Nothing
    Set gobjFSO = Nothing
    Set gcnOracle = Nothing
    Set gfrmParent = Nothing
    Set gobjPacsCore = Nothing
    Set gobjESign = Nothing
    Set gobjEmr = Nothing
End Sub
