VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRichReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const M_STR_MODULE_MENU_TAG = ""

Public Event BeforeEdit(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterOpen(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterSaved(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngSaveType As Long)
Public Event AfterClosed(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterPrinted(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event AfterDeleted(ByVal lngOrderID As Long, ByVal strDocID As String)
Public Event DataRefresh(ByVal strDocInfo As String)
Public Event DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)

Private WithEvents mobjInterfaceEvent As ZLPacsVBCommon.clsEvent
Attribute mobjInterfaceEvent.VB_VarHelpID = -1
Private mobjPacsInterface As Object
Attribute mobjPacsInterface.VB_VarHelpID = -1
Private mobjOwner As Object
Private mstrPrivs As String
Private mlngModule As Long
Private mlngDeptID As Long
Private mlngAdviceID As Long
Private mstrTmpReportId As String
Private mlngPatientId As Long
Private mlngStudyState As Long
Private mblnCanPrint As Boolean     '平诊需审核后才能打印
Private mblnDirectPrint As Boolean  '紧急病人和绿色通道病人可直接打印

Private mblnVisible As Boolean

Private mblnRefreshState As Boolean
Private mobjPacsCore As zl9PacsCore.clsViewer
Private mObjActiveMenuBar As CommandBars
Private mlngReportFormHandle As Long

Private mobjFtp As New clsFtp
Private mobjFtpInf As TFtpDeviceInf
Private mstrSubDir As String
Private mstrDestMainDir As String
Private mblnIsNoAskPrint As Boolean             '静默打印,True--静默打印,False--非静默打印
Private mintCompletePrint As Integer
Private mblnIsShowLastState As Boolean
Private mblnIsFirstImgUpLoadFiald As Boolean      '如果第一张图像上传失败，重试后还是失败，则后面的图像不再重试上传
Private mblnIsFirstImgDownLoadFiald As Boolean      '如果第一张图像下载失败，重试后还是失败，则后面的图像不再重试下载

Property Get AdviceId()
'获取当前报告的医嘱Id
    AdviceId = mlngAdviceID
End Property

'静默打印
Public Property Get IsNoAskPrint() As Boolean
    IsNoAskPrint = mblnIsNoAskPrint
End Property

Public Property Let IsNoAskPrint(ByVal vNewValue As Boolean)
    mblnIsNoAskPrint = vNewValue
    Call SetIsNoAskPrint(vNewValue)
End Property


'设置报告的图像处理对象
Property Get PacsCore() As zl9PacsCore.clsViewer
    Set PacsCore = mobjPacsCore
End Property

Property Set PacsCore(objPacsCore As zl9PacsCore.clsViewer)
    Set mobjPacsCore = objPacsCore
End Property


Public Sub NotificationRefresh()
'通知刷新
    mblnRefreshState = False
End Sub


Public Sub zlCreateMenu(ByVal objActiveMenuBar As Object)
'创建模块菜单栏按钮
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

On Error GoTo ErrorHand
    Set mObjActiveMenuBar = objActiveMenuBar

    Set cbrMenuBar = objActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "报告", 3, False)
    cbrMenuBar.ID = conMenu_EditPopup
    cbrMenuBar.Category = ""
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_NoAskPrint, "使用静默打印", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_FinalShowMode, "最终状态显示", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Preview, "预览", "", 102, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Print, "打印", "", 103, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_BatPrint, "批量打印", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 3002, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "清除状态", "", 21903, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "删除", "", 0, False)
    End With
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlCreateToolBar(ByVal objToolBar As Object)
'创建工具栏按钮
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long

On Error GoTo ErrorHand
    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue, , True)
    
    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "报告预览", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "报告打印", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 2607, False, lngIndex + 3) 'IconId=3002
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlExecuteMenu(ByVal lngMenuId As Long, Optional ByVal strDocID As String = 0)
'执行编辑器菜单
    Dim objControl As XtremeCommandBars.ICommandBarControl
        
On Error GoTo ErrorHand
    Set objControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    If objControl Is Nothing Then Exit Sub
    
    Select Case lngMenuId
        Case conMenu_PacsReport_FinalShowMode
            objControl.Checked = Not objControl.Checked
            mblnIsShowLastState = objControl.Checked
            zlDatabase.SetPara "最终状态显示", IIf(mblnIsShowLastState, 1, 0), glngSys, mlngModule
            SetReportFinalShowMode mblnIsShowLastState
            
        Case conMenu_File_Preview                   '预览
            Call OpenFormForReportPrint(strDocID, True)
            
        Case conMenu_File_Print                     '打印
            Call OpenFormForReportPrint(strDocID, False)
            
        Case conMenu_File_BatPrint                  '批量打印
            Call OpenFormForReportPrint(strDocID, False)
            
        Case conMenu_PacsReport_Open                '书写
            Call OpenFormForEditReport
        
        Case conMenu_PacsReport_ClearWritingState   '清除状态
            Call ClearWritingState
                
        Case conMenu_Edit_Delete                    '删除
            Call DeleteReport
        
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Sub ClearWritingState()
'清除状态
    Dim strDocID As String
    
On Error GoTo ErrorHand
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    strDocID = GetCurrDocId(mlngAdviceID)
    
    If Trim(strDocID) = "" Then Exit Sub
    
    Call mobjPacsInterface.EditLockedInfo(strDocID, "")
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlUpdateMenu(ByVal objControl As XtremeCommandBars.ICommandBarControl)
    
'更新编辑器菜单
On Error GoTo ErrorHand
    If objControl Is Nothing Then Exit Sub
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Select Case objControl.ID
        Case conMenu_PacsReport_FinalShowMode
            objControl.Checked = mblnIsShowLastState
            
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_BatPrint     '打印报告,预览报告,批量打印
            objControl.Visible = CheckPopedom(mstrPrivs, "PACS报告打印")
             
            If objControl.ID <> conMenu_File_BatPrint Then
                If Not mobjPacsInterface Is Nothing Then
                    objControl.Enabled = mobjPacsInterface.CanPrint("")
                Else
                    objControl.Enabled = False
                End If
            End If
            
        Case conMenu_PacsReport_Open                '书写
            Call InitEditButtonCaption(objControl)
            
        Case conMenu_Edit_Delete                    '删除
            objControl.Visible = (GetCurrDocId(mlngAdviceID) <> "" And (CheckPopedom(mstrPrivs, "PACS报告书写") Or CheckPopedom(mstrPrivs, "PACS报告删除")))
            If objControl.Visible Then objControl.Visible = (CheckPopedom(mstrPrivs, "PACS他人报告") Or GetDocCreator(mlngAdviceID) = UserInfo.姓名)
        Case conMenu_File_Open                      '查阅
            
        Case conMenu_Tool_Search                    '报告检索…
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub InitEditButtonCaption(ByVal objControl As XtremeCommandBars.ICommandBarControl)
'根据报告状态显示书写按钮名称
    Dim intReportState As Integer
On Error GoTo ErrorHand
    '1.在没有签名的情况下，书写按钮显示名称为“书写”
    '2.诊断签名后，如果没有报告审核权限，书写按钮则显示“修订”，如果有审核权限，书写按钮则显示“审核”
    '3.审核签名后，书写按钮则显示“审核”
    '4.终审签名后，书写按钮显示“查阅”
    '5.驳回处理，诊断驳回按诊断签名处理，审核驳回按审核签名处理
    If mlngAdviceID <= 0 Then
        objControl.Enabled = False
        objControl.Caption = "书写"
        objControl.ToolTipText = objControl.Caption
        Exit Sub
    End If
    
    intReportState = GetCurrReportState(mlngAdviceID)
    
    If intReportState = 1 Then                      '编辑中
        objControl.Caption = "书写"
        objControl.Enabled = CheckPopedom(mstrPrivs, "PACS报告书写") Or CheckPopedom(mstrPrivs, "PACS报告修订") Or CheckPopedom(mstrPrivs, "PACS报告终审")
        
    ElseIf intReportState = 2 Or intReportState = 5 Then                  '已诊断，诊断驳回
        If CheckPopedom(mstrPrivs, "PACS报告修订") Then '“PACS报告修订”权限，在报告文档编辑器中当 审核 权限使用
            objControl.Caption = "审核"
        Else
            objControl.Caption = "修订"
        End If
        
    ElseIf intReportState = 3 Or intReportState = 6 Then                  '已审核，审核驳回
        objControl.Caption = "审核"
        objControl.Enabled = CheckPopedom(mstrPrivs, "PACS报告终审")
        
    ElseIf intReportState = 4 Then                  '已终审
        objControl.Caption = "查阅"
        objControl.Enabled = True
    End If
    
    objControl.ToolTipText = objControl.Caption
    Exit Sub
ErrorHand:
    objControl.Caption = "书写"
    objControl.ToolTipText = objControl.Caption
End Sub

Public Function zlInit(ByVal lngSysNo As Long, ByVal lngDeptID As Long, ByVal strCode As String, ByVal strPrivs As String, objOwner As Object) As Boolean
'初始化窗体内容
    Dim strUser As String, strPwd As String, strServer As String
    Dim lngIndex As Long
    Dim blnResult As Boolean
    
On Error GoTo ErrorHand
    Set mobjOwner = objOwner
    mstrPrivs = strPrivs
    mlngDeptID = lngDeptID
    mlngModule = strCode
    
    strUser = UCase(gcnOracle.Properties(23))
    strPwd = gcnOracle.Properties(24)
    strServer = gcnOracle.Properties(8) 'UCase(Split(gcnOracle.Properties(8), "=")(2))
    
    lngIndex = InStr(strServer, "(DESCRIPTION=")
    If lngIndex >= 1 Then
        strServer = Mid(strServer, lngIndex, 255)
    Else
        strServer = UCase(Split(gcnOracle.Properties(8), "=")(2))
    End If
    
    mblnIsShowLastState = zlDatabase.GetPara("最终状态显示", glngSys, mlngModule, 0) = "1"
    mblnIsNoAskPrint = zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1"
    mblnCanPrint = GetDeptPara(mlngDeptID, "平诊需审核才能打报告") = "1"             '平诊需要审核才能打印 =true
    mintCompletePrint = Val(GetDeptPara(mlngDeptID, "终审后直接打印", 0))
    
    Set mobjInterfaceEvent = New ZLPacsVBCommon.clsEvent
    Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS智能报告编辑器接口部件zlPacsInterfaceC")
    
    If mobjPacsInterface Is Nothing Or mobjInterfaceEvent Is Nothing Then Exit Function
    If InitRegister = False Then Exit Function
    
    BUGEX "zlInit InitRegister " & gstrUserPswd
    
    blnResult = mobjPacsInterface.Init(mobjInterfaceEvent, lngSysNo, strCode, objOwner.Caption, gstrUserName, gstrUserPswd, gstrServerName, lngDeptID)
    
    '设置编辑器特性，不同版本特性方法可能有差异，因此需要忽略调用异常
    Call SetAttribute
    
    If blnResult Then
        Call SetReportFinalShowMode(mblnIsShowLastState)
        Call SetIsNoAskPrint(mblnIsNoAskPrint)
    End If
    
    zlInit = blnResult
    Exit Function
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub SetAttribute()
On Error Resume Next
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Call mobjPacsInterface.SetOwner(mobjOwner.hWnd)
err.Clear
End Sub

Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngPatientID As Long, ByVal lngStudyState As Long)
'同步医嘱信息
    mlngPatientId = lngPatientID
    mlngAdviceID = lngAdviceID
    mlngStudyState = lngStudyState
    
End Sub

Public Sub zlEditReport()
'书写报告
    Call OpenFormForEditReport
End Sub


Public Sub zlRefreshFace(ByVal blnCanPrint As Boolean, ByVal blnForceRefresh As Boolean)
'刷新
    mblnDirectPrint = blnCanPrint
    
    Call RefreshDocReportForm(blnForceRefresh)
    
    BUGEX "zlRefreshFaceRich mblnDirectPrint = " & mblnDirectPrint
    BUGEX "zlRefreshFaceRich mblnCanPrint = " & mblnCanPrint
End Sub


Public Function zlGetForm() As Long
'获取报告窗口句柄
    If mobjPacsInterface Is Nothing Then Exit Function
    
    If mlngReportFormHandle <= 0 Then mlngReportFormHandle = mobjPacsInterface.GetReportFormHandle(mlngAdviceID, mlngPatientId, mlngStudyState, mblnDirectPrint)
    zlGetForm = mlngReportFormHandle
    mblnVisible = True
End Function


Public Sub CloseForms()
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.CloseForms
    mblnVisible = False
End Sub

Private Sub OpenFormForReportPrint(ByVal strDocID As String, ByVal blnPreview As Boolean)
'预览，打印
    Dim strResultInput As String
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrorHand
    If Not blnPreview Then
        '打印前录入结果
        If Val(GetDeptPara(mlngDeptID, "诊断结果提示类型", 0)) = 2 Then
            strResultInput = ""
            If GetDeptPara(mlngDeptID, "无影像诊断为阴性", 0) = "1" Then '无影像诊断为阴性  -无提示自动标记
                Call EvaluatResult(mlngAdviceID, "0")
            End If
                
            strSql = "Select B.危急状态, A.结果阳性, B.影像质量, A.报告质量, B.符合情况 " & _
                     "From 影像报告记录 A, 影像检查记录 B " & _
                     "Where A.Id=[1] and A.医嘱Id = B.医嘱Id"
                     
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取结果阳性", strDocID)
        
            If IsNull(rsTemp!危急状态) And Val(GetDeptPara(mlngDeptID, "危急情况判断", 0)) <> 0 Then strResultInput = "危急状态|"
            If IsNull(rsTemp!结果阳性) And Not GetDeptPara(mlngDeptID, "忽略结果阴阳性", 0) = "1" Then strResultInput = strResultInput & "结果阳性|"
            If IsNull(rsTemp!影像质量) And NVL(GetDeptPara(mlngDeptID, "影像质量等级", "甲,乙")) <> "" And Val(GetDeptPara(mlngDeptID, "影像质量判定", "0")) <> 0 And CheckPopedom(mstrPrivs, "影像质控") Then strResultInput = strResultInput & "影像质量|"
            If IsNull(rsTemp!报告质量) And NVL(GetDeptPara(mlngDeptID, "报告质量等级", "甲,乙")) <> "" And Val(GetDeptPara(mlngDeptID, "报告质量判定", 0)) <> 0 And CheckPopedom(mstrPrivs, "报告质控") Then strResultInput = strResultInput & "报告质量|"
            If IsNull(rsTemp!符合情况) And Val(GetDeptPara(mlngDeptID, "符合情况判定", 0)) <> 0 Then strResultInput = strResultInput & "符合情况|"
            
            If strResultInput <> "" Then Call PromptResultRich(mlngAdviceID, strDocID, mlngModule, Me, mlngDeptID, strResultInput)
        End If
    End If
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.PrintReport(mobjOwner.hWnd, strDocID, blnPreview)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Function OpenFormForEditReport() As String
'书写--即独立报告窗口模式
'返回:错误信息
    If mobjPacsInterface Is Nothing Then Exit Function
    Call ShowErrMessage(mobjPacsInterface.OpenFormForEditReport(mobjOwner.hWnd, mlngAdviceID, mlngPatientId, mlngStudyState, mblnDirectPrint))
End Function


Private Function DeleteReport() As String
'删除报告
'返回:错误信息
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = GetCurrDocId(mlngAdviceID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    Call ShowErrMessage(mobjPacsInterface.DeleteReport(mobjOwner.hWnd, strDocID))
End Function


Private Function RefreshDocReportForm(ByVal isForce As Boolean) As String
'刷新嵌入式报告窗口
    If mobjPacsInterface Is Nothing Then Exit Function
    If mblnRefreshState = False Then
        isForce = True
        mblnRefreshState = True
    End If
    Call ShowErrMessage(mobjPacsInterface.RefreshReportForm(mlngAdviceID, mlngPatientId, mlngDeptID, mlngStudyState, mblnDirectPrint, isForce))
End Function

Private Sub ShowErrMessage(ByVal strErrMessage As String)
    If strErrMessage = "" Then Exit Sub
    
    MsgboxEx mobjOwner, strErrMessage, vbExclamation, gstrSysName
End Sub

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定Id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function


Private Sub Class_Terminate()
    If Not mobjInterfaceEvent Is Nothing Then Set mobjInterfaceEvent = Nothing
    If Not mobjPacsInterface Is Nothing Then Set mobjPacsInterface = Nothing
    
    mblnVisible = False
End Sub

Public Sub CreatePacsInterface()
    If mobjPacsInterface Is Nothing Then Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "PACS智能报告编辑器接口部件zlPacsInterfaceC")
End Sub

'事件
Private Sub mobjInterfaceEvent_OnDataRefresh(ByVal strDocID As String)
    RaiseEvent DataRefresh(strDocID)
End Sub

Private Sub mobjInterfaceEvent_OnDocPluginAction(ByVal strPluginID As String, ByVal strPluginName As String, ByVal lngActionType As Long, ByVal strData As String, ByVal strTag As String)
    RaiseEvent DocPluginAction(lngActionType, strData, strTag)
End Sub

Private Sub mobjInterfaceEvent_OnImageLoadFromFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImagePath As String, blnResult As Boolean)
'从FTP上下载报告图加载到报告中
    Dim objFSO As New Scripting.FileSystemObject
    Dim i As Integer
    Dim strImages() As String
    Dim strLoadFailureImgs As String
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    Dim blnLoadResult As Boolean
    
On Error GoTo ErrorHand
    mblnIsFirstImgDownLoadFiald = False
    
    If lngOrderID <= 0 Or strImageName = "" Or strImagePath = "" Then
        blnResult = True
        Exit Sub
    End If
    
    strImages = Split(strImageName, "★★")
    
    '创建本地路径
    If Not objFSO.FolderExists(strImagePath) Then Call MkLocalDir(strImagePath)
    
    '初始化FTP信息
    If Not InitFtpInf(strDocID) Then Exit Sub
    
    For i = 0 To UBound(strImages)
        strVirtualPath = Replace(mobjFtpInf.strFtpDir & mstrSubDir, "\", "/")
        strLocalFileName = Replace(strImagePath & "\" & strImages(i), "/", "\")
        
        '如果本地不存在则从FTP上下载
        If Not objFSO.FileExists(strLocalFileName) Then
            '连接FTP
            If Not ConnFtp Then Exit Sub
            
            '从存储设备上获取图像
            If mblnIsFirstImgDownLoadFiald Then
                blnLoadResult = mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strImages(i)) = 1
            Else
                blnLoadResult = FuncDownloadFile(mobjFtp, strVirtualPath, strLocalFileName, strImages(i))
            End If
            
            If Not blnLoadResult Then strLoadFailureImgs = strLoadFailureImgs & "、" & strImages(i)
        End If
    Next
    
    If Mid(strLoadFailureImgs, 2) <> "" Then
        If UBound(Split(Mid(strLoadFailureImgs, 2), "、")) = UBound(strImages) Then
            MsgBox "图像下载失败，可能是服务器上不存在图像或网络不稳定造成！" & vbCrLf & "图像名为：" & Mid(strLoadFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = False
        Else
            MsgBox "部分图像下载失败，可能是服务器上不存在图像或网络不稳定造成！" & vbCrLf & "图像名为：" & Mid(strLoadFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = True
        End If
    Else
        blnResult = True
    End If
    
    Exit Sub
ErrorHand:
    blnResult = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub mobjInterfaceEvent_OnImageSaveToFTP(ByVal lngOrderID As Long, ByVal strDocID As String, ByVal strImageName As String, ByVal strImageBase64 As String, ByVal strFTPDeviceNo As String, ByVal strRptCreateDate As String, blnResult As Boolean)
'将报告中的报告图保存到FTP
    Dim objFSO As New Scripting.FileSystemObject
    Dim strImages() As String
    Dim strImagesBase64() As String
    Dim strNeedUpImgs As String
    Dim strUpFailureImgs As String
    Dim i As Integer
    Dim blnLoadResult As Boolean
    
On Error GoTo ErrorHand
    mblnIsFirstImgUpLoadFiald = False
    
    If lngOrderID <= 0 Or strImageName = "" Then
        blnResult = True
        Exit Sub
    End If
    
    strImages = Split(strImageName, "★★")
    strImagesBase64 = Split(strImageBase64, "★★")
    
    '初始化FTP信息
    If Not InitFtpInf(strDocID, strFTPDeviceNo, Format(strRptCreateDate, "YYYYMMDD")) Then Exit Sub
    
    '创建本地路径
    If Not objFSO.FolderExists(mstrDestMainDir & mstrSubDir) Then Call MkLocalDir(mstrDestMainDir & mstrSubDir)
    
    '连接FTP
    If Not ConnFtp Then Exit Sub
    
    For i = 0 To UBound(strImages)
        If strImagesBase64(i) = "-1" Then
            '为空说明报告图已经删除,因此删除FTP服务器上对应的图像
            Call mobjFtp.FuncDelFile(mobjFtpInf.strFtpDir & mstrSubDir, strImages(i))
            
            '如果本地存在则删除
            If objFSO.FileExists(mstrDestMainDir & mstrSubDir & "\" & strImages(i)) Then Call objFSO.DeleteFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), True)
        Else
            strNeedUpImgs = strNeedUpImgs & "、" & strImages(i) '记录需要上传的图像名字
            
            '将图像串生成文件
            If WritTextFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImagesBase64(i)) Then
                '将图像文件上传到FTP存储设备
                mobjFtp.FuncFtpMkDir "/", mobjFtpInf.strFtpDir & mstrSubDir
                
                If mblnIsFirstImgUpLoadFiald Then
                    blnLoadResult = mobjFtp.FuncUploadFile(mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                Else
                    blnLoadResult = FuncUploadFile(mobjFtp, mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                End If
                
                If Not blnLoadResult Then strUpFailureImgs = strUpFailureImgs & "、" & strImages(i)
            Else
                strUpFailureImgs = strUpFailureImgs & "、" & strImages(i)
            End If
        End If
    Next
    
    '断开FTP连接
    mobjFtp.FuncFtpDisConnect
    
    If Mid(strUpFailureImgs, 2) <> "" Then
        If UBound(Split(Mid(strUpFailureImgs, 2), "、")) = UBound(Split(Mid(strNeedUpImgs, 2), "、")) Then
            MsgBox "图像上传至FTP服务器失败，可能是网络不稳定造成！" & vbCrLf & "图像名为：" & Mid(strUpFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = False
        Else
            MsgBox "部分图像上传至FTP服务器失败，可能是网络不稳定造成！" & vbCrLf & "图像名为：" & Mid(strUpFailureImgs, 2), vbExclamation, gstrSysName
            blnResult = True
        End If
    Else
        blnResult = True
    End If
    
    Exit Sub
ErrorHand:
    blnResult = False
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub mobjInterfaceEvent_OnReportAction(ByVal lngReportEvtAction As ZLPacsVBCommon.TReportEvtActionType, ByVal lngOrderID As Long, ByVal strDocID As String, ByVal lngReportState As Long)
    On Error GoTo ErrorHand
    
    Select Case lngReportEvtAction
        Case TReportEvtActionType.ratAfterClosed
            RaiseEvent AfterClosed(lngOrderID, strDocID)
            Call RefreshDocReportForm(True)
            
        Case TReportEvtActionType.ratAfterDeleted
            RaiseEvent AfterDeleted(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterOpen
            RaiseEvent AfterOpen(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterPrinted
            RaiseEvent AfterPrinted(lngOrderID, strDocID)
        
        Case TReportEvtActionType.ratAfterReportStateChanged
            '终审后直接打印
            BUGEX "mobjInterfaceEvent_OnReportAction mintCompletePrint = " & mintCompletePrint & "  reportState = " & lngReportState
            If mintCompletePrint = 1 And lngReportState = 4 Then
                Call OpenFormForReportPrint(strDocID, False)
            End If
            
            '诊断驳回或者审核驳回后，发起保存事件，修改检查状态
            If lngReportState = 5 Or lngReportState = 6 Then
                'reportState0-普通保存，1-诊断签名，2-审核（终审）签名
                RaiseEvent AfterSaved(lngOrderID, strDocID, lngReportState)
            End If
        Case TReportEvtActionType.ratAfterSaved
            'reportState0-普通保存，1-诊断签名，2-审核（终审）签名
            RaiseEvent AfterSaved(lngOrderID, strDocID, lngReportState)
        
        Case TReportEvtActionType.ratBeforeEdit
            RaiseEvent BeforeEdit(lngOrderID, strDocID)
            
    End Select
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function FuncDownloadFile(objFtp As clsFtp, strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Boolean
    If objFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strRemoteFileName) <> 0 Then
        If MsgBox("图像下载失败，可能是服务器上不存在图像或网络不稳定造成！", vbRetryCancel, gstrSysName) = vbRetry Then
            Call FuncDownloadFile(objFtp, strVirtualPath, strLocalFileName, strRemoteFileName)
        Else
            mblnIsFirstImgDownLoadFiald = True
            FuncDownloadFile = False
            Exit Function
        End If
    End If
    
    FuncDownloadFile = True
End Function

Private Function FuncUploadFile(objFtp As clsFtp, strVirtualPath As String, strLocalFileName As String, strRemoteFileName As String) As Boolean
    If objFtp.FuncUploadFile(strVirtualPath, strLocalFileName, strRemoteFileName) <> 0 Then
        If MsgBox("图像上传至FTP服务器失败，可能是网络不稳定造成！", vbRetryCancel, gstrSysName) = vbRetry Then
            Call FuncUploadFile(objFtp, strVirtualPath, strLocalFileName, strRemoteFileName)
        Else
            mblnIsFirstImgUpLoadFiald = True
            FuncUploadFile = False
            Exit Function
        End If
    End If
    
    FuncUploadFile = True
End Function

Private Function ConnFtp(Optional ByVal blnRefresh As Boolean = False) As Boolean
    If mobjFtp.hConnection = 0 Or blnRefresh Then
        '连接FTP存储设备
        If mobjFtp.FuncFtpConnect(mobjFtpInf.strFtpIp, mobjFtpInf.strFTPUser, mobjFtpInf.strFTPPwd) = 0 Then
            MsgBox "FTP服务连接失败，请检查网络或FTP设置." + vbCrLf + "FTP:" + mobjFtpInf.strFtpIp + "  USER:" + mobjFtpInf.strFTPUser, vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

Private Function InitFtpInf(ByVal strDocID As String, Optional ByVal strDeviceNO As String, Optional ByVal strCreateDate As String) As Boolean
'获取报告图像的FTP信息
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    If mstrTmpReportId = strDocID Then
        InitFtpInf = True
        Exit Function
    End If
    
    mstrTmpReportId = strDocID
    
    If strDeviceNO = "" Then
        strSql = "Select 'ReportImages/' || to_Char(b.创建时间,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
                 "a.设备号, a.FTP用户名, a.FTP密码, a.IP地址,'/'||a.Ftp目录||'/' As Root " & _
                 "From 影像设备目录 a, 影像报告记录 b where a.设备号 = b.设备号 And b.id = [1]"
                 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", strDocID)
    Else
        strSql = "Select 'ReportImages/'||[1]||'/'||[2]||'/' As URL," & _
                 "设备号, FTP用户名, FTP密码, IP地址,'/'||Ftp目录||'/' As Root From 影像设备目录 where 设备号 = [3]"
                 
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "获取FTP信息", strCreateDate, strDocID, strDeviceNO)
    End If
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInf.strFtpDir = NVL(rsTmp("Root"))
    mobjFtpInf.strFtpIp = NVL(rsTmp("IP地址"))
    mobjFtpInf.strFTPPwd = NVL(rsTmp("FTP密码"))
    mobjFtpInf.strFTPUser = NVL(rsTmp("FTP用户名"))
    mobjFtpInf.strDeviceId = Trim(NVL(rsTmp("设备号")))
    mstrSubDir = NVL(rsTmp("URL"))
    mstrDestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInf = True
End Function

Public Function GetCurrDocId(ByVal lngOrderID As Long) As String
'获取当前选中的报告的Id
On Error Resume Next
    If mobjPacsInterface Is Nothing Then Exit Function
    GetCurrDocId = mobjPacsInterface.GetCurrDocId(lngOrderID)
End Function

Public Sub EvaluatReportQuality(ByVal strDocID As String, ByVal strResult As String, _
     Optional ByVal lngOrderID As Long = 0)
'--------------------------------------------------------
'功  能：评定报告质量
'参数： strDocID -- 报告ID
'       strResult -- 报告质量级别
'       lngOrderID -- 医嘱ID，可选参数，如果strDocID=""，则通过医嘱ID提取报告ID
'返回：无
'-------------------------------------------------------
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    If Trim(strDocID) = "" And lngOrderID <> 0 Then
        strDocID = GetCurrDocId(lngOrderID)
    End If
    
    If Trim(strDocID) = "" Then Exit Sub
    Call mobjPacsInterface.EvaluatReportQuality(strDocID, strResult)
End Sub

Public Sub EvaluatResult(ByVal strDocID As String, ByVal strResult As String)
'评定阴阳性
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.EvaluatResult(strDocID, strResult)
End Sub

Public Sub ReportRelease(ByVal lngOrderID As Long, ByVal intType As Integer)
'报告发放或回收
'intType:0--报告回收，1--报告发放
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.ReportRelease(mobjOwner.hWnd, lngOrderID, intType)
End Sub

Public Function GetReportReleaseState(ByVal lngOrderID As Long) As Integer
'返回报告报告发放情况
'0-全部都未发放，1-部分发放(当前选中的报告未发放)，2-部分发放(当前选中的报告已发放)，3-全部都已发放
    If mobjPacsInterface Is Nothing Then Exit Function
    GetReportReleaseState = mobjPacsInterface.GetReportReleaseState(lngOrderID, mlngPatientId)
End Function

Public Function GetCurrReportState(ByVal lngOrderID As Long) As Integer
'返回当前报告的报告状态
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = Me.GetCurrDocId(lngOrderID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    GetCurrReportState = mobjPacsInterface.GetCurrReportState(strDocID)
End Function

Public Sub UpdateImageVideoState(ByVal lngEventType As Integer, ByVal lngAdviceID As Long, ByVal other As Variant)
'刷新前台和后台采集的图像
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Select Case lngEventType
        Case TVideoEventType.vetUpdateImg, TVideoEventType.vetCaptureFirstImg, TVideoEventType.vetDelAllImg, _
             TVideoEventType.vetAfterUpdateImg, TVideoEventType.vetExportImage, TVideoEventType.vetImportImage, TVideoEventType.vetImgDeled
            '报告Id传空，医嘱Id为0表示刷新后台采集的图像
            Call mobjPacsInterface.ExecutePlugInCmd(lngAdviceID, "", "图像刷新", lngAdviceID)
        Case TVideoEventType.vetLockStudy
            Call mobjPacsInterface.ExecutePlugInCmd(0, "", "锁定检查", "【" & other & "】视频采集")
        Case TVideoEventType.vetUnLockStudy
            Call mobjPacsInterface.ExecutePlugInCmd(0, "", "解锁检查", "视频采集")
    End Select
End Sub

Private Sub SetReportFinalShowMode(ByVal blnFinalShowMode As Boolean)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetReportFinalShowMode(blnFinalShowMode)
End Sub

Private Sub SetIsNoAskPrint(ByVal blnNoAskPrint As Boolean)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetIsNoAskPrint(blnNoAskPrint)
End Sub

'获取当前报告的创建人
Public Function GetDocCreator(ByVal lngOrderID As Long)
    Dim strDocID As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocID = Me.GetCurrDocId(lngOrderID)
    
    If Trim(strDocID) = "" Then Exit Function
    
    GetDocCreator = mobjPacsInterface.GetDocCreator(strDocID)
End Function

'报告执行
Public Sub SetDocCreator(ByVal strCreatorName)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetRecordDoctor(strCreatorName)
End Sub

Public Property Get Visible() As Boolean
    Visible = mblnVisible
End Property

Public Sub RefreshVideo()
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Call mobjPacsInterface.ExecutePlugInCmd(0, "", "视频刷新", "")
End Sub

Public Sub AddXWReportImgToRichRPT(ByVal lngOrderID As Long, strDcmImgPath As String)
    '将文件路径传到报告文档编辑器的缩略图插件中，在插件中在对图像进行操作
    If Trim(strDcmImgPath) = "" Then Exit Sub
    
    Call mobjPacsInterface.ExecutePlugInCmd(lngOrderID, "", "XW图像添加", strDcmImgPath)
End Sub
