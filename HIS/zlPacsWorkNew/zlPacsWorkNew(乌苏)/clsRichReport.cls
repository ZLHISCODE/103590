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

Public Event BeforeEdit(ByVal lngOrderID As Long, ByVal strDocId As String)
Public Event AfterOpen(ByVal lngOrderID As Long, ByVal strDocId As String)
Public Event AfterSaved(ByVal lngOrderID As Long, ByVal strDocId As String, ByVal lngSaveType As Long)
Public Event AfterClosed(ByVal lngOrderID As Long, ByVal strDocId As String)
Public Event AfterPrinted(ByVal lngOrderID As Long, ByVal strDocId As String)
Public Event AfterDeleted(ByVal lngOrderID As Long, ByVal strDocId As String)
Public Event DataRefresh(ByVal strDocInfo As String)
Public Event DocPluginAction(ByVal actionType As Long, ByVal data As String, ByVal tag As String)

Private WithEvents mobjInterfaceEvent As ZLPacsVBCommon.clsEvent
Attribute mobjInterfaceEvent.VB_VarHelpID = -1
Private mobjPacsInterface As Object
Attribute mobjPacsInterface.VB_VarHelpID = -1
Private mobjOwner As Object
Private mstrPrivs As String
Private mlngModule As Long
Private mlngDeptId As Long
Private mlngAdviceID As Long
Private mlngTempOrderId As Long
Private mlngPatientID As Long
Private mlngStudyState As Long
Private mblnCanPrint As Boolean     '平诊需审核后才能打印
Private mblnDirectPrint As Boolean  '紧急病人和绿色通道病人可直接打印

Private mblnRefreshState As Boolean
Private mobjPacsCore As zl9PacsCore.clsViewer
Private mObjActiveMenuBar As CommandBars

Private mobjFtp As New clsFtp
Private mobjBakFtp As New clsFtp
Private mobjFtpInf As TFtpDeviceInf
Private mobjBakFtpInf As TFtpDeviceInf
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


Public Sub zlExecuteMenu(ByVal lngMenuId As Long, Optional ByVal strDocId As String = 0)
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
            Call OpenFormForReportPrint(strDocId, True)
            
        Case conMenu_File_Print                     '打印
            Call OpenFormForReportPrint(strDocId, False)
            
        Case conMenu_File_BatPrint                  '批量打印
            Call OpenFormForReportPrint(strDocId, False)
            
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
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strDocId As String
    
On Error GoTo ErrorHand
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.EditLockedInfo(GetCurrDocId(mlngAdviceID), "")
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Public Sub zlUpdateMenu(ByVal objControl As XtremeCommandBars.ICommandBarControl)
    Dim lngReportState As Long
    
'更新编辑器菜单
On Error GoTo ErrorHand
    If objControl Is Nothing Then Exit Sub
    If mobjPacsInterface Is Nothing Then Exit Sub
    
    Select Case objControl.ID
        Case conMenu_PacsReport_FinalShowMode
            objControl.Checked = mblnIsShowLastState
            
        Case conMenu_File_Print, conMenu_File_Preview, conMenu_File_BatPrint     '打印报告,预览报告,批量打印
            If InStr(mstrPrivs, "PACS报告打印") <= 0 Then
                objControl.Visible = False
            Else
                objControl.Visible = True
            End If
            
            If objControl.ID <> conMenu_File_BatPrint Then
                lngReportState = GetCurrReportState(mlngAdviceID)
                
                objControl.Enabled = GetCurrDocId(mlngAdviceID) <> "" And lngReportState > 1 And lngReportState < 5
                
                If objControl.Enabled Then
                    If mblnDirectPrint Then '绿色通道和紧急病人，诊断签名后即可打印
                        objControl.Enabled = lngReportState > 1 And lngReportState < 5
                    ElseIf mblnCanPrint Then '平诊需审核后才能打印
                        objControl.Enabled = lngReportState = 4
                    End If
                End If
            End If
            
        Case conMenu_PacsReport_Open                '书写
            Call InitEditButtonCaption(objControl)
            
        Case conMenu_Edit_Delete                    '删除
            objControl.Visible = (GetCurrDocId(mlngAdviceID) <> "" And (InStr(mstrPrivs, "PACS报告书写") > 0 Or InStr(mstrPrivs, "PACS报告删除") > 0))
            If objControl.Visible Then objControl.Visible = (InStr(mstrPrivs, "PACS他人报告") > 0 Or GetDocCreator(mlngAdviceID) = UserInfo.姓名)
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
    Dim blnResult As Boolean
    
On Error GoTo ErrorHand
    Set mobjOwner = objOwner
    mstrPrivs = strPrivs
    mlngDeptId = lngDeptID
    mlngModule = strCode
    
    strUser = UCase(gcnOracle.Properties(23))
    strPwd = gcnOracle.Properties(24)
    strServer = UCase(Split(gcnOracle.Properties(8), "=")(2))
    
    mblnIsShowLastState = zlDatabase.GetPara("最终状态显示", glngSys, mlngModule, 0) = "1"
    mblnIsNoAskPrint = zlDatabase.GetPara("NoAsk", glngSys, 1070, 0) = "1"
    mblnCanPrint = GetDeptPara(mlngDeptId, "平诊需审核才能打报告") = "1"             '平诊需要审核才能打印 =true
    mintCompletePrint = Val(GetDeptPara(mlngDeptId, "终审后直接打印", 0))
    
    If mobjPacsInterface Is Nothing Then Exit Function
    blnResult = mobjPacsInterface.Init(mobjInterfaceEvent, lngSysNo, strCode, objOwner.Caption, strUser, strPwd, strServer, lngDeptID)
    
    If blnResult Then Call SetReportFinalShowMode(mblnIsShowLastState)
    
    zlInit = blnResult
    Exit Function
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function


Public Sub zlUpdateAdviceInf(ByVal lngAdviceID As Long, ByVal lngPatientID As Long, ByVal lngStudyState As Long)
'同步医嘱信息
    mlngPatientID = lngPatientID
    mlngAdviceID = lngAdviceID
    mlngStudyState = lngStudyState
    
    mblnRefreshState = True
End Sub

Public Sub zlEditReport()
'书写报告
    Call OpenFormForEditReport
End Sub


Public Sub zlRefreshFace(ByVal blnCanPrint As Boolean)
'刷新
    Call RefreshDocReportForm(True)
    
    mblnDirectPrint = blnCanPrint
    BUGEX "zlRefreshFaceRich mblnDirectPrint = " & mblnDirectPrint
    BUGEX "zlRefreshFaceRich mblnCanPrint = " & mblnCanPrint
End Sub


Public Function zlGetForm() As Long
'获取报告窗口句柄
    If mobjPacsInterface Is Nothing Then Exit Function
    zlGetForm = mobjPacsInterface.GetReportFormHandle(mlngAdviceID, mlngPatientID, mlngStudyState, mblnDirectPrint)
End Function


Public Sub CloseForms()
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.CloseForms
End Sub

Private Sub OpenFormForReportPrint(ByVal strDocId As String, ByVal blnPreview As Boolean)
'预览，打印
    Dim strResultInput As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrorHand
    If Not blnPreview Then
        '打印前录入结果
        If Val(GetDeptPara(mlngDeptId, "诊断结果提示类型", 0)) = 2 Then
            strResultInput = ""
            If GetDeptPara(mlngDeptId, "无影像诊断为阴性", 0) = "1" Then '无影像诊断为阴性  -无提示自动标记
                Call EvaluatResult(mlngAdviceID, "0")
            End If
                
            strSQL = "Select B.危急状态, A.结果阳性, B.影像质量, A.报告质量, B.符合情况 " & _
                     "From 影像报告记录 A, 影像检查记录 B " & _
                     "Where A.Id=[1] and A.医嘱Id = B.医嘱Id"
                     
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取结果阳性", strDocId)
        
            If IsNull(rsTemp!危急状态) And Val(GetDeptPara(mlngDeptId, "危急情况判断", 0)) <> 0 Then strResultInput = "危急状态|"
            If IsNull(rsTemp!结果阳性) And Not GetDeptPara(mlngDeptId, "忽略结果阴阳性", 0) = "1" Then strResultInput = strResultInput & "结果阳性|"
            If IsNull(rsTemp!影像质量) And Nvl(GetDeptPara(mlngDeptId, "影像质量等级", "甲,乙")) <> "" And Val(GetDeptPara(mlngDeptId, "影像质量判定", "0")) <> 0 And InStr(mstrPrivs, "影像质控") > 0 Then strResultInput = strResultInput & "影像质量|"
            If IsNull(rsTemp!报告质量) And Nvl(GetDeptPara(mlngDeptId, "报告质量等级", "甲,乙")) <> "" And Val(GetDeptPara(mlngDeptId, "报告质量判定", 0)) <> 0 And InStr(mstrPrivs, "报告质控") > 0 Then strResultInput = strResultInput & "报告质量|"
            If IsNull(rsTemp!符合情况) And Val(GetDeptPara(mlngDeptId, "符合情况判定", 0)) <> 0 Then strResultInput = strResultInput & "符合情况|"
            
            If strResultInput <> "" Then Call PromptResultRich(mlngAdviceID, strDocId, mlngModule, Me, mlngDeptId, strResultInput)
        End If
    End If
    
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.PrintReport(mobjOwner.hWnd, strDocId, blnPreview, mblnIsNoAskPrint)
    
    Exit Sub
ErrorHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub


Private Function OpenFormForEditReport() As String
'书写--即独立报告窗口模式
'返回:错误信息
    If mobjPacsInterface Is Nothing Then Exit Function
    Call ShowErrMessage(mobjPacsInterface.OpenFormForEditReport(mobjOwner.hWnd, mlngAdviceID, mlngPatientID, mlngStudyState, mblnDirectPrint))
End Function


Private Function DeleteReport() As String
'删除报告
'返回:错误信息
    If mobjPacsInterface Is Nothing Then Exit Function
    Call ShowErrMessage(mobjPacsInterface.DeleteReport(mobjOwner.hWnd, GetCurrDocId(mlngAdviceID)))
End Function


Private Function RefreshDocReportForm(ByVal isPorce As Boolean) As String
'刷新嵌入式报告窗口
    If mobjPacsInterface Is Nothing Then Exit Function
    Call ShowErrMessage(mobjPacsInterface.RefreshReportForm(mlngAdviceID, mlngPatientID, mlngDeptId, mlngStudyState, mblnDirectPrint, isPorce))
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


Private Sub Class_Initialize()
    If mobjInterfaceEvent Is Nothing Then Set mobjInterfaceEvent = New ZLPacsVBCommon.clsEvent
    Set mobjPacsInterface = DynamicCreate("ZLSoft.ZLPacs.Interface.PacsInterface", "报告接口部件PacsInterface")
End Sub


Private Sub Class_Terminate()
    If Not mobjInterfaceEvent Is Nothing Then Set mobjInterfaceEvent = Nothing
    If Not mobjPacsInterface Is Nothing Then Set mobjPacsInterface = Nothing
End Sub


'事件
Private Sub mobjInterfaceEvent_OnDataRefresh(ByVal strDocId As String)
    RaiseEvent DataRefresh(strDocId)
End Sub

Private Sub mobjInterfaceEvent_OnDocPluginAction(ByVal strPluginID As String, ByVal strPluginName As String, ByVal lngActionType As Long, ByVal strData As String, ByVal strTag As String)
    RaiseEvent DocPluginAction(lngActionType, strData, strTag)
End Sub

Private Sub mobjInterfaceEvent_OnImageLoadFromFTP(ByVal lngOrderID As Long, ByVal strDocId As String, ByVal strImageName As String, ByVal strImagePath As String, blnResult As Boolean)
'从FTP上下载报告图加载到报告中
    Dim objFso As New Scripting.FileSystemObject
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
    If Not objFso.FolderExists(strImagePath) Then Call MkLocalDir(strImagePath)
    
    '初始化FTP信息
    If Not InitFtpInfo(lngOrderID) Then Exit Sub
    
    For i = 0 To UBound(strImages)
        strVirtualPath = Replace(mobjFtpInf.strFtpDir & mstrSubDir, "\", "/")
        strLocalFileName = Replace(strImagePath & "\" & strImages(i), "/", "\")
        
        '如果本地不存在则从FTP上下载
        If Not objFso.FileExists(strLocalFileName) Then
            '连接FTP
            If Not ConnFtp Then Exit Sub
            
            '从存储设备上获取图像
            If mblnIsFirstImgDownLoadFiald Then
                blnLoadResult = mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strImages(i)) = 1
            Else
                blnLoadResult = FuncDownloadFile(mobjFtp, strVirtualPath, strLocalFileName, strImages(i))
            End If
            
            If Not blnLoadResult Then
                strVirtualPath = Replace(mobjBakFtpInf.strFtpDir & objFso.GetParentFolderName(mstrSubDir), "\", "/")
                
                '连接备份设备
                If mobjBakFtpInf.strDeviceId <> "" Then
                    If mobjBakFtp.hConnection = 0 Then Call mobjBakFtp.FuncFtpConnect(mobjBakFtpInf.strFTPIP, mobjBakFtpInf.strFTPUser, mobjBakFtpInf.strFTPPwd)
                       
                    '从备份设备上获取图像
                    If mobjBakFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, strImages(i)) <> 0 Then
                        strLoadFailureImgs = strLoadFailureImgs & "、" & strImages(i)
                    End If
                Else
                    strLoadFailureImgs = strLoadFailureImgs & "、" & strImages(i)
                End If
            End If
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
End Sub

Private Sub mobjInterfaceEvent_OnImageSaveToFTP(ByVal lngOrderID As Long, ByVal strDocId As String, ByVal strImageName As String, ByVal strImageBase64 As String, blnResult As Boolean)
'将报告中的报告图保存到FTP
    Dim objFso As New Scripting.FileSystemObject
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
    If Not InitFtpInfo(lngOrderID) Then Exit Sub
    
    '创建本地路径
    If Not objFso.FolderExists(mstrDestMainDir & mstrSubDir) Then Call MkLocalDir(mstrDestMainDir & mstrSubDir)
    
    '连接FTP
    If Not ConnFtp Then Exit Sub
    
    For i = 0 To UBound(strImages)
        If strImagesBase64(i) = "-1" Then
            '为空说明报告图已经删除,因此删除FTP服务器上对应的图像
            Call mobjFtp.FuncDelFile(mobjFtpInf.strFtpDir & mstrSubDir, strImages(i))
            If mobjBakFtpInf.strDeviceId <> "" Then Call mobjBakFtp.FuncDelFile(mobjBakFtpInf.strFtpDir & mstrSubDir, strImages(i))
            
            '如果本地存在则删除
            If objFso.FileExists(mstrDestMainDir & mstrSubDir & "\" & strImages(i)) Then Call objFso.DeleteFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), True)
        Else
            strNeedUpImgs = strNeedUpImgs & "、" & strImages(i) '记录需要上传的图像名字
            
            '将图像串生成文件
            If WritTextFile(mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImagesBase64(i)) Then
                '将图像文件上传到FTP存储设备
                If mblnIsFirstImgUpLoadFiald Then
                    blnLoadResult = mobjFtp.FuncUploadFile(mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                Else
                    blnLoadResult = FuncUploadFile(mobjFtp, mobjFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                End If
                
                If Not blnLoadResult Then
                    strUpFailureImgs = strUpFailureImgs & "、" & strImages(i)
                End If
                
                If mobjBakFtpInf.strDeviceId <> "" And blnLoadResult Then
                    '连接备份设备
                    If mobjBakFtp.hConnection = 0 Then Call mobjBakFtp.FuncFtpConnect(mobjBakFtpInf.strFTPIP, mobjBakFtpInf.strFTPUser, mobjBakFtpInf.strFTPPwd)
                    '将图像文件上传到FTP备份设备
                    Call mobjBakFtp.FuncUploadFile(mobjBakFtpInf.strFtpDir & mstrSubDir, mstrDestMainDir & mstrSubDir & "\" & strImages(i), strImages(i))
                End If
            Else
                strUpFailureImgs = strUpFailureImgs & "、" & strImages(i)
            End If
        End If
    Next
    
    '断开FTP连接
    mobjFtp.FuncFtpDisConnect
    mobjBakFtp.FuncFtpDisConnect
    
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
End Sub

Private Sub mobjInterfaceEvent_OnReportAction(ByVal lngReportEvtAction As ZLPacsVBCommon.TReportEvtActionType, ByVal lngOrderID As Long, ByVal strDocId As String, ByVal lngReportState As Long)
    On Error GoTo ErrorHand
    
    Select Case lngReportEvtAction
        Case TReportEvtActionType.ratAfterClosed
            RaiseEvent AfterClosed(lngOrderID, strDocId)
            
        Case TReportEvtActionType.ratAfterDeleted
            RaiseEvent AfterDeleted(lngOrderID, strDocId)
        
        Case TReportEvtActionType.ratAfterOpen
            RaiseEvent AfterOpen(lngOrderID, strDocId)
        
        Case TReportEvtActionType.ratAfterPrinted
            RaiseEvent AfterPrinted(lngOrderID, strDocId)
        
        Case TReportEvtActionType.ratAfterReportStateChanged
            '终审后直接打印
            BUGEX "mobjInterfaceEvent_OnReportAction mintCompletePrint = " & mintCompletePrint & "  reportState = " & lngReportState
            If mintCompletePrint = 1 And lngReportState = 4 Then
                Call OpenFormForReportPrint(strDocId, False)
            End If
            
        Case TReportEvtActionType.ratAfterSaved
            'reportState0-普通保存，1-诊断签名，2-审核（终审）签名
            RaiseEvent AfterSaved(lngOrderID, strDocId, lngReportState)
        
        Case TReportEvtActionType.ratBeforeEdit
            RaiseEvent BeforeEdit(lngOrderID, strDocId)
            
    End Select
    
    Exit Sub
ErrorHand:
    
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
        If mobjFtp.FuncFtpConnect(mobjFtpInf.strFTPIP, mobjFtpInf.strFTPUser, mobjFtpInf.strFTPPwd) = 0 Then
            '连接FTP备份设备
            If mobjBakFtp.hConnection = 0 Or blnRefresh Then
                If mobjBakFtp.FuncFtpConnect(mobjFtpInf.strFTPIP, mobjFtpInf.strFTPUser, mobjFtpInf.strFTPPwd) = 0 Then
                    MsgBox "FTP连接失败，请检查网络设置！", vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    ConnFtp = True
End Function

Private Function InitFtpInfo(ByVal lngOrderID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    If mlngTempOrderId = lngOrderID Then
        InitFtpInfo = True
        Exit Function
    End If
    
    mlngTempOrderId = lngOrderID
    
    strSQL = "Select Decode(C.接收日期,Null,'',to_Char(C.接收日期,'YYYYMMDD')||'/')||C.检查UId||'/' As URL," & _
             "D.设备号 as 设备号1,D.FTP用户名 as FTP用户名1,D.FTP密码 as FTP密码1,D.IP地址 as IP地址1,'/'||D.Ftp目录||'/' As Root1, " & _
             "E.设备号 as 设备号2,E.FTP用户名 as FTP用户名2,E.FTP密码 as FTP密码2,E.IP地址 as IP地址2,'/'||E.Ftp目录||'/' As Root2 " & _
             "From 影像检查记录 C,影像设备目录 D,影像设备目录 E Where  C.位置一=D.设备号(+) And C.位置二=E.设备号(+) and C.医嘱Id=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取FTP信息", lngOrderID)
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInf.strFtpDir = Nvl(rsTmp("Root1"))
    mobjFtpInf.strFTPIP = Nvl(rsTmp("IP地址1"))
    mobjFtpInf.strFTPPwd = Nvl(rsTmp("FTP密码1"))
    mobjFtpInf.strFTPUser = Nvl(rsTmp("FTP用户名1"))
    mobjFtpInf.strDeviceId = Trim(Nvl(rsTmp("设备号1")))
    
    mobjBakFtpInf.strFtpDir = Nvl(rsTmp("Root2"))
    mobjBakFtpInf.strFTPIP = Nvl(rsTmp("IP地址2"))
    mobjBakFtpInf.strFTPPwd = Nvl(rsTmp("FTP密码2"))
    mobjBakFtpInf.strFTPUser = Nvl(rsTmp("FTP用户名2"))
    mobjBakFtpInf.strDeviceId = Trim(Nvl(rsTmp("设备号2")))
    
    mstrSubDir = Nvl(rsTmp("URL"))
    mstrDestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
    
    InitFtpInfo = True
End Function

Public Function GetCurrDocId(ByVal lngOrderID As Long) As String
'获取当前选中的报告的Id
On Error Resume Next
    If mobjPacsInterface Is Nothing Then Exit Function
    GetCurrDocId = mobjPacsInterface.GetCurrDocId(lngOrderID)
End Function

Public Sub EvaluatReportQuality(ByVal strDocId As String, ByVal strResult As String)
 '评定报告质量
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.EvaluatReportQuality(strDocId, strResult)
End Sub

Public Sub EvaluatResult(ByVal strDocId As String, ByVal strResult As String)
'评定阴阳性
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.EvaluatResult(strDocId, strResult)
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
    GetReportReleaseState = mobjPacsInterface.GetReportReleaseState(lngOrderID, mlngPatientID)
End Function

Public Function GetCurrReportState(ByVal lngOrderID As Long) As Integer
'返回当前报告的报告状态
    If mobjPacsInterface Is Nothing Then Exit Function
    GetCurrReportState = mobjPacsInterface.GetCurrReportState(Me.GetCurrDocId(lngOrderID))
End Function

'刷新报告图插件中的缩略图和缓存图
Public Sub RefreshMiniImage(ByVal lngOrderID As Long)
    If mobjPacsInterface Is Nothing Then Exit Sub
    '刷新前台和后台采集的图像
    '报告Id传空，医嘱Id为0表示刷新后台采集的图像
    Call mobjPacsInterface.ExecutePlugInCmd(lngOrderID, "", "图像刷新", lngOrderID)
End Sub

Private Sub SetReportFinalShowMode(ByVal blnFinalShowMode As Boolean)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetReportFinalShowMode(blnFinalShowMode)
End Sub

'获取当前报告的创建人
Public Function GetDocCreator(ByVal lngOrderID As Long)
    Dim strDocId As String
    
    If mobjPacsInterface Is Nothing Then Exit Function
    
    strDocId = Me.GetCurrDocId(lngOrderID)
    
    If Trim(strDocId) = "" Then Exit Function
    
    GetDocCreator = mobjPacsInterface.GetDocCreator(strDocId)
End Function

'报告执行
Public Sub SetDocCreator(ByVal strCreatorName)
    If mobjPacsInterface Is Nothing Then Exit Sub
    Call mobjPacsInterface.SetRecordDoctor(strCreatorName)
End Sub
