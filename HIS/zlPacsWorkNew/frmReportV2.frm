VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "*\A..\ZL9PACSCONTROL\zl9PacsControl.vbp"
Begin VB.Form frmReportV2 
   Caption         =   "报告编辑"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13320
   Icon            =   "frmReportV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   13320
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer timerShake 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   1080
      Top             =   0
   End
   Begin VB.PictureBox picBack 
      Height          =   9135
      Left            =   120
      ScaleHeight     =   9075
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   840
      Width           =   12855
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   9075
         Left            =   4095
         TabIndex        =   1
         Top             =   0
         Width           =   110
         _ExtentX        =   185
         _ExtentY        =   16007
         SplitWidth      =   110
         SplitLevel      =   3
         Control1Name    =   "ucPacsHelper1"
         Control2Name    =   "ucReportEditor1"
      End
      Begin zl9PACSWork.ucReportEditor ucReportEditor1 
         Height          =   9075
         Left            =   4205
         TabIndex        =   3
         Top             =   0
         Width           =   8590
         _ExtentX        =   14208
         _ExtentY        =   15266
      End
      Begin zl9PACSWork.ucPacsHelper ucPacsHelper1 
         Height          =   9075
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   16007
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   120
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmReportV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements IWorkMenuV2

Private Const C_MODULE_NAME As String = "frmReportV2"

Private Const M_STR_HINT_NoSelectData As String = "无效的检查数据，请重新选择。"
Private Const M_STR_MODULE_MENU_TAG As String = "报告"
Private Const M_STR_LISTVIEWKEY_DESCRIBE As String = "describe"
Private Const M_STR_LISTVIEWKET_PROCESS As String = "process"

Private Const G_STR_TAG = "Po=息肉[+]E=糜烂区[+]M=镶嵌[+]L=粘膜白斑[+]C=湿疣[+]I=浸润性癌[+]W=醋酸白色上皮[+]AT=异常转化区[+]V=非典型血管[+]P=点状血管[+]Xn=直接活检部位"
Private Const conMenu_Process_MartItem As Long = 1000

Private Type rptFormat
    ID As Long          '报告格式ID
    strName As String   '报告格式名称
End Type
 

Private rptFormats() As rptFormat
Private mstr选中报表格式 As String
Private mstr报表编号 As String
Private mblOneReportFormat As Boolean           '是否只能选择一种打印格式
Private mstrRegPath As String
Private mStrTextMarks As String

Private mObjActiveMenuBar As CommandBars

Private mlngModuleId As Long
Private mstrPrivs As String
Private mintContextFontSize As Long

Private mblnMenuDownState As Boolean
Private mblnIsLinkHelper As Boolean
Private mobjCurPacsHelper As Object
Private mobjCapLinker As clsCapLinker

Private mlngDeptID As Long
 
Private mobjStudyInfo As clsStudyInfo
Private mlngFileFormatId As Long
Private mblnCheckPrintPara As Boolean
Private mblnExitAfterSign As Boolean    '报告签名后退出
Private mblnSetFocusWithReport As Boolean

Private mblnHasFace As Boolean
Private mObjNotify As IEventNotify
Private mblnHasExitSave As Boolean  '是否进行了退出时保存
 
Public mobjRichReportWrap As frmEPREditWrapV2
 


'医嘱ID
Property Get AdviceId() As Long
    If mobjStudyInfo Is Nothing Then
        AdviceId = 0
    Else
        AdviceId = mobjStudyInfo.lngAdviceId
    End If
End Property

Property Get StudyInfo() As clsStudyInfo
    Set StudyInfo = mobjStudyInfo
End Property

Property Set StudyInfo(value As clsStudyInfo)
    Set mobjStudyInfo = value
    
'    mblnIsRefreshStudy = False
End Property

Property Get IsLinkHelper() As Boolean
    IsLinkHelper = mblnIsLinkHelper
End Property


'报告编辑器对象
Property Get ReportEditor() As Object
    Set ReportEditor = ucReportEditor1
End Property

Property Get ReportHelper() As Object
    Set ReportHelper = ucPacsHelper1
End Property


'获取菜单接口对象
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property


Public Sub SetPrintFmts(ByVal strReportNo As String, ByVal strPrintFmts As String)
    If Len(strReportNo) > 0 Then mstr报表编号 = strReportNo
    If Len(strPrintFmts) > 0 Then mstr选中报表格式 = strPrintFmts
End Sub



Private Function HintError(objErr As ErrObject, ByVal strMethodName As String, _
    Optional ByVal blnIsDataErr As Boolean = True) As Long
    If blnIsDataErr Then
        HintError = mObjNotify.PrintErr(objErr, infDataErr, GetReportParentHwnd, C_MODULE_NAME, strMethodName)
    Else
        HintError = mObjNotify.PrintErr(objErr, infNormalErr, GetReportParentHwnd, C_MODULE_NAME, strMethodName)
    End If
End Function

Private Function HintMsg(ByVal strMsg As String, ByVal strMethodName As String, _
    Optional ByVal lngMsgType As Long = infHint) As Long
        HintMsg = mObjNotify.PrintInfo(strMsg, lngMsgType, GetReportParentHwnd, C_MODULE_NAME, strMethodName)
End Function


Private Function AllowPrint() As Boolean
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    AllowPrint = False
    
    strSQL = "Select a.报告人,a.复核人,b.紧急标志 ,b.Id From 影像检查记录 a ,病人医嘱记录 b Where a.医嘱id = b.Id And b.Id = [1] "
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "验证是否可以打印", mobjStudyInfo.lngAdviceId)

    If rsTemp.EOF = False Then
        AllowPrint = IIf(nvl(rsTemp!紧急标志, 0) = 1, nvl(rsTemp!报告人) <> "", (mblnCheckPrintPara And nvl(rsTemp!复核人) <> "") Or mblnCheckPrintPara = False)
    End If

    Exit Function
errH:
    If HintError(err, "AllowPrint") = 1 Then Resume
End Function

Public Sub AddRepImgFile(ByVal strFile As String, Optional ByVal lngImageAdviceId As Long = 0, Optional ByVal strFileName As String = "")
'添加报告图文件
    Call ucReportEditor1.AddRepImgFile(strFile, lngImageAdviceId, strFileName)
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errhandle
    Dim blnIsNewReport As Boolean
    Dim strHistoryData As String
    Dim blnIsCancel As Boolean
    Dim strReportImages As String
    Dim lngResult As Long
    
    If mblnMenuDownState Then Exit Sub
 
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 0, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
    
    mblnMenuDownState = True
    
    blnIsCancel = False
    
    Select Case Control.ID
        Case conMenu_PacsReport_Specialty
            If ucReportEditor1.HasSpeReport Then
                ucReportEditor1.IsSpeState = Not ucReportEditor1.IsSpeState
            End If
            
        Case conMenu_Edit_Modify, conMenu_File_Open, conMenu_File_ExportToXML, conMenu_Tool_Search   '使用病历编辑器方式进行打开
            Call EprEdit(Control)
             
            'TODO=病历编辑中的打印状态同步处理
            'TODO=审核签名后，需要写入复核人 ：ZL_影像报告保存_Update  Zl_影像检查_State
        Case conMenu_Process_AddMark, conMenu_Process_MartItem To conMenu_Process_MartItem + 99
            If Control.ID = conMenu_Process_AddMark Or Control.ID = conMenu_Process_MartItem Then
                Call ucReportEditor1.Mark(imtAuto)
            Else
                If Val(Control.Caption) <= 6 And Val(Control.Caption) > 0 Then
                    Call ucReportEditor1.Mark(imtSpecify, Control.Caption)
                Else
                    Call ucReportEditor1.Mark(imtSpecify, Split(Control.Caption & "=", "=")(0))
                End If
            End If
        
        Case conMenu_Process_MartItem + 100 '清除标注
            Call ucReportEditor1.ClearMark(True)
            
        Case conMenu_Process_MartItem + 101 '设置标注
            Call SetMarkTextLabel
            
        Case conMenu_Edit_Delete                '删除
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_DELETE, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            strHistoryData = ucReportEditor1.ReportID
            If DelReport Then
'                Call SendRequest(WM_LIST_SYNCROW, ,mobjStudyInfo.lngAdviceId)
                If ucPacsHelper1.Visible Then
                    Call ucPacsHelper1.ClearReportImgState
                End If
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_DELETE, 1, ucReportEditor1.AdviceId, strHistoryData)
            End If
            
        Case conMenu_PacsReport_Save            '保存
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            blnIsNewReport = IIf(ucReportEditor1.ReportID = 0, True, False)
            If ucReportEditor1.SaveReport(strReportImages) Then
                mblnHasExitSave = True
                
                '首次保存时需要刷新列表行
'                If blnIsNewReport Then Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                If ucPacsHelper1.Visible Then
                    Call ucPacsHelper1.SyncReportImgState(strReportImages)
                End If
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, strReportImages)
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
        Case conMenu_File_Print                '打印
            '判断是否选中了多份打印格式
            If Len(mstr选中报表格式) > 0 Then
                If UBound(Split(mstr选中报表格式, ",")) > 0 Then
                    If HintMsg("当前报告包含多种打印格式 [" & mstr选中报表格式 & "]，是否继续？", "cbrMain_Execute", vbYesNo) = vbNo Then
                        mblnMenuDownState = False
                        Exit Sub
                    End If
                End If
            End If
            
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            
            If blnIsCancel Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            If ucReportEditor1.ReportPrint(mstr报表编号, mstr选中报表格式) Then
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
            '打印后可能会自动完成等，因此需要对状态进行刷新
            Call ucReportEditor1.ConfigFaceState
            
        Case conMenu_File_Preview               '预览
            Call ucReportEditor1.ReportPreview(mstr报表编号, mstr选中报表格式)
            
        Case conMenu_PacsReport_AddNumber       '序号
            Call ucReportEditor1.AddNumber
            
        Case conMenu_PacsReport_History         '历史
            Call ucReportEditor1.RevisionHistory
            
        Case conMenu_PacsReport_Sign      '签名
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            lngResult = ucReportEditor1.Sign    '0-未签名，1-诊断签名，2-审核签名
            
            '判断是否需要签名后退出
            If lngResult = 0 Then
                mblnMenuDownState = False
                Exit Sub  '签名失败则退出
            End If
            
            '需要传递报表选择格式，因为签名后有可能会自动进行打印操作
            Select Case lngResult
                Case 1
    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, , mstr报表编号 & ":" & mstr选中报表格式)
                    
                    mobjStudyInfo.blnCanPrint = AllowPrint
                Case 2
    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_AUDIT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, , mstr报表编号 & ":" & mstr选中报表格式)
                    
                    mobjStudyInfo.blnCanPrint = AllowPrint
            End Select
            
            If mblnExitAfterSign And mblnIsLinkHelper = False Then
                mblnMenuDownState = False
                Unload Me
                Exit Sub
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
            If ucReportEditor1.IsReadOnly Then Call ResetEmbedVideoState(True)
                
        Case conMenu_PacsReport_DelSign         '回退
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_BACK, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.SignUntread Then
                mblnHasExitSave = True
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_BACK, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
            If ucReportEditor1.IsEditable Then Call ResetEmbedVideoState(False)
            
        Case conMenu_PacsReport_VerifySign_Item '验证
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_Verify, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.SignVerifiy(Val(Control.Parameter)) Then
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_Verify, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
        Case conMenu_PacsReport_Reject     '驳回
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.ReportReject Then
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, , ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
        Case conMenu_View_Refresh
            Call zlRefresh(mobjStudyInfo, True, False)
            
        Case conMenu_PacsReport_RejectHistory   '驳回历史
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJHISTORY, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            Call ucReportEditor1.RejectHistory
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, , ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
        
        Case conMenu_PacsReport_SelFormat_Item  '格式选择
            Call ucReportEditor1.ChangeReportFormat(Val(Control.Parameter))
            
        Case conMenu_PacsReport_RepFormat_Item  '打印格式
            Call ChangeRptFormat(Control.Index)
        
            
        Case conMenu_PacsReport_ClearWritingState   '清除报告“处理中”标记
            If mobjStudyInfo.blnMoved Then
                mblnMenuDownState = False
                Call HintMsg("报告数据已进行转储，不能执行当前操作。", "cbrMain_Execute", vbOKOnly)
                
                Exit Sub
            Else
                Call ClearReportState(mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_PrivOrder   '上条
            If ucPacsHelper1.Visible Then
                Call SendRequest(WM_LIST_GETLASTADVICE, mobjStudyInfo.lngAdviceId)
            Else
                Call SendRequest(WM_LIST_MOVEUP, mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_NextOrder   '下条
            If ucPacsHelper1.Visible Then
                Call SendRequest(WM_LIST_GETNEXTADVICE, mobjStudyInfo.lngAdviceId)
            Else
                Call SendRequest(WM_LIST_MOVEDOWN, mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_FontSet, conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser            '设置文本段字体
            Dim cbrEdit As CommandBarEdit
                
            Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
        
            If Control.ID = conMenu_PacsReport_FontSetUser Then
            '如果是自定义字号，判断是否符合规则
                
                If Not CheckUserFontValidate(cbrEdit.Text) Then
                '不符合规则，相当于设置失败
                    cbrEdit.Text = ""
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
                ucReportEditor1.EditFontSize = Abs(Val(cbrEdit.Text))
                Call zlDatabase.SetPara("报告显示字号", (Abs(Val(cbrEdit.Text))), glngSys, glngModul)
            Else
            '不是自定义字号，前面打勾，自定义text为空表示未选择自定义字号
                cbrEdit.Text = ""
                Control.Checked = True
                
                If Val(Control.Caption) = 0 Then
                    ucReportEditor1.EditFontSize = FontSize
                    Call zlDatabase.SetPara("报告显示字号", FontSize, glngSys, glngModul)
                Else
                    ucReportEditor1.EditFontSize = Val(Control.Caption)
                    Call zlDatabase.SetPara("报告显示字号", Val(Control.Caption), glngSys, glngModul)
                End If
            End If
            
    End Select
    
    Call mObjNotify.Broadcast(BM_SYS__EVENT_MENU, 1, mobjStudyInfo.lngAdviceId, Control.ID, Control.Category)
    
    mblnMenuDownState = False
Exit Sub
errhandle:
    mblnMenuDownState = False
    HintError err, "cbrMain_Execute", False
End Sub


Private Sub ResetEmbedVideoState(ByVal blnReadState As Boolean)
    If Not mobjCapLinker Is Nothing Then mobjCapLinker.ReportAdviceId = mobjStudyInfo.lngAdviceId
    
    If Not ucPacsHelper1.EmbedVideo Is Nothing Then
        If mobjCapLinker.LockAdviceId <> 0 And mobjCapLinker.LockAdviceId <> mobjStudyInfo.lngAdviceId Then Exit Sub    '锁定采集和当前报告不是相同检查时退出
        
        Call ucPacsHelper1.EmbedVideo.zlRestoreWindow(blnReadState And mobjCapLinker.ReportAdvReadOnly, , True)
    End If
End Sub

Public Sub ReadRepStateTag()
    If ucReportEditor1 Is Nothing Then Exit Sub
    
    Call ucReportEditor1.ReadResultTag(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
End Sub

Private Sub SetMarkTextLabel()
'------------------------------------------------
'功能：设置文字标注，并保存
'参数：
'返回：无
'------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    Dim objMenu As CommandBarControl
    On Error GoTo err
    
'    strTemp = InputBox("请输入新的文字标注配置，格式为“简码1=说明1|简码2=说明2|...”。", "文字标注设置", Replace(mstrTemp, "[+]", "|"))
    
    strTemp = frmInputBoxV2.ZlShowMe(mObjNotify.Owner, mStrTextMarks)
    
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        HintMsg "输入的格式不正确，应该按照“简码=说明”方式输入，请检查后重新设置。", "SetMarkTextLabel", vbOKOnly
        Exit Sub
    End If

    mStrTextMarks = strTemp
    Call SaveSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", Replace(strTemp, "|", "[+]"))

    Set objMenu = cbrMain.FindControl(, conMenu_Process_AddMark)
    
    If objMenu Is Nothing Then Exit Sub
    
    objMenu.CommandBar.Controls.DeleteAll
    Call LoadMark(objMenu)
    
    Exit Sub
err:
    If HintError(err, "SetMarkTextLabel", False) = 1 Then Resume
End Sub


Private Sub EprEdit(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If ucReportEditor1.IsModify Then Call ucReportEditor1.SaveReport
    
    If mobjRichReportWrap Is Nothing Then Set mobjRichReportWrap = New frmEPREditWrapV2
    
    If mobjRichReportWrap.InitEprEditor(Me, mObjNotify, mlngModuleId, mlngDeptID) = False Then Exit Sub
    
    If Control.ID = conMenu_Edit_Modify Then
        Call mobjRichReportWrap.OpenEprEditor(mobjStudyInfo, IIf(ucReportEditor1.SourceVer > 1, True, False))
        
        Call mObjNotify.Broadcast(BM_REPORT_EVENT_OPEN, "1", mobjStudyInfo.lngAdviceId, mobjStudyInfo.lngSendNo)
    Else
        Call mobjRichReportWrap.ExecuteMenu(mobjStudyInfo, Control.ID)
    End If
End Sub


Private Sub ChangeRptFormat(ByVal lngIndex As Long)
'更改被选中的自定义报表打印格式
    Dim cbrRptFormat As CommandBarControl
    Dim cbrRptFormatItem As CommandBarControl
    
    Dim i As Integer
    
    On Error GoTo err
    
    Set cbrRptFormat = cbrMain.FindControl(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, True)
    
    mstr选中报表格式 = ""
    
    If mblOneReportFormat Then
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If i = lngIndex Then
                cbrRptFormatItem.Checked = True
                mstr选中报表格式 = cbrRptFormatItem.Caption
            Else
                cbrRptFormatItem.Checked = False
            End If
        Next i
    Else
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If cbrRptFormatItem.Index = lngIndex Then cbrRptFormatItem.Checked = Not cbrRptFormatItem.Checked
            If cbrRptFormatItem.Checked = True Then
                mstr选中报表格式 = IIf(mstr选中报表格式 = "", cbrRptFormatItem.Caption, mstr选中报表格式 & "," & cbrRptFormatItem.Caption)
            End If
        Next i
    End If
    
    SaveSetting "ZLSOFT", mstrRegPath, "报表编号", mstr报表编号
    SaveSetting "ZLSOFT", mstrRegPath & "\" & mstr报表编号, "选中报表格式", mstr选中报表格式
    
    ucReportEditor1.ShowPrintFormat (Replace(mstr选中报表格式, ",", "  "))
    
    Exit Sub
err:
    If HintError(err, "ChangeRptFormat", False) = 1 Then
        Resume
    End If
End Sub


Private Function CheckUserFontValidate(ByVal strValue As String) As Boolean
'规则：经过abs(val(?))处理后是数字，否则验证不通过并且提示

    CheckUserFontValidate = True
    
    If Not IsNumeric(strValue) Or Val(strValue) < 1 Or Val(strValue) > 80 Then
        Call HintMsg("自定义字号只能设置为1到80中一个数字，请重新设置", "CheckUserFontValidate", vbOKOnly)
        CheckUserFontValidate = False
        Exit Function
    End If
    
End Function

Private Function GetReportParentHwnd() As Long
    If ucPacsHelper1.Visible Then
        GetReportParentHwnd = Me.hwnd
    Else
        GetReportParentHwnd = mObjNotify.Owner.hwnd
    End If
End Function

Private Function GetExistsWindow(ByVal lngAdviceId As Long) As Object
'判断是否已经打开弹出式报告窗口
    Dim objForm As Object
    
    Set GetExistsWindow = Nothing
    
    '判断是否存在已经打开的报告编辑框
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.AdviceId = lngAdviceId And objForm.IsLinkHelper = False Then
                Set GetExistsWindow = objForm
                
                Exit Function
            End If
        End If
    Next
End Function

Public Sub SendRequest(ByVal lngEventNo As Long, ByVal lngMainAdviceId As Long)
    Dim lngNewAdviceId As Long
    Dim lngSendNo As Long
    Dim blnIsMoved As Boolean
    Dim objStudyInfo As clsStudyInfo
    Dim objForm As Object
 
     
    If mObjNotify Is Nothing Then Exit Sub
    
    Select Case lngEventNo
        Case WM_LIST_MOVEUP, WM_LIST_MOVEDOWN
            lngNewAdviceId = lngMainAdviceId
            mObjNotify.SendRequest lngEventNo   'sendRequest后,mobjStudyInfo的医嘱ID会同步更新为最新检查的医嘱ID
            
            If lngNewAdviceId = mobjStudyInfo.lngAdviceId Then
                '弹出是报告窗口的msgboxd的parent必须为me
                HintMsg "已移动到" & IIf(lngEventNo = WM_LIST_MOVEUP, "起始检查行。", "末尾检查行。"), "SendRequest", vbOKOnly
                Exit Sub
            End If
            
        Case WM_LIST_GETLASTADVICE, WM_LIST_GETNEXTADVICE
            lngNewAdviceId = lngMainAdviceId
            mObjNotify.SendRequest lngEventNo, , lngNewAdviceId, lngSendNo, blnIsMoved  'lngNewAdviceId返回移动后最新的医嘱ID
            
            If lngNewAdviceId = lngMainAdviceId Then
                '弹出是报告窗口的msgboxd的parent必须为me
                HintMsg "已移动到" & IIf(lngEventNo = WM_LIST_GETLASTADVICE, "起始检查行。", "末尾检查行。"), "SendRequest", vbOKOnly
                Exit Sub
            End If
            
            Set objStudyInfo = zlGetStudyAdvice(lngNewAdviceId)
            
            If objStudyInfo Is Nothing Then
                HintMsg "未获取到对应的检查信息。", "SendRequest", infNormalErr
                Exit Sub
            End If
            
            If objStudyInfo.intStep <= 1 Then
                '跳过未报到或已拒绝的检查
                Call SendRequest(lngEventNo, objStudyInfo.lngAdviceId)
                Exit Sub
            End If
            
            '判断对应检查报告是否已经被其他窗口打开，如果已经打开，则进行提示,且跳过已经打开的报告
            Set objForm = GetExistsWindow(lngNewAdviceId)
            If Not objForm Is Nothing Then
                If HintMsg("对应报告窗口已经打开，是否切换？", "IsAllowExistsWindowChange", vbYesNo) = vbNo Then
                    Call SendRequest(lngEventNo, objStudyInfo.lngAdviceId)
                Else
                    objForm.WindowState = 0
                    objForm.Visible = True
                    objForm.ZOrder
                    
                    '抖动处理
                    Call objForm.Shake
                End If
                
                Exit Sub
            End If
            
            Call ucReportEditor1.PromptSave(lngNewAdviceId, 0)
      
            If mblnIsLinkHelper = False Then
                '移出之前的pacshelper对象关联
                If Not mobjCapLinker Is Nothing Then mobjCapLinker.RemoveRepPacsHelper mobjStudyInfo.lngAdviceId
            End If
            
            Call zlRefresh(objStudyInfo)
            Call SetReportTitle(objStudyInfo)
            
        Case WM_LIST_SYNCROW    '同步显示行
            Call mObjNotify.SendRequest(lngEventNo, , lngMainAdviceId)
            
    End Select
    
    Set objStudyInfo = Nothing
End Sub

Private Sub ClearReportState(ByVal lngAdviceId As Long)
'清除报告状态标记
    Dim strSQL As String
    Dim strInfo As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select 报告操作 from 影像检查记录 where 医嘱id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取报告操作人", lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If Trim(nvl(rsTemp!报告操作)) = "" Then Exit Sub
    
    
    strInfo = "本报告的目前状态为 [" & nvl(rsTemp!报告操作) & "] 处理中，确定要清除这份报告的状态吗？"
    If HintMsg(strInfo, "ClearReportState", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    
    Call UpdateReporter(lngAdviceId, "")
    Call ucReportEditor1.ConfigFaceState
End Sub


Private Sub InitReportPrintFormat(ByVal lngFileId As Long)
'初始化报告打印格式
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegReportNo As String
    
    mstrRegPath = "公共模块\" & App.ProductName & "\frmReport"
    
    mstr报表编号 = ""
    mstr选中报表格式 = ""
    
    '先判断是否使用自定义报表
    strSQL = "Select 通用,编号 From 病历文件列表  Where Id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取报告打印方式", lngFileId)
    If rsTemp.EOF = False Then
        If nvl(rsTemp!通用) = 2 Then
            '使用自定义报表格式打印
            mstr报表编号 = "ZLCISBILL" & Format(nvl(rsTemp!编号), "00000") & "-2"
            strRegReportNo = GetSetting("ZLSOFT", mstrRegPath, "报表编号", "")
            
            If mstr报表编号 <> strRegReportNo Then
                '当注册表中的报表编号与当前的报表编号不相同时，更新注册表中报表编号
                '以避免进行批量打印时，还是用前一次的报表进行打印
                SaveSetting "ZLSOFT", mstrRegPath, "报表编号", mstr报表编号
            End If
            
            mstr选中报表格式 = GetSetting("ZLSOFT", mstrRegPath & "\" & mstr报表编号, "选中报表格式", "")
        End If
    End If
    
    Call ucReportEditor1.ShowPrintFormat(Replace(mstr选中报表格式, ",", "  "))
End Sub


Private Sub InitCommandBar()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrPopControl As CommandBarControl
    Dim cbrEdit As CommandBarEdit
        
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = zlCommFun.GetPubIcons
    
    With Me.cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize True, 24, 24
    End With
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    If Me.cbrMain.Count > 1 Then Call Me.cbrMain(2).Delete
    '报告工具栏定义
    Set cbrToolBar = Me.cbrMain.Add("报告栏", IIf(mblnIsLinkHelper, xtpBarRight, xtpBarTop))
    
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Specialty, "专科"): cbrControl.iconid = 2558: cbrControl.ToolTipText = "专科报告编辑"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Save, "保存"): cbrControl.iconid = 3503: cbrControl.ToolTipText = "保存报告": cbrControl.BeginGroup = True

'        Set cbrControl = .Add(xtpControlSplitButtonPopup, 8352, "打印")
'        cbrControl.IconId = 103
'        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
                cbrControl.BeginGroup = True
                cbrControl.iconid = 103
                cbrControl.ToolTipText = "报告打印"
                
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
                cbrControl.iconid = 102
                cbrControl.ToolTipText = "报告预览"
'        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Sign, "签名")
            cbrControl.iconid = 3003
            cbrControl.ToolTipText = "签名"
            cbrControl.BeginGroup = True
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelSign, "回退")
            cbrControl.iconid = 3004
            cbrControl.ToolTipText = "回退签名"
        
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_VerifySign, "验证")
            cbrControl.iconid = 8044
            cbrControl.ToolTipText = "验证"
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            cbrControl.ToolTipText = "刷新"
            cbrControl.BeginGroup = True
            

        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, "格式")
            cbrControl.iconid = 3031
            cbrControl.ToolTipText = "选择自定义报表打印格式"
            
            
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_SelFormat, "模板")
            cbrControl.iconid = 227
            cbrControl.ToolTipText = "选择和更换报告单书写模板"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "病历")
            cbrControl.iconid = 3002
            cbrControl.ToolTipText = "用电子病历方式编辑报告"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_History, "修订史")
            cbrControl.iconid = 5015
            cbrControl.ToolTipText = "查看当前和历史报告的的修订情况"
            
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_AddNumber, "序号")
            cbrControl.BeginGroup = True
            cbrControl.iconid = 9023
            cbrControl.ToolTipText = "给段落文字添加序号"
            
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_AddMark, "标记")
            cbrControl.iconid = 3034
            cbrControl.ToolTipText = "给标记图添加标记"
            
            Call LoadMark(cbrControl)
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Reject, "驳回")
            cbrControl.iconid = 229
            cbrControl.ToolTipText = "报告驳回"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_RejectHistory, "驳回史")
            cbrControl.iconid = 8341
            cbrControl.ToolTipText = "驳回历史"
            
                
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_PrivOrder, "上条")
            cbrControl.BeginGroup = True
            cbrControl.iconid = 21802
            cbrControl.ToolTipText = "上一条检查记录"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_NextOrder, "下条")
            cbrControl.iconid = 21801
            cbrControl.ToolTipText = "下一条检查记录"
    

        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_FontSet, "字号")    'xtpControlSplitButtonPopup
            cbrControl.iconid = 509
            cbrControl.ToolTipText = "字体设置"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSetDefault, "默认", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet14, "14", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet16, "16", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet22, "22", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet28, "28", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet36, "36", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet42, "42", "", 0, False)
                
                
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlEdit, conMenu_PacsReport_FontSetUser, "自定义", "", 0, False)
                
                If mintContextFontSize <> 0 And IsCustomFont(mintContextFontSize) Then
                    Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
                    cbrEdit.Text = mintContextFontSize
                End If
                
            End With

        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.type = xtpControlButton) Or (cbrControl.type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '设置成主界面菜单
    Next
    
    cbrToolBar.Position = IIf(mblnIsLinkHelper, xtpBarRight, xtpBarTop)
End Sub


Private Sub LoadMark(mnuParent As Object)
'载入数字标注
    Dim objControl As CommandBarControl
    Dim arrTemp() As String
    Dim i As Long
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem, "①")
        objControl.ToolTipText = "自动递增数字编号"
        objControl.Category = 0
        objControl.iconid = 0
    
    For i = 1 To 6
        Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + i, i)
            objControl.ToolTipText = "数字编号" & i
            objControl.Category = i
            objControl.iconid = 0
    Next
     
    arrTemp = Split(mStrTextMarks, "|")
    
    For i = 0 To UBound(arrTemp)
        If Len(arrTemp(i)) > 0 Then
            Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 10 + i + 1, arrTemp(i))
                If i = 0 Then objControl.BeginGroup = True
                
                objControl.ToolTipText = "文本标注"
                objControl.Category = i + 1
                objControl.iconid = 0
        End If
    Next
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 100, "清除")
        objControl.BeginGroup = True
        objControl.ToolTipText = "清除标记"
        objControl.Category = 0
        objControl.iconid = 0
        
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 101, "设置")
        objControl.BeginGroup = True
        objControl.ToolTipText = "设置标记"
        objControl.Category = 0
        objControl.iconid = 0
    
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo errH
    
    Control.Enabled = True
    
    If mobjStudyInfo Is Nothing Then
        Control.Enabled = False
        Exit Sub
    End If
    
    Select Case Control.ID
        Case conMenu_PacsReport_Specialty
            Control.Visible = ucReportEditor1.HasSpeReport
            
            If Control.Visible Then
                Control.iconid = IIf(ucReportEditor1.IsSpeState, 3559, 3558)
                Control.Checked = IIf(ucReportEditor1.IsSpeState, True, False)
            End If
        
        Case conMenu_File_Print, conMenu_File_Preview      '打印报告,预览报告

            Control.Visible = CheckPopedom(mstrPrivs, "PACS报告打印")
            If Control.Visible = False Then Exit Sub
            
            '如果未找到对应的病历文件，那么打印预览按钮会被禁用
            If mlngFileFormatId = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = ucReportEditor1.ReportID <> 0
            End If
            
            If Control.Enabled Then Control.Enabled = mobjStudyInfo.blnCanPrint

        Case conMenu_Edit_Modify        '报告编辑（以病历编辑器方式进行编辑）
            Control.Enabled = ucReportEditor1.IsEditable
            

        Case conMenu_PacsReport_Save    '保存
            Control.Enabled = ucReportEditor1.IsModify

        Case conMenu_PacsReport_Reject
            '判断是否具备报告驳回权限
            '判断当前报告所在状态是否允许驳回
            Control.Visible = CheckPopedom(mstrPrivs, "报告驳回")
            If Control.Visible Then Control.Enabled = ucReportEditor1.ReportID <> 0 And Not ucReportEditor1.IsReadOnly

        Case conMenu_PacsReport_RejectHistory
            Control.Visible = True
            If Control.Visible Then Control.Enabled = ucReportEditor1.ReportID <> 0


        Case conMenu_PacsReport_Sign    '签名（能编辑书写报告一般都能进行签名，诊断签名）

            '在书写模式下，还没有签名的，可以签名
            '在修订模式下，签名数量没有超过16次的，可以签名。
            '只读模式下，什么都不能操作。
            Control.Enabled = (Not ucReportEditor1.IsReadOnly And ucReportEditor1.ReportID > 0) Or (ucReportEditor1.IsModify)

        Case conMenu_PacsReport_VerifySign  '签名验证
            '只有启用了数字签名，才显示签名验证按钮
            '只有报告书写，报告修订权限的人，才能对签名进行验证
            Control.Visible = IIf(ucReportEditor1.SignPassType = 0, False, True)
            If Control.Visible Then Control.Enabled = IIf(ucReportEditor1.SourceVer >= 1, True, False)
            
        Case conMenu_PacsReport_DelSign '回退

            '没有签名之前，不可以回退,只能回退自己的签名，或者通过“回退他人签名”的权限，回退本科室其他人的签名
            '只有签名过后才可以回退
            '回退自己的签名
             '有他人报告权限的,可以回退本科室的他人签名
             Control.Enabled = ucReportEditor1.SourceVer >= 1 And (Not ucReportEditor1.IsReadOnly)
             
        Case conMenu_View_Refresh
            Control.Visible = Not mblnIsLinkHelper

        Case conMenu_PacsReport_SelFormat  '选择格式 '修订模式下，不可以设置格式
            Control.Enabled = ucReportEditor1.IsEditable And IIf(ucReportEditor1.SourceVer < 1, True, False)
            
        Case conMenu_PacsReport_SelFormat_Item
            Control.Checked = IIf(Val(Control.Parameter) = ucReportEditor1.SampleId, True, False)
            
        Case conMenu_PacsReport_RepFormat   '选择打印格式
            Control.Visible = IIf(Len(mstr报表编号) > 0, True, False)

        Case conMenu_PacsReport_RepFormat_Item  '选择具体打印格式
            Control.Checked = InStr(mstr选中报表格式, Control.Caption)
            Control.iconid = IIf(Control.Checked, 90002, 90001)
 
        Case conMenu_PacsReport_FontSet, conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser   '设置字号
            Control.Checked = False
            If Val(Control.Caption) = 0 Then
                If FontSize = ucReportEditor1.EditFontSize Then Control.Checked = True
            Else
                If Val(Control.Caption) = ucReportEditor1.EditFontSize Then Control.Checked = True
            End If
 
        Case conMenu_Edit_Delete                            '删除报告
            '有他人报告和报告删除权限时，可以强制删除本科室其他人书写的报告
            '已签名报告不允许被删除
            Control.Visible = (ucReportEditor1.ReportID <> 0 And (CheckPopedom(mstrPrivs, "PACS报告书写") Or CheckPopedom(mstrPrivs, "PACS报告删除")))
            If Control.Visible Then
                '删除自己书写的未签名的报告或相同科室的他人未签名报告
                Control.Enabled = IIf(ucReportEditor1.SourceVer < 1, True, False) _
                                    And (ucReportEditor1.CreateUser = UserInfo.姓名 _
                                        Or (CheckPopedom(mstrPrivs, "PACS他人报告") _
                                            And CheckPopedom(mstrPrivs, "PACS报告删除") _
                                            And ucReportEditor1.CreateDeptId = mlngDeptID _
                                            ) _
                                        ) And (Not ucReportEditor1.IsReadOnly Or Not ucReportEditor1.IsComplete)
            End If

        Case conMenu_PacsReport_ClearWritingState       '清除报告“处理中”的状态,可以清除本科室的报告标记
            Control.Visible = CheckPopedom(mstrPrivs, "PACS报告删除")
            
        Case conMenu_PacsReport_AddNumber
            Control.Enabled = ucReportEditor1.IsEditable

        Case conMenu_Process_AddMark
            Control.Visible = IIf(ucReportEditor1.MarkImageCount > 0, True, False)
            Control.Enabled = ucReportEditor1.IsEditable
            
        Case conMenu_PacsReport_Default
    End Select
    Exit Sub
errH:
    If HintError(err, "cbrMain_Update", False) = 1 Then Resume
End Sub


Private Function DelReport() As Boolean
'删除报告
    DelReport = False
    
    If HintMsg("报告删除后将不能恢复，是否继续？", "DelReport", vbYesNo) = vbNo Then Exit Function
    
    If ucReportEditor1.DelReportData(ucReportEditor1.ReportID, True) = False Then Exit Function
    
    '清除锁定人
    Call ucReportEditor1.UnlockEditor
    
    '清除界面录入数据
    Call ucReportEditor1.ClearReport(True, True, True)
    
    '清除标记
    Call ucReportEditor1.ClearMark(False)
    
    '清除报告图
    Call ucReportEditor1.ClearReportImg
    
    '清除其他信息
    Call ucReportEditor1.ClearInfo
    
    '恢复初始状态
    Call ucReportEditor1.ConfigFaceState
    
    DelReport = True
End Function




Private Sub Form_Activate()
On Error GoTo errhandle
 
'    Debug.Print "TopWindow:" & GetTopWindow(App.hInstance) & " ForeWindow:" & GetForegroundWindow & "  CurWindow:" & Me.hWnd
    If mblnHasFace = False Then Exit Sub
    
    If mblnIsLinkHelper = False Then    '如果没有连接helper，则表示弹出式报告窗口
        '弹出式报告窗口需要判断GetForegroundWindow句柄是否与当前窗口句柄相同，如果不同，说明不是当前置顶的报告编辑窗口
        If GetForegroundWindow <> Me.hwnd Then Exit Sub
    End If
    
    If mblnIsLinkHelper = False Then
        If ucPacsHelper1.AllowEmbedVideo Then
            '如果报告窗口嵌入了视频采集，则切换到对应报告窗口后，在没有锁定采集情况下，可同步采集报告对应检查的图像
            '如果视频窗口嵌入成功，则需要设置caplinker的报告医嘱id为当前医嘱ID
            If Not mobjCapLinker Is Nothing Then
                mobjCapLinker.ReportAdviceId = mobjStudyInfo.lngAdviceId
                
                If ucPacsHelper1.ShowEmbedVideo(mobjCapLinker) = False Then
                    '如果视频采集嵌入失败，则设置caplinker的报告医嘱id为0
                    If Not mobjCapLinker Is Nothing Then mobjCapLinker.ReportAdviceId = 0
                End If
            Else
                ucPacsHelper1.HideEmbedVideo
            End If
        End If
    Else
        '嵌入式报告编辑窗口处理
        If Not mobjCapLinker Is Nothing And VideoIsAttachReportWindow = False Then
            mobjCapLinker.ReportAdviceId = 0
            
            Call mobjCurPacsHelper.ShowEmbedVideo(mobjCapLinker)
        End If
    End If
     
    '定位到内容编辑框，同时tab切换后，编辑框焦点可进行恢复
    If mblnIsLinkHelper = False Or mblnSetFocusWithReport Then
        If Not mObjNotify.Owner.ActiveControl Is Nothing Then
            If TypeOf mObjNotify.Owner.ActiveControl Is PatiIdentify Then
                '如果是查找状态下，不允许自动定位报告编辑
                Exit Sub
            End If
        End If
        
        ucReportEditor1.LocateEditBox
    End If
     
Exit Sub
errhandle:
    HintError err, "Form_Activate", False
End Sub


Private Function VideoIsAttachReportWindow()
'判断视频是否嵌入的弹出式报告窗口
    Dim objForm As Object
    Dim lngVideoRootHwnd As Long
    
    VideoIsAttachReportWindow = False
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.IsLinkHelper = False And Not objForm.ReportHelper Is Nothing Then  'And objForm.AdviceId = mobjCurStudyInfo.lngAdviceId Then
                If Not objForm.ReportHelper.EmbedVideo Is Nothing Then
                    lngVideoRootHwnd = GetAncestor(objForm.ReportHelper.EmbedVideo.VideoHwnd, GA_ROOT)
                    Exit For
                End If
            End If
        End If
    Next
    
    For Each objForm In Forms
        If TypeOf objForm Is frmReportV2 Then
            If objForm.IsLinkHelper = False And objForm.hwnd = lngVideoRootHwnd Then    'And objForm.AdviceId = mobjCurStudyInfo.lngAdviceId Then
                VideoIsAttachReportWindow = True
                Exit Function
            End If
        End If
    Next
End Function


Private Sub SendExitMsg()
On Error Resume Next
   Call mObjNotify.Broadcast(BM_REPORT_EVENT_POPUPEXIT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo errhandle
    '弹出报告内容保存提示
    If mblnIsLinkHelper = False Then
        mblnHasExitSave = mblnHasExitSave Or PromptSave
    
        '判断是否需要同步更新词句片段
        If ucPacsHelper1.IsSyncWordFragment Then Call mObjNotify.Broadcast(BM_REPORT_EVENT_REFFRAGMENT, , hwnd)
    End If
Exit Sub
errhandle:
    HintError err, "Form_QueryUnload", False
End Sub

Private Function JumpNextReportWindow() As Boolean
    Dim objForm As Object
    Dim i As Long
    
    JumpNextReportWindow = False
    
    For i = Forms.Count - 1 To 0 Step -1 ' Each objForm In Forms
        Set objForm = Forms(i)
        
        If TypeOf objForm Is frmReportV2 Then
            If objForm.IsLinkHelper = False And objForm.Visible Then
                
                Call objForm.LocateEditBox
                
                JumpNextReportWindow = True
                
                Exit Function
            End If
        End If
    Next
End Function


Private Sub Form_Terminate()
On Error GoTo errhandle
    If JumpNextReportWindow() = False Then
        If mblnIsLinkHelper = False Then Call SetForegroundWindow(MainForm.hwnd)
    End If
Exit Sub
errhandle:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strLayoutKey As String
    Dim strLayoutStr As String
    
On Error GoTo errhandle
 
    If mblnIsLinkHelper = False Then
        '弹出式窗口需要发送窗口退出消息以便嵌入式窗口刷新
        If mblnHasExitSave Then Call SendExitMsg
    End If

    Call ucReportEditor1.UnlockEditor
    
    strLayoutKey = ucReportEditor1.GetFaceKey
    strLayoutStr = ucReportEditor1.GetLayoutStr()
    
    If mblnIsLinkHelper = False Then
        Call SaveWinState(Me)
        '非嵌入式报告窗口保存ucpacsHelper布局串
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HELPER" & mlngModuleId, ucPacsHelper1.GetLayoutStr)
        
        If Me.ScaleWidth > 0 Then
            Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HelperWidth" & mlngModuleId, ucPacsHelper1.Width)
            Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "EditorWidth" & mlngModuleId, ucReportEditor1.Width)
        End If
        
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "POPUPEDITOR" & mlngModuleId & strLayoutKey, strLayoutStr)
    Else
        '保存嵌入式窗口状态
        Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "MAINEDITOR" & mlngModuleId & strLayoutKey, strLayoutStr)
    End If
    
    '嵌入主主视频窗口...
    If mblnHasFace Then
        If Not mobjCapLinker Is Nothing Then mobjCapLinker.ReportAdviceId = 0
        
'        If ChangeNextReportWindow(mobjStudyInfo.lngAdviceId) = False Then
'            If mblnIsLinkHelper = False Then Call SetForegroundWindow(mObjNotify.hwnd)
'        End If
        
        If Not mobjCapLinker Is Nothing Then Call mobjCapLinker.RemoveRepPacsHelper(mobjStudyInfo.lngAdviceId)
    End If

    Call ucSplitter1.Destory

    ucPacsHelper1.Destory
    ucReportEditor1.Destory
    
    Set mobjRichReportWrap = Nothing
    Set mobjCapLinker = Nothing
    Set mobjCurPacsHelper = Nothing
    Set mObjActiveMenuBar = Nothing
    Set mObjNotify = Nothing
    Set mobjStudyInfo = Nothing
Exit Sub
errhandle:
    Debug.Print "frmReportV2_UnLoad Err:" & err.Description
End Sub

'接口实现部分*********************************************************************************

Public Function IWorkMenuV2_zlBaseMenuID() As Long
End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'执行菜单命令

End Function
 
Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'判断菜单是否属于该模块菜单
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set mObjActiveMenuBar = objMenuBar
     
    Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "报告", 3, False)
    cbrMenuBar.ID = conMenu_EditPopup
    cbrMenuBar.Category = ""
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 3002, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "清除状态", "", 21903, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "删除", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Open, "查阅", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_ExportToXML, "导出XML…", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Search, "报告检索…", "", 0, False)
    End With
End Sub


Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
''创建工具栏
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long

    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue, , True)

    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "预览", "报告预览", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "打印", "报告打印", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "书写", "", 2607, False, lngIndex + 3) 'IconId=3002
End Sub


Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'清除所创建的菜单
    Exit Sub
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'清除创建的工具栏
    Exit Sub
End Sub



Public Sub IWorkMenuV2_zlUpdateMenu(ByVal strModuleName As String, Control As XtremeCommandBars.ICommandBarControl)
    Call cbrMain_Update(Control)
End Sub

Public Sub IWorkMenuV2_zlExecuteMenu(ByVal strModuleName As String, ByVal lngMenuId As Long)
    Dim objControl As XtremeCommandBars.ICommandBarControl
    Dim blnIsCancel As Boolean
    
    If mObjActiveMenuBar Is Nothing Then
        Set objControl = cbrMain.FindControl(, lngMenuId, , True)
    Else
        Set objControl = mObjActiveMenuBar.FindControl(, lngMenuId, , True)
    End If
    
    If objControl Is Nothing Then
'        '通过id执行对应功能
'        Select Case lngMenuId
'            Case conMenu_File_BatPrint  '批量打印操作
'               Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
'                If blnIsCancel Then mblnMenuDownState = False: Exit Sub
'
'                If ucReportEditor1.ReportPrint(mstr报表编号, mstr选中报表格式, True) Then
'    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
'                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
'                End If
'        End Select
        
        Exit Sub
    End If
    
    Call cbrMain_Execute(objControl)
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'配置右键菜单
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'刷新弹出的子菜单
    Exit Sub
End Sub

'*************************************************************************************************


Public Sub LocateEditBox()
'定位编辑框
    ucReportEditor1.LocateEditBox
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'工作站菜单栏改变字号
    Call SetMeneFontSize(bytFontSize)
    
    If mblnIsLinkHelper = False Then
        Call ucPacsHelper1.SetFontSize(bytFontSize)
    End If
    
    Call ucReportEditor1.SetFontSize(bytFontSize)
End Sub

Public Sub SetMeneFontSize(ByVal intFontSize As Integer)
    Me.FontSize = intFontSize
    Set cbrMain.options.Font = Me.Font
End Sub


Public Sub SetMenuDownState(ByVal blnValue As Boolean)
'功能：修改mblnMenuDownState的值，用于处理问题105988
    mblnMenuDownState = blnValue
End Sub


Public Function PrintPreview(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal blnIsPrint As Boolean = False, Optional ByVal lngSpecifyReportId As Long = 0, _
    Optional ByVal strPrintFmts As String = "") As Boolean
'打印和预览
    If blnIsPrint Then
        PrintPreview = ucReportEditor1.ReportPrintEx(lngAdviceId, blnIsMoved, lngSpecifyReportId, mblOneReportFormat, strPrintFmts)
    Else
        Call ucReportEditor1.ReportPreviewEx(lngAdviceId, blnIsMoved, lngSpecifyReportId, mblOneReportFormat)
    End If
    
    
End Function

Public Sub ReinitWordChar()
'同步常用词句
    Call ucReportEditor1.InitReportChar
End Sub

Public Sub ReinitWordFragment()
'刷新词句片段
    Call mobjCurPacsHelper.RefreshData("词句")
End Sub


Public Sub zlInit(objNotify As IEventNotify, ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, objCapLinker As Object, Optional objMainPacsHelper As Object = Nothing, _
    Optional ByVal blnHasFace As Boolean = True)
    Dim strLayout As String
'初始化
    mblnIsLinkHelper = False
    mlngModuleId = lngModuleNo
    mlngDeptID = lngDeptId
    mstrPrivs = strPrivs
    mblnHasFace = blnHasFace
    
    Set mObjNotify = objNotify
 
    '初始化参数
    Call InitParameters(lngDeptId)
    
    If blnHasFace Then
        Set mobjCapLinker = objCapLinker
        
        '判断是否需要继承显示词句，历史，图像等辅助模块
        Set mobjCurPacsHelper = ucPacsHelper1
        
        If Not objMainPacsHelper Is Nothing Then
            '嵌入式窗口
            mblnIsLinkHelper = True
            Set mobjCurPacsHelper = objMainPacsHelper
            
            ucPacsHelper1.Visible = False
            ucSplitter1.Visible = False
        Else
            '弹出式窗口
            Call ucPacsHelper1.Init(mObjNotify, lngModuleNo, lngDeptId, strPrivs)
            
            ucPacsHelper1.HideButtonEnable = False
            
            If Not objCapLinker Is Nothing Then
                Call ucPacsHelper1.ShowEmbedVideo(objCapLinker)
            Else
                Call ucPacsHelper1.HideEmbedVideo
            End If
            
            ucPacsHelper1.AllowLinkerViewer = False
            
'            mobjCapLinker.AddRepPacsHelper mobjStudyInfo.lngAdviceId, ucPacsHelper1
        End If
        
        Call InitCommandBar
        Call ucReportEditor1.Init(mObjNotify, lngModuleNo, lngDeptId, strPrivs, GetSignVerifyType(mlngDeptID), True)
    Else
'        ucSplitter1.Control1Name = ""
'        ucSplitter1.Control1Name = ""
        
        Call ucReportEditor1.Init(mObjNotify, lngModuleNo, lngDeptId, strPrivs, GetSignVerifyType(mlngDeptID), True)
    End If
     
    
    If blnHasFace Then
        ucReportEditor1.EditFontSize = mintContextFontSize
        Set mobjCurPacsHelper.LinkEditor = ucReportEditor1
    End If
    
    
    If mblnIsLinkHelper = False Then
        Call RestoreWinState(Me)
         
         '重试弹出窗口的pacshelper的宽度
        ucPacsHelper1.Width = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HelperWidth" & mlngModuleId, 750))
        ucReportEditor1.Width = Val(GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "EditorWidth" & mlngModuleId, 1000))
        
        strLayout = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "HELPER" & mlngModuleId, "")
        Call ucPacsHelper1.SetLayout(strLayout)
    End If
End Sub





Private Sub InitParameters(ByVal lngDeptId As Long)

    mblnCheckPrintPara = Val(GetDeptPara(mlngDeptID, "平诊需审核才能打报告", 0)) <> 0
    mblnSetFocusWithReport = Val(GetDeptPara(mlngDeptID, "检查切换时定位报告编辑", 1)) = 1
    
    mblnExitAfterSign = IIf(Val(zlDatabase.GetPara("PACS报告签名后退出", glngSys, mlngModuleId, True, "0")) = 0, False, True)
    mintContextFontSize = Val(zlDatabase.GetPara("报告显示字号", glngSys, mlngModuleId))
    mblOneReportFormat = GetDeptPara(lngDeptId, "单选报告格式", True)
    mStrTextMarks = GetSetting("ZLSOFT", "公共模块\zl9PACSWork\frmReportImageEdit", "简明文字标注", G_STR_TAG)
    
    mStrTextMarks = Replace(mStrTextMarks, "[+]", "|")
End Sub



Public Function GetFileFormatId(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean) As Long
'获取检查对应的诊疗单据格式ID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetFileFormatId = 0
    
    strSQL = "Select l.病人来源, a.病历文件id" & vbNewLine & _
            " From 病人医嘱记录 l, 病历单据应用 a" & vbNewLine & _
            " Where l.诊疗项目id = a.诊疗项目id(+) And a.应用场合(+) = Decode(l.病人来源, 2, 2, 4 ,4, 1) And l.Id = [1]"
            
    If blnIsMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询单据格式", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetFileFormatId = Val(nvl(rsData!病历文件id))
    
End Function

Public Function GetReportId(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As Long
'获取检查对应的报告ID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReportId = 0
    strSQL = "Select 病历ID,RawToHex(检查报告ID) 检查报告ID From 病人医嘱报告 Where 医嘱ID= [1]"
    If blnMoved Then
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If nvl(rsData!检查报告ID) <> "" Then
        GetReportId = -1
    Else
        GetReportId = Val(nvl(rsData!病历Id))
    End If
    
End Function


Public Sub SetReportTitle(objStudyInfo As clsStudyInfo)
    Me.Caption = "报告编辑    " & objStudyInfo.strPatientName & " (检查号:" & objStudyInfo.strStudyNum & Decode(objStudyInfo.lngPatientFrom, 1, "  门诊号:" & objStudyInfo.strMarkNum, 2, "  住院号:" & objStudyInfo.strMarkNum, "") & ")    " & objStudyInfo.strPatientAge & "    " & objStudyInfo.strPatientSex & "    " & objStudyInfo.strAdviceContext
End Sub

Public Sub SyncHelper(ByVal lngAdviceId As Long, ByVal lngSourceHwnd As Long, ByVal lngSyncType As Long)
'lngAdviceId:医嘱ID
'lngSourceHwnd:触发该方法的原始控件句柄
'lngSyncType:同步类型0-图像，1-词句,  2-历史   3-缓存
     If lngAdviceId <> mobjStudyInfo.lngAdviceId Then Exit Sub
     If ucPacsHelper1.Visible = False Then Exit Sub
     If lngSourceHwnd = Me.hwnd Or lngSourceHwnd = ucPacsHelper1.hwnd Then Exit Sub
     
     If lngSyncType = 0 And ucPacsHelper1.SelTabName <> "图像" Then Exit Sub
     If lngSyncType = 1 And ucPacsHelper1.SelTabName <> "词句" Then Exit Sub
     If lngSyncType = 2 And ucPacsHelper1.SelTabName <> "历史" Then Exit Sub
     If lngSyncType = 3 And ucPacsHelper1.SelTabName <> "缓存" Then Exit Sub
     
    Call ucPacsHelper1.zlRefresh(mobjStudyInfo, mlngFileFormatId, True)
End Sub

Public Sub zlRefresh(objStudyInfo As clsStudyInfo, _
    Optional ByVal blnIsForceRefresh As Boolean = False, Optional ByVal blnIsHistory As Boolean = False)
    Dim lngReportID As Long
    Dim strLayout As String
    Dim objFocus As Object
    
    Set objFocus = mObjNotify.Owner.ActiveControl
    
    If Not mobjStudyInfo Is Nothing And Not objStudyInfo Is Nothing Then
        If mobjStudyInfo.IsEquals(objStudyInfo) And blnIsForceRefresh = False Then Exit Sub
    End If
    
    Set mobjStudyInfo = objStudyInfo
     
    'mblnIsLinkHelper为false表示弹出窗口，没有关联主窗口的pacshelper对象
    If mblnIsLinkHelper = False Then
        If Not mobjCapLinker Is Nothing Then mobjCapLinker.AddRepPacsHelper mobjStudyInfo.lngAdviceId, ucPacsHelper1
    End If
     
    lngReportID = GetReportId(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    If lngReportID = -1 Then
        ucReportEditor1.ResetContext
        ucReportEditor1.IsEditable = False
   
        '使用非pacs报告编辑器书写的报告
        HintMsg "此检查已使用其他报告编辑器进行书写，不能打开。", "zlRefresh", vbExclamation
        'TASK:可弹出独立的预览窗口
        Exit Sub
    End If
    
    mlngFileFormatId = GetFileFormatId(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    If mblnIsLinkHelper = False Then
        Call ucPacsHelper1.zlRefresh(objStudyInfo, mlngFileFormatId)
        
        If Not mobjCapLinker Is Nothing Then mobjCapLinker.ReportAdviceId = mobjStudyInfo.lngAdviceId
    End If
    
    Call InitReportSampleFormat(mlngFileFormatId)
    Call InitReportPrintFormat(mlngFileFormatId)
    
    Call ucReportEditor1.PromptSave(mobjStudyInfo.lngAdviceId, lngReportID)
    

    Call ucReportEditor1.Refresh(mobjStudyInfo.lngAdviceId, mlngFileFormatId, 0, lngReportID, mobjStudyInfo.blnMoved, True)
    
    If Val(ucReportEditor1.tag) = 0 Then
        If mblnIsLinkHelper = False Then
            strLayout = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "POPUPEDITOR" & mlngModuleId & ucReportEditor1.GetFaceKey, "")
            Call ucReportEditor1.SetLayout(strLayout)
        Else
            strLayout = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\界面设置\" & App.ProductName & "\" & Me.Name, "MAINEDITOR" & mlngModuleId & ucReportEditor1.GetFaceKey, "")
            Call ucReportEditor1.SetLayout(strLayout)
        End If
        
        ucReportEditor1.tag = "1"
    End If
    
    If blnIsHistory Then
'        ucReportEditor1.IsEditable = False
        Call ucReportEditor1.ConfigFaceState(True, "历史查看")
    End If
    
    '判断焦点是否自动定位到报告编辑器
    If mblnSetFocusWithReport = False Then
        Call ResetFocus(objFocus)
    Else
    
        If Not objFocus Is Nothing Then
            If TypeOf objFocus Is PatiIdentify Then
                '如果是查找状态下，不允许自动定位报告编辑
                Exit Sub
            End If
        End If
        
        '如果有对应的弹出式报告编辑窗口，且检查相同，则不定位嵌入式窗口编辑器
        If mblnIsLinkHelper = True And IsSameReportWindow(mobjStudyInfo.lngAdviceId) Then
            Exit Sub
        End If
        
        Call ucReportEditor1.LocateEditBox
    End If
End Sub


Private Function IsSameReportWindow(ByVal lngAdviceId As Long) As Long
'是否存在相同的独立报告书写窗口
    Dim objForm As Object

    IsSameReportWindow = False
Exit Function

'    For Each objForm In Forms
'        If TypeOf objForm Is frmReportV2 Then
'            If objForm.IsLinkHelper = False And objForm.AdviceId = lngAdviceId Then
'                IsSameReportWindow = True
'                Exit Function
'            End If
'        End If
'    Next
End Function

Private Sub ResetFocus(objFocus As Object)
'重置焦点控件
On Error Resume Next
    Call objFocus.SetFocus
End Sub


Public Function PromptSave() As Boolean
'打开独立书写报告窗口时，调用此方法进行提示
    PromptSave = False
    
    If mobjStudyInfo Is Nothing Then Exit Function
    If mobjStudyInfo.lngAdviceId = 0 Then Exit Function
    
    PromptSave = ucReportEditor1.PromptSave(ucReportEditor1.AdviceId, ucReportEditor1.ReportID, True)
End Function


Private Sub InitReportSampleFormat(ByVal lngFileId As Long)
'初始化报告范文格式
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i  As Integer
    
    ReDim rptFormats(1) As rptFormat
    rptFormats(1).ID = 0
    rptFormats(1).strName = "标准格式"
    
    If lngFileId = 0 Then Exit Sub
    
    strSQL = "Select Id,名称 From 病历范文目录 Where 文件ID = [1] And 性质= 0 And (通用级=0 Or (通用级=1 And 科室ID=[2]) " & _
            " Or (通用级=2 And 人员ID= [3])) "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFileId, UserInfo.部门ID, UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        ReDim Preserve rptFormats(rsTemp.RecordCount + 1) As rptFormat
        For i = 1 To rsTemp.RecordCount
            rptFormats(i + 1).ID = rsTemp!ID
            rptFormats(i + 1).strName = rsTemp!名称
            
            rsTemp.MoveNext
        Next i
    End If
    Exit Sub
errH:
    If HintError(err, "InitReportSampleFormat") = 1 Then Resume
End Sub



Private Function IsCustomFont(ByVal intFontSize As Integer) As Boolean
'功能，判断是否使用自定义字号  返回 true-是
'规则，不能与103523字体重复
    IsCustomFont = True
    
    If intFontSize = 0 Or intFontSize = 14 Or intFontSize = 16 Or intFontSize = 22 Or intFontSize = 28 Or intFontSize = 36 Or intFontSize = 42 Then
        IsCustomFont = False
    End If
    
End Function


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'创建该模块内的菜单
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '如果这里不指定id，则不能将有些菜单添加到右键菜单中
    
    If lngIconId <> 0 Then CreateModuleMenu.iconid = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = "" 'M_STR_MODULE_MENU_TAG
End Function

Private Sub cbrMain_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
On Error GoTo errH
    Dim cbrControlItem As CommandBarControl
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset

    If CommandBar.Parent Is Nothing Then Exit Sub
    '添加格式选择弹出菜单（范文）
    If CommandBar.Parent.ID = conMenu_PacsReport_SelFormat Then
        CommandBar.Controls.DeleteAll

        '添加新的菜单项
        For i = 1 To UBound(rptFormats)
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_SelFormat_Item, rptFormats(i).strName, i)
            cbrControlItem.Parameter = rptFormats(i).ID
        Next i
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_RepFormat Then '(打印格式)
        CommandBar.Controls.DeleteAll
        
        If Len(mstr报表编号) <= 0 Then Exit Sub

        '添加新的菜单项
        strSQL = "Select a.编号,b.序号,b.说明 From zlreports a,zlrptfmts b Where a.Id=b.报表ID And a.编号=[1] Order By 序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取自定义报表格式", mstr报表编号)

        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_RepFormat_Item, rsTemp!序号 & "-" & nvl(rsTemp!说明))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = (InStr(mstr选中报表格式, cbrControlItem.Caption) <> 0)
            cbrControlItem.Parameter = rsTemp!序号
            cbrControlItem.CloseSubMenuOnClick = False

            rsTemp.MoveNext
        Wend
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_VerifySign Then
        '签名验证的弹出菜单，列出可以验证的签名版本
        CommandBar.Controls.DeleteAll

        '添加新的签名验证菜单
        strSQL = "Select 开始版,内容文本 as 签名医生 From 电子病历内容 Where 文件ID = [1] And 对象类型 =8  Order By 开始版"
        If ucReportEditor1.IsMoved Then
            strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取各个签名版本", ucReportEditor1.ReportID)

        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_VerifySign_Item, rsTemp!开始版 & "-" & nvl(rsTemp!签名医生))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = False
            cbrControlItem.Parameter = rsTemp!开始版
            rsTemp.MoveNext
        Wend
    End If
    Exit Sub
errH:
    If HintError(err, "cbrMain_InitCommandsPopup") = 1 Then Resume
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error GoTo errhandle
    If mblnHasFace = False Then Exit Sub
    
    picBack.Move Left, Top, Right - Left, Bottom - Top

    If ucPacsHelper1.Visible Then
        Call ucSplitter1.RePaint(False)
    Else
        ucReportEditor1.Move 0, 0, Right - Left, Bottom - Top
    End If
Exit Sub
errhandle:

End Sub
 
Private Sub timerShake_Timer()
On Error GoTo errhandle
    Select Case Val(timerShake.tag)
        Case 0
            Me.Left = Me.Left - 500
            timerShake.tag = 1
        Case 1
            Me.Left = Me.Left + 1000
            timerShake.tag = 2
        Case 2
            Me.Left = Me.Left - 1000
            timerShake.tag = 3
        Case Else
            Me.Left = Me.Left + 500
            
            timerShake.Enabled = False
            timerShake.tag = ""
    End Select
Exit Sub
errhandle:
    timerShake.Enabled = False
    timerShake.tag = ""
    
    HintError err, "", False
    
End Sub

Public Sub Shake()
'抖动
    timerShake.Enabled = True
End Sub


Private Sub ucReportEditor1_OnDelRepImg(ByVal strImgKey As String)
    If mobjCurPacsHelper Is Nothing Then Exit Sub
    If Len(strImgKey) <= 0 Then Exit Sub
    
    Call mobjCurPacsHelper.ClearReportImgState(strImgKey)
    
End Sub

Private Sub ucReportEditor1_OnOutlineChange(ByVal lngSelOutline As TOutlineType)
On Error GoTo errhandle
    If mobjCurPacsHelper Is Nothing Then
        HintMsg "mobjCurPacsHelper 对象无效。", "ucReportEditor1_OnOutlineChange", infNormalErr
        Exit Sub
    End If
    
    Select Case lngSelOutline
        Case otDesc
            Call mobjCurPacsHelper.SyncOutline("所见")
            
        Case otOpin
            Call mobjCurPacsHelper.SyncOutline("意见")
            
        Case otAdvi
            Call mobjCurPacsHelper.SyncOutline("建议")
        
        Case Else
            Call mobjCurPacsHelper.SyncOutline("")
            
    End Select
    
Exit Sub
errhandle:
    If HintError(err, "ucReportEditor1_OnOutlineChange", False) = 1 Then Resume
End Sub
 
