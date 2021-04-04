VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "*\A..\ZL9PACSCONTROL\zl9PacsControl.vbp"
Begin VB.Form frmReportV2 
   Caption         =   "����༭"
   ClientHeight    =   11190
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13320
   Icon            =   "frmReportV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11190
   ScaleWidth      =   13320
   StartUpPosition =   3  '����ȱʡ
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

Private Const M_STR_HINT_NoSelectData As String = "��Ч�ļ�����ݣ�������ѡ��"
Private Const M_STR_MODULE_MENU_TAG As String = "����"
Private Const M_STR_LISTVIEWKEY_DESCRIBE As String = "describe"
Private Const M_STR_LISTVIEWKET_PROCESS As String = "process"

Private Const G_STR_TAG = "Po=Ϣ��[+]E=������[+]M=��Ƕ[+]L=ճĤ�װ�[+]C=ʪ��[+]I=�����԰�[+]W=�����ɫ��Ƥ[+]AT=�쳣ת����[+]V=�ǵ���Ѫ��[+]P=��״Ѫ��[+]Xn=ֱ�ӻ�첿λ"
Private Const conMenu_Process_MartItem As Long = 1000

Private Type rptFormat
    ID As Long          '�����ʽID
    strName As String   '�����ʽ����
End Type
 

Private rptFormats() As rptFormat
Private mstrѡ�б����ʽ As String
Private mstr������ As String
Private mblOneReportFormat As Boolean           '�Ƿ�ֻ��ѡ��һ�ִ�ӡ��ʽ
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
Private mblnExitAfterSign As Boolean    '����ǩ�����˳�
Private mblnSetFocusWithReport As Boolean

Private mblnHasFace As Boolean
Private mObjNotify As IEventNotify
Private mblnHasExitSave As Boolean  '�Ƿ�������˳�ʱ����
 
Public mobjRichReportWrap As frmEPREditWrapV2
 


'ҽ��ID
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


'����༭������
Property Get ReportEditor() As Object
    Set ReportEditor = ucReportEditor1
End Property

Property Get ReportHelper() As Object
    Set ReportHelper = ucPacsHelper1
End Property


'��ȡ�˵��ӿڶ���
Property Get zlMenu() As IWorkMenuV2
    Set zlMenu = Me
End Property


Public Sub SetPrintFmts(ByVal strReportNo As String, ByVal strPrintFmts As String)
    If Len(strReportNo) > 0 Then mstr������ = strReportNo
    If Len(strPrintFmts) > 0 Then mstrѡ�б����ʽ = strPrintFmts
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
    
    strSQL = "Select a.������,a.������,b.������־ ,b.Id From Ӱ�����¼ a ,����ҽ����¼ b Where a.ҽ��id = b.Id And b.Id = [1] "
    If mobjStudyInfo.blnMoved Then
        strSQL = Replace(strSQL, "Ӱ�����¼", "HӰ�����¼")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��֤�Ƿ���Դ�ӡ", mobjStudyInfo.lngAdviceId)

    If rsTemp.EOF = False Then
        AllowPrint = IIf(nvl(rsTemp!������־, 0) = 1, nvl(rsTemp!������) <> "", (mblnCheckPrintPara And nvl(rsTemp!������) <> "") Or mblnCheckPrintPara = False)
    End If

    Exit Function
errH:
    If HintError(err, "AllowPrint") = 1 Then Resume
End Function

Public Sub AddRepImgFile(ByVal strFile As String, Optional ByVal lngImageAdviceId As Long = 0, Optional ByVal strFileName As String = "")
'��ӱ���ͼ�ļ�
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
            
        Case conMenu_Edit_Modify, conMenu_File_Open, conMenu_File_ExportToXML, conMenu_Tool_Search   'ʹ�ò����༭����ʽ���д�
            Call EprEdit(Control)
             
            'TODO=�����༭�еĴ�ӡ״̬ͬ������
            'TODO=���ǩ������Ҫд�븴���� ��ZL_Ӱ�񱨸汣��_Update  Zl_Ӱ����_State
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
        
        Case conMenu_Process_MartItem + 100 '�����ע
            Call ucReportEditor1.ClearMark(True)
            
        Case conMenu_Process_MartItem + 101 '���ñ�ע
            Call SetMarkTextLabel
            
        Case conMenu_Edit_Delete                'ɾ��
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
            
        Case conMenu_PacsReport_Save            '����
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            blnIsNewReport = IIf(ucReportEditor1.ReportID = 0, True, False)
            If ucReportEditor1.SaveReport(strReportImages) Then
                mblnHasExitSave = True
                
                '�״α���ʱ��Ҫˢ���б���
'                If blnIsNewReport Then Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                If ucPacsHelper1.Visible Then
                    Call ucPacsHelper1.SyncReportImgState(strReportImages)
                End If
                
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_SAVE, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, strReportImages)
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
        Case conMenu_File_Print                '��ӡ
            '�ж��Ƿ�ѡ���˶�ݴ�ӡ��ʽ
            If Len(mstrѡ�б����ʽ) > 0 Then
                If UBound(Split(mstrѡ�б����ʽ, ",")) > 0 Then
                    If HintMsg("��ǰ����������ִ�ӡ��ʽ [" & mstrѡ�б����ʽ & "]���Ƿ������", "cbrMain_Execute", vbYesNo) = vbNo Then
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
            
            If ucReportEditor1.ReportPrint(mstr������, mstrѡ�б����ʽ) Then
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
            '��ӡ����ܻ��Զ���ɵȣ������Ҫ��״̬����ˢ��
            Call ucReportEditor1.ConfigFaceState
            
        Case conMenu_File_Preview               'Ԥ��
            Call ucReportEditor1.ReportPreview(mstr������, mstrѡ�б����ʽ)
            
        Case conMenu_PacsReport_AddNumber       '���
            Call ucReportEditor1.AddNumber
            
        Case conMenu_PacsReport_History         '��ʷ
            Call ucReportEditor1.RevisionHistory
            
        Case conMenu_PacsReport_Sign      'ǩ��
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then
                mblnMenuDownState = False
                Exit Sub
            End If
            
            lngResult = ucReportEditor1.Sign    '0-δǩ����1-���ǩ����2-���ǩ��
            
            '�ж��Ƿ���Ҫǩ�����˳�
            If lngResult = 0 Then
                mblnMenuDownState = False
                Exit Sub  'ǩ��ʧ�����˳�
            End If
            
            '��Ҫ���ݱ���ѡ���ʽ����Ϊǩ�����п��ܻ��Զ����д�ӡ����
            Select Case lngResult
                Case 1
    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_SIGN, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, , mstr������ & ":" & mstrѡ�б����ʽ)
                    
                    mobjStudyInfo.blnCanPrint = AllowPrint
                Case 2
    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_AUDIT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, , mstr������ & ":" & mstrѡ�б����ʽ)
                    
                    mobjStudyInfo.blnCanPrint = AllowPrint
            End Select
            
            If mblnExitAfterSign And mblnIsLinkHelper = False Then
                mblnMenuDownState = False
                Unload Me
                Exit Sub
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
            If ucReportEditor1.IsReadOnly Then Call ResetEmbedVideoState(True)
                
        Case conMenu_PacsReport_DelSign         '����
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_BACK, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.SignUntread Then
                mblnHasExitSave = True
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_BACK, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
            Call ucReportEditor1.ConfigFaceState
            
            If ucReportEditor1.IsEditable Then Call ResetEmbedVideoState(False)
            
        Case conMenu_PacsReport_VerifySign_Item '��֤
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_Verify, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.SignVerifiy(Val(Control.Parameter)) Then
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_Verify, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
        Case conMenu_PacsReport_Reject     '����
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            If ucReportEditor1.ReportReject Then
'                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
                Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, , ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
            End If
            
        Case conMenu_View_Refresh
            Call zlRefresh(mobjStudyInfo, True, False)
            
        Case conMenu_PacsReport_RejectHistory   '������ʷ
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJHISTORY, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
            If blnIsCancel Then mblnMenuDownState = False: Exit Sub
            
            Call ucReportEditor1.RejectHistory
            Call mObjNotify.Broadcast(BM_REPORT_EVENT_REJECT, , ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
        
        Case conMenu_PacsReport_SelFormat_Item  '��ʽѡ��
            Call ucReportEditor1.ChangeReportFormat(Val(Control.Parameter))
            
        Case conMenu_PacsReport_RepFormat_Item  '��ӡ��ʽ
            Call ChangeRptFormat(Control.Index)
        
            
        Case conMenu_PacsReport_ClearWritingState   '������桰�����С����
            If mobjStudyInfo.blnMoved Then
                mblnMenuDownState = False
                Call HintMsg("���������ѽ���ת��������ִ�е�ǰ������", "cbrMain_Execute", vbOKOnly)
                
                Exit Sub
            Else
                Call ClearReportState(mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_PrivOrder   '����
            If ucPacsHelper1.Visible Then
                Call SendRequest(WM_LIST_GETLASTADVICE, mobjStudyInfo.lngAdviceId)
            Else
                Call SendRequest(WM_LIST_MOVEUP, mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_NextOrder   '����
            If ucPacsHelper1.Visible Then
                Call SendRequest(WM_LIST_GETNEXTADVICE, mobjStudyInfo.lngAdviceId)
            Else
                Call SendRequest(WM_LIST_MOVEDOWN, mobjStudyInfo.lngAdviceId)
            End If
            
        Case conMenu_PacsReport_FontSet, conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser            '�����ı�������
            Dim cbrEdit As CommandBarEdit
                
            Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
        
            If Control.ID = conMenu_PacsReport_FontSetUser Then
            '������Զ����ֺţ��ж��Ƿ���Ϲ���
                
                If Not CheckUserFontValidate(cbrEdit.Text) Then
                '�����Ϲ����൱������ʧ��
                    cbrEdit.Text = ""
                    mblnMenuDownState = False
                    Exit Sub
                End If
                
                ucReportEditor1.EditFontSize = Abs(Val(cbrEdit.Text))
                Call zlDatabase.SetPara("������ʾ�ֺ�", (Abs(Val(cbrEdit.Text))), glngSys, glngModul)
            Else
            '�����Զ����ֺţ�ǰ��򹴣��Զ���textΪ�ձ�ʾδѡ���Զ����ֺ�
                cbrEdit.Text = ""
                Control.Checked = True
                
                If Val(Control.Caption) = 0 Then
                    ucReportEditor1.EditFontSize = FontSize
                    Call zlDatabase.SetPara("������ʾ�ֺ�", FontSize, glngSys, glngModul)
                Else
                    ucReportEditor1.EditFontSize = Val(Control.Caption)
                    Call zlDatabase.SetPara("������ʾ�ֺ�", Val(Control.Caption), glngSys, glngModul)
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
        If mobjCapLinker.LockAdviceId <> 0 And mobjCapLinker.LockAdviceId <> mobjStudyInfo.lngAdviceId Then Exit Sub    '�����ɼ��͵�ǰ���治����ͬ���ʱ�˳�
        
        Call ucPacsHelper1.EmbedVideo.zlRestoreWindow(blnReadState And mobjCapLinker.ReportAdvReadOnly, , True)
    End If
End Sub

Public Sub ReadRepStateTag()
    If ucReportEditor1 Is Nothing Then Exit Sub
    
    Call ucReportEditor1.ReadResultTag(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
End Sub

Private Sub SetMarkTextLabel()
'------------------------------------------------
'���ܣ��������ֱ�ע��������
'������
'���أ���
'------------------------------------------------
    Dim strTemp As String
    Dim i As Integer
    Dim objMenu As CommandBarControl
    On Error GoTo err
    
'    strTemp = InputBox("�������µ����ֱ�ע���ã���ʽΪ������1=˵��1|����2=˵��2|...����", "���ֱ�ע����", Replace(mstrTemp, "[+]", "|"))
    
    strTemp = frmInputBoxV2.ZlShowMe(mObjNotify.Owner, mStrTextMarks)
    
    
    If strTemp = "" Then Exit Sub
    
    If InStr(strTemp, "=") = 0 Then
        HintMsg "����ĸ�ʽ����ȷ��Ӧ�ð��ա�����=˵������ʽ���룬������������á�", "SetMarkTextLabel", vbOKOnly
        Exit Sub
    End If

    mStrTextMarks = strTemp
    Call SaveSetting("ZLSOFT", "����ģ��\zl9PACSWork\frmReportImageEdit", "�������ֱ�ע", Replace(strTemp, "|", "[+]"))

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
'���ı�ѡ�е��Զ��屨���ӡ��ʽ
    Dim cbrRptFormat As CommandBarControl
    Dim cbrRptFormatItem As CommandBarControl
    
    Dim i As Integer
    
    On Error GoTo err
    
    Set cbrRptFormat = cbrMain.FindControl(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, True)
    
    mstrѡ�б����ʽ = ""
    
    If mblOneReportFormat Then
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If i = lngIndex Then
                cbrRptFormatItem.Checked = True
                mstrѡ�б����ʽ = cbrRptFormatItem.Caption
            Else
                cbrRptFormatItem.Checked = False
            End If
        Next i
    Else
        For i = 1 To cbrRptFormat.CommandBar.Controls.Count
            Set cbrRptFormatItem = cbrRptFormat.CommandBar.Controls(i)
            If cbrRptFormatItem.Index = lngIndex Then cbrRptFormatItem.Checked = Not cbrRptFormatItem.Checked
            If cbrRptFormatItem.Checked = True Then
                mstrѡ�б����ʽ = IIf(mstrѡ�б����ʽ = "", cbrRptFormatItem.Caption, mstrѡ�б����ʽ & "," & cbrRptFormatItem.Caption)
            End If
        Next i
    End If
    
    SaveSetting "ZLSOFT", mstrRegPath, "������", mstr������
    SaveSetting "ZLSOFT", mstrRegPath & "\" & mstr������, "ѡ�б����ʽ", mstrѡ�б����ʽ
    
    ucReportEditor1.ShowPrintFormat (Replace(mstrѡ�б����ʽ, ",", "  "))
    
    Exit Sub
err:
    If HintError(err, "ChangeRptFormat", False) = 1 Then
        Resume
    End If
End Sub


Private Function CheckUserFontValidate(ByVal strValue As String) As Boolean
'���򣺾���abs(val(?))����������֣�������֤��ͨ��������ʾ

    CheckUserFontValidate = True
    
    If Not IsNumeric(strValue) Or Val(strValue) < 1 Or Val(strValue) > 80 Then
        Call HintMsg("�Զ����ֺ�ֻ������Ϊ1��80��һ�����֣�����������", "CheckUserFontValidate", vbOKOnly)
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
'�ж��Ƿ��Ѿ��򿪵���ʽ���洰��
    Dim objForm As Object
    
    Set GetExistsWindow = Nothing
    
    '�ж��Ƿ�����Ѿ��򿪵ı���༭��
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
            mObjNotify.SendRequest lngEventNo   'sendRequest��,mobjStudyInfo��ҽ��ID��ͬ������Ϊ���¼���ҽ��ID
            
            If lngNewAdviceId = mobjStudyInfo.lngAdviceId Then
                '�����Ǳ��洰�ڵ�msgboxd��parent����Ϊme
                HintMsg "���ƶ���" & IIf(lngEventNo = WM_LIST_MOVEUP, "��ʼ����С�", "ĩβ����С�"), "SendRequest", vbOKOnly
                Exit Sub
            End If
            
        Case WM_LIST_GETLASTADVICE, WM_LIST_GETNEXTADVICE
            lngNewAdviceId = lngMainAdviceId
            mObjNotify.SendRequest lngEventNo, , lngNewAdviceId, lngSendNo, blnIsMoved  'lngNewAdviceId�����ƶ������µ�ҽ��ID
            
            If lngNewAdviceId = lngMainAdviceId Then
                '�����Ǳ��洰�ڵ�msgboxd��parent����Ϊme
                HintMsg "���ƶ���" & IIf(lngEventNo = WM_LIST_GETLASTADVICE, "��ʼ����С�", "ĩβ����С�"), "SendRequest", vbOKOnly
                Exit Sub
            End If
            
            Set objStudyInfo = zlGetStudyAdvice(lngNewAdviceId)
            
            If objStudyInfo Is Nothing Then
                HintMsg "δ��ȡ����Ӧ�ļ����Ϣ��", "SendRequest", infNormalErr
                Exit Sub
            End If
            
            If objStudyInfo.intStep <= 1 Then
                '����δ�������Ѿܾ��ļ��
                Call SendRequest(lngEventNo, objStudyInfo.lngAdviceId)
                Exit Sub
            End If
            
            '�ж϶�Ӧ��鱨���Ƿ��Ѿ����������ڴ򿪣�����Ѿ��򿪣��������ʾ,�������Ѿ��򿪵ı���
            Set objForm = GetExistsWindow(lngNewAdviceId)
            If Not objForm Is Nothing Then
                If HintMsg("��Ӧ���洰���Ѿ��򿪣��Ƿ��л���", "IsAllowExistsWindowChange", vbYesNo) = vbNo Then
                    Call SendRequest(lngEventNo, objStudyInfo.lngAdviceId)
                Else
                    objForm.WindowState = 0
                    objForm.Visible = True
                    objForm.ZOrder
                    
                    '��������
                    Call objForm.Shake
                End If
                
                Exit Sub
            End If
            
            Call ucReportEditor1.PromptSave(lngNewAdviceId, 0)
      
            If mblnIsLinkHelper = False Then
                '�Ƴ�֮ǰ��pacshelper�������
                If Not mobjCapLinker Is Nothing Then mobjCapLinker.RemoveRepPacsHelper mobjStudyInfo.lngAdviceId
            End If
            
            Call zlRefresh(objStudyInfo)
            Call SetReportTitle(objStudyInfo)
            
        Case WM_LIST_SYNCROW    'ͬ����ʾ��
            Call mObjNotify.SendRequest(lngEventNo, , lngMainAdviceId)
            
    End Select
    
    Set objStudyInfo = Nothing
End Sub

Private Sub ClearReportState(ByVal lngAdviceId As Long)
'�������״̬���
    Dim strSQL As String
    Dim strInfo As String
    Dim rsTemp As ADODB.Recordset
    
    strSQL = "select ������� from Ӱ�����¼ where ҽ��id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���������", lngAdviceId)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If Trim(nvl(rsTemp!�������)) = "" Then Exit Sub
    
    
    strInfo = "�������Ŀǰ״̬Ϊ [" & nvl(rsTemp!�������) & "] �����У�ȷ��Ҫ�����ݱ����״̬��"
    If HintMsg(strInfo, "ClearReportState", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    
    Call UpdateReporter(lngAdviceId, "")
    Call ucReportEditor1.ConfigFaceState
End Sub


Private Sub InitReportPrintFormat(ByVal lngFileId As Long)
'��ʼ�������ӡ��ʽ
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strRegReportNo As String
    
    mstrRegPath = "����ģ��\" & App.ProductName & "\frmReport"
    
    mstr������ = ""
    mstrѡ�б����ʽ = ""
    
    '���ж��Ƿ�ʹ���Զ��屨��
    strSQL = "Select ͨ��,��� From �����ļ��б�  Where Id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ӡ��ʽ", lngFileId)
    If rsTemp.EOF = False Then
        If nvl(rsTemp!ͨ��) = 2 Then
            'ʹ���Զ��屨���ʽ��ӡ
            mstr������ = "ZLCISBILL" & Format(nvl(rsTemp!���), "00000") & "-2"
            strRegReportNo = GetSetting("ZLSOFT", mstrRegPath, "������", "")
            
            If mstr������ <> strRegReportNo Then
                '��ע����еı������뵱ǰ�ı����Ų���ͬʱ������ע����б�����
                '�Ա������������ӡʱ��������ǰһ�εı�����д�ӡ
                SaveSetting "ZLSOFT", mstrRegPath, "������", mstr������
            End If
            
            mstrѡ�б����ʽ = GetSetting("ZLSOFT", mstrRegPath & "\" & mstr������, "ѡ�б����ʽ", "")
        End If
    End If
    
    Call ucReportEditor1.ShowPrintFormat(Replace(mstrѡ�б����ʽ, ",", "  "))
End Sub


Private Sub InitCommandBar()
    '���ܴ���������
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize True, 24, 24
    End With
    
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    If Me.cbrMain.Count > 1 Then Call Me.cbrMain(2).Delete
    '���湤��������
    Set cbrToolBar = Me.cbrMain.Add("������", IIf(mblnIsLinkHelper, xtpBarRight, xtpBarTop))
    
'    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Specialty, "ר��"): cbrControl.iconid = 2558: cbrControl.ToolTipText = "ר�Ʊ���༭"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Save, "����"): cbrControl.iconid = 3503: cbrControl.ToolTipText = "���汨��": cbrControl.BeginGroup = True

'        Set cbrControl = .Add(xtpControlSplitButtonPopup, 8352, "��ӡ")
'        cbrControl.IconId = 103
'        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
                cbrControl.BeginGroup = True
                cbrControl.iconid = 103
                cbrControl.ToolTipText = "�����ӡ"
                
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
                cbrControl.iconid = 102
                cbrControl.ToolTipText = "����Ԥ��"
'        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Sign, "ǩ��")
            cbrControl.iconid = 3003
            cbrControl.ToolTipText = "ǩ��"
            cbrControl.BeginGroup = True
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_DelSign, "����")
            cbrControl.iconid = 3004
            cbrControl.ToolTipText = "����ǩ��"
        
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_VerifySign, "��֤")
            cbrControl.iconid = 8044
            cbrControl.ToolTipText = "��֤"
            
            
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            cbrControl.ToolTipText = "ˢ��"
            cbrControl.BeginGroup = True
            

        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_RepFormat, "��ʽ")
            cbrControl.iconid = 3031
            cbrControl.ToolTipText = "ѡ���Զ��屨���ӡ��ʽ"
            
            
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_SelFormat, "ģ��")
            cbrControl.iconid = 227
            cbrControl.ToolTipText = "ѡ��͸������浥��дģ��"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����")
            cbrControl.iconid = 3002
            cbrControl.ToolTipText = "�õ��Ӳ�����ʽ�༭����"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_History, "�޶�ʷ")
            cbrControl.iconid = 5015
            cbrControl.ToolTipText = "�鿴��ǰ����ʷ����ĵ��޶����"
            
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_AddNumber, "���")
            cbrControl.BeginGroup = True
            cbrControl.iconid = 9023
            cbrControl.ToolTipText = "����������������"
            
        Set cbrControl = .Add(xtpControlSplitButtonPopup, conMenu_Process_AddMark, "���")
            cbrControl.iconid = 3034
            cbrControl.ToolTipText = "�����ͼ��ӱ��"
            
            Call LoadMark(cbrControl)
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_Reject, "����")
            cbrControl.iconid = 229
            cbrControl.ToolTipText = "���沵��"
            
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_RejectHistory, "����ʷ")
            cbrControl.iconid = 8341
            cbrControl.ToolTipText = "������ʷ"
            
                
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_PrivOrder, "����")
            cbrControl.BeginGroup = True
            cbrControl.iconid = 21802
            cbrControl.ToolTipText = "��һ������¼"
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsReport_NextOrder, "����")
            cbrControl.iconid = 21801
            cbrControl.ToolTipText = "��һ������¼"
    

        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_PacsReport_FontSet, "�ֺ�")    'xtpControlSplitButtonPopup
            cbrControl.iconid = 509
            cbrControl.ToolTipText = "��������"
            With cbrControl
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSetDefault, "Ĭ��", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet14, "14", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet16, "16", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet22, "22", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet28, "28", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet36, "36", "", 0, False)
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlButton, conMenu_PacsReport_FontSet42, "42", "", 0, False)
                
                
                Set cbrPopControl = CreateModuleMenu(.CommandBar.Controls, xtpControlEdit, conMenu_PacsReport_FontSetUser, "�Զ���", "", 0, False)
                
                If mintContextFontSize <> 0 And IsCustomFont(mintContextFontSize) Then
                    Set cbrEdit = cbrMain.FindControl(xtpControlEdit, conMenu_PacsReport_FontSetUser, True, True)
                    cbrEdit.Text = mintContextFontSize
                End If
                
            End With

        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If (cbrControl.type = xtpControlButton) Or (cbrControl.type = xtpControlSplitButtonPopup) Then cbrControl.Style = xtpButtonIconAndCaption
        If cbrControl.Category = "" Then cbrControl.Category = "Main" '���ó�������˵�
    Next
    
    cbrToolBar.Position = IIf(mblnIsLinkHelper, xtpBarRight, xtpBarTop)
End Sub


Private Sub LoadMark(mnuParent As Object)
'�������ֱ�ע
    Dim objControl As CommandBarControl
    Dim arrTemp() As String
    Dim i As Long
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem, "��")
        objControl.ToolTipText = "�Զ��������ֱ��"
        objControl.Category = 0
        objControl.iconid = 0
    
    For i = 1 To 6
        Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + i, i)
            objControl.ToolTipText = "���ֱ��" & i
            objControl.Category = i
            objControl.iconid = 0
    Next
     
    arrTemp = Split(mStrTextMarks, "|")
    
    For i = 0 To UBound(arrTemp)
        If Len(arrTemp(i)) > 0 Then
            Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 10 + i + 1, arrTemp(i))
                If i = 0 Then objControl.BeginGroup = True
                
                objControl.ToolTipText = "�ı���ע"
                objControl.Category = i + 1
                objControl.iconid = 0
        End If
    Next
    
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 100, "���")
        objControl.BeginGroup = True
        objControl.ToolTipText = "������"
        objControl.Category = 0
        objControl.iconid = 0
        
    Set objControl = mnuParent.CommandBar.Controls.Add(xtpControlButton, conMenu_Process_MartItem + 101, "����")
        objControl.BeginGroup = True
        objControl.ToolTipText = "���ñ��"
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
        
        Case conMenu_File_Print, conMenu_File_Preview      '��ӡ����,Ԥ������

            Control.Visible = CheckPopedom(mstrPrivs, "PACS�����ӡ")
            If Control.Visible = False Then Exit Sub
            
            '���δ�ҵ���Ӧ�Ĳ����ļ�����ô��ӡԤ����ť�ᱻ����
            If mlngFileFormatId = 0 Then
                Control.Enabled = False
            Else
                Control.Enabled = ucReportEditor1.ReportID <> 0
            End If
            
            If Control.Enabled Then Control.Enabled = mobjStudyInfo.blnCanPrint

        Case conMenu_Edit_Modify        '����༭���Բ����༭����ʽ���б༭��
            Control.Enabled = ucReportEditor1.IsEditable
            

        Case conMenu_PacsReport_Save    '����
            Control.Enabled = ucReportEditor1.IsModify

        Case conMenu_PacsReport_Reject
            '�ж��Ƿ�߱����沵��Ȩ��
            '�жϵ�ǰ��������״̬�Ƿ�������
            Control.Visible = CheckPopedom(mstrPrivs, "���沵��")
            If Control.Visible Then Control.Enabled = ucReportEditor1.ReportID <> 0 And Not ucReportEditor1.IsReadOnly

        Case conMenu_PacsReport_RejectHistory
            Control.Visible = True
            If Control.Visible Then Control.Enabled = ucReportEditor1.ReportID <> 0


        Case conMenu_PacsReport_Sign    'ǩ�����ܱ༭��д����һ�㶼�ܽ���ǩ�������ǩ����

            '����дģʽ�£���û��ǩ���ģ�����ǩ��
            '���޶�ģʽ�£�ǩ������û�г���16�εģ�����ǩ����
            'ֻ��ģʽ�£�ʲô�����ܲ�����
            Control.Enabled = (Not ucReportEditor1.IsReadOnly And ucReportEditor1.ReportID > 0) Or (ucReportEditor1.IsModify)

        Case conMenu_PacsReport_VerifySign  'ǩ����֤
            'ֻ������������ǩ��������ʾǩ����֤��ť
            'ֻ�б�����д�������޶�Ȩ�޵��ˣ����ܶ�ǩ��������֤
            Control.Visible = IIf(ucReportEditor1.SignPassType = 0, False, True)
            If Control.Visible Then Control.Enabled = IIf(ucReportEditor1.SourceVer >= 1, True, False)
            
        Case conMenu_PacsReport_DelSign '����

            'û��ǩ��֮ǰ�������Ի���,ֻ�ܻ����Լ���ǩ��������ͨ������������ǩ������Ȩ�ޣ����˱����������˵�ǩ��
            'ֻ��ǩ������ſ��Ի���
            '�����Լ���ǩ��
             '�����˱���Ȩ�޵�,���Ի��˱����ҵ�����ǩ��
             Control.Enabled = ucReportEditor1.SourceVer >= 1 And (Not ucReportEditor1.IsReadOnly)
             
        Case conMenu_View_Refresh
            Control.Visible = Not mblnIsLinkHelper

        Case conMenu_PacsReport_SelFormat  'ѡ���ʽ '�޶�ģʽ�£����������ø�ʽ
            Control.Enabled = ucReportEditor1.IsEditable And IIf(ucReportEditor1.SourceVer < 1, True, False)
            
        Case conMenu_PacsReport_SelFormat_Item
            Control.Checked = IIf(Val(Control.Parameter) = ucReportEditor1.SampleId, True, False)
            
        Case conMenu_PacsReport_RepFormat   'ѡ���ӡ��ʽ
            Control.Visible = IIf(Len(mstr������) > 0, True, False)

        Case conMenu_PacsReport_RepFormat_Item  'ѡ������ӡ��ʽ
            Control.Checked = InStr(mstrѡ�б����ʽ, Control.Caption)
            Control.iconid = IIf(Control.Checked, 90002, 90001)
 
        Case conMenu_PacsReport_FontSet, conMenu_PacsReport_FontSetDefault To conMenu_PacsReport_FontSetUser   '�����ֺ�
            Control.Checked = False
            If Val(Control.Caption) = 0 Then
                If FontSize = ucReportEditor1.EditFontSize Then Control.Checked = True
            Else
                If Val(Control.Caption) = ucReportEditor1.EditFontSize Then Control.Checked = True
            End If
 
        Case conMenu_Edit_Delete                            'ɾ������
            '�����˱���ͱ���ɾ��Ȩ��ʱ������ǿ��ɾ����������������д�ı���
            '��ǩ�����治����ɾ��
            Control.Visible = (ucReportEditor1.ReportID <> 0 And (CheckPopedom(mstrPrivs, "PACS������д") Or CheckPopedom(mstrPrivs, "PACS����ɾ��")))
            If Control.Visible Then
                'ɾ���Լ���д��δǩ���ı������ͬ���ҵ�����δǩ������
                Control.Enabled = IIf(ucReportEditor1.SourceVer < 1, True, False) _
                                    And (ucReportEditor1.CreateUser = UserInfo.���� _
                                        Or (CheckPopedom(mstrPrivs, "PACS���˱���") _
                                            And CheckPopedom(mstrPrivs, "PACS����ɾ��") _
                                            And ucReportEditor1.CreateDeptId = mlngDeptID _
                                            ) _
                                        ) And (Not ucReportEditor1.IsReadOnly Or Not ucReportEditor1.IsComplete)
            End If

        Case conMenu_PacsReport_ClearWritingState       '������桰�����С���״̬,������������ҵı�����
            Control.Visible = CheckPopedom(mstrPrivs, "PACS����ɾ��")
            
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
'ɾ������
    DelReport = False
    
    If HintMsg("����ɾ���󽫲��ָܻ����Ƿ������", "DelReport", vbYesNo) = vbNo Then Exit Function
    
    If ucReportEditor1.DelReportData(ucReportEditor1.ReportID, True) = False Then Exit Function
    
    '���������
    Call ucReportEditor1.UnlockEditor
    
    '�������¼������
    Call ucReportEditor1.ClearReport(True, True, True)
    
    '������
    Call ucReportEditor1.ClearMark(False)
    
    '�������ͼ
    Call ucReportEditor1.ClearReportImg
    
    '���������Ϣ
    Call ucReportEditor1.ClearInfo
    
    '�ָ���ʼ״̬
    Call ucReportEditor1.ConfigFaceState
    
    DelReport = True
End Function




Private Sub Form_Activate()
On Error GoTo errhandle
 
'    Debug.Print "TopWindow:" & GetTopWindow(App.hInstance) & " ForeWindow:" & GetForegroundWindow & "  CurWindow:" & Me.hWnd
    If mblnHasFace = False Then Exit Sub
    
    If mblnIsLinkHelper = False Then    '���û������helper�����ʾ����ʽ���洰��
        '����ʽ���洰����Ҫ�ж�GetForegroundWindow����Ƿ��뵱ǰ���ھ����ͬ�������ͬ��˵�����ǵ�ǰ�ö��ı���༭����
        If GetForegroundWindow <> Me.hwnd Then Exit Sub
    End If
    
    If mblnIsLinkHelper = False Then
        If ucPacsHelper1.AllowEmbedVideo Then
            '������洰��Ƕ������Ƶ�ɼ������л�����Ӧ���洰�ں���û�������ɼ�����£���ͬ���ɼ������Ӧ����ͼ��
            '�����Ƶ����Ƕ��ɹ�������Ҫ����caplinker�ı���ҽ��idΪ��ǰҽ��ID
            If Not mobjCapLinker Is Nothing Then
                mobjCapLinker.ReportAdviceId = mobjStudyInfo.lngAdviceId
                
                If ucPacsHelper1.ShowEmbedVideo(mobjCapLinker) = False Then
                    '�����Ƶ�ɼ�Ƕ��ʧ�ܣ�������caplinker�ı���ҽ��idΪ0
                    If Not mobjCapLinker Is Nothing Then mobjCapLinker.ReportAdviceId = 0
                End If
            Else
                ucPacsHelper1.HideEmbedVideo
            End If
        End If
    Else
        'Ƕ��ʽ����༭���ڴ���
        If Not mobjCapLinker Is Nothing And VideoIsAttachReportWindow = False Then
            mobjCapLinker.ReportAdviceId = 0
            
            Call mobjCurPacsHelper.ShowEmbedVideo(mobjCapLinker)
        End If
    End If
     
    '��λ�����ݱ༭��ͬʱtab�л��󣬱༭�򽹵�ɽ��лָ�
    If mblnIsLinkHelper = False Or mblnSetFocusWithReport Then
        If Not mObjNotify.Owner.ActiveControl Is Nothing Then
            If TypeOf mObjNotify.Owner.ActiveControl Is PatiIdentify Then
                '����ǲ���״̬�£��������Զ���λ����༭
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
'�ж���Ƶ�Ƿ�Ƕ��ĵ���ʽ���洰��
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
    '�����������ݱ�����ʾ
    If mblnIsLinkHelper = False Then
        mblnHasExitSave = mblnHasExitSave Or PromptSave
    
        '�ж��Ƿ���Ҫͬ�����´ʾ�Ƭ��
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
        '����ʽ������Ҫ���ʹ����˳���Ϣ�Ա�Ƕ��ʽ����ˢ��
        If mblnHasExitSave Then Call SendExitMsg
    End If

    Call ucReportEditor1.UnlockEditor
    
    strLayoutKey = ucReportEditor1.GetFaceKey
    strLayoutStr = ucReportEditor1.GetLayoutStr()
    
    If mblnIsLinkHelper = False Then
        Call SaveWinState(Me)
        '��Ƕ��ʽ���洰�ڱ���ucpacsHelper���ִ�
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HELPER" & mlngModuleId, ucPacsHelper1.GetLayoutStr)
        
        If Me.ScaleWidth > 0 Then
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HelperWidth" & mlngModuleId, ucPacsHelper1.Width)
            Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "EditorWidth" & mlngModuleId, ucReportEditor1.Width)
        End If
        
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "POPUPEDITOR" & mlngModuleId & strLayoutKey, strLayoutStr)
    Else
        '����Ƕ��ʽ����״̬
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "MAINEDITOR" & mlngModuleId & strLayoutKey, strLayoutStr)
    End If
    
    'Ƕ��������Ƶ����...
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

'�ӿ�ʵ�ֲ���*********************************************************************************

Public Function IWorkMenuV2_zlBaseMenuID() As Long
End Function

Public Function IWorkMenuV2_zlExecuteCmd(ByVal lngCmdType As Long)
'ִ�в˵�����

End Function
 
Public Function IWorkMenuV2_zlIsModuleMenu(ByVal strModuleName As String, objControlMenu As XtremeCommandBars.ICommandBarControl) As Boolean
'�жϲ˵��Ƿ����ڸ�ģ��˵�
    IWorkMenuV2_zlIsModuleMenu = IIf(objControlMenu.Category = M_STR_MODULE_MENU_TAG, True, False)
End Function


Public Sub IWorkMenuV2_zlCreateMenu(ByVal strModuleName As String, objMenuBar As Object)
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set mObjActiveMenuBar = objMenuBar
     
    Set cbrMenuBar = mObjActiveMenuBar.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "����", 3, False)
    cbrMenuBar.ID = conMenu_EditPopup
    cbrMenuBar.Category = ""
    
    With cbrMenuBar.CommandBar
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 3002, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_PacsReport_ClearWritingState, "���״̬", "", 21903, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Edit_Delete, "ɾ��", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_Open, "����", "", 0, True)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_File_ExportToXML, "����XML��", "", 0, False)
        Set cbrControl = CreateModuleMenu(.Controls, xtpControlButton, conMenu_Tool_Search, "���������", "", 0, False)
    End With
End Sub


Public Sub IWorkMenuV2_zlCreateToolBar(ByVal strModuleName As String, objToolBar As Object)
''����������
    Dim cbrControl As CommandBarControl
    Dim cbrLogOut As CommandBarControl
    Dim lngIndex As Long

    Set cbrLogOut = objToolBar.FindControl(, conMenu_Manage_InQueue, , True)

    lngIndex = 4
    If Not cbrLogOut Is Nothing Then lngIndex = cbrLogOut.Index

    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "����Ԥ��", 102, True, lngIndex + 1)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "�����ӡ", 103, False, lngIndex + 2)
    Set cbrControl = CreateModuleMenu(objToolBar.Controls, xtpControlButton, conMenu_PacsReport_Open, "��д", "", 2607, False, lngIndex + 3) 'IconId=3002
End Sub


Public Sub IWorkMenuV2_zlClearMenu(ByVal strModuleName As String)
'����������Ĳ˵�
    Exit Sub
End Sub


Public Sub IWorkMenuV2_zlClearToolBar(ByVal strModuleName As String)
'��������Ĺ�����
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
'        'ͨ��idִ�ж�Ӧ����
'        Select Case lngMenuId
'            Case conMenu_File_BatPrint  '������ӡ����
'               Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 0, ucReportEditor1.AdviceId, ucReportEditor1.ReportID, blnIsCancel)
'                If blnIsCancel Then mblnMenuDownState = False: Exit Sub
'
'                If ucReportEditor1.ReportPrint(mstr������, mstrѡ�б����ʽ, True) Then
'    '                Call SendRequest(WM_LIST_SYNCROW, mobjStudyInfo.lngAdviceId)
'                    Call mObjNotify.Broadcast(BM_REPORT_EVENT_PRINT, 1, ucReportEditor1.AdviceId, ucReportEditor1.ReportID)
'                End If
'        End Select
        
        Exit Sub
    End If
    
    Call cbrMain_Execute(objControl)
End Sub


Public Sub IWorkMenuV2_zlPopupMenu(ByVal strModuleName As String, objPopup As XtremeCommandBars.ICommandBar)
'�����Ҽ��˵�
    Exit Sub
End Sub

Public Sub IWorkMenuV2_zlRefreshSubMenu(ByVal strModuleName As String, objMenuBar As Object)
'ˢ�µ������Ӳ˵�
    Exit Sub
End Sub

'*************************************************************************************************


Public Sub LocateEditBox()
'��λ�༭��
    ucReportEditor1.LocateEditBox
End Sub

Public Sub ReSetFormFontSize(ByVal bytFontSize As Byte)
'����վ�˵����ı��ֺ�
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
'���ܣ��޸�mblnMenuDownState��ֵ�����ڴ�������105988
    mblnMenuDownState = blnValue
End Sub


Public Function PrintPreview(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean, _
    Optional ByVal blnIsPrint As Boolean = False, Optional ByVal lngSpecifyReportId As Long = 0, _
    Optional ByVal strPrintFmts As String = "") As Boolean
'��ӡ��Ԥ��
    If blnIsPrint Then
        PrintPreview = ucReportEditor1.ReportPrintEx(lngAdviceId, blnIsMoved, lngSpecifyReportId, mblOneReportFormat, strPrintFmts)
    Else
        Call ucReportEditor1.ReportPreviewEx(lngAdviceId, blnIsMoved, lngSpecifyReportId, mblOneReportFormat)
    End If
    
    
End Function

Public Sub ReinitWordChar()
'ͬ�����ôʾ�
    Call ucReportEditor1.InitReportChar
End Sub

Public Sub ReinitWordFragment()
'ˢ�´ʾ�Ƭ��
    Call mobjCurPacsHelper.RefreshData("�ʾ�")
End Sub


Public Sub zlInit(objNotify As IEventNotify, ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    ByVal strPrivs As String, objCapLinker As Object, Optional objMainPacsHelper As Object = Nothing, _
    Optional ByVal blnHasFace As Boolean = True)
    Dim strLayout As String
'��ʼ��
    mblnIsLinkHelper = False
    mlngModuleId = lngModuleNo
    mlngDeptID = lngDeptId
    mstrPrivs = strPrivs
    mblnHasFace = blnHasFace
    
    Set mObjNotify = objNotify
 
    '��ʼ������
    Call InitParameters(lngDeptId)
    
    If blnHasFace Then
        Set mobjCapLinker = objCapLinker
        
        '�ж��Ƿ���Ҫ�̳���ʾ�ʾ䣬��ʷ��ͼ��ȸ���ģ��
        Set mobjCurPacsHelper = ucPacsHelper1
        
        If Not objMainPacsHelper Is Nothing Then
            'Ƕ��ʽ����
            mblnIsLinkHelper = True
            Set mobjCurPacsHelper = objMainPacsHelper
            
            ucPacsHelper1.Visible = False
            ucSplitter1.Visible = False
        Else
            '����ʽ����
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
         
         '���Ե������ڵ�pacshelper�Ŀ��
        ucPacsHelper1.Width = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HelperWidth" & mlngModuleId, 750))
        ucReportEditor1.Width = Val(GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "EditorWidth" & mlngModuleId, 1000))
        
        strLayout = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "HELPER" & mlngModuleId, "")
        Call ucPacsHelper1.SetLayout(strLayout)
    End If
End Sub





Private Sub InitParameters(ByVal lngDeptId As Long)

    mblnCheckPrintPara = Val(GetDeptPara(mlngDeptID, "ƽ������˲��ܴ򱨸�", 0)) <> 0
    mblnSetFocusWithReport = Val(GetDeptPara(mlngDeptID, "����л�ʱ��λ����༭", 1)) = 1
    
    mblnExitAfterSign = IIf(Val(zlDatabase.GetPara("PACS����ǩ�����˳�", glngSys, mlngModuleId, True, "0")) = 0, False, True)
    mintContextFontSize = Val(zlDatabase.GetPara("������ʾ�ֺ�", glngSys, mlngModuleId))
    mblOneReportFormat = GetDeptPara(lngDeptId, "��ѡ�����ʽ", True)
    mStrTextMarks = GetSetting("ZLSOFT", "����ģ��\zl9PACSWork\frmReportImageEdit", "�������ֱ�ע", G_STR_TAG)
    
    mStrTextMarks = Replace(mStrTextMarks, "[+]", "|")
End Sub



Public Function GetFileFormatId(ByVal lngAdviceId As Long, ByVal blnIsMoved As Boolean) As Long
'��ȡ����Ӧ�����Ƶ��ݸ�ʽID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetFileFormatId = 0
    
    strSQL = "Select l.������Դ, a.�����ļ�id" & vbNewLine & _
            " From ����ҽ����¼ l, ��������Ӧ�� a" & vbNewLine & _
            " Where l.������Ŀid = a.������Ŀid(+) And a.Ӧ�ó���(+) = Decode(l.������Դ, 2, 2, 4 ,4, 1) And l.Id = [1]"
            
    If blnIsMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
            
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���ݸ�ʽ", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    GetFileFormatId = Val(nvl(rsData!�����ļ�id))
    
End Function

Public Function GetReportId(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean) As Long
'��ȡ����Ӧ�ı���ID
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    GetReportId = 0
    strSQL = "Select ����ID,RawToHex(��鱨��ID) ��鱨��ID From ����ҽ������ Where ҽ��ID= [1]"
    If blnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "��ѯ����ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then Exit Function
    
    If nvl(rsData!��鱨��ID) <> "" Then
        GetReportId = -1
    Else
        GetReportId = Val(nvl(rsData!����Id))
    End If
    
End Function


Public Sub SetReportTitle(objStudyInfo As clsStudyInfo)
    Me.Caption = "����༭    " & objStudyInfo.strPatientName & " (����:" & objStudyInfo.strStudyNum & Decode(objStudyInfo.lngPatientFrom, 1, "  �����:" & objStudyInfo.strMarkNum, 2, "  סԺ��:" & objStudyInfo.strMarkNum, "") & ")    " & objStudyInfo.strPatientAge & "    " & objStudyInfo.strPatientSex & "    " & objStudyInfo.strAdviceContext
End Sub

Public Sub SyncHelper(ByVal lngAdviceId As Long, ByVal lngSourceHwnd As Long, ByVal lngSyncType As Long)
'lngAdviceId:ҽ��ID
'lngSourceHwnd:�����÷�����ԭʼ�ؼ����
'lngSyncType:ͬ������0-ͼ��1-�ʾ�,  2-��ʷ   3-����
     If lngAdviceId <> mobjStudyInfo.lngAdviceId Then Exit Sub
     If ucPacsHelper1.Visible = False Then Exit Sub
     If lngSourceHwnd = Me.hwnd Or lngSourceHwnd = ucPacsHelper1.hwnd Then Exit Sub
     
     If lngSyncType = 0 And ucPacsHelper1.SelTabName <> "ͼ��" Then Exit Sub
     If lngSyncType = 1 And ucPacsHelper1.SelTabName <> "�ʾ�" Then Exit Sub
     If lngSyncType = 2 And ucPacsHelper1.SelTabName <> "��ʷ" Then Exit Sub
     If lngSyncType = 3 And ucPacsHelper1.SelTabName <> "����" Then Exit Sub
     
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
     
    'mblnIsLinkHelperΪfalse��ʾ�������ڣ�û�й��������ڵ�pacshelper����
    If mblnIsLinkHelper = False Then
        If Not mobjCapLinker Is Nothing Then mobjCapLinker.AddRepPacsHelper mobjStudyInfo.lngAdviceId, ucPacsHelper1
    End If
     
    lngReportID = GetReportId(mobjStudyInfo.lngAdviceId, mobjStudyInfo.blnMoved)
    
    If lngReportID = -1 Then
        ucReportEditor1.ResetContext
        ucReportEditor1.IsEditable = False
   
        'ʹ�÷�pacs����༭����д�ı���
        HintMsg "�˼����ʹ����������༭��������д�����ܴ򿪡�", "zlRefresh", vbExclamation
        'TASK:�ɵ���������Ԥ������
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
            strLayout = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "POPUPEDITOR" & mlngModuleId & ucReportEditor1.GetFaceKey, "")
            Call ucReportEditor1.SetLayout(strLayout)
        Else
            strLayout = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name, "MAINEDITOR" & mlngModuleId & ucReportEditor1.GetFaceKey, "")
            Call ucReportEditor1.SetLayout(strLayout)
        End If
        
        ucReportEditor1.tag = "1"
    End If
    
    If blnIsHistory Then
'        ucReportEditor1.IsEditable = False
        Call ucReportEditor1.ConfigFaceState(True, "��ʷ�鿴")
    End If
    
    '�жϽ����Ƿ��Զ���λ������༭��
    If mblnSetFocusWithReport = False Then
        Call ResetFocus(objFocus)
    Else
    
        If Not objFocus Is Nothing Then
            If TypeOf objFocus Is PatiIdentify Then
                '����ǲ���״̬�£��������Զ���λ����༭
                Exit Sub
            End If
        End If
        
        '����ж�Ӧ�ĵ���ʽ����༭���ڣ��Ҽ����ͬ���򲻶�λǶ��ʽ���ڱ༭��
        If mblnIsLinkHelper = True And IsSameReportWindow(mobjStudyInfo.lngAdviceId) Then
            Exit Sub
        End If
        
        Call ucReportEditor1.LocateEditBox
    End If
End Sub


Private Function IsSameReportWindow(ByVal lngAdviceId As Long) As Long
'�Ƿ������ͬ�Ķ���������д����
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
'���ý���ؼ�
On Error Resume Next
    Call objFocus.SetFocus
End Sub


Public Function PromptSave() As Boolean
'�򿪶�����д���洰��ʱ�����ô˷���������ʾ
    PromptSave = False
    
    If mobjStudyInfo Is Nothing Then Exit Function
    If mobjStudyInfo.lngAdviceId = 0 Then Exit Function
    
    PromptSave = ucReportEditor1.PromptSave(ucReportEditor1.AdviceId, ucReportEditor1.ReportID, True)
End Function


Private Sub InitReportSampleFormat(ByVal lngFileId As Long)
'��ʼ�����淶�ĸ�ʽ
On Error GoTo errH
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i  As Integer
    
    ReDim rptFormats(1) As rptFormat
    rptFormats(1).ID = 0
    rptFormats(1).strName = "��׼��ʽ"
    
    If lngFileId = 0 Then Exit Sub
    
    strSQL = "Select Id,���� From ��������Ŀ¼ Where �ļ�ID = [1] And ����= 0 And (ͨ�ü�=0 Or (ͨ�ü�=1 And ����ID=[2]) " & _
            " Or (ͨ�ü�=2 And ��ԱID= [3])) "
            
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngFileId, UserInfo.����ID, UserInfo.ID)
    If rsTemp.RecordCount <> 0 Then
        ReDim Preserve rptFormats(rsTemp.RecordCount + 1) As rptFormat
        For i = 1 To rsTemp.RecordCount
            rptFormats(i + 1).ID = rsTemp!ID
            rptFormats(i + 1).strName = rsTemp!����
            
            rsTemp.MoveNext
        Next i
    End If
    Exit Sub
errH:
    If HintError(err, "InitReportSampleFormat") = 1 Then Resume
End Sub



Private Function IsCustomFont(ByVal intFontSize As Integer) As Boolean
'���ܣ��ж��Ƿ�ʹ���Զ����ֺ�  ���� true-��
'���򣬲�����103523�����ظ�
    IsCustomFont = True
    
    If intFontSize = 0 Or intFontSize = 14 Or intFontSize = 16 Or intFontSize = 22 Or intFontSize = 28 Or intFontSize = 36 Or intFontSize = 42 Then
        IsCustomFont = False
    End If
    
End Function


Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngID As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngID, strCaption)
    End If
    
    CreateModuleMenu.ID = lngID '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
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
    '��Ӹ�ʽѡ�񵯳��˵������ģ�
    If CommandBar.Parent.ID = conMenu_PacsReport_SelFormat Then
        CommandBar.Controls.DeleteAll

        '����µĲ˵���
        For i = 1 To UBound(rptFormats)
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_SelFormat_Item, rptFormats(i).strName, i)
            cbrControlItem.Parameter = rptFormats(i).ID
        Next i
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_RepFormat Then '(��ӡ��ʽ)
        CommandBar.Controls.DeleteAll
        
        If Len(mstr������) <= 0 Then Exit Sub

        '����µĲ˵���
        strSQL = "Select a.���,b.���,b.˵�� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�Զ��屨���ʽ", mstr������)

        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_RepFormat_Item, rsTemp!��� & "-" & nvl(rsTemp!˵��))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = (InStr(mstrѡ�б����ʽ, cbrControlItem.Caption) <> 0)
            cbrControlItem.Parameter = rsTemp!���
            cbrControlItem.CloseSubMenuOnClick = False

            rsTemp.MoveNext
        Wend
    ElseIf CommandBar.Parent.ID = conMenu_PacsReport_VerifySign Then
        'ǩ����֤�ĵ����˵����г�������֤��ǩ���汾
        CommandBar.Controls.DeleteAll

        '����µ�ǩ����֤�˵�
        strSQL = "Select ��ʼ��,�����ı� as ǩ��ҽ�� From ���Ӳ������� Where �ļ�ID = [1] And �������� =8  Order By ��ʼ��"
        If ucReportEditor1.IsMoved Then
            strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ǩ���汾", ucReportEditor1.ReportID)

        While rsTemp.EOF = False
            Set cbrControlItem = CommandBar.Controls.Add(xtpControlButton, conMenu_PacsReport_VerifySign_Item, rsTemp!��ʼ�� & "-" & nvl(rsTemp!ǩ��ҽ��))
            cbrControlItem.Style = xtpButtonIconAndCaption
            cbrControlItem.Checked = False
            cbrControlItem.Parameter = rsTemp!��ʼ��
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
'����
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
        HintMsg "mobjCurPacsHelper ������Ч��", "ucReportEditor1_OnOutlineChange", infNormalErr
        Exit Sub
    End If
    
    Select Case lngSelOutline
        Case otDesc
            Call mobjCurPacsHelper.SyncOutline("����")
            
        Case otOpin
            Call mobjCurPacsHelper.SyncOutline("���")
            
        Case otAdvi
            Call mobjCurPacsHelper.SyncOutline("����")
        
        Case Else
            Call mobjCurPacsHelper.SyncOutline("")
            
    End Select
    
Exit Sub
errhandle:
    If HintError(err, "ucReportEditor1_OnOutlineChange", False) = 1 Then Resume
End Sub
 
