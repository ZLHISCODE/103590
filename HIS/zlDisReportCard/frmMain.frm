VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "�л����񹲺͹���Ⱦ�����濨"
   ClientHeight    =   9855
   ClientLeft      =   3270
   ClientTop       =   705
   ClientWidth     =   15105
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   15105
   Begin VB.TextBox txtFeedBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      Height          =   300
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   500
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtContent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   500
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   5000
   End
   Begin MSComctlLib.ProgressBar prgSaveData 
      Height          =   330
      Left            =   1365
      TabIndex        =   1
      Top             =   9540
      Visible         =   0   'False
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Max             =   44
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9480
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMain.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23733
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   930
      Top             =   375
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmMain.frx":115C
      Left            =   2730
      Top             =   570
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mfrmReport As Form      '�༭����
Attribute mfrmReport.VB_VarHelpID = -1
Private blnFirstActive As Boolean

Private mblnFeedbackReport As Boolean       '�����Ƿ��Ǵ����޵ı��棨��Ⱦ������վ���δͨ����
Private mlngFileID As Long                  '����ID
Private mStrFeedback As String              '����˵������
Private mblnǿ����д As Boolean              '��Ⱦ�����濨ǿ����д



Public Sub ShowMe(ByVal frmParent As Object, ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long, ByVal blnHand As Boolean)
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    
    mblnFeedbackReport = False
    mStrFeedback = ""
    Set mfrmReport = GetReport
    With mfrmReport.mclsReport
        .blnHaveStatus = True
    
        mlngFileID = lngFileId
        Call .InitReport(bytType, lngPatiID, lngPageID, bytFrom, bytBabyNo, lngDeptID, lngFileId)
        If frmParent.Name = "frmDiseaseStation" And InStr(frmParent.Caption, "��Ⱦ��������վ") > 0 Then
            gblnLock = True
        Else
            gblnLock = False
        End If
    
        If lngPatiID <> 0 Then
            If bytType = 1 Then
                mblnFeedbackReport = IsFeedbackReport(lngFileId)
                strSQL = "select t.���汾 from ���Ӳ�����¼ t where t.id=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", lngFileId)
                If rsTemp.RecordCount <> 0 Then
                    If Nvl(rsTemp!���汾) = 1 Then
                        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = False
                        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = True
                        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = False
                        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = False
                    Else
                        Call .CanWrite
                    End If
                End If
            Else
                Call .CanWrite
            End If
            Call .LoadData(bytType)
        End If
        
        '��Ⱦ������ϵͳ�Զ�ȡ�����
        If gblnLock = True Then
            Call cbrMain_Execute(cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel))
            cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Visible = False
        End If
        
        mblnǿ����д = ((Val(zlDatabase.GetPara("��Ⱦ�����濨ǿ����д", glngSys)) = 1) And (bytType = 0) And (blnHand = False And frmParent.Name <> "frmDiseaseStation"))

        Me.Show 1, frmParent
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Function IsFeedbackReport(ByVal lngFileId As Long) As Boolean
'���ܣ��жϸñ����Ƿ��Ǵ����޵ı���
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHand
    mStrFeedback = ""
    
    strSQL = "Select a.����״̬, b.�������� From �����걨��¼ A, �������淴�� B" & vbNewLine & _
             "Where a.�ļ�id = b.�ļ�id And a.�ļ�id = [1] And A.����״̬ = 4 And B.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From �������淴�� Where �ļ�id = [1])"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", lngFileId)
    If rsTemp.RecordCount > 0 Then
        mStrFeedback = rsTemp!�������� & ""
        IsFeedbackReport = True
    End If
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Private Function SavaProcessContent() As Boolean
'���ܣ��洢���ޱ���Ĵ���˵��
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strContent As String
    Dim strProcessDate As String
    Dim strPer As String
    
    On Error GoTo errHand
    strContent = txtContent.Text
    
    If Len(strContent) > 500 Then
        Call MsgBox("����˵������ܹ�����500���ַ���", vbInformation, gstrSysName)
        Exit Function
    End If
    strPer = "'" & UserInfo.���� & "'"
    strContent = "'" & strContent & "'"
    strProcessDate = "to_date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd HH24:MI:SS')"
    strSQL = "zl_�����걨��¼_update(" & CStr(mlngFileID) & " ,5,NULL,NULL,Null, " & strPer & "," & strProcessDate & "," & strContent & ")"

    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SavaProcessContent = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub InitSc()
'��ʼ���沼��
    Dim Pane1 As Pane
    
    On Error GoTo errHand
    
    With Me.dkpMain
        .SetCommandBars cbrMain
        .Options.HideClient = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
    End With
    
    dkpMain.DestroyAll
    
    Set Pane1 = dkpMain.CreatePane(1, 250, 250, DockLeftOf, Nothing)
    Pane1.Options = PaneNoCloseable + PaneNoCaption + PaneNoHideable + PaneNoFloatable
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    On Err GoTo errHand
    
    Select Case Control.ID
        Case conMenu_Manage_Exit
            Call Menu_Exit
        Case conMenu_Manage_Finish
            Call Menu_Finish
        Case conMenu_Manage_Cancel
            Call Menu_Cancel
             '�����ޱ����������д����˵�������Ƿ��ޱ����ȡ������������
            If mblnFeedbackReport Then
                txtContent.Locked = False
                txtContent.BackColor = &HFFFFFF
            Else
                Call mfrmReport.mclsReport.RelateFeedback(False)
            End If
        Case conMenu_Manage_Save
            Call Menu_Save

    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub Menu_Save()
'�ݴ�
    On Error GoTo errHand
    
    prgSaveData.Visible = True
    prgSaveData.Value = 0
    Call mfrmReport.mclsReport.ClearEnterInfo
    Call mfrmReport.mclsReport.SaveData(False)
    prgSaveData.Visible = False
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub Menu_Exit()
'�˳�
    Dim result As VbMsgBoxResult
    
    On Error GoTo errHand
    If mfrmReport.mclsReport.HaveChanged = True Then
        result = MsgBox("�Ƿ񱣴��޸����ݣ�", vbYesNoCancel + vbQuestion, gstrSysName)
        If result = vbYes Then
            If gblnLock Then
                Call Menu_Finish
            Else
                Call Menu_Save
            End If
            Unload Me
        ElseIf result = vbNo Then
            Unload Me
        Else
            Exit Sub
        End If
    Else
        Unload Me
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub Menu_Finish()
'���
    On Error GoTo errHand
    If CheckValidity = True Then
        '�����ޱ����Ҫ��д����˵��������
        If mblnFeedbackReport Then
            If Trim(txtContent.Text) = "" Then
                Call MsgBox("������д����˵����", vbInformation, gstrSysName)
                txtContent.SetFocus
                Exit Sub
            End If
            Call SavaProcessContent
        End If
        Call mfrmReport.mclsReport.SetEnterInfo
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = True
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = False
        prgSaveData.Visible = True
        prgSaveData.Value = 0
        Call mfrmReport.mclsReport.SaveData(True)
        prgSaveData.Visible = False
        '���Ƿ��ޱ���ͺͷ���������
        If (Not mblnFeedbackReport) And (Not gblnLock) Then
             Call mfrmReport.mclsReport.RelateFeedback(True)
        End If
        
		Me.Tag = 1
        Unload Me
    End If
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub
Private Sub Menu_Cancel()
'ȡ������
    Dim strSQL As String
    
    On Error GoTo errHand
    Call mfrmReport.mclsReport.CanWrite
    If Not gblnLock Then
        Call mfrmReport.mclsReport.ClearEnterInfo
    End If
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = True
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Cancel).Visible = False
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = True
    cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Visible = True

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Manage_Cancel
            If mblnFeedbackReport And cbrMain.FindControl(xtpControlButton, conMenu_Manage_Finish).Enabled = True Then
                txtContent.Locked = False
                txtContent.BackColor = &HFFFFFF
            End If
    End Select
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = mfrmReport.hWnd
    End If
End Sub

Private Sub Form_Activate()
    If blnFirstActive = True Then
        Call mfrmReport.mclsReport.SetMyFocus
        blnFirstActive = False
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo errHand
    blnFirstActive = True
    Me.WindowState = 2
    txtFeedBack.Text = mStrFeedback
    Call InitCommandBars

    If mblnǿ����д Then
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Exit).Visible = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Exit).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Enabled = False
        cbrMain.FindControl(xtpControlButton, conMenu_Manage_Save).Visible = False
    End If


    Call InitSc

    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Function CheckValidity() As Boolean
'���༭����ĺϷ���
    CheckValidity = mfrmReport.mclsReport.CheckValidity
End Function

Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Set Me.cbrMain.Icons = ZLCommFun.GetPubIcons
    
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.Visible = False
    
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagStretched
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Save, "�ݴ�(&S)", "��ʱ����", 3503, True)
    cbrControl.Style = xtpButtonIconAndCaption
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Finish, "���(&F)", "��ɱ༭", 804, False)
    cbrControl.Style = xtpButtonIconAndCaption
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Cancel, "ȡ�����(&C)", "ȡ�����", 3504, False)
    cbrControl.Style = xtpButtonIconAndCaption
    cbrControl.Visible = False
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_Manage_Exit, "�˳�(&E)", "�˳��༭", 191, True)
    cbrControl.Style = xtpButtonIconAndCaption

    If mblnFeedbackReport Then
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlLabel, 99999901, "��������:")
            cbrControl.Flags = xtpFlagRightAlign
            Set cbrCustom = .Add(xtpControlCustom, 99999902, "��������")
            cbrCustom.Handle = Me.txtFeedBack.hWnd
            cbrCustom.Flags = xtpFlagRightAlign
            Set cbrControl = .Add(xtpControlLabel, 99999903, "����˵��:")
            cbrControl.Flags = xtpFlagRightAlign
            Set cbrCustom = .Add(xtpControlCustom, 99999904, "����˵��")
            cbrCustom.Handle = Me.txtContent.hWnd
            cbrCustom.Flags = xtpFlagRightAlign
        End With
    End If
    
    With cbrMain.KeyBindings
        .Add FCONTROL, vbKeyS, conMenu_Manage_Save
        .Add FCONTROL, vbKeyF, conMenu_Manage_Finish
        .Add FCONTROL, vbKeyU, conMenu_Manage_Cancel
        .Add FCONTROL, vbKeyE, conMenu_Manage_Exit
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Private Function CreateModuleMenu(objMenuControl As CommandBarControls, _
    ByVal lngType As XTPControlType, ByVal lngId As Long, ByVal strCaption As String, _
    Optional strToolTip As String = "", Optional lngIconId As Long = 0, Optional blnStartGroup As Boolean = False, Optional ByVal lngIndex As Long = -1) As CommandBarControl
'������ģ���ڵĲ˵�
    
    On Error GoTo errHand
    
    If lngIndex >= 0 Then
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngId, strCaption, lngIndex)
    Else
        Set CreateModuleMenu = objMenuControl.Add(lngType, lngId, strCaption)
    End If

    CreateModuleMenu.ID = lngId '������ﲻָ��id�����ܽ���Щ�˵���ӵ��Ҽ��˵���
    
    If lngIconId <> 0 Then CreateModuleMenu.IconId = lngIconId
    If blnStartGroup Then CreateModuleMenu.BeginGroup = True
    If strToolTip <> "" Then CreateModuleMenu.ToolTipText = strToolTip
    
    CreateModuleMenu.Category = M_STR_MODULE_MENU_TAG

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Private Sub Form_Resize()
    On Error Resume Next
    prgSaveData.Width = Me.ScaleWidth - 1400
    prgSaveData.Left = 1400
    prgSaveData.Top = Me.ScaleTop + Me.ScaleHeight - prgSaveData.Height
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnǿ����д And Me.Tag <> "1" Then
        Call MsgBox("������ɴ�Ⱦ�����濨����д��", vbInformation, gstrSysName)
        Cancel = 1
        Exit Sub
    End If
    Unload mfrmReport
End Sub

Public Sub HaveSavedSQL()
    On Error Resume Next
    prgSaveData.Value = prgSaveData.Value + 1
    Err.Clear
End Sub

Private Sub txtContent_GotFocus()
    Me.txtContent.SelStart = 0: Me.txtContent.SelLength = 1000
    Call ZLCommFun.OpenIme(True)
End Sub

Private Sub txtContent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/<>", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtContent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(txtContent.hWnd, txtContent.Text, True, True)
End Sub

Private Sub txtFeedBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call zlCommFun.ShowTipInfo(txtFeedBack.hWnd, txtFeedBack.Text, True, True)
End Sub