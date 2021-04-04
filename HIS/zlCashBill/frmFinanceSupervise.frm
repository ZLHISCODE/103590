VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmFinanceSupervise 
   Caption         =   "�շѲ�����"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11730
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSupervise.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11730
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8055
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   635
      SimpleText      =   $"frmFinanceSupervise.frx":6852
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmFinanceSupervise.frx":6899
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13044
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "���˺�"
            TextSave        =   "���˺�"
            Object.ToolTipText     =   "��ǰ����Ա:���˺�"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   1605
      Left            =   -75
      TabIndex        =   1
      Top             =   1020
      Width           =   4290
      _Version        =   589884
      _ExtentX        =   7567
      _ExtentY        =   2831
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   -30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmFinanceSupervise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mfrmCollect As frmFinanceSupervisePersonList
Private mfrmHistory As frmFinaceSuperviseHistory
Private mfrmStandbyMoney As frmFinanceSuperviseStandbyMoenyList
Private mblnAllowZero As Boolean  '����������

Private Enum mPgIndex
    EM_PG_�տ��б� = 250101
    EM_PG_��ʷ�б� = 250102
    EM_PG_���ý��б� = 250103
End Enum

Private Sub initVar()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���
    '����:���˺�
    '����:2013-10-14 16:35:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dtDate As Date
    dtDate = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd")
    strSQL = "Select 1 From ��Ա�սɼ�¼ Where �Ǽ�ʱ��>=[1] And Rownum=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, dtDate)
    '���ʹ������:�շ�Ա�ɿ��������㣬��Ҫ����Ϊ��Щ�û�ʹ�õ��Ǳ����ӡ�ķ�ʽ���нɿ�,�������������
    '   �����һ���ڴ�����Ա�տ���ʼ�¼��������������ʹ��������ֱ�����θù���
    mblnAllowZero = rsTemp.EOF  '
    If Not rsTemp Is Nothing Then rsTemp.Close
    Set rsTemp = Nothing
End Sub


Public Sub zlShowFinanceSupervise(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal strPrivs As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շѲ����س������
    '���:frmMain-���õ�������
    '       lngModule-ģ���
    '       strPrivs-ģ��Ȩ�޴�
    '����:���˺�
    '����:2013-09-22 16:32:18
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
     
    If CheckDepend = False Then Exit Sub
    '��ʼ������
    Call InitFace
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show , frmMain
    End If
End Sub

Public Sub BHShowList(ByVal lngModule As Long, ByVal strPrivs As String, ByVal lngMain As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2013-10-17 18:17:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If CheckDepend = False Then Exit Sub
    '��ʼ������
    Call InitFace
    mlngModule = lngModule: mstrPrivs = strPrivs
    zlCommFun.ShowChildWindow Me.hWnd, lngMain
    Me.ZOrder 0
End Sub

Private Sub InitFace()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2013-09-03 14:43:09
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitPage
End Sub
Public Function zlDefCommandBars() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 18:21:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objComBar As CommandBarComboBox
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsThis.VisualTheme = xtpThemeOffice2003
    With cbsThis.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.ID = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        If mblnAllowZero And zlStr.IsHavePrivs(mstrPrivs, "���ʹ���") Then '�������ʹ������ʱ��������ִ�иù���
            Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain_Zero, "���ʹ���(&C)"): mcbrControl.BeginGroup = True
        End If
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.ID = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutOut, "���ű��ý�(&L)")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_OnWork, "�ϸڱ��ý�")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutIn, "�ջر��ý�(&H)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3017
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Manual, "�ֹ��տ�(&M)")
        mcbrControl.IconId = 6820
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_RollingCurtain, "�����տ�(&S)")
        mcbrControl.IconId = 3588
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Cancel, "�տ�����(&C)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3589
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Other, "������Ա�տ�(&O)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮(&E)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Personnel_Group, "��Ա����(&G)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChargeBook_Reprint, "�ش��տ��վ�(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_DrawBook_Reprint, "�ش��ý����õ�(&D)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.ID = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��(&G)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��(&M)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�(&L)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����(&D)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "�鿴��ϸ����(&V)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 2322
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.ID = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_Collect_Cancel
        .Add FCONTROL, Asc("R"), conMenu_Edit_ChargeBook_Reprint
        .Add FCONTROL, Asc("M"), conMenu_Edit_Collect_Manual
        .Add FCONTROL, Asc("O"), conMenu_Edit_Collect_Other
        .Add FCONTROL, Asc("T"), conMenu_View_Detail
        .Add 0, VK_F2, conMenu_Edit_Collect_RollingCurtain
        .Add 0, VK_F6, conMenu_Edit_CheckCash
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
    End With
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With
    
    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.ContextMenuPresent = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Manual, "�ֹ��տ�"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 6820
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_RollingCurtain, "�����տ�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Other, "������Ա�տ�(&O)"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 228
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Collect_Cancel, "�տ�����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutOut, "���ű��ý�")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_OnWork, "�ϸڱ��ý�")
        mcbrControl.IconId = 3011
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_StandbyMoeny_PutIn, "�ջر��ý�"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3017
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CheckCash, "�ֽ�㳮"): mcbrControl.BeginGroup = True
        mcbrControl.IconId = 3590
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Detail, "��ѯ��ϸ")
         mcbrControl.IconId = 2322
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")

    End With
    For Each mcbrControl In mcbrToolBar.Controls
        If mcbrControl.ID <> conMenu_COMBOX_INTERFACE Then
          mcbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
     zlDefCommandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    
    If zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�") _
        Or zlStr.IsHavePrivs(mstrPrivs, "�������տ�") _
        Or zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�") Then
        If mfrmCollect Is Nothing Then
            Set mfrmCollect = New frmFinanceSupervisePersonList
            Load mfrmCollect
        End If
        '��ʼ������
        Call mfrmCollect.zlInitVar(mlngModule, mstrPrivs, cbsThis)
    End If
    
    If mfrmHistory Is Nothing Then
        Set mfrmHistory = New frmFinaceSuperviseHistory
        Load mfrmHistory
    End If
    Call mfrmHistory.zlInitVar(mlngModule, mstrPrivs)
 
    If mfrmStandbyMoney Is Nothing Then
        Set mfrmStandbyMoney = New frmFinanceSuperviseStandbyMoenyList
        Load mfrmStandbyMoney
    End If
    Call mfrmStandbyMoney.zlInitVar(mlngModule, mstrPrivs)
    If zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�") _
      Or zlStr.IsHavePrivs(mstrPrivs, "�������տ�") _
      Or zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�") Then
        Set objItem = tbPage.InsertItem(EM_PG_�տ��б�, "�տ�", mfrmCollect.hWnd, 0)
        objItem.Tag = EM_PG_�տ��б�
    End If
    Set objItem = tbPage.InsertItem(EM_PG_��ʷ�б�, "��ʷ�տ���Ϣ", mfrmHistory.hWnd, 0)
    objItem.Tag = EM_PG_��ʷ�б�
    Set objItem = tbPage.InsertItem(EM_PG_���ý��б�, "���ý��б� ", mfrmStandbyMoney.hWnd, 0)
    objItem.Tag = EM_PG_���ý��б�
     With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

 
Private Sub cbsThis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With tbPage
        tbPage.Left = Left
        tbPage.Top = Top
        tbPage.Width = Right - Left
        tbPage.Height = Bottom - Top
    End With
End Sub

Private Sub Form_Activate()
    stbThis.Panels(3).Text = UserInfo.����
End Sub

Private Sub Form_Load()
    Call initVar
    RestoreWinState Me, App.ProductName
    Call zlDefCommandBars
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
End Sub
 

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mfrmCollect Is Nothing Then Unload mfrmCollect
    If Not mfrmHistory Is Nothing Then Unload mfrmHistory
    If Not mfrmStandbyMoney Is Nothing Then Unload mfrmStandbyMoney
    Set mfrmCollect = Nothing
    Set mfrmHistory = Nothing
    Set mfrmStandbyMoney = Nothing
End Sub
 
Private Sub ParameterSet()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2013-09-12 15:31:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If frmFinanceSuperviseParaSet.ShowMe(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub SaveCollect(Optional ByVal blnCustomCollect As Boolean = False)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʴ���
    '���:blnCustomCollect-true-�ֹ��տ�;false-�����տ�;
    '����:���˺�
    '����:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String
    If Val(tbPage.Selected.Tag) <> EM_PG_�տ��б� Then Exit Sub
    
    If Not (zlStr.IsHavePrivs(mstrPrivs, "�շ�Ա�տ�") _
        Or zlStr.IsHavePrivs(mstrPrivs, "�������տ�") _
        Or zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�")) Then Exit Sub
    
    If Not mfrmCollect.zlRollingCurtainCollect(Me, blnCustomCollect) Then Exit Sub
    Call zlRefresh
End Sub
Private Sub SaveCollectCancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�տ����ϴ���
    '����:���˺�
    '����:2013-09-12 15:34:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, blnDel As Boolean
    If Val(tbPage.Selected.Tag) <> EM_PG_��ʷ�б� Then Exit Sub
    If mfrmHistory.CancelData() Then Exit Sub
End Sub
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
     Select Case Control.ID
        Case conMenu_File_Exit: Unload Me: '�˳�(&X)
        Case conMenu_File_PrintSet: Call zlPrintSet '��ӡ����
        Case conMenu_File_Preview: Call zlPrintRpt(2)  'Ԥ��(&V)
        Case conMenu_File_Print: Call zlPrintRpt(1) '��ӡ(&P)
        Case conMenu_File_Excel: Call zlPrintRpt(3)  '�����&Excel��
        Case conMenu_File_Parameter: Call ParameterSet '��������
        Case conMenu_Edit_RollingCurtain_Zero: ExcuteRollingCurtainZero  '���ʹ���(&C)"
        Case conMenu_Edit_StandbyMoeny_PutOut: ExcutePutOutStandbyMoeny '���ű��ý�
        Case conMenu_Edit_StandbyMoeny_OnWork: ExcuteOnWorkStandbyMoeny
        Case conMenu_Edit_StandbyMoeny_PutIn: ExcutePutINStandbyMoeny '�ջر��ý�
        Case conMenu_Edit_Collect_Manual: Call SaveCollect(True)   '�ֹ��տ�
        Case conMenu_Edit_Collect_RollingCurtain: Call SaveCollect '�����տ�
        Case conMenu_Edit_Collect_Cancel: Call SaveCollectCancel '�տ�����
        Case conMenu_Edit_Collect_Other: Call SaveCollect  '������Ա�տ�
        Case conMenu_Edit_CheckCash: Call CheckCash '�ֽ�㳮(&E)
        Case conMenu_Edit_Personnel_Group: Call ExcuteSplitGroup '��Ա����
        Case conMenu_Edit_ChargeBook_Reprint:  Call RePrintBill(0) '�ش��տ��վ�(&R)
        Case conMenu_Edit_DrawBook_Reprint:  Call RePrintBill(1) '�ش��ý����õ�(&R)
        Case conMenu_View_Detail: Call ShowChargeList '�鿴��ϸ����(&V)
        Case conMenu_View_Refresh: zlRefresh 'ˢ��(&R)
        Case conMenu_View_LargeICO: SetPersonListShow (0) '��ͼ��(&G)
        Case conMenu_View_MinICO: SetPersonListShow (1)  'Сͼ��(&M)
        Case conMenu_View_ListICO: SetPersonListShow (2)  '�б�(&L)
        Case conMenu_View_DetailsICO:: SetPersonListShow (3)  '��ϸ����(&D)
        Case conMenu_View_StatusBar '״̬��(&S)
            stbThis.Visible = Not stbThis.Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Button
            cbsThis(2).Visible = Not cbsThis(2).Visible
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each mcbrControl In cbsThis(2).Controls
                mcbrControl.Style = IIf(mcbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
            cbsThis.RecalcLayout
        Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                'ִ�з�������ǰģ��ı���
                Call CallCustomRpt(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
            End If
        End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHavePrivs As Boolean, lngPage As Long
    Dim intView As Integer, blnEanbled As Boolean
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    If Not tbPage.Selected Is Nothing Then
        lngPage = Val(tbPage.Selected.Tag)
    End If
    Select Case Control.ID
    Case conMenu_Edit_RollingCurtain_Zero '���ʹ���
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "���ʹ���") And mblnAllowZero
        blnEanbled = lngPage = EM_PG_�տ��б�
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_PutOut '���ű��ý�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "���ű��ý�")
        blnEanbled = lngPage = EM_PG_���ý��б�
        Control.Visible = blnHavePrivs And blnEanbled:
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_OnWork '�ϸڱ��ý�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ϸڱ��ý�")
        blnEanbled = lngPage = EM_PG_���ý��б�
        Control.Visible = blnHavePrivs And blnEanbled:
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_StandbyMoeny_PutIn '�ջر��ý�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ջر��ý�")
        blnEanbled = lngPage = EM_PG_���ý��б�
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmStandbyMoney.IsAllowCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_Manual '�ֹ��տ�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ֹ��տ�")
        blnEanbled = lngPage = EM_PG_�տ��б�
        If blnEanbled Then
              blnEanbled = blnEanbled And mfrmCollect.IsAllowCustomCollect
        End If
       Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_RollingCurtain '�����տ�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�����տ�")
        blnEanbled = lngPage = EM_PG_�տ��б�
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmCollect.IsAllowCollect
        Control.Enabled = blnEanbled
    Case conMenu_Edit_Collect_Other  '������Ա�տ�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "������Ա�տ�")
        blnEanbled = lngPage = EM_PG_�տ��б�
        If blnEanbled Then blnEanbled = mfrmCollect.IsAllowOtherCollect
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_Collect_Cancel '�տ�����
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�տ�����")
        blnEanbled = lngPage = EM_PG_��ʷ�б�
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmHistory.IsAllowCollectCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_CheckCash ' "�ֽ�㳮(&E)")
        blnEanbled = lngPage <> EM_PG_���ý��б�
        Control.Visible = blnEanbled
        Control.Enabled = blnEanbled
    Case conMenu_Edit_Personnel_Group '��Ա����
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "��Ա����")
        Control.Visible = blnHavePrivs: Control.Enabled = blnHavePrivs
    Case conMenu_Edit_ChargeBook_Reprint ' "�ش��տ��վ�(&R)")
        blnEanbled = lngPage = EM_PG_��ʷ�б�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ش��տ��վ�") And zlStr.IsHavePrivs(mstrPrivs, "�տ��վݴ�ӡ")
        Control.Visible = blnHavePrivs And blnEanbled
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_Edit_DrawBook_Reprint ' "�ش��ý����õ�
        blnEanbled = lngPage = EM_PG_���ý��б�
        blnHavePrivs = zlStr.IsHavePrivs(mstrPrivs, "�ش��ý����õ�") _
            And zlStr.IsHavePrivs(mstrPrivs, "���ý����õ���ӡ")
        Control.Visible = blnHavePrivs And blnEanbled
        If blnEanbled Then blnEanbled = mfrmStandbyMoney.IsAllowCancel
        Control.Enabled = blnHavePrivs And blnEanbled
    Case conMenu_View_Detail '�鿴��ϸ����
        blnEanbled = lngPage <> EM_PG_���ý��б�
        Control.Visible = blnEanbled
        If blnEanbled Then
            If lngPage <> EM_PG_��ʷ�б� Then
                blnEanbled = mfrmCollect.IsAllowViewChargeList
            Else
                blnEanbled = mfrmHistory.IsAllowViewChargeList
            End If
        End If
        Control.Enabled = blnEanbled
    Case conMenu_View_LargeICO  '��ͼ��(&G)
        Control.Visible = lngPage = EM_PG_�տ��б�
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 0
        End If
    Case conMenu_View_MinICO  'Сͼ��(&M)
        Control.Visible = lngPage = EM_PG_�տ��б�
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 1
        End If
    Case conMenu_View_ListICO  '�б�(&L)
        Control.Visible = lngPage = EM_PG_�տ��б�
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 2
        End If
    Case conMenu_View_DetailsICO  '��ϸ����(&D)
        Control.Visible = lngPage = EM_PG_�տ��б�
        If Control.Visible Then
            intView = mfrmCollect.zlPersonListShowMode
            Control.Checked = intView = 3
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    End Select
End Sub
Private Function CheckDepend() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2013-09-04 17:10:03
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    CheckDepend = False
    On Error GoTo errHandle
    Set rsTemp = Get���㷽ʽ
    rsTemp.Filter = "����=1"
    If rsTemp.EOF Then
        rsTemp.Filter = 0
        ShowMsgbox "���㷽ʽ�в�����һ�������ֽ����ʵĽ��㷽ʽ,���ڽ��㷽ʽ����������!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Filter = 0
    rsTemp.Close
    If UserInfo.���� = "" Then
        MsgBox "��ǰ��¼�û�δָ����Ӧ����Ա,����ʹ�ñ����ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub zlPrintRpt(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����б�
    '���:bytMode=1-��ӡ,2-Ԥ��,3-�����Excel
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-09-13 10:23:30
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then
        '��ӡ������Ϣ
        Call mfrmCollect.zlPrint(bytMode)
        Exit Sub
    End If
    If Val(tbPage.Selected.Tag) = EM_PG_���ý��б� Then
        mfrmStandbyMoney.zlPrint (bytMode): Exit Sub
    End If
    Call mfrmHistory.zlPrint(bytMode)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RePrintBill(ByVal bytRePrintType As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ش򵥾�
    '���:bytRePrintType-0-�վ�;1-���ý����õ�
    '����:���˺�
    '����:2013-09-13 16:00:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then Exit Sub
    If Val(tbPage.Selected.Tag) = EM_PG_���ý��б� Then
        mfrmStandbyMoney.RePrintBill: Exit Sub
    End If
    Call mfrmHistory.RePrintBill
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub CheckCash()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ֽ�㳮
    '����:���˺�
    '����:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then
        dblMoney = mfrmCollect.GetCashMoney
    End If
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub

Private Sub zlRefresh()
    '���½�������ˢ��
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then
        Call mfrmCollect.zlRefresh
    ElseIf Val(tbPage.Selected.Tag) = EM_PG_���ý��б� Then
         Call mfrmStandbyMoney.zlRefresh: Exit Sub
    Else
        Call mfrmHistory.zlRefresh
    End If
End Sub

Private Sub ShowChargeList()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��ϸ�տ�����
    '����:���˺�
    '����:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
     If Val(tbPage.Selected.Tag) = EM_PG_���ý��б� Then Exit Sub
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then
         Call mfrmCollect.ShowChargeList(Me)
         Exit Sub
    End If
    '��ʷ������ʾ
    Call mfrmHistory.ShowChargeList(Me)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CallCustomRpt(ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����Զ��屨��
    '���:lngSys-ϵͳ��
    '        strRptCode-������
    '����:���˺�
    '����:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Val(tbPage.Selected.Tag) = EM_PG_�տ��б� Then
         Call mfrmCollect.CallCustomRpt(Me, lngSys, strRptCode)
         Exit Sub
    End If
    If Val(tbPage.Selected.Tag) = EM_PG_���ý��б� Then
         Call mfrmStandbyMoney.CallCustomRpt(Me, lngSys, strRptCode)
    End If
    '��ʷ������ʾ
    Call mfrmHistory.CallCustomRpt(Me, lngSys, strRptCode)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub SetPersonListShow(ByVal intICOType As Integer)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����б����ʾ��ʽ
    '���:intType-ͼ������(0-��ͼ��;1-Сͼ��;2-�б�;3-��ϸ����)
    '����:���˺�
    '����:2013-09-27 15:30:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Val(tbPage.Selected.Tag) <> EM_PG_�տ��б� Then Exit Sub
    mfrmCollect.zlPersonListShowMode = intICOType
End Sub

Private Sub ExcuteOnWorkStandbyMoeny()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з����ϸڱ��ý����
    '����:������
    '����:2013-12-4
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.zlPayOnWorkMoney(Me) = False Then Exit Sub
End Sub

 Private Sub ExcutePutOutStandbyMoeny()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�з��ű��ý����
    '����:���˺�
    '����:2013-10-12 16:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.zlPayStandbyMoney(Me) = False Then Exit Sub
 End Sub
 Private Sub ExcutePutINStandbyMoeny()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���ջط��ű��ý����
    '����:���˺�
    '����:2013-10-12 16:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mfrmStandbyMoney Is Nothing Then Exit Sub
    If mfrmStandbyMoney.CancelStandbyMoney() = False Then Exit Sub
 End Sub
Private Sub ExcuteRollingCurtainZero()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ʹ��㴦��
    '����:���˺�
    '����:2013-10-14 11:49:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim str�շ�Ա As String, blnStrans As Boolean, intStep As Integer, intCount As Integer
    Dim strRand As String, strTemp As String, lngTop As Long, lngLeft As Long
    Dim i As Long, intTemp As Integer
    On Error GoTo errHandle
    '��ʾ����
    If zlStr.IsHavePrivs(mstrPrivs, "���ʹ���") = False Then Exit Sub
    Randomize
    For i = 1 To 3
       intTemp = Asc("A") + Int(Rnd * 10)
        If intTemp > Asc("Z") Then intTemp = Asc("A")
       strRand = strRand & Chr(intTemp)
    Next
    
    lngLeft = Me.Left + 2500: lngTop = Me.Top + 2500
    strTemp = InputBox("  ���ʹ������������������շ���Ա�Ľɿ�����, " & _
                          "����㲻���ù������ã��벻Ҫʹ�øù��ܡ� " & _
                          "�����ȷ��Ҫ������ʹ���,�����������ַ�:" & vbCrLf & " " & vbCrLf & " " & _
                          "" & strRand, "����", "", lngLeft, lngTop)
    If strTemp = "" Then Exit Sub
    If UCase(strTemp) <> UCase(strRand) Then
         MsgBox "�������,�������������!", vbInformation + vbOKOnly, gstrSysName
         Exit Sub
    End If
    frmWait.OpenWait Me, "���˹��㴦��", True
    frmWait.WaitInfo = "������ȡ����..."
    strSQL = "Select Distinct �տ�Ա From ��Ա�ɿ���� Where ����=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.EOF Then
        frmWait.CloseWait
        MsgBox "û��Ҫ���ʹ�������ݴ�������Ҫ���ʹ���", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Sub
    End If
    str�շ�Ա = ""
    gcnOracle.BeginTrans
    blnStrans = True
    intStep = 1: intCount = rsTemp.RecordCount
    With rsTemp
        Do While Not .EOF
            frmWait.WaitInfo = "�������[" & Nvl(!�տ�Ա) & "�Ľɿ�����..."
            If zlCommFun.ActualLen(str�շ�Ա & "," & Nvl(!�տ�Ա)) > 4000 Then
                str�շ�Ա = Mid(str�շ�Ա, 2)
                ' Zl_���ʹ����¼_Insert
                strSQL = "Zl_���ʹ����¼_Insert("
                '  �Ǽ���_In   In ��Ա�սɼ�¼.�Ǽ���%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  �Ǽ�ʱ��_In In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
                strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
                '  �շ�Ա_In   In Varchar2 := Null
                '  �շ�Ա_in-ָ�����շ�Ա,Ϊ��ʱ,Ϊ�����շ�Ա;�ǿ�ʱ,Ϊָ�����շ�Ա(����Ϊ���,����ö��ŷָ�
                strSQL = strSQL & "'" & str�շ�Ա & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            str�շ�Ա = str�շ�Ա & "," & Nvl(!�տ�Ա)
            frmWait.pgb.Value = intStep \ intCount
            intStep = intStep + 1
            .MoveNext
        Loop
    End With
    If intCount = 0 Then intCount = 1
    frmWait.pgb.Value = intStep \ intCount
    strDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If str�շ�Ա <> "" Then
        str�շ�Ա = Mid(str�շ�Ա, 2)
        ' Zl_���ʹ����¼_Insert
        strSQL = "Zl_���ʹ����¼_Insert("
        '  �Ǽ���_In   In ��Ա�սɼ�¼.�Ǽ���%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  �Ǽ�ʱ��_In In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
        strSQL = strSQL & "to_date('" & strDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  �շ�Ա_In   In Varchar2 := Null
        '  �շ�Ա_in-ָ�����շ�Ա,Ϊ��ʱ,Ϊ�����շ�Ա;�ǿ�ʱ,Ϊָ�����շ�Ա(����Ϊ���,����ö��ŷָ�
        strSQL = strSQL & "'" & str�շ�Ա & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans: blnStrans = False
    frmWait.CloseWait
    
    MsgBox "���ʹ�������ɹ�!", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    frmWait.CloseWait
    If blnStrans Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub ExcuteSplitGroup()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ�в������
    '����:���˺�
    '����:2013-10-15 15:22:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If frmGroupAndPesons.ShowGroups(Me, mlngModule, mstrPrivs) = False Then Exit Sub
    '���¼��ز�����
    Call mfrmCollect.zlRefresh
End Sub
