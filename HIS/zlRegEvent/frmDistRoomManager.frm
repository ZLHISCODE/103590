VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmDistRoomManager 
   Caption         =   "����������"
   ClientHeight    =   10890
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15105
   Icon            =   "frmDistRoomManager.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10890
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picTabPage 
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   270
      ScaleHeight     =   3945
      ScaleWidth      =   7665
      TabIndex        =   4
      Top             =   1425
      Width           =   7665
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   30
         TabIndex        =   5
         Top             =   75
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1500
      ScaleHeight     =   375
      ScaleWidth      =   3810
      TabIndex        =   1
      Top             =   0
      Width           =   3810
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   -45
         TabIndex        =   3
         Top             =   15
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483634
      End
      Begin VB.TextBox txtValue 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   600
         TabIndex        =   2
         ToolTipText     =   "��λF3"
         Top             =   30
         Width           =   3165
      End
   End
   Begin VB.Timer tmrBrush 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   1620
      Top             =   435
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   10530
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDistRoomManager.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21564
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   60
      Top             =   -15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmDistRoomManager.frx":115E
      Left            =   375
      Top             =   345
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmDistRoomManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModul As Long, mstrFindKey As String
Private mbytViewScrop(0 To 3) As Byte  '0-��ʾ�ѷ��ﲡ��;1-��ʾ�ѽ��ﲡ��;2-��ʾ����ɲ���;3-��ʾ�����ﲡ��
Private mblnCard As Boolean     '�Ƿ�ˢ��
Private mobjFindKey As CommandBarPopup
Private WithEvents mfrmTriageMgr  As frmTriageManager
Attribute mfrmTriageMgr.VB_VarHelpID = -1
Private WithEvents mobjQueue As zlQueueManage.clsQueueManage
Attribute mobjQueue.VB_VarHelpID = -1
Private mstrQueuePrivs As String '�Ŷӽк�����ģ��Ȩ��
Private mlngTimerState As Boolean        '��ʱ���timer״̬�ı���
Private mblnȱʡ���� As Boolean
Private Enum pg_Page
    pg_����ҳ = 1
    pg_�Ŷ�ҳ = 2
End Enum
Private Type ty_Para
    str������� As String
    int������Ч���� As Integer
    byt�Ŷӽк�ģʽ As Byte '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    byt����վ�� As Byte   '0-�������̨�������;1-����ҽ����������
    bln������� As Boolean
    blnAutoRefresh  As Boolean
    strcurQueueName As String '��ǰ��������
    lngcurQueueҵ��ID As Long     '��ǰ����ҵ��ID
    str�ٴ����� As String
    byt��������ʽ As Byte  '���ﲡ�˵�����ʽ,0-���ұ���,����,���ݺ�;1-���ұ���,����,�Һ�ʱ��;
    bln��Һ�ģʽ As Boolean '�Ƿ����ģʽ,���̣�ֱ���ڷ���̨ȡ�ţ�Ȼ���ڽ���ʱ���������۵�
End Type
Private mcllFilter As Variant
Private mcbrToolControl As CommandBarControl
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mTy_Para As ty_Para

'-----------------------------------------------------------------------------------
'��Ϣ��ر���
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1
Private mstrRegistIdsed As String '�Ѿ�ˢ�µĹҺ�ID,�ö��ŷ���
Private mblnExistNewMsg As Boolean    '�Ƿ��������Ϣ
'-----------------------------------------------------------------------------------
'���㿨���
Private mcllBrushCard As Collection
Private mstrCaption As String
Private mintFindType As Integer

Private Type ty_Square
    lngȱʡ�����ID As Long
    lng�����ID  As Long
    bln�������� As Boolean
    intҽ�ƿ����� As Long
End Type

Private mty_Square As ty_Square

Private Type ty_Queue
    str�������� As String
    lngҵ��ID As Long
    lng����ID As Long
    str�ŶӺ��� As String
End Type

Private Const conPane_OfferWin = 1  'ȡ�Ŵ���
Private Const conPane_Pages = 2 '��ҳ
Private mty_Queue As ty_Queue
Private WithEvents mfrmOferWin As frmTriageRoomRegNum   'ȡ�Ŵ���
Attribute mfrmOferWin.VB_VarHelpID = -1
Private mblnUnload As Boolean
Private mblnFirst As Boolean
Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���򻮷�
    '����:���˺�
    '����:2018-01-09 16:25:28
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim sngWidth As Single, lngHeight As Long, strReg As String, panThis As Pane
    Dim panTop As Pane
    If mTy_Para.bln��Һ�ģʽ Then
        Set mfrmOferWin = New frmTriageRoomRegNum
        If mfrmOferWin.zlInitVar(Me, mTy_Para.str�������, mlngModul, mstrPrivs, gobjSquare.objSquareCard, gobjRegist) = False Then mblnUnload = True: Exit Sub
        lngHeight = 1035 \ Screen.TwipsPerPixelY    '�߶ȹ̶�
        
        Set panTop = dkpMan.CreatePane(conPane_OfferWin, 200, lngHeight, DockTopOf, Nothing)
        panTop.Title = "ȡ����Ϣ": panTop.Tag = conPane_OfferWin
        panTop.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        panTop.Handle = mfrmOferWin.Hwnd
        panTop.MaxTrackSize.Height = lngHeight
        panTop.MinTrackSize.Height = lngHeight
        
        Set panThis = dkpMan.CreatePane(conPane_Pages, 250, 580, DockBottomOf, panTop)
    Else
      Set panThis = dkpMan.CreatePane(conPane_Pages, 250, 580, DockLeftOf, Nothing)
    End If
    panThis.Title = "������Ϣ"
    panThis.Tag = conPane_Pages
    panThis.Handle = picTabPage.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.SetCommandBars cbsThis
    dkpMan.Options.HideClient = True
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_OfferWin   'ȡ����Ϣ
        Item.Handle = mfrmOferWin.Hwnd
    Case conPane_Pages  '������Ϣ
        Item.Handle = picTabPage.Hwnd
    End Select
End Sub

'-----------------------------------------------------------------------------------
Private Sub ClearMenuItem()
    'ɾ�����ڵĹ������������˵���
    Dim lngCount As Long
    For lngCount = cbsThis.ActiveMenuBar.Controls.Count To 1 Step -1
        cbsThis.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsThis.Count To 2 Step -1
        cbsThis(lngCount).Delete
    Next
End Sub


Public Function zlDefCommandBars() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ���˵���������
    '���أ����óɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-01 11:04:33
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl, cbrCustom As CommandBarControlCustom
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar, i As Long, strKey As String
    
    Err = 0: On Error GoTo Errhand:
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
    
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.EnableCustomization False
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BillPrint, "�ش��Ŷӵ�(&R)"): cbrControl.BeginGroup = True
         '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BarcodePrint, "�����ӡ(&B)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "����ǩ��(&Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ��ǩ��(&X)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Triage, "����(&M)"): cbrControl.BeginGroup = True 'Ctrl+T
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_ChangeNum, "����(&C)") 'CTRL+M
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_Leave, "���˲�����(&L)")
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_Wait, "���˴���(&W)")
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalize, "����(&H)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalizeCancel, "ȡ������(&E)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "��ɾ���(&O)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Redo, "ȡ�����(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "������Ϣ(&I)"): cbrControl.BeginGroup = True 'Ctrl+I
        '73743:���ϴ�,2014-7-21,���˻�����Ϣ����
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPatiBaseInfo, "���˻�����Ϣ����(&D)")
    End With
 
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "��������(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_View_TriagePati, "��ʾ�ѷ��ﲡ��(&1)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conmenu_View_AdmissionsPati, "��ʾ�ѽ��ﲡ��(&2)")
        Set cbrControl = .Add(xtpControlButton, conmenu_View_OverPati, "��ʾ����ɲ���(&3)")
        Set cbrControl = .Add(xtpControlButton, conmenu_View_Leave, "��ʾ�����ﲡ��(&4)")
        
        Set cbrControl = .Add(xtpControlButton, conmenu_View_AutoRefresh, "�Զ�ˢ��(&A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    Set cbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '���˵��Ҳ�Ĳ���
    Set cbrCustom = cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Find, "")
    cbrCustom.Handle = picSearch.Hwnd
    cbrCustom.flags = xtpFlagRightAlign
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("T"), conMenu_Edit_Triage    '����
        .Add FCONTROL, Asc("M"), conmenu_Edit_ChangeNum '����
        .Add FCONTROL, Asc("I"), conMenu_Edit_ModiyPati     '������Ϣ
        .Add FCONTROL, Asc("F"), conMenu_View_Filter     '��������
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F3, conMenu_View_Find
    End With
    
    
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    Dim blnAddTools As Boolean
    
    blnAddTools = False
    If tbPage.Selected Is Nothing Then
        blnAddTools = True
    Else
        blnAddTools = Not (tbPage.Selected.Tag = pg_�Ŷ�ҳ And mTy_Para.bln��Һ�ģʽ)
    End If
    
    If blnAddTools Then
        '-----------------------------------------------------
        '����������
        Set cbrToolBar = cbsThis.Add("������", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.ContextMenuPresent = False
        cbrToolBar.EnableDocking xtpFlagStretched
        
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Plan, "����ǩ��"): cbrControl.BeginGroup = True
           'Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Logout, "ȡ��ǩ��")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Triage, "����"): cbrControl.BeginGroup = True
            Set mcbrToolControl = cbrControl
            Set cbrControl = .Add(xtpControlButton, conmenu_Edit_BackHospitalize, "���˻���"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Finish, "��ɾ���"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "����")
            Set cbrControl = .Add(xtpControlButton, conMenu_View_Filter, "����"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        End With
    End If
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
    
    If Not gobjRegist Is Nothing And mTy_Para.bln��Һ�ģʽ = False Then gobjRegist.zlDefCommandBars Me, cbsThis, True, , mcbrToolControl
    If Not cbrToolBar Is Nothing Then
        For Each cbrControl In cbrToolBar.Controls
            cbrControl.Style = xtpButtonIconAndCaption
        Next
    End If
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function zlGetDept(ByVal str������� As String) As String
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡ������Ϣ
    '����:������ϢIDs:��:123;234;24
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:40:14
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strDeptIds As String, rsTemp As ADODB.Recordset
    On Error GoTo Hd
    Set rsTemp = GetDepartments("'�ٴ�'", "1,3", InStr(mstrPrivs, "���п���") = 0)
    
    With rsTemp
        strDeptIds = ""
        Do While Not .EOF
            If InStr("," & str������� & ",", "," & Nvl(rsTemp!ID) & ",") > 0 Or str������� = "" Then
                strDeptIds = strDeptIds & "," & Nvl(rsTemp!ID)
            End If
            .MoveNext
        Loop
    End With
    If strDeptIds <> "" Then zlGetDept = Mid(strDeptIds, 2)
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Private Sub zlRefreshQueueData()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����»�ȡ��������
    '���ƣ����˺�
    '���ڣ�2010-06-02 17:53:32
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, rsTemp As ADODB.Recordset, strSQL As String
    Dim strTemp As String
    Dim strQueue() As String, i As Long
    If mobjQueue Is Nothing Or mTy_Para.byt�Ŷӽк�ģʽ = 0 Then Exit Sub
    If Not (InStr(mstrQueuePrivs, ";����;") > 0) Then Exit Sub
    
    strTemp = IIf(mTy_Para.str������� = "", mTy_Para.str�ٴ�����, mTy_Para.str�������)
    varData = Split(strTemp, ",")
    i = UBound(varData) + 1
    ReDim Preserve strQueue(1 To i) As String
    For i = 0 To UBound(varData)
        strQueue(i + 1) = varData(i)
    Next
    '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    'zlRefresh(cnOracle As ADODB.Connection, str��������() As String, strCurrent�������� As String, lngCurrentWorkID As Long) As Long
    '����:����ˢ��ָ��ҽ��id�ı������ݣ�����������ṩ�༭����
    '����:  lngOrderId-ҽ��id;
    '����:�ɹ�����0,���򷵻ش������
    Call mobjQueue.zlRefresh(strQueue, mTy_Para.strcurQueueName, mTy_Para.lngcurQueueҵ��ID)
End Sub


Private Sub InitVar(Optional blnPatiSet As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ����ر����Ͳ���
    '���:
    '���ƣ����˺�
    '���ڣ�2010-06-01 16:25:23
    '------------------------------------------------------------------------------------------------------------------------
    Dim Curdate As Date, byt�Ŷӽк�ģʽ As Boolean
    Dim bytNoDay As Byte
   
    byt�Ŷӽк�ģʽ = mTy_Para.byt�Ŷӽк�ģʽ
    mstrQueuePrivs = ";" & GetPrivFunc(glngSys, 1160) & ";"
    
    '143274:���ϴ���2019/7/26,�������Ա�����С����п��ҡ�Ȩ�ޣ�����Ҫ����������Ƿ��ǲ���Ա����������
    mTy_Para.str������� = Get�������(glngSys, mlngModul, mstrPrivs)
    mTy_Para.str�ٴ����� = zlGetDept(mTy_Para.str�������)
    mTy_Para.int������Ч���� = zlDatabase.GetPara("������Ч����", glngSys, mlngModul, "1")  '����:27600
    mTy_Para.byt�Ŷӽк�ģʽ = Val(zlDatabase.GetPara("�Ŷӽк�ģʽ", glngSys, mlngModul))
    mTy_Para.byt����վ�� = Val(zlDatabase.GetPara("�ŶӺ���վ��", glngSys, mlngModul))
    mTy_Para.bln������� = Val(zlDatabase.GetPara("�������������", glngSys, mlngModul)) = 1
    mTy_Para.blnAutoRefresh = Val(zlDatabase.GetPara("�Զ�ˢ��", glngSys, mlngModul, 0)) = 1
    mTy_Para.byt��������ʽ = Val(zlDatabase.GetPara("��������ʽ", glngSys, mlngModul, 0)) '���ﲡ�˵�����ʽ,0-���ұ���,����,���ݺ�;1-���ұ���,����,�Һ�ʱ��;
    
    mTy_Para.bln��Һ�ģʽ = Val(zlDatabase.GetPara("��Һ�ģʽ", glngSys)) = 1
    
    mbytViewScrop(0) = IIf(Val(zlDatabase.GetPara("��ʾ���ﲡ��", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(1) = IIf(Val(zlDatabase.GetPara("��ʾ���ﲡ��", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(2) = IIf(Val(zlDatabase.GetPara("��ʾ���ﲡ��", glngSys, mlngModul, 0)) = 1, 1, 0)
    mbytViewScrop(3) = IIf(Val(zlDatabase.GetPara("��ʾ�����ﲡ��", glngSys, mlngModul, 0)) = 1, 1, 0)
    
    Curdate = zlDatabase.Currentdate
    Set mcllFilter = New Collection
    bytNoDay = IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency)
    
    mcllFilter.Add Array(Format(DateAdd("D", -1 * bytNoDay, Curdate), "yyyy-mm-dd 00:00:00"), Format(Curdate, "yyyy-mm-dd 23:59:59")), "�Һ�ʱ��"
    mcllFilter.Add Array("", ""), "�Һ�NO"
    mcllFilter.Add Array("", ""), "��Ʊ��"
    mcllFilter.Add "", "�Һ�Ա"
    mcllFilter.Add "", "����"
    mcllFilter.Add "", "�����": mcllFilter.Add "", "���￨��"
    mcllFilter.Add "", "ҽ����": mcllFilter.Add "", "��������"
    mcllFilter.Add 0, "KIND"
    mcllFilter.Add 0, "����ID"
    mcllFilter.Add "  And A.����ʱ�� Between [1] And [2]", "����"
    mfrmTriageMgr.zlSetFilterCons mcllFilter
    Call mfrmTriageMgr.zlSetViewScrop(0, mbytViewScrop(0))
    Call mfrmTriageMgr.zlSetViewScrop(1, mbytViewScrop(1))
    Call mfrmTriageMgr.zlSetViewScrop(2, mbytViewScrop(2))
    Call mfrmTriageMgr.zlSetViewScrop(3, mbytViewScrop(3))
    
    mfrmTriageMgr.zl������� = mTy_Para.str�������
    mfrmTriageMgr.zl��Ч���� = mTy_Para.int������Ч����
    tmrBrush.Enabled = mTy_Para.blnAutoRefresh
    Call mfrmTriageMgr.zlInitVar(Me, mTy_Para.byt��������ʽ)
    If blnPatiSet And byt�Ŷӽк�ģʽ <> mTy_Para.byt�Ŷӽк�ģʽ Then
        Call Check�Ŷӽк�
        Call InitPage: cbsThis.RecalcLayout
    End If
End Sub

Private Sub InitPage()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ��ҳ��
    '���ƣ����˺�
    '���ڣ�2010-06-01 16:12:58
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    
    Err = 0: On Error GoTo Errhand:
    Call tbPage.RemoveAll
     
    Set ObjItem = tbPage.InsertItem(pg_Page.pg_����ҳ, "�������", mfrmTriageMgr.Hwnd, 0)
    ObjItem.Tag = pg_Page.pg_����ҳ
    '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    If Not mobjQueue Is Nothing And InStr(mstrQueuePrivs, ";����;") > 0 And mTy_Para.byt�Ŷӽк�ģʽ <> 0 Then
        Set ObjItem = tbPage.InsertItem(pg_�Ŷ�ҳ, "�Ŷӽк�", mobjQueue.zlGetForm.Hwnd, 0)
        ObjItem.Tag = pg_�Ŷ�ҳ
        If mTy_Para.bln��Һ�ģʽ Then
            ObjItem.Selected = True
        Else
             tbPage.Item(0).Selected = True
        End If
    Else
         tbPage.Item(0).Selected = True
    End If
    
     With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 Private Sub SubPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Select Case tbPage.Selected.Tag
    Case pg_Page.pg_����ҳ
        mfrmTriageMgr.zlSubPrint (bytMode)
    Case pg_Page.pg_�Ŷ�ҳ
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    '------------------------------------
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call SubPrint(2)
    Case conMenu_File_Print: Call SubPrint(1)
    Case conMenu_File_Excel: Call SubPrint(3)
    Case conMenu_Manage_Plan 'ǩ��
        Call mfrmTriageMgr.zlExcǩ��(False, True)
        Call zlRefreshQueueData
    Case conMenu_File_BillPrint '�Ŷӵ���ӡ
            Call mfrmTriageMgr.zlRePrintBill
    '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    Case conMenu_File_BarcodePrint
        Call mfrmTriageMgr.zlPrintBarcode
    Case conMenu_Manage_Logout 'ȡ��ǩ��
        Call mfrmTriageMgr.zlExcǩ��(True, True)
        Call zlRefreshQueueData
    Case conmenu_Edit_BackHospitalize  '����
        Call mfrmTriageMgr.zlExc����(False, True)
        Call zlRefreshQueueData
    Case conmenu_Edit_BackHospitalizeCancel 'ȡ������
        Call mfrmTriageMgr.zlExc����(True, True)
        Call zlRefreshQueueData
    Case conMenu_Edit_Triage   ' ����
        Call mfrmTriageMgr.zlExecuteTriage(Me)
    Case conmenu_Edit_ChangeNum    '���
        Call mfrmTriageMgr.zlExcuteChangeNum(Me)
    Case conMenu_Edit_ModiyPati  '����������Ϣ
        Call mfrmTriageMgr.zlExcuteEditPati(Me)
    '73743:���ϴ�,2014-7-3,���˻�����Ϣ����
    Case conMenu_Edit_ModiyPatiBaseInfo  '���˻�����Ϣ����
        Call mfrmTriageMgr.zlModiyPatiBaseInfo(Me)
    Case conmenu_Edit_Leave  '���˲�����
        Call mfrmTriageMgr.zlExcutePatiLeave(Me)
    Case conmenu_Edit_Wait '���˴���
        Call mfrmTriageMgr.zlExcutePatiWait(Me)
    Case conMenu_Manage_Finish '��ɾ���
        Call zlExcutePatiOver: Call tmrBrush_Timer
    Case conMenu_Manage_Redo  '�ָ�����
         Call mfrmTriageMgr.zlExcutePatiCancelOver(Me): Call tmrBrush_Timer
    Case conmenu_View_TriagePati     '��ʾ���ﲡ��
        mbytViewScrop(0) = IIf(mbytViewScrop(0) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(0, mbytViewScrop(0), True)
        zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    Case conmenu_View_AdmissionsPati    '��ʾ���ﲡ��
        mbytViewScrop(1) = IIf(mbytViewScrop(1) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(1, mbytViewScrop(1), True)
        zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(1), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    Case conmenu_View_OverPati    '��ʾ�Ѿ��ﲡ��
        mbytViewScrop(2) = IIf(mbytViewScrop(2) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(2, mbytViewScrop(2), True)
        zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(2), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    Case conmenu_View_Leave    '��ʾ�����ﲡ��
        mbytViewScrop(3) = IIf(mbytViewScrop(3) = 1, 0, 1)
        Call mfrmTriageMgr.zlSetViewScrop(3, mbytViewScrop(3), True)
        zlDatabase.SetPara "��ʾ�����ﲡ��", mbytViewScrop(3), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    Case conmenu_View_AutoRefresh    '�Զ�ˢ��
        
        mTy_Para.blnAutoRefresh = Not mTy_Para.blnAutoRefresh
        zlDatabase.SetPara "�Զ�ˢ��", IIf(mTy_Para.blnAutoRefresh, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
        tmrBrush.Enabled = mTy_Para.blnAutoRefresh
        Call zlRefreshData
    Case conMenu_View_Refresh   'ˢ��
        Call zlRefreshData
    Case conMenu_View_Filter  '����
        Call zlSetFilterCons
    Case conMenu_View_StatusBar
        stbThis.Visible = Not stbThis.Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Button
        cbsThis(2).Visible = Not cbsThis(2).Visible
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        cbsThis.Options.LargeIcons = Not cbsThis.Options.LargeIcons
        cbsThis.RecalcLayout
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_Parameter: Call zlParaSet
    Case conMenu_View_Find
           If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
    Case conMenu_File_Exit: Unload Me
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                mfrmTriageMgr.zlExcuteReport Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1))
        Else
             If Check�Ŷӽк� Then mobjQueue.zlExecuteCommandBars Control
        End If
        Dim strOut As String
        gobjRegist.zlExecuteCommandBars Me, Control, strOut
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnTriagePage As Boolean, bytQueue As Byte
    
    Err = 0: On Error Resume Next
    blnTriagePage = pg_Page.pg_����ҳ = Val(tbPage.Selected.Tag)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.index
        Case conMenu_EditPopup
          Control.Visible = blnTriagePage
        End Select
    End If
    Select Case Control.ID
    Case conMenu_View_Refresh
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Visible = blnTriagePage
            Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsHaveData
    Case conMenu_Manage_Plan
        '95637:���ϴ�,2016/7/17,�Һ������Ŷ���������ǩ��
        If Check�Ŷӽк� Then
             Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False
             Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs����ǩ��(bytQueue)
             Control.Visible = False
             Control.Caption = IIf(bytQueue = 0, "����ǩ��(&Q)", "����ǩ��(&Q)")
             Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False  'ˢ�±���
        Else
            Control.Visible = False
        End If
    Case conMenu_File_BillPrint '�ش��Ŷӵ�
            Control.Visible = InStr(1, mstrPrivs, ";�����Ŷӵ�;") > 0 And blnTriagePage
    '77412:���ϴ���2014/9/3,���ﲡ�������ӡ
    Case conMenu_File_BarcodePrint '�����ӡ
            Control.Visible = InStr(1, mstrPrivs, ";�����ӡ;") > 0 And blnTriagePage
    Case conMenu_Manage_Logout  'ȡ��ǩ��
            '95637:���ϴ�,2016/7/17,�Һ������Ŷ�����ȡ��ǩ��
            If Check�Ŷӽк� Then
                 Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False
                 Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs����ȡ��ǩ��
            Else
                Control.Visible = False
            End If
    Case conmenu_Edit_BackHospitalize   '����
        If Check�Ŷӽк� Then
            Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False
            Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs�������(bytQueue)
            Control.Visible = False
            
            Control.Caption = IIf(bytQueue = 0, "����(&H)", "��������ǩ��(&H)")
            Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False
        Else
            Control.Visible = False: Control.Enabled = False
        End If
    Case conmenu_Edit_BackHospitalizeCancel  'ȡ������
            If Check�Ŷӽк� Then
                Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False
                Control.Enabled = Control.Visible And mfrmTriageMgr.zlIs����ȡ������
            Else
                Control.Visible = False: Control.Enabled = False
            End If
    Case conMenu_Edit_Triage    '����
        Control.Visible = blnTriagePage
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsTriage
    Case conmenu_Edit_ChangeNum   '����
        Control.Visible = blnTriagePage
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsTriage
        
    Case conMenu_Edit_ModiyPati  '����������Ϣ
        Control.Visible = blnTriagePage And InStr(mstrPrivs, ";�����޸�;") > 0
        '73743:���ϴ�,2014-7-21,���˻�����Ϣ����
    Case conMenu_Edit_ModiyPatiBaseInfo  '�������˻�����Ϣ
        Control.Visible = blnTriagePage And InStr(1, ";" & GetPrivFunc(glngSys, 9003) & ";", ";������Ϣ����;") > 0
    Case conmenu_Edit_Leave  '���˲�����
        Control.Visible = blnTriagePage
         Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiLeave
    Case conmenu_Edit_Wait '���˴���
        Control.Visible = blnTriagePage
         Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiWait
    Case conMenu_Manage_Finish '��ɾ���
        Control.Visible = blnTriagePage And InStr(mstrPrivs, "��ɾ���") > 0 'ֻ��"��ɾ���"�Ĳſ��Խ��б�ע������ɹ���
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiFinish
    Case conMenu_Manage_Redo  '�ָ�����
        Control.Visible = blnTriagePage And InStr(mstrPrivs, "��ɾ���") > 0   'ֻ��"��ɾ���"�Ĳſ��Խ��б�ע������ɹ���
        Control.Enabled = Control.Visible And mfrmTriageMgr.zlIsPatiReDo
    Case conMenu_EditPopup  '�༭
        Control.Visible = blnTriagePage
    Case conmenu_View_TriagePati    '��ʾ�ѷ��ﲡ��
        Control.Checked = (mbytViewScrop(0) = 1)
        Control.Visible = blnTriagePage
    Case conmenu_View_AdmissionsPati    '��ʾ�ѽ��ﲡ��
        Control.Checked = (mbytViewScrop(1) = 1)
        Control.Visible = blnTriagePage
    Case conmenu_View_OverPati    '��ʾ����ɲ���
        Control.Checked = (mbytViewScrop(2) = 1)
        Control.Visible = blnTriagePage
    Case conmenu_View_Leave    '��ʾ�����ﲡ��
        Control.Checked = (mbytViewScrop(3) = 1)
        Control.Visible = blnTriagePage
    Case conmenu_View_AutoRefresh   '�Զ�ˢ��
        If Not IsStartMsgModule Then    'ֱ�ӵ���,��������������(�Ѿ���ѯ���̸���)
            Control.Checked = mTy_Para.blnAutoRefresh
        Else
            '��������Ϣƽ̨,�����������Զ�ˢ��
            Control.Visible = False
        End If
    Case conMenu_View_Filter  '��������
        Control.Visible = blnTriagePage
    Case conMenu_View_LocationItem, conMenu_View_Find 'ֻ�з���ҳ��Ŵ���.
        Control.Visible = blnTriagePage   'And mTy_Para.bln��Һ�ģʽ = False
        
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case conMenu_View_FindType       'ָ������
        If Control.Parent Is cbsThis.ActiveMenuBar And mTy_Para.bln��Һ�ģʽ = False Then
            Control.Caption = "" & mstrCaption & "��"
        End If
        Control.Visible = blnTriagePage And mTy_Para.bln��Һ�ģʽ = False '42532
    Case Else
        If Check�Ŷӽк� Then mobjQueue.zlUpdateCommandBars Control
        gobjRegist.zlUpdateCommandBars Control
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call SetFocusPatiTextBox
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        DoEvents
        If txtValue.Visible = True And txtValue.Enabled Then
            Call txtValue.SetFocus
        End If
    Else
        IDKind.ActiveFastKey
    End If
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Val(tbPage.Selected.Tag) = pg_Page.pg_�Ŷ�ҳ Then Exit Sub
    
    If KeyAscii = vbKeyReturn And Not Me.ActiveControl Is txtValue And mTy_Para.bln��Һ�ģʽ = False Then
        Call mfrmTriageMgr.zlExcuteFunction
    End If
End Sub

Private Sub Form_Load()
    Err = 0: On Error Resume Next
    mblnFirst = True
    mblnUnload = False
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.Hwnd)
    Call mobjICCard.SetParent(Me.Hwnd)
    
    Set mfrmTriageMgr = New frmTriageManager
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitVar
    Call Check�Ŷӽк�
    Call InitIDKind
    Call InitRegist
    Call zlDefCommandBars
    Call InitPancel
    Call InitPage
    Call zlRefreshQueueData
'    ����108110,��ε���ˢ�·����б�
'    Call zlRefreshData
    '��ʼ����Ϣ���Ͷ���
    Call InitMsgModule
    Call mfrmTriageMgr.zlSetobjMsgModule(mobjMsgModule)
    
    If mblnUnload Then Unload Me
End Sub

Private Sub InitRegist()
    '��ʼ���Һ�
    Dim strDept As String
    If gobjRegist Is Nothing Then Call CreateRegisterObject
    gobjRegist.zlInitData 0, , mTy_Para.str�������
End Sub
 
 
Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If txtValue.Visible Then txtValue.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtValue.Text = objPatiInfor.����
    If objCard.���� Like "*���֤*" Then
        Call zlRefreshData(True, Trim(txtValue.Text), 2, True)
    ElseIf objCard.���� Like "*IC��*" Or objCard.���� Like "*�ɣÿ�*" Then
        Call zlRefreshData(True, Trim(txtValue.Text), 3, True)
    Else
        Call zlRefreshData(True, Trim(txtValue.Text), 1, True)
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    tmrBrush.Enabled = False
    Call SaveWinState(Me, App.ProductName)
    Err = 0: On Error Resume Next
    If Not mobjIDCard Is Nothing Then
         Call mobjIDCard.SetEnabled(False)
         Set mobjIDCard = Nothing
     End If
     If Not mobjICCard Is Nothing Then
         Call mobjICCard.SetEnabled(False)
         Set mobjICCard = Nothing
     End If
    zlDatabase.SetPara "�Զ�ˢ��", IIf(mTy_Para.blnAutoRefresh, 1, 0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(0), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(1), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʾ���ﲡ��", mbytViewScrop(2), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    zlDatabase.SetPara "��ʾ�����ﲡ��", mbytViewScrop(3), glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    If Not mfrmTriageMgr Is Nothing Then Unload mfrmTriageMgr
    Set mfrmTriageMgr = Nothing
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    
    If Not mfrmOferWin Is Nothing Then Unload mfrmOferWin
    Set mfrmOferWin = Nothing
    mblnFirst = False
    
    '��ж��Ϣ���Ͷ���
    Call UnloadMsgModule
End Sub

 
Private Sub mfrmOferWin_GetNumSucces(ByVal strNO As String)
    '����ˢ������
    Call zlRefreshData
End Sub

Private Sub mfrmTriageMgr_zlPopuMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    Set objPopup = cbsThis.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
End Sub
 

Private Sub mfrmTriageMgr_zlQueueAsk(intType As Integer, strNO As String, lng����ID As Long, Cancel As Boolean)
  '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ܲ�����,��
    '��Σ�intType:1-����;2-����;3-���˲�����;4-���˴���;5-������ɾ���;6-����ȡ������,7-����
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-06-03 14:15:46
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strQueueName As String, lngID As Long
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim i As Byte
    Err = 0: On Error GoTo Errhand: '48792
    If Check�Ŷӽк� = False Then Exit Sub
    
    strSQL = "SELECT ID,ִ�в���ID,����,ִ����,nvl(����ID,0) as ����ID  From ���˹Һż�¼ where NO=[1] and ��¼����=1 and ��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    strQueueName = Nvl(rsTemp!ִ�в���id)
'    If Nvl(rsTemp!ִ����) <> "" Then
'        strQueueName = strQueueName & ":" & Nvl(rsTemp!ִ����)
'    ElseIf Nvl(rsTemp!����) <> "" Then
'        strQueueName = strQueueName & ":" & Nvl(rsTemp!����)
'    End If
    lngID = Val(Nvl(rsTemp!ID))
    Select Case intType
    Case 1  '-����;
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
        If mTy_Para.bln������� = False Then Exit Sub
        mobjQueue.zlQueueExec strQueueName, 0, lngID, IIf(mTy_Para.byt�Ŷӽк�ģʽ = 2, 5, 1)
    Case 2  '����
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
        If mTy_Para.bln������� = False Then Exit Sub
        mobjQueue.zlQueueExec strQueueName, 0, lngID, IIf(mTy_Para.byt�Ŷӽк�ģʽ = 2, 5, 1)
    Case 3   ' ���˲�����;
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 3
    Case 4, 6   '���˴���,'����ȡ������
        ' 0-����,1-ֱ��,2-����,3-��ͣ,4-��ɾ���,5-�㲥
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 0
    Case 5  '������ɾ���
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 4
    Case 7  '����
        mobjQueue.zlQueueExec strQueueName, 0, lngID, 6
    End Select
    Call zlRefreshQueueData
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub mfrmTriageMgr_zlShowInfor(strShowInfor As String)
    Me.stbThis.Panels(2).Text = strShowInfor
End Sub

 Private Sub zlParaSet()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������
    '���ƣ����˺�
    '���ڣ�2010-06-01 15:47:06
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    frmDistPara.mstrPrivs = mstrPrivs
    frmDistPara.mlngModul = mlngModul
    mlngTimerState = IIf(tmrBrush.Enabled, 1, 0): tmrBrush.Enabled = False
    
    frmDistPara.Show 1, Me
    Call InitVar(True)
    Call zlRefreshData
    
    gobjRegist.zlInitData 0, , mTy_Para.str�������
    If Not mfrmOferWin Is Nothing Then
        If mfrmOferWin.zlInitVar(Me, mTy_Para.str�������, mlngModul, mstrPrivs, gobjSquare.objSquareCard, gobjRegist) = False Then Unload Me
    End If
    tmrBrush.Enabled = mlngTimerState = 1
End Sub

Private Sub zlSetFilterCons()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ù�������
    '���ƣ����˺�
    '���ڣ�2010-06-01 16:00:34
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim cllFilter As Variant
    If mTy_Para.blnAutoRefresh Then
        If MsgBox("�Զ�ˢ��״̬�������������ˡ�" & vbCrLf & "���ڽ�ֹ�Զ�ˢ����", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
        mTy_Para.blnAutoRefresh = False:  tmrBrush.Enabled = False
    End If
    Set cllFilter = mcllFilter
    If frmDistFilter.zlShowMe(Me, mlngModul, cllFilter, mstrPrivs) = False Then
        Exit Sub
    End If
    Set mcllFilter = cllFilter
    txtValue.Text = ""
    Call mfrmTriageMgr.zlSetFilterCons(cllFilter)
    
    mfrmTriageMgr.zlintFindKeys = mintFindType
    Call zlRefreshData(True)
End Sub
 
Private Sub mobjMsgModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '������Ϣ��������
    Dim strRegistIds As String, strRegisterID As String, strRegisterDeptdId  As String
    
    If mblnExistNewMsg Then Exit Sub '������Ϣ,�Ͳ�����ȷ��,ֱ���˳�
    If UCase(strMsgItemIdentity) <> "ZLHIS_REGIST_001" Then Exit Sub
    If strMsgContent = "" Then Exit Sub
    If mfrmTriageMgr Is Nothing Then Exit Sub
    
    If Val(tbPage.Selected.Tag) = pg_Page.pg_�Ŷ�ҳ Then
        strRegistIds = "," & mobjQueue.GetQueueBusinessDataIDs() & ","
    Else
        strRegistIds = "," & mfrmTriageMgr.zlGetRegistIDsed & ","
    End If
    
    If zlXML.OpenXMLDocument(strMsgContent) = False Then Exit Sub
    If zlXML.GetSingleNodeValue("register_info/register_id", strRegisterID) = False Then Exit Sub
    If zlXML.GetSingleNodeValue("register_info/register_dept_id", strRegisterDeptdId) = False Then Exit Sub

    If InStr(1, strRegistIds, "," & Val(strRegisterDeptdId) & ",") = 0 _
        And (InStr(1, "," & mTy_Para.str������� & ",", "," & strRegisterDeptdId & ",") = True _
              Or mTy_Para.str������� = "") Then
            mblnExistNewMsg = True
    End If
End Sub

Private Sub mobjQueue_OnQueueChange(ByVal lngQueueId As Long, ByVal strQueue As String, ByVal strPatient As String, ByVal strRoom As String, ByVal strDoctor As String, blnIsAllowChange As Boolean, blnIsAlreadyProcess As Boolean)
    '����������Ϣ�󣬸����ŶӶ���
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str�ŶӺ��� As String, str�Ŷ���� As String, dat�Ŷ�ʱ�� As Date
    On Error GoTo Errhand
    
    strSQL = "Select ��������,ҵ��ID,����ID,�ŶӺ��� From �ŶӽкŶ��� Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ", lngQueueId)
    If rsTemp.EOF Then Exit Sub
    mty_Queue.str�������� = Nvl(rsTemp!��������)
    mty_Queue.lngҵ��ID = Val(Nvl(rsTemp!ҵ��ID))
    mty_Queue.lng����ID = Val(Nvl(rsTemp!����ID))
    mty_Queue.str�ŶӺ��� = Nvl(rsTemp!�ŶӺ���)
    
    '�������ƶ������˵�������Ҫ�����Ŷ�
    If mty_Queue.str�������� <> strQueue Then blnIsAllowChange = True: Exit Sub
    
    '�ŶӺ���
    strSQL = "Select zl_Get_ReQueue(4,[1],[2],[3],[4]) as �ŶӺ��� From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ŶӺ���", mty_Queue.lngҵ��ID, mty_Queue.lng����ID, strDoctor, strRoom)
    If Not rsTemp.EOF Then str�ŶӺ��� = Nvl(rsTemp!�ŶӺ���)
    '�Ŷ�ʱ��
    strSQL = "Select zl_Get_ReQueueDate(4,[1],[2],[3],[4]) as �Ŷ�ʱ�� From Dual "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ŶӺ���", mty_Queue.lngҵ��ID, mty_Queue.lng����ID, strDoctor, strRoom)
    If Not rsTemp.EOF Then dat�Ŷ�ʱ�� = rsTemp!�Ŷ�ʱ��
    
    If mty_Queue.str�ŶӺ��� <> str�ŶӺ��� Then
        '���뷢���˱仯���Ŷ����Ҳ���»�ȡ
        strSQL = "Select Zlgetsequencenum(0, [1], 1) as �Ŷ���� From Dual "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�Ŷ����", mty_Queue.lngҵ��ID)
        If Not rsTemp.EOF Then str�Ŷ���� = Nvl(rsTemp!�Ŷ����)
        
        strSQL = "ZL_�ŶӽкŶ���_UPDATE('" & strQueue & "'," & 0 & ",'" & mty_Queue.lngҵ��ID _
                    & "'," & mty_Queue.lng����ID & ",'" & strPatient & "','" _
                    & strRoom & "','" & strDoctor & "','" & str�ŶӺ��� & "','" & str�Ŷ���� & "',To_Date('" & CStr(dat�Ŷ�ʱ��) & "', 'YYYY-MM-DD hh24:mi:ss'))"
    Else
        'ֻ�����������ƣ��������������ң�ҽ����Ϣ
        strSQL = "ZL_�ŶӽкŶ���_UPDATE('" & strQueue & "'," & 0 & ",'" & mty_Queue.lngҵ��ID _
                    & "'," & mty_Queue.lng����ID & ",'" & strPatient & "','" _
                    & strRoom & "','" & strDoctor & "')"
    End If
    zlDatabase.ExecuteProcedure strSQL, "�޸Ķ�����Ϣ"
    blnIsAllowChange = True: blnIsAlreadyProcess = True
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub mobjQueue_OnQueueRoomLoad(ByVal strҵ��ID As String, rsRoomData As ADODB.Recordset, rsDoctorData As ADODB.Recordset)
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lngִ�п���ID As Long, lngת�����ID As Long
    Dim bytRegistMode As Byte
    On Error GoTo Errhand
    
    If gbytRegistMode = 0 Then
        bytRegistMode = 0
    ElseIf Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        bytRegistMode = 0
    Else
        bytRegistMode = 1
    End If
    
    If bytRegistMode = 0 Then
        strSQL = "Select ת�����ID,ִ�в���ID,�ű� From ���˹Һż�¼ Where ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�Һ���Ŀ", CLng(strҵ��ID))
        If rsTemp.EOF Then Exit Sub
        lngִ�п���ID = Val(Nvl(rsTemp!ִ�в���id))
        lngת�����ID = Val(Nvl(rsTemp!ת�����ID))
        
        If lngת�����ID = 0 Then
            strSQL = _
                " Select  B.ID As RoomId,B.���� As RoomName,B.���� As RoomCode,B.����" & vbNewLine & _
                " From �ҺŰ������� a, �������� b, �ҺŰ��� c, ���˹Һż�¼ d" & vbNewLine & _
                " Where a.�������� = b.���� And a.�ű�id = c.Id And c.���� = d.�ű� And d.ID = [1] " & vbNewLine & _
                "  And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
                " Order by B.����"
        Else
            '����Ѿ�ת���ˣ���ֻ��ͨ��ת�����ȥȷ����������
            strSQL = _
                " Select Distinct B.ID As RoomId,B.���� As RoomName,B.���� As RoomCode,B.����" & vbNewLine & _
                " From �ҺŰ������� a, �������� b, �ҺŰ��� c" & vbNewLine & _
                " Where a.�������� = b.���� And a.�ű�id = c.Id And c.ID IN (Select ID From �ҺŰ��� Where ����ID=[2]) " & vbNewLine & _
                "  And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
                " Order by B.����"
        End If
        Set rsRoomData = zlDatabase.OpenSQLRecord(strSQL, "��������", CLng(strҵ��ID), lngת�����ID)
    Else
        strSQL = "Select ת�����ID,ִ�в���ID,�����¼ID From ���˹Һż�¼ Where ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�Һ���Ŀ", CLng(strҵ��ID))
        If rsTemp.EOF Then Exit Sub
        lngִ�п���ID = Val(Nvl(rsTemp!ִ�в���id))
        lngת�����ID = Val(Nvl(rsTemp!ת�����ID))
        
        If lngת�����ID = 0 Then
            strSQL = _
                " Select B.ID As RoomId,B.���� As RoomName,B.���� As RoomCode,B.����" & vbNewLine & _
                " From �ٴ��������Ҽ�¼ a, �������� b, ���˹Һż�¼ d" & vbNewLine & _
                " Where a.����id = b.id And a.��¼id = d.�����¼id And d.ID = [1] " & _
                "  And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
                " Order by B.����"
        Else
            strSQL = _
                " Select Distinct B.ID As RoomId,B.���� As RoomName,B.���� As RoomCode,B.����" & vbNewLine & _
                " From �ٴ��������Ҽ�¼ a, �������� b" & vbNewLine & _
                " Where a.����id = b.id And a.��¼id IN (Select ID From �ٴ������¼  Where ����ID=[2]) " & _
                "  And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
                " Order by B.����"
        End If
        Set rsRoomData = zlDatabase.OpenSQLRecord(strSQL, "��������", CLng(strҵ��ID), lngת�����ID)
        
        If rsRoomData.EOF Then
            strSQL = _
                " Select B.ID As RoomId,B.���� As RoomName,B.���� As RoomCode,B.����" & vbNewLine & _
                " From �����������ÿ��� a, �������� b" & vbNewLine & _
                " Where a.����ID = B.ID And a.����ID = [1] " & vbNewLine & _
                "  And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null) " & vbNewLine & _
                " Order by B.����"
            Set rsRoomData = zlDatabase.OpenSQLRecord(strSQL, "��������", IIf(lngת�����ID = 0, lngִ�п���ID, lngת�����ID))
        End If
    End If
    
    '��ȡ�����µ�ҽ��
    strSQL = "Select c.���� as DoctorIdName,c.���� as DoctorIdCode,c.id as DoctorId From ��Ա����˵�� a, ������Ա b ,��Ա�� c" & vbNewLine & _
            " Where b.��Աid=c.id And b.��Աid=a.��Աid  And  a.��Ա����=[1] And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) " & vbNewLine & _
            " And (c.վ��='" & gstrNodeNo & "' Or c.վ�� is Null) And b.����id = [2]"
    Set rsDoctorData = zlDatabase.OpenSQLRecord(strSQL, "����ҽ��", "ҽ��", IIf(lngת�����ID = 0, lngִ�п���ID, lngת�����ID))
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub picSearch_Resize()
    Err = 0: On Error Resume Next
    With picSearch
        txtValue.Width = .ScaleWidth - IDKind.Width
    End With
End Sub

 

Private Sub picTabPage_Resize()
    Err = 0: On Error Resume Next
    With picTabPage
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
        tbPage.Width = .ScaleWidth
    End With
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim cbrControl As CommandBarControl
    Dim panThis As Pane
    Dim i As Long
    
    Call ShowAndHideOfferWin
      
    Call LockWindowUpdate(Me.Hwnd)
    Call ClearMenuItem
    Call zlDefCommandBars
    If Check�Ŷӽк� Then GoTo GoEnd:
    
    If Val(tbPage.Selected.Tag) = pg_Page.pg_�Ŷ�ҳ Then
        '�����ŶӶ�����Ϣ
        Call mobjQueue.zlDefCommandBars(cbsThis)
        For i = 1 To cbsThis.Count
            If i <> 1 Then
                For Each cbrControl In cbsThis(i).Controls
                    cbrControl.Style = xtpButtonIconAndCaption
                Next
            End If
        Next
    End If
        
GoEnd:
    Call LockWindowUpdate(0)
    Call zlRefreshData
    Call SetFocusPatiTextBox
End Sub
Private Sub SetFocusPatiTextBox()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ƶ������������
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�Fale
    '����:���˺�
    '����:2018-01-26 09:48:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mTy_Para.bln��Һ�ģʽ = False Then Exit Sub
    
    If Val(tbPage.Selected.Tag) <> pg_�Ŷ�ҳ Then Exit Sub
    
    If Not mfrmOferWin Is Nothing Then
        With mfrmOferWin.PatiIdentify
            If .Enabled And .Visible Then .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub



Private Sub txtValue_Change()
    If Me.ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '0-�����;1-����;2-�Һŵ�;3-���￨��;4-ҽ����
        If IDKind.GetCurCard.���� = "�Һŵ�" And txtValue.Text <> "" Then
            If Not (InStr(txtValue.Text, "-") = 1 Or InStr(txtValue.Text, "*") = 1) Then txtValue.Text = GetFullNO(txtValue.Text, 12)
        End If
        Call zlRefreshData(True, Trim(txtValue.Text), , True)
        zlControl.TxtSelAll txtValue
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    '0-�����,1-����,2-�Һŵ�,3-���￨��,4-ҽ����
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    strKind = IDKind.GetCurCard.����
    txtValue.PasswordChar = IIf(IDKind.GetCurCard.�������Ĺ��� <> "", "*", "")
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtValue.IMEMode = 0
    
    'ȡȱʡ��ˢ����ʽ
            '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
            '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
            '��7λ��,��ֻ��������,��Ȼȡ������
    Select Case strKind
    Case "����"
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, gobjSquare.blnȱʡ��������)
        intLen = gobjSquare.intȱʡ���ų���
    Case "�����"
        If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "�Һŵ�"
    Case "ҽ����"
    Case Else
            If IDKind.GetCurCard.�ӿ���� <> 0 Then
                blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.GetCurCard.�������Ĺ��� <> "")
                intLen = IDKind.GetCurCard.���ų���
            End If
    End Select
    
    'ˢ����ϻ���������س�
    If blnCard And Len(txtValue.Text) = intLen - 1 And KeyAscii <> 8 Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0: mblnCard = True
         Call zlRefreshData(True, Trim(txtValue.Text), 1, True)
        mblnCard = False
        zlControl.TxtSelAll txtValue
   End If
End Sub
Private Sub txtvalue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
    txtValue.Text = Trim(txtValue.Text)
End Sub

Private Sub tmrBrush_Timer()
    Static intNum As Integer
    If IsStartMsgModule Then
        '1.���ӳɹ���,��Ҫ1���Ӳ���ˢ��һ��
        '2.������Ҫ��������Ϣʱ,����ˢ��
        intNum = intNum + 1
        If intNum >= 2 Then 'ÿ��30��ִ��һ��,����Ϊ1����
           intNum = 0
           If mblnExistNewMsg Then
                mblnExistNewMsg = False
                Call zlRefreshData
           End If
        End If
    Else
        intNum = 0
        Call zlRefreshData
    End If
End Sub

Private Sub zlRefreshData(Optional blnFilter As Boolean = False, _
    Optional strFindValue As String = "", Optional bytReadType As Byte = 0, Optional ByVal blnAutoǩ�� As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ˢ������
    '��Σ�blnFilter-�Ƿ����
    '          bytReadType-��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
    '���ƣ����˺�
    '���ڣ�2010-06-02 09:43:08
    '------------------------------------------------------------------------------------------------------------------------
    mlngTimerState = Me.tmrBrush.Enabled: Me.tmrBrush.Enabled = False
    If Val(tbPage.Selected.Tag) = pg_Page.pg_�Ŷ�ҳ Then
        Call zlRefreshQueueData
    Else
        mfrmTriageMgr.zlintFindKeys = mintFindType
        Call mfrmTriageMgr.zlRefreshData(blnFilter, strFindValue, bytReadType, IDKind.GetCurCard, blnAutoǩ��)
    End If
    Me.tmrBrush.Enabled = mlngTimerState
End Sub

Public Sub zlExcutePatiOver()
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ɾ���
    '���ƣ����˺�
    '���ڣ�2010-05-31 15:52:52
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim strMsgbox As String, lng����ID As Long, lngִ��״̬ As Long
    Dim strNO As String, strȱʡ���� As String, strȱʡҽ�� As String
    Dim rsTmp As ADODB.Recordset, lngID As Long
    Dim i As Long, strSQL As String
    
    If InStr(mstrPrivs, "��ɾ���") = 0 Then Exit Sub
    lng����ID = mfrmTriageMgr.zlGet����ID
    If lng����ID = 0 Then
        MsgBox "�����ڵĲ��ˣ�", vbInformation, gstrSysName: Exit Sub
    End If
    lngִ��״̬ = mfrmTriageMgr.zlGet�Һ�ִ��״̬
    If lngִ��״̬ = 1 Then Exit Sub
    If lngִ��״̬ = 2 Then
        strMsgbox = "ҽ���Ѿ��Ըò��˽���������Ӧ��ҽ��ȷ����ɣ�" & vbCrLf & _
                    "�����������(��ҽ����������ǰ����޷���������)" & vbCrLf & _
                    "���򣬽��鲻Ҫ���иò�����" & vbCrLf & vbCrLf & _
                    "���Ҫ��������"
        If MsgBox(strMsgbox, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    strNO = mfrmTriageMgr.zlGet�Һ�NO: strȱʡҽ�� = mfrmTriageMgr.zlGet�Һ�ҽ��
    strȱʡ���� = mfrmTriageMgr.zlGet�Һ�����
    lngID = mfrmTriageMgr.zlGet�Һ�ID
    
    On Error GoTo errHandle
    If frmDistOver.zlShowEdit(Me, mstrPrivs, mstrQueuePrivs, mobjQueue, mlngModul, strNO, lng����ID, strȱʡ����, strȱʡҽ��, mTy_Para.byt�Ŷӽк�ģʽ, lngID) = False Then Exit Sub
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check�Ŷӽк�() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ʹ����ŶӽкŹ���
    '���أ��ŶӽкŹ������еĶ��Ϸ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-06 10:19:43
    '˵��������: Ȩ�޺Ϸ����;�������Ŷӽкŵ�;�����Ŷӽкųɹ�!
    '------------------------------------------------------------------------------------------------------------------------
    '�ŶӽкŴ���ģʽ:1.�������̨������л�ҽ����������;2-�ȷ������,��ҽ�����о���.0-���Ŷӽк�
    If mTy_Para.byt�Ŷӽк�ģʽ = 0 Then GoTo GoEnd:
    If Not (InStr(mstrQueuePrivs, ";����;") > 0) Then GoTo GoEnd:
    Err = 0: On Error GoTo GoEnd:
    If mobjQueue Is Nothing Then
        Set mobjQueue = CreateObject("zlQueueManage.clsQueueManage")
        mobjQueue.zlInitVar gcnOracle, glngSys, 0, mTy_Para.int������Ч����, mstrQueuePrivs, ""
    End If
    Check�Ŷӽк� = True
    Exit Function
GoEnd:
    If Not mobjQueue Is Nothing Then mobjQueue.CloseWindows
    Set mobjQueue = Nothing
End Function

'Private Sub InitMenus()
'    Dim varData As Variant, varTemp As Variant, strKind As String
'    Dim i As Long
'
'    Set mcllBrushCard = New Collection
'    strKind = "��|����|0|0|" & zlGetPatiInforMaxLen.intPatiName & "|0|0||"
'    strKind = strKind & ";" & "��|�����|0|0|18|0|0||"
'    strKind = strKind & ";" & "��|�Һŵ�|0|0|18|0|0||"
'    strKind = strKind & ";" & "��|���￨|0|0|18|0|0||"
'    strKind = strKind & ";" & "ҽ|ҽ����|0|0|64|0|0||"
'    strKind = strKind & ";" & "��|���֤��|0|0|18|0|0||"
'    strKind = strKind & ";" & "IC|IC����|0|0|50|0|0||"
'    If Not gobjSquare.objSquareCard Is Nothing Then
'        strKind = gobjSquare.objSquareCard.zlGetIDKindStr(strKind)
'    End If
'    varData = Split(strKind, ";")
'    For i = 0 To UBound(varData)
'        varTemp = Split(varData(i), "|")
'        'ȡȱʡ��ˢ����ʽ
'        '����|�����|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
'        '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
'        '��7λ��,��ֻ��������,��Ȼȡ������
'        mcllBrushCard.Add varTemp, varTemp(1)
'        If Val(varTemp(5)) = 1 Then
'            gobjSquare.blnȱʡ�������� = Trim(varTemp(7)) <> ""
'            mty_Square.lngȱʡ�����ID = Val(varTemp(3))
'            gobjSquare.intȱʡ���ų��� = Val(varTemp(4))
'            mblnȱʡ���� = Val(varTemp(2)) = 1
'        End If
'    Next
'    Call InitCardType
'End Sub

'��ʼ��IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "��|����|0;ҽ|ҽ����|0;��|���֤��|0;IC|IC����|1;��|�����|0;��|�Һŵ�|0", txtValue)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.��ȱʡ������ And Not objCard Is Nothing Then
        gobjSquare.blnȱʡ�������� = objCard.�������Ĺ��� <> ""
        gobjSquare.intȱʡ���ų��� = objCard.���ų���
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.blnȱʡ�������� = IDKind.Cards.������ʾ
        gobjSquare.intȱʡ���ų��� = 100
    End If
End Function


Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    If txtValue.Text = "" And Me.ActiveControl Is txtValue Then
        txtValue.Text = strID:
        If txtValue.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        '��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
        Call zlRefreshData(True, Trim(txtValue.Text), 2)
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    If txtValue.Text = "" And Me.ActiveControl Is txtValue Then
        txtValue.Text = strNO
        If txtValue.Text = "" Then
            Call mobjICCard.SetEnabled(False) '��������Ϸ������������ü����Զ���ȡ
            Exit Sub
        End If
        '��ȡ����(0-������;1-ˢ��;2-��ȡ���֤;3-��ȡIC��)
        Call zlRefreshData(True, Trim(txtValue.Text), 3)
    End If
End Sub
 
Private Sub InitMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    Call IsStartMsgModule   '�����Զ�ˢ��
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub UnloadMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ж��Ϣģ��
    '����:���˺�
    '����:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo Errhand:
    
    If mobjMsgModule Is Nothing Then Exit Sub
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function IsStartMsgModule() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ���������Ϣģ������(�������ӳɹ�)
    '����:������Ϣģ����������ӳɹ��ķ���true,���򷵻�False
    '����:���˺�
    '����:2014-03-11 14:42:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjMsgModule Is Nothing Then Exit Function
    If mobjMsgModule.IsConnect = False Then Exit Function
    If tmrBrush.Enabled = False Then tmrBrush.Enabled = True
    IsStartMsgModule = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ShowAndHideOfferWin()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������ȡ�Ŵ���
    '����:���˺�
    '����:2018-01-17 16:16:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    If mTy_Para.bln��Һ�ģʽ = False Then Exit Sub
  
    Set panThis = dkpMan.FindPane(conPane_OfferWin)
    If panThis Is Nothing Then Exit Sub
    
    If Val(tbPage.Selected.Tag) = pg_�Ŷ�ҳ Then
        If Not panThis.Selected Then panThis.Select
        Call SetFocusPatiTextBox
        Exit Sub
    End If
    If Not panThis.Closed Then panThis.Close
End Sub

