VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm���ֱ�׼ά�� 
   Caption         =   "���ֱ�׼ά��"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   270
   ClientWidth     =   11115
   Icon            =   "frm���ֱ�׼ά��.frx":0000
   ScaleHeight     =   8040
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   8580
      TabIndex        =   18
      ToolTipText     =   "��ݼ���F3"
      Top             =   210
      Width           =   1320
   End
   Begin VB.PictureBox picRightUp 
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   4590
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   16
      Top             =   870
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid fgMain 
         Height          =   5475
         Left            =   45
         TabIndex        =   17
         Top             =   165
         Width           =   5235
         _cx             =   9234
         _cy             =   9657
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483630
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm���ֱ�׼ά��.frx":1272
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   1
         ExplorerBar     =   0
         PicturesOver    =   -1  'True
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7680
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   635
      SimpleText      =   $"frm���ֱ�׼ά��.frx":1389
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm���ֱ�׼ά��.frx":13D0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16695
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
   Begin VB.PictureBox pic��Ŀ��Ϣ_S 
      BackColor       =   &H00FAFAFA&
      Height          =   1935
      Left            =   4605
      Picture         =   "frm���ֱ�׼ά��.frx":1C64
      ScaleHeight     =   1875
      ScaleWidth      =   5040
      TabIndex        =   13
      Top             =   3765
      Width           =   5100
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   165
         TabIndex        =   15
         Top             =   420
         Width           =   6360
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   90
         Width           =   1095
      End
   End
   Begin VB.PictureBox picLeft_S 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5610
      Left            =   135
      ScaleHeight     =   5610
      ScaleWidth      =   3255
      TabIndex        =   1
      Top             =   690
      Width           =   3255
      Begin VB.PictureBox pic������Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   5
         Top             =   1965
         Width           =   2790
         Begin VB.PictureBox picFAXX 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2415
            ScaleHeight     =   225
            ScaleWidth      =   255
            TabIndex        =   6
            Top             =   75
            Width           =   255
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   90
            Width           =   1095
         End
         Begin VB.Label lbl�������� 
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   450
            Width           =   2580
         End
         Begin VB.Label lbl��ֵ 
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ:"
            Height          =   195
            Left            =   225
            TabIndex        =   10
            Top             =   1146
            Width           =   2580
         End
         Begin VB.Label lbl��ֵ 
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ:"
            Height          =   195
            Left            =   225
            TabIndex        =   9
            Top             =   1380
            Width           =   2580
         End
         Begin VB.Label lbl�ܷ� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ܷ�:"
            Height          =   195
            Left            =   225
            TabIndex        =   8
            Top             =   914
            Width           =   2580
         End
         Begin VB.Label lbl���� 
            BackStyle       =   0  'Transparent
            Caption         =   "����:"
            Height          =   195
            Left            =   225
            TabIndex        =   7
            Top             =   682
            Width           =   2580
         End
      End
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   90
         ScaleHeight     =   1590
         ScaleWidth      =   2940
         TabIndex        =   2
         Top             =   45
         Width           =   2940
         Begin MSComctlLib.TreeView tvw���� 
            Height          =   1200
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   2117
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "���ַ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   4
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1155
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ֱ�׼ά��.frx":2161
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm���ֱ�׼ά��.frx":2FB3
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   2
      Left            =   9825
      Top             =   8475
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5415
      Picture         =   "frm���ֱ�׼ά��.frx":3227
      Top             =   8595
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   2565
      Picture         =   "frm���ֱ�׼ά��.frx":3276
      Top             =   8565
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   2940
      Picture         =   "frm���ֱ�׼ά��.frx":32CB
      Top             =   8520
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   75
      Picture         =   "frm���ֱ�׼ά��.frx":3489
      Top             =   8505
      Visible         =   0   'False
      Width           =   2790
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   270
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frm���ֱ�׼ά��.frx":3649
      Left            =   735
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
      ScaleMode       =   1
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   0
      Left            =   7725
      Picture         =   "frm���ֱ�׼ά��.frx":365D
      Top             =   8490
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Image imgBG_fg 
      Height          =   2970
      Index           =   1
      Left            =   5865
      Picture         =   "frm���ֱ�׼ά��.frx":3E81
      Top             =   8535
      Visible         =   0   'False
      Width           =   3105
   End
End
Attribute VB_Name = "frm���ֱ�׼ά��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////
'
'       ģ�飺���ֱ�׼ά��
'       ���ܣ��������ֱ�׼��¼�롢�޸ġ�ɾ������ӡ��ѡ�õȡ�
'       ��д������ΰ
'       ���ڣ�2005��1��5��
'
'///////////////////////////////////////////////////////////////////////////////


Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private mstrPrivs               As String               'Ȩ�޴�
Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private mlngModule              As Long                 'ģ���
Private m_lngOldRow             As Long                 '
Private m_lngCurRow             As Long                 '
Private m_lngCurID              As Long                 '��¼��ǰ��¼ID
Private m_lngCurFAID            As Long                 '����ID
Private m_lngCurSJID            As Long                 '�ϼ�ID
Private m_strTreeKey            As String
Private m_lngOldSJID            As Long
Private mzlPrintModeS           As gzlPrintModeS        '��ӡ
Private mintItemID              As Long                 '��׼ID
Private mcbrPopupBarProg        As CommandBar           '�������ڡ����ࡿ
Private mcbrPopupBarItem        As CommandBar           '�������ڡ���Ŀ��
Private mblnProgUsed            As Boolean              '�����Ƿ���ʹ��
Dim cbrPopupItem                As CommandBarControl    '������

'==============================================================================
'=���ܣ� �ؼ���ʼ��
'==============================================================================
Private Sub InitControl()
    On Error GoTo errH
    '�˵�����
    Call InitCommandBar
    '��������
    Call InitDockPannel
    '���Tree
    Call InitTreeView
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ�˵�������
'==============================================================================
Private Sub InitCommandBar()
    Dim objMenu         As CommandBarPopup
    Dim objBar          As CommandBar
    Dim objExtendedBar  As CommandBar
    Dim objPopup        As CommandBarPopup
    Dim objControl      As CommandBarControl
    Dim cbrCustom       As CommandBarControlCustom
    
    On Error GoTo errH
    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsMain)
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    '------------------------------------------------------------------------------------------------------------------
    '�ļ�
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Excel, "�����&Excel")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_File_Exit, "�˳�(&X)", True)
    '------------------------------------------------------------------------------------------------------------------
    '�༭
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewKind, "��������(&N)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyKind, "�޸ķ���(&F)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteKind, "ɾ������(&L)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Import, "���뷽��(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "������Ŀ(&X)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Insert, "������Ŀ(&G)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "�޸���Ŀ(&R)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ����Ŀ(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "���ӱ�׼(&A)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "�����Ʊ�׼(&I)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸ı�׼Ŀ(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ����׼(&D)")
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)", True)
    '------------------------------------------------------------------------------------------------------------------
    '����
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_Help, "��������(&H)")
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & ParamInfo.��Ʒ����)
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Home, ParamInfo.��Ʒ���� & "��ҳ(&H)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Forum, ParamInfo.��Ʒ���� & "��̳(&F)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Help_About, "����(&A)��", True)
    '------------------------------------------------------------------------------------------------------------------
    '���˵��Ҳ�Ĳ���
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    mstrFindKey = Trim(GetPara("��λ����", mlngModule, "��Ŀ", True))
    If mstrFindKey = "" Then mstrFindKey = "��Ŀ"
    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.��Ŀ", , , "��Ŀ")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.ȱ��", , , "ȱ��")
    Set cbrCustom = cbsMain.ActiveMenuBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    cbrCustom.Handle = txtLocation.hWnd
    cbrCustom.Flags = xtpFlagRightAlign
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "ǰһ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    Set objControl = cbsMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "��һ��")
    objControl.Flags = xtpFlagRightAlign
    objControl.STYLE = xtpButtonIcon
    '��׼������
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Print, "��ӡ")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Preview, "Ԥ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_NewItem, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Modify, "�޸�")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "ɾ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
        .Add FCONTROL, vbKeyI, conMenu_Edit_CopyNewItem     '����
        .Add FCONTROL, vbKeyE, conMenu_Edit_Modify          '�޸�
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete       'ɾ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Transf_Save     '����
        .Add 0, vbKeyF3, conMenu_View_Location              '��λ
        .Add 0, vbKeyF4, conMenu_View_Option                'ѡ��λ����
        .Add FCONTROL, vbKeyLeft, conMenu_View_Forward      'ǰһ��
        .Add FCONTROL, vbKeyRight, conMenu_View_Backward    '��һ��
    End With
    '------------------------------------------------------------------------------------------------------------------
    '�����˵�����
    Set mcbrPopupBarProg = cbsMain.Add("���������˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "��������(&N)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "�޸ķ���(&F)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "ɾ������(&L)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_Import, "���뷽��(&P)")
    Set cbrPopupItem = mcbrPopupBarProg.Controls.Add(xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
    Set mcbrPopupBarItem = cbsMain.Add("������Ŀ�˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "������Ŀ(&X)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Insert, "������Ŀ(&G)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸���Ŀ(&R)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ����Ŀ(&C)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "���ӱ�׼(&A)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "�����׼(&I)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�޸ı�׼(&M)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "ɾ����׼(&D)")
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ָ�
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = picLeft_S.hWnd
        Case 2
            Item.Handle = picRightUp.hWnd
        Case 3
            Item.Handle = pic��Ŀ��Ϣ_S.hWnd
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ���򻮷�
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane As Pane
    
    On Error GoTo errH
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 200, 100, DockRightOf, Nothing)
    objPane.Title = "��Ŀ"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(3, 200, 100, DockBottomOf, objPane)
    objPane.Title = "��׼"
    objPane.Options = PaneNoCaption
    dkpMain.SetCommandBars cbsMain
    Call DockPannelInit(dkpMain)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����fgMain������ˢ��״̬��Ϣ
'==============================================================================
Private Sub fgMain_Click()
    On Error GoTo errH
    fgMain_SelChange
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� fgMain
'==============================================================================
Private Sub fgMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    Select Case Button
        Case 2          '�����˵�����
            Call SendLMouseButton(fgMain.hWnd, X, Y)
            mcbrPopupBarItem.ShowPopup
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ˫���༭�����ֱ�׼����Ŀ��������Ŀʱ��
'==============================================================================
Private Sub fgMain_DblClick()
    On Error GoTo errH
    
    If InStr(mstrPrivs, "��ɾ��") = 0 Then Exit Sub
    If fgMain.MouseRow = 0 Or mblnProgUsed Then Exit Sub
    Call StandardEdit
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub fgMain_RowColChange()
    On Error GoTo errH
    Call fgMain_SelChange
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ѡ�����б仯ʱ���»�����Ϣ
'==============================================================================
Private Sub fgMain_SelChange()
    Dim lngID       As Long
    Dim rsTemp      As ADODB.Recordset
    On Error GoTo errH
    
    fgMain.WallPaper = imgBG_fg(2).Picture
    m_lngCurRow = fgMain.Row
    mblnProgUsed = False
    If m_lngCurRow <= 0 Then
        m_lngCurSJID = 0
        m_lngCurID = 0
        fgMain.WallPaper = imgBG_fg(0).Picture
        Exit Sub
    End If
    
    m_lngCurID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 4)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 4)))    '��ȡID
    m_lngCurSJID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 5)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 5)))     '��ȡID
    m_lngCurFAID = IIf(Len(fgMain.Cell(flexcpText, m_lngCurRow, 6)) = 0, 0, Val(fgMain.Cell(flexcpText, m_lngCurRow, 6)))     '��ȡID
    If m_lngCurSJID = 0 Then
        lngID = m_lngCurID
    Else
        lngID = m_lngCurSJID
    End If
    
    Show����Ҫ�� lngID, fgMain.Cell(flexcpText, m_lngCurRow, 0), fgMain.Cell(flexcpText, m_lngCurRow, 1)
    m_lngOldRow = m_lngCurRow
    
    gstrSQL = "select count(*) from �������ֽ�� where ����ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If rsTemp(0).Value > 0 Then
        '�÷����Ѿ�ʹ��
        mblnProgUsed = True
        fgMain.WallPaper = imgBG_fg(1).Picture
    End If
    Call SetMenu
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڳ�ʼ��ʱ���ڿ�λλ�ÿ���
'==============================================================================
Private Sub Form_Activate()
    On Error GoTo errH
    Call Form_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڳ�ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڳ�ʼ��
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    
    KeyPreview = True
    m_lngOldRow = -1
    m_lngCurRow = -1
    m_lngCurID = -1
    m_lngOldSJID = -1
    mblnProgUsed = False
    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngModule = ParamInfo.ģ���
    If GetPersonSet Then
        mstrFindKey = Trim(GetPara("��λ����", mlngModule, "����", True))
    End If
    '�ؼ���ʼ��
    Call InitControl
    
    '����б�
    Call DataLoad

    '�ָ�����λ��
    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs)
    Call SetMenu
    
    picFAXX.Picture = imgClose.Picture
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����λ�ñ仯
'==============================================================================
Private Sub Form_Resize()
    On Error GoTo errH

    Call SetPaneRange(dkpMain, 1, 100, 60, 450, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 2, 400, 100, ScaleHeight, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 400, 100, ScaleHeight, Me.ScaleHeight)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ڹر�ʱ�������
'==============================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errH
    m_strTreeKey = ""
    'ʹ�ø��Ի�����
    Call SetPara("��λ����", mstrFindKey, mlngModule)
    SaveWinState Me, App.ProductName
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�رջ���ʾ
'==============================================================================
Private Sub picFAXX_Click()
    On Error GoTo errH
    
    If picFAXX.Tag = "" Then
        picFAXX.Tag = "Opened"
        picFAXX.Picture = imgOpen.Picture
        pic������Ϣ.Height = 340
    Else
        picFAXX.Tag = ""
        picFAXX.Picture = imgClose.Picture
        pic������Ϣ.Height = 1695
    End If
    picFAXX.Refresh
    Call picLeft_S_Resize
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ͳ��
'==============================================================================
Private Sub DataUpdate()
    Dim rs              As ADODB.Recordset
    Dim lng�ܷ�         As Double
    On Error GoTo errH
    gstrSQL = "select ����,����,��ֵ,��ֵ,�ܷ� from �������ַ��� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        lbl��������.Caption = rs("����")
        lbl����.Caption = "����:" & rs("����")
        lbl��ֵ.Caption = "��ֵ:" & rs("��ֵ")
        lbl��ֵ.Caption = "��ֵ:" & rs("��ֵ")
        lbl�ܷ�.Caption = "�ܷ�:" & rs("�ܷ�")
        lng�ܷ� = rs("�ܷ�")
    Else
        lbl��������.Caption = ""
        lbl����.Caption = ""
        lbl��ֵ.Caption = ""
        lbl��ֵ.Caption = ""
        lbl�ܷ�.Caption = ""
    End If

    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        If Abs(lng�ܷ� - rs.Fields(0)) > 0.01 Then
            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & rs.Fields(0)
            lbl�ܷ�.ForeColor = vbRed
        Else
            lbl�ܷ�.ForeColor = vbBlack
        End If
    Else
        lbl�ܷ�.ForeColor = vbRed
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�����ɫ
'==============================================================================
Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    If X >= 0 And X <= picFAXX.ScaleWidth And Y >= 0 And Y <= picFAXX.ScaleHeight Then
        SetCapture picFAXX.hWnd
        '������룡����
        picFAXX.Line (0, 0)-(picFAXX.ScaleWidth - Screen.TwipsPerPixelX, picFAXX.ScaleHeight - Screen.TwipsPerPixelY), vbBlue, B
    Else
        '����Ƴ�������
        picFAXX.Cls
        ReleaseCapture
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���picLeft���ؼ�λ�ÿ���
'==============================================================================
Private Sub picLeft_S_Resize()
On Error Resume Next
    pic������Ϣ.Move 135, picLeft_S.ScaleHeight - pic������Ϣ.Height - 270 * 2, picLeft_S.ScaleWidth - 270
    With picTree
        .Move 135, 135, pic������Ϣ.Width, Abs(picLeft_S.ScaleHeight - pic������Ϣ.Height - 270 * 3)
        .Cls
        .PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
        .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    End With
    tvw����.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
    With pic������Ϣ
        .Cls
        .PaintPicture imgBGBlue.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBGBlue.Width, 360
        .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
        .PaintPicture imgBGBlue.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
    End With
    picFAXX.Move pic������Ϣ.ScaleWidth - picFAXX.Width - 80
    Refresh
End Sub

'==============================================================================
'=���ܣ� �Ҳ�picRightUp���ؼ�λ�ÿ���
'==============================================================================
Private Sub picRightUp_Resize()
    On Error GoTo errH
    fgMain.Move 15, 15, picRightUp.Width - 30, picRightUp.Height - 30
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҳ�picRightUp���ؼ�λ�ÿ���
'==============================================================================
Private Sub pic��Ŀ��Ϣ_S_Resize()
    On Error GoTo errH
    lblInfo.Move lblInfo.Left, lblInfo.Top, Abs(pic��Ŀ��Ϣ_S.ScaleWidth - 2 * lblInfo.Left), Abs(pic��Ŀ��Ϣ_S.ScaleHeight - lblInfo.Top)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ϲ�picTree����˫������
'==============================================================================
Private Sub picTree_DblClick()
    On Error GoTo errH
    If Left(tvw����.SelectedItem.Key, 4) = "Root" Then Exit Sub
    Call ProgEdit
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ���ϲ�picTree����˫������
'==============================================================================
Private Sub picTree_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If IsNumeric(Mid(m_strTreeKey, 2)) Then
        If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call ProgEdit
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �һ�����ʱ�����������Ҽ��˵�
'==============================================================================
Private Sub tvw����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo errH
    If InStr(mstrPrivs, "��ɾ��") = 0 Then Exit Sub
    Select Case Button
        Case 2          '�����˵�����
            Call SendLMouseButton(tvw����.hWnd, X, Y)
            mcbrPopupBarProg.ShowPopup
    End Select
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������º�ѡ����Ӧ����Ŀ�ͱ�׼
'==============================================================================
Private Sub tvw����_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rs              As ADODB.Recordset
    Dim lng�ܷ�         As Double
    
    On Error GoTo errH
    
    If m_strTreeKey = Node.Key Then Exit Sub     '�����ظ�ˢ��
    m_strTreeKey = Node.Key
    m_lngCurFAID = Val(Mid(m_strTreeKey, 2))
    
    gstrSQL = "select ����,����,��ֵ,��ֵ,�ܷ� from �������ַ��� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        lbl�������� = rs("����")
        lbl���� = "����:" & rs("����")
        lbl��ֵ = "��ֵ:" & rs("��ֵ")
        lbl��ֵ = "��ֵ:" & rs("��ֵ")
        lbl�ܷ� = "�ܷ�:" & rs("�ܷ�")
        lng�ܷ� = rs("�ܷ�")
    Else
        lbl�������� = ""
        lbl���� = ""
        lbl��ֵ = ""
        lbl��ֵ = ""
        lbl�ܷ� = ""
    End If
    
    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        If Abs(lng�ܷ� - rs.Fields(0)) > 0.01 Then
            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & rs.Fields(0)
            lbl�ܷ�.ForeColor = vbRed
        Else
            lbl�ܷ�.ForeColor = vbBlack
        End If
    Else
        lbl�ܷ�.ForeColor = vbRed
    End If
    '����б�
    Call DataLoad
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=����:�����ݱ���д�ӡ,Ԥ���������EXCEL
'=����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'==============================================================================
Private Sub subPrint(bytMode As Byte)
    Dim objPrint            As New zlPrint1Grd
    Dim objAppRow           As zlTabAppRow
    Dim bytR                As Byte
    Dim rs                  As ADODB.Recordset
    
    On Error GoTo errH
    
    Set objPrint.Body = fgMain
    objPrint.Title.Text = tvw����.SelectedItem.Text
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    Set objAppRow = New zlTabAppRow
    gstrSQL = "select ID,����,�ܷ�,��ֵ,��ֵ,����,����,ѡ��,����ʱ��,ͣ��ʱ�� from �������ַ��� where ID= [1]" ' & m_lngCurFAID
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
    If Not rs.EOF Then
        objAppRow.Add "�ܷ�:" & NVL(rs("�ܷ�"), 0)
        objAppRow.Add "�׼�������:" & NVL(rs("��ֵ"), 0)
        objAppRow.Add "�Ҽ�������:" & NVL(rs("��ֵ"), 0)
        
        objPrint.UnderAppRows.Add objAppRow
    End If
    Set objAppRow = New zlTabAppRow
    objAppRow.Add "��ӡ�ˣ�" & gstrUserName
    objAppRow.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objAppRow
    
    If bytMode = 1 Then
        bytR = zlPrintAsk(objPrint)
        If bytR <> 0 Then zlPrintOrView1Grd objPrint, bytR
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� װ�����ַ��� Ŀǰֻ����סԺ����
'==============================================================================
Private Sub InitTreeView()
    Dim rsTemp          As ADODB.Recordset
    Dim nod             As Node
    Dim i               As Long
    Dim FirstKey        As String
    Dim v               As Variant
    
    On Error GoTo errH
    
    fgMain.Tag = ""
    'Tree�ĳ�ʼ��
    Set tvw����.ImageList = ils16
    tvw����.Nodes.Clear
    
    'ע����ø�ʽ���ȸ�ֵgstrSQL,Ȼ������ݼ�
    gstrSQL = "select ID,����,ѡ�� from �������ַ��� where ����='סԺ' Order by ѡ�� desc,����,����ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    i = 1
    Do Until rsTemp.EOF
        '����ӽڵ�
        Set nod = tvw����.Nodes.Add(, , "A" & rsTemp("ID"), rsTemp("����"), IIf(rsTemp("ѡ��") = 1, "RootSel", "Root"), IIf(rsTemp("ѡ��") = 1, "RootSel", "Root"))
        If rsTemp("ѡ��") = 1 Then
            nod.Bold = True
        Else
            nod.Bold = False
        End If
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTemp.MoveNext
    Loop
    'm_strTreeKey��Ϊ�գ�������û���ҵ���
    If i = 1 Then m_strTreeKey = FirstKey
    For Each v In tvw����.Nodes
        If v.Key = FirstKey Then
            '����ѡ��
            v.Selected = True
            v.EnsureVisible
            If picTree.Visible = True Then picTree.SetFocus
        End If
    Next
    If Not tvw����.SelectedItem Is Nothing Then tvw����_NodeClick tvw����.SelectedItem
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� װ���Ӧ���������ֱ�׼
'==============================================================================
Public Sub DataLoad()
    Dim rsTemp      As ADODB.Recordset
    Dim i           As Long
    
    On Error GoTo errH
    
    With fgMain
        .Tag = ""
        .Redraw = flexRDNone
        .Rows = 1
        .Clear
        .Cell(flexcpText, 0, 0) = "��Ŀ"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 1) = "��׼��ֵ"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 2) = "��׼����"
        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 3) = "���ֱ�׼"
        .Cell(flexcpAlignment, 0, 3) = flexAlignCenterCenter
        .Cell(flexcpText, 0, 4) = "ID"
        .Cell(flexcpText, 0, 5) = "�ϼ�ID"
        .Cell(flexcpText, 0, 6) = "����ID"
        .Cell(flexcpText, 0, 7) = "���"
        
        'ȷ����������
        If tvw����.SelectedItem Is Nothing Then .Redraw = flexRDDirect: Exit Sub
        With tvw����.SelectedItem
            Select Case Left(.Key, 1)
                Case "A", "B"
                    m_lngCurFAID = Val(Mid(.Key, 2))
                    gstrSQL = "select �ϼ����,���,ID,�ϼ�ID,����ID,��Ŀ,��׼��ֵ,����Ҫ��,ȱ������,�۷ֱ�׼,���� from �������ֱ�׼��ͼ Where ����='��' and ����ID = [1]"
                Case Else
                    Call SetMenu
                    fgMain.Redraw = flexRDDirect
                    Exit Sub
            End Select
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(CStr(Mid(.Key, 2))))
        End With
        
        .FocusRect = flexFocusSolid
        '��������
        .Cols = 8
        .Rows = rsTemp.RecordCount + 1
        i = 1
        Do Until rsTemp.EOF
            .Cell(flexcpText, i, 0) = NVL(rsTemp.Fields("��Ŀ"))
            .Cell(flexcpAlignment, i, 0) = flexAlignCenterCenter
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("��׼��ֵ")), " ", Format(rsTemp.Fields("��׼��ֵ"), "####��"))
            .Cell(flexcpAlignment, i, 1) = flexAlignCenterCenter
            .Cell(flexcpText, i, 2) = NVL(rsTemp.Fields("ȱ������"))
            .Cell(flexcpAlignment, i, 2) = flexAlignLeftTop
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("�۷ֱ�׼")), "", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�׼�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "�Ҽ�", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "����", IIf(rsTemp.Fields("�۷ֱ�׼") = "��", "������", rsTemp.Fields("�۷ֱ�׼"))))))
            .Cell(flexcpAlignment, i, 3) = flexAlignCenterCenter
            .Cell(flexcpText, i, 4) = NVL(rsTemp.Fields("ID"), 0)
            .Cell(flexcpText, i, 5) = NVL(rsTemp.Fields("�ϼ�ID"), 0)
            .Cell(flexcpText, i, 6) = NVL(rsTemp.Fields("����ID"), 0)
            .Cell(flexcpText, i, 7) = NVL(rsTemp.Fields("���"), 0)
            rsTemp.MoveNext
            i = i + 1
        Loop
        '�Զ�����
        .WordWrap = True
        '�ϲ���Ԫ��
        .MergeCells = 2
        .MergeCol(.ColIndex("��Ŀ")) = True
        .MergeCol(.ColIndex("��׼��ֵ")) = True
        '��������
        .ColAlignment(.ColIndex("��Ŀ")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("��׼��ֵ")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("���ֱ�׼")) = flexAlignCenterCenter
        '���ص�Ԫ��
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("�ϼ�ID")) = 0
        .ColWidth(.ColIndex("����ID")) = 0
        .ColWidth(.ColIndex("���")) = 0
        '�������
        .ColWidth(.ColIndex("��Ŀ")) = 1500
        .ColWidth(.ColIndex("��׼��ֵ")) = 850
        .ColWidth(.ColIndex("ȱ������")) = 3700
        .ColWidth(.ColIndex("���ֱ�׼")) = 1100
        '�и�����
'        .RowHeightMin = 300
        '���������
'        .ColWidthMax = 7000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize .ColIndex("ȱ������")
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        'ѡ����ǰ����
        If m_lngOldRow > 0 And m_lngOldRow < i Then
            .Row = m_lngOldRow
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            fgMain_SelChange
        ElseIf fgMain.Tag = "" And i > 1 And .Rows > 1 Then
            m_lngOldRow = 1
            fgMain.Tag = "ѡ�е�һ��"
            .Row = 1
            .Col = 2
            .ShowCell m_lngOldRow, 2
            On Error Resume Next
            If .Visible = True Then .SetFocus
            fgMain_SelChange
        Else
            lblInfo = "������"
        End If
    End With
    
    Call DataUpdate
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʾ��¼����Ϣ
'==============================================================================
Private Sub SetMenu()
    On Error GoTo errH
    stbThis.Panels(2).Text = "�б��й���ʾ��" & fgMain.Rows - 1 & "�����ݡ�"
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������ĿID��ʾ����Ҫ��
'==============================================================================
Private Sub Show����Ҫ��(lngID As Long, ��Ŀ As String, ��׼��ֵ As String)
    Dim rs          As ADODB.Recordset
    
    On Error GoTo errH
    gstrSQL = "select ID,���� as ����Ҫ��,�ϼ�ID from �������ֱ�׼ Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
    
    If Not rs.EOF Then
        If m_lngOldSJID > 0 And m_lngOldSJID = lngID Then Exit Sub
        If IsNull(rs.Fields("����Ҫ��")) Then
                lblInfo = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
                lblInfo = lblInfo + vbCrLf
        Else
            If Len(rs.Fields("����Ҫ��")) > 0 Then
                lblInfo = "���ƣ�" + ��Ŀ + "  " + IIf(Len(Trim(��׼��ֵ)) = 0, "", "(" + ��׼��ֵ + ")")
                lblInfo = lblInfo + vbCrLf + rs.Fields("����Ҫ��")
            End If
        End If
    Else
        lblInfo.Caption = "������":
    End If
    m_lngOldSJID = m_lngCurSJID
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȡ���Ƿ񱣴沼��
'==============================================================================
Private Function GetPersonSet() As Boolean
    
    On Error GoTo errH
    
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� �˵����ܿ���
'==============================================================================
Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo errH
    
    Select Case Control.ID
        Case conMenu_Edit_NewKind           '���ӷ���
            Call ProgAdd
        Case conMenu_Edit_ModifyKind        '�޸ķ���
            Call ProgEdit
        Case conMenu_Edit_DeleteKind        'ɾ������
            Call ProgDel
        Case conMenu_Edit_Import            '���뷽��
            Call ProgImp
        Case conMenu_Edit_Select            'ѡ�÷���
            Call ProgSele
        Case conMenu_Edit_NewParent         '������Ŀ
            Call ItemAdd
        Case conMenu_Edit_Insert            '������Ŀ
            Call ItemInsert
        Case conMenu_Edit_ModifyParent      '�޸���Ŀ
            Call ItemEdit
        Case conMenu_Edit_DeleteParent      'ɾ����Ŀ
            Call ItemDel
        Case conMenu_Edit_NewItem           '���ӱ�׼
            Call StandardAdd
        Case conMenu_Edit_CopyNewItem       '�����׼
            Call StandardInsrt
        Case conMenu_Edit_Modify            '�޸ı�׼
            Call StandardEdit
        Case conMenu_Edit_Delete            'ɾ����׼
            Call StandardDel
        Case conMenu_View_Refresh           'ˢ������
            Call InitTreeView
        Case conMenu_File_Preview           'Ԥ��
            Call ItemPrint
        Case conMenu_File_Print             '��ӡ
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel             '�����&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_View_Forward           '��һ��
            With fgMain
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
        Case conMenu_View_Backward          '��һ��
 
            With fgMain
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    .ShowCell .Row, .Col
                End If
            End With
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '��ҵ���޹صĹ��ܣ������Ĺ���
                Call CommandBarExecutePublic(Control, Me, fgMain, "�������ֱ�׼ά��")
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �˵�Ȩ�޿���
'==============================================================================
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo errH
    
    With fgMain
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel
                Control.Enabled = ((fgMain.Rows > 1) And IsPrivs(mstrPrivs, "��ɾ��"))
            Case conMenu_Edit_NewKind           '���ӷ���
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_ModifyKind       '�޸ķ���
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_DeleteKind       'ɾ������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_Import           '���뷽��
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed And (fgMain.Rows <= 1)) And tvw����.Nodes.count > 0
            Case conMenu_Edit_Select           'ѡ�÷���
                Control.Enabled = (fgMain.Rows > 0) And tvw����.Nodes.count > 0
            Case conMenu_Edit_NewParent        '������Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_Insert           '������Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_ModifyParent     '�޸���Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_DeleteParent     'ɾ����Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_NewItem          '���ӱ�׼
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_CopyNewItem      '�����׼
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvw����.Nodes.count > 0
            Case conMenu_Edit_Modify           '�޸ı�׼
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (fgMain.Row > 0) Or (Control.Visible And Not mblnProgUsed And m_lngCurSJID <> 0)
            Case conMenu_Edit_Delete           'ɾ����׼
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed And m_lngCurSJID <> 0)
            Case conMenu_View_Forward
                Control.Enabled = (Control.Visible And fgMain.Row > 1)
            Case conMenu_View_Backward
                Control.Enabled = (Control.Visible And fgMain.Row + 1 < fgMain.Rows)
            Case conMenu_View_Refresh
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = Control.Visible
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
    
            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

'==============================================================================
'=���ܣ� ��ӡ ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo errH
    subPrint (mzlPrintModeS)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�������ַ���
'==============================================================================
Private Sub ProgAdd()
    Dim f As New frm���ַ����༭
    On Error GoTo errH
    f.ShowForm   '����
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '����б�
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ѡ�÷���
'==============================================================================
Private Sub ProgSele()
    Dim intIndex        As Long
    Dim bln��ʹ��       As Boolean
    
    On Error GoTo errH
    
    If m_lngCurFAID < 1 Then Exit Sub
    If MsgBox("ע�⣺���ְַ���ѡ����һ���ǳ����ص����飬ͨ����Ҫ������ģ�" & vbCrLf & "��ȷ��ѡ�ñ����ַ�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select count(*) from �������ֽ�� where ����ID=(select ID from �������ַ��� where ����='סԺ' and ѡ��=1)"
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    If rsTemp(0).Value > 0 Then
        'Ĭ��סԺ�����Ѿ�ʹ��
        If MsgBox("ע�⣺ϵͳĬ�����ְַ�����ʹ�õ��У��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    rsTemp.Close
    
    gstrSQL = "ZL_�������ַ���_ѡ��(" & CStr(m_lngCurFAID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call InitTreeView
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�޸����ַ���
'==============================================================================
Private Sub ProgEdit()
    Dim f               As New frm���ַ����༭
    Dim lng�ܷ�         As Double
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurFAID < 1 Then Exit Sub
    f.ShowForm m_lngCurFAID   '�޸ģ�����ID
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '����б�
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�������з���
'==============================================================================
Private Sub ProgImp()
    Dim lID             As Long     'ѡ�еķ���ID
    Dim lNewID          As Long
    Dim f               As New frmѡ�����ַ���
    Dim rs              As ADODB.Recordset
    Dim lng�ܷ�         As Double
    Dim rsTmp           As ADODB.Recordset
    Dim strT            As String

    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurFAID <= 0 Then Exit Sub
    
    f.FillCmbSelFA m_lngCurFAID
    f.Show 1
    lID = f.ID_From
    
    'ִ�е������������
    'ԴIDΪ��lID   Ŀ��IDΪ�� m_lngCurFAID
    gstrSQL = "Select ID, �ϼ�id, ����id, ����, ����, ��׼��ֵ, ȱ�ݵȼ�, ���ֵ�λ, �ϼ����, ���, �ж�����, ����ȼ�, ����Դ" & vbNewLine & _
                "From �������ֱ�׼" & vbNewLine & _
                "Where �ϼ�id Is Null And ����id = [1]" & vbNewLine & _
                "Order By �ϼ����, ���, ID"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lID)
    zlCommFun.ShowFlash "���Ժ�ϵͳ���ڵ������ַ�������", Me
    DoEvents
    
    On Error GoTo LL
    
    gcnOracle.BeginTrans
    Do While Not rs.EOF
        '�ҵ�����Ŀ�������Ŀ
        lNewID = zlDatabase.GetNextId("�������ֱ�׼")
        gstrSQL = "ZL_�������ֱ�׼_Insert" & _
            "(" & lNewID & "," & NVL(rs!�ϼ�ID, "NULL") & "," & m_lngCurFAID & _
            ",'" & NVL(rs!����) & "','" & NVL(rs!����) & "'," & NVL(rs!��׼��ֵ, "NULL") & ",'" & NVL(rs!ȱ�ݵȼ�) & _
            "','" & NVL(rs!���ֵ�λ) & "',0,'" & Replace(NVL(rs!�ж�����), "'", "''") & "','" & NVL(rs!����ȼ�) & "'," & rs!����Դ & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        '��һ�������¼���Ŀ��ѭ�����֮��
        gstrSQL = "Select ID, �ϼ�id, ����id, ����, ����, ��׼��ֵ, ȱ�ݵȼ�, ���ֵ�λ, �ϼ����, ���, �ж�����, ����ȼ�, ����Դ" & vbNewLine & _
                "From �������ֱ�׼" & vbNewLine & _
                "Where �ϼ�id = [1] And ����id = [2]" & vbNewLine & _
                "Order By �ϼ����, ���, ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(rs("ID")), lID)
        Do While Not rsTmp.EOF
            gstrSQL = "ZL_�������ֱ�׼_Insert" & _
                "(" & zlDatabase.GetNextId("�������ֱ�׼") & "," & lNewID & "," & m_lngCurFAID & _
                ",'" & NVL(rsTmp!����) & "','" & NVL(rsTmp!����) & "'," & NVL(rsTmp!��׼��ֵ, "NULL") & ",'" & NVL(rsTmp!ȱ�ݵȼ�) & _
                "','" & NVL(rsTmp!���ֵ�λ) & "',0,'" & Replace(NVL(rs!�ж�����), "'", "''") & "','" & NVL(rs!����ȼ�) & "'," & rs!����Դ & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            rsTmp.MoveNext
        Loop
        rs.MoveNext
    Loop
    
    'ˢ�½����
    gcnOracle.CommitTrans
    
    Call DataLoad
    zlCommFun.StopFlash
    Exit Sub
LL:
    gcnOracle.RollbackTrans
    zlCommFun.StopFlash
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ɾ�����ַ���
'==============================================================================
Private Sub ProgDel()
    Dim intIndex        As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurFAID < 1 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ������������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_�������ַ���_Delete(" & CStr(m_lngCurFAID) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call InitTreeView
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=������Ŀ
'==============================================================================
Private Sub ItemAdd()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    f.ShowForm "����", m_lngCurFAID
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=������Ŀ
'==============================================================================
Private Sub ItemInsert()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then 'Ϊ����������
        f.ShowForm "����", m_lngCurFAID, 0, m_lngCurID
    Else
        f.ShowForm "����", m_lngCurFAID, 0, m_lngCurSJID
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�޸���Ŀ
'==============================================================================
Private Sub ItemEdit()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If m_lngCurSJID < 1 Then
        f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurID
    Else
        f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurSJID
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ɾ����Ŀ
'==============================================================================
Private Sub ItemDel()
    
    Dim intIndex As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    
    If m_lngCurID < 1 Then Exit Sub
    
    If MsgBox("��ȷ��Ҫɾ������������Ŀ��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If m_lngCurSJID = 0 Then
        gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurID) & ",0)"
    Else
        gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurSJID) & ",0)"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=������׼
'==============================================================================
Private Sub StandardAdd()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then 'Ϊ����������
        f.ShowForm "����", m_lngCurFAID, m_lngCurID
    Else
        f.ShowForm "����", m_lngCurFAID, m_lngCurSJID
    End If
    
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�����׼
'==============================================================================
Private Sub StandardInsrt()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurSJID < 1 Then 'Ϊ����������
        f.ShowForm "����", m_lngCurFAID, m_lngCurID, m_lngCurSJID
    Else
        f.ShowForm "����", m_lngCurFAID, m_lngCurSJID, m_lngCurID
    End If
    
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�޸ı�׼
'==============================================================================
Private Sub StandardEdit()
    Dim f As New frm���ֱ�׼�༭
    On Error GoTo errH
    If ObjPtr(tvw����.SelectedItem) = 0 Then Exit Sub
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If fgMain.Col < 2 Then  'һ����Ŀ
        If m_lngCurSJID < 1 Then
            f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurID, Not mblnProgUsed
        Else
            f.ShowForm "�޸�", m_lngCurFAID, , m_lngCurSJID, Not mblnProgUsed
        End If
    Else                    '����Ŀ
        f.ShowForm "�޸�", m_lngCurFAID, m_lngCurSJID, m_lngCurID, Not mblnProgUsed
    End If
    Call DataUpdate
    If f.Moded Then
        Call DataLoad
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ɾ����׼
'==============================================================================
Private Sub StandardDel()
    Dim intIndex As Long
    
    On Error GoTo errH
    
    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If m_lngCurID < 1 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ���������ֱ�׼��", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    gstrSQL = "ZL_�������ֱ�׼_Delete(" & CStr(m_lngCurID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call DataLoad
    Call SetMenu
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=��λ�õ�����ѡ��
'==============================================================================
Private Sub txtLocation_GotFocus()
    On Error GoTo errH
    Call zlControl.TxtSelAll(txtLocation)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ٶ�λ
'==============================================================================
Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo errH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If mstrFindKey = "ȱ��" Then mstrFindKey = "ȱ������"
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = fgMain.Row + 1 To fgMain.Rows - 1
            If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To fgMain.Row
                If InStr(UCase(fgMain.TextMatrix(lngLoop, fgMain.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If fgMain.Rows > 1 And lngRow >= 1 Then fgMain.Row = lngRow
        
        Call LocationObj(txtLocation)
    End If
    If mstrFindKey = "ȱ������" Then mstrFindKey = "ȱ��"
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
