VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmAuditItem 
   Caption         =   "��������׼"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   Icon            =   "frmAuditItem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   11655
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000015&
      BorderStyle     =   0  'None
      Height          =   5520
      Index           =   2
      Left            =   315
      ScaleHeight     =   5520
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   480
      Width           =   3015
      Begin VB.PictureBox picTree 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1590
         Left            =   60
         ScaleHeight     =   1590
         ScaleWidth      =   2940
         TabIndex        =   14
         Top             =   240
         Width           =   2940
         Begin MSComctlLib.TreeView tvwAuditType 
            Height          =   1200
            Left            =   495
            TabIndex        =   16
            Top             =   420
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   2117
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   494
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ils16"
            Appearance      =   0
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "����׼"
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
            TabIndex        =   15
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox pic������Ϣ 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   45
         ScaleHeight     =   1695
         ScaleWidth      =   2790
         TabIndex        =   7
         Top             =   2565
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
            TabIndex        =   8
            Top             =   75
            Width           =   255
         End
         Begin VB.Label lbl�ֶ��� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ֶ���:"
            Height          =   195
            Left            =   225
            TabIndex        =   13
            Top             =   1035
            Width           =   2580
         End
         Begin VB.Label lbl�ܷ� 
            BackStyle       =   0  'Transparent
            Caption         =   "�ܷ�:"
            Height          =   195
            Left            =   225
            TabIndex        =   12
            Top             =   705
            Width           =   2580
         End
         Begin VB.Label lbl����ʱ�� 
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��:"
            Height          =   195
            Left            =   225
            TabIndex        =   11
            Top             =   1365
            Width           =   2580
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
            TabIndex        =   10
            Top             =   450
            Width           =   2580
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
            TabIndex        =   9
            Top             =   90
            Width           =   1095
         End
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   1
      Left            =   3240
      ScaleHeight     =   2865
      ScaleWidth      =   5880
      TabIndex        =   3
      Top             =   4215
      Width           =   5880
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   2025
         Left            =   690
         TabIndex        =   4
         Top             =   240
         Width           =   2700
         _Version        =   589884
         _ExtentX        =   4762
         _ExtentY        =   3572
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   4350
      ScaleHeight     =   2715
      ScaleWidth      =   5010
      TabIndex        =   2
      Top             =   825
      Width           =   5010
      Begin VSFlex8Ctl.VSFlexGrid vsfAuditItem 
         Height          =   4695
         Left            =   105
         TabIndex        =   6
         Top             =   150
         Width           =   6270
         _cx             =   11060
         _cy             =   8281
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.TextBox txtLocation 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7875
      TabIndex        =   0
      ToolTipText     =   "��ݼ���F3"
      Top             =   90
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   7320
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15690
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   26
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   2400
      Top             =   120
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
            Picture         =   "frmAuditItem.frx":08CA
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAuditItem.frx":171C
            Key             =   "RootSel"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgOpen 
      Height          =   225
      Left            =   2670
      Picture         =   "frmAuditItem.frx":1990
      Top             =   11235
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5535
      Picture         =   "frmAuditItem.frx":19E5
      Top             =   11250
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgBGBlue 
      Height          =   1530
      Left            =   210
      Picture         =   "frmAuditItem.frx":1A34
      Top             =   11190
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Image imgBG 
      Height          =   1695
      Left            =   3060
      Picture         =   "frmAuditItem.frx":1BF4
      Top             =   11205
      Visible         =   0   'False
      Width           =   2790
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3135
      Top             =   720
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmAuditItem.frx":1DB2
      Left            =   1185
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmAuditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long      '�����ؼ�����ˢ��
Private mstrPrivs               As String               'Ȩ�޴�
Private mobjFindKey             As CommandBarPopup      '��ѯ
Private mstrFindKey             As String               '��ѯ��
Private mlngModule              As Long                 'ģ���
Private mstrSaveKey             As String               '������ϴεķ���ѡ��ؼ���
Private mRsAuditItem            As ADODB.Recordset      '���ݼ�
Private mblnCheckAll            As Boolean              '�Ƿ���ʾ�¼�
Private zlCheck                 As New clsCheck         '�����
Private mzlPrintModeS           As gzlPrintModeS        '��ӡ
Private mstrSortID              As String               '����λ
Private mblnProgUsed            As Boolean              '�����Ƿ���ʹ��
Private mlngCurFAID             As Long                 '��ǰ����ID
Private mintTypeID              As Integer              '�����������޸ġ�ɾ��ʱ��ID
Private mintItemID              As Integer              '��Ŀ�������޸ġ�ɾ��ʱ��ID
Private mDataChange             As Boolean              '�����Ƿ񱻱༭��
Private menuEditMode            As �༭ģʽ
Private mblPopType              As Boolean
Private mcbrPopupBarType        As CommandBar           '�������ڡ����ࡿ
Private mcbrPopupBarItem        As CommandBar           '�������ڡ���Ŀ��
Dim cbrPopupItem                As CommandBarControl    '������
Private Const con_vsfField = "/*+ rule */ '' as ͼ��,a.id, a.����id,a.����,a.����,a.����,a.��ֵ,a.����,b.���� as ����,decode(a.���ö���,1,'סԺҽ��',2,'סԺ����',3,'������',4,'�����¼',5,'��ҳ��¼',6,'ҽ������',7,'����֤��',8,'֪���ļ�','δ����') as ���ö���,a.˵��,a.�������,���ö��� as ���ñ���,�ļ�ID,���û���,����Դ"
Private Const conFieldFiles = "Select /*+ rule */ a.id as �ļ�ID,a.��� as �ļ�����,a.���� as �ļ�����,a.˵�� as �ļ�˵��" & vbCrLf & _
                         "from �����ļ��б� A, Table (Cast(f_Str2List([1])  As zlTools.t_StrList)) B " & vbCrLf & _
                         "where /*+ rule */a.id = b.COLUMN_VALUE And a.���� = [2]"
Private Const conEmrField = "Select /*+ Rule*/ Rawtohex(b.Id) As �ļ�id, b.Code As �ļ�����, b.Title As �ļ�����, b.Note As �ļ�˵��" & vbNewLine & _
                        "From (Select Hextoraw(Column_Value) As ID From Table(Zlcommunal.f_Str2list(:p0, ','))) A, Antetype_List B" & vbNewLine & _
                        "Where Hextoraw(a.Id) = b.Id And b.Kind = :p1" & vbNewLine & _
                        "Order By �ļ�����"
Public Enum �༭ģʽ
    ��� = 0
    ���� = 1
    �޸� = 2
    �������� = 3
End Enum

'��ӡģʽ
Enum gzlPrintModeS
    zlPrint = 1         '��ӡ
    zlView = 2          '�鿴
    zlExcel = 3         '�����Excel
End Enum

'���ڵ㶨λ
Dim nod                         As Node
Dim i                           As Long
Dim FirstKey                    As String
Dim v                           As Variant

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
    
    On Error GoTo ErrH

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
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Show, "���뷽��(&P)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Hide, "��������(&C)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewParent, "����(&A)...", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_ModifyParent, "�޸�(&E)...")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_DeleteParent, "ɾ��(&X)")
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_NewItem, "������Ŀ(&N)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Modify, "�޸���Ŀ(&M)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_CopyNewItem, "��������(&C)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Delete, "ɾ����Ŀ(&D)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Save, "�������(&S)", True)
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ������(&R)")
    
    '------------------------------------------------------------------------------------------------------------------
    '�鿴
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    Set objPopup = NewCommandBar(objMenu, xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)")
    Set objControl = NewCommandBar(objPopup, xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_ShowAll, "�����¼�(&A)", True)
    
    Set objControl = NewCommandBar(objMenu, xtpControlButton, conMenu_View_Find, "����(&F)...")
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
    
    '���˵��Ҳ�Ĳ���
    '------------------------------------------------------------------------------------------------------------------
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    
    mstrFindKey = Trim(GetPara("��λ����", mlngModule, "����", True))
    If mstrFindKey = "" Then mstrFindKey = "����"

    Set mobjFindKey = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_View_LocationItem, mstrFindKey)
    mobjFindKey.IconId = conMenu_View_Find
    mobjFindKey.Flags = xtpFlagRightAlign
    mobjFindKey.STYLE = xtpButtonIconAndCaption
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&1.����", , , "����")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&2.����", , , "����")
    Set objControl = NewCommandBar(mobjFindKey, xtpControlButton, conMenu_View_LocationItem, "&3.����", , , "����")
    
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
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Save, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��")
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "����", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "�˳�")
    
    '------------------------------------------------------------------------------------------------------------------
    '����Ŀ����:���������������Ѵ���

    With cbsMain.KeyBindings
        
        .Add 0, vbKeyF5, conMenu_View_Refresh               'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help                  '����
        
        .Add FCONTROL, vbKeyF, conMenu_View_Find            '����
        .Add FCONTROL, vbKeyP, conMenu_File_Print           '��ӡ
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem         '����
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
    Set mcbrPopupBarType = cbsMain.Add("��������˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_NewKind, "��������(&N)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_ModifyKind, "�޸ķ���(&F)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_DeleteKind, "ɾ������(&L)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_View_Show, "���뷽��(&P)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_View_Hide, "��������(&C)")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_Select, "ѡ�÷���(&S)")
    
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_NewParent, "����(&A)...", True)
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_ModifyParent, "�޸�(&E)...")
    Set cbrPopupItem = mcbrPopupBarType.Controls.Add(xtpControlButton, conMenu_Edit_DeleteParent, "ɾ��(&X)")
    
    Set mcbrPopupBarItem = cbsMain.Add("������Ŀ�˵�", xtpBarPopup)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_NewItem, "����(&N)", True)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_CopyNewItem, "����(&C)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Save, "����(&S)", True)
    Set cbrPopupItem = mcbrPopupBarItem.Controls.Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "ȡ��(&R)")
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ���򻮷�
'==============================================================================
Private Sub InitDockPannel()
    Dim objPane As Pane

    On Error GoTo ErrH
    
    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "����"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 400, 100, DockRightOf, Nothing)
    objPane.Title = "�¼�"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 300, 100, DockBottomOf, objPane)
    objPane.Title = "��ϸ"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �ؼ���ʼ��
'==============================================================================
Private Sub InitControl()
    
    On Error GoTo ErrH
    
    Call InitVsflexGrid
    Call InitCommandBar
    Call InitDockPannel
    Call InitTabControl
    Call InitTreeView
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼ������ VsflexGrid
'==============================================================================
Private Sub InitVsflexGrid()
    Dim strField        As String
    Dim strFieldWidth   As String
    Dim varField        As Variant
    Dim varFieldWidth   As Variant
    Dim i               As Integer
    On Error GoTo ErrH
    vsfAuditItem.FocusRect = flexFocusNone
    vsfAuditItem.ExtendLastCol = True
    vsfAuditItem.ExplorerBar = flexExSortShowAndMove
    vsfAuditItem.AutoResize = False
    gstrSQL = "" & _
        "Select " & con_vsfField & vbCrLf & _
        "From �������Ŀ¼ a,(SELECT /*+ rule */ id,���� FROM ���������� START WITH id=[1] CONNECT BY PRIOR ID = �ϼ�ID)b " & vbCrLf & _
        "Where a.����id = b.ID and 1=0"
    Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, -1)
    Set vsfAuditItem.DataSource = mRsAuditItem
    With vsfAuditItem
        .ColWidth(0) = 250
        .MergeCol(.ColIndex("����id")) = True
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(.ColIndex("ͼ��")) = 450
        .ColWidth(.ColIndex("���ö���")) = 2000
        .ColWidth(.ColIndex("��ֵ")) = 500
        .ColWidthMin = 450
        
'        .FrozenCols = 3
        If GetPersonSet Then
            'ʹ�ø��Ի����á����ѱ���ĸ�ʽ��
            strField = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "����", "")
            strFieldWidth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & Me.Name & "\VSFlexGrid", vsfAuditItem.Name & "���", "")
            varField = Split(strField, ",")
            varFieldWidth = Split(strFieldWidth, ",")
            For i = 0 To UBound(varField)
                If varField(i) <> "" And Val(varFieldWidth(i)) <> 0 Then
                    .ColPosition(.ColIndex(varField(i))) = i
                    .ColWidth(i) = Val(varFieldWidth(i))
                End If
            Next
        End If
        .ColWidth(.ColIndex("ID")) = 0: .ColHidden(.ColIndex("ID")) = True
        .ColWidth(.ColIndex("����id")) = 0: .ColHidden(.ColIndex("����id")) = True
        .ColWidth(.ColIndex("���ñ���")) = 0: .ColHidden(.ColIndex("���ñ���")) = True
        .ColWidth(.ColIndex("�������")) = 0: .ColHidden(.ColIndex("�������")) = True
        .ColWidth(.ColIndex("�ļ�ID")) = 0: .ColHidden(.ColIndex("�ļ�ID")) = True
        .ColWidth(.ColIndex("���û���")) = 0: .ColHidden(.ColIndex("���û���")) = True
        .ColWidth(.ColIndex("����")) = 0: .ColHidden(.ColIndex("����")) = True
        .ColWidth(.ColIndex("����Դ")) = 0: .ColHidden(.ColIndex("����Դ")) = True
    End With
    DoEvents
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����������
'==============================================================================
Private Sub InitTreeView()
    Dim rsTree      As ADODB.Recordset
    Dim intStartid As Integer
    On Error GoTo ErrH

    'Tree�ĳ�ʼ��
    Set tvwAuditType.ImageList = GetImageList(16)
    tvwAuditType.Nodes.Clear
    
    gstrSQL = "Select ID,����,����ʱ�� From ������鷽��"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    Do Until rsTree.EOF
        If zlCommFun.NVL(rsTree!����ʱ��) <> "" Then
            intStartid = rsTree!ID
        End If
        Set nod = tvwAuditType.Nodes.Add(, , "Root" & rsTree!ID, zlCommFun.NVL(rsTree!����, "Ĭ�Ϸ���"), 20, 20)
        nod.Expanded = True
            
        rsTree.MoveNext
    Loop
    
'    '��Ӹ��ڵ�
'    Set nod = tvwAuditType.Nodes.Add(, , "Root", "����", 20, 20)
'    nod.Expanded = True

    gstrSQL = "SELECT /*+ rule */ id,�ϼ�ID,����ID,����,���� FROM ���������� START WITH �ϼ�ID is NULL CONNECT BY PRIOR ID = �ϼ�ID"
    Set rsTree = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    If rsTree.RecordCount = 0 Then Exit Sub
    rsTree.Sort = "����"
    i = 1
    Do Until rsTree.EOF
        '����ӽڵ�
        Set nod = tvwAuditType.Nodes.Add(IIf("" & rsTree("�ϼ�ID") = "", "Root" & rsTree("����ID"), "A" & rsTree("�ϼ�ID")), tvwChild, "A" & rsTree("ID"), "��" + "" & rsTree("����") + "��" + "" & rsTree("����"), 23, 24)
        If i = 1 Then FirstKey = nod.Key
        If FirstKey = nod.Key Then i = 2
        If FirstKey = "" And i = 1 Then FirstKey = nod.Key: i = 2
        rsTree.MoveNext
    Loop
    FirstKey = "A" & mintTypeID
    For Each v In tvwAuditType.Nodes
        If v.Key = FirstKey Then
            '����ѡ��
            v.Selected = True
            v.EnsureVisible
        End If
    Next
    If tvwAuditType.SelectedItem Is Nothing Then
        tvwAuditType.Nodes("Root" & intStartid).Selected = True
        tvwAuditType.Nodes("Root" & intStartid).Bold = True
        tvwAuditType.Nodes("Root" & intStartid).Tag = 1
    End If
    DoEvents
    tvwAuditType_NodeClick tvwAuditType.SelectedItem
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʼTab�ؼ�
'==============================================================================
Private Function InitTabControl() As Boolean
    
    On Error GoTo ErrH
    
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With
        Set .Icons = zlCommFun.GetPubIcons
        .InsertItem 0, " ������Ϣ ", frmAuditItemEdit.hWnd, 0
        .Item(0).Selected = True
    End With

    InitTabControl = True

    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ� λ������
'==============================================================================
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub


'==============================================================================
'=���ܣ� ����ָ�
'==============================================================================
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo ErrH
    
    Select Case Item.ID
        Case 1
            Item.Handle = picPane(2).hWnd
        Case 2
            Item.Handle = picPane(0).hWnd
        Case 3
            Item.Handle = picPane(1).hWnd
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �����λ��¼ vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow      As Long
    On Error GoTo ErrH
    lngRow = vsfAuditItem.FindRow(mstrSortID, -1, vsfAuditItem.ColIndex("ID"), False, True)
    If lngRow > 0 Then vsfAuditItem.Row = lngRow
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ĳ�в����ƶ�λ�� vsfAuditItem[ͼ��]
'==============================================================================
Private Sub vsfAuditItem_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    If Col = vsfAuditItem.ColIndex("ͼ��") Then
        Position = -1
    Else
        If Position <= vsfAuditItem.ColIndex("ͼ��") Then Position = Col
    End If
End Sub

'==============================================================================
'=���ܣ� ����ǰ��¼ID vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_BeforeSort(ByVal Col As Long, Order As Integer)
    On Error GoTo ErrH
    mstrSortID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ĳ�в����϶���С vsfAuditItem[ͼ��]
'==============================================================================
Private Sub vsfAuditItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfAuditItem.ColIndex("ͼ��") Then Cancel = True
End Sub

'==============================================================================
'=���ܣ� ˫������޸Ĺ��� vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_DblClick()
    On Error GoTo ErrH
    If vsfAuditItem.MouseRow <= 0 Then Exit Sub
    Call ExecuteCommand("�޸���Ŀ")
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �Ҽ��˵� vsfAuditItem
'==============================================================================
Private Sub vsfAuditItem_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
        
            Call SendLMouseButton(vsfAuditItem.hWnd, X, Y)

            mcbrPopupBarItem.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
 
'==============================================================================
'=���ܣ����б任ʱ
'==============================================================================
Private Sub vsfAuditItem_RowColChange()
    Dim rsTemp          As ADODB.Recordset
    Dim varPos          As Variant, strReturn As String
    On Error GoTo ErrH
    DoEvents
    If vsfAuditItem.Rows = 1 Then
        With frmAuditItemEdit
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = ""
            .txtName.Text = ""
            .txtCode.Text = ""
            .txtMnemonicCode.Text = ""
            .cboUsed.ListIndex = -1
            .cboLink.ListIndex = -1
            .txtDescription.Text = ""
            .txtAudit_NotCheck.Text = ""
            .txtNumValue = ""
            .CboPalValue.ListIndex = -1
            .blnProgUsed = False
            Set .vsfFiles.DataSource = Nothing
        End With
        stbThis.Panels(2) = "��ǰ��ʾ�� 0 ����Ŀ��"
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    If vsfAuditItem.ColIndex("ID") <= 0 Then Exit Sub
    If Val(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))) <= 0 Then
        frmAuditItemEdit.vsfFiles.Rows = 1
        Exit Sub
    End If
    With frmAuditItemEdit
        
        .txtTypeID.Tag = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        
        gstrSQL = "select /*+ rule */id,�ϼ�ID,����,���� from ���������� a Where a.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, CStr(Val("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����ID")))))
        If Not zlCheck.Connection_ChkRsState(rsTemp) Then
            .txtTypeID.Tag = "" & rsTemp!ID
            .txtTypeID.Text = "[" + rsTemp!���� + "]" & rsTemp!����
        Else
            .txtTypeID.Tag = "-1"
            .txtTypeID.Text = "[ȫ��]����"
        End If
        
        If vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����Դ")) = "0" Then
            .optSource(0).Value = True:             .optSource(1).Value = False
        Else
            .optSource(0).Value = False:             .optSource(1).Value = True
        End If
        .txtName.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .txtCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .txtMnemonicCode.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        .cboUsed.ListIndex = zlCheck.Cmb_EditIndex(.cboUsed, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("���ñ���")))
        .cboLink.ListIndex = zlCheck.Cmb_EditIndex(.cboLink, "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("���û���")))
        .txtDescription.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("˵��"))
        .txtAudit_NotCheck.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("�������"))
        .txtFileID.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("�ļ�ID"))
        .txtNumValue.Text = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("��ֵ"))
        .CboPalValue.ListIndex = IIf(vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����")) = "", 0, vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����")))
        
        .blnProgUsed = mblnProgUsed
        If .optSource(0).Value Then
            gstrSQL = conFieldFiles
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, .txtFileID.Text, AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 0))
        Else
            gstrSQL = conEmrField
            strReturn = gobjEmr.OpenSQLRecordset(gstrSQL, .txtFileID.Text & "^" & DbType.T_String & "^p0|" & AuditFileTran(zlCheck.Cmb_ID(.cboUsed), 1) & "^" & DbType.T_String & "^p1", rsTemp)
            If strReturn <> "" Then
                zlCheck.Msg_OK strReturn
                Exit Sub
            End If
        End If
        Set .vsfFiles.DataSource = rsTemp
    End With
    stbThis.Panels(2) = "��ǰ��ʾ�� " & vsfAuditItem.Rows - 1 & " ����Ŀ��"
    varPos = zlCheck.Connection_GetBookMark(mRsAuditItem, "ID=" & CStr("" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))))
    DoEvents
    If Not IsNull(varPos) Then
        If Val(varPos) > 0 Then mRsAuditItem.Bookmark = varPos
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������ݼ��� vsfAuditItem
'==============================================================================
Private Sub DataAuditItem(Optional strWhere As String)
    Dim strKey      As String
    Dim i           As Long
    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    If strWhere = "" Then
        If ObjPtr(tvwAuditType.SelectedItem) = 0 Then Exit Sub
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
            Exit Sub
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            mlngCurFAID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            If tvwAuditType.SelectedItem.Tag = "1" Then
                mblnProgUsed = True
            Else
                mblnProgUsed = False
            End If
        Else
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend
            
            If InStrRev(nTmpNode.Key, "Root") > 0 Then
                mlngCurFAID = Replace(nTmpNode.Key, "Root", "")
                If nTmpNode.Tag = "1" Then
                    mblnProgUsed = True
                Else
                    mblnProgUsed = False
                End If
            End If
        End If
        
        
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            strKey = Mid(tvwAuditType.SelectedItem.Key, 5)
        Else
            strKey = Mid(tvwAuditType.SelectedItem.Key, 2)
        End If
        
        If mblnCheckAll Then
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From �������Ŀ¼ a,(SELECT /*+ rule */ id,���� FROM ���������� START WITH id=[1] CONNECT BY PRIOR ID = �ϼ�ID)b " & vbCrLf & _
                    "Where a.����id = b.ID"
        Else
            gstrSQL = "" & _
                    "Select " & con_vsfField & vbCrLf & _
                    "From �������Ŀ¼ a,���������� b" & vbCrLf & _
                    "Where a.����id = b.ID and a.����id=[1]"
        End If
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name, strKey)
    Else
        gstrSQL = "" & _
                "Select " & con_vsfField & vbCrLf & _
                "From �������Ŀ¼ a,���������� b" & vbCrLf & _
                "Where a.����id = b.ID And" & vbCrLf & strWhere
        Set mRsAuditItem = zlDatabase.OpenSQLRecord(gstrSQL, Me.Name)
    End If
    Set vsfAuditItem.DataSource = mRsAuditItem
       
    With vsfAuditItem
        If .Rows > 1 Then
            For i = .FixedRows To .Rows - 1
                .Cell(flexcpPictureAlignment, i, .ColIndex("ͼ��")) = flexPicAlignCenterCenter
                Select Case .Cell(flexcpText, i, .ColIndex("���ñ���"))
                    Case "1"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(2).Picture
                    Case "2"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(15).Picture
                    Case "3"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(16).Picture
                    Case "4"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(17).Picture
                    Case "5"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(18).Picture
                    Case "6"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(6).Picture
                    Case "7"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(3).Picture
                    Case "8"
                        .Cell(flexcpPicture, i, .ColIndex("ͼ��")) = frmPubResource.ils16.ListImages(20).Picture
                End Select
            Next i
            .Row = 1
        End If
    End With
    Call DataUpdate
    Call vsfAuditItem_RowColChange
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ӷ������� TypeInsert
'==============================================================================
Public Sub TypeInsert()

    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    
    With frmAuditItemTypeEdit
        .EditMode = ����
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            .strID = "-1"
            .lngProjectID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            .strProjectName = tvwAuditType.SelectedItem.Text
        Else
            .strID = Mid(tvwAuditType.SelectedItem.Key, 2)
            
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend
            
            .lngProjectID = Replace(nTmpNode.Key, "Root", "")
            .strProjectName = nTmpNode.Text
        End If
        
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeEdit = Nothing: Exit Sub
        mintTypeID = .strID
    End With
    Set frmAuditItemTypeEdit = Nothing
    'ˢ����
    Call InitTreeView
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �޸ķ������� TypeUpdate
'==============================================================================
Public Sub TypeUpdate()

    Dim nTmpNode As Node
    
    On Error GoTo ErrH
    If InStrRev(tvwAuditType.SelectedItem.Key, "Root") > 0 Then
        MsgBox "��Ŀ¼,�����޸ġ�", vbInformation, "������ʾ"
        Exit Sub
    End If
    With frmAuditItemTypeEdit
        .EditMode = �޸�
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            .strID = Mid(tvwAuditType.SelectedItem.Key, 5)
            .lngProjectID = Replace(tvwAuditType.SelectedItem.Key, "Root", "")
            .strProjectName = tvwAuditType.SelectedItem.Text
        Else
            .strID = Mid(tvwAuditType.SelectedItem.Key, 2)
            
            Set nTmpNode = tvwAuditType.SelectedItem
            While Not nTmpNode.Parent Is Nothing
                Set nTmpNode = nTmpNode.Parent
            Wend

            .lngProjectID = Replace(nTmpNode.Key, "Root", "")
            .strProjectName = nTmpNode.Text
        End If
        
        .Show vbModal
        If .blnCancel Then Set frmAuditItemTypeEdit = Nothing: Exit Sub
    End With
    Set frmAuditItemTypeEdit = Nothing
    If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
        mintTypeID = Mid(tvwAuditType.SelectedItem.Key, 5)
    Else
        mintTypeID = Mid(tvwAuditType.SelectedItem.Key, 2)
    End If

    'ˢ����
    Call InitTreeView
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ӷ������� TypeUpdate
'==============================================================================
Public Sub TypeDelete()
    Dim strKey      As String
    
    On Error GoTo ErrH
    If InStrRev(tvwAuditType.SelectedItem.Key, "Root") > 0 Then
        MsgBox "��Ŀ¼������ɾ����", vbInformation, "������ʾ"
        Exit Sub
    Else
        If MsgBox("ȷ��ɾ������""" & tvwAuditType.SelectedItem.Text & """����������Ŀ��", vbOKCancel + vbDefaultButton2, "������ʾ") <> vbOK Then Exit Sub
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            mintTypeID = Val(Mid(tvwAuditType.SelectedItem.Key, 5))
        Else
            mintTypeID = Val(Mid(tvwAuditType.SelectedItem.Key, 2))
        End If
    End If
    frmAuditItemTypeEdit.strID = CStr(mintTypeID)
    frmAuditItemTypeEdit.AuditItemTypeDelete
    Set frmAuditItemTypeEdit = Nothing
    tvwAuditType.Nodes.Remove tvwAuditType.SelectedItem.Index
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ӡ ItemPrint
'==============================================================================
Private Sub ItemPrint()
    On Error GoTo ErrH
    subPrint (mzlPrintModeS)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub subPrint(ByVal bytMode As Byte)
    Dim lngLoop         As Long
    Dim objControl      As Object
    Dim objPrint        As New zlPrint1Grd
    Dim objAppRow       As zlTabAppRow
    
    If vsfAuditItem Is Nothing Then Exit Sub
    LockWindowUpdate vsfAuditItem.hWnd
    vsfAuditItem.ColHidden(vsfAuditItem.ColIndex("ͼ��")) = True
    Call SearchPrintData(vsfAuditItem, frmPubResource.msfPrint)
    vsfAuditItem.ColHidden(vsfAuditItem.ColIndex("ͼ��")) = False
    LockWindowUpdate 0
    '���ô�ӡ��������
    Set objPrint.Body = frmPubResource.msfPrint
    objPrint.Title.Text = Me.Caption
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ�ˣ�" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ�䣺" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    zlPrintOrView1Grd objPrint, bytMode
End Sub

'==============================================================================
'=���ܣ� ��ѯ ItemFind
'==============================================================================
Private Sub ItemFind()
    Dim strWhere        As String
    On Error GoTo ErrH
    With frmAuditItemFind
        .Show vbModal
        If .blnCancel Then Set frmAuditItemFind = Nothing: Exit Sub
        strWhere = .strWhere
    End With
    Set frmAuditItemFind = Nothing
    DataAuditItem (strWhere)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrH
    mstrPrivs = UserInfo.ģ��Ȩ��
    mlngModule = ParamInfo.ģ���
    Call ExecuteCommand("��ʼ�ؼ�")
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, gblnShowInTaskBar)
    Call zlDatabase.ShowReportMenu(Me, ParamInfo.ϵͳ��, ParamInfo.ģ���, UserInfo.ģ��Ȩ��)
        
    menuEditMode = ���
    Call ExecuteCommand("��ȡ���������Ŀ") '    mblnProgUsed = False
    frmAuditItemEdit.WinLock
    picFAXX.Picture = imgClose.Picture
    
    Call ExecuteCommand("��ע���")
        
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrH
    menuEditMode = ���
    Call SetPaneRange(dkpMain, 1, 100, 60, 450, Me.ScaleHeight)
    Call SetPaneRange(dkpMain, 3, 15, 300, Me.ScaleWidth, 350)
    
    dkpMain.RecalcLayout
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error GoTo ErrH
    
    Call ExecuteCommand("дע���")
    Call SaveWinState(Me, App.ProductName)
    Set mobjFindKey = Nothing
    Set frmAuditItemEdit = Nothing
    
    SaveFlexState vsfAuditItem, Me.Name
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next

    Select Case Index
    Case 0
        vsfAuditItem.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    Case 1
        tbcPage.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 2
'        tvwAuditType.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
        
        On Error Resume Next
        pic������Ϣ.Move 0, picPane(2).ScaleHeight - pic������Ϣ.Height, picPane(2).ScaleWidth
        With picTree
            .Move 0, 0, pic������Ϣ.Width, picPane(2).Height - pic������Ϣ.Height
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, picTree.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, picTree.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, picTree.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, picTree.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, picTree.ScaleHeight - Screen.TwipsPerPixelY, picTree.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        
        tvwAuditType.Move Screen.TwipsPerPixelX * 4, 390, Abs(picTree.ScaleWidth - 8 * Screen.TwipsPerPixelX), Abs(picTree.ScaleHeight - 390 - Screen.TwipsPerPixelY * 4)
        With pic������Ϣ
            .Cls
            .PaintPicture imgBGBlue.Picture, 0, 0, pic������Ϣ.Width, 360, 0, 0, imgBGBlue.Width, 360
            .PaintPicture imgBGBlue.Picture, 0, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, 0, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, pic������Ϣ.ScaleWidth - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, pic������Ϣ.Height - 360, imgBGBlue.Width - Screen.TwipsPerPixelX, 360, Screen.TwipsPerPixelX, imgBGBlue.Height - 360
            .PaintPicture imgBGBlue.Picture, 0, pic������Ϣ.ScaleHeight - Screen.TwipsPerPixelY, pic������Ϣ.Width, Screen.TwipsPerPixelY, 0, imgBGBlue.Height - Screen.TwipsPerPixelY, imgBGBlue.Width, Screen.TwipsPerPixelY
        End With
        picFAXX.Move pic������Ϣ.ScaleWidth - picFAXX.Width - 80
        Refresh
        
    End Select
End Sub

Private Sub tvwAuditType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH

    Select Case Button
        Case 2          '�����˵�����
        
            Call SendLMouseButton(tvwAuditType.hWnd, X, Y)

            mcbrPopupBarType.ShowPopup
    End Select
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tvwAuditType_NodeClick(ByVal Node As MSComctlLib.Node)
        
    If mstrSaveKey = Node.Key Then Exit Sub
    If Left(Node.Key, 4) = "Root" Then
        vsfAuditItem.Rows = 1
        mstrSaveKey = Node.Key
        mlngCurFAID = Replace(mstrSaveKey, "Root", "")
        If Node.Tag = "1" Then
            mblnProgUsed = True
        Else
            mblnProgUsed = False
        End If
        Call DataUpdate
        Exit Sub
    End If
    mstrSaveKey = Node.Key
    
    Call ExecuteCommand("��ȡ���������Ŀ")
    
End Sub

Private Sub txtLocation_GotFocus()
    Call zlControl.TxtSelAll(txtLocation)
End Sub

Private Sub txtLocation_KeyPress(KeyAscii As Integer)
    Dim lngRow      As Long
    Dim intCol      As Integer
    Dim bytMatch    As Byte
    Dim lngLoop     As Long
    
    On Error GoTo ErrH
    
    lngRow = 0
    If txtLocation.Locked Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        '��ȡ���ڵ�ǰ�еļ�¼����
        For lngLoop = vsfAuditItem.Row + 1 To vsfAuditItem.Rows - 1
            If InStr(UCase(vsfAuditItem.TextMatrix(lngLoop, vsfAuditItem.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                lngRow = lngLoop
                Exit For
            End If
        Next
        '��ȡС�ڵ�ǰ�еļ�¼����
        If lngRow = 0 Then
            For lngLoop = 0 To vsfAuditItem.Row
                If InStr(UCase(vsfAuditItem.TextMatrix(lngLoop, vsfAuditItem.ColIndex(mstrFindKey))), UCase(txtLocation.Text)) > 0 Then
                    lngRow = lngLoop
                    Exit For
                End If
            Next
        End If
        If vsfAuditItem.Rows > 1 And lngRow >= 1 Then vsfAuditItem.Row = lngRow
        
        Call LocationObj(txtLocation)
    End If
    
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Private Function ExecuteCommand(ByVal strCommand As String) As Boolean
    
    On Error GoTo ErrH
    Dim strF As String
    Dim strTvwName As String
    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        Call InitControl
    Case "��ȡ���������Ŀ"
        Call DataAuditItem
    Case "���ӷ���"
        Call TypeInsert
    Case "�޸ķ���"
        Call TypeUpdate
    Case "ɾ������"
        Call TypeDelete
    Case "������Ŀ"
        If Left(tvwAuditType.SelectedItem.Key, 4) = "Root" Then
            frmAuditItemEdit.lngItemTypeID = Mid(tvwAuditType.SelectedItem.Key, 5)
        Else
            frmAuditItemEdit.lngItemTypeID = Mid(tvwAuditType.SelectedItem.Key, 2)
        End If
        Call frmAuditItemEdit.ItemInsert
    Case "�޸���Ŀ"
        frmAuditItemEdit.lngItemID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        Call frmAuditItemEdit.ItemUpdate
    Case "ɾ����Ŀ"
        Dim varPos      As Variant
        frmAuditItemEdit.lngItemID = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
        frmAuditItemEdit.strItemCode = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        frmAuditItemEdit.strItemName = "" & vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("����"))
        
        Call frmAuditItemEdit.ItemDelete
        
        
        varPos = vsfAuditItem.Row
        Call DataAuditItem
        Call vsfAuditItem_RowColChange
        If varPos <= vsfAuditItem.Rows - 1 Then vsfAuditItem.Row = varPos
        
    Case "������Ŀ"
        If frmAuditItemEdit.ItemSave Then: Exit Function
        mstrSaveKey = "A" & frmAuditItemEdit.lngItemTypeID
        Dim lngRow      As Long
        '����ѡ�����
        FirstKey = "A" & CStr(frmAuditItemEdit.lngItemTypeID)
        For Each v In tvwAuditType.Nodes
            If v.Key = FirstKey Then
                '����ѡ��
                v.Selected = True
                v.EnsureVisible
            End If
        Next
        frmAuditItemEdit.WinLock
        Call DataAuditItem
        DoEvents
        lngRow = vsfAuditItem.FindRow(CStr(frmAuditItemEdit.lngItemID), -1, vsfAuditItem.ColIndex("ID"), False, True)
        If lngRow > 0 Then vsfAuditItem.Row = lngRow

    Case "ȡ����Ŀ"
        If frmAuditItemEdit.ItemCancel Then Exit Function
        
        Call vsfAuditItem_RowColChange
        
    Case "����������Ŀ"
        Call frmAuditItemEdit.ItemCopy
    Case "���뷽��"
         '��XML�ļ�����
        Dim strXML As String
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        
        On Error GoTo err1
        dlgThis.ShowOpen
        
err1:
        If Err.Number = cdlCancel Then
           Err.Clear
           strXML = ""
           Exit Function
        End If
        
        strXML = dlgThis.FileName
        If gobjFSO.FileExists(strXML) Then
            If ImportFromXMLFile(Me.tvwAuditType, strXML) Then
                'ˢ��
                Call InitTreeView
            End If
        End If
        
    Case "��������"
        If ObjPtr(tvwAuditType.SelectedItem) = 0 Then Exit Function
        strTvwName = tvwAuditType.SelectedItem.Text
        
        dlgThis.FileName = "�������_" & strTvwName & "_����.xml"
        dlgThis.Filter = "*.XML|*.xml|*.*|*.*"
        dlgThis.CancelError = True
        On Error Resume Next
        dlgThis.ShowSave
        If Err.Number <> 0 Then Err.Clear: Exit Function
        strF = dlgThis.FileName
        On Error GoTo ErrH
        If gobjFSO.FileExists(strF) Then
            DoEvents
            If MsgBox("���ļ��Ѿ����ڣ��Ƿ񸲸ǣ�", vbOKCancel + vbQuestion, gstrSysName) = vbCancel Then Exit Function
        End If
        
        If ExportToXMLFile(tvwAuditType, strF) Then
            DoEvents
            MsgBox "�ɹ�����ΪXML�ļ���" & vbCrLf & "�ļ���:" & strF, vbOKOnly + vbInformation, gstrSysName
        End If
        
    Case "ǰһ��"
        With vsfAuditItem
            If .Row > 1 Then
                .Row = .Row - 1
                .ShowCell .Row, .Col
            End If
        End With
    Case "��һ��"
        With vsfAuditItem
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
                .ShowCell .Row, .Col
            End If
        End With
    Case "��ע���"
        mblnCheckAll = (Val(GetPara("�����¼�", mlngModule, "0", False)) = 1)
        If GetPersonSet Then
            'ʹ�ø��Ի�����
            dkpMain.LoadStateFromString GetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, "")
            mstrFindKey = Trim(GetPara("��λ����", mlngModule, "����", True))
            Call RestoreWinState(Me, App.ProductName)
        End If
    Case "дע���"
        'ʹ�ø��Ի�����
        Call SetPara("��λ����", mstrFindKey, mlngModule)
        Call SetPara("�����¼�", IIf(mblnCheckAll, 1, 0), mlngModule)
        Call SetRegister(˽��ģ��, Me.Name & "\��������\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
        Call SaveWinState(Me, App.ProductName)
    End Select
    ExecuteCommand = True
    
    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPersonSet() As Boolean
    
    On Error GoTo ErrH
    GetPersonSet = False
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then GetPersonSet = True

    Exit Function
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnNewCancel        As Boolean
    On Error GoTo ErrH
    
    Select Case Control.ID
        Case conMenu_Edit_NewKind                       '��������
            Call ProgAdd
        Case conMenu_Edit_ModifyKind                    '�޸ķ���
            Call ProgEdit
        Case conMenu_Edit_DeleteKind                    'ɾ������
            Call ProgDel
        Case conMenu_View_Show                          '���뷽��
            Call ExecuteCommand("���뷽��")
        Case conMenu_View_Hide                          '��������
            Call ExecuteCommand("��������")
        Case conMenu_Edit_Select                        'ѡ�÷���
            Call ProgSele
        Case conMenu_Edit_NewParent                     '���ӷ���
            Call ExecuteCommand("���ӷ���")
        Case conMenu_Edit_ModifyParent                  '�޸ķ���
            Call ExecuteCommand("�޸ķ���")
        Case conMenu_Edit_DeleteParent                  'ɾ������
            Call ExecuteCommand("ɾ������")
        Case conMenu_Edit_NewItem                       '������Ŀ
            vsfAuditItem.Rows = vsfAuditItem.Rows + 1
            vsfAuditItem.Row = vsfAuditItem.Rows - 1
            Call ExecuteCommand("������Ŀ")
        Case conMenu_Edit_Modify                        '�޸���Ŀ
            Call ExecuteCommand("�޸���Ŀ")
        Case conMenu_Edit_CopyNewItem                   '����������Ŀ
            vsfAuditItem.Rows = vsfAuditItem.Rows + 1
            vsfAuditItem.Row = vsfAuditItem.Rows - 1
            Call ExecuteCommand("����������Ŀ")
        Case conMenu_Edit_Delete                        'ɾ����Ŀ
            Call ExecuteCommand("ɾ����Ŀ")
        Case conMenu_Edit_Transf_Save                   '������Ŀ
            Call ExecuteCommand("������Ŀ")
        Case conMenu_Edit_Transf_Cancle                  'ȡ����Ŀ
            blnNewCancel = frmAuditItemEdit.EditMode = ���� Or frmAuditItemEdit.EditMode = ��������
            
            Call ExecuteCommand("ȡ����Ŀ")
            If blnNewCancel And frmAuditItemEdit.EditMode = ��� Then
                vsfAuditItem.Rows = vsfAuditItem.Rows - 1
                vsfAuditItem.Row = vsfAuditItem.Rows - 1
            End If
            
        Case conMenu_View_ShowAll                       '�����¼�
            mblnCheckAll = Not mblnCheckAll
            Control.Checked = mblnCheckAll
            DataAuditItem
        Case conMenu_View_Find                          '��������
            Call ItemFind
        Case conMenu_File_Preview   'Ԥ��
            mzlPrintModeS = zlView
            Call ItemPrint
        Case conMenu_File_Print   '��ӡ
            mzlPrintModeS = zlPrint
            Call ItemPrint
        Case conMenu_File_Excel '�����&Excel
            mzlPrintModeS = zlExcel
            Call ItemPrint
        Case conMenu_View_Forward
            Call ExecuteCommand("ǰһ��")
        Case conMenu_View_Backward
            Call ExecuteCommand("��һ��")
        Case conMenu_View_Option
            mobjFindKey.Execute
        Case conMenu_View_LocationItem
            mstrFindKey = Control.Parameter
            mobjFindKey.Caption = mstrFindKey
            cbsMain.RecalcLayout
        Case conMenu_View_Location
            LocationObj txtLocation
        Case conMenu_View_Refresh               'ˢ��
            Dim lngRow      As Long
            If vsfAuditItem.Rows = 1 Then Exit Sub
            mintItemID = vsfAuditItem.TextMatrix(vsfAuditItem.Row, vsfAuditItem.ColIndex("ID"))
            Call DataAuditItem
            lngRow = vsfAuditItem.FindRow(mintItemID, -1, vsfAuditItem.ColIndex("ID"), False, True)
            If lngRow > 0 Then vsfAuditItem.Row = lngRow
        Case Else
            If Control.ID > 400 And Control.ID < 500 Then
                Call ReportOpen(gcnOracle, Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1), Me)
            Else
                 '��ҵ���޹صĹ��ܣ������Ĺ���
                Call CommandBarExecutePublic(Control, Me, vsfAuditItem, "�������Ŀ¼�嵥")
            End If
    End Select
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error GoTo ErrH

    picPane(0).Enabled = (frmAuditItemEdit.EditMode = ���)
    picPane(2).Enabled = (frmAuditItemEdit.EditMode = ���)
    txtLocation.Locked = (frmAuditItemEdit.EditMode <> ���)
    With vsfAuditItem
        Select Case Control.ID
            Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel 'Ԥ��,��ӡ,�����Excel
                Control.Enabled = ((vsfAuditItem.Rows > 1) And IsPrivs(mstrPrivs, "��ɾ��"))
            Case conMenu_EditPopup
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
            Case conMenu_Edit_NewParent
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_ModifyParent, conMenu_Edit_DeleteParent    '
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_NewItem                    '������Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_Modify                        '�޸���Ŀ
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_Delete                  'ɾ��
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_CopyNewItem                   '��������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If tvwAuditType.SelectedItem Is Nothing Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (InStrRev(tvwAuditType.SelectedItem.Key, "Root") = 0 And Control.Visible And (Not zlCheck.Connection_ChkRsState(mRsAuditItem)) And frmAuditItemEdit.EditMode = ���)
                End If
            Case conMenu_Edit_Transf_Save, conMenu_Edit_Transf_Cancle '�����ȡ������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = ((frmAuditItemEdit.EditMode <> ���) And Control.Visible)
            Case conMenu_View_Forward
                Control.Enabled = (frmAuditItemEdit.EditMode = ���) And vsfAuditItem.Row > 1
            Case conMenu_View_Backward
                Control.Enabled = (frmAuditItemEdit.EditMode = ���) And vsfAuditItem.Row + 1 < vsfAuditItem.Rows
            Case conMenu_View_Find, conMenu_View_Refresh
                Control.Enabled = (frmAuditItemEdit.EditMode = ���)
            Case conMenu_View_ShowAll                       '�����¼�
                Control.Checked = mblnCheckAll
                Control.Enabled = (frmAuditItemEdit.EditMode = ���)
            Case conMenu_View_LocationItem, conMenu_View_LocationItem, conMenu_View_LocationItem
                If InStr(Control.Caption, mstrFindKey) > 0 Then
                    Control.Checked = True
                Else
                    Control.Checked = False
                End If
                
            Case conMenu_Edit_NewKind                       '��������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = Control.Visible
            Case conMenu_Edit_ModifyKind                    '�޸ķ���
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.Nodes.count > 0
                
            Case conMenu_Edit_DeleteKind                    'ɾ������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.Nodes.count > 0
            Case conMenu_View_Show                          '���뷽��
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = (Control.Visible And Not mblnProgUsed) And tvwAuditType.SelectedItem.Children = 0 And Left(tvwAuditType.SelectedItem.Key, 4) = "Root"
                End If
            Case conMenu_View_Hide                          '��������
                Control.Visible = IsPrivs(mstrPrivs, "��ɾ��")
                If ObjPtr(tvwAuditType.SelectedItem) = 0 Then
                    Control.Enabled = False
                Else
                    Control.Enabled = Control.Visible And tvwAuditType.SelectedItem.Children > 0 And Left(tvwAuditType.SelectedItem.Key, 4) = "Root"
                End If
            Case conMenu_Edit_Select                        'ѡ�÷���
                Control.Enabled = tvwAuditType.Nodes.count > 0
            Case Else
                Call CommandBarUpdatePublic(Control, Me)
        End Select
    End With
    Exit Sub
ErrH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�رջ���ʾ
'==============================================================================
Private Sub picFAXX_Click()
    On Error GoTo ErrH
    
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
    Call picPane_Resize(2)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ������Ϣ�����ɫ
'==============================================================================
Private Sub picFAXX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
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
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=������鷽��
'==============================================================================
Private Sub ProgAdd()
    Dim f As New frm��鷽���༭
    On Error GoTo ErrH
    f.ShowForm   '����
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '����б�
'        Call DataLoad
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=�޸���鷽��
'==============================================================================
Private Sub ProgEdit()
    Dim f               As New frm��鷽���༭
    Dim lng�ܷ�         As Double
    On Error GoTo ErrH
    
'    mlngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If mlngCurFAID < 1 Then Exit Sub
    f.ShowForm mlngCurFAID   '�޸ģ�����ID
    Call DataUpdate
    If f.Moded Then
        Call InitTreeView
        '����б�
        Call DataAuditItem
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ɾ����鷽��
'==============================================================================
Private Sub ProgDel()
    Dim intIndex        As Long
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrH
    
'    m_lngCurFAID = Mid(tvw����.SelectedItem.Key, 2)
    If mlngCurFAID < 1 Then Exit Sub
    If MsgBox("��ȷ��Ҫɾ������������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '����Ƿ��ڲ���������¼�д���
    Set rsTmp = gclsPackage.GetProjectUse(mlngCurFAID)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!���� > 0 Then
            Call MsgBox("�÷����Ѿ���ʹ�ù�,�ݲ���ɾ��!", vbInformation, gstrSysName)
            Exit Sub
        End If
    End If
    gstrSQL = "ZL_������鷽��_Delete(" & CStr(mlngCurFAID) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Call InitTreeView
    Call DataAuditItem
    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=ѡ�÷���
'==============================================================================
Private Sub ProgSele()
    Dim intIndex        As Long
    Dim bln��ʹ��       As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lngDefaultID As Long
    
    On Error GoTo ErrH
    
    If mlngCurFAID < 1 Then Exit Sub
    If MsgBox("ע�⣺���ְ���ѡ����һ���ǳ����ص����飬ͨ����Ҫ������ģ�" & vbCrLf & "��ȷ��ѡ�ñ���鷽����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    For i = 1 To tvwAuditType.Nodes.count
        If tvwAuditType.Nodes(i).Bold Then
            lngDefaultID = Replace(tvwAuditType.Nodes(i).Key, "Root", "")
        End If
    Next
    
    Dim rsTemp As New ADODB.Recordset
    Set rsTmp = gclsPackage.GetProjectUse(lngDefaultID)
    If rsTmp.RecordCount = 1 Then
        If rsTmp!���� > 0 Then
            If MsgBox("ע�⣺ϵͳĬ�����ְ�����ʹ�õ��У��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    gstrSQL = "ZL_������鷽��_ѡ��(" & CStr(mlngCurFAID) & ",1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Call InitTreeView
    Call SetMenu
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ͳ��
'==============================================================================
Private Sub DataUpdate()
    Dim rs              As ADODB.Recordset
    Dim lng�ܷ�         As Double
    On Error GoTo ErrH
    gstrSQL = "Select ����,�ܷ�,�ֶ���,����ʱ��,ͣ��ʱ��,˵�� From ������鷽�� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngCurFAID)
    If Not rs.EOF Then
        lbl��������.Caption = rs("����")
        lbl�ֶ���.Caption = "�ֶ���:" & rs("�ֶ���")
        lbl����ʱ��.Caption = "����ʱ��:" & zlCommFun.NVL(rs("����ʱ��"))
        lbl�ܷ�.Caption = "�ܷ�:" & rs("�ܷ�")
        lng�ܷ� = rs("�ܷ�")
    Else
        lbl��������.Caption = ""
        lbl�ֶ���.Caption = ""
        lbl����ʱ��.Caption = ""
        lbl�ܷ�.Caption = ""
    End If
    
'''    gstrSQL = "select sum(��׼��ֵ) from �������ֱ�׼ where �ϼ�ID is null and ����ID = [1]"
'''    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngCurFAID)
'''    If Not rs.EOF Then
'''        If Abs(lng�ܷ� - rs.Fields(0)) > 0.01 Then
'''            lbl�ܷ� = lbl�ܷ� + "����Ŀ������Ϊ:" & rs.Fields(0)
'''            lbl�ܷ�.ForeColor = vbRed
'''        Else
'''            lbl�ܷ�.ForeColor = vbBlack
'''        End If
'''    Else
'''        lbl�ܷ�.ForeColor = vbRed
'''    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʾ��¼����Ϣ
'==============================================================================
Private Sub SetMenu()
    On Error GoTo ErrH
    stbThis.Panels(2).Text = "�б��й���ʾ��" & vsfAuditItem.Rows - 1 & "�����ݡ�"
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
