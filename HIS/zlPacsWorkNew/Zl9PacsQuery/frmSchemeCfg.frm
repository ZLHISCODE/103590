VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSchemeCfg 
   Caption         =   "��ѯ��������"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   16935
   Icon            =   "frmSchemeCfg.frx":0000
   LinkTopic       =   "Form2"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   16935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSchemeName 
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   6615
      TabIndex        =   10
      Top             =   600
      Width           =   6615
      Begin VB.ComboBox cbxDeptFilter 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   240
         Width           =   3855
      End
      Begin VB.PictureBox picBasic 
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4695
         ScaleWidth      =   5775
         TabIndex        =   14
         Top             =   3720
         Width           =   5775
         Begin VB.CheckBox chkFindRealTimeFilter 
            Caption         =   "���Ҷ���ɸѡ"
            Height          =   195
            Left            =   3480
            TabIndex        =   35
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CheckBox chkRealTimeFilter 
            Caption         =   "״̬ʵʱɸѡ"
            Height          =   195
            Left            =   3480
            TabIndex        =   34
            Top             =   2250
            Width           =   1095
         End
         Begin VB.CheckBox chkSelRowTransparent 
            Caption         =   "ѡ����͸��"
            Height          =   180
            Left            =   3600
            TabIndex        =   5
            Top             =   3000
            Width           =   1215
         End
         Begin VB.ComboBox cboColor 
            Height          =   300
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   4080
            Width           =   1455
         End
         Begin VB.CheckBox chkEmbedFind 
            Caption         =   "Ƕ�����ݼ���"
            Height          =   180
            Left            =   1920
            TabIndex        =   4
            Top             =   2640
            Width           =   1425
         End
         Begin VB.CheckBox chkTab 
            Caption         =   "��ʷ�������"
            Enabled         =   0   'False
            Height          =   180
            Left            =   240
            TabIndex        =   3
            Top             =   2640
            Width           =   1425
         End
         Begin VB.ComboBox cbxDept 
            Height          =   300
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   630
            Width           =   3135
         End
         Begin VB.Frame fraLine 
            Height          =   30
            Left            =   0
            TabIndex        =   19
            Top             =   2040
            Width           =   4695
         End
         Begin VB.CheckBox chkTrance 
            Caption         =   "�б��ܸ���"
            Height          =   195
            Left            =   240
            TabIndex        =   1
            Top             =   2250
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.TextBox txtDate 
            Height          =   270
            Left            =   3960
            TabIndex        =   7
            Text            =   "0"
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txtSchemeMemo 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox txtDay 
            Height          =   270
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "3"
            Top             =   3300
            Width           =   495
         End
         Begin VB.CheckBox chkLocate 
            Caption         =   "��λ������ʾ"
            Height          =   195
            Left            =   1920
            TabIndex        =   2
            Top             =   2250
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.TextBox txtSchemeName 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   120
            Width           =   3135
         End
         Begin VB.ComboBox cboRefreshTime 
            Height          =   300
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3720
            Width           =   1455
         End
         Begin VB.CommandButton cmdBack 
            Appearance      =   0  'Flat
            Caption         =   "����"
            Height          =   375
            Left            =   2760
            TabIndex        =   16
            Top             =   2880
            Width           =   735
         End
         Begin VB.CommandButton cmdFore 
            Appearance      =   0  'Flat
            Caption         =   "����"
            Height          =   375
            Left            =   2040
            TabIndex        =   15
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label labPatiTypeColor 
            AutoSize        =   -1  'True
            Caption         =   "����������ɫ��ʾ�� "
            Height          =   180
            Left            =   240
            TabIndex        =   33
            Top             =   4080
            Width           =   1710
         End
         Begin VB.Label lblRefreshTime 
            AutoSize        =   -1  'True
            Caption         =   "����Զ�ˢ�¼��   "
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   3720
            Width           =   1710
         End
         Begin VB.Label lblQueryDay 
            AutoSize        =   -1  'True
            Caption         =   "Ĭ�ϲ�ѯ����"
            Height          =   180
            Left            =   240
            TabIndex        =   31
            Top             =   3360
            Width           =   1080
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "��ѯ���ڷ�Χ����"
            Height          =   180
            Left            =   2400
            TabIndex        =   30
            Top             =   3360
            Width           =   1440
         End
         Begin VB.Label labDept 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ר�ÿ���:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   660
            Width           =   885
         End
         Begin VB.Label labFore 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "ѡ������ɫ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   480
            TabIndex        =   25
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label lblYears 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   4560
            TabIndex        =   24
            Top             =   3360
            Width           =   180
         End
         Begin VB.Label labSchemeMemo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����˵��:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   2040
            TabIndex        =   22
            Top             =   3360
            Width           =   180
         End
         Begin VB.Label labObj 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   180
            Width           =   975
         End
         Begin VB.Label labBack 
            BackColor       =   &H00FEE0E2&
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   3000
            Width           =   1695
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   4815
         _cx             =   8493
         _cy             =   5106
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   0
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label labDeptFilter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ɸѡ:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.PictureBox picSchemeContent 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   6960
      ScaleHeight     =   8415
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   720
      Width           =   10935
      Begin XtremeSuiteControls.TabControl tbcMain 
         Height          =   4095
         Left            =   480
         TabIndex        =   13
         Top             =   2160
         Width           =   9615
         _Version        =   589884
         _ExtentX        =   16960
         _ExtentY        =   7223
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   9570
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2963
            MinWidth        =   1764
            Picture         =   "frmSchemeCfg.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24192
            MinWidth        =   6068
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
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmSchemeCfg.frx":6DEE
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSchemeCfg.frx":6E02
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1320
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchemeCfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const SW_SHOW = 5
Private Const M_STR_GRIDDATA = "���|ID|����ID|��������|ʹ��״̬|�Ƿ�Ĭ��|�Ƿ���|ר�ÿ���|����˵��"       '�����ʾ
Private Const M_STR_CROOK = "��"
Private Enum ColTitle
    ct��� = 0
    ctID = 1
    ct����ID = 2
    ct�������� = 3
    ctʹ��״̬ = 4
    ct�Ƿ�Ĭ�� = 5
    ct�Ƿ��� = 6
    ctר�ÿ��� = 7
    ct����˵�� = 8
End Enum

Private Const conMenu_View_ToolBar = 301              '������(&T)
Private Const conMenu_View_ToolBar_Button = 302         '��׼��ť(&S)
Private Const conMenu_View_ToolBar_Text = 303           '�ı���ǩ(&T)
Private Const conMenu_View_ToolBar_Size = 304           '��ͼ��(&B)
Private Const conMenu_View_StatusBar = 305            '״̬��(&S)

Private Const conMenu_View_FontSize_S = 306            'С����
Private Const conMenu_View_FontSize_M = 307            '������
Private Const conMenu_View_FontSize_L = 308            '������

'�˵�����ö�ٶ���
Private Enum TMenuType
    mtFile = 1                  '�ļ�
    mtSave = 101                '����
    mtCancel = 102              '�ر�
    mtImport = 103              '����
    mtExport = 104              '����
    mtQuit = 105                '�˳�
    
    mtEdit = 2                  '�༭
    mtNewScheme = 201           '����
    mtModifyScheme = 202        '�޸�
    mtDelScheme = 203           'ɾ��
    mtUsually = 204             '����
    mtSetDefault = 205           'Ĭ��
    mtRecover = 206              '�ָ�
    mtUseScheme = 207            '����/����
    mtCheckScheme = 208          '����
    mtMoveLastScheme = 209       '����
    mtMoveNextScheme = 210       '����
    mtSetSysQuery = 214          '�û���ѯ����
    mtResource = 215             '��Դ����
    mtCopy = 216                '����
    
    mtViewPopup = 3            '�鿴
    
'    mtHelpPopup = 4            '����
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private mlngModuleNo As Long     '����ģ��
Private mblnNewState As Boolean     '�Ƿ�������
Private mblnClose As Boolean
Private mblnIsEdit As Boolean   '�����Ƿ��ѱ༭
Private mlngRow As Long
Private mobjIconManage As New frmIconManage
Private WithEvents mobjQuerySet As frmScheme_BaseQueryCfg
Attribute mobjQuerySet.VB_VarHelpID = -1
Private mobjFilterSet As New frmScheme_FilterCfg
Private mobjDisPlaySet As New frmScheme_DisplayCfg
Private mobjSetRelated As New frmSetRelated
Private mobjSqlScheme As New clsSqlScheme
Private mblnEdit As Boolean
Private mlngVer As Long
Private mlngIndex As Long
Private mlngCurSchemeId As Long '��ǰ����ʹ�õķ���


Public Sub ShowMe(ByVal lngModuleNo As Long, ByVal strSysPara As String, ByVal strBasePara As String, ByVal bytFontSize As Byte, owner As Object, ByVal lngWindowShemeID As Long)
    mlngCurSchemeId = lngWindowShemeID
    mlngModuleNo = lngModuleNo
    gstrPara = strSysPara
    gstrBasePara = strBasePara
    gbytFontSize = IIf(bytFontSize = 0, 9, IIf(bytFontSize = 1, 12, IIf(bytFontSize = 2, 15, bytFontSize)))
    mblnEdit = False
    mblnNewState = False
    mblnClose = False
    mblnIsEdit = False
    mlngRow = 0
    mlngVer = -1
    
    Me.Show , owner
End Sub

Private Sub cboColor_Click()
On Error GoTo errHandle

    If mblnClose And Val(cboColor.Tag) <> cboColor.ListIndex Then
        mblnIsEdit = True
    End If
    
    cboColor.Tag = cboColor.ListIndex
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cboRefreshTime_Click()
On Error GoTo errHandle

    If mblnClose And mlngIndex <> cboRefreshTime.ListIndex Then
        mblnIsEdit = True
    End If
    
    mlngIndex = cboRefreshTime.ListIndex
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbrMain_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle
    
    Select Case control.Id
        Case TMenuType.mtSave       '����
            Call SaveScheme
        Case TMenuType.mtCancel     '�ر�
            Call CancelScheme
        Case TMenuType.mtImport     '����
            Call ImportScheme
        Case TMenuType.mtExport     '����
            Call ExportScheme
            tbcMain.SetFocus
        Case TMenuType.mtQuit       '�˳�
            Call UnloadMe
        Case TMenuType.mtNewScheme  '����
            Call NewScheme
        Case TMenuType.mtModifyScheme   '�޸�
            Call ModifyScheme
        Case TMenuType.mtDelScheme      'ɾ��
            Call DeleteScheme
        Case TMenuType.mtCopy       '����
            Call CopyScheme
        Case TMenuType.mtUsually    '����
            Call SetUsualScheme
        Case TMenuType.mtSetDefault 'Ĭ��
            Call SetDefaultScheme
        Case TMenuType.mtRecover    '�ָ�
            Call RecoverScheme
        Case TMenuType.mtUseScheme  '����/����
            Call SetUseScheme
        Case TMenuType.mtCheckScheme    '����
        Case TMenuType.mtMoveLastScheme     '����
            Call MoveLastScheme
        Case TMenuType.mtMoveNextScheme     '����
            Call MoveNextScheme
        Case TMenuType.mtSetSysQuery    '�û���ѯ����
            Call ShowUserScheme
        Case TMenuType.mtResource       '��Դ����
            Call mobjIconManage.ShowIconWindow("", Me)
            tbcMain.SetFocus
'---------------------------�鿴----------------
        Case conMenu_View_ToolBar_Button '������
            Call Menu_View_ToolBar_Button_click(control)
        Case conMenu_View_ToolBar_Text '��ť����
            Call Menu_View_ToolBar_Text_click(control)
        Case conMenu_View_ToolBar_Size '��ͼ��
            Call Menu_View_ToolBar_Size_click(control)
        Case conMenu_View_StatusBar '״̬��
            Call Menu_View_StatusBar_click(control)
        Case conMenu_View_FontSize_S    'С����
            Call SetFontSize(0)
        Case conMenu_View_FontSize_M    '������
            Call SetFontSize(1)
        Case conMenu_View_FontSize_L    '������
            Call SetFontSize(2)
            
    End Select
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    Dim blnIsAllowEdit As Boolean
    Dim lngSelDeptId As Long
    
On Error GoTo errHandle
    blnIsAllowEdit = True
    If vsfMain.Rows > 1 Then
        If cbxDeptFilter.ListIndex > 0 Then
            lngSelDeptId = Val(cbxDeptFilter.ItemData(cbxDeptFilter.ListIndex))
            
            If lngSelDeptId > 0 Then
                If Val(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����ID)) <> lngSelDeptId Then blnIsAllowEdit = False
            End If
        End If
    End If

    Select Case control.Id
        Case TMenuType.mtSave       '����
            If mblnClose Then
                control.Enabled = IsEdit
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtCancel     '�ر�
            control.Enabled = mblnClose
        Case TMenuType.mtImport     '����
            control.Enabled = Not mblnClose
        Case TMenuType.mtExport     '����
            control.Enabled = Not mblnClose
        Case TMenuType.mtNewScheme  '����
            control.Enabled = Not mblnClose
        Case TMenuType.mtModifyScheme   '�޸�
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1) And blnIsAllowEdit
            
        Case TMenuType.mtDelScheme      'ɾ��
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1) And blnIsAllowEdit
        Case TMenuType.mtRecover
            If Not mobjSqlScheme Is Nothing Then
                control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1) And Len(Trim(mobjSqlScheme.Store)) > 0 And blnIsAllowEdit
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtCopy      '����
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtUseScheme
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1) And blnIsAllowEdit
            If IsSelectionRow(vsfMain) Then
                control.Caption = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", "����", "����")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", 211, 207)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtUsually
            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����") And blnIsAllowEdit
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)) = 0, "���ó���", "ȡ������")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) = M_STR_CROOK, 212, 204)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtSetDefault

            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����") And blnIsAllowEdit
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)) = 0, "����Ĭ��", "ȡ��Ĭ��")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) = M_STR_CROOK, 213, 205)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtCheckScheme    '����
            control.Enabled = Not (vsfMain.Rows <= 1)
        Case TMenuType.mtMoveLastScheme     '����
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtMoveNextScheme     '����
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
    End Select

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbxDept_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbxDept_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbxDeptFilter_Change()
On Error GoTo errHandle
    Dim lngDeptId As Long
    
    lngDeptId = 0
    If cbxDeptFilter.ListIndex > 0 Then lngDeptId = Val(cbxDeptFilter.ItemData(cbxDeptFilter.ListIndex))
    
    Call FilterDeptScheme(lngDeptId)
    Call ShowScheme
Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub FilterDeptScheme(ByVal lngDeptId As Long)
'�������ˢ��
    Dim i As Long
    Dim strSql As String
    Dim lngRowDeptId As Long
    
    For i = 1 To vsfMain.Rows - 1
        lngRowDeptId = Val(vsfMain.TextMatrix(i, ColTitle.ct����ID))
        
        If lngRowDeptId = 0 _
            Or lngRowDeptId = lngDeptId Or lngDeptId = 0 Then
            vsfMain.RowHidden(i) = False
        Else
            vsfMain.RowHidden(i) = True
        End If
    Next
    
    If lngDeptId <> 0 Then
        For i = 1 To vsfMain.Rows - 1
            If Val(vsfMain.TextMatrix(i, ColTitle.ct����ID)) = 0 Then
                vsfMain.Cell(flexcpBackColor, i, 0, i, vsfMain.Cols - 1) = &HE0E0E0
            Else
                vsfMain.Cell(flexcpBackColor, i, 0, i, vsfMain.Cols - 1) = vbWhite
            End If
        Next
    Else
        vsfMain.Cell(flexcpBackColor, 1, 0, vsfMain.Rows - 1, vsfMain.Cols - 1) = vbWhite
    End If

End Sub

Private Sub cbxDeptFilter_Click()
On Error GoTo errHandle
    Dim lngDeptId As Long
    
    lngDeptId = 0
    If cbxDeptFilter.ListIndex > 0 Then lngDeptId = Val(cbxDeptFilter.ItemData(cbxDeptFilter.ListIndex))
    
    Call FilterDeptScheme(lngDeptId)
    Call ShowScheme
Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkEmbedFind_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkFindRealTimeFilter_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkRealTimeFilter_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkLocate_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkSelRowTransparent_Click()
On Error GoTo errHandle
    If chkSelRowTransparent.Value = 1 Then
        labBack.BackColor = picBasic.BackColor
    Else
        labBack.BackColor = Val(labBack.Tag)
    End If

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    If App.LogMode = 0 Then MsgBox "chkSelRowTransparent_Click"
End Sub

Private Sub chkTrance_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cmdBack_Click()
    Dim lngColor As Long
    On Error Resume Next
    
    dlgFile.Flags = cdlCCFullOpen
    dlgFile.Color = labBack.BackColor
    dlgFile.ShowColor
    lngColor = dlgFile.Color
    
    labBack.BackColor = lngColor
    labBack.Tag = lngColor
    
    If mblnClose Then
        mblnIsEdit = True
    End If
End Sub

Private Sub cmdfore_Click()
    On Error Resume Next
    Dim lngColor As Long
    
    dlgFile.Flags = cdlCCFullOpen
    dlgFile.Color = labFore.ForeColor
    dlgFile.ShowColor
    
    lngColor = dlgFile.Color
    
    labFore.ForeColor = lngColor
    
    If mblnClose Then
        mblnIsEdit = True
    End If
End Sub

Private Sub Form_Activate()
    Call picSchemeContent_Resize
    Call picSchemeName_Resize
End Sub

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Me.ZOrder
    
    Set mobjQuerySet = New frmScheme_BaseQueryCfg
    
    Call InitCommandBars
    Call InitTabControl
    Call InitTime
    Call GridInit(M_STR_GRIDDATA, vsfMain)
    
    Call RefreshWindowState(False)
    Call initWindow
    
    Call RefreshDept
    Call RefreshList
    Call RefreshScheme
    Call ShowScheme
    Call SetFontSize(gbytFontSize)
    Call InitDockPannel
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub InitTime()
    With cboRefreshTime
        .AddItem "��ˢ��", 0
        .AddItem "1����", 1
        .AddItem "2����", 2
        .AddItem "3����", 3
        .AddItem "5����", 4
        .AddItem "10����", 5
        
        .ListIndex = 0
    End With
    
End Sub


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    
    '���ò˵����͹��������
    With cbrMain.Options
        .ShowExpandButtonAlways = False                         '�����ڹ������Ҳ���ʾѡ�ť,��ʹ�������㹻��
        .ToolBarAccelTips = True                                '��ʾ��ť��ʾ
        .AlwaysShowFullMenus = False                            '�����õĲ˵���������
        .UseFadedIcons = False                                  'ͼ����ʾΪ��ɫЧ��
        .IconsWithShadow = True                                 '���ָ�������ͼ����ʾ��ӰЧ��
        .UseDisabledIcons = True                                '��������ť����ʱͼ����ʾΪ������ʽ
        .LargeIcons = True                                      '��������ʾΪ��ͼ��
        .SetIconSize True, 32, 32                               '���ô�ͼ��ĳߴ�
        .SetIconSize False, 16, 16                              '����Сͼ��ĳߴ�
        .ShowTextBelowIcons = True
    End With
    With cbrMain
        .VisualTheme = xtpThemeOffice2003                        '���ÿؼ���ʾ���
        .EnableCustomization False                               '�Ƿ������Զ�������
        Set .Icons = imgMain.Icons                               '���ù�����ͼ��ؼ�
    End With

    Me.cbrMain.EnableCustomization False
    Me.cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '�˵�����
'Begin------------------------�༭�˵�--------------------------------------Ĭ�Ͽɼ�
    cbrMain.ActiveMenuBar.Title = "�˵�"
    
    
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtFile, "�ļ�(&F)")
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSave, "����(&S)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancel, "�ر�(&C)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtImport, "����(&I)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtExport, "����(&E)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtQuit, "�˳�(&Q)"): cbrControl.BeginGroup = True
    End With
        
        Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, TMenuType.mtEdit, "�༭(&E)")
        
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewScheme, "����(&N)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtModifyScheme, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelScheme, "ɾ��(&D)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCopy, "����(&C)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "Ĭ��(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "����(&Y)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "�ָ�(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "����(&A)")
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "����(&V)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����(&X)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetSysQuery, "�û���ѯ����(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtResource, "��Դ����(&Z)")
    End With
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mtViewPopup, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)

    
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSave, "����", "���淽��")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancel, "�ر�", "�رձ༭")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewScheme, "����", "��������"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtModifyScheme, "�޸�", "�޸ķ���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelScheme, "ɾ��", "ɾ������")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCopy, "����", "���Ʒ���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "Ĭ��", "����Ĭ�Ϸ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "����", "���÷���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "�ָ�", "�ָ�����")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "����", "���÷���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����", "���Ʒ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����", "���Ʒ���")
'        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "����", "���Է���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtQuit, "�˳�", "�˳�"): cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
End Sub

Private Sub CreateViewAndHelpMenu(objViewMenu As Object, objHelpMenu As Object, _
    Optional ByVal strMenuTag As String = "")
    
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrPopControl As CommandBarControl
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    If Not (objViewMenu Is Nothing) Then
        Set cbrMenuBar = objViewMenu
        
        With cbrMenuBar.CommandBar
            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(T)")
                cbrControl.Category = strMenuTag
            
                With cbrControl.CommandBar '�����˵�
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(0)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                        
                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(1)")
                        cbrPopControl.Checked = True
                        cbrPopControl.Category = strMenuTag
                End With
            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(S)")
                cbrControl.Checked = True
                        cbrPopControl.Category = strMenuTag
        End With
    End If

End Sub

'���ְ�
Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error GoTo errHandle
    
    Select Case Item.Id
        Case 1
            Item.Handle = picSchemeName.hwnd
        Case 2
            Item.Handle = picSchemeContent.hwnd
    End Select
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

'���沼��
Private Sub InitDockPannel()
    Dim objPane As Pane
    Dim intCx As Integer '���ڿ��Ƴ�ʼ���
    
    On Error GoTo errHandle

    dkpMain.SetCommandBars cbrMain
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    If gbytFontSize = 15 Then
        intCx = 150
    ElseIf gbytFontSize = 12 Then
        intCx = 140
    Else
        intCx = 115
    End If
    
    Set objPane = dkpMain.CreatePane(1, intCx, 100, DockLeftOf, Nothing)
    objPane.Title = "picSchemeName"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, objPane)
    objPane.Title = "picSchemeContent"
    objPane.Options = PaneNoCaption
    
    Set objPane = Nothing
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    If cbrMain.FindControl(, TMenuType.mtSave).Enabled Then
        If MsgBox("��δ���淽�����Ƿ񱣴棿", vbYesNo, Me.Caption) = vbYes Then
            Call SaveScheme
        End If
    End If
    
    If mblnEdit Then
        MsgBox "�����иĶ�������������վ��Ч��", vbInformation, Me.Caption
    End If
    
    If Not mobjQuerySet Is Nothing Then mobjQuerySet.UnloadMe
    If Not mobjFilterSet Is Nothing Then mobjFilterSet.UnloadMe
    If Not mobjDisPlaySet Is Nothing Then mobjDisPlaySet.UnloadMe
    If Not mobjSetRelated Is Nothing Then mobjSetRelated.UnloadMe
    If Not mobjIconManage Is Nothing Then mobjIconManage.UnloadMe
    
    
    For i = 1 To dkpMain.PanesCount
        dkpMain.Panes(i).Handle = 0
    Next
    
    dkpMain.CloseAll
    

    
    Set mobjQuerySet = Nothing
    Set mobjFilterSet = Nothing
    Set mobjDisPlaySet = Nothing
    Set mobjSqlScheme = Nothing
    Set mobjSetRelated = Nothing
    Set mobjIconManage = Nothing
End Sub

Private Sub mobjQuerySet_DoCheckVerify(ByVal strHint As String)
    On Error GoTo errHandle
    
    stbThis.Panels(2).Text = strHint
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub picBasic_Resize()
  On Error Resume Next
    
    Dim intSize As Integer
    Dim intWidth As Integer
    Dim lngW6words As Long
    Dim lngW5words As Long

    If gbytFontSize = 9 Then
        lngW6words = 1450
        lngW5words = 1200
    ElseIf gbytFontSize = 12 Then
        lngW6words = 1800
        lngW5words = 1550
    Else
        lngW6words = 2200
        lngW5words = 1900
    End If
    
    txtSchemeName.Left = labObj.Left + labObj.Width + 100
    txtSchemeName.Width = picBasic.Width - txtSchemeName.Left - 100
    
    cbxDept.Left = labDept.Left + labDept.Width + 100
    cbxDept.Width = picBasic.Width - cbxDept.Left - 100
    
    txtSchemeMemo.Left = labSchemeMemo.Left + labSchemeMemo.Width + 100
    txtSchemeMemo.Width = picBasic.Width - txtSchemeMemo.Left - 100
    
    fraLine.Width = picBasic.Width
    
    
    chkTrance.Width = lngW6words
    chkTrance.Left = picBasic.Left + 100
    chkTrance.Top = fraLine.Top + fraLine.Height + 50

    chkLocate.Width = lngW6words
    chkLocate.Top = chkTrance.Top
    chkLocate.Left = chkTrance.Left + chkTrance.Width + 100
    
    chkRealTimeFilter.Width = lngW6words
    chkRealTimeFilter.Top = chkLocate.Top
    chkRealTimeFilter.Left = chkLocate.Left + chkLocate.Width + 100

    chkTab.Width = lngW6words
    chkTab.Top = chkTrance.Top + chkTrance.Height + 100
    chkTab.Left = chkTrance.Left

    chkEmbedFind.Width = lngW6words
    chkEmbedFind.Top = chkTab.Top
    chkEmbedFind.Left = chkTab.Left + chkTab.Width + 100
    
    chkFindRealTimeFilter.Width = lngW6words
    chkFindRealTimeFilter.Top = chkEmbedFind.Top
    chkFindRealTimeFilter.Left = chkEmbedFind.Left + chkEmbedFind.Width + 100
    
    intSize = labFore.FontSize

    Call labBack.Move(chkTrance.Left, chkTab.Top + chkTab.Height + 100, labBack.Width, 360) 'todo
    Call labFore.Move(labBack.Left + (labBack.Width - labFore.Width) / 2, labBack.Top + (labBack.Height - labFore.Height) / 2)
    
    Call cmdFore.Move(labBack.Left + labBack.Width + 100, labBack.Top)
    Call cmdBack.Move(cmdFore.Left + cmdFore.Width, cmdFore.Top)
    
    Call chkSelRowTransparent.Move(cmdBack.Left + cmdBack.Width + 40, cmdBack.Top, lngW5words)
    lblQueryDay.Top = labBack.Top + labBack.Height + 240
    lblQueryDay.Left = labBack.Left
    txtDay.Left = lblQueryDay.Left + lblQueryDay.Width
    txtDay.Top = Abs(lblQueryDay.Top - Abs(txtDay.Height - lblQueryDay.Height) / 2)
    lblDay.Left = txtDay.Left + txtDay.Width
    lblDay.Top = lblQueryDay.Top
    
    Call lblDate.Move(lblDay.Left + lblDay.Width + 100, lblDay.Top)
    Call txtDate.Move(lblDate.Left + lblDate.Width, Abs(lblQueryDay.Top - Abs(txtDate.Height - lblQueryDay.Height) / 2))
    
    Call lblYears.Move(txtDate.Left + txtDate.Width + 100, lblDate.Top)
    Call lblRefreshTime.Move(lblQueryDay.Left, lblDate.Top + lblDate.Height + 180)

    cboRefreshTime.Left = lblRefreshTime.Left + lblRefreshTime.Width
    cboRefreshTime.Top = Abs(lblRefreshTime.Top - Abs(cboRefreshTime.Height - lblRefreshTime.Height) / 2)
    
    Call labPatiTypeColor.Move(lblRefreshTime.Left, lblRefreshTime.Top + lblRefreshTime.Height + 180)
    Call cboColor.Move(labPatiTypeColor.Left + labPatiTypeColor.Width, Abs(labPatiTypeColor.Top - Abs(cboColor.Height - labPatiTypeColor.Height) / 2))
End Sub

Private Sub picSchemeContent_Resize()
    On Error Resume Next

    tbcMain.Move picSchemeContent.Left, picSchemeContent.Top, picSchemeContent.Width, picSchemeContent.Height
End Sub

Private Sub picSchemeName_Resize()
    On Error Resume Next

    labDeptFilter.Left = picSchemeName.Left + 30
    
    cbxDeptFilter.Left = labDeptFilter.Left + labDeptFilter.Width
    cbxDeptFilter.Top = picSchemeName.Top + 30
    cbxDeptFilter.Width = picSchemeName.Width - labDeptFilter.Width - 60
    
    labDeptFilter.Top = cbxDeptFilter.Top + 30
    
    vsfMain.Move picSchemeName.Left, cbxDeptFilter.Top + cbxDeptFilter.Height + 30, picSchemeName.Width - 30, picSchemeName.Height - cbxDeptFilter.Height - IIf(stbThis.Visible, stbThis.Height, 0) - picBasic.Height + 90
    
    picBasic.Move vsfMain.Left, vsfMain.Top + vsfMain.Height, vsfMain.Width
End Sub

Private Sub InitTabControl()
    With tbcMain
        With .PaintManager
            .BoldSelected = True
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameNone
            .Position = xtpTabPositionTop
            .OneNoteColors = False
            .BoldSelected = True
            .ColorSet.ButtonSelected = &HFFC0C0
            .ColorSet.ButtonNormal = &HE0E0E0
            .Layout = xtpTabLayoutAutoSize
            .ButtonMargin.Top = 3
            .ButtonMargin.Bottom = 4
            .ShowIcons = True
        End With
        .InsertItem 0, "������ѯ����", mobjQuerySet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "������ѯ����"
        .InsertItem 1, "��ѯ��������", mobjFilterSet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "��ѯ��������"
        .InsertItem 2, "������ʾ����", mobjDisPlaySet.hwnd, 0
        .Item(.ItemCount - 1).Tag = "������ʾ����"
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strResult As String
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    Dim rsRecord As Recordset
    Dim strItem As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If Item.Tag = "��ѯ��������" Or Item.Tag = "������ʾ����" Then
        If Len(mobjQuerySet.GetQuerySql) = 0 Then
            MsgBox "���������ѯ��䡣", vbInformation, Me.Caption
            tbcMain.Item(0).Selected = True
            Exit Sub
        End If
        
        If mobjQuerySet.mintVerify = 0 Then
            ShowStateBar "���ڽ������ݲ�ѯ�����֤...�����ʱ�������������䣩"
            strResult = SqlVerify(mobjQuerySet.GetQuerySql)
            
            If Len(strResult) = 0 Then
                mobjQuerySet.mintVerify = 2
                strResult = IsHaveID(mobjQuerySet.GetQuerySql)
            End If
            
            ShowStateBar ""
            If Len(strResult) > 0 Then
                mobjQuerySet.mintVerify = 1
                MsgBox "��ѯ�����֤ʧ�ܣ�" & vbCrLf & "ԭ���ǣ�" & strResult, vbInformation, Me.Caption
                tbcMain.Item(0).Selected = True
                Call mobjQuerySet.rtbCheckSQLSetFocus
                Exit Sub
            End If
        ElseIf mobjQuerySet.mintVerify = 1 Then
            MsgBox "��ѯ�����֤ʧ�ܡ�", vbInformation, Me.Caption
            tbcMain.Item(0).Selected = True
            Call mobjQuerySet.rtbCheckSQLSetFocus
            Exit Sub
        End If
        
        Call RefreshShowScheme
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub RefreshList()
'�������ˢ��
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset

    vsfMain.Clear
    
    strSql = "select '' ���,a.ID,nvl(a.��������,0) as ����ID, ��������,ʹ��״̬,�Ƿ�Ĭ��,�Ƿ���,b.���� as ר�ÿ���,����˵�� from Ӱ���ѯ���� a, ���ű� b where a.��������=b.ID(+) And ����ģ�� = [1] Order By �������"
    Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", mlngModuleNo)
    Set vsfMain.DataSource = rsData
    
    vsfMain.ColHidden(ColTitle.ctID) = True
    vsfMain.ColHidden(ColTitle.ct����ID) = True
    
    Call DataConvert
    Call SchemeNo

    vsfMain.ColWidth(ColTitle.ct��������) = 2000
End Sub


Private Sub RefreshDept()
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errHandle
    strSql = "select ID,����  from  ���ű� a, ��������˵�� b where a.id=b.����id and b.��������='���' order by ����"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "��ѯ��鲿��", Me.Caption)
    
    cbxDept.Clear
    cbxDeptFilter.Clear
    
    cbxDept.AddItem ""
    cbxDept.ItemData(0) = 0
    
    cbxDeptFilter.AddItem ""
    cbxDeptFilter.ItemData(0) = 0
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    While Not rsData.EOF
        cbxDept.AddItem NVL(rsData!����)
        cbxDept.ItemData(cbxDept.ListCount - 1) = Val(rsData!Id)
        
        cbxDeptFilter.AddItem NVL(rsData!����)
        cbxDeptFilter.ItemData(cbxDept.ListCount - 1) = Val(rsData!Id)
        
        rsData.MoveNext
    Wend
    
Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
End Sub

Private Sub DataConvert()
'�������ת��
    Dim i As Long
    
    If vsfMain.Rows <= 1 Then
        Exit Sub
    End If
    
    For i = 1 To vsfMain.Rows - 1
        If Val(vsfMain.TextMatrix(i, ColTitle.ctʹ��״̬)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ctʹ��״̬) = "����"
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ctʹ��״̬)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ctʹ��״̬) = "����"
        End If
        
        If Val(vsfMain.TextMatrix(i, ColTitle.ct�Ƿ���)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ct�Ƿ���) = ""
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ct�Ƿ���)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ct�Ƿ���) = M_STR_CROOK
        End If
        
        If Val(vsfMain.TextMatrix(i, ColTitle.ct�Ƿ�Ĭ��)) = 0 Then
            vsfMain.TextMatrix(i, ColTitle.ct�Ƿ�Ĭ��) = ""
        ElseIf Val(vsfMain.TextMatrix(i, ColTitle.ct�Ƿ�Ĭ��)) = 1 Then
            vsfMain.TextMatrix(i, ColTitle.ct�Ƿ�Ĭ��) = M_STR_CROOK
        End If
        
    Next
End Sub

Private Sub txtDate_Change()
    On Error GoTo errHandle
    If Val(txtDate.Text) > 99 Then
        txtDate.Text = 99
    End If
    
    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDay_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    
    If Val(txtDay.Text) > 15 Then
        txtDay.Text = 15
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDay_LostFocus()
    If Val(txtDay.Text) < 0 Then
        txtDay.Text = 0
    End If
End Sub

Private Sub txtSchemeMemo_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtSchemeMemo_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If Len(txtSchemeMemo.Text) > 511 And Chr(KeyAscii) <> vbBack Then KeyAscii = 0
End Sub

Private Sub txtSchemeName_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtSchemeName_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If Len(txtSchemeName.Text) > 29 And Chr(KeyAscii) <> vbBack Then KeyAscii = 0
End Sub

Private Sub vsfMain_DblClick()
    On Error GoTo errHandle
    
    If vsfMain.Rows <= 1 Then Exit Sub
    Call ModifyScheme
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub vsfMain_RowColChange()
    On Error GoTo errHandle
    
    If vsfMain.Row <> mlngRow Then
        mlngRow = vsfMain.Row
        Call ShowScheme
        Call RefreshShowScheme
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub NewScheme()
'��������
    mblnNewState = True
    mblnClose = True
    cbxDeptFilter.Enabled = False
    
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
    Call NewRow(vsfMain)
    Call InitScheme
    Call SetNoneEdit
End Sub

Private Sub ModifyScheme()
'�޸ķ���

    mblnNewState = False
    mblnClose = True
    cbxDeptFilter.Enabled = False
    
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
    Call SetNoneEdit
End Sub

Private Sub SaveScheme()
    Dim strSql As String
    Dim strText As String
    Dim rsData As Recordset
    Dim lngDeptId As Long
    Dim blnIsChangeDept As Boolean
    Dim lngDefaultTag As Long
    Dim blnIsDefalut As Boolean '�������Ƿ�Ĭ�Ϸ���

    blnIsDefalut = IsDefault()
    strText = mobjQuerySet.GetQuerySql
    
    If blnIsDefalut Then
        If Not IsEnabledToSave() Then Exit Sub
        If Not VerBeforeSaveScheme(strText) Then Exit Sub
    Else
        If IsEnabledToSave() Then
            Call VerBeforeSaveScheme(strText)
        End If
    End If
    
    strText = GetSchemeContent
    
    If cbxDept.ListIndex >= 0 Then
        lngDeptId = cbxDept.ItemData(cbxDept.ListIndex)
    Else
        lngDeptId = 0
    End If

    If mblnNewState Then
        strSql = "Zl_Ӱ���ѯ_��������('" & Replace(txtSchemeName.Text, "'", "''") & _
                                        "','" & Replace(txtSchemeMemo.Text, "'", "''") & _
                                        "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) & _
                                        "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) & _
                                        "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) & _
                                        "','" & mlngModuleNo & _
                                        "','" & Replace(strText, "'", "''") & _
                                        "'," & lngDeptId & ")"
    Else
        blnIsChangeDept = IIf(lngDeptId <> Val(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����ID)), True, False)
        
        If blnIsChangeDept Then
            lngDefaultTag = 0
        Else
            lngDefaultTag = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)) = 0, 0, 1)
        End If
        
        strSql = "Zl_Ӱ���ѯ_���·���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & _
                                            ",'" & Replace(txtSchemeName.Text, "'", "''") & _
                                            "','" & Replace(txtSchemeMemo.Text, "'", "''") & _
                                            "'," & lngDefaultTag & _
                                            "," & IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", 1, 0) & _
                                            "," & IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)) = 0, 0, 1) & _
                                            "," & mlngModuleNo & _
                                            ",'" & Replace(strText, "'", "''") & _
                                            "'," & lngDeptId & ")"
    End If
    
    Call zlDatabase.CallProcedure("Zl_Ӱ���ѯ_���·���", "�༭����", vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID), txtSchemeName.Text, txtSchemeMemo.Text, lngDefaultTag, IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", 1, 0), IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)) = 0, 0, 1), mlngModuleNo, strText, lngDeptId)
    
    If mblnNewState Then
        strSql = "select ID from Ӱ���ѯ���� where �������� = [1] and ����ģ�� = [2]"
        Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", txtSchemeName.Text, mlngModuleNo)
        If rsData.RecordCount < 1 Then
            Exit Sub
        End If
        
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) = rsData.Fields!Id
    Else
        If blnIsChangeDept Then vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) = ""
    End If
    
    If lngDeptId = 0 Then
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctר�ÿ���) = ""
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����ID) = 0
    Else
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctר�ÿ���) = cbxDept.Text
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����ID) = lngDeptId
    End If
    
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct��������) = txtSchemeName.Text
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����˵��) = txtSchemeMemo.Text
    vsfMain.RowData(vsfMain.Row) = strText

    Call ShowScheme
    Call RefreshShowScheme
    If mblnNewState Then
        Call SetUseScheme
        Call SchemeNo
    End If

    mblnNewState = False
    Call SetNoneEdit
    
    mblnEdit = True
End Sub

Private Function GetSchemeContent() As String
    Dim objSqlScheme As clsSqlScheme
    Dim objScSearchCfg As clsScSerachCfg
    Dim strText As String
    Dim strQuery As String
    Dim strDetail As String
    Dim strSql As String
    Dim rsData As Recordset
    
    
    Set objSqlScheme = New clsSqlScheme

    '������Ϣ
    With objSqlScheme
        If Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)) > 0 Then
            .SchemeId = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)
            
            If mlngVer >= 0 Then
                strSql = "select �汾 from Ӱ���ѯ���� where id = [1] "
                Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", .SchemeId)
                
                If rsData.RecordCount > 0 Then
                    .Ver = Val(NVL(rsData!�汾)) + 1
                End If
            End If
        Else
            .Ver = 0
        End If
        .UseFuncFollow = chkTrance.Value = 1
        .DataRange = Val(txtDate.Text)
        .AutoRefreshTimeLen = cboRefreshTime.ListIndex
        .LocSerachHint = chkLocate.Value = 1
        .RealTimeFilter = chkRealTimeFilter.Value = 1
        .FindRealTimeFilter = chkFindRealTimeFilter.Value = 1
        .DefaultQueryDays = Val(txtDay.Text)
        .SchemeName = txtSchemeName
        .Descript = txtSchemeMemo
        .Store = mobjSqlScheme.Store
        .BackColor = labBack.BackColor
        .ForeColor = labFore.ForeColor
        .EmbedFind = chkEmbedFind.Value = 1
        .SelRowTransparent = chkSelRowTransparent.Value = 1
        .OldHistoryStyle = True
        .PatiColor = cboColor.Text
    End With


    '��ѯ���ģ��
    Call mobjQuerySet.SetQueryCfg(objSqlScheme)

    '¼������ģ��
    Call mobjFilterSet.SetConditionCfg(objSqlScheme)

    Call mobjDisPlaySet.SetShowCfg(objSqlScheme)

    strText = objSqlScheme.GetScheme

    GetSchemeContent = strText

    Set objSqlScheme = Nothing
    Set objScSearchCfg = Nothing
End Function

Private Sub MoveLastScheme()
'���Ʒ���
    Dim strSql As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveUp(vsfMain)

    strSql = "zl_Ӱ���ѯ_�ƶ�����(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "���Ʒ���")
    Call SchemeNo
    mblnEdit = True
End Sub



Private Sub MoveNextScheme()
'���Ʒ���
    Dim strSql As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveDown(vsfMain)

    strSql = "zl_Ӱ���ѯ_�ƶ�����(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "���Ʒ���")
    Call SchemeNo
    mblnEdit = True
End Sub

Private Sub SetDefaultScheme()
'����Ĭ�Ϸ���
    Dim strSql As String
    Dim strCurDefaultState As String
    Dim i As Long
    Dim lngDeptId As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurDefaultState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)
    lngDeptId = Val(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct����ID))
    
    strSql = "Zl_Ӱ���ѯ_Ĭ�Ϸ���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & _
                                    "," & IIf(strCurDefaultState = M_STR_CROOK, 0, 1) & _
                                    "," & mlngModuleNo & _
                                    "," & lngDeptId & ")"
    Call ExecuteCmd(strSql, "����Ĭ��")
     
    For i = 1 To vsfMain.Rows - 1
        If lngDeptId = Val(vsfMain.TextMatrix(i, ColTitle.ct����ID)) Then
            vsfMain.TextMatrix(i, ColTitle.ct�Ƿ�Ĭ��) = ""
        End If
    Next
  
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) = IIf(Len(strCurDefaultState) = 0, M_STR_CROOK, "")
    
    mblnEdit = True
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUsualScheme()
'�����Ƿ���
    Dim strSql As String
    Dim strCurUsualState As String
    Dim i As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUsualState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)
    strSql = "Zl_Ӱ���ѯ_���÷���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUsualState = M_STR_CROOK, 0, 1) & ")"
    Call ExecuteCmd(strSql, "���ó���")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) = IIf(Len(strCurUsualState) = 0, M_STR_CROOK, "")
    mblnEdit = True
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUseScheme()
'���÷���
    Dim strSql As String
    Dim strCurUseState As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUseState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬)

    strSql = "Zl_Ӱ���ѯ_���÷���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUseState = "����", 0, 1) & ")"
    Call ExecuteCmd(strSql, "ʹ��״̬����")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = IIf(strCurUseState = "����", "����", "����")
    mblnEdit = True
End Sub


Private Sub DeleteScheme()
'ɾ��ѡ�з���
    Dim strSql As String
    Dim lngRow As Long
    Dim lngID As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    If Not MsgBox("�Ƿ�ɾ��ѡ�з�����", vbYesNo, Me.Caption) = vbYes Then
        Exit Sub
    End If
    
    lngRow = vsfMain.Row
    lngID = vsfMain.TextMatrix(lngRow, ColTitle.ctID)
    
    If lngID = mlngCurSchemeId Then
        MsgBox "��ǰ���������У�����ɾ�����������ڹ���վ�л�������������Ȼ����ɾ���˷���", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strSql = "Zl_Ӱ���ѯ_ɾ������(" & lngID & ")"
    Call ExecuteCmd(strSql, "ɾ������")

    vsfMain.RemoveItem (lngRow)
    mblnEdit = True
End Sub

Private Sub CancelScheme()
    Dim i As Long
    
    If IsEdit Then
        If MsgBox("������δ���棬�Ƿ񱣴棿", vbYesNo, Me.Caption) = vbYes Then
            Call SaveScheme
        End If
    End If
    If mblnNewState Then
        vsfMain.RemoveItem vsfMain.Row
    End If
    
    For i = vsfMain.Row To 1 Step -1
        If vsfMain.RowHidden(i) = False Then
            vsfMain.Row = i
            Exit For
        End If
    Next
    
    Call ShowScheme
    Call RefreshShowScheme
    
    mblnNewState = False
    mblnClose = False
    cbxDeptFilter.Enabled = True
    
    Call RefreshWindowState(False)
    Call RefreshSubWindowState(False)
    
'    Call RefreshScheme
End Sub

Private Sub RefreshWindowState(blnState As Boolean)
    vsfMain.Enabled = Not blnState
    txtSchemeName.Enabled = blnState
    cbxDept.Enabled = blnState
    txtSchemeMemo.Enabled = blnState
    txtDate.Enabled = blnState
    cboRefreshTime.Enabled = blnState
    txtDay.Enabled = blnState
    chkTrance.Enabled = blnState
    chkLocate.Enabled = blnState
    chkRealTimeFilter.Enabled = blnState
    cmdFore.Enabled = blnState
    cmdBack.Enabled = blnState
    chkEmbedFind.Enabled = blnState
    chkSelRowTransparent.Enabled = blnState
    cboColor.Enabled = blnState
    chkFindRealTimeFilter.Enabled = blnState
End Sub

Private Sub InitScheme()
    chkTrance.Value = 1
    txtDate.Text = 0
    txtDay.Text = 3
    txtSchemeName.Text = ""
    chkTab.Value = 0
    chkEmbedFind.Value = 1
    
    If cbxDeptFilter.ListIndex > 0 Then
        cbxDept.ListIndex = cbxDeptFilter.ListIndex
    Else
        cbxDept.ListIndex = 0
    End If
    
    txtSchemeMemo.Text = ""
    chkLocate.Value = 1
    cboRefreshTime.ListIndex = 0
    chkRealTimeFilter.Value = 0
    chkFindRealTimeFilter.Value = 0

    tbcMain.Item(0).Selected = True
End Sub


Private Sub RefreshScheme()
'ˢ�·���
    Dim i As Long
    Dim lngNumber As Long
    Dim strSchemeXml As String

    If vsfMain.Rows < 2 Then Exit Sub
    For i = 1 To vsfMain.Rows - 1
        If Not Len(vsfMain.TextMatrix(i, ColTitle.ctID)) = 0 Then
            lngNumber = vsfMain.TextMatrix(i, ColTitle.ctID)
            strSchemeXml = ReadSchemeXml(lngNumber, "")
            vsfMain.RowData(i) = strSchemeXml
        End If
    Next

End Sub

Private Sub ShowScheme()
'��ʾ�����Ļ�����Ϣ����
    Dim i As Long
    
    mobjSqlScheme.OpenScheme vsfMain.RowData(vsfMain.Row)

    With mobjSqlScheme
        chkTrance.Value = IIf(.UseFuncFollow, 1, 0)
        chkLocate.Value = IIf(.LocSerachHint, 1, 0)
        chkRealTimeFilter.Value = IIf(.RealTimeFilter, 1, 0)
        chkFindRealTimeFilter.Value = IIf(.FindRealTimeFilter, 1, 0)
        txtDate.Text = .dateRange
        cboRefreshTime.ListIndex = IIf(.AutoRefreshTimeLen <= 0, 0, .AutoRefreshTimeLen)
        txtDay.Text = IIf(.DefaultQueryDays >= 0, .DefaultQueryDays, 3)
        txtSchemeName.Text = .SchemeName
        txtSchemeMemo.Text = .Descript
'        chkTab.Value = IIf(.QuickShowScheme, 1, 0)
        chkEmbedFind.Value = IIf(.EmbedFind, 1, 0)
        chkSelRowTransparent.Value = IIf(.SelRowTransparent, 1, 0)

        For i = 0 To cbxDept.ListCount - 1
            If cbxDept.List(i) = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctר�ÿ���) Then
                cbxDept.ListIndex = i
                Exit For
            End If
        Next
        
        labFore.ForeColor = .ForeColor
        labBack.BackColor = .BackColor
        labBack.Tag = labBack.BackColor
        mlngVer = .Ver
        
        '��ʾ��ѯ���ģ��
        Call mobjQuerySet.ShowQuerySet(mobjSqlScheme)
        '��ʾ���ٹ���ģ��
        Call mobjFilterSet.ShowFilterSet(mobjSqlScheme)
        '��ʾ��ʾ����ģ��
        Call mobjDisPlaySet.ShowDisplaySet(mobjSqlScheme)
        
        Call initCboColor(mobjSqlScheme)
        If .PatiColor <> "" Then
            For i = 0 To cboColor.ListCount - 1
                If cboColor.List(i) = .PatiColor Then
                    cboColor.ListIndex = i
                    Exit For
                End If
            Next
        End If
        
    End With
End Sub

Private Sub RefreshShowScheme()
'ˢ�·�������ʾ
    Dim strQuerySql As String

    '��ȡ��ǰ��ѯ���
    strQuerySql = mobjQuerySet.GetQuerySql

    If tbcMain.Item(1).Selected = True Then
        Call mobjFilterSet.RefreshFilterSet(strQuerySql, mobjSqlScheme)
    ElseIf tbcMain.Item(2).Selected = True Then
        Call mobjDisPlaySet.RefreshDisplaySet(strQuerySql)
    End If
End Sub

Private Function ExeSqlTrans(strSql As String) As String
    Dim i As Long
    
    On Error GoTo errRollback

    ExeSqlTrans = ""

    gcnOracle.BeginTrans

    If Len(strSql) > 0 Then
        Call ExecuteCmd(strSql, "���淽������")
    End If

    gcnOracle.CommitTrans
    Exit Function
errRollback:
    gcnOracle.RollbackTrans
    ExeSqlTrans = Err.Description
End Function

Private Function IsEnabledToSave() As Boolean
'��������
    Dim i As Long
    Dim strResult As String

    IsEnabledToSave = False
    If Len(Replace(txtSchemeName.Text, " ", "")) = 0 Then
        MsgBox "��������Ϊ�գ������뷽�����ơ�", vbInformation, Me.Caption
        txtSchemeName.SetFocus
        Exit Function
    End If

    '�����������ظ�
    For i = 1 To vsfMain.Rows - 2
        If vsfMain.TextMatrix(i, 2) = txtSchemeName.Text And (i <> vsfMain.Row) Then
            MsgBox "�������Ѵ��ڣ����顣", vbInformation, Me.Caption
            txtSchemeName.SetFocus
            Exit Function
        End If
    Next

    If Not mobjQuerySet.IsEnabledToSave Then
        MsgBox "��¼���ѯ���", vbInformation, Me.Caption
        Exit Function
    End If
    
    
    If mobjQuerySet.mintVerify = 0 Then
        ShowStateBar "���ڽ������ݲ�ѯ�����֤...�����ʱ�������������䣩"
        
        strResult = SqlVerify(mobjQuerySet.GetQuerySql)
        
        If Len(strResult) = 0 Then
            mobjQuerySet.mintVerify = 2
            strResult = IsHaveID(mobjQuerySet.GetQuerySql)
        End If
        
        ShowStateBar ""
        If Len(strResult) > 0 Then
            mobjQuerySet.mintVerify = 1
            MsgBox "��ѯ�����֤ʧ�ܣ�ԭ��Ϊ��" & strResult, vbInformation, Me.Caption
            Exit Function
        End If
    ElseIf mobjQuerySet.mintVerify = 1 Then
        MsgBox "��ѯ�����֤ʧ�ܣ���������֤��", vbInformation, Me.Caption
        Exit Function
    End If

    If Not mobjFilterSet.IsEnabledSave Then
        Exit Function
    End If
    
    If Not mobjFilterSet.IsSetted Then
        If MsgBox("��ѯ����������δ���á�", vbYesNo, Me.Caption) = vbNo Then
            Exit Function
        End If
    End If
    IsEnabledToSave = True
End Function

Private Sub RefreshSubWindowState(blnState As Boolean)
'ˢ���ӽ������״̬
    Call mobjFilterSet.RefreshWindowState(blnState)
    Call mobjQuerySet.RefreshWindowState(blnState)
    Call mobjDisPlaySet.RefreshWindowState(blnState)
End Sub

Private Sub ExportScheme()
'����
    Dim objExportScheme As New frmExportScheme
    Dim arrID() As Long
    Dim strFile As String
    Dim blnIcon As Boolean
On Error Resume Next
    objExportScheme.ShowMe mlngModuleNo, True, arrID, strFile, blnIcon, Me
    
    If Not objExportScheme Is Nothing Then Unload frmExportScheme
    Set objExportScheme = Nothing
    
    If Err.Number <> 0 Then MsgBox Err.Description, vbOKOnly, Me.Caption
End Sub

Private Sub ImportScheme()
'����
    Dim objExportScheme As New frmExportScheme
    Dim arrID() As Long
    Dim strFile As String
    Dim blnIcon As Boolean
On Error Resume Next

    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"

    dlgFile.FileName = ""
    dlgFile.ShowOpen

    If dlgFile.FileName = "" Then Exit Sub
    If Len(Dir(dlgFile.FileName)) = 0 Then
        MsgBox "�ļ������ڡ�", vbInformation, Me.Caption
        Exit Sub
    End If
    
    strFile = dlgFile.FileName

    If objExportScheme.ShowMe(mlngModuleNo, False, arrID, strFile, blnIcon, Me) Then
        ShowStateBar "���ڵ��뷽��..."
        Call ImportContent(arrID, strFile, blnIcon)
        
    End If
    
    If Not objExportScheme Is Nothing Then Unload objExportScheme
    Set objExportScheme = Nothing
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbOKOnly, Me.Caption
        Exit Sub
    End If
    
    Call RefreshList
    Call RefreshScheme
    Call ShowScheme
End Sub


Private Sub ImportContent(arrID() As Long, strFile As String, blnIcon As Boolean)
'���뷽��
    Dim rsData As ADODB.Recordset
    Dim lngOldSchemeId As Long
    Dim lngNewSchemeId As Long
    Dim strSql As String
    Dim strExeSql() As String
    Dim strResult As String
    Dim strText As String
    Dim strLog As String
    Dim i As Long
    Dim rsTemp As ADODB.Recordset
    Dim strSchemeName As String
    Dim lngCount As Long
    Dim lngScheme As Long
    Dim blnIsImport As Boolean
    Dim strOldName As String
    Dim blnImportIcon As Boolean
    Dim arrIcon() As String
    Dim strPath As String
    Dim strName As String
    Dim blnIsHave As Boolean
    Dim lngIconNum As Long
    Dim lngDefeated As Long
    Dim objSqlScheme As New clsSqlScheme

    If blnIcon Then
    '����ͼ��
        If Len(strFile) = 0 Then
            MsgBox "û��ͼ���ļ�������", vbInformation, Me.Caption
            Exit Sub
        End If
        strPath = Replace(strFile, ".XML", "\")
        If Len(Dir(strPath)) = 0 Then
            MsgBox "δ�ҵ�ͼ���ļ�������", vbInformation, Me.Caption
            Exit Sub
        End If
        strName = Dir(strPath & "*.ico", 7)

        strSql = "select ��Դ���� from Ӱ���ѯ��Դ where ��Դ���� = [1]"
        Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", 1)
        lngIconNum = 0
        Do
            If Len(Trim(strName)) = 0 Then Exit Do
            If Not IsHaveIcon(Replace(strName, ".ico", ""), rsData) Then
                strSql = "Zl_Ӱ���ѯ_����ͼ��('" & Replace(strName, ".ico", "") & "','1')"
                Call ExecuteCmd(strSql, "����ͼ��")
                Call zlBlobSave(Replace(strName, ".ico", ""), strPath & strName)
                lngIconNum = lngIconNum + 1
            Else
                strLog = "ͼ�����ơ�" & Replace(strName, ".ico", "") & "���Ѵ��ڣ�����ʱ�Ѻ��Ը�ͼ��"
                Call LogFile(strLog)
            End If
            strName = Dir()
        Loop
    End If

'
    Set rsData = New ADODB.Recordset
    Call rsData.Open(strFile)

    If rsData.RecordCount <= 0 Then
        MsgBox "û�п����ڵ�������ݣ������ļ��Ƿ���ȷ��", vbInformation, Me.Caption
        Exit Sub
    End If


    rsData.Sort = "id"

    lngScheme = 0
    lngDefeated = 0
    lngOldSchemeId = 0
    rsData.MoveFirst
    ReDim Preserve strExeSql(1)

    While Not rsData.EOF
        blnIsHave = False
        For i = 0 To UBound(arrID)
            If Val(NVL(rsData!Id)) = arrID(i) Then
                blnIsHave = True
                Exit For
            End If
        Next

        If blnIsHave Then
            lngCount = 0
            strOldName = ""
            strText = ""
            blnIsImport = True
            If lngOldSchemeId <> Val(NVL(rsData!Id)) Then
                '����Ӱ���ѯ������¼

                strSql = "select Ӱ���ѯ����_ID.NextVal as ID from dual"
                Set rsTemp = ExecuteSql(strSql, "��ȡ�·���ID")
                If rsTemp.RecordCount <= 0 Then
                    MsgBox "���ܻ�ȡ�����ķ���ID��ϵͳ���˳����档", vbExclamation, Me.Caption
                    Exit Sub
                End If

                lngNewSchemeId = Val(NVL(rsTemp!Id))
                strSchemeName = NVL(rsData!��������)

                For i = 1 To vsfMain.Rows - 1
                    If vsfMain.TextMatrix(i, ColTitle.ct��������) = strSchemeName Then
                        strOldName = strSchemeName
                        
                        Do While True
                            strSchemeName = strSchemeName & lngCount
                            If IsHaveScheme(strSchemeName) Then
                                Exit Do
                            End If

                            lngCount = lngCount + 1
                        Loop
                        
                        If Not MsgBox("�Ѵ�����Ϊ��" & strOldName & "���ķ���������󷽰�������Ϊ��" & strSchemeName & "�����Ƿ�������룿", vbYesNo, Me.Caption) = vbYes Then
                            blnIsImport = False
                        End If
                    End If
                Next
                If blnIsImport Then
                    strText = strText & rsData.Fields(3).Value
                    objSqlScheme.OpenScheme strText
                    strResult = SqlVerify(objSqlScheme.Query)
                    If Len(strResult) = 0 Then
                        If Len(strOldName) > 0 Then
                            strText = Replace(strText, "name=""" & strOldName & """", "name=""" & strSchemeName & """")
                        End If
                        ReDim Preserve strExeSql(UBound(strExeSql) + 1)
                        strExeSql(UBound(strExeSql) - 1) = "zl_Ӱ���ѯ_��������('" & _
                                                                    Replace(strSchemeName, "'", "''") & "','" & _
                                                                    NVL(rsData!����˵��) & "'," & _
                                                                    "'','1','','" & mlngModuleNo & "','" & _
                                                                    Replace(NVL(strText), "'", "''") & "')"
                    Else
                        strLog = "������" & NVL(rsData!��������) & "������ʧ�ܣ�ԭ��Ϊ��������֤ʧ�ܣ�" & strResult & "��"
                        Call LogFile(strLog)
                        lngDefeated = lngDefeated + 1
                    End If
                End If
                lngOldSchemeId = Val(NVL(rsData!Id))
            End If
        End If

        rsData.MoveNext
    Wend

    'д�뷽�������������
    For i = 0 To UBound(strExeSql)
        strSql = strExeSql(i)
        If Len(strSql) > 0 Then
            strResult = ExeSqlTrans(strSql)
            
            If Len(strResult) > 0 Then
                strLog = "������" & Mid(strSql, InStr(strSql, "('") + 2, InStr(strSql, "',") - InStr(strSql, "('") - 2) & "������ʧ�ܣ�ԭ��Ϊ��" & strResult
                Call LogFile(strLog)
                lngDefeated = lngDefeated + 1
            Else
                lngScheme = lngScheme + 1
            End If
        End If
    Next i
    
    ShowStateBar ""
    If blnIcon Then
        MsgBox "�ѵ���ɹ�" & lngScheme & "�����ݣ�ʧ��" & lngDefeated & "�����ݣ�" & lngIconNum & "��ͼ����Դ��", vbInformation, Me.Caption
    Else
        MsgBox "�ѵ���ɹ�" & lngScheme & "�����ݣ�ʧ��" & lngDefeated & "�����ݡ�", vbInformation, Me.Caption
    End If
    If lngDefeated > 0 Then
        ShellExecute Me.hwnd, "open", App.Path & "\" & "SchemeImport" & ".log", "", vbNullString, SW_SHOW
    End If
End Sub

Private Sub LogFile(ByVal strInfo As String)
    Dim lngFileNum As Long
    Dim FilePath As String
    Dim objFSO As Object
    Dim objLogFile As Object
    
    FilePath = App.Path & "\" & "SchemeImport" & ".log"

    lngFileNum = FreeFile
 
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Len(Dir(FilePath)) = 0 Then
        objFSO.CreateTextFile FilePath, True
    End If
    Set objLogFile = objFSO.GetFile(FilePath)
    If objLogFile = Empty Then
        Open FilePath For Output As #lngFileNum
    Else
        If objLogFile.Size > 2097152 Then
            objLogFile.Copy App.Path & "\" & App.EXEName & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".log"
            Open FilePath For Output As #lngFileNum
        Else
            Open FilePath For Append As #lngFileNum
        End If
    End If
 
    Print #lngFileNum, CStr(Now()) & ": " & strInfo
    Close #lngFileNum
 
End Sub

Private Sub UnloadMe()
'�˳�

    Unload Me
End Sub

Private Sub RecoverScheme()
    Dim strStore As String

    If MsgBox("�Ƿ�ȷ���ָ�������", vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    strStore = mobjSqlScheme.Store
    If Len(strStore) < 1 Then
        MsgBox "�÷���û�����ûָ����ԣ��޷��ָ���", vbInformation, Me.Caption
        Exit Sub
    End If
    If Not mobjSqlScheme.OpenScheme(strStore) Then
        MsgBox "�����ָ�ʧ�ܣ�����ָ������Ƿ���ȷ��", vbInformation, Me.Caption
        Exit Sub
    End If
    mobjSqlScheme.Store = strStore

    vsfMain.RowData(vsfMain.Row) = strStore
    Call ShowScheme
    Call SaveScheme
End Sub

Private Sub Menu_View_ToolBar_Button_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'�鿴
    Dim i As Integer

    On Error GoTo errHandle

    For i = 2 To cbrMain.Count
        Me.cbrMain(i).Visible = Not Me.cbrMain(i).Visible
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_ToolBar_Text_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'��ť
    Dim i As Integer, cbrControl As CommandBarControl
    Dim intStyle As Integer

    On Error GoTo errHandle

    For i = 2 To cbrMain.Count
        If Me.cbrMain(i).Controls.Count >= 1 Then
            intStyle = Me.cbrMain(i).Controls(i).Style
            If intStyle = xtpButtonIconAndCaption Then
                intStyle = xtpButtonIcon
                Me.cbrMain(i).ShowTextBelowIcons = False
            Else
                intStyle = xtpButtonIconAndCaption
                Me.cbrMain(i).ShowTextBelowIcons = True
            End If
        End If

        For Each cbrControl In Me.cbrMain(i).Controls
            cbrControl.Style = intStyle
        Next
    Next

    control.Checked = Not control.Checked
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_ToolBar_Size_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'ͼ��
    On Error GoTo errHandle

    Me.cbrMain.Options.LargeIcons = Not Me.cbrMain.Options.LargeIcons
    control.Checked = Not control.Checked

    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub Menu_View_StatusBar_click(ByVal control As XtremeCommandBars.ICommandBarControl)
'״̬��
    On Error GoTo errHandle

    Me.stbThis.Visible = Not Me.stbThis.Visible
    control.Checked = Not control.Checked
    picSchemeName.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    picSchemeContent.Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    Me.cbrMain.RecalcLayout
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Function IsHaveScheme(strName As String) As Boolean
    Dim i As Long

    IsHaveScheme = False
    For i = 1 To vsfMain.Rows - 1
        If UCase(Trim(vsfMain.TextMatrix(i, ColTitle.ct��������))) = UCase(Trim(strName)) Then
            Exit Function
        End If
    Next
    IsHaveScheme = True
End Function

Private Sub SchemeNo()
'�����������
    Dim i As Long

    If vsfMain.Rows < 2 Then Exit Sub
    For i = 1 To vsfMain.Rows - 1
        vsfMain.TextMatrix(i, ColTitle.ct���) = i
    Next
End Sub

Private Function IsEdit() As Boolean
'�жϷ��������Ƿ����ı�
    IsEdit = False
    
    If mblnClose Then
        If mblnIsEdit Or mobjQuerySet.mblnIsEdit Or mobjDisPlaySet.mblnIsEdit Or mobjFilterSet.mblnIsEdit Then
            IsEdit = True
        End If
    End If
End Function

Private Sub SetNoneEdit()
    mblnIsEdit = False
    mobjQuerySet.mblnIsEdit = False
    mobjDisPlaySet.mblnIsEdit = False
    mobjFilterSet.mblnIsEdit = False
End Sub

Private Function IsHaveIcon(strName As String, rsRecord As Recordset) As Boolean
    IsHaveIcon = False
    If rsRecord.RecordCount < 1 Then
        IsHaveIcon = False
        Exit Function
    End If
    rsRecord.MoveFirst
    Do While Not rsRecord.EOF
        If UCase(Trim(strName)) = UCase(Trim(NVL(rsRecord.Fields!��Դ����))) Then
            IsHaveIcon = True
            Exit Function
        End If
        rsRecord.MoveNext
    Loop
End Function

Public Function ShowUserScheme() As Boolean
'��ʾ�û����ò�ѯ��������...
'����е�������true��û�е�������false
    Dim objQueryCfg As New frmUserQueryReleation

    On Error GoTo errHandle

    ShowUserScheme = objQueryCfg.ShowUserScheme(Me, mlngModuleNo, 0)
    
    If Not objQueryCfg Is Nothing Then Unload objQueryCfg
    Set objQueryCfg = Nothing
    
    Exit Function
errHandle:
    If Not objQueryCfg Is Nothing Then Unload objQueryCfg
    Set objQueryCfg = Nothing
    
    Err.Raise -1, "clsPacsQuery.ShowUserScheme", "�û���ѯ������������ʧ��:" & Err.Description
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
    Dim bytFontSize As Byte
    Dim CtlFont As StdFont
    '���������С
    gbytFontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, IIf(bytSize = 2, 15, bytSize)))
    bytFontSize = gbytFontSize
    Set CtlFont = New StdFont
    CtlFont.Name = "����"
    'frmSchemeCfg
    labDeptFilter.FontSize = bytFontSize
    cbxDeptFilter.FontSize = bytFontSize
    vsfMain.FontSize = bytFontSize
    labObj.FontSize = bytFontSize
    labDept.FontSize = bytFontSize
    txtSchemeName.FontSize = bytFontSize
    cbxDept.FontSize = bytFontSize
    labSchemeMemo.FontSize = bytFontSize
    txtSchemeMemo.FontSize = bytFontSize
    chkTrance.FontSize = bytFontSize
    lblQueryDay.FontSize = bytFontSize
    txtDay.FontSize = bytFontSize
    lblDay.FontSize = bytFontSize
    lblDate.FontSize = bytFontSize
    txtDate.FontSize = bytFontSize
    lblYears.FontSize = bytFontSize
    lblRefreshTime.FontSize = bytFontSize
    cboRefreshTime.FontSize = bytFontSize
    
    CtlFont.Size = bytFontSize
    chkLocate.FontSize = bytFontSize
    chkRealTimeFilter.FontSize = bytFontSize
    chkFindRealTimeFilter.FontSize = bytFontSize
    chkTab.FontSize = bytFontSize
    chkEmbedFind.FontSize = bytFontSize
    labFore.FontSize = bytFontSize
    
    chkSelRowTransparent.FontSize = bytFontSize
    labPatiTypeColor.FontSize = bytFontSize
    cboColor.FontSize = bytFontSize
    
    If bytFontSize = 15 Then
        cmdFore.FontSize = 14.5
        cmdBack.FontSize = 14.5
    Else
        cmdFore.FontSize = bytFontSize
        cmdBack.FontSize = bytFontSize
    End If
    
    If bytFontSize = 9 Then
        picBasic.Height = 4355
        vsfMain.ColWidth(ColTitle.ct���) = 500
    ElseIf bytFontSize = 12 Then
        picBasic.Height = 4655
        vsfMain.ColWidth(ColTitle.ct���) = 650
    Else
        picBasic.Height = 4955
        vsfMain.ColWidth(ColTitle.ct���) = 800
    End If
    
    
    Set tbcMain.PaintManager.Font = CtlFont
    
    Set cbrMain.Options.Font = CtlFont
    
    Call picSchemeContent_Resize
    Call picSchemeName_Resize
    
    If Not mobjQuerySet Is Nothing Then
        mobjQuerySet.SetFontSize bytFontSize
    End If
    
    If Not mobjFilterSet Is Nothing Then
        mobjFilterSet.SetFontSize bytFontSize
    End If
    
    If Not mobjDisPlaySet Is Nothing Then
        mobjDisPlaySet.SetFontSize bytFontSize
    End If
    
    Call picBasic_Resize
End Sub

Private Sub CopyScheme()
    Dim lngRow As Long
    
    lngRow = vsfMain.Row
    If lngRow <= 0 Then Exit Sub
    
    NewRow vsfMain
    vsfMain.RowData(vsfMain.Rows - 1) = vsfMain.RowData(lngRow)
    Call ShowScheme
    
    txtSchemeName.Text = txtSchemeName.Text & "��������"
    
    If cbxDeptFilter.ListIndex > 0 Then
        cbxDept.ListIndex = cbxDeptFilter.ListIndex
    Else
        cbxDept.ListIndex = 0
    End If
    
    mblnNewState = True
    mblnClose = True
    
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
End Sub

Private Sub ShowStateBar(ByVal strHint As String)
'״̬����ʾ
    stbThis.Panels(2).Text = strHint
End Sub

Private Sub initWindow()
    On Error Resume Next
    Dim lngLeft As Long
    Dim lngTop As Long
    
    lngLeft = (Screen.Width - Me.Width) / 2
    lngTop = (Screen.Height - Me.Height) / 2

    Call Me.Move(lngLeft, lngTop)
    Me.WindowState = 0
    
End Sub

Private Function VerBeforeSaveScheme(ByVal strSql As String) As Boolean
'����ǰ��Ҫ������֤�����⣬����frmScheme_FilterCfg�б��е�����
On Error GoTo errH
    '�ж�where ���� filter ���Ƿ����
    Dim strCurPara As String
    Dim strSearchNames As String
    Dim i As Integer
    Dim objSqlParse As New clsSqlParse
    Dim strSysPara As String
    
    VerBeforeSaveScheme = True
    If Not mobjSqlScheme Is Nothing Then
        strSysPara = gstrPara & gstrBasePara
        strSysPara = Replace(strSysPara, "[", ";")
        strSysPara = Replace(strSysPara, "]", ";")
        Call objSqlParse.init(strSql)

        strSearchNames = mobjFilterSet.GetSearchNames()
        
        For i = 1 To objSqlParse.SqlStruct.ParCount
            strCurPara = objSqlParse.SqlStruct.AllParameter(i)
                
            If (InStr(strCurPara, "[@") > 0) Or (InStr(strCurPara, "[*") > 0) Then
                strCurPara = Mid$(strCurPara, 3, InStr(strCurPara, ",") - 3)
            Else
                strCurPara = Mid(strCurPara, 2, Len(strCurPara) - 2)
            End If
            strCurPara = ";" & strCurPara & ";"
            
            If InStr(strSearchNames, strCurPara) = 0 Then

                If Not (InStr(strSysPara, strCurPara) > 0 And InStr(strCurPara, "ϵͳ") > 0) Then
                    strCurPara = Replace(strCurPara, ";", "")
                    MsgBox "SQL�е�����[" & strCurPara & "]�ڲ�ѯ��������-����¼��������δ�ҵ���Ӧ���ã����顣", vbInformation, Me.Caption
                    VerBeforeSaveScheme = False
                    Exit Function
                End If
            End If
        Next
    End If
    Exit Function
errH:
    MsgBox Err.Description, vbExclamation, Me.Caption
End Function

Private Function IsDefault() As Boolean
'���أ���ǰ�б�ѡ�з�����Ĭ�� ����Ӱ���ѯ�����б�����ΪĬ��
On Error GoTo errH
    Dim strSql As String
    Dim rsTmp As Recordset
    
    IsDefault = False
    
    If vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) <> "" Then
        IsDefault = True
        Exit Function
    End If
    
    IsDefault = True
    strSql = "Select ID From Ӱ���ѯ���� Where �Ƿ�Ĭ�� = 1 And ��ѯ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "�Ƿ�Ĭ�Ϸ���", Val(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)))
    
    If rsTmp.EOF Then IsDefault = False
    Exit Function
errH:
    IsDefault = True
End Function

Private Sub initCboColor(ByRef objSqlScheme As clsSqlScheme)
On Error GoTo errH
    Dim i As Long
    Dim rsTmp As Recordset
    
    Dim objSqlParse As New clsSqlParse
    Dim objQuery As New clsPacsQuery
    
    objSqlParse.init mobjQuerySet.GetQuerySql
    Set rsTmp = objQuery.GetQueryField(objSqlParse)
    If rsTmp Is Nothing Then
        Exit Sub
    End If
    
    cboColor.Clear
    Call cboColor.AddItem("")
    
    For i = 0 To rsTmp.Fields.Count - 1
        Call cboColor.AddItem(rsTmp.Fields(i).Name)
    Next
    Call zlControl.CboSetIndex(cboColor.hwnd, 0)
    
    Exit Sub
errH:
    MsgBox Err.Description, vbExclamation, Me.Caption
End Sub
