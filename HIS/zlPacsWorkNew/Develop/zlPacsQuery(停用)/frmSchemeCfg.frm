VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSchemeCfg 
   Caption         =   "��ѯ��������"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   15120
   Icon            =   "frmSchemeCfg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   15120
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2760
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picSchemeName 
      Height          =   5295
      Left            =   960
      ScaleHeight     =   5235
      ScaleWidth      =   2715
      TabIndex        =   17
      Top             =   1080
      Width           =   2775
      Begin VSFlex8Ctl.VSFlexGrid vsfMain 
         Height          =   4935
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   2415
         _cx             =   4260
         _cy             =   8705
         Appearance      =   1
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
   End
   Begin VB.PictureBox picSchemeContent 
      Height          =   7215
      Left            =   3960
      ScaleHeight     =   7155
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   480
      Width           =   10935
      Begin VB.Frame fraBasic 
         Height          =   1335
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   10575
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
            TabIndex        =   1
            Top             =   270
            Width           =   2475
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
            Height          =   375
            Left            =   4680
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtPage 
            Height          =   270
            Left            =   6690
            TabIndex        =   7
            Text            =   "0"
            Top             =   892
            Width           =   495
         End
         Begin VB.TextBox txtDate 
            Height          =   270
            Left            =   8730
            TabIndex        =   8
            Text            =   "0"
            Top             =   892
            Width           =   495
         End
         Begin VB.CheckBox chkTrance 
            Caption         =   "�б��ܸ���"
            Height          =   375
            Left            =   4080
            TabIndex        =   6
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkHistory 
            Caption         =   "��ʾ�����ʷ"
            Height          =   375
            Left            =   2520
            TabIndex        =   5
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkGroup 
            Caption         =   "���÷���"
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   840
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkCard 
            Caption         =   "����ˢ��"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.Frame fraLine 
            Height          =   30
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   10215
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
            Left            =   3720
            TabIndex        =   16
            Top             =   330
            Width           =   975
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
            TabIndex        =   15
            Top             =   330
            Width           =   975
         End
         Begin VB.Label lblPage 
            AutoSize        =   -1  'True
            Caption         =   "��ҳ��С��"
            Height          =   180
            Left            =   5760
            TabIndex        =   14
            Top             =   937
            Width           =   900
         End
         Begin VB.Label lblRows 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   7290
            TabIndex        =   13
            Top             =   937
            Width           =   180
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            Caption         =   "���ڷ�Χ��"
            Height          =   180
            Left            =   7770
            TabIndex        =   12
            Top             =   937
            Width           =   900
         End
         Begin VB.Label lblYears 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Left            =   9330
            TabIndex        =   11
            Top             =   937
            Width           =   180
         End
      End
      Begin XtremeSuiteControls.TabControl tbcMain 
         Height          =   4095
         Left            =   480
         TabIndex        =   20
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
      TabIndex        =   19
      Top             =   7860
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSchemeCfg.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12938
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
      Bindings        =   "frmSchemeCfg.frx":70E6
      Left            =   480
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSchemeCfg.frx":70FA
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
Private mlngModuleNo As Long     '����ģ��
Private mblnNewState As Boolean     '�Ƿ�������
Private mblnClose As Boolean
Private mblnIsEdit As Boolean   '�����Ƿ��ѱ༭
Private mlngRow As Long
Private mobjIconManage As New frmIconManage
Private mobjQuerySet As New frmScheme_BaseQueryCfg
Private mobjFilterSet As New frmScheme_FilterCfg
Private mobjDisPlaySet As New frmScheme_DisplayCfg
Private mobjSetRelated As New frmSetRelated
Private mobjSqlScheme As New clsSqlScheme

Private Const SW_SHOW = 5
Private Const M_STR_GRIDDATA = "���|ID|��������|ʹ��״̬|�Ƿ�Ĭ��|�Ƿ���|����˵��"       '�����ʾ
Private Const M_STR_CROOK = "��"
Private Enum ColTitle
    ct��� = 0
    ctID = 1
    ct�������� = 2
    ctʹ��״̬ = 3
    ct�Ƿ�Ĭ�� = 4
    ct�Ƿ��� = 5
    ct����˵�� = 6
End Enum

Private Const conMenu_View_ToolBar = 301              '������(&T)
Private Const conMenu_View_ToolBar_Button = 302         '��׼��ť(&S)
Private Const conMenu_View_ToolBar_Text = 303           '�ı���ǩ(&T)
Private Const conMenu_View_ToolBar_Size = 304           '��ͼ��(&B)
Private Const conMenu_View_StatusBar = 305            '״̬��(&S)

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
    
    mtViewPopup = 3            '�鿴
    
'    mtHelpPopup = 4            '����
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub ShowMe(ByVal lngModuleNo As Long, ByVal strSysPara As String, owner As Object)
    mlngModuleNo = lngModuleNo
    gstrPara = strSysPara
    Me.Show , owner
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
            
''--------------------------����-----------------
'        Case conMenu_Help_Help
'            Call Menu_Help_Help_click
'        Case conMenu_Help_Web_Forum
'            Call Menu_Help_Web_Forum_click
'        Case conMenu_Help_Web_Home
'            Call Menu_Help_Web_Home_click
'        Case conMenu_Help_Web_Mail
'            Call Menu_Help_Web_Mail_click
'        Case conMenu_Help_About
'            Call Menu_Help_About_click
    End Select
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub cbrMain_Update(ByVal control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo errHandle

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
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtDelScheme      'ɾ��
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtRecover
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
        Case TMenuType.mtUseScheme
            control.Enabled = Not (mblnClose Or vsfMain.Rows <= 1)
            If IsSelectionRow(vsfMain) Then
                control.Caption = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", "����", "����")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", 211, 207)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtUsually
            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����")
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)) = 0, "��     ��", "ȡ������")
                control.IconId = IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) = M_STR_CROOK, 212, 204)
            Else
                control.Enabled = False
            End If
        Case TMenuType.mtSetDefault

            If IsSelectionRow(vsfMain) Then
                control.Enabled = Not (mblnClose Or vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����")
                control.Caption = IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)) = 0, "Ĭ     ��", "ȡ��Ĭ��")
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

Private Sub chkCard_Click()
    On Error GoTo errHandle
    
    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkGroup_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub chkHistory_Click()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
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

Private Sub Form_Load()
    On Error GoTo errHandle
    
    Me.ZOrder
    Call InitCommandBars
    Call InitDockPannel
    Call InitTabControl

    Call GridInit(M_STR_GRIDDATA, vsfMain)
    Call RefreshWindowState(False)
    Call RefreshList
    Call RefreshScheme
    Call ShowScheme
    
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
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
        .SetIconSize True, 24, 24                               '���ô�ͼ��ĳߴ�
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
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "Ĭ��(&F)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "����(&Y)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "�ָ�(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "����(&V)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����(&L)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����(&X)")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetSysQuery, "�û���ѯ����(&U)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtResource, "��Դ����(&Z)")
    End With
    
    
    'Begin----------------------�鿴�˵�--------------------------------------
    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mtViewPopup, "�鿴(V)")
    Call CreateViewAndHelpMenu(cbrMenuBar, Nothing)
    
'    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
'    Set cbrMenuBar = cbrMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mtHelpPopup, "����(H)")
'    Call CreateViewAndHelpMenu(Nothing, cbrMenuBar)
    
    
    '---------------------����������------------------------------------------
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSave, "����", "���淽��")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCancel, "�ر�", "�رձ༭")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtNewScheme, "����", "��������"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtModifyScheme, "�޸�", "�޸ķ���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtDelScheme, "ɾ��", "ɾ������")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtSetDefault, "Ĭ��", "����Ĭ�Ϸ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUsually, "����", "���÷���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtRecover, "�ָ�", "�ָ�����")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtUseScheme, "����", "���÷���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveLastScheme, "����", "���Ʒ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtMoveNextScheme, "����", "���Ʒ���")
        Set cbrControl = .Add(xtpControlButton, TMenuType.mtCheckScheme, "����", "���Է���"): cbrControl.BeginGroup = True
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

'    'Begin----------------------�����˵�--------------------------------------Ĭ�Ͽɼ�
'    If Not (objHelpMenu Is Nothing) Then
'        Set cbrMenuBar = objHelpMenu
'
'        With cbrMenuBar.CommandBar
'            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_Help, "��������(M)")
'                cbrControl.Category = strMenuTag
'
'            Set cbrControl = .Controls.Add(xtpControlButtonPopup, conMenu_Help_Web, "WEB�ϵ�����(W)")
'                cbrControl.Category = strMenuTag
'
'                With cbrControl.CommandBar
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Forum, "������̳(0)")
'                        cbrPopControl.Category = strMenuTag
'
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Home, "������ҳ(1)")
'                        cbrPopControl.Category = strMenuTag
'
'                    Set cbrPopControl = .Controls.Add(xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(2)")
'                        cbrPopControl.Category = strMenuTag
'                End With
'
'            Set cbrControl = .Controls.Add(xtpControlButton, conMenu_Help_About, "���ڡ�(A)")
'                cbrControl.Category = strMenuTag
'        End With
'    End If
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
    
    On Error GoTo errHandle

    dkpMain.SetCommandBars cbrMain
    
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
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

Private Sub picSchemeContent_Resize()
    On Error Resume Next

    fraBasic.Move picSchemeContent.Left, picSchemeContent.Top - 100, picSchemeContent.Width
    tbcMain.Move picSchemeContent.Left, picSchemeContent.Top + fraBasic.Height - 100, picSchemeContent.Width, picSchemeContent.Height - fraBasic.Height + IIf(Not stbThis.Visible, stbThis.Height, 0) + 100
    fraLine.Move fraBasic.Left, fraBasic.Top + fraBasic.Height / 2 + 200, fraBasic.Width
    txtSchemeMemo.Move txtSchemeMemo.Left, txtSchemeMemo.Top, fraBasic.Width - txtSchemeMemo.Left - 500
End Sub

Private Sub picSchemeName_Resize()
    On Error Resume Next

    vsfMain.Move picSchemeName.Left, picSchemeName.Top, picSchemeName.Width, picSchemeName.Height - IIf(stbThis.Visible, stbThis.Height, 0)
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
    
        strResult = SqlVerify(mobjQuerySet.GetQuerySql)
        
        If Len(strResult) = 0 Then
            strResult = IsHaveID(mobjQuerySet.GetQuerySql)
        End If
        
        If Len(strResult) > 0 Then
            MsgBox "��ѯ�����֤ʧ�ܣ�" & vbCrLf & "ԭ���ǣ�" & strResult, vbInformation, Me.Caption
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
    strSql = "select '' ���,ID,��������,ʹ��״̬,�Ƿ�Ĭ��,�Ƿ���,����˵�� from Ӱ���ѯ���� where ����ģ�� = [1] Order By �������"
    Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", mlngModuleNo)
    Set vsfMain.DataSource = rsData
    vsfMain.ColHidden(ColTitle.ctID) = True
    
    Call DataConvert
    Call SchemeNo

    vsfMain.ColWidth(ColTitle.ct���) = 500
    vsfMain.ColWidth(ColTitle.ct��������) = 2000
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

Private Sub txtPage_Change()
    On Error GoTo errHandle

    If mblnClose Then
        mblnIsEdit = True
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, vbExclamation, Me.Caption
    Err.Clear
End Sub

Private Sub txtPage_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If InStr(1, "0123456789", Chr(KeyAscii)) = 0 And Chr(KeyAscii) <> vbBack Then
        KeyAscii = 0
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
    Call RefreshWindowState(True)
    Call RefreshSubWindowState(True)
    Call SetNoneEdit
End Sub

Private Sub SaveScheme()
    Dim strSql As String
    Dim strText As String
    Dim rsData As Recordset

    If Not IsEnabledToSvae Then Exit Sub
    strText = GetSchemeContent
    If mblnNewState Then
        strSql = "Zl_Ӱ���ѯ_��������('" & Replace(txtSchemeName.Text, "'", "''") & "','" & Replace(txtSchemeMemo.Text, "'", "''") & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) & "','" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) & "','" & mlngModuleNo & "','" & Replace(strText, "'", "''") & "')"
    Else
        strSql = "Zl_Ӱ���ѯ_���·���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & ",'" & Replace(txtSchemeName.Text, "'", "''") & "','" & Replace(txtSchemeMemo.Text, "'", "''") & "'," & IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)) = 0, 0, 1) & "," & IIf(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = "����", 1, 0) & "," & IIf(Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)) = 0, 0, 1) & "," & mlngModuleNo & ",'" & Replace(strText, "'", "''") & "')"
    End If
    Call ExecuteCmd(strSql, "�༭����")
    If mblnNewState Then
        strSql = "select ID from Ӱ���ѯ���� where �������� = [1] and ����ģ�� = [2]"
        Set rsData = ExecuteSql(strSql, "��ѯ������Ϣ", txtSchemeName.Text, mlngModuleNo)
        If rsData.RecordCount < 1 Then
            Exit Sub
        End If
        vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) = rsData.Fields!Id

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

End Sub

Private Function GetSchemeContent() As String
    Dim objSqlScheme As clsSqlScheme
    Dim objScSearchCfg As clsScSerachCfg
    Dim strText As String
    Dim strQuery As String
    Dim strDetail As String


    Set objSqlScheme = New clsSqlScheme

    '������Ϣ
    With objSqlScheme
        If Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)) > 0 Then
            .SchemeId = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)
        End If
        .UseCard = chkCard.value
        .UseGroup = chkGroup.value
        .ShowHistory = chkHistory.value
        .UseFuncFollow = chkTrance.value
        .PageRecord = Val(txtPage.Text)
        .DataRange = Val(txtDate.Text)
        .SchemeName = txtSchemeName
        .Descript = txtSchemeMemo
        .Store = mobjSqlScheme.Store
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
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveUp(vsfMain)

    strSql = "zl_Ӱ���ѯ_�ƶ�����(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "���Ʒ���")
    Call SchemeNo
End Sub



Private Sub MoveNextScheme()
'���Ʒ���
    Dim strSql As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    Call MoveDown(vsfMain)

    strSql = "zl_Ӱ���ѯ_�ƶ�����(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & vsfMain.Row & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "���Ʒ���")
    Call SchemeNo
End Sub

Private Sub SetDefaultScheme()
'����Ĭ�Ϸ���
    Dim strSql As String
    Dim strCurDefaultState As String
    Dim i As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurDefaultState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��)
    strSql = "Zl_Ӱ���ѯ_Ĭ�Ϸ���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurDefaultState = M_STR_CROOK, 0, 1) & "," & mlngModuleNo & ")"
    Call ExecuteCmd(strSql, "����Ĭ��")
    
    vsfMain.Cell(flexcpText, 1, ColTitle.ct�Ƿ�Ĭ��, vsfMain.Rows - 1, ColTitle.ct�Ƿ�Ĭ��) = ""
    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ�Ĭ��) = IIf(Len(strCurDefaultState) = 0, M_STR_CROOK, "")
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUsualScheme()
'�����Ƿ���
    Dim strSql As String
    Dim strCurUsualState As String
    Dim i As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUsualState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���)
    strSql = "Zl_Ӱ���ѯ_���÷���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUsualState = M_STR_CROOK, 0, 1) & ")"
    Call ExecuteCmd(strSql, "���ó���")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ct�Ƿ���) = IIf(Len(strCurUsualState) = 0, M_STR_CROOK, "")
'    cbrMain.RecalcLayout
End Sub

Private Sub SetUseScheme()
'���÷���
    Dim strSql As String
    Dim strCurUseState As String

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    strCurUseState = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬)

    strSql = "Zl_Ӱ���ѯ_���÷���(" & vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID) & "," & IIf(strCurUseState = "����", 0, 1) & ")"
    Call ExecuteCmd(strSql, "ʹ��״̬����")

    vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctʹ��״̬) = IIf(strCurUseState = "����", "����", "����")

End Sub


Private Sub DeleteScheme()
'ɾ��ѡ�з���
    Dim strSql As String
    Dim lngRow As Long

    If Not IsSelectionRow(vsfMain) Then
        MsgBox Me, "��ѡ����Ҫ�����Ĳ�ѯ������", vbOKOnly, Me.Caption
        Exit Sub
    End If

    If Not MsgBox("�Ƿ�ɾ��ѡ�з�����", vbYesNo, Me.Caption) = vbYes Then
        Exit Sub
    End If

    lngRow = vsfMain.Row
    strSql = "Zl_Ӱ���ѯ_ɾ������(" & vsfMain.TextMatrix(lngRow, ColTitle.ctID) & ")"
    Call ExecuteCmd(strSql, "ɾ������")

    vsfMain.RemoveItem (lngRow)
End Sub

Private Sub CancelScheme()
    If IsEdit Then
        If MsgBox("������δ���棬�Ƿ񱣴棿", vbYesNo, Me.Caption) = vbYes Then
            Call SaveScheme
        End If
    End If
    If mblnNewState Then
        vsfMain.RemoveItem vsfMain.Row
    End If
    Call ShowScheme
    Call RefreshShowScheme
    mblnNewState = False
    mblnClose = False
    Call RefreshWindowState(False)
    Call RefreshSubWindowState(False)
    
'    Call RefreshScheme
End Sub

Private Sub RefreshWindowState(blnState As Boolean)
    vsfMain.Enabled = Not blnState
    txtSchemeName.Enabled = blnState
    txtSchemeMemo.Enabled = blnState
    txtDate.Enabled = blnState
    txtPage.Enabled = blnState
    chkCard.Enabled = blnState
    chkGroup.Enabled = blnState
    chkHistory.Enabled = blnState
    chkTrance.Enabled = blnState
End Sub

Private Sub InitScheme()
    chkCard.value = 1
    chkHistory.value = 1
    chkGroup.value = 1
    chkTrance.value = 1
    txtPage.Text = 0
    txtDate.Text = 0
    txtSchemeName.Text = ""
    txtSchemeMemo.Text = ""

    tbcMain.Item(0).Selected = True
End Sub


Private Sub RefreshScheme()
'ˢ�·���
    Dim i As Long
    Dim lngNumber As Long
    Dim strSchemeXml As String
'
'    Set mobjSqlScheme = Nothing
'    If Not IsSelectionRow(vsfMain) Then
'        Exit Sub
'    End If
'    If Len(vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)) = 0 Then
'        Exit Sub
'    End If
'
'    lngNumber = vsfMain.TextMatrix(vsfMain.Row, ColTitle.ctID)
'
'    strSchemeXml = ReadSchemeXml(lngNumber, "")
'
'    Call mobjSqlScheme.OpenScheme(strSchemeXml)
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
    Call mobjSqlScheme.OpenScheme(vsfMain.RowData(vsfMain.Row))
    With mobjSqlScheme
        chkCard.value = IIf(.UseCard, 1, 0)
        chkHistory.value = IIf(.ShowHistory, 1, 0)
        chkGroup.value = IIf(.UseGroup, 1, 0)
        chkTrance.value = IIf(.UseFuncFollow, 1, 0)
        txtPage.Text = .PageRecord
        txtDate.Text = .DateRange
        txtSchemeName.Text = .SchemeName
        txtSchemeMemo.Text = .Descript

        '��ʾ��ѯ���ģ��
        Call mobjQuerySet.ShowQuerySet(mobjSqlScheme)
        '��ʾ���ٹ���ģ��
        Call mobjFilterSet.ShowFilterSet(mobjSqlScheme)
        '��ʾ��ʾ����ģ��
        Call mobjDisPlaySet.ShowDisplaySet(mobjSqlScheme)
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

Private Function IsEnabledToSvae() As Boolean
'��������
    Dim i As Long
    Dim strResult As String

    IsEnabledToSvae = False
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

    strResult = SqlVerify(mobjQuerySet.GetQuerySql)
    If Len(strResult) = 0 Then
        strResult = IsHaveID(mobjQuerySet.GetQuerySql)
    End If
    If Len(strResult) > 0 Then
        MsgBox "��ѯ�����֤ʧ�ܣ�ԭ��Ϊ��" & strResult, vbInformation, Me.Caption
        Exit Function
    End If

    If Not mobjFilterSet.IsEnabledSave Then
        Exit Function
    End If

    IsEnabledToSvae = True
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

    objExportScheme.ShowMe mlngModuleNo, True, arrID, strFile, blnIcon, Me
    Set objExportScheme = Nothing
End Sub

Private Sub ImportScheme()
'����
    Dim objExportScheme As New frmExportScheme
    Dim arrID() As Long
    Dim strFile As String
    Dim blnIcon As Boolean

    dlgFile.Filter = "(*.XML)|*.XML|(*.*)|*.*"
    dlgFile.DefaultExt = "*.XML"

    dlgFile.FileName = ""
    dlgFile.ShowOpen

    If dlgFile.FileName = "" Then Exit Sub
    strFile = dlgFile.FileName

    If objExportScheme.ShowMe(mlngModuleNo, False, arrID, strFile, blnIcon, Me) Then
        Call ImportContent(arrID, strFile, blnIcon)
        Set objExportScheme = Nothing
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
                        If MsgBox("�Ѵ�����Ϊ����" & strSchemeName & "���ķ���,�Ƿ��������", vbYesNo, Me.Caption) = vbYes Then
                            Do While True
                                strSchemeName = strSchemeName & lngCount
                                If IsHaveScheme(strSchemeName) Then
                                    Exit Do
                                End If

                                lngCount = lngCount + 1
                            Loop
                        Else
                            blnIsImport = False
                        End If
                    End If
                Next
                If blnIsImport Then
                    strText = strText & rsData.Fields(3).value
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
    If cbrMain.FindControl(, TMenuType.mtSave).Enabled Then
        If MsgBox("��δ���淽�����Ƿ��˳���", vbYesNo, Me.Caption) = vbNo Then
            Exit Sub
        End If
    End If

    mobjFilterSet.UnloadMe
    mobjDisPlaySet.UnloadMe
    mobjSetRelated.UnloadMe
    
    Set mobjQuerySet = Nothing
    Set mobjFilterSet = Nothing
    Set mobjDisPlaySet = Nothing
    Set mobjSqlScheme = Nothing
    Set mobjSetRelated = Nothing
    Set mobjIconManage = Nothing

    Unload Me
End Sub

Private Sub RecoverScheme()
    Dim strStore As String

    If MsgBox("�Ƿ�ȷ���ָ�������", vbYesNo, Me.Caption) = vbNo Then
        Exit Sub
    End If
    strStore = mobjSqlScheme.Store
    If Len(strStore) < 1 Then
        MsgBox "�÷���û�����ûָ����ԣ��޷��ָ�", vbInformation, Me.Caption
        Exit Sub
    End If
    Call mobjSqlScheme.OpenScheme(strStore)
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
'
'Private Sub Menu_Help_Help_click()
''���ܣ����ð�������
'    On Error GoTo errHandle
'
'    ShowHelp App.ProductName, Me.hWnd, Me.Name
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Forum_click()
'    On Error GoTo errHandle
'
'    Call zlWebForum(Me.hWnd)
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Home_click()
'    On Error GoTo errHandle
'
'    zlHomePage hWnd
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_Web_Mail_click()
'    On Error GoTo errHandle
'
'    zlMailTo hWnd
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub
'
'Private Sub Menu_Help_About_click()
'    On Error GoTo errHandle
'
'    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
'    Exit Sub
'errHandle:
'    MsgBox Err.Description, vbExclamation, Me.Caption
'    Err.Clear
'End Sub

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
    Exit Function
errHandle:
    Set objQueryCfg = Nothing
    Err.Raise -1, "clsPacsQuery.ShowUserScheme", "�û���ѯ������������ʧ��:" & Err.Description
End Function

