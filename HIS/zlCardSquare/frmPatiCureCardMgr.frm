VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmPatiCureCardMgr 
   Caption         =   "ҽ�ƿ�����"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   Icon            =   "frmPatiCureCardMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   12405
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   5700
      ScaleHeight     =   1875
      ScaleWidth      =   5070
      TabIndex        =   25
      Top             =   5115
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   8520
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiCureCardMgr.frx":1CFA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12885
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   180
      Top             =   675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardMgr.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiCureCardMgr.frx":28E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6765
      Left            =   0
      ScaleHeight     =   6765
      ScaleWidth      =   3690
      TabIndex        =   29
      Top             =   1110
      Width           =   3690
      Begin VB.TextBox txtName 
         Height          =   350
         Left            =   960
         TabIndex        =   5
         Top             =   960
         Width           =   2445
      End
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   0
         Left            =   960
         TabIndex        =   17
         Top             =   4215
         Width           =   2445
      End
      Begin VB.TextBox txtEdit 
         Height          =   350
         Index           =   1
         Left            =   960
         TabIndex        =   19
         Top             =   4635
         Width           =   2445
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   2325
         TabIndex        =   20
         Top             =   5100
         Width           =   1100
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "����ʧʱ�����(&G)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   2925
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   8
         Top             =   1920
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "���������ڲ���(&S)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   1590
         Value           =   1  'Checked
         Width           =   2745
      End
      Begin VB.TextBox txtCard 
         Height          =   350
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   525
         Width           =   2445
      End
      Begin VB.TextBox txtCard 
         Height          =   350
         Index           =   0
         Left            =   960
         TabIndex        =   1
         Top             =   80
         Width           =   2445
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   10
         Top             =   2310
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp��ʼ���� 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   13
         Top             =   3300
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin MSComCtl2.DTPicker dtp�������� 
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   3735
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   81526787
         CurrentDate     =   40722
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   2
         Left            =   390
         TabIndex        =   16
         Top             =   4290
         Width           =   540
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "��ʧ��"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   18
         Top             =   4710
         Width           =   540
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "��������(&M)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   14
         Top             =   3780
         Width           =   990
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����(&L)"
         Height          =   180
         Index           =   1
         Left            =   375
         TabIndex        =   12
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label lblEndDate 
         AutoSize        =   -1  'True
         Caption         =   "��������(&E)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   9
         Top             =   2370
         Width           =   990
      End
      Begin VB.Label lblStartDate 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����(&S)"
         Height          =   180
         Index           =   0
         Left            =   375
         TabIndex        =   7
         Top             =   1980
         Width           =   990
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   2
         Top             =   615
         Width           =   720
      End
      Begin VB.Label lblEDIT 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ����"
         Height          =   180
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   165
         Width           =   720
      End
   End
   Begin VB.PictureBox picCardList 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   4305
      ScaleHeight     =   2565
      ScaleWidth      =   8175
      TabIndex        =   27
      Top             =   1410
      Width           =   8175
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   22
         Top             =   0
         Width           =   2280
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCardList 
         Height          =   2055
         Left            =   300
         TabIndex        =   23
         Top             =   465
         Width           =   6825
         _cx             =   12039
         _cy             =   3625
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPatiCureCardMgr.frx":2C36
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   28
            Top             =   60
            Width           =   210
            Begin VB.Image imgCol 
               Height          =   195
               Left            =   0
               Picture         =   "frmPatiCureCardMgr.frx":2C8F
               ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   330
         Left            =   720
         TabIndex        =   30
         Top             =   0
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   582
         Appearance      =   2
         IDKindStr       =   $"frmPatiCureCardMgr.frx":31DD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "����"
         IDKind          =   -1
         DefaultCardType =   "0"
         BackColor       =   -2147483633
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "�ֿ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   21
         Top             =   75
         Width           =   630
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   495
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmPatiCureCardMgr.frx":32C0
      Left            =   1515
      Top             =   90
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPatiCureCardMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar, mcbrComboxToolBar As CommandBar
Private mblnFirst  As Boolean, mstrPrivs As String, mstrTitle As String    '���ܱ���
Private mlngModule As Long, mstrKey As String
Private Enum mPgIndex
    Pg_�䶯��¼ = 250101
    Pg_�ʻ������¼ = 250102
    Pg_������ϵ = 250103
End Enum
Private Enum mPaneID
    Pane_Search = 1     '��������
    Pane_CardLists = 2  '���б�
    Pane_CardDetails = 3    '��ϸ�б�
End Enum
Private Enum mtxtIdx
    idx_������ = 0
    idx_��ʧ�� = 1
End Enum
Private WithEvents mobjIDCard As zlIDCard.clsIDCard  '���֤�ӿ�
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard   'IC���ӿ�
Attribute mobjICCard.VB_VarHelpID = -1
Private mlngCardTypeID As Long
Private mPanSearch As Pane
Private mobjSubFrm As Collection
Private mArrFilter As Collection
Private mblnInited As Boolean
Private Const mconMenu_Lable = 3999
Private WithEvents mfrmChage As frmPatiCureCardChangeMgr
Attribute mfrmChage.VB_VarHelpID = -1
Private WithEvents mfrmConsume As frmPatiCureCardConsumeMgr
Attribute mfrmConsume.VB_VarHelpID = -1
Private WithEvents mfrmFamily As frmPatiCureCardFamilyMgr
Attribute mfrmFamily.VB_VarHelpID = -1
Private mcolCard As Collection
Private mblnNotRefresh As Boolean  '��ˢ������
Private mblnNotClick As Boolean
Private mstrPrepayPrivs As String 'Ԥ��������Ȩ��
Private mlng�����ID As Long
Private mbln���ƿ� As Boolean '��ǰ�Ƿ����ƿ�
Private mbln�ظ�ʹ�� As Boolean '57899
Private mbln���� As Boolean '��ǰ�Ƿ񷢿�;�����:56599
Private mstrCurStatus As String '��ǰ״̬
Private mblnSeekName As Boolean '����ģ������
Private mintNameDays As Integer '������������

Private mlngCurPatient As Long '��ǰѡ�еĲ���ID '״̬���
'-------------------------------------------------------------------------
'����ش���
'Private mPatiCard As SquareCard 'ˢ�������
Private mstrPassWord As String
Private mobjPatiCardObject As clsCardObject
Private mblnDefaultPassInputCardNo As Boolean
'-------------------------------------------------------------------------
Private mstrListReportName As String    '������� ����50122
Private mlngListReportID As Long    '��� ����50122
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��
Private mobjPubPatient As Object
Private mstrPubPatiPrivs As String '����������ϢȨ��
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private Type Ty_PrintProperty
    intPrintMode As Integer '��ӡģʽ 3-�ش� ��4-����5-��ӡƾ��
    strUseType As String 'ʹ�����
    strInvoice As String 'Ʊ�ݺ�
    strPrintNo As String '�������ݺ�
    lng����ID As Long '����Ʊ������ID
    strBackInvoice As String  '����Ʊ��
    bytPrintPayCard As Byte
    bytPrintBoundCard As Byte
End Type
Private mPrint As Ty_PrintProperty

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2011-06-28 15:22:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsCardList
        'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        .ColData(.ColIndex("����")) = "1|0"
        .ColData(.ColIndex("��־")) = "-1|1"
        If .ColIndex("ID") >= 0 Then
            .ColData(.ColIndex("ID")) = "-1|1"
            .ColHidden(.ColIndex("ID")) = True
        End If
    End With
End Sub
Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2011-06-28 15:22:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:
    Set mobjSubFrm = New Collection
    
    Set mfrmChage = New frmPatiCureCardChangeMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�䶯��¼, "�䶯���", mfrmChage.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_�䶯��¼
    mobjSubFrm.Add mfrmChage, CStr(objItem.Tag)
    
    Set mfrmConsume = New frmPatiCureCardConsumeMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_�ʻ������¼, "�ʻ������Ϣ", mfrmConsume.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_�ʻ������¼
    mobjSubFrm.Add mfrmConsume, CStr(objItem.Tag)
    
    Set mfrmFamily = New frmPatiCureCardFamilyMgr
    Set objItem = tbPage.InsertItem(mPgIndex.Pg_������ϵ, "������ϵ", mfrmFamily.hWnd, 0)
    objItem.Tag = mPgIndex.Pg_������ϵ
    mobjSubFrm.Add mfrmFamily, CStr(objItem.Tag)
    
    mblnNotClick = True
     With tbPage
        tbPage.Item(i).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    mblnNotClick = False
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2009-11-18 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
     With dkpMan
        .ImageList = imlPaneIcons
        Set mPanSearch = .CreatePane(mPaneID.Pane_Search, 400, 400, DockLeftOf, Nothing)
        mPanSearch.Title = "��������": mPanSearch.Options = PaneNoCloseable
        mPanSearch.MinTrackSize.Width = picFilter.Width / Screen.TwipsPerPixelX
        mPanSearch.MaxTrackSize.Width = picFilter.Width / Screen.TwipsPerPixelX
        
        Set objPane = .CreatePane(mPaneID.Pane_CardLists, 400, 400, DockRightOf, mPanSearch)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        Set objPane = .CreatePane(mPaneID.Pane_CardDetails, 400, 400, DockBottomOf, objPane)
        objPane.Title = ""
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picList.hWnd
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    zlRestoreDockPanceToReg Me, dkpMan, "����"
End Function
Private Function zlIsHaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������
     '����:��ǰ�ؼ�������,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-28 18:17:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
    zlIsHaveData = False
    If Me.ActiveControl Is vsCardList Then
        zlIsHaveData = vsCardList.TextMatrix(1, vsCardList.ColIndex("����")) <> ""
    End If
End Function

Private Function zlIsCardBinding() As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:�Ƿ����ƿ��ظ�ʹ�ý��а󶨵��������
'����:
'����:
'����:
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
    zlIsCardBinding = False
    If Me.ActiveControl Is vsCardList Then
        With vsCardList
            If .TextMatrix(.Row, .ColIndex("����")) <> "" Then
                zlIsCardBinding = Val(.TextMatrix(.Row, .ColIndex("�䶯���"))) = 11 And Val(.TextMatrix(.Row, .ColIndex("�Ƿ��ظ�ʹ��"))) = 1 And _
                             mbln���ƿ�
            End If
        End With
    End If
End Function
 

Private Sub cmdFilter_Click()
    Call InitFilterToVar
    Call LoadDataToGrid
End Sub
Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
'    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case mPaneID.Pane_Search    '������������
        Item.Handle = picFilter.hWnd
    Case mPaneID.Pane_CardDetails   '��ϸ����Ϣ
        Item.Handle = picList.hWnd
    Case mPaneID.Pane_CardLists '���б�
        Item.Handle = picCardList.hWnd
    End Select
End Sub
Private Sub zl_OpenReport(ByVal lngSys As Long, ByVal strReportCode As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ������
    '���:lngSys-ϵͳ��
    '     strReportCode������
    '����:���˺�
    '����:2011-06-28 18:20:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���� As String
    Dim str������ As String, str�������� As String
    With vsCardList
        If .Row < 0 Then Exit Sub
        str���� = Trim(.TextMatrix(.Row, .ColIndex("����")))
        If str���� = "" Then Exit Sub
        
        str������ = Trim(.TextMatrix(.Row, .ColIndex("������")))
        str�������� = Trim(.TextMatrix(.Row, .ColIndex("��������")))
    End With
    Call ReportOpen(gcnOracle, lngSys, strReportCode, Me, "�����ID=" & mlngCardTypeID, "����=" & str����, "������=" & str������, "��������=" & str��������)
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
    
    cbsThis.EnableCustomization False
    '-----------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched)
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.id = conMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "�շ�����(&M)")
        mcbrControl.IconId = 227
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_PrintSingleBill, "�ش�ɿ(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_BarcodePrint, "�ش򷢿�ƾ��(&W)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.id = conMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����(&S)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "�󶨿�(&B)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelCardBound, "ȡ����(&C)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�(&T)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "����(&H)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "����(&F)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardLoss, "��ʧ(&G)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardCancelLoss, "ȡ����ʧ(&O)")
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��Ԥ��(&J)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFullBack, "��Ԥ��(&Y)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBackMoney, "����˿�(&R)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_MzToZy, "����תסԺ(&M)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ZyToMz, "סԺת����(&Z)")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ModiyPati, "����������Ϣ(&X)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_ChangePassWord, "�������(&P)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Family, "�����Ǽ�(&D)"): mcbrControl.BeginGroup = True
        
        '95809:���ϴ�,2016/8/26,�˲�����ʷ������
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Medical, "�˲�����(&N)"): mcbrControl.BeginGroup = True
        '104726:���ϴ�,2017/4/17,�շѷ�Ʊ�ش򲹴�
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Wham, "�ش�Ʊ��(&W)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Make, "����Ʊ��(&M)")
    End With

    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    mcbrMenuBar.id = conMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_View_Family, "������Ϣ(&V)"): mcbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBar = cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.id = conMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set mcbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
    End With
    
    '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_CardPay
        .Add FCONTROL, Asc("T"), conMenu_Edit_CardBack
        
        .Add FCONTROL, Asc("J"), conMenu_Edit_CardInFull
        .Add FCONTROL, Asc("B"), conMenu_Edit_CardBackMoney
        
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        .Add 0, VK_F11, conMenu_Edit_RollingCurtain
    End With
    '���ò����ò˵�
    With cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
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
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardPay, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBound, "�󶨿�"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CancelCardBound, "ȡ����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardBack, "�˿�")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_Cardtrade, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardFill, "����")
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_CardInFull, "��Ԥ��"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Edit_RollingCurtain, "�շ�����(&M)")
        mcbrControl.IconId = 227
        
        Set mcbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
        Set objComBar = .Add(xtpControlComboBox, conMenu_COMBOX_INTERFACE, "ҽ�ƿ����")
        objComBar.Flags = xtpFlagRightAlign
        objComBar.HideFlags = xtpNoHide
        objComBar.Width = (TextWidth("��") * 16) / Screen.TwipsPerPixelX
         objComBar.Style = xtpComboLabel
    End With
    For Each mcbrControl In mcbrToolBar.Controls
          If mcbrControl.id <> conMenu_COMBOX_INTERFACE Then
            mcbrControl.Style = xtpButtonIconAndCaption
          End If
    Next
    '��������
     Call LoadTypeData(objComBar)
     zlDefCommandBars = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub LoadTypeData(ByVal cbrCmb As CommandBarComboBox)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������б�����
    '����:���˺�
    '����:2011-06-29 16:51:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, intIndex As Integer, strSQL As String
    
    On Error GoTo errHandle
    
    intIndex = 1
    '�����:56599
    strSQL = "Select ID,����,����,�Ƿ�����,�Ƿ񷢿�,Nvl(�Ƿ��ظ�ʹ��,0) as �Ƿ��ظ�ʹ�� From ҽ�ƿ���� where �Ƿ�����=1 And Nvl(�Ƿ�֤��,0)=0 Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Set mcolCard = New Collection
    With rsTemp
        cbrCmb.Clear
        Do While Not .EOF
            cbrCmb.AddItem CStr(Nvl(!����)) & "-" & CStr(Nvl(!����))
            cbrCmb.ItemData(intIndex) = Val(Nvl(!id))
            mcolCard.Add Array(Val(Nvl(rsTemp!id)), Val(rsTemp!�Ƿ�����) & "-" & Val(rsTemp!�Ƿ񷢿�), Val(Nvl(rsTemp!�Ƿ��ظ�ʹ��))), "K" & rsTemp!id
            If mlngCardTypeID = Val(Nvl(!id)) Then
               cbrCmb.ListIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
    End With
    If intIndex > 1 And cbrCmb.ListIndex <= 0 Then
        cbrCmb.ListIndex = 1:
    End If
    If cbrCmb.ListIndex > 0 Then
        mlngCardTypeID = cbrCmb.ItemData(cbrCmb.ListIndex)
        mbln���� = Split(mcolCard(cbrCmb.ListIndex)(1), "-")(1) = 1 '�����:56599
        mbln���ƿ� = Split(mcolCard(cbrCmb.ListIndex)(1), "-")(0) = 1 '�����:56599
        mbln�ظ�ʹ�� = mcolCard(cbrCmb.ListIndex)(2) = 1 '57899
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngID As Long, strTemp As String, strCardNo As String
    Dim ctrCombox As CommandBarComboBox, lng����ID As Long
    Dim str�������� As String 'ȡ����,�˿�,����;�����:56599
    Dim objfrmPrint As frmPrint
    Dim strSelect As String
    '---------------------------------------------
    Set objfrmPrint = New frmPrint
    Load objfrmPrint
    Select Case Control.id
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_PrintSingleBill       '"�ش�ɿ(&R)")
        strSelect = zlCommFun.ShowMsgbox("�ɿ��ӡ", "��ѡ����Ҫ��ӡ�Ľɿ", "����(&F),Ԥ��(&I),ȡ��(&C)", Me, _
                                         vbDefaultButton2)
        If Not (strSelect = "ȡ��" Or strSelect = "") Then
            Call objfrmPrint.PrintReBill(strSelect, Trim(vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("����"))), _
                                         mlngCardTypeID, mPrint.bytPrintPayCard)
        End If
    Case conMenu_Edit_RollingCurtain   '�շ�Ա����
        Call zlExecuteChargeRollingCurtain(Me)
    Case conMenu_Edit_CardPay    '����(&S)")
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_����, mlngCardTypeID) = False Then Exit Sub
            If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call LoadDataToGrid
      Case conMenu_Edit_CardBound
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_�󶨿�, mlngCardTypeID) = False Then Exit Sub
            If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call LoadDataToGrid
      Case conMenu_Edit_CardBack    '�˿�(&B)")
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
           ' If strCardNo = "" Then Exit Sub
        End With
        '�����:56599
        str�������� = Check�˿�(strCardNo)
        Select Case str��������
            Case "ȡ����"
                zlExecuteCommandBars cbsThis.FindControl("", conMenu_Edit_CancelCardBound)
            Case "�˿�"
                If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_�˿�, mlngCardTypeID, strCardNo) = False Then Exit Sub
                If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                Call LoadDataToGrid
            Case "����"
            
        End Select
    Case conMenu_Edit_Cardtrade   '����
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
            '90233:���ϴ�,2015/11/5,�����Ͳ������벡��ID
            lng����ID = Val(Trim(.TextMatrix(.Row, .ColIndex("����ID"))))
            If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_����, mlngCardTypeID, strCardNo, lng����ID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardFill    '����(&B)")
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
            lng����ID = Val(Trim(.TextMatrix(.Row, .ColIndex("����ID"))))
          '  If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_����, mlngCardTypeID, strCardNo, lng����ID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardLoss        '��ʧ
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
            lng����ID = Val(Trim(.TextMatrix(.Row, .ColIndex("����ID"))))
        
            If strCardNo = "" Then Exit Sub
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_��ʧ, mlngCardTypeID, strCardNo, lng����ID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_CardCancelLoss        'ȡ����ʧ
        Call SaveCardCancelLose
   Case conMenu_Edit_CancelCardBound
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then Exit Sub
        End With
        '�����:56599
        str�������� = Checkȡ��Ժ�⿨��(lng����ID, strCardNo)
        Select Case str��������
            Case "ȡ����"
                frmPaticurCardCancelBound.mstrPrepayPrivs = mstrPrepayPrivs
                If frmPaticurCardCancelBound.zlCancelBand(Me, mlngModule, mlngCardTypeID, lng����ID, strCardNo, False) = False Then Exit Sub
                Call LoadDataToGrid
            Case "�˿�"
                zlExecuteCommandBars cbsThis.FindControl("", conMenu_Edit_CardBack)
                Exit Sub
            Case "����"
                
        End Select
    Case conMenu_Edit_CardInFull    '��Ԥ��(&J)"
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With
        'intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
        Call zlPrepayFunc(1, lng����ID)
    
    Case conMenu_Edit_CardBackMoney '�˿�
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With
        'intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
     Call zlPrepayFunc(2, lng����ID)
    Case conMenu_Edit_CardInFullBack    '��Ԥ��
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With        'intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
        Call zlPrepayFunc(3, lng����ID)
    Case conMenu_Edit_MzToZy
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With        'intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
        Call zlPrepayFunc(4, lng����ID)
    Case conMenu_Edit_ZyToMz
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With        'intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
        Call zlPrepayFunc(5, lng����ID)
    Case conMenu_Edit_ModiyPati  '����������Ϣ(&M)
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        End With
        If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_����������Ϣ, mlngCardTypeID, "", lng����ID) = False Then Exit Sub
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_ChangePassWord    '�������
         If frmModiPatiPass.zlModifyPass(Me, mlngModule, mlngCardTypeID, , , InStr(1, mstrPrivs, ";ǿ���޸�����;") = 0) Then
              Exit Sub
         End If
        If MsgBox("�����Ѿ������ı�,�Ƿ�����ˢ������?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Call LoadDataToGrid
    Case conMenu_Edit_Family
        '����:���˼�������
        If Not CreatePublicPatient Then Exit Sub
        Call mobjPubPatient.MakePatiFamily(Me, 0, 2, mlngModule) '�༭
        zlRefrshListData
    Case conMenu_Edit_Wham
        mPrint.intPrintMode = 3
        Call PrintBill
    Case conMenu_Edit_Make
        mPrint.intPrintMode = 4
        Call PrintBill
    Case conMenu_File_BarcodePrint
        mPrint.intPrintMode = 0
        Call PrintBill
    Case conMenu_View_Family
        If Not CreatePublicPatient Then Exit Sub
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        End With
        If lng����ID <= 0 Then Exit Sub
        Call mobjPubPatient.MakePatiFamily(Me, lng����ID, 1, mlngModule) '�鿴
    Case conMenu_Edit_Medical
        '���ܣ��˲����� ����в���ID�Ϳ������Զ���λ������ʱ�Ĳ����Ѽ�¼
        With vsCardList
            strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
            lng����ID = Val(Trim(.TextMatrix(.Row, .ColIndex("����ID"))))
        End With
        Call frmPatiBooks.ShowMe(Me, 1, mlngModule, lng����ID, strCardNo)
    Case conMenu_COMBOX_INTERFACE   '���ѡ��
        Set ctrCombox = Control
        mlngCardTypeID = ctrCombox.ItemData(ctrCombox.ListIndex)
        mbln���ƿ� = Split(mcolCard(ctrCombox.ListIndex)(1), "-")(0) = 1
        mbln���� = Split(mcolCard(ctrCombox.ListIndex)(1), "-")(1) = 1 '�����:56599
        '115505:���ϴ�,2017/10/23,���¿�����
        mbln�ظ�ʹ�� = mcolCard(ctrCombox.ListIndex)(2) = 1
        Call LoadDataToGrid
    Case conMenu_View_Refresh   'ˢ��
        '����ˢ������
        Call LoadDataToGrid
    Case mlngListReportID    '����50122
        If vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("����ID")) = "" Then
            '�����:57285
            ShowMsgbox "����û��ѡ���Ӧ����֧��ϸ,�뵽�ʻ������Ϣѡ���Ӧ����֧��ϸ!"
            Exit Sub
        End If
        mfrmConsume.zlShowReport CLng(vsCardList.TextMatrix(vsCardList.Row, vsCardList.ColIndex("����ID")))
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Call zl_OpenReport(Val(Split(Control.Parameter, ",")(0)), Split(Control.Parameter, ",")(1))
        End If
    End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
 Private Function zlPopuReportMenus() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������֧���˵�
    '����:���˺�
    '����:2012-06-12 15:28:03
    '����:50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    If mlngListReportID = 0 Then Exit Function
    Set cbrPopupBar = Me.cbsThis.Add("��������˵�", xtpBarPopup)
    Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mlngListReportID, mstrListReportName)
    cbrPopupBar.ShowPopup
 End Function
 
 
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng�����ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String

    If IsCardType(IDKind, "IC����") And Not txtPatient.Locked Then
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call txtPatient_KeyPress(vbKeyReturn)
            End If
        End If
        Exit Sub
    End If

    lng�����ID = IDKind.GetCurCard.�ӿ����
    If lng�����ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '����:�����ӿ�
    '    '���:frmMain-���õĸ�����
    '    '       lngModule-���õ�ģ���
    '    '       strExpand-��չ����,������
    '    '       blnOlnyCardNO-������ȡ����
    '    '����:strOutCardNO-���صĿ���
    '    '       strOutPatiInforXML-(������Ϣ����.XML��)
    '    '����:��������    True:���óɹ�,False:����ʧ��\
   
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng�����ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    Set gobjSquare.objCurCard = objCard
    '105155:���ϴ�,2017/2/8,����������ʾ�жϲ���ȷ
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    If objCard.�ӿ���� > 0 Then
        txtPatient.MaxLength = objCard.���ų���
    Else
        txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    End If
    
    '��Ҫ�����Ϣ,����ˢ����,���л�,���������ʾʧȥ����
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = "": mlngCurPatient = 0
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
    If mlngCurPatient <> 0 Then
        txtPatient.PasswordChar = ""  '����Ѿ�ͨ�����ַ�ʽ��ȡ���˲���,��ʱ��ʾ���ǲ�������,��Ӧ����������
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
End Sub

 
Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean
    '60010
    If txtPatient.Locked Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.����)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.����
    Call txtPatient_KeyPress(vbKeyReturn)
'    If mrsInfo Is Nothing Then
'        blnNew = True
'    ElseIf mrsInfo.State <> 1 Then
'        blnNew = True
'    End If
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

 '����50122
Private Sub mfrmConsume_zlPopupMenus(ByVal vsGrid As VSFlex8Ctl.VSFlexGrid)
    If vsGrid.Rows = 1 Or vsGrid.Row = 1 Then Exit Sub
    zlPopuReportMenus
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean, lngID As Long, blnEnabled As Boolean, bln��ʧ As Boolean
    Dim lng����ID As Long
    Dim blnIsBind As Boolean '���ƿ��Ƿ�ͨ���󶨵ķ�ʽ�����ظ�ʹ�õ�
    If Me.Visible = False Then Exit Sub
    
    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = zlIsHaveData
    Case conMenu_Edit_RollingCurtain        '�շ�Ա����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs_RollingCurtain, "����")
        Control.Enabled = Control.Visible
        
    Case conMenu_File_PrintSingleBill           '"�ش�ɿ(&R)"
        blnIsBind = True
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�"))) <> "" And .TextMatrix(.Row, .ColIndex("״̬")) = "��Ч��" And lng����ID > 0
        End With
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "Ԥ���վ�") Or zlstr.IsHavePrivs(mstrPrivs, "ҽ�ƿ��վ�")) And blnIsBind And Not gbln�շѷ�Ʊ
        Control.Enabled = Control.Visible
    Case conMenu_File_BarcodePrint
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "�ش򷢿�ƾ��")) And lng����ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardPay
        Control.Visible = (zlstr.IsHavePrivs(mstrPrivs, "����") And mbln���ƿ�) Or (zlstr.IsHavePrivs(mstrPrivs, "����") And mbln���ƿ� = False And mbln���� = True) '�����:56599
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardBound
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�󶨿�") And ((Not mbln���ƿ�) Or (mbln���ƿ� And mbln�ظ�ʹ��))
        Control.Enabled = Control.Visible
        
    Case conMenu_Edit_CancelCardBound
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With
        blnIsBind = True
        If mbln���ƿ� Then
            '���ƿ� ʹ�ð󶨽����ظ�ʹ��
            blnIsBind = zlIsCardBinding
        End If
        '�������ƿ����ð�,���ƿ� �ظ�����ʹ�ð󶨵ķ�ʽ
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "ȡ���󶨿�") And ((Not mbln���ƿ�) Or (mbln���ƿ� And blnIsBind))
        Control.Enabled = Control.Visible And lng����ID > 0
    Case conMenu_Edit_CardBack
        blnIsBind = False
        If mbln���ƿ� Then
            '���ƿ� ʹ�ð󶨽����ظ�ʹ��
            blnIsBind = zlIsCardBinding
        End If
        '԰�ڿ��˿�,�������˿�,ֻ��ȡ��
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�˿�") And mbln���ƿ� And Not blnIsBind
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Cardtrade '����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����") And mbln���ƿ�
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardFill  '����
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����") And mbln���ƿ�
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardLoss  '��ʧ
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "��ʧ") And mstrCurStatus = "��Ч��"
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardCancelLoss  'ȡ����ʧ
        bln��ʧ = mstrCurStatus = "�ѹ�ʧ"
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "ȡ����ʧ") ' And mbln���ƿ�
        Control.Enabled = Control.Visible And zlIsHaveData And bln��ʧ
    Case conMenu_Edit_CardInFull  '��Ԥ��
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "����Ԥ��") Or zlstr.IsHavePrivs(mstrPrepayPrivs, "סԺԤ��") Or zlstr.IsHavePrivs(mstrPrepayPrivs, "����Ԥ��")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardInFullBack  '��Ԥ��
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "Ԥ���˿�")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_CardBackMoney  '�˿�
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "Ԥ���˿�")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_MzToZy    '����תסԺ
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "����Ԥ��תסԺ")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_ZyToMz    'סԺת����
        Control.Visible = zlstr.IsHavePrivs(mstrPrepayPrivs, "סԺԤ��ת����")
        Control.Enabled = Control.Visible
   Case conMenu_Edit_ModiyPati
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����������Ϣ")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_ChangePassWord    '�������
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�޸�����")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Family
        Control.Visible = zlstr.IsHavePrivs(mstrPubPatiPrivs, "���˼���")
        Control.Enabled = Control.Visible
    Case conMenu_View_Family
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPubPatiPrivs, "���˼���") And lng����ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Wham
        blnIsBind = True
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�"))) <> "" And .TextMatrix(.Row, .ColIndex("״̬")) = "��Ч��"
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "�ش�Ʊ") And gbln�շѷ�Ʊ And blnIsBind And lng����ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Make
        blnIsBind = True
        With vsCardList
            lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
            If lng����ID <= 0 Then lng����ID = 0
            blnIsBind = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�"))) <> "" And .TextMatrix(.Row, .ColIndex("״̬")) = "��Ч��"
        End With
        Control.Visible = zlstr.IsHavePrivs(mstrPrivs, "����Ʊ") And gbln�շѷ�Ʊ And blnIsBind And lng����ID > 0
        Control.Enabled = Control.Visible
    Case conMenu_View_Refresh   'ˢ��
    Case Else
        If (Control.id >= conMenu_ReportPopup * 100# + 1 And Control.id <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            Control.Visible = Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1503_1" And Split(Control.Parameter, ",")(1) <> "ZL" & glngSys \ 100 & "_INSIDE_1107_2"
        End If
    End Select
End Sub
 
'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    '------------------------------------
    Select Case Control.id
        Case conMenu_File_Exit: Unload Me
        Case conMenu_File_PrintSet: Call zlPrintSet
        Case conMenu_View_StatusBar
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
        Case conMenu_File_Parameter     '��������
             If frmPatiCureCardPara.zlSetPara(Me, mlngModule, mstrPrivs) = False Then Exit Sub
        Case Else   '�����������ܵ���
            Call zlExecuteCommandBars(Control)
        End Select
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub
Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub
Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnHaveData As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Me.Visible = False Then Exit Sub

    If Control.type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    Err = 0: On Error Resume Next
    Select Case Control.id
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = stbThis.Visible
    Case Else
        Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    zlControl.ControlSetFocus vsCardList
    Call vsCardList_GotFocus
    mblnFirst = False
End Sub

Private Sub Form_Initialize()
  Call InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Dim strShow As String
    Dim i As Long
    If mblnInited = False Then
        mblnInited = True
    Else
        Exit Sub
    End If
    mblnFirst = True: mstrCurStatus = ""
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    mstrPrepayPrivs = ";" & GetPrivFunc(glngSys, 1103) & ";"
    mstrPubPatiPrivs = ";" & GetPrivFunc(glngSys, 9003) & ";"
    Call InitFace
    mlngCardTypeID = Val(zlDatabase.GetPara("�ϴ�ҽ�����", glngSys, mlngModule, 0, , InStr(1, mstrPrivs, ";��������;") > 0))
    'ֻ�з����Ż����ش򲹴�
    mPrint.bytPrintPayCard = Split(zlDatabase.GetPara("ҽ�ƿ��վݸ�ʽ", glngSys, mlngModule, "0|0"), "|")(0)
    mPrint.bytPrintBoundCard = Split(zlDatabase.GetPara("ҽ�ƿ��վݸ�ʽ", glngSys, mlngModule, "0|0"), "|")(1)
    RestoreWinState Me, App.ProductName, mstrTitle
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    Call InitData: Call InitPanel: Call InitPage
    Call zlDefCommandBars '��ʼ�˵���������
    Call InitFilterToVar
    Call LoadDataToGrid(-1)
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
    Call InitListReport
End Sub
Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-06-21 13:19:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strKindStr As String, blnVisible As Boolean
    Dim intKind As Integer, strKey As String
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    
    Call InitIDKind
     
    
    'ȡȱʡ��ˢ����ʽ
    '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|
    '�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������)
    '��7λ��,��ֻ��������,��Ȼȡ������
    
    '89086:���ϴ�,2015/10/9,�����������ģ������
    mblnSeekName = zlDatabase.GetPara("����ģ������", glngSys, mlngModule) = "1"
    mintNameDays = Val(zlDatabase.GetPara("������������", glngSys, mlngModule))
    
    gobjSquare.blnȱʡ�������� = IDKind.ShowPassText
    
    Call GetRegInFor(g˽��ģ��, Me.Name, "idkind", strKey)
    intKind = Val(strKey)
    If intKind > 0 And intKind <= IDKind.ListCount Then
        IDKind.IDKind = intKind
    End If
    
 End Sub
Private Function InitIDKind() As Boolean
    Dim lngCardID As Long
    If gobjSquare Is Nothing Then Exit Function
    gobjSquare.objSquareCard.mblnYLMgr = True
    lngCardID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModule, 0))
    On Error GoTo ErrEnd
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, IDKind.IDKindStr, txtPatient)
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    IDKind.ShowPropertySet = InStr(";" & mstrPrivs & ";", "��������") > 0
ErrEnd:
End Function
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:���˺�
    '����:2011-06-29 18:08:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtp��������(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtp��������(0).value = Format(dtp��������(0).MaxDate, "yyyy-mm-dd 23:59:59")
    dtp��ʼ����(0).MaxDate = dtp��������(0).MaxDate
    dtp��ʼ����(0).value = Format(DateAdd("d", -7, dtp��ʼ����(0).MaxDate), "yyyy-mm-dd 00:00:00")
    dtp��������(1).MaxDate = dtp��������(0).MaxDate
    dtp��������(1).value = Format(dtp��������(1).MaxDate, "yyyy-mm-dd 23:59:59")
    dtp��ʼ����(1).MaxDate = dtp��������(1).MaxDate
    dtp��ʼ����(1).value = dtp��ʼ����(0).value
 
     
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Dim i As Long, strTemp As String
   If Me.Visible = False Then Exit Sub
   SaveWinState Me, App.ProductName, mstrTitle
   zlDatabase.SetPara "�ϴ�ҽ�����", mlngCardTypeID, glngSys, mlngModule, InStr(1, mstrPrivs, ";��������;") > 0
   zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
   
   
   zlSaveDockPanceToReg Me, dkpMan, "����"
   If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled False
        Set mobjIDCard = Nothing
   End If
   If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled False
        Set mobjICCard = Nothing
   End If
   If Not mobjReport Is Nothing Then Set mobjReport = Nothing
   If Not mobjPubPatient Is Nothing Then Set mobjPubPatient = Nothing
    '�ر��Ӵ���
    If Not mobjSubFrm Is Nothing Then
        For i = 1 To mobjSubFrm.count
            If Not mobjSubFrm(i) Is Nothing Then Unload mobjSubFrm(i)
        Next
    End If
    
    Call SaveRegInFor(g˽��ģ��, Me.Name, "idkind", IDKind.IDKind)
End Sub
 Private Function zlPopuMenus(ByVal blnListView As Boolean) As Boolean
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Err = 0: On Error Resume Next
    If Me.cbsThis.ActiveMenuBar.Controls(2).Visible = False Then Exit Function
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = Me.cbsThis.Add("�����˵�", xtpBarPopup)
    For Each mcbrControl In mcbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
        cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
    Next

    If Me.cbsThis.ActiveMenuBar.Controls(3).Visible Then
        Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls(3)
        For Each mcbrControl In mcbrMenuBar.CommandBar.Controls

            Select Case mcbrControl.id
            Case conMenu_View_ShowStoped, conMenu_View_ShowAll, conMenu_View_Refresh
                Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, mcbrControl.id, mcbrControl.Caption)
                cbrPopupItem.BeginGroup = mcbrControl.BeginGroup
                cbrPopupItem.Checked = mcbrControl.Checked
            End Select
        Next
    End If
    cbrPopupBar.ShowPopup
End Function

Private Function zlCheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '����:���ݺϷ�,����true�����򷵻�False
    '����:���˺�
    '����:2009-11-19 15:37:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, strSQL As String
    zlCheckDepend = False
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ����   From ���㷽ʽ Where ���� = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ֽ���㷽ʽ", UserInfo.id)
    If rsTemp.EOF Then
        ShowMsgbox "���㷽ʽ�в�����һ�������ֽ����ʵĽ��㷽ʽ,���ڽ��㷽ʽ����������!"
        rsTemp.Close
        Exit Function
    End If
    rsTemp.Close
    '76009,Ƚ����,2014-7-30
    strSQL = "Select 1 From ҽ�ƿ���� Where Nvl(�Ƿ�����, 0) = 1 And Rownum <= 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ҽ�ƿ����")
    If rsTemp.RecordCount = 0 Then
        ShowMsgbox "ҽ�ƿ�����в������κο���������ڡ�ҽ�ƿ��������н���ά����"
        Exit Function
    End If
    zlCheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub ShowList(ByVal lngModule As Long, ByVal strTitle As String, ByVal frmMain As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������,��ʾ��ص���Ŀ��������Ϣ
    '����:���˺�
    '����:2009-11-19 15:38:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrTitle = strTitle: mstrPrivs = gstrPrivs
    If Not zlCheckDepend Then Exit Sub            '���������Բ���
    Me.Caption = strTitle
    If IsObject(frmMain) Then
        Me.Show , frmMain
    Else
        zlCommFun.ShowChildWindow Me.hWnd, frmMain
    End If
    Me.ZOrder 0
End Sub

 
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long, vRect As RECT
    vRect = zlControl.GetControlRect(picImgList.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsCardList, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub InitFilterToVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������������
    '����:���˺�
    '����:2011-06-28 23:56:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mArrFilter = New Collection
    mArrFilter.Add Array(Trim(txtCard(0).Text), Trim(txtCard(1).Text)), "���ŷ�Χ"
    If chkFilter(0).value Then
        mArrFilter.Add Array(Format(dtp��ʼ����(0).value, "yyyy-mm-dd HH:MM:SS"), Format(dtp��������(0).value, "yyyy-mm-dd HH:MM:SS")), "����ʱ��"
    Else
        mArrFilter.Add Array("1901-01-01", "1901-01-01"), "����ʱ��"
    End If
    If chkFilter(1).value Then
        mArrFilter.Add Array(Format(dtp��ʼ����(1).value, "yyyy-mm-dd HH:MM:SS"), Format(dtp��������(1).value, "yyyy-mm-dd HH:MM:SS")), "��ʧʱ��"
    Else
        mArrFilter.Add Array("1901-01-01", "1901-01-01"), "��ʧʱ��"
    End If
    mArrFilter.Add Trim(txtEdit(mtxtIdx.idx_������)), "������"
    mArrFilter.Add Trim(txtEdit(mtxtIdx.idx_��ʧ��)), "��ʧ��"
    mArrFilter.Add Trim(txtName.Text), "����"
End Sub
Private Function LoadDataToGrid(Optional lng����ID As Long = 0, Optional strCardNo As String = "") As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
'����:�������ݸ�����
'���:lng����ID-��ָ���Ĳ���ID;
'       strCardNo-��ָ���Ŀ���
'����:���سɹ�,����true,���򷵻�False
'����:���˺�
'����:2009-11-19 15:43:29
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strWhere As String, lngRow As Long, strPreCardNO As String, lngPreTypeID As Long
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long
    Err = 0: On Error GoTo Errhand:
    strWhere = ""
    If lng����ID <> 0 Then
        strWhere = strWhere & " And A.����ID=[10] "
    ElseIf strCardNo <> "" Then
        strWhere = strWhere & " And A.����=[11] "
    Else
        If mArrFilter("����ʱ��")(0) <> "1901-01-01" And mArrFilter("��ʧʱ��")(0) <> "1901-01-01" Then
            strWhere = strWhere & " And (A.�������� Between [1] And [2] Or A.��ʧʱ�� Between [3] And [4])"
        ElseIf mArrFilter("����ʱ��")(0) = "1901-01-01" And mArrFilter("��ʧʱ��")(0) <> "1901-01-01" Then
            strWhere = strWhere & " And (A.��ʧʱ�� Between [3] And [4])"
        ElseIf mArrFilter("����ʱ��")(0) <> "1901-01-01" And mArrFilter("��ʧʱ��")(0) = "1901-01-01" Then
            strWhere = strWhere & " And (A.�������� Between [1] And [2])"
        End If
        If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
            strWhere = strWhere & " And (A.���� Between [5] And [6])"
        ElseIf mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
            strWhere = strWhere & " And A.����=[6]"
        ElseIf mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) = "" Then
            strWhere = strWhere & " And A.����=[5]"
        End If
        If mArrFilter("��ʧ��") <> "" Then strWhere = strWhere & " and  A.��ʧ�� like [7]"
        If mArrFilter("������") <> "" Then strWhere = strWhere & " and  A.������ like [8]"
        If mArrFilter("����") <> "" Then
            If zlstr.ActualLen(mArrFilter("����")) < 4 And _
                   (DateDiff("d", CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1))) > 7 Or _
                   DateDiff("d", CDate(mArrFilter("��ʧʱ��")(0)), CDate(mArrFilter("��ʧʱ��")(1))) > 7) Then
                    MsgBox "�������Ϣ̫�٣����������������������������ֻ��ĸ��ַ�!", vbInformation, gstrSysName
                    Exit Function
            Else
                strWhere = strWhere & " and  B.���� like [12]"
            End If
        End If
        If mArrFilter("��ʧ��") = "" And mArrFilter("������") = "" And mArrFilter("����") = "" And _
           mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) = "" Then
            If DateDiff("d", CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1))) > 30 Or _
               DateDiff("d", CDate(mArrFilter("��ʧʱ��")(0)), CDate(mArrFilter("��ʧʱ��")(1))) > 30 Then
                If MsgBox("ѡ���ʱ�䷶Χ������30��,�Ƿ����?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Function
            End If
        End If
    End If
    Call zlCommFun.ShowFlash("���ڼ��ز���ҽ�ƿ���Ϣ,���Ե�...", Me)
    '    strSQL = "" & _
         '         "    Select  A.����ID,A.�����ID,C.����||'-'|| C.���� As ҽ�ƿ����,A.����,  " & _
         '         "             case when Nvl(A.״̬,0)=0 then '��Ч��' " & _
         '         "                     when Nvl(A.״̬,0)=2 then '����ͣ��' " & _
         '         "                     When Nvl(��ʧʱ��,to_date('3000-01-01','yyyy-mm-dd'))+ Nvl(D.��Ч����,0)<=sysdate then '�ѹ�ʧ' " & _
         '         "                     Else ''  end as ״̬, " & _
         '         "              A.��ʧ��,A.��ʧ��ʽ, to_char(A.��ʧʱ��,'yyyy-mm-dd hh24:mi:ss') as ��ʧʱ��," & _
         '         "             A.������,to_char(A.��������,'yyyy-mm-dd hh24:mi:ss')  as ��������, " & _
         '         "             B.����,B.�Ա�,B.����,to_char(B.��������,'yyyy-mm-dd hh24:mi:ss')  as  ��������,B.�����ص�," & _
         '         "             B.���֤��,B.�����,B.סԺ��,b.����,b.��ͥ��ַ,B.��ͥ�绰,b.�໤��,b.��ϵ������, " & _
         '         "             b.��ϵ�˹�ϵ,b.��ϵ�˵�ַ,b.��ϵ�˵绰,b.������λ,b.��λ�绰,b.��ͥ��ַ�ʱ�,decode(b.��Ժ,1,'��','') as ��Ժ " & _
         '         "     From ����ҽ�ƿ���Ϣ A,������Ϣ B,ҽ�ƿ���� C, ҽ�ƿ���ʧ��ʽ D " & _
         '         "     Where A.����ID=B.����ID And A.�����ID=C.Id And A.��ʧ��ʽ=D.����(+)  and A.�����ID=[9] " & strWhere

    strSQL = " Select * " & vbNewLine & _
           "   From (With ҽ�ƿ��䶯 As (" & vbNewLine & _
           "                      Select Max(Bd.ID) As �䶯id, C.�Ƿ��ϸ����, C.�Ƿ��ظ�ʹ��, A.����id, A.�����id," & vbNewLine & _
           "                              C.���� || '-' || C.���� As ҽ�ƿ����, A.����," & vbNewLine & _
           "                              Case" & vbNewLine & _
           "                                When Nvl(A.״̬, 0) = 0 Then" & vbNewLine & _
           "                                 '��Ч��'" & vbNewLine & _
           "                                When Nvl(A.״̬, 0) = 2 Then" & vbNewLine & _
           "                                 '����ͣ��'" & vbNewLine & _
    "                                When Nvl(��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(D.��Ч����, 0) <= Sysdate Then"
    strSQL = strSQL & vbNewLine & _
           "                                 '�ѹ�ʧ'" & vbNewLine & _
           "                                Else" & vbNewLine & _
           "                                 ''" & vbNewLine & _
           "                              End As ״̬, A.��ʧ��, A.��ʧ��ʽ," & vbNewLine & _
           "                              To_Char(A.��ʧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��ʧʱ��, A.������," & vbNewLine & _
           "                              To_Char(A.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, B.����, B.�Ա�, B.����," & vbNewLine & _
           "                              To_Char(B.��������, 'yyyy-mm-dd hh24:mi:ss') As ��������, B.�����ص�, B.���֤��, B.�����," & vbNewLine & _
           "                              B.סԺ��,B.�ֻ���, B.����, B.��ͥ��ַ as ��סַ, B.��ͥ�绰, B.�໤��, B.��ϵ������, B.��ϵ�˹�ϵ," & vbNewLine & _
           "                              B.��ϵ�˵�ַ, B.��ϵ�˵绰, B.������λ, B.��λ�绰, B.��ͥ��ַ�ʱ�," & vbNewLine & _
           "                              Decode(B.��Ժ, 1, '��', '') As ��Ժ" & vbNewLine & _
           "                      From ����ҽ�ƿ���Ϣ A, ������Ϣ B, ҽ�ƿ���� C, ҽ�ƿ���ʧ��ʽ D, ����ҽ�ƿ��䶯 Bd" & vbNewLine & _
           "                      Where A.����id = B.����id And A.�����id = C.ID And A.��ʧ��ʽ = D.����(+) And" & vbNewLine & _
           "                            A.�����id=[9]   " & strWhere & " And A.����id = Bd.����id(+) And A.�����id = Bd.�����id(+) And" & vbNewLine & _
           "                            A.���� = Bd.����(+)" & vbNewLine & _
           "                      Group By C.�Ƿ��ϸ����, C.�Ƿ��ظ�ʹ��, A.����id, A.�����id, C.���� || '-' || C.����, A.����," & vbNewLine & _
           "                                Case" & vbNewLine & _
           "                                  When Nvl(A.״̬, 0) = 0 Then" & vbNewLine & _
           "                                   '��Ч��'" & vbNewLine & _
           "                                  When Nvl(A.״̬, 0) = 2 Then" & vbNewLine & _
           "                                   '����ͣ��'" & vbNewLine & _
           "                                  When Nvl(��ʧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) + Nvl(D.��Ч����, 0) <=Sysdate Then" & vbNewLine & _
           "                                   '�ѹ�ʧ'" & vbNewLine & _
           "                                  Else " & vbNewLine & _
    "                                   '' "
    strSQL = strSQL & vbNewLine & _
           "                                End, A.��ʧ��, A.��ʧ��ʽ, To_Char(A.��ʧʱ��, 'yyyy-mm-dd hh24:mi:ss'), A.������," & vbNewLine & _
           "                                To_Char(A.��������, 'yyyy-mm-dd hh24:mi:ss'), B.����, B.�Ա�, B.����," & vbNewLine & _
           "                                To_Char(B.��������, 'yyyy-mm-dd hh24:mi:ss'), B.�����ص�, B.���֤��, B.�����, B.סԺ��,B.�ֻ���," & vbNewLine & _
           "                                B.����, B.��ͥ��ַ, B.��ͥ�绰, B.�໤��, B.��ϵ������, B.��ϵ�˹�ϵ, B.��ϵ�˵�ַ," & vbNewLine & _
           "                                B.��ϵ�˵绰, B.������λ, B.��λ�绰, B.��ͥ��ַ�ʱ�, Decode(B.��Ժ, 1, '��', '')" & vbNewLine & _
           "                      )" & vbNewLine & _
           "   Select T.����id, T.�����id, T.ҽ�ƿ����, T.����, T.״̬, T.��ʧ��, T.��ʧ��ʽ, T.��ʧʱ��, T.������, T.��������," & vbNewLine & _
           "          T.����, T.�Ա�, T.����, T.��������, T.�����ص�, T.���֤��, T.�����, T.סԺ��,T.�ֻ���, T.����, T.��סַ," & vbNewLine & _
           "           T.��ͥ�绰, T.�໤��, T.��ϵ������, T.��ϵ�˹�ϵ, T.��ϵ�˵�ַ, T.��ϵ�˵绰, T.������λ, T.��λ�绰," & vbNewLine & _
           "          T.��ͥ��ַ�ʱ�, T.��Ժ, Nvl(Bd.�䶯���, 0) As �䶯���, T.�Ƿ��ϸ����, T.�Ƿ��ظ�ʹ��, Z.NO as ���ݺ�," & vbNewLine & _
           "          LTrim(To_Char(max(Decode(Y.����, 1, Nvl(Y.Ԥ�����,0),0)),'99999999990.00')) As ����Ԥ�����, " & _
           "          LTrim(To_Char(max(Decode(Y.����, 1, 0, Nvl(Y.Ԥ�����,0))),'99999999990.00')) As סԺԤ����� " & _
           "   From ҽ�ƿ��䶯 T, ����ҽ�ƿ��䶯 Bd , סԺ���ü�¼ Z,������� Y" & vbNewLine & _
           "   Where T.�䶯id = Bd.ID(+) And T.����ID = Z.����ID(+) And T.����ID = Y.����ID(+) and Y.����(+)=1" & vbNewLine & _
           "         And Z.��¼����(+) = 5 And z.��¼״̬(+) = 1 And T.���� = Z.ʵ��Ʊ��(+) And T.�����ID = Nvl(Z.����(+),0)" & vbNewLine & _
           "   Group by T.����id, T.�����id, T.ҽ�ƿ����, T.����, T.״̬, T.��ʧ��, T.��ʧ��ʽ, T.��ʧʱ��, T.������, T.��������," & vbNewLine & _
           "         T.����, T.�Ա�, T.����, T.��������, T.�����ص�, T.���֤��, T.�����, T.סԺ��,T.�ֻ���, T.����, T.��סַ," & vbNewLine & _
           "         T.��ͥ�绰, T.�໤��, T.��ϵ������, T.��ϵ�˹�ϵ, T.��ϵ�˵�ַ, T.��ϵ�˵绰, T.������λ, T.��λ�绰," & vbNewLine & _
           "         T.��ͥ��ַ�ʱ�, T.��Ժ, Bd.�䶯���, T.�Ƿ��ϸ����, T.�Ƿ��ظ�ʹ��, Z.NO) T"

    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
                                          CDate(mArrFilter("����ʱ��")(0)), CDate(mArrFilter("����ʱ��")(1)), _
                                          CDate(mArrFilter("��ʧʱ��")(0)), CDate(mArrFilter("��ʧʱ��")(1)), _
                                          CStr(mArrFilter("���ŷ�Χ")(0)), CStr(mArrFilter("���ŷ�Χ")(1)), _
                                          CStr(mArrFilter("��ʧ��")), CStr(mArrFilter("������")), mlngCardTypeID, _
                                          lng����ID, strCardNo, Trim(txtName.Text) & "%")
    With vsCardList
        If .Row > 0 And .ColIndex("����") >= 0 Then
            strPreCardNO = Trim(.TextMatrix(.Row, .ColIndex("����")))
            If strPreCardNO <> "" And .ColIndex("�����ID") >= 0 Then
                lngPreTypeID = Val(.TextMatrix(.Row, .ColIndex("�����ID")))
            End If
        End If
        .Redraw = flexRDNone
        .Clear: .Rows = 2: .Cols = 1
        .Cell(flexcpForeColor, 1, .FixedCols - 1, .Rows - 1, .Cols - 1) = .ForeColor
        .Cell(flexcpData, 0, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpText, 0, 0, .Rows - 1, .Cols - 1) = ""
        mblnNotRefresh = True
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        .Row = 1
        If lngPreTypeID = mlngCardTypeID Then   '�ж�λ
            i = .FindRow(strPreCardNO, 1, .ColIndex("����"))
            If i >= 1 Then .Row = i
        End If
        mblnNotRefresh = False

        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = True
                .ColData(i) = "-1|1"    '����ѡ��
            ElseIf .ColKey(i) Like "*ʱ��" Or .ColKey(i) Like "*����" Or .ColKey(i) = "״̬" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf InStr(";�䶯���;�Ƿ��ϸ����;�Ƿ��ظ�ʹ��;���ݺ�;", ";" & .ColKey(i) & ";") > 0 Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True: .ColWidth(i) = 0
                .ColData(i) = "-1|1"    '����ѡ��
            ElseIf .ColKey(i) Like "*���" Then
                .ColAlignment(i) = flexAlignRightCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .ColData(.ColIndex("����")) = "1|0": .ColData(.ColIndex("��־")) = "-1|1"
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '�ָ�������
        zl_vsGrid_Para_Restore mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, True
        .ColWidth(.ColIndex("��־")) = 285
        .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter
        Call SetGridRowForeColor     '��������ɫ
        .Redraw = flexRDBuffered
    End With
    Call vsCardList_AfterRowColChange(-1, 0, vsCardList.Row, 0)
    Call zlCommFun.StopFlash
    LoadDataToGrid = True
    Exit Function
Errhand:
    Call zlCommFun.StopFlash
    vsCardList.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SetGridRowForeColor(Optional ByVal lngRow As Long = 0)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������ɫ
    '���:lngRow=0,��ʾ�������������е���ɫ
    '����:���˺�
    '����:2011-06-29 15:04:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long, int״̬ As Integer, lngRows As Long, i As Long
    With vsCardList
        If lngRow = 0 Then lngRows = .Rows - 1: lngRow = 1
        For i = lngRow To lngRows
            lngColor = IIf(Trim(.TextMatrix(i, .ColIndex("��ʧʱ��"))) <> "", vbRed, &H80000008)
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = lngColor
        Next
    End With
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
    Dim lng����ID As Long, datDate As Date
    Dim strSQL As String
    
    On Error GoTo errH
    If gblnBill���� Then
        lng����ID = GetInvoiceGroupID(1, TotalPages, mPrint.lng����ID, glngShareUseID, mPrint.strInvoice, mPrint.strUseType)
        If lng����ID <= 0 Then
            Select Case lng����ID
                Case -1
                    MsgBox "����[" & mPrint.strPrintNo & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "��û���㹻�����ú͹��õ�Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -2
                    MsgBox "����[" & mPrint.strPrintNo & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "��û���㹻�ĵĹ���Ʊ�ݣ�������һ�������ñ��ع���Ʊ�ݺ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -3
                    MsgBox "����[" & mPrint.strPrintNo & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "Ʊ�ݺ�[" & mPrint.strInvoice & "]���ڿ����������ε���ЧƱ�ݺŷ�Χ�ڣ�" & _
                        "������������Ч��Ʊ�ݺź��ش�õ��ݣ�", vbInformation, gstrSysName
                Case -4
                    MsgBox "����[" & mPrint.strPrintNo & "]" & "����Ҫ" & TotalPages & "��Ʊ�ݣ�" & vbCrLf & _
                        "Ʊ�ݺ�[" & mPrint.strInvoice & "]���ڵ���������û���㹻��Ʊ�ݣ�" & _
                        "���ȴ�ӡ����Ʊ��,���굱ǰ�������κ��ش�õ��ݣ�", vbInformation, gstrSysName
                Case Else
                    MsgBox "Ʊ��������Ϣ����ʧ�ܣ�������������ش򵥾�[" & mPrint.strInvoice & "]��", vbInformation, gstrSysName
            End Select
            Cancel = True: Exit Sub
        End If
    End If
    datDate = zlDatabase.Currentdate
    strSQL = "Zl_���˷���Ʊ��_Print("
    '  No_In           Varchar2,
    strSQL = strSQL & "'" & Replace(mPrint.strPrintNo, "'", "") & "'" & ","
    '  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "'" & mPrint.strInvoice & "',"
    '  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "" & ZVal(lng����ID) & ","
    '  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(datDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    '  ��������_In     Number
    strSQL = strSQL & mPrint.intPrintMode & ","
    '  Ʊ������_In     Number := 1,
    strSQL = strSQL & "" & TotalPages & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, "Ʊ����������")
    
    '���ϸ����Ʊ��ʱ���浽ע���
    '���±���Ʊ��
    If Not gblnBill���� Then
        zlDatabase.SetPara "��ǰ�շ�Ʊ�ݺ�", mPrint.strInvoice, glngSys, 1121
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Cancel = True
End Sub

Private Sub picCardList_Resize()
    Err = 0: On Error Resume Next
    With picCardList
        IDKind.Top = .ScaleTop + 100
        txtPatient.Top = IDKind.Top
        lblPati.Top = IDKind.Top + (IDKind.Height - lblPati.Height) \ 2
        vsCardList.Left = .ScaleLeft
        vsCardList.Width = .ScaleWidth
        vsCardList.Height = .ScaleHeight - vsCardList.Top
    End With
End Sub
Private Sub picFilter_Resize()
    Err = 0: On Error Resume Next
    With picFilter
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub

 

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnNotClick Then Exit Sub
    Call zlRefrshListData
End Sub

Private Sub txtCard_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mtxtIdx.idx_������ Or Index = mtxtIdx.idx_��ʧ�� Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
        zlCommFun.OpenIme False
End Sub

 

Private Sub txtName_Change()
    If zlstr.ActualLen(txtName.Text) < 4 Then
        If chkFilter(0).value = 0 And chkFilter(1).value = 0 Then chkFilter(0).value = 1
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

 Private Sub vsCardList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strCardNo  As String, lng����ID As Long
    If NewRow <= 0 Then Exit Sub
    If mblnNotRefresh Then Exit Sub
    zl_VsGridRowChange vsCardList, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow = NewRow Then Exit Sub
    Call zlRefrshListData
End Sub
Private Sub zlRefrshListData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����ϸ����
    '����:���˺�
    '����:2012-06-12 14:43:06
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, lng����ID As Long
    
    If tbPage.Selected Is Nothing Then Exit Sub
    On Error GoTo errHandle
    zlCommFun.ShowFlash "����װ������,���Ժ�..."
    With vsCardList
        If .ColIndex("����") < 0 Or .Row < 0 Then Exit Sub
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        mstrCurStatus = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
        If tbPage.Selected.Tag = Pg_�䶯��¼ Then
            Call mfrmChage.zlReLoadData(mlngCardTypeID, strCardNo)
        ElseIf tbPage.Selected.Tag = Pg_�ʻ������¼ Then
            Call mfrmConsume.zlReLoadData(lng����ID, mlngCardTypeID, strCardNo)
        Else
            Call mfrmFamily.zlReLoadData(lng����ID, mlngCardTypeID, strCardNo)
        End If
    End With
    zlCommFun.StopFlash
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub vsCardList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub
Private Sub vsCardList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsCardList, Me.Name, "����Ϣ�б�", True, zlstr.IsHavePrivs(mstrPrivs, "��������")
End Sub

Private Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���д�ӡ,Ԥ���������EXCEL
    '���:bytFunc=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '����:
    '����:
    '����:���˺�
    '����:2009-11-20 15:34:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long, lngRow As Long
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim blnCardList As Boolean
    blnCardList = Me.ActiveControl Is vsCardList
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    objPrint.Title.Text = gstrUnitName & "ҽ�ƿ����"
    
    If CStr(mArrFilter("����ʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "����ʱ�䣺" & CStr(mArrFilter("����ʱ��")(0)) & "��" & CStr(mArrFilter("����ʱ��")(1))
    End If
    If CStr(mArrFilter("��ʧʱ��")(0)) <> "1901-01-01" Then
        objRow.Add "��ʧʱ�䣺" & CStr(mArrFilter("��ʧʱ��")(0)) & "��" & CStr(mArrFilter("��ʧʱ��")(1))
    End If
    
    If objRow.count > 1 Then
        objPrint.UnderAppRows.Add objRow
        Set objRow = New zlTabAppRow
    End If
    If mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ŷ�Χ��" & CStr(mArrFilter("���ŷ�Χ")(0)) & "��" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) = "" And mArrFilter("���ŷ�Χ")(1) <> "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(1))
    ElseIf mArrFilter("���ŷ�Χ")(0) <> "" And mArrFilter("���ŷ�Χ")(1) = "" Then
        objRow.Add "���ţ�" & CStr(mArrFilter("���ŷ�Χ")(0))
    End If
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & UserInfo.����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    '���ڴ�ӡ�ؼ�����ʶ������������
    With vsCardList
        .Redraw = flexRDNone
        .GridColor = .ForeColor
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = 0
            End If
            
        Next
    End With
    
    Err = 0: On Error GoTo Errhand:
    Set objPrint.Body = vsCardList
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    '�ָ�
    With vsCardList
        For i = 0 To .Cols - 1
            .Cell(flexcpData, 0, i) = .ColWidth(i)
            If .ColHidden(i) = True Or i = 0 Then
                .ColWidth(i) = Val(.Cell(flexcpData, 0, i))
            End If
        Next
        .GridColor = &H8000000F
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub
Private Function SaveCardCancelLose() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ȡ����ʧ
    '����:���˺�
    '����:2011-06-28 22:41:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String, strCardNo As String, lngRow As Long, i As Long
    Dim lng����ID As Long, strSQL As String
    
    With vsCardList
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
        If strCardNo = "" Then Exit Function
        If .TextMatrix(.Row, .ColIndex("��ʧʱ��")) = "" Then Exit Function
        If MsgBox("�����Ҫ�Կ���Ϊ:��" & .TextMatrix(.Row, .ColIndex("����")) & "���ļ�¼����ȡ����ʧ������" & vbCrLf & _
                    "   ���ǡ�: ����ȡ����ʧ����,ȡ����Ŀ�Ƭ���ܽ���ˢ���Ȳ�����" & vbCrLf & _
                    "   ����:��������ȡ����ʧ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
    End With
    
        'Zl_����ҽ�ƿ���Ϣ_ȡ����ʧ
        strSQL = "Zl_����ҽ�ƿ���Ϣ_ȡ����ʧ("
        '  ����id_In     In ����ҽ�ƿ���Ϣ.����id%Type,
        strSQL = strSQL & "" & lng����ID & ","
        '  �����id_In   In ����ҽ�ƿ���Ϣ.�����id%Type,
        strSQL = strSQL & "" & mlngCardTypeID & ","
        '  ����_In       In ����ҽ�ƿ���Ϣ.����%Type,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ����Ա����_In In ���˱䶯��¼.����Ա����%Type
        strSQL = strSQL & "'" & UserInfo.���� & "')"
    
        Err = 0: On Error GoTo Errhand:
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        With vsCardList
            .TextMatrix(.Row, .ColIndex("��ʧ��")) = ""
            .TextMatrix(.Row, .ColIndex("��ʧ��ʽ")) = ""
            .TextMatrix(.Row, .ColIndex("��ʧʱ��")) = ""
            .TextMatrix(.Row, .ColIndex("״̬")) = "��Ч��"
        End With
        SaveCardCancelLose = True
        Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
   
Private Sub vsCardList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCardList
        Select Case Col
        Case .ColIndex("��־")
            Cancel = True
        Case Else
        End Select
    End With
End Sub

Private Sub vsCardList_DblClick()
    Dim strCardNo As String, lng����ID As Long
    '84755:���ϴ�,2015/5/15,�鿴ҽ�ƿ���Ϣʱ���벡��id
    With vsCardList
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        If strCardNo = "" Then Exit Sub
    End With
    If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_��ѯ, mlngCardTypeID, strCardNo, lng����ID) = False Then Exit Sub
End Sub

Private Sub vsCardList_GotFocus()
    zl_VsGridGotFocus vsCardList, gSysColor.lngGridColorSel
End Sub

Private Sub vsCardList_LostFocus()
    zl_VsGridLostFocus vsCardList, gSysColor.lngGridColorLost
End Sub

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
End Sub
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    Select Case Index
    Case mtxtIdx.idx_������
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case mtxtIdx.idx_��ʧ��
        If Select��Աѡ����(Me, txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then
            Exit Sub
        End If
    Case Else
        '���ڿ��Ų�֪����,�����޷���λ
    End Select
End Sub

Private Sub chkFilter_Click(Index As Integer)
    Select Case Index
    Case 0
        If chkFilter(Index).value = 0 And zlstr.ActualLen(Trim(txtName.Text)) < 4 Then
           If chkFilter(1).value = 0 Then chkFilter(1).value = 1
        End If
    Case 1
        If chkFilter(Index).value = 0 And zlstr.ActualLen(Trim(txtName.Text)) < 4 Then
           If chkFilter(0).value = 0 Then chkFilter(0).value = 1
        End If
    End Select
    dtp��ʼ����(Index).Enabled = chkFilter(Index).value = 1
    dtp��������(Index).Enabled = chkFilter(Index).value = 1
End Sub

Private Sub chkFilter_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtp��������_Change(Index As Integer)
     If dtp��������(Index).value > dtp��ʼ����(Index).MaxDate Then dtp��������(Index).value = dtp��ʼ����(Index).MaxDate
    If dtp��������(Index).value < dtp��ʼ����(Index).value Then
        dtp��ʼ����(Index).value = dtp��������(Index).value
    End If
End Sub
Private Sub dtp��������_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtp��ʼ����_Change(Index As Integer)
    If dtp��ʼ����(Index).value > dtp��������(Index).MaxDate Then dtp��ʼ����(Index).value = dtp��������(Index).MaxDate
    If dtp��������(Index).value < dtp��ʼ����(Index).value Then
        dtp��������(Index).value = dtp��ʼ����(Index).value
    End If
End Sub
Private Sub dtp��ʼ����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub zlPrepayFunc(ByVal intFunc As Integer, ByVal lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Ԥ���
    '���:intFunc-1-��Ԥ��;2-��Ԥ��;3-����,4-����תסԺ;5-סԺת����;
    '����:���˺�
    '����:2011-07-24 18:25:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun As Object, intԤ������ As Integer
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    'bytԤ������: 0-��Ԥ����(ȱʡ,���л�����),1-�������(1),2-����״̬(1); 3-����˿�(37770), 4-����תסԺ;5-סԺת����
    Select Case intFunc
    Case 1  '1.��Ԥ��
        intԤ������ = 0
    Case 2 '�˿�
        intԤ������ = 3
    Case 3: intԤ������ = 2
    Case 4: intԤ������ = 4
    Case 5: intԤ������ = 5
    End Select
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ� ����Ԥ�����տ��
    '������
    '   lngModul:��Ҫִ�еĹ������
    '   cnMain:����������ݿ�����
    '   frmMain:������
    '   strDBUser:��ǰ���ݿ��¼�û���
    '  bytCallObject:���˺����(0-Ԥ�������(ȱʡ��);1-���˷��ò�ѯ����,2-ҽ�ƿ�����)
    '  lng����ID-ȱʡ�Ĳ���ID
    '  lng��ҳID-ȱʡ����ҳID
    '  dblDefPrePayMoney-ȱʡ��Ԥ�����
    Set gfrmCardMgr = Me
    If objFun.PlusDeposit(glngSys, gcnOracle, Me, gstrDBUser, 2, lng����ID, 0, 0, intԤ������) = False Then
        Set gfrmCardMgr = Nothing
        Exit Sub
    End If
    Set gfrmCardMgr = Nothing
End Sub
 

Private Sub txtPatient_Change()
    Call AutoBrushSet(txtPatient.Text = "")
    If Trim(txtPatient.Text) = "" Then Call ClearData
End Sub
Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    Call AutoBrushSet(txtPatient.Text = "")
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "����") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub
Private Sub txtPatient_LostFocus()
    Call AutoBrushSet(False)
    Call zlCommFun.OpenIme(False)
End Sub
Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub
Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "����") Then
        '105567:���ϴ�,2017/5/25,���ż��ܵ��µ�һ������ƴ�����ܴ������뷨
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                txtPatient.Text = ""
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "�����") Or IsCardType(IDKind, "סԺ��") Or IsCardType(IDKind, "�ֻ���") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
         txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
        txtPatient.IMEMode = 0
    End If
    
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '����ˢ���ͻس�,���˳�
        Exit Sub
    End If
    
    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    If Not GetPatient(txtPatient.Text, blnCard) Then
        If blnCard Then
            Call ClearData: txtPatient.Text = ""
        Else
            Call ClearData: zlControl.TxtSelAll txtPatient
        End If
        Exit Sub
    End If
    txtPatient.PasswordChar = ""
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub AutoBrushSet(blnAutoRefrsh As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ�ˢ������
    '����:���˺�
    '����:2011-06-20 13:31:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(blnAutoRefrsh)
   If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(blnAutoRefrsh)
   Call IDKind.SetAutoReadCard(blnAutoRefrsh)
End Sub
Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
   Dim lngPreIDKind As Long
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        IDKind.IDKind = IDKind.GetKindIndex("IC����")
        txtPatient.Text = strCardNo
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
   Dim lngPreIDKind As Long
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("���֤��")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Public Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2011-06-20 09:29:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    vsCardList.Clear 1
    vsCardList.Rows = 2
End Sub

Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTmp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng����ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, blnBrushCurCardType As Boolean '�Ƿ�ˢ�ĵ�ǰ��
    Dim strCardNo As String, rsInfor As ADODB.Recordset
    Dim blnIsMobileNO As Boolean
    mlngCurPatient = 0 '��ձ���
    txtPatient.ForeColor = &HFF0000
    blnIsMobileNO = IDKind.IsMobileNo(strInput)
    If IsCardType(IDKind, "IC����") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   'ˢ����ȱʡ�Ŀ�
        mlng�����ID = Val(IDKind.GetCurCard.�ӿ����)
        If mlng�����ID <= 0 Then
            mlng�����ID = IDKind.GetDefaultCardTypeID
        End If
        strCardNo = strInput
        '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If mlng�����ID = mlngCardTypeID Then blnBrushCurCardType = True
        If GetPatiID(mlng�����ID, strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
            If blnIsMobileNO Then
                '�ֻ��Ų���
                If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
                    If lng����ID = 0 Then GoTo NotFoundPati:
                    Set rsInfor = New ADODB.Recordset
                    txtPatient.Text = "": Exit Function
                End If
            Else
                If lng����ID = 0 Then GoTo NotFoundPati:
                Set rsInfor = New ADODB.Recordset
                txtPatient.Text = "": Exit Function
            End If
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng����ID
        strWhere = strWhere & " And A.����ID=[1]"
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strWhere = strWhere & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strWhere = strWhere & " And A.����ID = (Select Nvl(Max(����ID),0) as ����ID From ������ҳ Where סԺ��=[1])"
    ElseIf IsCardType(IDKind, "����") And blnIsMobileNO Then
        '�ֻ��Ų���
        If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
    Else
        Select Case IDKind.GetCurCard.����
            Case "����", "��������￨"
                '�����:116787,����,2017/12/12,���ݡ�������ģ����ѯ����Ҫ�����������ֲ�ȥ��ѯ,
                '                              ����"����Ԥ�����"��"סԺԤ�����"��ʾ������
                If Not mblnSeekName Or zlstr.ActualLen(strInput) < 4 Then Exit Function
                strPati = _
                "Select /*+Rule */" & vbNewLine & _
                "       a.����id As ID, a.����id, max(a.����)as ����, max(a.�Ա�)as �Ա�, max(a.����)as ����, max(a.��������) as ��������, max(a.����)as ����,max(a.�����)as �����," & vbNewLine & _
                "       max( a.סԺ��)as סԺ��, max(a.��������)as ��������, max(a.���֤��)as ���֤��, max(a.��ͥ��ַ) As ��סַ, max(a.������λ)as ������λ," & vbNewLine & _
                "       LTrim(To_Char(max(Decode(b.����, 1, Nvl(b.Ԥ�����, 0), 0)), '99999999990.00')) As ����Ԥ�����," & vbNewLine & _
                "       LTrim(To_Char(max(Decode(b.����, 1, 0, Nvl(b.Ԥ�����, 0))), '99999999990.00')) As סԺԤ�����," & vbNewLine & _
                "       max(c.����) as ���� " & vbNewLine & _
                "From ������Ϣ A, ������� B,����ҽ�ƿ���Ϣ C" & vbNewLine & _
                "Where a.ͣ��ʱ�� Is Null And a.����id = b.����id(+) And a.����id=c.����id(+) And b.����(+) = 1 And Rownum < 101 And " & vbNewLine & _
                "      a.���� Like [1] And c.�����ID(+)=[2] " & vbNewLine & _
                "group by a.����id" & vbNewLine & _
                "Order by  ����,����"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����ѡ��", 1, "", "��ѡ����", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, _
                                                     blnCancel, False, True, strInput & "%", mlngCardTypeID)
                If blnCancel Then
                    Set rsInfor = New ADODB.Recordset: Exit Function
                End If
                If rsTmp Is Nothing Then GoTo NotFoundPati:
                If rsTmp.State <> 1 Then GoTo NotFoundPati:
                If rsTmp.RecordCount = 0 Then GoTo NotFoundPati:
                lng����ID = Val(Nvl(rsTmp!����ID))
                mlngCurPatient = lng����ID
                '84490:���ϴ�,2015/5/15,ͨ���������Ҳ��˳ɹ����ȡ��������
                txtPatient.Text = Nvl(rsTmp!����)
                '74309:���ϴ���2014-7-7������������ʾ��ɫ����
                Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), txtPatient.ForeColor, vbRed))
                Call LoadDataToGrid(lng����ID)
                GetPatient = True
                Exit Function
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.ҽ����=[2]"
             Case "���֤��", "�������֤��", "�������֤", "���֤"
                strInput = UCase(strInput)
                If GetPatiID("���֤", strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If GetPatiID("IC��", strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case Else
                '�������ĺ���
                If Val(IDKind.GetCurCard.�ӿ����) > 0 Then
                    mlng�����ID = IDKind.GetCurCard.�ӿ����
                     If mlng�����ID = mlngCardTypeID Then blnBrushCurCardType = True
                    If GetPatiID(mlng�����ID, strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then
                        If lng����ID = 0 Then GoTo NotFoundPati:
                        Set rsInfor = New ADODB.Recordset
                        txtPatient.Text = "": Exit Function
                    End If
                    If lng����ID = 0 Then GoTo NotFoundPati:
                    strInput = "-" & lng����ID
                    strWhere = strWhere & " And A.����ID=[1]"
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.����, strInput, False, lng����ID, strPassWord, , , , , , , , , , mlngCardTypeID) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
        End Select
    End If
    
    On Error GoTo errH
    '��ȡ������Ϣ
    strSQL = "" & _
    "   Select  A.����ID, A.����, A.����֤��,A.��������,A.����" & _
    "   From ������Ϣ A" & _
    "   Where A.ͣ��ʱ�� is NULL " & strWhere
    Set rsInfor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    
    If rsInfor.EOF Then GoTo NotFoundPati:
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtPatient.PasswordChar = ""
    '74309:���ϴ���2014-7-7������������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(rsInfor!��������), IIf(IsNull(rsInfor!����), txtPatient.ForeColor, vbRed))
    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
    txtPatient.IMEMode = 0
    txtPatient.Text = Nvl(rsInfor!����)
    lng����ID = Val(Nvl(rsInfor!����ID))
    mlngCurPatient = lng����ID
    '���Ե�ǰ����ˢ��
    Call LoadDataToGrid(lng����ID, IIf(blnBrushCurCardType, strInput, ""))
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
NotFoundPati:
    If blnBrushCurCardType And strInput <> "" Then
        If Not mbln���ƿ� Then
            If MsgBox("δ�ҵ�ָ�����Ĳ�����Ϣ,�Ƿ���п���?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then Exit Function
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_�󶨿�, mlngCardTypeID, strCardNo) = False Then Exit Function
        Else
            If MsgBox("δ�ҵ�ָ�����Ĳ�����Ϣ,�Ƿ���з���?", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbNo Then Exit Function
            If frmPatiCureCardEdit.zlShowCard(Me, mlngModule, mstrPrivs, Cr_����, mlngCardTypeID, strCardNo) = False Then Exit Function
        End If
        Call LoadDataToGrid(lng����ID, strCardNo)
    End If
    If blnCard Then
        MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����    ", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Else
        MsgBox "������Ϣδ�ҵ�,�����Ƿ�������ȷ!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
    Set rsInfor = New ADODB.Recordset
End Function
Private Sub InitListReport()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������(��֧���)
    '����:���˺�
    '����:2012-06-12 15:26:26
    '����:50122
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, objPop As Object
    Dim i As Long, j As Long
    With cbsThis.ActiveMenuBar
        For i = 1 To .Controls.count
             If .Controls(i).id = conMenu_ReportPopup Or .Controls(i).Caption Like "����" Then
                 Set objPop = cbsThis.ActiveMenuBar.Controls(i)
                   
                With objPop.CommandBar
                     For j = 1 To .Controls.count
                        varData = Split(.Controls(j).Parameter & ",,", ",")
                        If varData(1) = "ZL" & glngSys \ 100 & "_INSIDE_1107_2" Then
                            mlngListReportID = .Controls(j).id
                            mstrListReportName = .Controls(j).Caption
                            Exit Sub
                        End If
                     Next
                End With
             End If
        Next
    End With
End Sub
'��ȡidkind��Ĭ��kindֵ
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind��Ĭ��Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.����)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case "סԺ��"
          IsCardType = IDKindCtl.GetCurCard.���� = "סԺ��"
     Case "�ֻ���"
          IsCardType = IDKindCtl.GetCurCard.���� = "�ֻ���"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then Exit Function
            If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
            IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
     End Select
End Function

Private Function Checkȡ��Ժ�⿨��(lng����ID As Long, strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ��Ժ�⿨�������
    '���:lng����ID - ����ID; strCardNo - ����
    '����:��������:ȡ����,�˿�,����
    '����:����
    '����:2012-12-19 15:26:26
    '����:56599
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��Ժ���� As Boolean 'True - ��Ժ���� False - �����󶨿�
    Dim strSQL As String, msgBoxResult As String
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:
    strSQL = "" & _
    "   Select count(1) as ��Ժ���� From סԺ���ü�¼ Where ��¼����=5 And ����ID=[1] And ʵ��Ʊ��=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, strCardNo)
    If rsTemp.EOF = False Then
        bln��Ժ���� = Val(Nvl(rsTemp!��Ժ����)) > 0
        If bln��Ժ���� = True And mbln���� = True Then
            msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "����:" & strCardNo & "Ϊ��Ժ����,���Ƿ����ȡ���ÿ��İ�?", "ȡ����,�˿�,����", Me, vbQuestion)
            Checkȡ��Ժ�⿨�� = msgBoxResult
            If Checkȡ��Ժ�⿨�� = "" Then Checkȡ��Ժ�⿨�� = "����"
            Exit Function
        End If
    End If
    Checkȡ��Ժ�⿨�� = "ȡ����"
    Exit Function
ErrHandl:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check�˿�(strCardNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�˿��������
    '���:strCardNo - ����
    '����:��������:ȡ����,�˿�,����
    '����:����
    '����:2012-12-19 15:26:26
    '����:56599
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln�󶨿� As Boolean
    Dim strSQL As String, msgBoxResult As String
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandl:
     
    bln�󶨿� = zlIsCardBinding
    If bln�󶨿� = True And mbln���� = True And mbln���ƿ� = False Then
        msgBoxResult = zl9ComLib.zlCommFun.ShowMsgbox(gstrSysName, "����:" & strCardNo & "��Ϊ�󶨿�,�Ƿ�ȡ����?", "��,��", Me, vbQuestion)
        Select Case msgBoxResult
            Case "��"
                Check�˿� = "ȡ����"
            Case "��"
                Check�˿� = "����"
            Case Else
                Check�˿� = "����"
        End Select
        Exit Function
    End If
    Check�˿� = "�˿�"
    Exit Function
ErrHandl:
     If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-07-22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CreatePublicPatient = True
End Function

Private Sub PrintBill()
'���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
'bytMode=0-�ش�,1-����
    Dim strCardNo As String, lng����ID As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strOperName As String, strDate As String
    Dim blnStartFactUseType  As Boolean, strUseType As String
    Dim blnHaveData As Boolean, strFormat As String
    Dim objfrmPrint As frmPrint
    
    Set objfrmPrint = New frmPrint
    Load objfrmPrint
    With vsCardList
        lng����ID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        strCardNo = Trim(.TextMatrix(.Row, .ColIndex("����")))
        strOperName = Trim(.TextMatrix(.Row, .ColIndex("������")))
        strDate = .TextMatrix(.Row, .ColIndex("��������"))
        If strCardNo = "" Then ShowMsgbox "ûѡ����ص�ҽ�ƿ���": Exit Sub
    End With
    
    If mPrint.intPrintMode = 0 Then
        '��ӡ����/�󶨿�ƾ��
        strFormat = IIf(mPrint.bytPrintBoundCard = 0, "", "ReportFormat=" & mPrint.bytPrintBoundCard)
        Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1107", Me, "�����ID=" & mlngCardTypeID, "NO=" & strCardNo, "����=" & strCardNo, "�ɿ�=" & 0, "�Ҳ�=" & 0, "PrintEmpty=0", strFormat, 2)
        Exit Sub
    End If
    
    strSQL = "Select A.No,B.ID From סԺ���ü�¼ A,(Select A.ID,A.NO From Ʊ�ݴ�ӡ���� A,Ʊ��ʹ����ϸ B Where A.ID = B.��ӡID And A.�������� = 5 And B.Ʊ�� = 1) B " & _
            " Where A.no=B.NO(+) And A.ʵ��Ʊ�� = [1] And Nvl(A.����,0) = [2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "������¼", strCardNo, mlngCardTypeID)
    With rsTemp
        If .EOF Then
            MsgBox "��ǰ����" & strCardNo & "�ķ��������ں����ݱ���!" & vbCrLf _
                & "����ϵͳ����Ա��ϵ,ת�뵽�������ݱ��ٲ���!", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    If mPrint.intPrintMode = 3 Then
        If Not BillOperCheck(8, strOperName, CDate(strDate), "�ش�") Then Exit Sub
    Else
        If Not IsNull(rsTemp!id) Then
            MsgBox "��ǰ���������Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    If gblnStartFactUseType Then
        mPrint.strUseType = zl_GetInvoiceUserType(lng����ID, 0, 0)
    End If
    mPrint.strPrintNo = Nvl(rsTemp!NO)
    
    If Not objfrmPrint.RePrintBill(Me, strCardNo, mlngCardTypeID, mPrint.strUseType, _
                mPrint.strPrintNo, mPrint.intPrintMode, mPrint.bytPrintPayCard, True) Then Exit Sub
End Sub

