VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPriceGradeManage 
   Caption         =   "�۸�ȼ�����"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10530
   Icon            =   "frmPriceGradeManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPriceGrade 
      BorderStyle     =   0  'None
      Height          =   3585
      Left            =   1410
      ScaleHeight     =   3585
      ScaleWidth      =   2505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2505
      Begin MSComctlLib.ListView lvwPriceGrade 
         Height          =   1905
         Left            =   300
         TabIndex        =   5
         Top             =   990
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   3360
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils32"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         OLEDragMode     =   1
         NumItems        =   0
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPriceGrade 
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   540
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�۸�ȼ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Shape shpPriceGrade 
         BorderColor     =   &H80000003&
         Height          =   255
         Left            =   210
         Top             =   60
         Width           =   675
      End
   End
   Begin VB.PictureBox picGradeApply 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   4170
      ScaleHeight     =   3615
      ScaleWidth      =   2325
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2325
      Begin VSFlex8Ctl.VSFlexGrid vsfGradeApply 
         Height          =   2145
         Left            =   270
         TabIndex        =   2
         Top             =   930
         Width           =   1785
         _cx             =   3149
         _cy             =   3784
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
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin VB.Shape shpGradeApply 
         BorderColor     =   &H80000003&
         Height          =   345
         Left            =   270
         Top             =   180
         Width           =   495
      End
      Begin XtremeSuiteControls.ShortcutCaption sccGradeApply 
         Height          =   300
         Left            =   210
         TabIndex        =   3
         Top             =   630
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   529
         _StockProps     =   6
         Caption         =   "�۸�ȼ�Ӧ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7275
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   635
      SimpleText      =   $"frmPriceGradeManage.frx":030A
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPriceGradeManage.frx":0351
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13494
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
   Begin MSComctlLib.ImageList ils32 
      Left            =   2625
      Top             =   345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":0BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":103D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":1357
            Key             =   "Default"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":1C31
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1980
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":250B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":2963
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":2C7D
            Key             =   "Default"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceGradeManage.frx":3217
            Key             =   "Stop"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPriceGradeManage.frx":37B1
      Left            =   1080
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPriceGradeManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModule As Long
Private mblnUnload As Boolean
Private mblnFirst As Boolean

Private Enum PaneIndex
    Pane_PriceGrade = 1
    Pane_GradeApply = 2
End Enum

Private Enum ColIndex
    'LVW_���� = 0
    LVW_���� = 1
    LVW_���� = 2
    LVW_����ҩƷ = 3
    LVW_�������� = 4
    LVW_������ͨ��Ŀ = 5
    LVW_����ʱ�� = 6
    LVW_����ʱ�� = 7
    LVW_�Ƿ�ͣ�� = 8
    
    VSF_Ӧ�ó��� = 0
    VSF_���� = 1
    VSF_���� = 2
End Enum

Private mblnShowStopedGrade As Boolean
Private mbytLvwViewType As Byte

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnStop As Boolean
    
    On Error Resume Next
    Select Case Control.ID
    Case conMenu_EditPopup '�༭
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��ɾ��") _
                        Or zlStr.IsHavePrivs(mstrPrivs, "ͣ��") _
                        Or zlStr.IsHavePrivs(mstrPrivs, "����")
    Case conMenu_Edit_NewItem '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = Control.Visible
    Case conMenu_Edit_Modify '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_�Ƿ�ͣ��)) = 0
        End If
    Case conMenu_Edit_Delete 'ɾ��
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "��ɾ��")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_�Ƿ�ͣ��)) = 0
        End If
    Case conMenu_Edit_Reuse 'ͣ��
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ͣ��")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_�Ƿ�ͣ��)) = 0
        End If
    Case conMenu_Edit_Stop '����
        Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
        Control.Enabled = Control.Visible And Not lvwPriceGrade.SelectedItem Is Nothing
        If Control.Enabled Then
            Control.Enabled = Val(lvwPriceGrade.SelectedItem.SubItems(LVW_�Ƿ�ͣ��)) = 1
        End If
    End Select
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHandler
    If mblnUnload Then Unload Me: Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Exit Sub
ErrHandler:
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim strLvwCols As String
    Err = 0: On Error GoTo ErrHandler
    mblnUnload = False
    mblnFirst = True
    mstrPrivs = gstrPrivs
    mlngModule = glngModul
    
    mblnShowStopedGrade = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", 0)) = 1
    mbytLvwViewType = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ListView��ͼ", 0))
    
    If DefMainCommandBars() = False Then mblnUnload = True: Exit Sub
    If InitPanel() = False Then mblnUnload = True: Exit Sub
    If InitVsfGrid() = False Then mblnUnload = True: Exit Sub
    
    Call RestoreWinState(Me, App.ProductName)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    
    '���ListView�Ļ�δ�����ã������һ��ʹ�ã��Ǿ͵���ȱʡ�ĳ�ʼ��
    strLvwCols = "����,1600,0,1;����,1000,0,1;����,1200,0,0;" & _
        "����ҩƷ,0,2,0;��������,0,2,0;������ͨ��Ŀ,0,2,0;����ʱ��,1900,2,0;����ʱ��,1900,2,0;�Ƿ�ͣ��,0,0,1"
    If lvwPriceGrade.ColumnHeaders.Count = 0 Then
        zlControl.LvwSelectColumns lvwPriceGrade, strLvwCols, True
    End If
    Call SetLvwViewType(mbytLvwViewType)
    Call LoadPriceGrade
    Exit Sub
ErrHandler:
    mblnUnload = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitPanel() As Boolean
    '����:��ʼ�����沼��
    '����:���óɹ�,����true,���򷵻�False
    Dim objPane As Pane

    Err = 0: On Error GoTo ErrHandler
    Set objPane = dkpMain.CreatePane(Pane_PriceGrade, 200, 120, DockLeftOf, Nothing)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 65

    Set objPane = dkpMain.CreatePane(Pane_GradeApply, 700, 400, DockRightOf, objPane)
    objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
    objPane.MinTrackSize.Width = 65

    With dkpMain
        .SetCommandBars cbsMain
        .VisualTheme = ThemeOffice2003
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = True 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    InitPanel = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DefMainCommandBars() As Boolean
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    Dim cbrControl As CommandBarControl, cbrSubControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup, cbrToolBar As CommandBar
    Dim objPopupControl As CommandBarControl

    Err = 0: On Error GoTo ErrHandler
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .ShowExpandButtonAlways = False
    End With
    cbsMain.EnableCustomization False

    '�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagHideWrap Or xtpFlagStretched

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ͣ��(&S)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "����(&T)")
    End With

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        With cbrControl.CommandBar.Controls
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False): cbrSubControl.Checked = True
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False): cbrSubControl.Checked = True
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False): cbrSubControl.Checked = True
        End With
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.BeginGroup = True
        cbrControl.Checked = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��(&G)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����(&D)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾͣ�õȼ�(&S)"): cbrControl.BeginGroup = True
        cbrControl.Checked = mblnShowStopedGrade
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�����")
        With cbrControl.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "������ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "������̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '����������
    Set cbrToolBar = cbsMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.ContextMenuPresent = False
    cbrToolBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ͣ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Stop, "����")
        
        Set objPopupControl = .Add(xtpControlSplitButtonPopup, conMenu_View_Append, "�鿴"): objPopupControl.BeginGroup = True
        objPopupControl.IconId = conMenu_View_LargeICO
        With objPopupControl.CommandBar.Controls
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_LargeICO, "��ͼ��")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_MinICO, "Сͼ��")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_ListICO, "�б�")
            Set cbrSubControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "��ϸ����")
        End With
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

    '�����
    With cbsMain.KeyBindings
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F12, conMenu_File_Parameter
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("B"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("D"), conMenu_Edit_Delete
    End With

    '���ò����ò˵�
    With cbsMain.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
    End With

    DefMainCommandBars = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitVsfGrid() As Boolean
    '���ܣ���ʼ������ؼ�
    Dim strHead As String, varData As Variant
    Dim i As Long

    Err = 0: On Error GoTo ErrHandler
    With vsfGradeApply
        .redraw = flexRDNone
        .Rows = 1
        .FixedCols = 0: .FixedRows = 1
        '
        strHead = ",1,250|����,4,700|����,1,2000"
        varData = Split(strHead, "|")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedAlignment(-1) = flexAlignCenterCenter
        .RowHeightMin = 300
        '.ColHidden(VSF_Ӧ�ó���) = True'��������"Ӧ�ó���"�У����غ���鲻��չ����������ȱʡ����"Ӧ�ó���"���Ϊ10

        .AllowSelection = False
        .AllowBigSelection = False
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow

        .HighLight = flexHighlightAlways
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
        .PicturesOver = True '������ͼƬ����

        .BackColorBkg = vbWindowBackground
        .SheetBorder = vbWindowBackground
        
'        '����������,�����û�ѡ����ʾ��
'        For i = 0 To .Cols - 1
'            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
'            Select Case i
'            Case LVW_ID
'                 .ColData(i) = "-1|1"
'            Case LVW_����
'                .ColData(i) = "1|0"
'            End Select
'        Next
        .redraw = flexRDBuffered
    End With
    InitVsfGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim frmEdit As frmPriceGradeEdit
    Dim strItem As String

    Err = 0: On Error GoTo ErrHandler
    Select Case Control.ID
    'bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Case conMenu_File_PrintSet '��ӡ����
        Call zlPrintSet
    Case conMenu_File_Preview 'Ԥ��
        Call ZlDataPrint(2)
    Case conMenu_File_Print '��ӡ
        Call ZlDataPrint(1)
    Case conMenu_File_Excel '�����Excel��
        Call ZlDataPrint(3)
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_Edit_NewItem '����
        Set frmEdit = New frmPriceGradeEdit
        If frmEdit.ShowMe(Me, 0, , strItem) Then Call LoadPriceGrade(strItem)
    Case conMenu_Edit_Modify '����
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        Set frmEdit = New frmPriceGradeEdit
        strItem = lvwPriceGrade.SelectedItem.Text
        If frmEdit.ShowMe(Me, 1, strItem) Then Call LoadPriceGrade
    Case conMenu_Edit_Delete 'ɾ��
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        Set frmEdit = New frmPriceGradeEdit
        strItem = lvwPriceGrade.SelectedItem.Text
        If frmEdit.ShowMe(Me, 2, strItem) Then Call LoadPriceGrade
    Case conMenu_Edit_Reuse 'ͣ��
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        strItem = lvwPriceGrade.SelectedItem.Text
        If StopAndStartPriceGrade(strItem, True) Then Call LoadPriceGrade
    Case conMenu_Edit_Stop '����
        If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
        strItem = lvwPriceGrade.SelectedItem.Text
        If StopAndStartPriceGrade(strItem, False) Then Call LoadPriceGrade
    Case conMenu_View_ToolBar_Button '��׼��ť
        Control.Checked = Not Control.Checked
        cbsMain(2).Visible = Control.Checked
        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Text, , True)
        objControl.Enabled = Control.Checked
        Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_View_ToolBar_Size, , True)
        objControl.Enabled = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '�ı���ǩ
        Control.Checked = Not Control.Checked
        For Each objControl In cbsMain(2).Controls
            objControl.Style = IIF(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Not Control.Checked
        cbsMain.Options.LargeIcons = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Not Control.Checked
        stbThis.Visible = Control.Checked
        cbsMain.RecalcLayout
    Case conMenu_View_Append
        mbytLvwViewType = mbytLvwViewType + 1
        If mbytLvwViewType < lvwIcon Or mbytLvwViewType > lvwReport Then
            mbytLvwViewType = lvwIcon
        End If
        Call SetLvwViewType(mbytLvwViewType)
    Case conMenu_View_LargeICO '��ͼ��
        Call SetLvwViewType(lvwIcon)
        mbytLvwViewType = lvwIcon
    Case conMenu_View_MinICO 'Сͼ��
        Call SetLvwViewType(lvwSmallIcon)
        mbytLvwViewType = lvwSmallIcon
    Case conMenu_View_ListICO '�б�
        Call SetLvwViewType(lvwList)
        mbytLvwViewType = lvwList
    Case conMenu_View_DetailsICO '��ϸ����
        Call SetLvwViewType(lvwReport)
        mbytLvwViewType = lvwReport
    Case conMenu_View_ShowStoped '��ʾͣ�õȼ�
        Control.Checked = Not Control.Checked
        zlDatabase.SetPara "��ʾͣ�õȼ�", IIF(Control.Checked, 1, 0), glngSys, mlngModule
        mblnShowStopedGrade = Control.Checked
        cbsMain.RecalcLayout
        Call LoadPriceGrade
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPriceGrade
    Case conMenu_Help_Help '��������
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home '������ҳ
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '���ڡ�
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case Else
        If (Control.ID >= conMenu_ReportPopup * 100# + 1 And Control.ID <= conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '�����Զ��屨��
            Call ZlCallCustomReprot(Me, Val(Split(Control.Parameter, ",")(0)), Trim(Split(Control.Parameter, ",")(1)))
        End If
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetLvwViewType(ByVal bytView As Byte)
    '����:����ListView�ؼ��鿴�˵�״̬���Լ�����ListView�ؼ��鿴��ʽ
    '��Σ�
    '   bytView: 0-lvwIcon ��ȱʡ��ͼ��
    '            1-lvwSmallIcon  Сͼ��
    '            2-lvwList �б�
    '            3-lvwReport ����
    Dim objControl As CommandBarControl
    Dim objPopupControl As CommandBarControl
    
    Err = 0: On Error GoTo ErrHandler
    '�˵���
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Find(, conMenu_View_LargeICO, , True): objControl.Checked = (bytView = lvwIcon)
        Set objControl = .Find(, conMenu_View_MinICO, , True): objControl.Checked = (bytView = lvwSmallIcon)
        Set objControl = .Find(, conMenu_View_ListICO, , True): objControl.Checked = (bytView = lvwList)
        Set objControl = .Find(, conMenu_View_DetailsICO, , True): objControl.Checked = (bytView = lvwReport)
    End With
    '������
    With cbsMain(2).Controls
        Set objPopupControl = .Find(, conMenu_View_Append, , True)
        Set objControl = .Find(, conMenu_View_LargeICO, , True): objControl.Checked = (bytView = lvwIcon)
        Set objControl = .Find(, conMenu_View_MinICO, , True): objControl.Checked = (bytView = lvwSmallIcon)
        Set objControl = .Find(, conMenu_View_ListICO, , True): objControl.Checked = (bytView = lvwList)
        Set objControl = .Find(, conMenu_View_DetailsICO, , True): objControl.Checked = (bytView = lvwReport)
    End With

    Select Case bytView
    Case lvwIcon
        objPopupControl.IconId = conMenu_View_LargeICO
        lvwPriceGrade.View = lvwIcon
    Case lvwSmallIcon
        objPopupControl.IconId = conMenu_View_MinICO
        lvwPriceGrade.View = lvwSmallIcon
    Case lvwList
        objPopupControl.IconId = conMenu_View_ListICO
        lvwPriceGrade.View = lvwList
    Case lvwReport
        objPopupControl.IconId = conMenu_View_DetailsICO
        lvwPriceGrade.View = lvwReport
    End Select
    cbsMain.RecalcLayout
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlCallCustomReprot(ByVal frmMain As Form, ByVal lngSys As Long, strReprotName As String)
    '����:������ص��Զ��屨��
    Err = 0: On Error GoTo ErrHandler
    Call ReportOpen(gcnOracle, lngSys, strReprotName, frmMain)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Err = 0: On Error Resume Next
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Err = 0: On Error GoTo ErrHandler
    Select Case Item.ID
    Case Pane_PriceGrade
        Item.Handle = picPriceGrade.hwnd
    Case Pane_GradeApply
        Item.Handle = picGradeApply.hwnd
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error GoTo ErrHandler
    
    mblnUnload = False
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "��ʾͣ��", IIF(mblnShowStopedGrade, 1, 0)
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "ListView��ͼ", mbytLvwViewType
    Call SaveWinState(Me, App.ProductName)
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlDataPrint(bytMode As Byte)
    '����:���д�ӡ,Ԥ���������EXCEL
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objVsfPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim objLvwPrint As New zlPrintLvw
    Dim bytR As Byte
    
    Err = 0: On Error GoTo ErrHandler
    If Me.ActiveControl Is vsfGradeApply Then
        'VSFlexGrid
        Set objVsfPrint.Body = vsfGradeApply
        objVsfPrint.Title.Text = "�۸�ȼ�Ӧ��"
        
        objVsfPrint.Title.Font.Name = "����_GB2312"
        objVsfPrint.Title.Font.Size = 18
        objVsfPrint.Title.Font.Bold = True
        
        
        Set objRow = New zlTabAppRow
        objRow.Add "��ӡ�ˣ�" & gstrUserName
        objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
        objVsfPrint.BelowAppRows.Add objRow

        If bytMode = 1 Then
            bytR = zlPrintAsk(objVsfPrint)
            If bytR <> 0 Then zlPrintOrView1Grd objVsfPrint, bytR
        Else
            zlPrintOrView1Grd objVsfPrint, bytMode
        End If
    Else
        'ListView
        Set objLvwPrint.Body.objData = lvwPriceGrade
        objLvwPrint.Title.Text = "�۸�ȼ�"

        objLvwPrint.Title.Font.Name = "����_GB2312"
        objLvwPrint.Title.Font.Size = 18
        objLvwPrint.Title.Font.Bold = True
        
        objLvwPrint.BelowAppItems.Add "��ӡ�ˣ�" & gstrUserName
        objLvwPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
        
        If bytMode = 1 Then
            bytR = zlPrintAsk(objLvwPrint)
            If bytR <> 0 Then zlPrintOrViewLvw objLvwPrint, bytR
        Else
            zlPrintOrViewLvw objLvwPrint, bytMode
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Err = 0: On Error Resume Next
    '����б�ͷ����
    lvwPriceGrade.Sorted = True
    lvwPriceGrade.SortKey = ColumnHeader.Index - 1
    lvwPriceGrade.SortOrder = IIF(lvwPriceGrade.SortOrder = lvwDescending, lvwAscending, lvwDescending)
End Sub

Private Sub lvwPriceGrade_DblClick()
    Dim strItem As String
    Dim frmEdit As New frmPriceGradeEdit
    
    Err = 0: On Error GoTo ErrHandler
    If lvwPriceGrade.SelectedItem Is Nothing Then Exit Sub
    strItem = lvwPriceGrade.SelectedItem.Text
    
    If zlStr.IsHavePrivs(mstrPrivs, "��ɾ��") _
        And Val(lvwPriceGrade.SelectedItem.SubItems(LVW_�Ƿ�ͣ��)) = 0 Then '����
        If frmEdit.ShowMe(Me, 1, strItem) Then Call LoadPriceGrade
    Else '�鿴
        frmEdit.ShowMe Me, 3, strItem
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHandler
    If lvwPriceGrade.Tag = Item.Text Then Exit Sub
    lvwPriceGrade.Tag = Item.Text
    Call LoadPriceGradeApply(Item.Text)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwPriceGrade_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    Err = 0: On Error GoTo ErrHandler
'    If lvwPriceGrade.Visible And lvwPriceGrade.Enabled Then lvwPriceGrade.SetFocus
    If Not (Button = vbRightButton) Then Exit Sub
    If Not Me.ActiveControl Is lvwPriceGrade Then Exit Sub
    
    Set objPopup = cbsMain.FindControl(xtpControlPopup, conMenu_EditPopup, , True)
    If objPopup Is Nothing Then Exit Sub
    If objPopup.Visible = False Then Exit Sub
    objPopup.CommandBar.ShowPopup
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPriceGrade_Resize()
    Err = 0: On Error Resume Next
    shpPriceGrade.Move 0, 0, picPriceGrade.ScaleWidth, picPriceGrade.ScaleHeight
    sccPriceGrade.Move 10, 10, picPriceGrade.ScaleWidth - 30
    With lvwPriceGrade
        .Left = sccPriceGrade.Left
        .Top = sccPriceGrade.Top + sccPriceGrade.Height
        .Width = picPriceGrade.ScaleWidth - .Left - 20
        .Height = picPriceGrade.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub picGradeApply_Resize()
    Err = 0: On Error Resume Next
    shpGradeApply.Move 0, 0, picGradeApply.ScaleWidth, picGradeApply.ScaleHeight
    sccGradeApply.Move 10, 10, picGradeApply.ScaleWidth - 30
    With vsfGradeApply
        .Left = sccGradeApply.Left
        .Top = sccGradeApply.Top + sccGradeApply.Height
        .Width = picGradeApply.ScaleWidth - .Left - 20
        .Height = picGradeApply.ScaleHeight - .Top - 20
    End With
End Sub

Private Sub vsfGradeApply_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = VSF_Ӧ�ó��� Then Cancel = True
End Sub

Private Sub vsfGradeApply_DblClick()
    Call lvwPriceGrade_DblClick
End Sub

Private Sub vsfGradeApply_GotFocus()
    vsfGradeApply.BackColorSel = vbHighlight
End Sub

Private Sub vsfGradeApply_LostFocus()
    vsfGradeApply.BackColorSel = &HE0E0E0
End Sub

Private Function LoadPriceGrade(Optional ByVal strSelectItem As String) As Boolean
    '���ؼ۸�ȼ�
    '��Σ�
    '   strSelectItem ȱʡѡ����Ŀ,�շѼ۸�ȼ�����
    Dim strSQL As String, strWhere As String
    Dim rsData As ADODB.Recordset
    Dim objListItem As ListItem, i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If strSelectItem = "" Then
        If Not lvwPriceGrade.SelectedItem Is Nothing Then
            strSelectItem = lvwPriceGrade.SelectedItem.Text
        End If
    End If
    
    lvwPriceGrade.ListItems.Clear
    lvwPriceGrade.Tag = ""
    vsfGradeApply.Clear 1: vsfGradeApply.Rows = vsfGradeApply.FixedRows
    If mblnShowStopedGrade = False Then
        '����ʾͣ�ü۸�ȼ�
        strWhere = " And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01','yyyy-mm-dd'))"
    End If
    strSQL = "Select ����, ����, ����, �Ƿ�����ҩƷ, �Ƿ���������, �Ƿ�������ͨ��Ŀ, ����ʱ��, ����ʱ��," & vbNewLine & _
            "        Decode(Nvl(����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), 0, 1) As �Ƿ�ͣ��" & vbNewLine & _
            " From �շѼ۸�ȼ�" & vbNewLine & _
            " Where 1 = 1 " & strWhere & vbNewLine & _
            " Order By ����"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "�շѼ۸�ȼ�")
    If rsData.RecordCount = 0 Then LoadPriceGrade = True: Exit Function
    
    '����,����,����,����ҩƷ,��������,������ͨ��Ŀ,����ʱ��,����ʱ��
    Do While Not rsData.EOF
        Set objListItem = lvwPriceGrade.ListItems.Add(, "K" & Nvl(rsData!����), Nvl(rsData!����), "Default", "Default")
        objListItem.SubItems(LVW_����) = Nvl(rsData!����)
        objListItem.SubItems(LVW_����) = Nvl(rsData!����)
        objListItem.SubItems(LVW_����ҩƷ) = IIF(Val(Nvl(rsData!�Ƿ�����ҩƷ)) = 1, "��", "")
        objListItem.SubItems(LVW_��������) = IIF(Val(Nvl(rsData!�Ƿ���������)) = 1, "��", "")
        objListItem.SubItems(LVW_������ͨ��Ŀ) = IIF(Val(Nvl(rsData!�Ƿ�������ͨ��Ŀ)) = 1, "��", "")
        objListItem.SubItems(LVW_����ʱ��) = Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
        objListItem.SubItems(LVW_����ʱ��) = Format(Nvl(rsData!����ʱ��), "yyyy-mm-dd hh:mm:ss")
        objListItem.SubItems(LVW_�Ƿ�ͣ��) = Val(Nvl(rsData!�Ƿ�ͣ��))
        If Val(Nvl(rsData!�Ƿ�ͣ��)) = 1 Then
            '�ı�ͣ�ü۸�ȼ���ͼ���������ɫ
            objListItem.Icon = "Stop"
            objListItem.SmallIcon = "Stop"
            objListItem.ForeColor = vbRed
            For i = 1 To objListItem.ListSubItems.Count
                objListItem.ListSubItems(i).ForeColor = vbRed
            Next
        End If
        If Nvl(rsData!����) = strSelectItem Then
            objListItem.Selected = True
        End If
        rsData.MoveNext
    Loop
    If Not lvwPriceGrade.SelectedItem Is Nothing Then
        Call lvwPriceGrade_ItemClick(lvwPriceGrade.SelectedItem)
    End If
    LoadPriceGrade = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPriceGradeApply(ByVal strPriceGrade As String) As Boolean
    '���ؼ۸�ȼ�Ӧ��
    '��Σ�
    '   strPriceGrade �շѼ۸�ȼ�����
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim i  As Long, j  As Long, strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    With vsfGradeApply
        .redraw = flexRDNone
        .Clear 1
        .Rows = 1
        
        strSQL = "Select Decode(Nvl(a.����, 0), 0, 'Ժ��', 'ҽ�Ƹ��ʽ') As Ӧ�ó���," & vbNewLine & _
                "        Decode(Nvl(a.����, 0), 0, b.���, c.����) As ����," & vbNewLine & _
                "        Decode(Nvl(a.����, 0), 0, b.����, c.����) As ����" & vbNewLine & _
                " From �շѼ۸�ȼ�Ӧ�� A, Zlnodelist B, ҽ�Ƹ��ʽ C" & vbNewLine & _
                " Where a.վ�� = b.���(+) And a.ҽ�Ƹ��ʽ = c.����(+) And �۸�ȼ� = [1]" & vbNewLine & _
                " Order By a.����, ����"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "�շѼ۸�ȼ�Ӧ��", strPriceGrade)
        If rsData.RecordCount = 0 Then
            .redraw = flexRDBuffered
            LoadPriceGradeApply = True
            Exit Function
        End If
        
        .Rows = rsData.RecordCount + 1
        lngRow = 1
        Do While Not rsData.EOF
            .TextMatrix(lngRow, VSF_Ӧ�ó���) = Nvl(rsData!Ӧ�ó���)
            .TextMatrix(lngRow, VSF_����) = Nvl(rsData!����)
            .TextMatrix(lngRow, VSF_����) = Nvl(rsData!����)
            lngRow = lngRow + 1
            rsData.MoveNext
        Loop
        
        '������ʾ
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True

        .Subtotal flexSTNone, VSF_Ӧ�ó���, , , , , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline VSF_Ӧ�ó���
        .OutlineCol = VSF_Ӧ�ó���

        .MergeCells = flexMergeRestrictRows
        .MergeRow(-1) = False
        
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) Then
                .Cell(flexcpText, i, 0, i, .Cols - 1) = .TextMatrix(i + 1, VSF_Ӧ�ó���)
                .MergeRow(i) = True '���кϲ�
                .IsCollapsed(i) = flexOutlineExpanded  '�Ƿ�չ��״̬
            End If
        Next
        .redraw = flexRDBuffered
    End With
    LoadPriceGradeApply = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function StopAndStartPriceGrade(ByVal str�۸�ȼ� As String, _
    ByVal blnStop As Boolean) As Boolean
    'ͣ��/���ü۸�ȼ�
    '��Σ�
    '   str�۸�ȼ� �շѼ۸�ȼ�����
    '   blnStop �Ƿ�ͣ��
    Dim strSQL As String, strWhere As String
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select Decode(Nvl(����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')), To_Date('3000-01-01', 'yyyy-mm-dd'), 0, 1) As �Ƿ�ͣ��" & vbNewLine & _
            " From �շѼ۸�ȼ�" & vbNewLine & _
            " Where ���� = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�۸�ȼ�", str�۸�ȼ�)
    If rsTemp.EOF Then
        MsgBox "��ǰ�۸�ȼ������ѱ�����ɾ������ˢ�º�鿴...", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Val(Nvl(rsTemp!�Ƿ�ͣ��)) = 1 Then
        If blnStop Then
            MsgBox "��ǰ�۸�ȼ��ѱ�ͣ�ã������ٴ�ͣ�á�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf blnStop = False Then
        MsgBox "��ǰ�۸�ȼ���������״̬�����������á�", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If blnStop = False Then
        '����ü۸�ȼ����ú󣬻ᵼ��һ��վ����ڶ����Ч�ļ۸�ȼ�����һ��ҽ�Ƹ��ʽ���ڶ����Ч�ļ۸�ȼ�����������
        strSQL = "Select Decode(Nvl(a.����, 0), 0, 'Ժ��', 'ҽ�Ƹ��ʽ') As Ӧ�ó���," & vbNewLine & _
                "        Decode(Nvl(a.����, 0), 0, c.����, a.ҽ�Ƹ��ʽ) As ����, a.�۸�ȼ�" & vbNewLine & _
                " From �շѼ۸�ȼ�Ӧ�� A, �շѼ۸�ȼ�Ӧ�� B, Zlnodelist C, �շѼ۸�ȼ� D" & vbNewLine & _
                " Where a.���� = b.���� And (a.վ�� = b.վ�� Or a.ҽ�Ƹ��ʽ = b.ҽ�Ƹ��ʽ) And a.վ�� = c.���(+)" & vbNewLine & _
                "       And a.�۸�ȼ� = d.���� And b.�۸�ȼ� = [1] And a.�۸�ȼ� <> [1]" & vbNewLine & _
                "       And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�۸�ȼ�", str�۸�ȼ�)
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                strTemp = strTemp & vbCrLf & Nvl(rsTemp!����) & "��" & Nvl(rsTemp!�۸�ȼ�)
                rsTemp.MoveNext
            Loop
            If MsgBox("����һ��Ժ����һ��ҽ�Ƹ��ʽֻ������һ����Ч�ļ۸�ȼ��������������õļ۸�ȼ�Ӧ���е�" & _
                "����Ժ����ҽ�Ƹ��ʽ������������Ч�ļ۸�ȼ�������������������������ЩԺ����ҽ�Ƹ��ʽ��������Ч�۸�ȼ���" & _
                "Ȼ��Ӧ�õ�ǰ�۸�ȼ����Ƿ������" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    
    If MsgBox("��ȷ��Ҫ" & IIF(blnStop, "ͣ��", "����") & "����Ϊ��" & str�۸�ȼ� & "���ļ۸�ȼ���", _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    
    If blnStop Then 'ͣ��
        'Zl_�շѼ۸�ȼ�_Stop(
        strSQL = "Zl_�շѼ۸�ȼ�_Stop("
        '   ����_In �շѼ۸�ȼ�.����%Type)
        strSQL = strSQL & "'" & str�۸�ȼ� & "')"
    Else '����
        'Zl_�շѼ۸�ȼ�_Start(
        strSQL = "Zl_�շѼ۸�ȼ�_Start("
        '   ����_In �շѼ۸�ȼ�.����%Type)
        strSQL = strSQL & "'" & str�۸�ȼ� & "')"
    End If
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    StopAndStartPriceGrade = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
