VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmLisView 
   Caption         =   "������"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15540
   Icon            =   "frmLISView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   15540
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6465
      Left            =   270
      ScaleHeight     =   6465
      ScaleWidth      =   3705
      TabIndex        =   6
      Top             =   435
      Width           =   3705
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   3765
         Left            =   30
         TabIndex        =   7
         Top             =   1530
         Width           =   2970
         _Version        =   589884
         _ExtentX        =   5239
         _ExtentY        =   6641
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFind 
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   45
         ScaleHeight     =   1275
         ScaleWidth      =   3210
         TabIndex        =   8
         Top             =   105
         Width           =   3210
         Begin VB.ComboBox cboPages 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   885
            Width           =   1335
         End
         Begin VB.CommandButton cmd��Ŀ 
            Caption         =   "��"
            Height          =   300
            Left            =   2325
            TabIndex        =   15
            Top             =   480
            Width           =   350
         End
         Begin VB.TextBox txt��Ŀ 
            Height          =   300
            Left            =   510
            TabIndex        =   13
            Top             =   495
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   510
            TabIndex        =   10
            Top             =   105
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106233859
            CurrentDate     =   39819
         End
         Begin VB.CommandButton cmdOK 
            Height          =   300
            Left            =   2745
            Picture         =   "frmLISView.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   350
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1830
            TabIndex        =   11
            Top             =   105
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   106233859
            CurrentDate     =   39819
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            Caption         =   "סԺ����"
            Height          =   180
            Left            =   90
            TabIndex        =   17
            Top             =   930
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   90
            TabIndex        =   14
            Top             =   555
            Width           =   660
         End
         Begin VB.Label lblinfo 
            Caption         =   "����"
            Height          =   180
            Left            =   90
            TabIndex        =   12
            Top             =   165
            Width           =   660
         End
      End
   End
   Begin VB.PictureBox PicTab 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   4515
      ScaleHeight     =   2535
      ScaleWidth      =   3990
      TabIndex        =   4
      Top             =   2085
      Width           =   3990
      Begin XtremeSuiteControls.TabControl TabCtlWindow 
         Height          =   2280
         Left            =   90
         TabIndex        =   5
         Top             =   105
         Width           =   3765
         _Version        =   589884
         _ExtentX        =   6641
         _ExtentY        =   4022
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox PicImage 
      BorderStyle     =   0  'None
      Height          =   2595
      Left            =   9570
      ScaleHeight     =   2595
      ScaleWidth      =   1935
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      Begin VB.VScrollBar VScroll 
         Height          =   1245
         Left            =   1620
         Max             =   0
         TabIndex        =   2
         Top             =   150
         Width           =   225
      End
      Begin C1Chart2D8.Chart2D ChartThis 
         Height          =   735
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   120
         Width           =   885
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1561
         _ExtentY        =   1296
         _StockProps     =   0
         ControlProperties=   "frmLISView.frx":685E
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7020
      Width           =   15540
      _ExtentX        =   27411
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLISView.frx":6DE1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24500
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
   Begin MSComctlLib.ImageList Imglist 
      Left            =   135
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":7675
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":7C0F
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":81A9
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":8743
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":8CDD
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9277
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9611
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":99AB
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLISView.frx":9D45
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrThis 
      Left            =   1020
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmLISView.frx":A0DF
      Left            =   2565
      Top             =   180
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLisView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000

Private mlng����ID As Long
Private mlng��ҳID As Long

Private mlngҽ��ID As Long
Private mlng�걾ID As Long
Private mlng������� As Long

Private mstrPrivs As String
Private mstrLike As String

Private mfrmLisRptGeneral   As frmLisRptGeneral                  '����鿴
Attribute mfrmLisRptGeneral.VB_VarHelpID = -1
Private mfrmLisRptMicrobiology As frmLisRptMicrobiology                ' ΢���ﱨ��鿴

Private Const ID_MENU_MOUSE = 90

Private Const Dkp_ID_Request As Integer = 3                         '�˶ԵǼǴ���
Private Const Dkp_ID_Append As Integer = 4                          '���渽�Ӵ���
Private Const Dkp_ID_Image As Integer = 5                           '��ʾ����ͼ��

Private Enum mCol
    ID = 0: ����: ����: ����ʱ��: ������Ŀ: ��Դ: סԺ����: �걾��: ΢����걾: ҽ��ID: ������� ': ����: �Ա�: ����: ����id:  ������: �����: Ӥ��: ��ҳID: ���ʱ��: ��λ
End Enum
Dim blnLoad As Boolean
Private mlngItemID As Long      '�ϴ�ѡ����
Private mstrWhere As String     '��������

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mlngMod As Integer '����ģ���
Private mblnShowBorder As Boolean       '�Ƿ���ʾ�����border
Private mlngPageId As Long                'סԺ������ҳ

Private Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Sub ShowMe(ByVal lng����ID As Long, ByVal lngMod As Long, ByVal frmMain As Form, Optional ByVal blnShowBorder As Boolean = True, _
                Optional ByRef objOutFrm As Object)
    On Error GoTo errHandle
    
    mblnShowBorder = blnShowBorder
    mlng����ID = 0
    If lng����ID = 0 Then Exit Sub
    mlng����ID = lng����ID
    mlngMod = lngMod
    If blnShowBorder Then
        Me.Show , frmMain  '�������ʾ����ı߿����ʾ�ô���ΪǶ��ʽ���ã����ǵ���show����
    Else
        Call YSystemMenu(Me.hWnd)
    End If
    Set objOutFrm = Me

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub CreateCbs()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbrThis.Icons = zlCommFun.GetPubIcons
    With Me.cbrThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrThis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrThis.ActiveMenuBar.Title = "�˵�"
'    Me.cbrthis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&T)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")

       ' Set cbrControl = .Add(xtpControlButton, conMenu_File_Parameter, "��������(&O)"): cbrControl.BeginGroup = True

        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With
    

    'conMenu_EditPopup
'    '�Ҽ��˵�
    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_MENU_MOUSE, "�Ҽ��˵�", -1, False)
    cbrMenuBar.ID = ID_MENU_MOUSE
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "����Ԥ��(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "�����ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Transfer_Force, "�����ѯ(&P)"): cbrControl.BeginGroup = True

    End With
    cbrMenuBar.Visible = False

    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)")
        With cbrControl.CommandBar.Controls
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): cbrPopControl.BeginGroup = True
            Set cbrPopControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Location, "��ʾ�����(&S)"): cbrPopControl.BeginGroup = True
    
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
            cbrPopControl.Checked = True
            Set cbrPopControl = cbrControl.CommandBar.Controls.Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
            cbrPopControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)"): cbrControl.Checked = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Backward, "ǰһ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Forward, "��һ��(&L)")


        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_LeaveMedi, "���ؼ���ͼ��"): 'cbrControl.BeginGroup = True
        
        If zlDatabase.GetPara("���ؼ���ͼ��", glngSys, mlngMod, "True") = "True" Then
            cbrControl.Checked = True
        End If

        Set cbrControl = .Add(xtpControlButton, conMenu_LIS_HideList, "�����б�(&P)"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&F)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbrThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With

    '�����
    
    With Me.cbrThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_ESCAPE, conMenu_LIS_Cancel
        .Add 0, VK_PAGEUP, conMenu_Tool_Reference_1
        .Add 0, VK_PAGEDOWN, conMenu_Tool_Reference_2

    End With
    Me.cbrThis.ActiveMenuBar.Visible = mblnShowBorder
    '���ò����ò˵�
'    With Me.cbrthis.Options
'        .AddHiddenCommand conMenu_File_PrintSet
'    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbrThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next

End Sub

Private Sub CreateDockPane()
    Dim Pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Dim lngPane5Width As Long, lngPane2Height As Long, lngPane2Width As Long, lngPane3Height As Long
    

    dkpMain.Options.HideClient = True
    
    Set Pane3 = dkpMain.CreatePane(Dkp_ID_Request, 100, 600, DockLeftOf)
    Pane3.Title = "���յǼ�"
    Pane3.Handle = Me.PicInfo.hWnd
    Pane3.Options = PaneNoCaption
    
    
    Set Pane4 = dkpMain.CreatePane(Dkp_ID_Append, 9800, 790, DockRightOf, Pane3)
    Pane4.Title = "���Ӵ���"
    Pane4.Handle = Me.PicTab.hWnd
    Pane4.Options = PaneNoCaption
    
    lngPane5Width = 200
    Set Pane5 = dkpMain.CreatePane(Dkp_ID_Image, lngPane5Width, 200, DockRightOf, Pane4)
    Pane5.Title = "ͼ����ʾ"
    Pane5.Handle = Me.PicImage.hWnd
'    Pane5.Options = PaneNoCaption
    Pane4.Select
    
End Sub

Private Sub CreateTableControl()
    
    On Error Resume Next

    With Me.TabCtlWindow
        Set .Icons = zlCommFun.GetPubIcons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem(0, "���鱨��", mfrmLisRptGeneral.hWnd, conMenu_Tool_Report).Tag = "��ͨ������"
        .InsertItem(1, "���鱨��", mfrmLisRptMicrobiology.hWnd, conMenu_Tool_Report).Tag = "΢���ﱨ����"
        
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True
        .Item(0).Selected = True
        DoEvents
        .Item(1).Visible = False
    End With

End Sub

Private Sub CreateRptListHead()
    Dim Column As ReportColumn
    Dim i As Integer

    With Me.rptList.Columns


        rptList.SetImageList Imglist

        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False
        
        Set Column = .Add(mCol.����, "��������", 90, True): Column.Groupable = True
        
        Set Column = .Add(mCol.����, "", 18, False): Column.Icon = 0

        Set Column = .Add(mCol.������Ŀ, "������Ŀ", 90, True): Column.Groupable = True

        Set Column = .Add(mCol.����ʱ��, "����ʱ��", 80, True): Column.Groupable = True
        Column.Sortable = True: Column.SortAscending = False: Me.rptList.SortOrder.Add Column
        Set Column = .Add(mCol.��Դ, "��Դ", 30, True): Column.Groupable = True
        Set Column = .Add(mCol.סԺ����, "סԺ����", 65, True): Column.Groupable = False
        
        Set Column = .Add(mCol.�걾��, "�걾��", 65, True): Column.Groupable = False
        Set Column = .Add(mCol.΢����걾, "΢����걾", 30, True): Column.Visible = False: Column.Groupable = True
        Set Column = .Add(mCol.ҽ��ID, "ҽ��id", 30, True): Column.Visible = False: Column.Groupable = False
        Set Column = .Add(mCol.�������, "�������", 30, True): Column.Visible = False: Column.Groupable = False

    End With
    
    With rptList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
        
    End With
        
    '�������
    Me.rptList.GroupsOrder.DeleteAll
    Me.rptList.GroupsOrder.Add Me.rptList.Columns.Find(mCol.����)
    Me.rptList.GroupsOrder(0).SortAscending = True
    Me.rptList.Columns.Find(mCol.����).Visible = False
    Me.rptList.Populate
End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnReset As Boolean = False)
    '����           �Լ���ͼ������Ű�
    '����           intCount = ͼ����
    '               blnReset = �Ƿ���Ҫ���¶���
    Dim intLoop As Integer
    Dim Pane5 As Pane

    On Error Resume Next
    
    For intLoop = 0 To intCount
        If intLoop = 0 Then
            With Me.ChartThis(intLoop)
                .Visible = True
                .Top = 0
                .Left = 0
                .Width = IIF(Me.PicImage.ScaleWidth - Me.VScroll.Width - 20 <= 100, 100, Me.PicImage.ScaleWidth - Me.VScroll.Width - 20)
                .Height = .Width
            End With
        Else
            If blnReset = True And Me.ChartThis.UBound < intLoop Then
                Load Me.ChartThis(intLoop)
            End If
            With Me.ChartThis(intLoop)
                .Visible = True
                .Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
                .Left = 0
                .Width = Me.ChartThis(intLoop - 1).Width
                .Height = .Width
                .IsBatched = False
            End With
        End If
    Next
    
    '���ض����Chart�ؼ�
    For intLoop = intCount + 1 To Me.ChartThis.UBound
        Me.ChartThis(intLoop).Visible = False
    Next
    
    Set Pane5 = Me.dkpMain.FindPane(Dkp_ID_Image)
    If Not Pane5 Is Nothing Then
'        If intCount < 0 Then
'            Pane5.Close
'        Else
            If Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = False Then

                Me.dkpMain.ShowPane (Dkp_ID_Image)
                Me.dkpMain.FindPane(Dkp_ID_Request).Select
                Me.dkpMain.RecalcLayout
            Else
                Pane5.Close
            End If
'        End If
    End If
    With Me.VScroll
        .Top = 0
        .Left = Me.PicImage.ScaleWidth - .Width - 10
        .Height = Me.PicImage.ScaleHeight
        .Max = intCount
        .SmallChange = 1
        .LargeChange = 1
    End With
End Sub

Private Function GetLastPageId(ByVal lngPatientID As String) As Long
    Dim strSQL As String
    Dim rsTmp As Recordset
    
    GetLastPageId = 0
    
    strSQL = "Select Distinct a.����id, b.�����, a.סԺ��, a.��Ժ����, a.����, a.��ҳid," & _
            " a.��Ժ����, a.��Ժ���� From ������ҳ a, ������Ϣ b where a.����id=b.����id and a.����id=[1] order by ��ҳid"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���鼼ʦվ", lngPatientID)
    
    With Me.cboPages
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem "�� " & rsTmp("��ҳID") & " ��"
            .ItemData(.NewIndex) = rsTmp("����id")
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveLast
            GetLastPageId = rsTmp("��ҳID")
            .Text = "�� " & rsTmp("��ҳID") & " ��"
        Else
            .Visible = False
            lblPages.Visible = False
        End If
    End With
End Function

Private Sub LoadAllData(ByVal lng����ID As Long)
    '��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer
    Dim objRow As ReportRow
    Dim blnHave As Boolean
    Dim dateMin As Date, dateMax As Date, str��Ŀ As String
    Dim lngPageId As Long
    
    On Error GoTo errHandle
    
    lngPageId = mlngPageId
    
    dateMin = CDate(0)
    dateMax = CDate(0)
    
    Me.rptList.Records.DeleteAll
    If lngPageId > 0 Or (cboPages.Text = "����" And cboPages.Visible = True) Then
        rptList.Columns(mCol.סԺ����).Visible = True
    Else
        rptList.Columns(mCol.סԺ����).Visible = False
    End If
    
    If mstrWhere = "" Then
        strSQL = "Select A.ID, A.����, A.����ʱ��, A.������Ŀ, A.�걾���, Nvl(A.΢����걾,0) as ΢����걾, A.����, A.�Ա�, A.����, A.������ as �������, A.ҽ��id, A.����id, A.������, A.�����," & vbNewLine & _
                "       A.Ӥ��, a.��ҳid סԺ����, A.���ʱ��, decode(A.������Դ,1,'����',2,'סԺ','����') as ������Դ, A.��������" & vbNewLine & _
                "From ����걾��¼ A,������Ϣ B  " & vbNewLine & _
                "Where A.����id = [1] And A.����� is Not null And A.����id=B.����id " & IIF(lngPageId > 0, " And (a.��ҳid = [2] or a.��ҳid is null) ", "") & vbNewLine & _
                "Order By A.����ʱ��, A.���ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lngPageId)
    Else
        dateMin = CDate(Split(mstrWhere, "|")(0))
        dateMax = CDate(Split(mstrWhere, "|")(1))
        str��Ŀ = CStr(Split(mstrWhere, "|")(2))
        dateMax = DateAdd("d", 1, dateMax)
        
        strSQL = "Select A.ID, A.����, A.����ʱ��, A.������Ŀ, A.�걾���,  Nvl(A.΢����걾,0) as ΢����걾, A.����, A.�Ա�, A.����, A.������ as �������, A.ҽ��id, A.����id, A.������, A.�����," & vbNewLine & _
                "       A.Ӥ��, a.��ҳid סԺ����, A.���ʱ��, decode(A.������Դ,1,'����',2,'סԺ','����') as ������Դ, A.��������" & vbNewLine & _
                "From ����걾��¼  A ,������Ϣ B " & vbNewLine & _
                "Where A.����id = [1] And A.����id=B.����id And A.����� is Not null And A.����ʱ�� Between [2] And [3] " & vbNewLine & _
                IIF(str��Ŀ = "", "", " And instr(A.������Ŀ,[4])>0  ") & IIF(lngPageId > 0, " And (a.��ҳid = [5] or a.��ҳid is null) ", "") & _
                "Order By A.����ʱ��, A.���ʱ�� Desc"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, dateMin, dateMax, str��Ŀ, lngPageId)
    End If
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.Count + 1
            Record.AddItem ""
        Next
        Record.Item(mCol.����).value = "" & rsTmp("��������")
        If Val("" & rsTmp("����")) = 1 Then
            Record.Item(mCol.����).Icon = 1
        End If
        Record.Item(mCol.������Ŀ).value = Trim("" & rsTmp("������Ŀ"))
        Record.Item(mCol.����ʱ��).Caption = Format("" & rsTmp("����ʱ��"), "MM-dd HH:mm:ss")
        Record.Item(mCol.����ʱ��).value = Format("" & rsTmp("����ʱ��"), "YYYY-MM-dd HH:mm:ss")
        Record.Item(mCol.�걾��).value = Trim("" & rsTmp("�걾���"))
        Record.Item(mCol.΢����걾).value = Trim("" & rsTmp("΢����걾"))
        Record.Item(mCol.ҽ��ID).value = Val("" & rsTmp("ҽ��ID"))
        Record.Item(mCol.ID).value = Val("" & rsTmp!ID)
        Record.Item(mCol.�������).value = Val("" & rsTmp("�������"))
        Record.Item(mCol.��Դ).value = "" & rsTmp("������Դ")
        Record.Item(mCol.סԺ����).value = "" & rsTmp("סԺ����")
        If mstrWhere = "" And IsNull(rsTmp("����ʱ��")) = False Then
            If dateMin = CDate(0) Then
                dateMin = CDate(Format("" & rsTmp("����ʱ��"), "YYYY-MM-dd"))
                dateMax = CDate(Format("" & rsTmp("����ʱ��"), "YYYY-MM-dd"))
            Else
                If CDate(Format("" & rsTmp("����ʱ��"), "YYYY-MM-dd")) > dateMax Then
                    dateMax = CDate(Format("" & rsTmp("����ʱ��"), "YYYY-MM-dd"))
                End If
            End If
        End If
        blnHave = True
        rsTmp.MoveNext
    Loop
    
    If mstrWhere = "" Then
        dtpStart.MaxDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
        dtpEnd.MaxDate = dtpStart.MaxDate
        
        txt��Ŀ.Text = ""
        txt��Ŀ.Tag = ""
        strSQL = "select ����,�Ǽ�ʱ�� from ������Ϣ Where ����id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        Do Until rsTmp.EOF
            Me.Caption = "�������(" & rsTmp.Fields("����") & ")"
            dtpEnd.MinDate = rsTmp!�Ǽ�ʱ��
            dtpStart.MinDate = rsTmp!�Ǽ�ʱ��
            rsTmp.MoveNext
        Loop
        
        strSQL = "Select Min(��Ժ����) as ��Ժ���� From ������ҳ Where ����ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        
        If Not rsTmp.EOF Then
            If rsTmp!��Ժ���� < dtpEnd.MinDate Then
                dtpEnd.MinDate = rsTmp!��Ժ����
                dtpStart.MinDate = rsTmp!��Ժ����
            End If
        End If
        
        dtpStart.value = dtpStart.MinDate
        dtpEnd.value = dtpEnd.MaxDate
    End If
    
    '1-ˢ��
    rptList.Populate

    '2-�۵�������
    For Each objRow In rptList.Rows
        If objRow.GroupRow Then objRow.Expanded = False
    Next
    
    '3-��λ���ϴ�ѡ����
    If mlngItemID <> 0 Then
        For Each objRow In Me.rptList.Rows
            If objRow.GroupRow = False Then
                If Val(objRow.Record(mCol.ID).value) = mlngItemID Then
                    Set Me.rptList.FocusedRow = objRow
                    Exit For
                End If
            End If
        Next
    End If
    
    '4-չ��ѡ����
    If Me.rptList.FocusedRow Is Nothing And Me.rptList.Rows.Count > 0 Then
        If Me.rptList.Rows(0).GroupRow Then
            Me.rptList.Rows(0).Expanded = True
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0).Childs(0)
        Else
            Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
        End If
    End If
    '5-�����¼�
    Call rptList_SelectionChanged
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function ReadImageData(lngKeyID As Long, blnSave As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim strTime As Date
    Dim objLisDev As Object, strFilename As String, strErr As String
    
    On Error GoTo errH
    strTime = Now
    gstrSQL = "select id ,�걾ID,ͼ������ from ����ͼ���� where �걾id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKeyID)
    'ͼ���Ű�
    ImageTypeSet rsTmp.RecordCount - 1, True
    '����ʾʱ������
    If Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked = True Then Exit Function
    
    Set objLisDev = CreateObject("zlLISDev.clsDrawGraph")
    
    If objLisDev.GetSampleImgInit(glngSys, gcnOracle, strErr) Then
        Do Until rsTmp.EOF
            If Dir(App.Path & "\" & rsTmp("ID") & ".cht") = "" Then
                If Not objLisDev Is Nothing Then
                    strFilename = ""
                    
                    strFilename = objLisDev.GetImage(Val("" & rsTmp("ID")), App.Path, False, strErr)
                    If strFilename <> "" Then
                        Me.ChartThis(DrawIndex).Load App.Path & "\" & strFilename
                    Else
                        '��ȡ�ļ� ʧ��
                         
                    End If
                Else
                    '��������ʧ��!
                End If
            Else
                Me.ChartThis(DrawIndex).Load App.Path & "\" & rsTmp("ID") & ".cht"
                 
            End If
            DrawIndex = DrawIndex + 1
            rsTmp.MoveNext
        Loop
        Call objLisDev.GetSampleImgExit(strErr)
    End If
    ImageTypeSet DrawIndex - 1, False

    ReadImageData = True
'    Debug.Print "ID=" & lngKeyID & ",��ʱ:" & DateDiff("s", strTime, Now)
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintSetup()
    '��ӡ����
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    lngҽ��ID = mlngҽ��ID
    lng����ID = mlng����ID
    
    strSQL = "select ���ͺ� from ����ҽ������ a , ����ҽ����¼ b where b.id = a.ҽ��id and b.id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, gstrSysName, lngҽ��ID)
    If rsTmp.EOF = False Then
        lng���ͺ� = Val("" & rsTmp(0))
    End If
    
    If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        ReportPrintSet gcnOracle, glngSys, strReportCode, Me
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetReportCode(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    If lngҽ��ID = 0 And lng���ͺ� = 0 Then Exit Function
    
    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.���,'00000'))||'-2' AS ������," & _
                       "A.NO," & _
                       "A.��¼���� " & _
                "FROM ����ҽ������ A,�����ļ��б� C,����ҽ����¼ D,��������Ӧ�� E " & _
                "Where E.�����ļ�id = C.ID " & _
                        "AND D.������ĿID=E.������ĿID " & _
                      "AND A.ҽ��ID=D.ID AND E.Ӧ�ó���=Decode(D.������Դ,2,2,4,4,1) " & _
                      " AND D.���id= [1] "
    If DataMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    On Error GoTo errH
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lngҽ��ID, lng���ͺ�)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.NVL(rs("������"))
        strNO = zlCommFun.NVL(rs("NO"))
        bytMode = zlCommFun.NVL(rs("��¼����"), 1)
    End If
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReportPrint(ByVal blnPrint As Boolean)
    '���������ӡ
    
    Dim strReportCode As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng����ID As Long
    Dim strSQL As String

    Dim intLoop As Integer
    On Error GoTo errH
    If Me.rptList.FocusedRow Is Nothing Then Exit Sub
    
    'blnCurrMoved = rptList.SelectedRows(0).Record.Item(mCol.ת��).Value = "��"
    Call Open_LIS_Report(Me, mlngҽ��ID, mlng����ID, blnCurrMoved, blnPrint)

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub ShowOrHideItem(Control As CommandBarControl, DkpID As Integer)
    '����               '��ʾ������
    Dim Pane As Pane
    Set Pane = Me.dkpMain.FindPane(DkpID)
    If Control.Checked = True Then
        Pane.Close
    Else
        
        If Pane.Closed Then Me.dkpMain.ShowPane (DkpID)
        Pane.Select
    End If
    If DkpID = Dkp_ID_Image Then ReadImageData mlng�걾ID, False
    Me.dkpMain.RecalcLayout
    Me.cbrThis.RecalcLayout
End Sub

Private Sub BackOrNextPatient(Move As Integer)
    '����                 �ƶ�����һ�����˻���һ������
    '����                 Move = 1 ��һ���� =2 ��һ����
    Dim Rerow As ReportRow
    Dim i As Long
    With Me.rptList
        If .Rows.Count <= 0 Then Exit Sub
        i = .SelectedRows(0).Index
        If Move = 1 Then            '�����ƶ�
            If i - 1 >= 0 Then
                i = i - 1
                .FocusedRow = .Rows(i)
            End If
        Else
            If i < .Rows.Count - 1 Then
                i = i + 1
                .FocusedRow = .Rows(i)
            End If
        End If
    End With
End Sub

Private Sub cboPages_Click()
    mlngPageId = Val(Trim(Replace(Replace(cboPages.Text, "��", ""), "��", "")))
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    
    Select Case Control.ID
        
        '''''''''''''''''''''''''''''''''''''''�ļ�''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_File_PrintSet                                                      '��ӡ����
             PrintSetup
            
        Case conMenu_File_Preview                                                       '����Ԥ��
            ReportPrint False
        
        Case conMenu_File_Print                                                         '�����ӡ
            ReportPrint True
        Case conMenu_File_Exit                                                          '�˳�
            Unload Me
        Case conMenu_View_Refresh
            Call zlRefresh
        '''''''''''''''''''''''''''''''''''''''�鿴'''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_View_ToolBar_Button                                                '��׼��ť
            Control.Checked = Not Control.Checked
            Me.cbrThis(2).Visible = Control.Checked
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Text                                                  '�ı���ǩ
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbrThis(2).Controls
                cbrControl.Style = IIF(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_ToolBar_Size                                                  '��ͼ��
            Control.Checked = Not Control.Checked
            Me.cbrThis.Options.LargeIcons = Not Me.cbrThis.Options.LargeIcons
            Me.cbrThis.RecalcLayout
        
        Case conMenu_View_StatusBar                                                     '״̬��
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Control.Checked
            Me.cbrThis.RecalcLayout
'''
        Case conMenu_View_Expend_CurCollapse                            '�۵���ǰ��
            If rptList.SelectedRows.Count > 0 Then
                If rptList.SelectedRows(0).GroupRow Then
                    rptList.SelectedRows(0).Expanded = False
                ElseIf Not rptList.SelectedRows(0).ParentRow Is Nothing Then
                    If rptList.SelectedRows(0).ParentRow.GroupRow Then
                        rptList.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '���۵���λ��������,�����Զ�������¼�
            Call rptList_SelectionChanged
    
        Case conMenu_View_Expend_CurExpend                              'չ����ǰ��
            If rptList.SelectedRows.Count > 0 Then
                rptList.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse                            '�۵�������
            For Each objRow In rptList.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '���۵���λ��������,�����Զ�������¼�
            Call rptList_SelectionChanged
        Case conMenu_View_Expend_AllExpend                              'չ��������
            For Each objRow In rptList.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Location
            rptList.ShowGroupBox = Not rptList.ShowGroupBox             '��ʾ�����

'''
        Case conMenu_View_Forward                                                       'ǰһ��
            BackOrNextPatient 2
        
        Case conMenu_View_Backward                                                      '��һ��
            BackOrNextPatient 1
            
        Case conMenu_LIS_HideList                                                       '�����б�
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Request
        
        Case conMenu_Manage_LeaveMedi                                                   '���ؼ���ͼ��
            Control.Checked = Not Control.Checked
            ShowOrHideItem Control, Dkp_ID_Image
        ''''''''''''''''''''''''''''''''''''''����''''''''''''''''''''''''''''''''''''''''''''''''''
        Case conMenu_Help_Help                                                          '��������
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
        Case conMenu_Help_Web                                                           'WEB�ϵ�
            Call zlHomePage(hWnd)
        
        Case conMenu_Help_Web_Home                                                      '��ҳ
            Call zlHomePage(Me.hWnd)
        
        Case conMenu_Help_Web_Mail                                                      '���ͷ���
            Call zlMailTo(Me.hWnd)
        
        Case conMenu_Help_About                                                         '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
    End Select
End Sub

Private Sub cbrThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    
    Case conMenu_File_Print, conMenu_File_Preview  '�����ӡ,Ԥ����Ԥ����Ҳ�ɴ�ӡ��
        Control.Enabled = InStr(mstrPrivs, "�����ӡ") > 0
    Case conMenu_File_Exit
        Control.Visible = mblnShowBorder
    End Select
End Sub

Private Sub cmdOK_Click()
    Call zlRefresh
End Sub

Private Sub cmd��Ŀ_Click()
    Call ShowSelect
End Sub

Private Sub dkpMain_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionClosed Then Cancel = True
    If Pane.ID = Dkp_ID_Append Then Cancel = False

End Sub

Private Sub dkpMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim lngTop As Long, lngLeft As Long, lngRight As Long, lngBottom As Long
    
    Me.cbrThis.GetClientRect lngLeft, lngTop, lngRight, lngBottom
    Top = lngTop
    Bottom = Me.ScaleHeight - lngBottom
End Sub

Private Sub dkpMain_Resize()
    Me.cbrThis.RecalcLayout
    
    Call ImageTypeSet(Me.VScroll.Max)
End Sub



Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Dkp_ID_Request
        Item.Handle = Me.PicInfo.hWnd
    Case Dkp_ID_Append
        Item.Handle = Me.PicTab.hWnd
    Case Dkp_ID_Image
        Item.Handle = Me.PicImage.hWnd
    End Select
End Sub

Private Sub cbrthis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub


Private Sub Form_Activate()
'   Call rptList_SelectionChanged '����ѡ���¼�
End Sub

Private Sub Form_Load()

    On Error Resume Next
    mstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    

    
    If Dir(App.Path & "\zlLisPic*.Bmp") <> "" Then
        Kill App.Path & "\zlLisPic*.Bmp"
    End If
    If Dir(App.Path & "\*.cht") <> "" Then Kill App.Path & "\*.cht"
    '=====================================================
    'Call RestoreWinState(Me, App.ProductName)                   '����ָ�

    Set mfrmLisRptGeneral = frmLisRptGeneral                     '��ͨ�걾����
    Set mfrmLisRptMicrobiology = frmLisRptMicrobiology           '΢����걾����
    mfrmLisRptGeneral.mlngMod = mlngMod
    mfrmLisRptMicrobiology.mlngMode = mlngMod
    mstrPrivs = IIF(Right(gMainPrivs, 1) = ";", gMainPrivs, gMainPrivs & ";") & gcolPrivs(glngSys & "_" & mlngMod)
    CreateCbs                           '����������
    CreateDockPane                      '������������
    CreateTableControl                  '����TAB
    CreateRptListHead
    mstrWhere = ""
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    End If
    
    mlngPageId = GetLastPageId(mlng����ID)
    Call LoadAllData(mlng����ID)
    
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    Dim Pane1 As Pane
    Dim intLoop As Integer
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub

    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Append)
    Pane1.MinTrackSize.SetSize 6954 / Screen.TwipsPerPixelX, 380 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize Pane1.MaxTrackSize.Width, 380 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Request)
    Pane1.MinTrackSize.SetSize 3080 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    Pane1.MaxTrackSize.SetSize 3980 / Screen.TwipsPerPixelX, 2295 / Screen.TwipsPerPixelY
    
    Set Pane1 = Me.dkpMain.FindPane(Dkp_ID_Image)
    Pane1.MinTrackSize.SetSize 1880 / Screen.TwipsPerPixelX, 500 / Screen.TwipsPerPixelY
    
    Me.dkpMain.RecalcLayout
    Me.dkpMain.NormalizeSplitters
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnCheck As Boolean
    blnCheck = Me.cbrThis.FindControl(, conMenu_Manage_LeaveMedi, , True).Checked
    
    Call zlDatabase.SetPara("���ؼ���ͼ��", IIF(blnCheck, "True", "False"), glngSys, mlngMod)
    Call SaveWinState(Me, App.ProductName)
    Unload mfrmLisRptGeneral
    Unload mfrmLisRptMicrobiology
    Set mfrmLisRptGeneral = Nothing
    Set mfrmLisRptMicrobiology = Nothing
    
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    End If
    mlng����ID = 0
    mlng��ҳID = 0
    mlngҽ��ID = 0
    mlng�걾ID = 0
    mlng������� = 0
End Sub


Private Sub picFind_Resize()
    On Error Resume Next
    dtpStart.Width = (picFind.ScaleWidth - dtpStart.Left - 90) / 2
    dtpEnd.Left = dtpStart.Left + dtpStart.Width + 45
    dtpEnd.Width = picFind.ScaleWidth - dtpEnd.Left - 45
    
    cmdOK.Left = picFind.ScaleWidth - cmdOK.Width - 45
    
    cmd��Ŀ.Left = cmdOK.Left - 45 - cmd��Ŀ.Width
    txt��Ŀ.Width = cmd��Ŀ.Left - txt��Ŀ.Left - 10
    
End Sub

Private Sub PicInfo_Resize()
    On Error Resume Next

    picFind.Top = 0
    picFind.Left = 0
    picFind.Width = PicInfo.ScaleWidth

    Me.rptList.Left = 0
    Me.rptList.Top = picFind.Top + picFind.Height

    Me.rptList.Width = PicInfo.ScaleWidth
    Me.rptList.Height = PicInfo.ScaleHeight - Me.rptList.Top
    

End Sub

Private Sub picList_Resize()

End Sub

Private Sub picTab_Resize()
    
    Me.TabCtlWindow.Top = 0
    Me.TabCtlWindow.Left = 0
    Me.TabCtlWindow.Width = Me.PicTab.ScaleWidth
    Me.TabCtlWindow.Height = Me.PicTab.ScaleHeight
    Call ImageTypeSet(VScroll.Max)
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objPopup As CommandBarPopup
    On Error Resume Next
    If Button = 2 Then
        If rptList.Records.Count <= 0 Then Exit Sub
        If Not rptList.SelectedRows(0).GroupRow Then
            Set objPopup = cbrThis.ActiveMenuBar.FindControl(, ID_MENU_MOUSE)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub rptList_SelectionChanged()
    On Error GoTo errHandle
    Dim i As Integer
     '---------------�ı�ѡ��ǰ��������м�¼----------------------
'    Call mfrmLisRptGeneral.zlRefresh(0)
'    Call mfrmLisRptMicrobiology.zlRefresh(0, 0)
'    Call ReadImageData(0, False)
    '-------------------------------------------------------------
    If rptList.SelectedRows.Count = 0 Then
        If rptList.Rows.Count > 0 Then
            '�м�¼,ȡ�ڸ��Ƿ�����,����ǰ��
            For i = 0 To rptList.Rows.Count - 1
                If Not rptList.Rows(i).GroupRow Then
                    rptList.Rows(i).Selected = True
                    
                    mlngҽ��ID = Val(rptList.Rows(i).Record(mCol.ҽ��ID).value)
                    mlng�걾ID = Val(rptList.Rows(i).Record(mCol.ID).value)
                    mlng������� = Val(rptList.Rows(i).Record(mCol.�������).value)
                    
                    If rptList.Rows(i).Record(mCol.΢����걾).value = "0" Then
                        Me.TabCtlWindow.Item(0).Visible = True
                        Me.TabCtlWindow.Item(1).Visible = False
                        Me.TabCtlWindow.Item(0).Selected = True
                        mfrmLisRptGeneral.zlRefresh (mlngҽ��ID)
                        ReadImageData mlng�걾ID, False
                    Else
                        Me.TabCtlWindow.Item(0).Visible = False
                        Me.TabCtlWindow.Item(1).Visible = True
                        Me.TabCtlWindow.Item(1).Selected = True
                        Call mfrmLisRptMicrobiology.zlRefresh(mlng�걾ID, mlng�������)
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If rptList.FocusedRow Is Nothing Then
        mlngҽ��ID = 0
        mlng�걾ID = 0
        mlng������� = 0
        Exit Sub
    End If
    If rptList.FocusedRow.GroupRow Then Exit Sub
    
    If Val(rptList.FocusedRow.Record(mCol.ҽ��ID).value) <> 0 And _
      (mlngҽ��ID <> Val(rptList.FocusedRow.Record(mCol.ҽ��ID).value) Or _
      mlng�걾ID <> Val(rptList.FocusedRow.Record(mCol.ID).value)) Then
        mlngҽ��ID = Val(rptList.FocusedRow.Record(mCol.ҽ��ID).value)
        mlng�걾ID = Val(rptList.FocusedRow.Record(mCol.ID).value)
        mlng������� = Val(rptList.FocusedRow.Record(mCol.�������).value)
        If rptList.FocusedRow.Record(mCol.΢����걾).value = "0" Then
            Me.TabCtlWindow.Item(0).Visible = True
            Me.TabCtlWindow.Item(1).Visible = False
            Me.TabCtlWindow.Item(0).Selected = True
            mfrmLisRptGeneral.zlRefresh (mlngҽ��ID)
            
        Else
            Me.TabCtlWindow.Item(0).Visible = False
            Me.TabCtlWindow.Item(1).Visible = True
            Me.TabCtlWindow.Item(1).Selected = True
            Call mfrmLisRptMicrobiology.zlRefresh(mlng�걾ID, mlng�������)
        End If
    Else
        mlngҽ��ID = 0
        mfrmLisRptGeneral.zlRefresh (0)
        Call mfrmLisRptMicrobiology.zlRefresh(0, 0)
    End If
    ReadImageData mlng�걾ID, False
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
    Resume
    End If
End Sub

Private Sub zlRefresh()
    '��ˢ��ʱ����
    If dtpStart.value > dtpEnd.value Then
        MsgBox "��ѯ��ʼ���ڲ��ܴ��ڽ������ڣ�"
        Exit Sub
    End If
    mstrWhere = Format(dtpStart.value, "yyyy-MM-dd") & "|" & Format(dtpEnd.value, "yyyy-MM-dd") & "|" & Trim(txt��Ŀ.Text)
    Call LoadAllData(mlng����ID)
End Sub

Private Function ShowSelect()

    Dim vRect As RECT, strSQL As String, rsTmp As ADODB.Recordset
    Dim str���� As String, blnCanel As Boolean, strSel��Ŀ As String
    Dim strInput As String
    
    On Error GoTo errHandle
    str���� = Trim(txt��Ŀ.Text)
    If Trim(str����) <> "" Then
        str���� = Replace(str����, "%", "")
        str���� = Replace(str����, "'", "")
        str���� = Replace(UCase(str����), "AND", "")
        str���� = Replace(UCase(str����), "OR", "")
    Else
        str���� = ""
    End If
    vRect = GetControlRect(txt��Ŀ.hWnd)

    
    If str���� <> "" Then
        strInput = " And (B.���� Like [1] Or A.���� Like [1] Or A.���� Like [1])"
        If IsNumeric(str����) Then
            '1X.����ȫ������ʱֻƥ�����
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.���� Like [1]"
        ElseIf zlCommFun.IsCharAlpha(str����) Then
            'X1.����ȫ����ĸʱֻƥ�����
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.���� Like [1]"
        ElseIf zlCommFun.IsCharChinese(str����) Then
            '��������,��ֻƥ������
            strInput = " And A.���� Like [1]"
        End If
        
        str���� = IIF(Len(str����) < 3, "", mstrLike) & str���� & "%"
        strSQL = "Select Distinct A.ID, A.�������� As ����, A.����, A.����, Decode(A.�����Ŀ, 1, '��', '��') As �����Ŀ" & vbNewLine & _
                "From ������ĿĿ¼ A, ������Ŀ���� B" & vbNewLine & _
                "Where A.ID = B.������Ŀid And A.��� = 'C' And A.����Ӧ�� = 1 And" & vbNewLine & _
                "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & strInput
    
    Else
        strSQL = "Select A.ID, A.�������� As ����, A.����, A.����, Decode(A.�����Ŀ, 1, '��', '��') As �����Ŀ" & vbNewLine & _
                "From ������ĿĿ¼ A" & vbNewLine & _
                "Where A.��� = 'C' And A.����Ӧ�� = 1 And" & vbNewLine & _
                "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) "
    End If
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��Ŀ", False, "", "ѡ����Ŀ", False, False, True, _
                                         vRect.Left, vRect.Top, txt��Ŀ.Height, blnCanel, False, True, str����)
    
    If Not blnCanel And Not rsTmp Is Nothing Then
        Do Until rsTmp.EOF
            strSel��Ŀ = strSel��Ŀ & "," & rsTmp!����
            rsTmp.MoveNext
        Loop
        txt��Ŀ.Text = ""
        If strSel��Ŀ <> "" Then txt��Ŀ.Text = Mid(strSel��Ŀ, 2)
    Else
        If txt��Ŀ.Enabled Then
            txt��Ŀ.SelStart = 0: txt��Ŀ.SelLength = Len(txt��Ŀ.Text)
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub txt��Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ShowSelect
End Sub

Private Sub VScroll_Change()
    Dim intLoop As Integer
    If Me.Visible = False Then Exit Sub
    For intLoop = 0 To Me.VScroll.Max
        If intLoop < Me.VScroll.value Then
            Me.ChartThis(intLoop).Visible = False
        Else
            Me.ChartThis(intLoop).Visible = True
            If intLoop = Me.VScroll.value Then
                Me.ChartThis(intLoop).Top = 0
            Else
                Me.ChartThis(intLoop).Top = Me.ChartThis(intLoop - 1).Top + Me.ChartThis(intLoop - 1).Height + 10
            End If
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/5/25
'��    ��:����API��̬���ô����border
'��    ��:
'           new_Hwnd    ����ľ��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub
