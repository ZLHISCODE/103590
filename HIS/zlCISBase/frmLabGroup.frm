VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmLabGroup 
   Caption         =   "����С������"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   13050
   Icon            =   "frmLabGroup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   13050
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox pic���� 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5640
      Left            =   4215
      ScaleHeight     =   5640
      ScaleWidth      =   4095
      TabIndex        =   12
      Top             =   1245
      Width           =   4095
      Begin VB.PictureBox picEdit���� 
         BorderStyle     =   0  'None
         Height          =   2715
         Left            =   -45
         ScaleHeight     =   2715
         ScaleWidth      =   3855
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3270
         Width           =   3855
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   350
            Index           =   1
            Left            =   2160
            TabIndex        =   15
            Top             =   105
            Width           =   1635
         End
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   350
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   75
            Width           =   1740
         End
         Begin MSComctlLib.ListView lvw���� 
            Height          =   1980
            Left            =   195
            TabIndex        =   16
            Top             =   660
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   3493
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblEdit���� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   30
            Left            =   165
            TabIndex        =   17
            Top             =   540
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg���� 
         Height          =   3195
         Left            =   15
         TabIndex        =   18
         Top             =   60
         Width           =   3825
         _cx             =   6747
         _cy             =   5636
         Appearance      =   0
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
         BackColorFixed  =   15790320
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
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
   Begin VB.PictureBox PicList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6525
      Left            =   15
      ScaleHeight     =   6525
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   825
      Width           =   3945
      Begin XtremeReportControl.ReportControl rptList 
         Height          =   6210
         Left            =   105
         TabIndex        =   1
         Top             =   165
         Width           =   3720
         _Version        =   589884
         _ExtentX        =   6562
         _ExtentY        =   10954
         _StockProps     =   0
         FocusSubItems   =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   7575
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLabGroup.frx":000C
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17939
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.PictureBox pic��Ա 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   8550
      ScaleHeight     =   5895
      ScaleWidth      =   4080
      TabIndex        =   3
      Top             =   1125
      Width           =   4080
      Begin VB.PictureBox picEdit��Ա 
         BorderStyle     =   0  'None
         Height          =   3675
         Left            =   90
         ScaleHeight     =   3675
         ScaleWidth      =   3855
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3390
         Width           =   3855
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1350
            TabIndex        =   8
            Top             =   240
            Width           =   2115
         End
         Begin VB.CommandButton cmdFind 
            Height          =   300
            Left            =   3480
            Picture         =   "frmLabGroup.frx":08A0
            Style           =   1  'Graphical
            TabIndex        =   7
            TabStop         =   0   'False
            ToolTipText     =   "���ҷ�����������Ŀ"
            Top             =   225
            Width           =   360
         End
         Begin VB.CommandButton cmd��Ա 
            Caption         =   "��"
            Height          =   350
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   60
            Width           =   390
         End
         Begin VB.CommandButton cmd��Ա 
            Caption         =   "��"
            Height          =   350
            Index           =   1
            Left            =   465
            TabIndex        =   5
            Top             =   30
            Width           =   555
         End
         Begin MSComctlLib.ListView lvw��Ա 
            Height          =   1830
            Left            =   435
            TabIndex        =   9
            Top             =   600
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   3228
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label lblEdit��Ա 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������Ա:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   450
            TabIndex        =   10
            Top             =   360
            Width           =   810
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfg��Ա 
         Height          =   3105
         Left            =   15
         TabIndex        =   11
         Top             =   30
         Width           =   3705
         _cx             =   6535
         _cy             =   5477
         Appearance      =   0
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
         BackColorFixed  =   15790320
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
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   930
      Top             =   330
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmLabGroup.frx":0E2A
      Left            =   2520
      Top             =   195
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmLabGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conPane_List = 201
Const conPane_Edit = 202
Const conPane_���� = 203
Const conPane_��Ա = 204
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mLngEditWidth As Long       'Ϊ��Ӧ����������´�����.�ȶ��봰���С.
Private mintEditState As Integer    '��ǰ�༭״̬��0-�Ǳ༭״̬,1-�༭״̬
Private mlngGroupID As Long         '��ǰ����С��ID
Private mstr������Ա As String      '��ǰС�����Ա
Private mstr�������� As String      '��ǰС�������

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long

Private Enum mCol
    ID = 0:  С�����: С��
    ����Id = 0: ��������: ����: �鿴: ����: ֻ����ɨ��������
    ��ԱID = 0: ��Ա����: ����: Ĭ��: ��ע
End Enum
Private mblnInit As Boolean
Private mstr�������� As String

Private Sub initMenu()
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set cbsThis.Icons = zlcommfun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��(&D)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "��������(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "������Ա(&M)")
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        
        
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Untread

        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FSHIFT, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_Edit_Untread
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "�޸�")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "ɾ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Compend, "��Ա")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
        
End Sub

Private Sub initPic()
    '-----------------------------------------------------
    '����ͣ������
    Dim panThis As Pane, panSub1 As Pane, panSub2 As Pane
    
    mblnInit = True
    Set panThis = dkpMan.CreatePane(conPane_List, 200, 580, DockLeftOf, Nothing)
    panThis.Title = "����С���б�"
    panThis.Options = PaneNoCaption
    
'    Set panThis = dkpMan.CreatePane(conPane_Edit, 350, 50, DockRightOf, Nothing)
'    panThis.Title = "С���������"
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption

    Set panSub1 = dkpMan.CreatePane(conPane_����, 550, 600, DockRightOf, panThis)
    panSub1.Title = "С������"
    panSub1.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable

    Set panSub2 = dkpMan.CreatePane(conPane_��Ա, 550, 600, DockRightOf, panThis)
    panSub2.Title = "С���Ա"
    panSub2.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panSub2.AttachTo panSub1
    
    panSub2.Select

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    Me.dkpMan.VisualTheme = ThemeOffice2003
End Sub

Private Function zlEditStart(ByVal intAdd As Integer) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    
    Select Case mintEditState
    Case 1  '��������
        If mlngGroupID = 0 And intAdd <> 1 Then Exit Function
        If frmLabGroupEdit.ShowMe(mlngGroupID, intAdd, mstr��������, Me) Then
            Call zlSaveData
        End If
        mintEditState = 0

    Case 2  '����
        If mlngGroupID = 0 Then Exit Function
        strSQL = "Select ID,����,���� From ��������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        Me.lvw����.ListItems.Clear
        Do Until rsTmp.EOF
            If InStr(mstr�������� & ",", "," & rsTmp.Fields("id") & ",") <= 0 Then
                Me.lvw����.ListItems.Add , "_" & rsTmp.Fields("id"), "" & rsTmp.Fields("����") & " " & rsTmp.Fields("����")
            End If
            rsTmp.MoveNext
        Loop
        Call pic����_Resize
        
        zlEditStart = True
    Case 3  '��Ա
        If mlngGroupID = 0 Then Exit Function
        Me.lvw��Ա.ListItems.Clear
        Call pic��Ա_Resize
        zlEditStart = True
    End Select
    Exit Function
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlEditCancle()
    mintEditState = 0: Me.PicList.Enabled = True
    Me.vfg����.Enabled = True
    Me.vfg��Ա.Enabled = True
    Call pic��Ա_Resize
    Call pic����_Resize
    Call zlLoadData
End Sub

Private Sub zlSaveData()
    '��������
    Dim strSQL As String, lngID As Long
    Dim str���� As String, str���� As String, intRow As Integer
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
    Select Case mintEditState
    Case 1  '��������

        
    Case 2  '����
        strSQL = ""
        With vfg����
            For intRow = .FixedRows To .Rows - 1
                strSQL = strSQL & "|" & .TextMatrix(intRow, mCol.����Id) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.�鿴) = flexChecked, 1, 0) & _
                "," & IIf(.Cell(flexcpChecked, intRow, mCol.����) = flexChecked, 1, 0) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.ֻ����ɨ��������) = flexChecked, 1, 0)
            Next
        End With
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "zl_����С������_Edit(" & mlngGroupID & ",'" & strSQL & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    Case 3  '��Ա
        strSQL = ""
        With vfg��Ա
            For intRow = .FixedRows To .Rows - 1
                strSQL = strSQL & "|" & .TextMatrix(intRow, mCol.��ԱID) & "," & IIf(.Cell(flexcpChecked, intRow, mCol.Ĭ��) = flexChecked, 1, 0)
            Next
        End With
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            strSQL = "zl_����С���Ա_Edit(" & mlngGroupID & ",'" & strSQL & "')"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End Select
    Call zlLoadData
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub zlEditDelete()
    '#
    Dim strSQL As String
    On Error GoTo ErrHandle
    
    strSQL = "zl_����С��_Edit(3," & mlngGroupID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Call zlLoadData
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub CreateRptListHead()
    Dim Column As ReportColumn
    Dim i As Integer

    With Me.rptList.Columns

        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        
        With rptList.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        
        End With
        'rptList.SetImageList Imglist

        Set Column = .Add(mCol.ID, "ID", 30, True): Column.Visible = False
        Set Column = .Add(mCol.С�����, "����", 100, True)
        Column.Sortable = True: Column.SortAscending = False: Me.rptList.SortOrder.Add Column
        
        Set Column = .Add(mCol.С��, "С��", 120, True)
    End With
End Sub

Private Sub zlLoadData()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim Record As ReportRecord
    Dim intLoop As Integer

    On Error GoTo ErrHandle
    rptList.Records.DeleteAll
    strSQL = "select id,����,���� from ����С��"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    Do Until rsTmp.EOF
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.count + 1
            Record.AddItem ""
        Next
        Record.Item(mCol.ID).Value = Val("" & rsTmp!ID)
        Record.Item(mCol.С�����).Value = Trim("" & rsTmp!����)
        Record.Item(mCol.С��).Value = Trim("" & rsTmp!����)
        rsTmp.MoveNext
    Loop
    rptList.Populate
    
    Dim rptParent As ReportRow
    If mlngGroupID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = mlngGroupID Then
                    Set rptParent = rptRow.ParentRow
                    Set Me.rptList.FocusedRow = rptRow
                    Exit For
                End If
            End If
        Next
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then
                If Not (rptRow Is rptParent) Then rptRow.Expanded = False
            End If
        Next
        Set Me.rptList.FocusedRow = Me.rptList.FocusedRow
    Else
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then rptRow.Expanded = False
        Next
    End If
    
    
    Call rptList_SelectionChanged
    
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
End Sub


Private Sub initVfg����()
    With vfg����
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = 6: .FixedCols = 0
        
        .TextMatrix(0, mCol.����Id) = "": .ColWidth(mCol.����Id) = 0: .ColAlignment(mCol.����Id) = flexAlignRightCenter
        .ColHidden(mCol.����Id) = True
        .TextMatrix(0, mCol.��������) = "����": .ColWidth(mCol.��������) = 1000: .ColAlignment(mCol.��������) = flexAlignLeftCenter
        .TextMatrix(0, mCol.����) = "����": .ColWidth(mCol.����) = 2000: .ColAlignment(mCol.����) = flexAlignLeftCenter
        .TextMatrix(0, mCol.�鿴) = "�鿴": .ColWidth(mCol.�鿴) = 600: .ColAlignment(mCol.�鿴) = flexAlignLeftCenter
        .TextMatrix(0, mCol.����) = "���� ": .ColWidth(mCol.����) = 600: .ColAlignment(mCol.����) = flexAlignLeftCenter
        .TextMatrix(0, mCol.ֻ����ɨ��������) = "ֻ����ɨ��������": .ColWidth(mCol.ֻ����ɨ��������) = 600: .ColAlignment(mCol.ֻ����ɨ��������) = flexAlignLeftCenter
    End With
End Sub

Private Sub initVfg��Ա()
    With vfg��Ա
        .BackColor = &H80000005
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .BackColorFixed = &HFDD6C6
        .GridLinesFixed = flexGridFlat
        .RowHeightMin = 300
        .Editable = flexEDNone
        
        .Rows = 2: .FixedRows = 1
        .Cols = 5: .FixedCols = 0
        
        .TextMatrix(0, mCol.��ԱID) = "": .ColWidth(mCol.��ԱID) = 0: .ColAlignment(mCol.��ԱID) = flexAlignRightCenter
        .ColHidden(mCol.��ԱID) = True
        .TextMatrix(0, mCol.��Ա����) = "���": .ColWidth(mCol.��Ա����) = 1000: .ColAlignment(mCol.��Ա����) = flexAlignLeftCenter
        .TextMatrix(0, mCol.����) = "����": .ColWidth(mCol.����) = 2000: .ColAlignment(mCol.����) = flexAlignLeftCenter
        .TextMatrix(0, mCol.Ĭ��) = "Ĭ��С��": .ColWidth(mCol.Ĭ��) = 1000: .ColAlignment(mCol.Ĭ��) = flexAlignLeftCenter
        .TextMatrix(0, mCol.��ע) = "��ע": .ColWidth(mCol.Ĭ��) = 1000: .ColAlignment(mCol.��ע) = flexAlignLeftCenter
    End With
End Sub

Private Sub zlRefresh()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo ErrHandle
        
    vfg����.Clear
    mstr�������� = ""
    Call initVfg����
    strSQL = "Select A.����id, B.����, B.����, A.�鿴, A.����,a.��������" & vbNewLine & _
            "From ����С������ A, �������� B" & vbNewLine & _
            "Where A.����id = B.ID And A.С��id = [1] Order by B.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    With vfg����
        Do Until rsTmp.EOF
            .TextMatrix(.Rows - 1, mCol.����Id) = Val("" & rsTmp.Fields("����id"))
            .TextMatrix(.Rows - 1, mCol.��������) = "" & rsTmp.Fields("����")
            .TextMatrix(.Rows - 1, mCol.����) = "" & rsTmp.Fields("����")
            
            mstr�������� = mstr�������� & "," & Val("" & rsTmp.Fields("����id"))
            If Val("" & rsTmp.Fields("�鿴")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.�鿴) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.�鿴) = flexUnchecked
            End If
            If Val("" & rsTmp.Fields("����")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.����) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.����) = flexUnchecked
            End If
            If Val("" & rsTmp.Fields("��������")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.ֻ����ɨ��������) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.ֻ����ɨ��������) = flexUnchecked
            End If
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Val(.TextMatrix(.Rows - 1, mCol.����Id)) = 0 Then .Rows = .Rows - 1
    End With
    
    vfg��Ա.Clear
    mstr������Ա = ""
    Call initVfg��Ա
    strSQL = "Select A.��Աid, B.���, B.����, A.Ĭ��С��, A.��ע" & vbNewLine & _
            "From ����С���Ա A, ��Ա�� B" & vbNewLine & _
            "Where A.��Աid = B.ID And A.С��id = [1] order by B.���"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngGroupID)
    With vfg��Ա
        Do Until rsTmp.EOF
            mstr������Ա = mstr������Ա & "," & Val("" & rsTmp.Fields("��Աid"))
            .TextMatrix(.Rows - 1, mCol.��ԱID) = Val("" & rsTmp.Fields("��Աid"))
            .TextMatrix(.Rows - 1, mCol.��Ա����) = "" & rsTmp.Fields("���")
            .TextMatrix(.Rows - 1, mCol.����) = "" & rsTmp.Fields("����")
            If Val("" & rsTmp.Fields("Ĭ��С��")) = 1 Then
                .Cell(flexcpChecked, .Rows - 1, mCol.Ĭ��) = flexChecked
            Else
                .Cell(flexcpChecked, .Rows - 1, mCol.Ĭ��) = flexUnchecked
            End If
            .TextMatrix(.Rows - 1, mCol.��ע) = "" & rsTmp.Fields("��ע")
            .Rows = .Rows + 1
            rsTmp.MoveNext
        Loop
        If Val(.TextMatrix(.Rows - 1, mCol.��ԱID)) = 0 Then .Rows = .Rows - 1
    End With
    Exit Sub
ErrHandle:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lng
    Select Case Control.ID
        Case conMenu_View_Refresh: Call zlLoadData
        Case conMenu_Help_Help:     Call ShowHelp(gstrLisHelp, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_View_ToolBar_Button
            Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In Me.cbsThis(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsThis.RecalcLayout
        Case conMenu_View_ToolBar_Size
            Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
            Me.cbsThis.RecalcLayout
        Case conMenu_View_StatusBar
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsThis.RecalcLayout

        Case conMenu_Edit_Save      '����
            Call zlSaveData
            mintEditState = 0
            Me.PicList.Enabled = True
            Me.vfg����.Enabled = True
            Me.vfg��Ա.Enabled = True
            Call pic��Ա_Resize
            Call pic����_Resize
        Case conMenu_Edit_Untread   '����
            Call zlEditCancle

        Case conMenu_File_Exit      '�˳�
            If mintEditState <> 0 Then
                If MsgBox("�ڱ༭״̬�У���������˳���δ������޸Ľ���ʧ���Ƿ������", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
            
        Case conMenu_Edit_NewItem   '����
            mintEditState = 1
            If zlEditStart(1) Then
                Me.PicList.Enabled = False
                Me.vfg��Ա.Enabled = False
                Me.vfg����.Enabled = False
            End If
        Case conMenu_Edit_Modify    '�޸�
            mintEditState = 1
            If zlEditStart(0) Then
                Me.PicList.Enabled = False
                Me.vfg��Ա.Enabled = False
                Me.vfg����.Enabled = False
            End If
        Case conMenu_Edit_Delete    'ɾ��
            If MsgBox("��ɾ����С�������е���Ա���������ã��Ƿ������", vbExclamation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                Call zlEditDelete
            End If
        Case conMenu_Edit_Append    '����
            mintEditState = 2
            If zlEditStart(0) Then
                Me.dkpMan.FindPane(conPane_����).Select
                Me.PicList.Enabled = False
                Me.vfg��Ա.Enabled = False
                Me.vfg����.Enabled = True
            End If
            
        Case conMenu_Edit_Compend   '��Ա
            mintEditState = 3
            If zlEditStart(1) Then
                Me.dkpMan.FindPane(conPane_��Ա).Select
                Me.PicList.Enabled = False
                Me.vfg����.Enabled = False
                Me.vfg��Ա.Enabled = True
            End If
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Edit_Save, conMenu_Edit_Untread     '����
            Control.Enabled = mintEditState <> 0
        Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Append, _
             conMenu_Edit_Compend, conMenu_View_Refresh
            Control.Enabled = mintEditState = 0
    End Select
End Sub

Private Sub cmdFind_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strFind As String
    
    On Error GoTo ErrHandle
    strFind = DelInvalidChar(UCase(Trim(Me.txtFind)))
    If strFind <> "" Then
        strSQL = "Select Distinct /*+Rule */" & vbNewLine & _
                " D.ID, D.���, D.����, D.�Ա�" & vbNewLine & _
                "From ��������˵�� B, ������Ա C, ��Ա�� D, ���ű� A" & vbNewLine & _
                "Where A.ID = B.����id And B.�������� = '����' And A.ID = C.����id And C.��Աid = D.ID And (" & _
                zlcommfun.GetLike("D", "���", strFind) & " or " & zlcommfun.GetLike("D", "����", strFind) & " or " & zlcommfun.GetLike("D", "����", strFind) & ")"
    Else
        strSQL = "Select Distinct /*+Rule */" & vbNewLine & _
                " D.ID, D.���, D.����, D.�Ա�" & vbNewLine & _
                "From ��������˵�� B, ������Ա C, ��Ա�� D, ���ű� A" & vbNewLine & _
                "Where A.ID = B.����id And B.�������� = '����' And A.ID = C.����id And C.��Աid = D.ID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFind)
    Me.lvw��Ա.ListItems.Clear
    Do Until rsTmp.EOF
        If InStr(mstr������Ա & ",", "," & rsTmp.Fields("id") & ",") <= 0 Then
            Me.lvw��Ա.ListItems.Add , "_" & rsTmp.Fields("id"), "" & rsTmp.Fields("���") & " " & rsTmp.Fields("����")
        End If
        rsTmp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd��Ա_Click(Index As Integer)
    Dim ObjItem  As ListItem
    With vfg��Ա
    If Index = 0 Then

        If Me.lvw��Ա.SelectedItem Is Nothing Then Exit Sub
        Set ObjItem = Me.lvw��Ա.SelectedItem
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.��ԱID) = Mid(ObjItem.Key, 2)
        .TextMatrix(.Rows - 1, mCol.��Ա����) = Split(ObjItem.Text, " ")(0)
        .TextMatrix(.Rows - 1, mCol.����) = Split(ObjItem.Text, " ")(1)
        
        .Cell(flexcpChecked, .Rows - 1, mCol.Ĭ��) = flexChecked
        
        If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
        Me.lvw��Ա.ListItems.Remove ObjItem.Key: Me.lvw��Ա.SetFocus
    Else
        If .Row < .FixedRows Then Exit Sub
        Set ObjItem = Me.lvw��Ա.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.��ԱID), .TextMatrix(.Row, mCol.��Ա����) & " " & .TextMatrix(.Row, mCol.����))
        ObjItem.Selected = True
        .RemoveItem .Row
    End If
    End With
End Sub

Private Sub cmd����_Click(Index As Integer)
    Dim ObjItem  As ListItem
    With vfg����
    If Index = 0 Then

        If Me.lvw����.SelectedItem Is Nothing Then Exit Sub
        Set ObjItem = Me.lvw����.SelectedItem
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, mCol.����Id) = Mid(ObjItem.Key, 2)
        .TextMatrix(.Rows - 1, mCol.��������) = Split(ObjItem.Text, " ")(0)
        .TextMatrix(.Rows - 1, mCol.����) = Split(ObjItem.Text, " ")(1)
        
        .Cell(flexcpChecked, .Rows - 1, mCol.�鿴) = flexChecked
        .Cell(flexcpChecked, .Rows - 1, mCol.����) = flexChecked
        .Cell(flexcpChecked, .Rows - 1, mCol.ֻ����ɨ��������) = flexChecked
        
        If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
        Me.lvw����.ListItems.Remove ObjItem.Key: Me.lvw����.SetFocus
    Else
        If .Row < .FixedRows Then Exit Sub
        Set ObjItem = Me.lvw����.ListItems.Add(, "_" & .TextMatrix(.Row, mCol.����Id), .TextMatrix(.Row, mCol.��������) & " " & .TextMatrix(.Row, mCol.����))
        ObjItem.Selected = True
        .RemoveItem .Row
    End If
    End With
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
    If Action = PaneActionFloating Then Cancel = True
    If Action = PaneActionClosing Then Cancel = True
    
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    
    Select Case Item.ID
    Case conPane_List
        Item.Handle = Me.PicList.hWnd
    Case conPane_����
        Item.Handle = Me.pic����.hWnd
    Case conPane_��Ա
        Item.Handle = Me.pic��Ա.hWnd
    End Select

End Sub

Private Sub Form_Load()
    
    mstrPrivs = gstrPrivs
    
    mintEditState = 0
    Call zlcommfun.SetWindowsInTaskBar(Me.hWnd, False)

    Me.lvw����.ListItems.Clear
    With Me.lvw����.ColumnHeaders
        .Clear
        .Add , "_ID", "�����б�", 3000
    End With
    With Me.lvw����
        .SortKey = .ColumnHeaders("_ID").Index - 1
        .SortOrder = lvwAscending
    End With

    Me.lvw��Ա.ListItems.Clear
    With Me.lvw��Ա.ColumnHeaders
        .Clear
        .Add , "_ID", "��Ա�б�", 3000
    End With
    With Me.lvw��Ա
        .SortKey = .ColumnHeaders("_ID").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call initMenu
    Call initPic
    Call CreateRptListHead
    Call zlLoadData
    Call RestoreWinState(Me, App.ProductName)
    '-----------------------------------------------------
End Sub

Private Sub Form_Resize()
'    Dim panBase As Pane
'    If Me.WindowState = vbMinimized Then Exit Sub
'    Set panBase = Me.dkpMan.FindPane(conPane_Edit)
'    mLngEditWidth = picGroupBase.ScaleHeight
'    panBase.MinTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'    panBase.MaxTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'    Me.dkpMan.RecalcLayout
'    Me.dkpMan.NormalizeSplitters
'
'    panBase.MinTrackSize.SetSize 0, 0
'    panBase.MaxTrackSize.SetSize 350, mLngEditWidth / Screen.TwipsPerPixelX
'
'    Me.dkpMan.RecalcLayout
'    Me.dkpMan.NormalizeSplitters

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvw��Ա_DblClick()
    Call cmd��Ա_Click(0)
End Sub

Private Sub lvw����_DblClick()
    Call cmd����_Click(0)
End Sub

Private Sub picEdit��Ա_Resize()
    Err = 0: On Error Resume Next
    Me.cmd��Ա(0).Left = Me.picEdit��Ա.ScaleLeft
    Me.cmd��Ա(0).Top = Me.picEdit��Ա.ScaleTop
    Me.cmd��Ա(0).Width = Me.picEdit��Ա.ScaleWidth / 2
    
    Me.cmd��Ա(1).Left = Me.cmd��Ա(0).Left + Me.cmd��Ա(0).Width
    Me.cmd��Ա(1).Top = Me.picEdit��Ա.ScaleTop
    Me.cmd��Ա(1).Width = Me.picEdit��Ա.ScaleWidth / 2
    
    
    Me.txtFind.Top = Me.cmd��Ա(0).Top + Me.cmd��Ա(0).Height + 15
    Me.txtFind.Left = Me.picEdit��Ա.ScaleLeft + Me.lblEdit��Ա.Width + 15
    Me.txtFind.Width = Me.picEdit��Ա.ScaleWidth - Me.txtFind.Left - Me.cmdFind.Width - 10
    
    Me.cmdFind.Left = Me.txtFind.Left + Me.txtFind.Width + 10
    Me.cmdFind.Top = Me.txtFind.Top
    
    
    Me.lblEdit��Ա.Top = Me.txtFind.Top + 25
    Me.lblEdit��Ա.Left = Me.picEdit��Ա.ScaleLeft
    
    
    Me.lvw��Ա.Left = Me.picEdit��Ա.ScaleLeft
    Me.lvw��Ա.Top = Me.txtFind.Top + Me.txtFind.Height + 15
    Me.lvw��Ա.Width = Me.picEdit��Ա.ScaleWidth
    Me.lvw��Ա.Height = Me.picEdit��Ա.ScaleHeight - Me.lvw��Ա.Top

End Sub

Private Sub picEdit����_Resize()
    Err = 0: On Error Resume Next
    Me.cmd����(0).Left = Me.picEdit����.ScaleLeft
    Me.cmd����(0).Top = Me.picEdit����.ScaleTop
    Me.cmd����(0).Width = Me.picEdit����.ScaleWidth / 2
    
    Me.cmd����(1).Left = Me.cmd����(0).Left + Me.cmd����(0).Width
    Me.cmd����(1).Top = Me.picEdit����.ScaleTop
    Me.cmd����(1).Width = Me.picEdit����.ScaleWidth / 2
    
    Me.lblEdit����.Top = Me.cmd����(0).Top + Me.cmd����(0).Height + 15
    Me.lblEdit����.Left = Me.picEdit����.ScaleLeft
    
    Me.lvw����.Left = Me.picEdit����.ScaleLeft
    Me.lvw����.Top = Me.lblEdit����.Top + Me.lblEdit����.Height + 15
    Me.lvw����.Width = Me.picEdit����.ScaleWidth
    Me.lvw����.Height = Me.picEdit����.ScaleHeight - Me.lvw����.Top
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With Me.rptList
        .Left = Me.PicList.ScaleLeft: .Width = Me.PicList.ScaleWidth - .Left
        .Height = Me.PicList.ScaleHeight - .Top
    End With
End Sub

Private Sub pic��Ա_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg��Ա
        .Left = Me.pic��Ա.ScaleLeft
        .Top = Me.pic��Ա.ScaleTop
        .Width = Me.pic��Ա.ScaleWidth
        If mintEditState = 3 Then
            .Height = Me.pic��Ա.ScaleHeight - Me.picEdit��Ա.Height
            With Me.picEdit��Ա
                .Left = Me.pic��Ա.ScaleLeft
                .Top = Me.vfg��Ա.Top + Me.vfg��Ա.Height
                .Width = Me.pic��Ա.ScaleWidth
                .Visible = True
            End With
            
        Else
            .Height = Me.pic��Ա.ScaleHeight
            Me.picEdit��Ա.Visible = False
        End If
    End With
End Sub

Private Sub pic����_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg����
        .Left = Me.pic����.ScaleLeft
        .Top = Me.pic����.ScaleTop
        .Width = Me.pic����.ScaleWidth
        If mintEditState = 2 Then
            .Height = Me.pic����.ScaleHeight - Me.picEdit����.Height
            With Me.picEdit����
                .Left = Me.pic����.ScaleLeft
                .Top = Me.vfg����.Top + Me.vfg����.Height
                .Width = Me.pic����.ScaleWidth
                .Visible = True
            End With
            
        Else
            .Height = Me.pic����.ScaleHeight
            Me.picEdit����.Visible = False
        End If
    End With
End Sub

Private Sub rptList_SelectionChanged()
    Dim i As Integer
    
    mstr�������� = ""
    If rptList.SelectedRows.count = 0 Then
        If rptList.Rows.count > 0 Then
            '�м�¼,ȡ�ڸ��Ƿ�����,����ǰ��
            For i = 0 To rptList.Rows.count - 1
                If Not rptList.Rows(i).GroupRow Then
                    rptList.Rows(i).Selected = True
                    mlngGroupID = Val(Me.rptList.Rows(i).Record(mCol.ID).Value)
                    mstr�������� = Me.rptList.Rows(i).Record(mCol.С�����).Value & "|" & Me.rptList.Rows(i).Record(mCol.С��).Value
                    Exit For
                End If
            Next
        End If
    End If
        
    If Not Me.rptList.FocusedRow Is Nothing Then
        If Me.rptList.FocusedRow.GroupRow = True Then
            mlngGroupID = 0
        Else
            mlngGroupID = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
            mstr�������� = Me.rptList.FocusedRow.Record(mCol.С�����).Value & "|" & Me.rptList.FocusedRow.Record(mCol.С��).Value
        End If
    End If
    
    Dim panThis As Pane, panSub As Pane
    If mblnInit Then
        Set panThis = Me.dkpMan.FindPane(conPane_����)
        Set panSub = Me.dkpMan.FindPane(conPane_��Ա)
        panSub.AttachTo panThis
        panThis.Select
        mblnInit = False
    End If
    
    Me.dkpMan.RecalcLayout
    
    Call zlRefresh
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlcommfun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0

End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
    End If
End Sub

Private Sub vfg��Ա_DblClick()
    Dim blnCheck As Boolean
'    If mintEditState = 3 Then
'        If InStr("," & mCol.Ĭ�� & ",", "," & vfg����.MouseCol & ",") <= 0 Then
'            Call cmd��Ա_Click(1)
'        Else
             blnCheck = Me.vfg��Ա.Cell(flexcpChecked, vfg��Ա.MouseRow, vfg��Ա.MouseCol) = flexChecked
             If blnCheck Then
                Me.vfg��Ա.Cell(flexcpChecked, vfg��Ա.MouseRow, vfg��Ա.MouseCol) = flexUnchecked
             Else
                Me.vfg��Ա.Cell(flexcpChecked, vfg��Ա.MouseRow, vfg��Ա.MouseCol) = flexChecked
             End If
'        End If
'    End If

End Sub

Private Sub vfg����_DblClick()
    Dim blnCheck As Boolean
'    If mintEditState = 2 Then
'        If InStr("," & mCol.���� & ",", "," & vfg����.MouseCol & ",") <= 0 Then
'            Call cmd����_Click(1)
'        Else
             blnCheck = Me.vfg����.Cell(flexcpChecked, vfg����.MouseRow, vfg����.MouseCol) = flexChecked
             If blnCheck Then
                Me.vfg����.Cell(flexcpChecked, vfg����.MouseRow, vfg����.MouseCol) = flexUnchecked
             Else
                Me.vfg����.Cell(flexcpChecked, vfg����.MouseRow, vfg����.MouseCol) = flexChecked
             End If
'        End If
'    End If
End Sub
