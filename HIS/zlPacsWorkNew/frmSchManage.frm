VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{79EB16A5-917F-4145-AB5F-D3AEA60612D8}#16.3#0"; "CODEJOCK.CALENDAR.V16.3.1.OCX"
Begin VB.Form frmSchManage 
   Caption         =   "���ԤԼ����"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmSchManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   11715
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picTimeTable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   5040
      ScaleHeight     =   4935
      ScaleWidth      =   6255
      TabIndex        =   1
      Top             =   600
      Width           =   6255
      Begin TabDlg.SSTab sstTimeTable 
         Height          =   4575
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "ԤԼʱ���"
         TabPicture(0)   =   "frmSchManage.frx":0442
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "schTimeTable"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "ԤԼ�б�"
         TabPicture(1)   =   "frmSchManage.frx":045E
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "vsfSchList"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin zl9PACSWork.ucScheduleTimetable schTimeTable 
            Height          =   4935
            Left            =   -74760
            TabIndex        =   9
            Top             =   600
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   8705
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfSchList 
            Height          =   3735
            Left            =   360
            TabIndex        =   10
            Top             =   600
            Width           =   4695
            _cx             =   8281
            _cy             =   6588
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
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
      End
   End
   Begin VB.PictureBox picCalendar 
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   120
      ScaleHeight     =   7335
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VSFlex8Ctl.VSFlexGrid vsfSchDevice 
         Height          =   2070
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2895
         _cx             =   5106
         _cy             =   3651
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
      Begin XtremeCalendarControl.DatePicker dpCalendar 
         Height          =   2655
         Left            =   720
         TabIndex        =   5
         Top             =   2040
         Width           =   3135
         _Version        =   1048579
         _ExtentX        =   5530
         _ExtentY        =   4683
         _StockProps     =   64
         AutoSize        =   0   'False
         ShowTodayButton =   0   'False
         ShowNoneButton  =   0   'False
         ShowNonMonthDays=   0   'False
         Show3DBorder    =   2
         AskDayMetrics   =   -1  'True
         TextTodayButton =   "���ؽ���"
      End
      Begin VB.Frame frmChangeDay 
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   4560
         Width           =   3135
         Begin VB.CommandButton cmdChangeDay 
            Caption         =   "ǰһ��"
            Height          =   495
            Index           =   1
            Left            =   10
            TabIndex        =   11
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdChangeDay 
            Caption         =   "��һ��"
            Height          =   495
            Index           =   2
            Left            =   2250
            TabIndex        =   4
            Top             =   120
            Width           =   855
         End
         Begin VB.CommandButton cmdToday 
            Caption         =   "����"
            Height          =   495
            Left            =   870
            TabIndex        =   3
            Top             =   120
            Width           =   1380
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfWeekView 
         Height          =   1095
         Left            =   240
         TabIndex        =   12
         Top             =   5760
         Width           =   3255
         _cx             =   5741
         _cy             =   1931
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
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
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7575
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   4154
            MinWidth        =   4154
            Picture         =   "frmSchManage.frx":047A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13785
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
      Left            =   3480
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmSchManage.frx":0D0E
      Left            =   4080
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSchManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngOrderID As Long             'ҽ��ID
Private mlngSchDeviceID As Long         'ԤԼ�豸ID
Private mstrDeptIDs As String           '����ID
Private mschDate As Date                '��ǰ����
Private mfrmParent As Object            '������
Private mstrModifiedOrderID As String   '�����ԤԼ��Ϣ��ҽ��ID�����á�,������
Private mstrSchRestDate As String       '������Ϣ��
Private mstrPrivs As String             '�����ߵ�Ȩ��

Private mlngColorLblWaiting As Long     'ԤԼ��ǩ��ԤԼ�Ⱥ���ɫ
Private mlngColorLblDone As Long        'ԤԼ��ǩ�������ɫ
Private mlngColorLblPassed As Long      'ԤԼ��ǩ��������ɫ
Private mblnAutoPrint As Boolean        '�Ƿ��Զ���ӡԤԼ��

'���ԤԼ�豸
Private Enum constScheduleDeviceList
    col_SchDevice_ID = 0
    col_SchDevice_Ӱ����� = 1
    col_SchDevice_�豸���� = 2
    col_SchDevice_�豸˵�� = 3
End Enum

'���ԤԼ�б�
Private Enum constScheduleList
    col_SchList_ID = 0
    col_SchList_��� = 1
    col_SchList_���� = 2
    col_SchList_ҽ������ = 3
    col_SchList_�������� = 4
    col_SchList_ԤԼ��ʼʱ�� = 5
    col_SchList_ԤԼ����ʱ�� = 6
    col_SchList_ִ�й��� = 7
    col_SchList_�ֻ��� = 8
    col_SchList_���ע�� = 9
End Enum

'���ԤԼ����ͼ
Private Enum constWeekView
    col_WeekView_���� = 0
    col_WeekView_���� = 1
    col_WeekView_��ԤԼ = 2
    col_WeekView_������ = 3
End Enum

Private Sub InitCommandBar()
'------------------------------------------------
'���ܣ���ʼ��������
'������ ��
'���أ� ��
'------------------------------------------------
    Dim cbrToolBar As CommandBar
    Dim cbrControl As CommandBarControl
    
    On Error GoTo err
    
    '�ⲿ��ȫ�����ã��Ƿ��Ҫ��
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = zlCommFun.GetPubIcons
        
    With cbrMain.options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True         '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    
    '����ʾ�˵�
    cbrMain.ActiveMenuBar.Visible = False
    
    '��ʾ������
    Set cbrToolBar = cbrMain.Add("ԤԼ������", xtpBarLeft)
    cbrToolBar.EnableDocking xtpFlagStretched
    cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.Closeable = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Save, "����ԤԼ")
        cbrControl.iconid = 6823
        cbrControl.ToolTipText = "����ԤԼ��Ϣ"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Print, "��ӡԤԼ��")
        cbrControl.iconid = 103
        cbrControl.ToolTipText = "��ӡ���ߵ�ԤԼ֪ͨ��"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Modify, "�޸�ԤԼ")
        cbrControl.iconid = 6886
        cbrControl.ToolTipText = "�޸ļ��ԤԼ"
        
        
'        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Query, "ԤԼ��ѯ")
'        cbrControl.IconId = 3946
'        cbrControl.ToolTipText = "��ѯԤԼ���"
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Delete, "ɾ��ԤԼ")
        cbrControl.iconid = 6822
        cbrControl.ToolTipText = "ɾ��һ�����ԤԼ"
        
        cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Refresh, "ˢ��")
        cbrControl.iconid = 791
        cbrControl.ToolTipText = "ˢ������"
        
        cbrControl.BeginGroup = True
        
        
        Set cbrControl = .Add(xtpControlButton, conMenu_PacsSchdule_Quit, "�˳�")
        cbrControl.iconid = 191
        cbrControl.ToolTipText = "�رմ���"
        
        cbrControl.BeginGroup = True
    End With
    
    cbrToolBar.Position = xtpBarTop
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Modify     '�޸�ԤԼ
            '�򿪼��ԤԼ����
            Call ModifySchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Delete     'ɾ��ԤԼ
            Call DelSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Refresh    'ˢ��
            Call RefreshSchedule
            
'        Case conMenu_PacsSchdule_Query      '��ѯ
'            Call frmSchQuery.zlShowMe(mstrDeptIDs, Me)
            
        Case conMenu_PacsSchdule_Print      '��ӡԤԼ��
            Call PrintSchedule(mlngOrderID)
            
        Case conMenu_PacsSchdule_Save       '����ԤԼ
            Call SaveSchedule
            Call RefreshSchedule
            
        Case conMenu_PacsSchdule_Quit       '�˳�
            Unload Me
            
    End Select
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_PacsSchdule_Refresh, conMenu_PacsSchdule_Quit   'ˢ�� ,�˳�
            'ʲô������
        Case Else
            Control.Enabled = IIf(sstTimeTable.Tab = 1, False, True)
    End Select
End Sub

Private Sub cmdChangeDay_Click(Index As Integer)
    Select Case Index
        Case 1
            mschDate = mschDate - 7
        Case 2
            mschDate = mschDate + 7
    End Select
    
    Call ChangeCalendar(mschDate)
    Call RefreshSchedule
End Sub

Private Sub cmdToday_Click()
    mschDate = Format(Now, "YYYY-MM-DD")
    Call ChangeCalendar(mschDate)
    Call RefreshSchedule
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picCalendar.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picTimeTable.hwnd
    End If
End Sub

Private Sub dpCalendar_DayMetrics(ByVal Day As Date, ByVal Metrics As XtremeCalendarControl.IDatePickerDayMetrics)
    If InStr(mstrSchRestDate, Format(Day, "YYYY-MM-DD")) > 0 Then
        Metrics.ForeColor = vbRed
        Metrics.Font.Bold = True
    End If
End Sub

Private Sub dpCalendar_MonthChanged()
    Call RefreshCalendar
End Sub

Private Sub dpCalendar_SelectionChanged()
    '���������ڣ�����ˢ��ʱ���
    
    mschDate = dpCalendar.Selection.Blocks(0).DateBegin
    ChangeCalendar (mschDate)
    Call RefreshSchedule
End Sub

Private Sub InitFaceScheme()
'------------------------------------------------
'���ܣ���ʼ�����沼��
'������ ��
'���أ� ��
'------------------------------------------------
    Dim Pane1 As Pane, Pane2 As Pane
    
    On Error GoTo err
    
    '����������ʾ����
    With dkpMain
        .VisualTheme = ThemeOffice2003
        .SetCommandBars cbrMain
        .options.HideClient = True
        .options.UseSplitterTracker = False 'ʵʱ�϶�
        .options.ThemedFloatingFrames = True
        .options.AlphaDockingContext = True
        dkpMain.options.DefaultPaneOptions = PaneNoCaption
    End With
    
    '�ȴ�ע����ȡԤ�����úõĴ��ڲ��֣�Ȼ�����������
    dkpMain.LoadStateFromString GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, "")
    
    '���ע����б���Ľ��沼��Pane�������ԣ������Ĭ�ϵ�Pane����
    If dkpMain.PanesCount <> 2 Then
        dkpMain.DestroyAll
        
        Set Pane1 = dkpMain.CreatePane(1, 350, 150, DockLeftOf, Nothing)
        Pane1.title = "ԤԼ��Ϣ"
        Pane1.options = PaneNoCaption
        
        Set Pane2 = dkpMain.CreatePane(2, 650, 300, DockRightOf, Pane1)
        Pane2.title = "ԤԼʱ���"
        Pane2.options = PaneNoCaption
    End If
    
    'Ĭ����ʾʱ���
    sstTimeTable.Tab = 0
    vsfSchList.Visible = False
    
    mlngColorLblWaiting = zlDatabase.GetPara("���ԤԼ��ǩ��ԤԼ��ɫ", glngSys, 1292, "0")
    mlngColorLblDone = zlDatabase.GetPara("���ԤԼ��ǩ�������ɫ", glngSys, 1292, "12632256")
    mlngColorLblPassed = zlDatabase.GetPara("���ԤԼ��ǩ�ѹ�����ɫ", glngSys, 1292, "255")
    
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    Dim strOrders() As String
    
    '�ϲ�ɾ���ͱ����ҽ��ID
    strOrders = Split(schTimeTable.strModifiedOrderID, ",")
    For i = 0 To UBound(strOrders)
        If InStr(mstrModifiedOrderID, CStr(strOrders(i))) = 0 Then
            mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(strOrders(i))
        End If
    Next i
    
    If Trim(mstrModifiedOrderID) <> "" Then
        mstrModifiedOrderID = Mid(mstrModifiedOrderID, 2)
    End If
    
    '�رմ����ʱ�򣬱�����沼��
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.�û��� & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name, dkpMain.SaveStateToString)
    
    Call SaveWinState(Me, App.ProductName)
    
    Set mfrmParent = Nothing
    
    '�ͷ�DockingPane
    For i = 1 To dkpMain.PanesCount
        dkpMain.Panes(i).Handle = 0
    Next i
    dkpMain.CloseAll
End Sub

Private Sub picCalendar_Resize()
    On Error Resume Next
    
    vsfSchDevice.Left = 0
    vsfSchDevice.Top = 0
    vsfSchDevice.Width = picCalendar.ScaleWidth
    
    dpCalendar.Left = 0
    dpCalendar.Top = vsfSchDevice.Height
    dpCalendar.Width = picCalendar.ScaleWidth
    
    frmChangeDay.Left = 0
    frmChangeDay.Top = dpCalendar.Top + dpCalendar.Height - 50
    frmChangeDay.Width = picCalendar.ScaleWidth
    
    cmdToday.Width = frmChangeDay.Width - cmdChangeDay(2).Width * 2 - 20
    
    cmdChangeDay(2).Left = frmChangeDay.Width - cmdChangeDay(2).Width - 10
    
    vsfWeekView.Left = 0
    vsfWeekView.Top = frmChangeDay.Top + frmChangeDay.Height + 40
    vsfWeekView.Width = picCalendar.ScaleWidth
    vsfWeekView.Height = picCalendar.ScaleHeight - vsfWeekView.Top - 50
    
End Sub

Private Sub picTimeTable_Resize()
    On Error Resume Next
    
    sstTimeTable.Left = 0
    sstTimeTable.Top = 0
    sstTimeTable.Width = picTimeTable.ScaleWidth
    sstTimeTable.Height = picTimeTable.ScaleHeight - stbThis.Height
    
    schTimeTable.Left = 0
    schTimeTable.Top = 300
    schTimeTable.Width = sstTimeTable.Width - 20
    schTimeTable.Height = sstTimeTable.Height - 300
'
    vsfSchList.Left = 0
    vsfSchList.Top = 300
    vsfSchList.Width = sstTimeTable.Width - 20
    vsfSchList.Height = sstTimeTable.Height - 300
End Sub

Private Sub ChangeCalendar(dtDate As Date)
'------------------------------------------------
'���ܣ��޸�ԤԼ����������
'������dtDate -- ����������
'���أ���
'------------------------------------------------
    On Error GoTo err
    
    dpCalendar.ClearSelection
    Call dpCalendar.Select(dtDate)
    dpCalendar.EnsureVisibleSelection
    If dpCalendar.Visible = True Then
        dpCalendar.SetFocus
    End If
    
    Call LoadWeekView
    Exit Sub
err:
    If ErrCenter() = 1 Then Resume
End Sub

Public Function ZlShowMe(strPrivs As String, strDeptIDs As String, lngOrderID As Long, frmParent As Object) As String
'------------------------------------------------
'���ܣ��򿪴���
'������ strDeptIDs -- ����ID��
'       lngOrderID -- ҽ��ID������ԤԼ������ʾ������
'       frmParent -- ������
'       strPrivs -- �����ߵ�Ȩ��
'���أ������ԤԼ��Ϣ��ҽ��ID�����á�,������
'------------------------------------------------
    On Error GoTo err
    
    mlngOrderID = 0
    mlngSchDeviceID = 0
    mstrPrivs = strPrivs
    
    mstrDeptIDs = strDeptIDs
    Set mfrmParent = frmParent
    
    mstrModifiedOrderID = ""
    
    '��ʼ�����沼��
    Call InitFaceScheme
    
    '����������
    Call InitCommandBar
    
    '��ȡϵͳ����
    mblnAutoPrint = IIf(Val(zlDatabase.GetPara("����ԤԼ���Զ���ӡԤԼ��", glngSys, 1292)) = 1, True, False)
    
    '�ȳ�ʼ��ʱ���ؼ�
    Call schTimeTable.Init(2)   'ԤԼ����ģʽ
    
    Call RestoreWinState(Me, App.ProductName)
    
    '������������
    dpCalendar.ShowNonMonthDays = False
    dpCalendar.AskDayMetrics = True
    
    mschDate = GetOrderSchDate(lngOrderID)
    
    If LoadData = False Then
        Exit Function
    End If
    
    Me.Show 1, mfrmParent
    
    ZlShowMe = mstrModifiedOrderID
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Function GetOrderSchDate(lngOrderID As Long) As Date
'------------------------------------------------
'���ܣ�����ҽ��ID����ȡ����ҽ����ԤԼʱ�䣬���û��ԤԼ�����ؽ���
'������
'       lngOrderID -- ҽ��ID������ԤԼ������ʾ������
'���أ�ԤԼ������ʾ������
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSQL = "select ԤԼ��ʼʱ�� from Ӱ��ԤԼ��¼ where ҽ��ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ����", lngOrderID)
    
    If rsTemp.EOF = False Then
        GetOrderSchDate = Format(rsTemp!ԤԼ��ʼʱ��, "YYYY-MM-DD")
    Else
        GetOrderSchDate = Format(Now, "YYYY-MM-DD")
    End If
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function


Private Function LoadData() As Boolean
'------------------------------------------------
'���ܣ����ش������������
'������
'���أ�True -- �ɹ��� False -- ʧ��
'------------------------------------------------
    On Error GoTo err
    
    '���Ⱥ�˳��
    If LoadSchDevice = False Then
        Exit Function
    End If
    
    '��������
    Call ChangeCalendar(mschDate)
    Call RefreshCalendar
    
    'ˢ��ʱ���
    Call RefreshSchedule
    
    LoadData = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadSchDevice() As Boolean
'------------------------------------------------
'���ܣ�����ԤԼ�豸
'������
'���أ�True -- �ɹ� ��False -- ʧ��
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo err
    
    strSQL = "select ID,�豸����,Ӱ���豸��,Ӱ�����,�豸˵�� from Ӱ��ԤԼ�豸 where ����ID in (" & mstrDeptIDs & ") and �Ƿ�����=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ�豸")
    
    With vsfSchDevice
        .Clear
        .Cols = 4
        .Rows = rsTemp.RecordCount + 1
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .RowHeightMin = 400
        
        .ColWidth(col_SchDevice_ID) = 50
        .ColWidth(col_SchDevice_�豸����) = 2000
        .ColWidth(col_SchDevice_Ӱ�����) = 5000
        .ColWidth(col_SchDevice_�豸˵��) = 500
        
        '�ϲ���һ��
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        For i = 0 To 3
            .TextMatrix(0, i) = "ԤԼ�豸"
        Next i
        
        '�����ݿ��������
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_SchDevice_ID) = rsTemp!ID
            .TextMatrix(i, col_SchDevice_�豸����) = rsTemp!�豸����
            .TextMatrix(i, col_SchDevice_Ӱ�����) = rsTemp!Ӱ�����
            .TextMatrix(i, col_SchDevice_�豸˵��) = NVL(rsTemp!�豸˵��)
            rsTemp.MoveNext
        Next i
        
        '���غ�̨������
        .ColHidden(col_SchDevice_ID) = True
        .ColHidden(col_SchDevice_Ӱ�����) = True
        
        'ѡ���һ��
        If .Rows > 1 Then
            Call .Select(1, 1)
            mlngSchDeviceID = Val(.TextMatrix(1, col_SchDevice_ID))
        Else
            mlngSchDeviceID = 0
            Call MsgBoxD(Me, "û�п�����ԤԼ��Ӱ���豸���������ԤԼ�豸��", vbOKOnly, "���ԤԼ��ʾ")
            Exit Function
        End If
    End With
    
    LoadSchDevice = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function RefreshSchedule() As Boolean
'------------------------------------------------
'���ܣ�����ˢ��ʱ�������
'������
'���أ�True -- �ɹ� �� False -- ʧ��
'------------------------------------------------
    On Error GoTo err
    
    'ˢ��ԤԼ�б�
    Call LoadSchList
    
    '�Ѿ�����ԤԼ��Ϣ��ֱ����ʾ����
    If schTimeTable.RefreshSchedule(mlngSchDeviceID, mschDate, mlngOrderID) = True Then
        mlngOrderID = schTimeTable.LabelOrderID
    Else
        mlngOrderID = 0
    End If
        
    RefreshSchedule = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub DelSchedule(lngOrderID As Long)
'------------------------------------------------
'���ܣ�ɾ��ԤԼ
'������ lngOrderID -- ҽ��ID
'���أ���
'------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo err
    
    If lngOrderID = 0 Then
        Exit Sub
    End If
    
    If InStr(mstrModifiedOrderID, CStr(lngOrderID)) = 0 Then
        mstrModifiedOrderID = mstrModifiedOrderID & "," & CStr(lngOrderID)
    End If
    
    strSQL = "Zl_Ӱ��ԤԼ��¼_ɾ��(" & lngOrderID & ")"
    zlDatabase.ExecuteProcedure strSQL, "ɾ�����ԤԼ"
    
    Call RefreshSchedule
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub schTimeTable_OnChangeOrder(ByVal lngOrderID As Long, ByVal strOrderInfo As String)
    mlngOrderID = lngOrderID
    stbThis.Panels(2).Text = strOrderInfo
End Sub

Private Sub schTimeTable_OnMenuScheduleModify()
    '�򿪼��ԤԼ����
    Call ModifySchedule(mlngOrderID)
End Sub

Private Sub schTimeTable_OnMenuSchedulePrint()
    Call PrintSchedule(mlngOrderID)
End Sub

Private Sub schTimeTable_OnSchLabelModifed(ByVal iIndex As Integer)
    stbThis.Panels(2).Text = schTimeTable.LabelOrderInfo
End Sub

Private Sub sstTimeTable_Click(PreviousTab As Integer)
    '�л���ҳ��
    If PreviousTab <> sstTimeTable.Tab Then
        If sstTimeTable.Tab = 0 Then
            schTimeTable.Visible = True
            vsfSchList.Visible = False
        Else
            '��ʾԤԼ�б�
            schTimeTable.Visible = False
            vsfSchList.Visible = True
            Call LoadSchList
        End If
    End If
    
End Sub

Private Sub vsfSchDevice_Click()
    If vsfSchDevice.Rows > 1 Then
        '�޸ĵ�ǰ��ѡ�е�ԤԼ�豸ID
        mlngSchDeviceID = vsfSchDevice.TextMatrix(vsfSchDevice.RowSel, col_SchDevice_ID)
        Call RefreshSchedule
        Call RefreshCalendar
    End If
End Sub

Private Sub SaveSchedule()
'------------------------------------------------
'���ܣ�����ԤԼ
'������
'���أ���
'------------------------------------------------
    Dim i As Integer
    Dim arrOrderID() As String
    
    On Error GoTo err
        
    Call schTimeTable.SaveAllSchedule
    
    '�Զ���ӡԤԼ��
    If mblnAutoPrint = True Then
        arrOrderID = Split(schTimeTable.strModifiedOrderID, ",")
        For i = 0 To UBound(arrOrderID)
            Call PrintSchedule(arrOrderID(i))
        Next i
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
        
End Sub

Private Sub PrintSchedule(ByVal lngOrderID As Long)
'------------------------------------------------
'���ܣ���ӡ��ǰԤԼ��
'������ lngOrderID -- ҽ��ID
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsReports As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim blnPrinted As Boolean
    Dim lngUniFmt As Long           'ͨ�ñ����ʽ���
    
    On Error GoTo err
    
    '��ӡԤԼ��
    If lngOrderID <> 0 Then
        '���ȼ�鱨���Ƿ�ֻ��һ����ʽ
        strSQL = "Select a.ID,a.���,b.���,b.˵�� From zlreports a,zlrptfmts b Where a.Id=b.����ID And a.���=[1] Order By ���"
        Set rsReports = zlDatabase.OpenSQLRecord(strSQL, "��ѯԤԼ�������ʽ", "ZL1_INSIDE_1290_01")

        If rsReports.EOF = True Then
            Call MsgBox("����ZL1_INSIDE_1290_01�������ڣ�����ϵ����Ա��Ӵ˱���", vbInformation, "���ԤԼ��ʾ")
            Exit Sub
        End If
        '����ж����ʽ������������ĿID�����Ҷ�Ӧ�ı����ʽ����
        If rsReports.RecordCount > 1 Then
            strSQL = "Select a.���� From �����ļ��б� A, ��������Ӧ�� B, ����ҽ����¼ C " _
                & " Where c.������Ŀid = b.������Ŀid And decode(c.������Դ, 3, 1, c.������Դ) = b.Ӧ�ó��� " _
                & "And b.�����ļ�id = a.ID And c.ID = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ļ�����", lngOrderID)
            
            If rsTemp.EOF = False Then
            While rsReports.EOF = False And blnPrinted = False
                If NVL(rsReports!˵��) = "ͨ�ü��ԤԼ��" Then
                    lngUniFmt = rsReports!���
                End If
                
                If NVL(rsReports!˵��) = NVL(rsTemp!����) Then
                    If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & lngOrderID, "ReportFormat=" & rsReports!���, 2) = False Then
                        Call MsgBox("����ZL1_INSIDE_1290_01���У���ʽΪ��" & NVL(rsReports!˵��) & "�ı����򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
                    Else
                        '��ӡ���˳�ѭ��
                        blnPrinted = True
                    End If
                Else
                    rsReports.MoveNext
                End If
            Wend
            End If
            '���û�У�����ҡ�ͨ�ü��ԤԼ������������ӡ
            If blnPrinted = False Then
                If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & lngOrderID, "ReportFormat=" & lngUniFmt, 2) = False Then
                    Call MsgBox("����ZL1_INSIDE_1290_01���У���ʽΪ����ͨ�ü��ԤԼ�����ı����򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
                Else
                    blnPrinted = True
                End If
            End If
        Else
            If ReportOpen(gcnOracle, 100, "ZL1_Inside_1290_01", Me, "ҽ��ID=" & lngOrderID, 2) = False Then
                Call MsgBox("����ZL1_INSIDE_1290_01���򿪲��ɹ�������ϵ����Ա�����˱���", vbInformation, "���ԤԼ��ʾ")
            Else
                blnPrinted = True
            End If
        End If
        
        'д���ӡ��¼
        strSQL = "Zl_Ӱ��ԤԼ��¼_��ӡ(" & lngOrderID & ")"
        zlDatabase.ExecuteProcedure strSQL, "���ԤԼ����ӡ"
        
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifySchedule(lngOrderID As Long)
'------------------------------------------------
'���ܣ��޸�ԤԼ
'������lngOrderID -- ҽ��ID
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    If schTimeTable.LabelOrderID <> 0 Then
        strSQL = "select ִ�й��� from ����ҽ������ where ҽ��ID =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯִ�й���", lngOrderID)
        
        If rsTemp.EOF = False Then
            If NVL(rsTemp!ִ�й���, 0) = 0 Or NVL(rsTemp!ִ�й���, 0) = 1 Then
                Call frmSchSchedule.ZlShowMe(mstrPrivs, lngOrderID, mstrDeptIDs, Me)
                Call RefreshSchedule
            Else
                MsgBox "����ԤԼ�Ѿ���ִ�У������޸ġ�", vbOKOnly, "���ԤԼ��ʾ"
            End If
        End If
    End If
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSchList()
'------------------------------------------------
'���ܣ�����ԤԼ��¼�б�
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim rsBaby As ADODB.Recordset
    Dim i As Integer
    Dim lngColor As Long
    
    On Error GoTo err
    
    strSQL = "Select d.ID, d.ҽ��ID, d.���, d.��������, d.ԤԼ��ʼʱ��, d.ԤԼ����ʱ��, " _
        & " d.ԤԼ��ʼʱ���, d.ԤԼ����ʱ���, b.����, b.ҽ������, b.Ӥ��, c.ִ�й���,d.���ע��,e.�ֻ��� " _
        & " From ����ҽ����¼ B, ����ҽ������ C,Ӱ��ԤԼ��¼ D,������Ϣ E Where b.id in " _
        & " (Select  a.ҽ��ID From Ӱ��ԤԼ��¼ A Where a.ԤԼ�豸ID = [1] And " _
        & " a.ԤԼ��ʼʱ�� Between [2] And [3] )And c.ҽ��id = b.id and d.ҽ��id=B.id  And b.����id = e.����id Order By cast(d.��� as int)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�����ԤԼ��¼", mlngSchDeviceID, CDate(Format(mschDate, "yyyy-MM-dd 00:00:00")), CDate(Format(mschDate, "yyyy-MM-dd 23:59:59")))
    
    With vsfSchList
        .Clear
        .Cols = 10
        .Rows = IIf(rsTemp.EOF = True, 1, rsTemp.RecordCount + 1)
        .FixedRows = 1
        .FixedCols = 0
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSort
        .Cell(flexcpAlignment, 0, 0, 0, 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .RowHeightMin = 350
        
        .ColWidth(col_SchList_ID) = 50
        .ColWidth(col_SchList_���) = 450
        .ColWidth(col_SchList_����) = 800
        .ColWidth(col_SchList_ҽ������) = 3000
        .ColWidth(col_SchList_ԤԼ��ʼʱ��) = 1800
        .ColWidth(col_SchList_ԤԼ����ʱ��) = 1800
        .ColWidth(col_SchList_�ֻ���) = 2000
        
        '��ʾ����
        .TextMatrix(0, col_SchList_ID) = "ID"
        .TextMatrix(0, col_SchList_����) = "����"
        .TextMatrix(0, col_SchList_���) = "���"
        .TextMatrix(0, col_SchList_ҽ������) = "ҽ������"
        .TextMatrix(0, col_SchList_ԤԼ��ʼʱ��) = "��ʼʱ��"
        .TextMatrix(0, col_SchList_ԤԼ����ʱ��) = "����ʱ��"
        .TextMatrix(0, col_SchList_��������) = "��������"
        .TextMatrix(0, col_SchList_ִ�й���) = "ִ�й���"
        .TextMatrix(0, col_SchList_���ע��) = "���ע��"
        .TextMatrix(0, col_SchList_�ֻ���) = "�ֻ���"
        '�����ݿ��������
        If rsTemp.EOF = False Then
            For i = 1 To rsTemp.RecordCount
                If rsTemp!Ӥ�� <> 0 Then
                    strSQL = "Select A.����ʱ��,Nvl(B.Ӥ������, A.���� || '֮��' || Trim(To_Char(B.���, '9'))) As Ӥ������, B.Ӥ���Ա�, B.����ʱ��" & vbNewLine & _
                                 "  From ����ҽ����¼ A, ������������¼ B " & vbNewLine & _
                                 "  Where a.����ID = b.����ID  And b.��� = [2] And a.ID = [1]"
                    Set rsBaby = zlDatabase.OpenSQLRecord(strSQL, "��ȡӤ����Ϣ", CLng(rsTemp!ҽ��ID), CLng(rsTemp!Ӥ��))
                    .TextMatrix(i, col_SchList_����) = rsBaby!Ӥ������
                Else
                    .TextMatrix(i, col_SchList_����) = rsTemp!����
                End If
                
                .TextMatrix(i, col_SchList_ID) = rsTemp!ID
                .TextMatrix(i, col_SchList_���) = rsTemp!���
                .TextMatrix(i, col_SchList_ҽ������) = rsTemp!ҽ������
                .TextMatrix(i, col_SchList_ԤԼ��ʼʱ��) = Format(rsTemp!ԤԼ��ʼʱ��, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, col_SchList_ԤԼ����ʱ��) = Format(rsTemp!ԤԼ����ʱ��, "YYYY-MM-DD HH:MM:SS")
                .TextMatrix(i, col_SchList_��������) = NVL(rsTemp!��������)
                .TextMatrix(i, col_SchList_�ֻ���) = NVL(rsTemp!�ֻ���)
                .TextMatrix(i, col_SchList_ִ�й���) = IIf(NVL(rsTemp!ִ�й���, 0) = -1, "�Ѳ���", IIf(NVL(rsTemp!ִ�й���, 0) = 0 Or NVL(rsTemp!ִ�й���, 0) = 1, "�ѵǼ�", IIf(NVL(rsTemp!ִ�й���, 0) = 2, "�ѱ���", IIf(NVL(rsTemp!ִ�й���, 0) = 3, "�Ѽ��", IIf(NVL(rsTemp!ִ�й���, 0) = 4, "�ѱ���", IIf(NVL(rsTemp!ִ�й���, 0) = 5, "�����", IIf(NVL(rsTemp!ִ�й���, 0) = 6, "�����", "δ���")))))))
                'ִ�й���: -1-���أ�0��1-�ѵǼǣ�2-�ѱ�����3-�Ѽ�飻4-�ѱ��棻5-����ˣ�6-�����
                .TextMatrix(i, col_SchList_���ע��) = NVL(rsTemp!���ע��)
                
                '������ɫ
                 '������ɫ
                If Not (NVL(rsTemp!ִ�й���, 0) = 0 Or NVL(rsTemp!ִ�й���, 0) = 1 Or NVL(rsTemp!ִ�й���, 0) = 2) Then
                    lngColor = mlngColorLblDone
                ElseIf Format(rsTemp!ԤԼ��ʼʱ��, "YYYY-MM-DD HH:MM:SS") < Format(Now, "YYYY-MM-DD HH:MM:SS") Then
                    lngColor = mlngColorLblPassed
                Else
                    lngColor = mlngColorLblWaiting
                End If
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = lngColor
                
                rsTemp.MoveNext
            Next i
        End If
        
        '���غ�̨������
        .ColHidden(col_SchList_ID) = True
        
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshCalendar()
'------------------------------------------------
'���ܣ�ˢ������
'������
'���أ���
'------------------------------------------------
    
    On Error GoTo err
    
    mstrSchRestDate = RefeshSchRestDay(mlngOrderID, mlngSchDeviceID, dpCalendar.LastVisibleDay)
    
    dpCalendar.RedrawControl
    
    'ˢ��ԤԼ����ͼ
    Call LoadWeekView
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadWeekView()
'------------------------------------------------
'���ܣ�����ԤԼ������ͼ
'������
'���أ���
'------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim dtMonday As Date
    Dim lngCapacity As Long
    Dim lngScheduledCount As Long
    Dim lngVacancy As Long
    
    On Error GoTo err
    
    With vsfWeekView
        .Clear
        .Cols = 4
        .Rows = 8
        .FixedRows = 1
        .FixedCols = 1
        .AllowUserResizing = flexResizeColumns
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .ScrollBars = flexScrollBarVertical
        .CellAlignment = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .ExtendLastCol = True
        .ColWidthMin = 300
        
        .RowHeight(0) = 300
        For i = 1 To .Rows - 1
            .RowHeight(i) = 550
        Next i
        
        .ColWidth(0) = 1000
        
        .TextMatrix(0, col_WeekView_����) = "����"
        .TextMatrix(0, col_WeekView_����) = "����"
        .TextMatrix(0, col_WeekView_��ԤԼ) = "��ԤԼ"
        .TextMatrix(0, col_WeekView_������) = "������"
        
        '���ҵ�����һ
        dtMonday = mschDate - Weekday(mschDate, vbMonday) + 1
        
        '�����ݿ��������
        For i = 1 To 7
            .TextMatrix(i, col_WeekView_����) = "��" & WeekdayChinese(CLng(i)) & vbCrLf & Format(dtMonday + i - 1, "M��D��")
            .RowData(i) = dtMonday + i - 1
            
            If DateScheduleInfo(mlngOrderID, dtMonday + i - 1, mlngSchDeviceID, lngCapacity, lngScheduledCount, lngVacancy) = True Then
                .Cell(flexcpBackColor, i, col_WeekView_����) = vbGreen
                .TextMatrix(i, col_WeekView_����) = lngVacancy
            Else
                .Cell(flexcpBackColor, i, col_WeekView_����) = vbRed
                .TextMatrix(i, col_WeekView_����) = 0
            End If
            
            .TextMatrix(i, col_WeekView_��ԤԼ) = lngScheduledCount
            .TextMatrix(i, col_WeekView_������) = lngCapacity
            
            'ѡ�е���
            If mschDate = dtMonday + i - 1 Then
                .Select i, 0, i, .Cols - 1
            End If
        Next i
        
    End With
    
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function WeekdayChinese(lngWeekday As Long) As String
'------------------------------------------------
'���ܣ������ֵ�����ţ����������ĵ���һ������
'������ lngWeekday -- ����ţ�1-7
'���أ��ܵ������ַ���һ�����������ġ��塢������
'------------------------------------------------
    On Error GoTo err
    
    Select Case Val(lngWeekday)
        Case 1:
            WeekdayChinese = "һ"
        Case 2:
            WeekdayChinese = "��"
        Case 3:
            WeekdayChinese = "��"
        Case 4:
            WeekdayChinese = "��"
        Case 5:
            WeekdayChinese = "��"
        Case 6:
            WeekdayChinese = "��"
        Case 7:
            WeekdayChinese = "��"
        Case Else
            WeekdayChinese = "��"
    End Select
    
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function DateScheduleInfo(ByVal lngOrderID As Long, ByVal dtDate As Date, _
    ByVal lngDeviceID As Long, ByRef lngDayCapacity As Long, ByRef lngScheduledCount As Long, _
    ByRef lngVacancy As Long) As Boolean
'-----------------------------------------------------------
'����:��ȡ�����ԤԼ���
'���:  lngOrderID -- ҽ��ID
'       dtDate -- ����
'       lngSchDeviceID -- ԤԼ�豸ID
'       lngDayCapacity -- ԤԼ������
'       lngScheduledCount -- ��ԤԼ����
'       lngVacancy -- ʣ������
'����: True -- ��ԤԼ��False -- ����ԤԼ
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngPlanID As Long
    Dim dtStartTime As Date

    On Error GoTo err
    
    lngDayCapacity = 0
    lngScheduledCount = 0
    lngVacancy = 0
    
    lngPlanID = schTimeTable.GetSchPlanID(lngDeviceID, dtDate, False, True)
    
    If lngPlanID <> 0 Then
        strSQL = "Select nvl(Sum(b.ԤԼ����), 0) as thecount From Ӱ��ԤԼʱ��ƻ� B Where b.ԤԼ����id =[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡԤԼ����", lngPlanID)
        
        If rsTemp.EOF = False Then
            lngDayCapacity = rsTemp!thecount
        End If
            
        strSQL = "Select Count(A.ID) as thecount From Ӱ��ԤԼ��¼ A Where a.ԤԼ�豸id = [1] And " _
            & " a.ԤԼ��ʼʱ�� Between to_date(to_char([2], 'yyyy-mm-dd') || ' 00:00:01', 'yyyy-mm-dd hh24:mi:ss') And " _
            & " to_date(to_char([2], 'yyyy-mm-dd') || ' 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ԤԼ����", lngDeviceID, CDate(Format(dtDate, "yyyy-MM-dd 00:00:00")))
        
        If rsTemp.EOF = False Then
            lngScheduledCount = rsTemp!thecount
        End If
        
        If Format(dtDate, "YYYY-MM-DD") = Format(Now, "YYYY-MM-DD") Then
            '���죬ȡ��ǰʱ��2Сʱ֮���ԤԼʱ���
            strSQL = " Select a.id, a.��ʼʱ��, a.����ʱ��,a.ԤԼ���� From Ӱ��ԤԼʱ��ƻ� A " _
                    & " Where a.ԤԼ����id = [1] and " _
                    & " to_char(a.����ʱ��, 'hh24:mi:ss') > to_char(sysdate + 2 / 24, 'hh24:mi:ss') " _
                    & " Order By a.��ʼʱ�� desc"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�����ԤԼ����", lngPlanID)
            
            While rsTemp.EOF = False
                lngVacancy = lngVacancy + rsTemp!ԤԼ����
                dtStartTime = rsTemp!��ʼʱ��
                rsTemp.MoveNext
            Wend
            
            strSQL = "Select Count(A.ID) as thecount From Ӱ��ԤԼ��¼ A Where a.ԤԼ�豸id = [1] And " _
                    & "  to_char(a.ԤԼ��ʼʱ��, 'hh24:mi:ss') > to_char([2], 'hh24:mi:ss') " _
                    & " And trunc(a.ԤԼ��ʼʱ��) = trunc(sysdate)"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��ԤԼ����", lngDeviceID, CDate(Format(dtStartTime, "yyyy-MM-dd hh:mm:ss")))
            
            If rsTemp.EOF = False Then
                lngVacancy = lngVacancy - rsTemp!thecount
            End If
        Else
            If InStr(mstrSchRestDate, Format(dtDate, "YYYY-MM-DD")) > 0 Then
                lngVacancy = 0
            Else
                lngVacancy = lngDayCapacity - lngScheduledCount
            End If
        End If
    End If
    
        DateScheduleInfo = (lngVacancy <> 0)
    Exit Function
err:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfWeekView_Click()
    On Error GoTo err
    
    If vsfWeekView.Rows > 1 Then
        mschDate = Format(vsfWeekView.RowData(vsfWeekView.RowSel), "YYYY-MM-DD")
        Call ChangeCalendar(mschDate)
        Call RefreshSchedule
    End If
    Exit Sub
err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
