VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStPathMain 
   Caption         =   "��׼·���ο�"
   ClientHeight    =   9435
   ClientLeft      =   3240
   ClientTop       =   1395
   ClientWidth     =   12765
   Icon            =   "frmStPathMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   12765
   WindowState     =   2  'Maximized
   Begin XtremeReportControl.ReportControl rptStPath 
      Height          =   6495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5100
      _Version        =   589884
      _ExtentX        =   8996
      _ExtentY        =   11456
      _StockProps     =   0
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "���Ҳ���(Ctrl+F)"
      Top             =   0
      Width           =   1155
   End
   Begin VB.PictureBox picStPathDetial 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   5300
      ScaleHeight     =   7095
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   270
      Width           =   7575
      Begin VB.PictureBox picPathTable 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   720
         ScaleHeight     =   2295
         ScaleWidth      =   6495
         TabIndex        =   3
         Top             =   360
         Width           =   6495
         Begin VB.Frame fraTableTile 
            BackColor       =   &H00F0F4E4&
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   6255
            Begin VB.Label lblTableTile 
               AutoSize        =   -1  'True
               BackColor       =   &H00F0F4E4&
               Height          =   180
               Left            =   120
               TabIndex        =   7
               Top             =   0
               Width           =   90
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPathTable 
            Height          =   975
            Left            =   0
            TabIndex        =   4
            Top             =   1320
            Width           =   3585
            _cx             =   6324
            _cy             =   1720
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   3
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   20
            RowHeightMax    =   5000
            ColWidthMin     =   100
            ColWidthMax     =   12000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStPathMain.frx":058A
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
      Begin XtremeSuiteControls.TabControl tbcStPath 
         Height          =   7335
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6615
         _Version        =   589884
         _ExtentX        =   11668
         _ExtentY        =   12938
         _StockProps     =   64
      End
      Begin VB.PictureBox picPathCourse 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   240
         ScaleHeight     =   3735
         ScaleWidth      =   6015
         TabIndex        =   8
         Top             =   3000
         Width           =   6015
         Begin RichTextLib.RichTextBox rtfPathCourse 
            Height          =   4095
            Left            =   120
            TabIndex        =   9
            Top             =   60
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   7223
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmStPathMain.frx":05F6
         End
      End
   End
   Begin VB.Frame fraSplit 
      Caption         =   "Frame1"
      Height          =   7335
      Left            =   5200
      MousePointer    =   9  'Size W E
      TabIndex        =   0
      Top             =   0
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTemp 
      Height          =   900
      Left            =   7200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7320
      Visible         =   0   'False
      Width           =   1080
      _cx             =   1905
      _cy             =   1587
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeSuiteControls.TabControl tbcPathName 
      Height          =   975
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   2535
      _Version        =   589884
      _ExtentX        =   4471
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblFind 
      Caption         =   "·������"
      Height          =   220
      Left            =   1440
      TabIndex        =   11
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmStPathMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrs�� As New ADODB.Recordset '��׼·��������
Private mrs��ͷ��Ϣ As New ADODB.Recordset '��׼·�����ı�ͷ�Լ��׶�����������Ŀ����Ϣ
Private mlngStPathID As Long 'ѡ�еı�׼·����ID
Private Const M_INT_STEPNUM = 3 '�̶���ʾ�׶���
Private mstrTilePos As String '��׼·��������·�����̶���Ŀ�ʼλ�ã���ʽΪ������1��ʼλ�ã�0��,����2������3

'��ö��
Private Enum PathListCols
    COL_ID = 0
    COL_�������� = 1
    COL_���� = 2
    COL_·������ = 3
    COL_�汾˵�� = 4
    COL_�������� = 5
    COL_�������� = 6
End Enum
'��ǰ��ý���Ŀؼ�ö��
Private Enum FocusContrl
    FC_PathList = 0
    FC_Course = 1
    FC_TBC = 2
    FC_Table = 3
End Enum
Private mintFocus As Integer '��ǰ��ý���Ŀؼ�

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long

    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
        Case conMenu_File_Preview
            Call zlRptPrint(0)
        Case conMenu_File_Print
            Call zlRptPrint(1)
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_View_ToolBar_Button '������
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            Control.Checked = Not Control.Checked
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    If objControl.ID = conMenu_Help_Help Or objControl.ID = conMenu_File_Exit Or objControl.ID = conMenu_File_Print Or objControl.ID = conMenu_File_Preview Then
                        objControl.Style = xtpButtonIcon
                    Else
                        objControl.Style = IIf(Control.Checked, xtpButtonIconAndCaption, xtpButtonIcon)
                    End If
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    rptStPath.SelectedRows(0).Expanded = False
                ElseIf Not rptStPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptStPath.SelectedRows(0).ParentRow.GroupRow Then
                        rptStPath.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '���۵���λ��������,�����Զ�������¼�
            Call rptStPath_SelectionChanged
        Case conMenu_View_Expend_CurExpend 'չ����ǰ��
            If rptStPath.SelectedRows.Count > 0 Then
                rptStPath.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse '�۵�������
            For Each objRow In rptStPath.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '���۵���λ��������,�����Զ�������¼�
            Call rptStPath_SelectionChanged
        Case conMenu_View_Expend_AllExpend 'չ��������
            For Each objRow In rptStPath.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Find '����
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus '��ʱ��Ҫ��λһ��
                If txtFind.Text <> "" Then
                    Call FindPath(False)
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext '������һ��
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call FindPath(True)
            End If
        '·��Ŀ¼��ɾ��
        Case conMenu_Edit_NewPath
            Call InsertItem(0)
        Case conMenu_Edit_ModifyPath
            Call ModItem(0)
        Case conMenu_Edit_DelPath
            Call DeleteItem(0, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '·�����̶�����ɾ��
        Case conMenu_Edit_NewCourseItem
            Call InsertItem(1)
        Case conMenu_Edit_ModifyCourseItem
            Call ModItem(1)
        Case conMenu_Edit_DelCourseItem
            Call DeleteItem(1, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '·������ɾ��
        Case conMenu_Edit_NewTable
            Call InsertItem(2)
        Case conMenu_Edit_ModifyTable
            Call ModItem(2)
        Case conMenu_Edit_DelTable
            Call DeleteItem(2, Val(rptStPath.SelectedRows(0).Record(COL_ID).Value))
        '�������޸�
        Case conMenu_Edit_ModifyTableContent
            Call ModItem(3)
            
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 10))
        Case conMenu_File_Exit '�˳�
            Unload Me
        Case conMenu_File_ImportPathTable
            If frmImportPath.ShowMe(Me, mlngStPathID) = True Then
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngTop As Long, lngLeft As Long, lngBottom As Long, lngRight As Long
    
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    tbcPathName.Left = lngLeft
    tbcPathName.Top = lngTop
    tbcPathName.Height = lngBottom - lngTop
    tbcPathName.Width = (lngRight - lngLeft) * 0.3
    
    fraSplit.Top = lngTop
    fraSplit.Left = tbcPathName.Left + tbcPathName.Width + 30
    fraSplit.Height = lngBottom - lngTop
    
    picStPathDetial.Top = lngTop
    picStPathDetial.Left = fraSplit.Left + fraSplit.Width + 30
    picStPathDetial.Width = lngRight - picStPathDetial.Left
    picStPathDetial.Height = lngBottom - lngTop
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean, blnParent As Boolean
    
    '���ݵ�ǰ��Ŀؼ���ÿ�����
    blnEnabled = True
    
    Select Case mintFocus
        Case FC_Course
            If rtfPathCourse.SelStart >= Len(rtfPathCourse.Text) And Len(rtfPathCourse.Text) <> 0 Or Len(rtfPathCourse.Text) = 0 Then
                blnEnabled = False
            End If
        Case FC_TBC
            If tbcStPath.Selected.Index = 0 Then blnEnabled = False
        Case FC_PathList
            blnEnabled = True
        Case FC_Table
            If tbcStPath.Selected.Index = 0 Then blnEnabled = False
    End Select
    If rptStPath.Rows.Count <> 0 Then
        If rptStPath.SelectedRows.Count <> 0 Then
            blnParent = rptStPath.SelectedRows(0).GroupRow
        End If
    Else
        blnParent = True
    End If
    
    blnEnabled = blnEnabled And Not blnParent
    
    Select Case Control.ID
        '·��Ŀ¼��ɾ��
        Case conMenu_Edit_NewPath
            Control.Enabled = True
        Case conMenu_Edit_ModifyPath
            Control.Enabled = Not blnParent
        Case conMenu_Edit_DelPath
            Control.Enabled = Not blnParent
        '·�����̶�����ɾ��
        Case conMenu_Edit_NewCourseItem
            Control.Enabled = Not blnParent
        Case conMenu_Edit_ModifyCourseItem
            Control.Enabled = (mintFocus = FC_Course) And blnEnabled
        Case conMenu_Edit_DelCourseItem
            Control.Enabled = (mintFocus = FC_Course) And blnEnabled
        '·������ɾ��
        Case conMenu_Edit_NewTable
            Control.Enabled = Not blnParent
        Case conMenu_Edit_ModifyTable
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled
        Case conMenu_Edit_DelTable
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled And tbcStPath.ItemCount > 2
        '�������޸�
        Case conMenu_Edit_ModifyTableContent
            Control.Enabled = (mintFocus = FC_Table Or mintFocus = FC_TBC) And blnEnabled
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_Expend_CurExpend 'չ����ǰ��
            blnEnabled = False
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptStPath.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
            blnEnabled = False
            If rptStPath.SelectedRows.Count > 0 Then
                If rptStPath.SelectedRows(0).GroupRow Then
                    blnEnabled = rptStPath.SelectedRows(0).Expanded
                ElseIf Not rptStPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptStPath.SelectedRows(0).ParentRow.GroupRow Then
                        blnEnabled = rptStPath.SelectedRows(0).ParentRow.Expanded
                    End If
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend '�۵�/չ����
            Control.Enabled = rptStPath.GroupsOrder.Count > 0 And rptStPath.Rows.Count > 0
    End Select
End Sub

Private Sub Form_Load()

    mlngStPathID = 0
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = False
        .ShowTextBelowIcons = False
        .AlwaysShowFullMenus = True
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization True
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    'tbcPathName·���ο�
    With Me.tbcPathName
        With .PaintManager
            .Appearance = xtpTabAppearanceExcel
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "��ҽ�ο�", rptStPath.hwnd, 0
        .InsertItem 1, "��ҽ�ο�", rptStPath.hwnd, 0
        
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    '��ʼ��tbcControl
    With tbcStPath
    
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        
        .AllowReorder = False
        '���μ�������ֻ����ѡ��Լ���׼סԺ����
        Call .InsertItem(0, "��׼סԺ����", picPathCourse.hwnd, 0)
        .Item(0).Selected = True 'Ĭ��ѡ���׼סԺ����
        
    End With
    Call MainDefCommandBar
    '��ʼ����׼·���б�
    Call InitPathList
    '���ر�׼·��Ŀ¼
    Call LoadStPathList(tbcPathName.Selected.Index)
End Sub

Private Sub Form_Resize()
'���ܣ�����tbcPathNameh��picStPathDetial��λ�ô�С
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
        Me.Height = IIf(Me.Height < 9000, 9000, Me.Height)
        Me.Width = IIf(Me.Width < 12000, 12000, Me.Width)
    End If
    Call cbsMain_Resize
End Sub


Private Sub Form_Unload(Cancel As Integer)
'���ܣ����ģ�鼶����ֵ
    Set mrs�� = Nothing
    Set mrs��ͷ��Ϣ = Nothing
    mlngStPathID = 0
    mstrTilePos = ""
    mintFocus = -1
    
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'���ܣ�ʵ�ֱ�׼·���嵥���׼·�����������϶���С

    If Button = 1 Then
        If tbcPathName.Width + x > 11000 Or tbcPathName.Width + x < 2000 Then Exit Sub
        
        fraSplit.Left = fraSplit.Left + x
        tbcPathName.Width = fraSplit.Left - 30 - tbcPathName.Left
        
        picStPathDetial.Left = fraSplit.Left + 30
        picStPathDetial.Width = Me.ScaleWidth - picStPathDetial.Left
        
        Me.Refresh
    End If
    
End Sub

Private Sub picPathCourse_Resize()
'���ܣ�ʵ�ֱ�׼·���������ݵĴ�С����

    rtfPathCourse.Width = picPathCourse.Width - rtfPathCourse.Left - 120
    rtfPathCourse.Height = picPathCourse.Height - rtfPathCourse.Top
    
End Sub

Private Sub picPathTable_Resize()
'���ܣ����ñ���ͷ������ڿؼ���λ�����С

    fraTableTile.Height = lblTableTile.Height + 60
    fraTableTile.Width = picPathTable.Width
    lblTableTile.Width = fraTableTile.Width
    lblTableTile.Width = fraTableTile.Width - lblTableTile.Left
    
    vsPathTable.Top = fraTableTile.Top + fraTableTile.Height
    vsPathTable.Height = picPathTable.Height - vsPathTable.Top
    vsPathTable.Width = picPathTable.Width - vsPathTable.Left
    
End Sub

Private Sub picStPathDetial_Resize()
'���ܣ���׼·���������Ĵ�С����

    tbcStPath.Top = 0
    tbcStPath.Left = 0
    tbcStPath.Width = picStPathDetial.Width
    tbcStPath.Height = picStPathDetial.Height
    picPathTable.Width = tbcStPath.Width
    picPathCourse.Width = tbcStPath.Width
    
End Sub

Private Sub rptStPath_GotFocus()
    mintFocus = FC_PathList
End Sub

Private Sub rptStPath_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBar
    
    mintFocus = FC_PathList

    If Button = 2 Then
        Set objHitTest = rptStPath.HitTest(x, y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
            With objPopup.Controls
                 .Add xtpControlButton, conMenu_Edit_NewPath, "����·��(&A)"
                 .Add xtpControlButton, conMenu_Edit_ModifyPath, "�޸�·��(&Q)"
                 .Add xtpControlButton, conMenu_Edit_DelPath, "ɾ��·��(&W)"
            End With
        End If
        
        rptStPath.SetFocus
        If Not objPopup Is Nothing Then objPopup.ShowPopup
    End If
    
End Sub

Private Sub rptStPath_SelectionChanged()
'���ܣ�����ѡ���·��ID,������ID���ر�׼·�������Լ���

    mintFocus = FC_PathList
    If rptStPath.Rows.Count <> 0 Then
        If Not rptStPath.SelectedRows(0).GroupRow Then
            If mlngStPathID <> Val(rptStPath.SelectedRows(0).Record.Tag) And Val(rptStPath.SelectedRows(0).Record.Tag) <> 0 Then
                mlngStPathID = Val(rptStPath.SelectedRows(0).Record.Tag)
                Call LoadPathByID(mlngStPathID, True, 0)
            End If
        End If
    End If
End Sub

Private Sub rtfPathCourse_GotFocus()
    mintFocus = FC_Course
End Sub

Private Sub rtfPathCourse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar

    mintFocus = FC_Course
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_NewCourseItem, "��������(&Z)"
            .Add xtpControlButton, conMenu_Edit_ModifyCourseItem, "�޸Ķ���(&U)"
            .Add xtpControlButton, conMenu_Edit_DelCourseItem, "ɾ������(&D)"
        End With
        
        rtfPathCourse.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub tbcPathName_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'����:����ѡ�ѡ���������
    If Me.Visible Then
        mlngStPathID = 0
        Call LoadStPathList(Item.Index)
        Call LoadPathByID(mlngStPathID, True, 0)
    End If
End Sub

Private Sub tbcStPath_GotFocus()
    mintFocus = FC_TBC
End Sub

Private Sub tbcStPath_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar
    
    mintFocus = FC_TBC
    
    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_NewTable, "������(&K)"
            .Add xtpControlButton, conMenu_Edit_ModifyTable, "�޸ı�(&M)"
            .Add xtpControlButton, conMenu_Edit_DelTable, "ɾ����(&Y)"
        End With
        
        tbcStPath.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Sub tbcStPath_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'���ܣ�ѡ���ʱ���ر�����

    If Me.Visible Then
        Call LoadPathByID(mlngStPathID, False, Item.Index)
        mintFocus = FC_TBC
        picPathCourse.Visible = Item.Index = 0
        picPathTable.Visible = Item.Index <> 0
    End If
    
End Sub

Private Sub LoadStPathList(ByVal lngIndex As Long)
'���ܣ����ر�׼·��Ŀ¼
'����:lngIndex 0-��ҽ�ο�,1-��ҽ�ο�

    Dim objRecord     As ReportRecord
    Dim objItem       As ReportRecordItem
    Dim i As Long, strDept As String
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = "Select a.Id, a.��������, a.����, a.·������, a.�汾˵��,��������,�������� " & vbNewLine & _
        "From   ��׼·��Ŀ¼ A,��׼·������ B where A.ID=B.��׼·��ID And Nvl(a.���,0)=[1] " & vbNewLine & _
        " order by ��������,���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngIndex)
    rptStPath.Records.DeleteAll
    For i = 0 To rsTemp.RecordCount - 1
        Set objRecord = rptStPath.Records.Add
        Set objItem = objRecord.AddItem(rsTemp!ID & "")
        Set objItem = objRecord.AddItem(rsTemp!�������� & "")
        Set objItem = objRecord.AddItem(rsTemp!���� & "")
        Set objItem = objRecord.AddItem(rsTemp!·������ & "")
        Set objItem = objRecord.AddItem(rsTemp!�汾˵�� & "")
        Set objItem = objRecord.AddItem(rsTemp!�������� & "")
        Set objItem = objRecord.AddItem(rsTemp!�������� & "")
        objRecord.Tag = CStr(rsTemp!ID)
        rsTemp.MoveNext
    Next
    rptStPath.Populate
    
    For i = 0 To rptStPath.Rows.Count - 1
        If mlngStPathID = 0 Then
            If rptStPath.Rows(i).GroupRow = False Then
                rptStPath.Rows(i).Selected = True
                mlngStPathID = Val(rptStPath.Rows(i).Record(COL_ID).Value)
                rptStPath.SelectedRows(0).ParentRow.Expanded = True
                Exit For
            End If
        Else
            If rptStPath.Rows(i).GroupRow = False Then
                If mlngStPathID = Val(rptStPath.Rows(i).Record(COL_ID).Value) Then
                    rptStPath.Rows(i).Selected = True
                    rptStPath.SelectedRows(0).ParentRow.Expanded = True
                    Exit For
                End If
            End If
        End If
    Next
    
    '�����Զ�����LoadPathByID(mlngStPathID, True)
    Call LoadPathByID(mlngStPathID, True)
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub LoadPathByID(ByVal lngId As Long, Optional ByVal blnReadData As Boolean, Optional ByVal lng��� As Long)
'���ܣ�����ѡ��ı�׼·��ID��ȡ���ݣ������ݱ���ż���·�����̣�·����������ͷ
'������lngID   ѡ���·��ID
'      blnReadData �Ƿ��ȡ��׼·����Ϣ���ڱ�׼·�����μ��ػ��߱�׼·���л�ʱ�Ǿ���Ҫ��ȡ��
'      lng���  0 ��׼·�����̣�1 ��1��2����2...
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, k As Long
    Dim strSql As String, strFilter As String
    Dim strTilePos As String '��¼������λ�ø�ʽΪ������1��ʼλ�ã�����;����2��ʼλ�ã�����
    Dim lngColCount As Long, lng������ As Long, lngBeginRow As Long
    Dim lngRowCount As Long
    Dim strContent As String
    Dim strB As String
    Dim arrTemp() As String
    Dim n As Integer, intPos As Integer
    
    On Error GoTo errH
    
    If blnReadData Then
        'ɾ��ѡ������vs����
        vsPathTable.Delete
        For i = tbcStPath.ItemCount - 1 To 1 Step -1
            tbcStPath.RemoveItem (i)
        Next
        
        '���ر�׼סԺ����
        rtfPathCourse.Visible = False
        rtfPathCourse.Text = ""
        strSql = "Select ����, ���� From ��׼·������ Where ��׼·��id = [1] Order By ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        mstrTilePos = ""
        strB = "��,��,��,��,��,��,��,��,��,��,��"
        arrTemp = Split(strB, ",")
        If rsTmp.RecordCount <> 0 Then
            For i = 1 To rsTmp.RecordCount
                strTilePos = strTilePos & ";" & Len(strContent) & "," & Len(rsTmp!����) & "," & n
                mstrTilePos = mstrTilePos & "," & Len(strContent)
                n = 0
                For j = LBound(arrTemp) To UBound(arrTemp)
                    Do
                        intPos = InStr(intPos + 1, "" & rsTmp!����, arrTemp(j))
                        If intPos = 0 Then Exit Do
                        n = n + 1
                        DoEvents
                    Loop
                Next
                strContent = strContent & rsTmp!���� & vbNewLine & vbNewLine & rsTmp!���� & vbNewLine & vbNewLine
                rsTmp.MoveNext
            Next
            rtfPathCourse.Text = strContent
            mstrTilePos = Mid(mstrTilePos, 2)
        End If
               
        
        Call SetStPathCourceFont(Mid(strTilePos, 2)) '��������
        rtfPathCourse.Visible = True
        
        '��ȡ��������Ϣ
        strSql = "Select a.����� �����, b.������, b.����ͷ, a.����, a.����" & vbNewLine & _
                "From (Select �����, Max(�������) ����, Max(�׶����) ���� From ��׼·���� Where ��׼·��id = [1] Group By �����) A, ��׼·���� B" & vbNewLine & _
                "Where b.��׼·��id =[1] And a.����� = b.����� And b.������� = 1 And b.�׶���� = 1" & vbNewLine & _
                "Order By �����"

        Set mrs��ͷ��Ϣ = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        
        '���ر�׼·����ѡ��
        If mrs��ͷ��Ϣ.RecordCount > 0 Then
            j = mrs��ͷ��Ϣ.RecordCount
            For i = 1 To j
                mrs��ͷ��Ϣ.Filter = "����� =" & i
                Call tbcStPath.InsertItem(i, mrs��ͷ��Ϣ!������, picPathTable.hwnd, 0)
            Next
            '��ȡ������
            strSql = "Select  �����, ������, ����ͷ, �������, ��������, �׶����, �׶�����, ·������" & vbNewLine & _
                "From   ��׼·����" & vbNewLine & _
                "where ��׼·��id=[1]"
            Set mrs�� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngId)
        End If
    End If
    
    If lng��� <> 0 Then
        'û�б���Ϣ�����򲻼��ر���Ϣ
        
        mrs��.Filter = ""
        If mrs��.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        mrs��ͷ��Ϣ.Filter = " ����� =" & lng���
        If mrs��ͷ��Ϣ.RecordCount = 0 Then tbcStPath.Item(0).Selected = True: Exit Sub
        
        '���ر���ͷ
        lblTableTile.Caption = ""
        lblTableTile.Caption = vbNewLine & mrs��ͷ��Ϣ!����ͷ
        
        With vsPathTable
            .Redraw = False
            .Rows = 0
            .Cols = 0
            'ȷ������
            lngColCount = Val(mrs��ͷ��Ϣ!���� & "")
            'ȷ��������
            lng������ = Val(mrs��ͷ��Ϣ!���� & "")
            lngRowCount = IntEx(lngColCount / (M_INT_STEPNUM + 1)) * lng������ + IntEx(lngColCount / (M_INT_STEPNUM + 1)) - 1
            If lngRowCount = 1 And lngColCount = 1 Then
                .Rows = 0
                .Cols = 0
                Call SetVsStyle
                Call picPathTable_Resize '����lblTableTile��autoSize�������Ҫ����resize
                tbcStPath.Item(lng���).Selected = True
                Exit Sub
            Else
                .Rows = lngRowCount
                .Cols = IIf(lngColCount > M_INT_STEPNUM, M_INT_STEPNUM + 1, lngColCount)
            End If
    
    
            For k = 1 To IntEx(lngColCount / (M_INT_STEPNUM + 1))
                lngBeginRow = (k - 1) * lng������ + (k - 1)
                For i = lngBeginRow To lngBeginRow + lng������ - 1
                    For j = 0 To .Cols - 1
                        'ÿ�����������ĵ�һ����Ԫ��Ϊʱ��
                        If i = lngBeginRow And j = 0 Then
                            .TextMatrix(i, j) = "ʱ��"
                        Else
                            If Not (i = lngBeginRow Or j = 0) Then
                                strFilter = "�����=" & lng��� & " and �������=" & i - lngBeginRow + 1 & " and �׶����=" & (k - 1) * 3 + j + 1
                                mrs��.Filter = strFilter
                                If mrs��.RecordCount = 1 Then
                                    .TextMatrix(i, j) = Nvl(mrs��!·������, " ")
                                    .TextMatrix(i, 0) = Replace(Replace(Replace(mrs��!�������� & "", Chr(13), ""), Chr(10), ""), " ", "")
                                    .TextMatrix(lngBeginRow, j) = mrs��!�׶����� & ""
                                End If
                            End If
                        End If
                    Next
                Next
            Next
            
            Call SetVsStyle
            .Redraw = True
            Call picPathTable_Resize '����lblTableTile��autoSize�������Ҫ����resize
        
            
        End With
    End If
    
    tbcStPath.Item(lng���).Selected = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStPathCourceFont(ByVal strTilePos As String)
'���ܣ���RichTextBox������������
'���� strTilePos ��¼������λ�ø�ʽΪ������1��ʼλ�ã����ⳤ��;����2��ʼλ�ã����ⳤ��
    Dim arrTmp As Variant, i As Long, j As Long
    
    On Error Resume Next
    If Len(Trim(strTilePos)) = 0 Then Exit Sub
    arrTmp = Split(Trim(strTilePos), ";")
    
    With rtfPathCourse
        For i = LBound(arrTmp) To UBound(arrTmp)
            If Val(Split(arrTmp(i), ",")(2)) = 0 Then
                .SelStart = Split(arrTmp(i), ",")(0)
            Else
                j = i
                Exit For
            End If
            .SelLength = Split(arrTmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "����"
            .SelBold = True
            .SelLength = 0
        Next
        For i = j To UBound(arrTmp)
            .SelStart = Val(Split(arrTmp(i), ",")(0)) - Val(Split(arrTmp(j), ",")(2))
            .SelLength = Split(arrTmp(i), ",")(1)
            .SelFontSize = 14
            .SelFontName = "����"
            .SelBold = True
            .SelLength = 0
        Next
        .SelStart = 0 '����ƶ�����ʼ
        
    End With
    
End Sub

Private Sub SetVsStyle()
'���ܣ������������ñ����ĵ�Ԫ��ĸ߶�����,�Լ�������ɫ�ȣ��Լ���Ԫ��ĺϲ���

    Dim i As Long, j As Long
    Dim lngmaxHeight As Long
    
    
   On Error GoTo errH
    With vsPathTable
        If .Rows = 0 And .Cols = 0 Then Exit Sub
        '�޸ķ������ƣ��׶Σ�����Ӵ־���
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, 0) = 4 '����
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &HE1FFE1
        
        .AutoResize = False
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1, False, 0) '�Զ�������С
        '���ý׶����壬��ɫ�����뷽ʽ
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = "ʱ��" Then
                .Cell(flexcpAlignment, i, 0, i, .Cols - 1) = 4
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = False '���üӴ�ǰҪ������Ӵ�
                .Cell(flexcpFontBold, i, 0, i, .Cols - 1) = True
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE1FFE1
            Else
                If .Cols > 1 Then
                    .Cell(flexcpAlignment, i, 1, i, .Cols - 1) = 0
                End If
            End If
        Next
        
        '��ȡͬһ����ߵĵ�Ԫ��߶ȸ�ֵ���и�
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) <> "" Then
                For j = 0 To .Cols - 1
                    If j = 0 Then
                        lngmaxHeight = ComputerLines(.TextMatrix(i, j))
                    Else
                        lngmaxHeight = IIf(lngmaxHeight > ComputerLines(.TextMatrix(i, j)), lngmaxHeight, ComputerLines(.TextMatrix(i, j)))
                    End If
                Next
                .RowHeight(i) = IIf(lngmaxHeight = 0, 5, lngmaxHeight) * Me.TextHeight("��") * 1.5
            Else
                For j = 0 To .Cols - 1
                    .TextMatrix(i, j) = " " 'Ϊ�˺ϲ���Ԫ��
                Next
            End If
        Next
        '�ָ��е�Ԫ��ϲ����Լ��߿���ɫ����
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            If .TextMatrix(i, 0) = " " Then
                Call .CellBorderRange(i, 0, i, .Cols - 1, &HFFFFFF, 1, 0, 1, 0, 1, 0)
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HFFFFFF
                .MergeRow(i) = True
            End If
        Next
        
        For i = 1 To .Cols - 1
            .ColWidth(i) = 4000
        Next
        .ColWidth(0) = 1500
        'ʵ�������϶��п�
        .FixedRows = 1
        Call .CellBorderRange(0, 0, 0, .Cols - 1, &H8000&, 0, 0, 1, 1, 1, 1)
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function ComputerLines(ByVal strInput As String) As Long
'���ܣ����������ı��лس����ĸ���
'������  strInput   Ҫ����س������ַ���
'���أ�   �س����ĸ���

    Dim strTmp As String
    Dim Count  As Long, lngPos As Long, lngLen As Long
    
    lngPos = InStr(strInput, Chr(13))
    lngLen = Len(strInput)
    strTmp = strInput
    
    Do While lngPos <> 0
        If Trim(strTmp) = "" Then Exit Do
        If lngPos + 1 <= lngLen Then
            strTmp = Mid(strTmp, lngPos + 1)
            Count = Count + 1
            lngPos = InStr(strTmp, Chr(13))
            lngLen = Len(strTmp)
        End If
    Loop
    
    ComputerLines = Count + 2
    
End Function

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_ImportPathTable, "����·����")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewPath, "����·��(&A)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyPath, "�޸�·��(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelPath, "ɾ��·��(&W)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewCourseItem, "��������(&Z)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyCourseItem, "�޸Ķ���(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelCourseItem, "ɾ������(&D)")

        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "������(&K)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTable, "�޸ı�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelTable, "ɾ����(&Y)")
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTableContent, "�޸�����(&G)")
        objControl.BeginGroup = True
    End With

   Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False) '����
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)") '����
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False)
                objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False)
                objControl.Checked = True
            Set objControl = .Add(xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False)
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False): objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&I)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_FindNext, "������һ��(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True '����
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): objControl.BeginGroup = True
    End With
    
    '���������⴦��
    '-----------------------------------------------------
    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
        Set objControl = .Add(xtpControlLabel, 0, "����")
        objControl.IconId = conMenu_View_Find
        objControl.Flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Caption = ""
        objCustom.Handle = txtFind.hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With
    

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewPath, "����·��")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewCourseItem, "��������")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyCourseItem, "�޸Ķ���")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelCourseItem, "ɾ������")
            objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NewTable, "������")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTable, "�޸ı�")
            objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_Edit_DelTable, "ɾ����")
            objControl.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ModifyTableContent, "�޸�����")
            objControl.BeginGroup = True: objControl.Style = xtpButtonIconAndCaption
            
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    

    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewPath '����·��
        .Add FCONTROL, vbKeyM, conMenu_Edit_DelPath '�޸�·��
        .Add FSHIFT, vbKeyD, conMenu_Edit_DelPath 'ɾ��·��
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With
End Sub

Private Sub InitPathList()
'���ܣ���ʼ��·���б�
    Dim objCol        As ReportColumn
    
    With rptStPath
        '��ʼ��Report�ؼ�����������
        Set objCol = .Columns.Add(PathListCols.COL_ID, "ID", 20, False)
            objCol.Alignment = xtpAlignmentCenter: objCol.Resizable = True: objCol.AllowDrag = False: objCol.Visible = False
        Set objCol = .Columns.Add(PathListCols.COL_��������, "��������", 80, False)
            objCol.Resizable = True: objCol.Alignment = xtpAlignmentLeft: objCol.AllowDrag = False: objCol.TreeColumn = True: objCol.Groupable = True
        Set objCol = .Columns.Add(PathListCols.COL_����, "����", 50, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_·������, "·������", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_�汾˵��, "�汾˵��", 70, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_��������, "��������", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        Set objCol = .Columns.Add(PathListCols.COL_��������, "��������", 200, False)
            objCol.Alignment = xtpAlignmentLeft: objCol.Resizable = True: objCol.AllowDrag = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoItemsText = "û�п���ʾ�Ķ���."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        
        .GroupsOrder.Add rptStPath.Columns(COL_��������)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add rptStPath.Columns(COL_��������)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add rptStPath.Columns(COL_·������)
        .SortOrder(1).SortAscending = True
    '��λ��ѡ��ı�׼·��
    End With
End Sub

Private Function DeleteItem(ByVal lngDelType As Long, Optional ByVal lngStPathID As Long)
'���ܣ�ɾ���ض�����
'lngDelType 0-ɾ����׼·��
'           1-ɾ��·�����̶���
'           2-ɾ����׼·����
'lngStPathID ��׼·��ID
    Dim strSql As String
    
    On Error GoTo errH
    
    Select Case lngDelType
        Case 0
            If MsgBox("��ȷ��Ҫɾ��" & rptStPath.SelectedRows(0).Record(COL_·������).Value & "��", vbYesNo, gstrSysName) = vbYes Then
                strSql = "Zl_��׼·��Ŀ¼_Delete(" & mlngStPathID & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                If rptStPath.SelectedRows(0).ParentRow.Childs.Count > 1 Then
                    If Val(rptStPath.SelectedRows(0).ParentRow.Childs(0).Record(COL_ID).Value) <> mlngStPathID Then
                        mlngStPathID = Val(rptStPath.SelectedRows(0).ParentRow.Childs(0).Record(COL_ID).Value)
                    Else
                        mlngStPathID = Val(rptStPath.SelectedRows(1).ParentRow.Childs(0).Record(COL_ID).Value)
                    End If
                Else
                    mlngStPathID = 0
                End If
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
        Case 1
            If frmStPathItemEdit.ShowMe(Me, 2, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart)) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 2
            If MsgBox("��ȷ��Ҫɾ��" & tbcStPath.Selected.Caption & "��", vbYesNo, gstrSysName) = vbYes Then
                strSql = "Zl_��׼·����_Delete(" & mlngStPathID & "," & tbcStPath.Selected.Index & ")"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
                Call LoadPathByID(mlngStPathID, True, 0)
            End If
    End Select
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ModItem(ByVal lngDelType As Long, Optional ByVal lngStPathID As Long)
'���ܣ��޸��ض�����
'lngDelType 0-�޸ı�׼·��
'           1-�޸�·�����̶���
'           2-�޸ı�׼·����
'lngStPathID ��׼·��ID
    Dim strSql As String
    
    Select Case lngDelType
        Case 0
            With rptStPath.SelectedRows(0)
                If frmStPathEdit.ShowMe(Me, 1, Val(.Record(COL_ID).Value), .Record(COL_·������).Value, .Record(COL_����).Value, _
                    .Record(COL_��������).Value, .Record(COL_�汾˵��).Value, .Record(COL_��������).Value, .Record(COL_��������).Value, tbcPathName.Selected.Index) = True Then
                    Call LoadStPathList(tbcPathName.Selected.Index)
                End If
            End With
        Case 1
            If frmStPathItemEdit.ShowMe(Me, 1, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart)) = True Then
                Call LoadPathByID(mlngStPathID, False, tbcStPath.Selected.Index)
            End If
        Case 2
            If frmStTableEdit.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index, tbcStPath.Selected.Caption, lblTableTile.Caption) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 3
            If frmStTableContent.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
    End Select

End Function

Private Function InsertItem(ByVal lngDelType As Long)
'���ܣ������ض�����
'lngDelType 0-�����׼·��
'           1-����·�����̶���
'           2-�����׼·����
'lngStPathID ��׼·��ID
    Dim strSql As String, strDep As String
    
    Select Case lngDelType
        Case 0
            If rptStPath.Rows.Count > 0 Then
                If Not rptStPath.SelectedRows(0).GroupRow Then
                    strDep = rptStPath.SelectedRows(0).Record(COL_��������).Value
                Else
                    strDep = rptStPath.SelectedRows(0).Childs(0).Record(COL_��������).Value
                End If
            End If
            
            If frmStPathEdit.ShowMe(Me, 0, mlngStPathID, , , strDep, , , , tbcPathName.Selected.Index) = True Then
                Call LoadStPathList(tbcPathName.Selected.Index)
            End If
        Case 1
            If tbcStPath.Selected.Index <> 0 Then tbcStPath.Item(0).Selected = True
            If frmStPathItemEdit.ShowMe(Me, 0, mlngStPathID, GetCourseItemNo(rtfPathCourse.SelStart, True)) = True Then
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
        Case 2
            If frmStTableEdit.ShowMe(Me, mlngStPathID) = True Then
                tbcStPath.Item(tbcStPath.ItemCount - 1).Selected = True
                Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
            End If
    End Select

End Function

Private Sub txtFind_GotFocus()
'���ܣ���ý��㣬ȫѡ
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    
End Sub

Private Sub vsPathTable_DblClick()
'���ܣ������׼·�����༭����
    If frmStTableContent.ShowMe(Me, mlngStPathID, tbcStPath.Selected.Index) = True Then
        Call LoadPathByID(mlngStPathID, True, tbcStPath.Selected.Index)
    End If
End Sub

Private Sub vsPathTable_GotFocus()
    mintFocus = FC_Table
End Sub

Private Sub vsPathTable_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBar

    If Button = 2 Then
        Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
        With objPopup.Controls
            .Add xtpControlButton, conMenu_Edit_ModifyTableContent, "�޸�����(&G)"
        End With
        vsPathTable.SetFocus
        objPopup.ShowPopup
    End If
End Sub

Private Function GetCourseItemNo(ByVal lngSelStar As Long, Optional ByVal blnNew As Boolean = False) As Long
'���ܣ���õ�ǰ�������Ӧ��·�����̶������
'������lngSelStar ��ǰ���λ��
'      blnNew    �Ƿ�����������
'���أ� ������   ��ǰ�������
    Dim arrTmp As Variant, i As Long, lngNo As Long
    
    'û�ж���·�����̶���
    If Len(Trim(mstrTilePos)) = 0 Then GetCourseItemNo = 1: Exit Function
    '����ڿ�ʼλ��
    If lngSelStar = 0 Then GetCourseItemNo = 1: Exit Function
    
    arrTmp = Split(Trim(mstrTilePos), ",")
    '����һ�����̶���ʱ�Ĵ���
    If LBound(arrTmp) = UBound(arrTmp) Then GetCourseItemNo = IIf(blnNew, 2, 1): Exit Function
    '��������Ĵ���
    lngNo = 0
    For i = LBound(arrTmp) To UBound(arrTmp)
        If i < UBound(arrTmp) Then
            If lngSelStar >= Val(arrTmp(i)) And lngSelStar < Val(arrTmp(i + 1)) Then
                lngNo = IIf(blnNew, i + 2, i + 1): Exit For
            End If
        End If
    Next
    If lngNo = 0 Then lngNo = IIf(blnNew, UBound(arrTmp) + 2, UBound(arrTmp) + 1)
    
    GetCourseItemNo = lngNo
    
End Function

Private Sub FindPath(Optional ByVal blnNext As Boolean)
'���ܣ�����(��һ��������
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long
    Dim strInput As String
    Dim intType As Integer
    Dim lngRow As Long
    
    
    If Trim(txtFind.Text) = "" Then Exit Sub
     
    '��ʼ������
    With rptStPath
        If .SelectedRows.Count > 0 Then
            If Not .SelectedRows(0).GroupRow Then blnHave = True: lngRow = .SelectedRows(0).Index
        End If
        If Not blnNext Or blnReStart Or Not blnHave Then
            i = 0 'ReportControl����������0��ʼ
        Else
            i = .SelectedRows(0).Index + 1
        End If

        '����·��
        For i = i To .Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(COL_·������).Value Like "*" & Trim(txtFind.Text) & "*" Then Exit For
            End If
        Next
        
        If i <= .Rows.Count - 1 Then
            blnReStart = False
            '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
            .SetFocus
            Set .FocusedRow = .Rows(i)
            .Rows(i).ParentRow.Expanded = True
        Else
            blnReStart = True
            MsgBox IIf(blnNext, "������", "") & "�Ҳ������������ı�׼·����", vbInformation, gstrSysName
        End If
    
    End With
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    
    Set objReport = rptStPath
    strSubhead = "��׼·���嵥"
    
    If objReport.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '�������ݱ��
    If zlControl.RPTCopyToVSF(objReport, vsTemp) Is Nothing Then Exit Sub
    
    '���ô�ӡ��������
    
    Set objPrint.Body = Me.vsTemp
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub



