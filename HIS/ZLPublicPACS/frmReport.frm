VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{257A5750-6F4D-4A7A-A149-21D28B3E6EAA}#6.1#0"; "ZLPacsRichPages.ocx"
Begin VB.Form frmReport 
   Caption         =   "�������"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12960
   Icon            =   "frmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8805
   ScaleWidth      =   12960
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBarControl 
      BorderStyle     =   0  'None
      Height          =   700
      Left            =   0
      ScaleHeight     =   705
      ScaleWidth      =   12735
      TabIndex        =   2
      Top             =   0
      Width           =   12735
      Begin XtremeCommandBars.CommandBars cbrMain 
         Left            =   2520
         Top             =   240
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7455
      ScaleWidth      =   13815
      TabIndex        =   0
      Top             =   960
      Width           =   13815
      Begin VB.PictureBox picVLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5940
         Left            =   5160
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5940
         ScaleWidth      =   30
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   600
         Width           =   30
      End
      Begin VB.PictureBox picReportContent 
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   8280
         ScaleHeight     =   7095
         ScaleWidth      =   5295
         TabIndex        =   3
         Top             =   120
         Width           =   5295
         Begin ZLPacsRichPageScale.ZLRichPageScaleAct zlDocEditor 
            Bindings        =   "frmReport.frx":6852
            Height          =   1575
            Left            =   1800
            TabIndex        =   4
            Top             =   960
            Width           =   2415
            Object.Visible         =   -1  'True
            AutoScroll      =   0   'False
            AutoSize        =   0   'False
            AxBorderStyle   =   1
            BorderWidth     =   0
            Caption         =   "ZLRichPages"
            Color           =   -16777201
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            KeyPreview      =   0   'False
            PixelsPerInch   =   96
            PrintScale      =   1
            Scaled          =   -1  'True
            DropTarget      =   0   'False
            HelpFile        =   ""
            PopupMode       =   0
            ScreenSnap      =   0   'False
            SnapBuffer      =   10
            DockSite        =   0   'False
            DoubleBuffered  =   0   'False
            ParentDoubleBuffered=   0   'False
            UseDockManager  =   0   'False
            Enabled         =   -1  'True
            AlignWithMargins=   0   'False
            HMenuVisible    =   -1  'True
            VMenuVisible    =   -1  'True
            ReadOnly        =   0   'False
            Orientation     =   0
            BottomMagin     =   2.54
            BoundLeftRight  =   20
            FooterVisible   =   -1  'True
            FooterY         =   10
            MaxPageBreakHeight=   25
            MinPageBreakHeight=   5
            PageBreakHeight =   20
            PageNoFirst     =   1
            PageNoFromNumber=   1
            PageNoHAlign    =   0
            PageNoVAlign    =   0
            PageNoVisible   =   -1  'True
            PageViewMode    =   -1  'True
            RightMargin     =   3.17
            TopMargin       =   2.54
            BackgroundStyle =   3
            CtlColor        =   10070188
            IsShowHint      =   -1  'True
            TabNavigation   =   1
            NoReadOnlyJumps =   0   'False
            NoCaretHighLightJumps=   0   'False
            NoImageResize   =   0   'False
            HideReadOnlyCaret=   -1  'True
            AutoSwitchLang  =   0   'False
            WantTabs        =   -1  'True
            DoNotWantShiftReturns=   0   'False
            DoNotWantReturns=   0   'False
            CtrlJumps       =   -1  'True
            ClearTagOnStyleApp=   0   'False
            IsShowCheckPoints=   0   'False
            IsShowPageBreaks=   0   'False
            IsShowSpecialCharacters=   0   'False
            IsShowHiddenText=   0   'False
            IsShowItemHints =   0   'False
            IsDblClickSelectsWord=   -1  'True
            IsRClickDeselects=   -1  'True
            AlignPageH      =   0
            AlignPageV      =   0
            ViewMode        =   0
            ZoomMode        =   0
            EditZoomMode    =   2
            ZoomPercent     =   68
            ZoomPercentEdit =   100
            ParentCustomHint=   0   'False
            Modified        =   -1  'True
            UndoLimit       =   -1
            IsMarginRectVisible=   -1  'True
            StateView       =   -1  'True
            BackGroundPicture=   "frmReport.frx":6866
            BeginProperty PageNoFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderVisible   =   -1  'True
            HeaderY         =   10
            TableAutoAddRow =   -1  'True
            ThumbilsVisible =   -1  'True
            SimpleViewMode  =   0   'False
            SclRVRulerVVisible=   0   'False
            SclRVRulerHVisible=   0   'False
            ScrollBarVVisible=   -1  'True
            ScrollBarHVisible=   -1  'True
            BackGroudVisible=   -1  'True
            BorderPenStyle  =   0
            Ver             =   "2.1"
            StatusBarVisible=   -1  'True
            CanEdit         =   -1  'True
            DisableCopyElement=   0   'False
            PageWidth       =   21
            PageHeight      =   29.7
            CanPopMenu      =   0   'False
            LeftMargin      =   3.17
            CanInput        =   -1  'True
            TableGridVisible=   0   'False
            CanEditHeader   =   -1  'True
            CanEditFooter   =   -1  'True
            IsRevision      =   0   'False
            RevisionTag     =   ""
            RevisionAddColor=   0
            RevisionDelColor=   0
            MaskText        =   ""
            AllowSelection  =   -1  'True
            BeginProperty MaskTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FinalShowMode   =   0   'False
            DocMasterId     =   ""
            PageSetupInPre  =   0   'False
            ServerTime      =   "1899-12-30"
            XMLEncoding     =   ""
            HScrollPos      =   0
            VScrollPos      =   4
            IsShowMargin    =   0   'False
            IsAutoPageWidth =   0   'False
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Bindings        =   "frmReport.frx":6B2E
            Height          =   1095
            Left            =   2640
            TabIndex        =   5
            Top             =   3480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1931
            _Version        =   393217
            ScrollBars      =   3
            TextRTF         =   $"frmReport.frx":6B81
         End
      End
      Begin VB.PictureBox picReportList 
         BorderStyle     =   0  'None
         Height          =   7095
         Left            =   120
         ScaleHeight     =   7095
         ScaleWidth      =   5700
         TabIndex        =   1
         Top             =   120
         Width           =   5700
         Begin VB.PictureBox picReportControl 
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   360
            ScaleHeight     =   2055
            ScaleWidth      =   6615
            TabIndex        =   8
            Top             =   600
            Width           =   6615
            Begin VB.OptionButton optTime 
               Caption         =   "3����"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   510
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.ComboBox cboDoctor 
               Height          =   300
               Left            =   3600
               Style           =   2  'Dropdown List
               TabIndex        =   20
               Top             =   840
               Width           =   1500
            End
            Begin VB.OptionButton optTime 
               Caption         =   "7����"
               Height          =   180
               Index           =   1
               Left            =   1005
               TabIndex        =   19
               Top             =   510
               Width           =   855
            End
            Begin VB.OptionButton optTime 
               Caption         =   "�������"
               Height          =   180
               Index           =   2
               Left            =   1875
               TabIndex        =   18
               Top             =   510
               Width           =   1095
            End
            Begin VB.OptionButton optTime 
               Caption         =   "һ������"
               Height          =   180
               Index           =   3
               Left            =   3000
               TabIndex        =   17
               Top             =   510
               Width           =   1095
            End
            Begin VB.ComboBox cboKind 
               Enabled         =   0   'False
               Height          =   300
               ItemData        =   "frmReport.frx":6C10
               Left            =   2520
               List            =   "frmReport.frx":6C23
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   1230
               Width           =   975
            End
            Begin VB.TextBox txtKind 
               Enabled         =   0   'False
               Height          =   270
               Left            =   3600
               TabIndex        =   15
               Top             =   1245
               Width           =   1500
            End
            Begin VB.ComboBox cboDept 
               Height          =   300
               IMEMode         =   1  'ON
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   1230
               Width           =   1500
            End
            Begin VB.ComboBox cbo�������� 
               Height          =   300
               Left            =   960
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   840
               Width           =   1500
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "�� ��"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Tag             =   "��"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "ס Ժ"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   11
               Tag             =   "ס"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "�� ��"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   10
               Tag             =   "��"
               Top             =   1680
               Width           =   855
            End
            Begin VB.CheckBox chkPatFrom 
               Caption         =   "�� ��"
               Height          =   255
               Index           =   3
               Left            =   3000
               TabIndex        =   9
               Tag             =   "��"
               Top             =   1680
               Width           =   855
            End
            Begin MSComCtl2.DTPicker dtpStart 
               Height          =   255
               Left            =   960
               TabIndex        =   22
               Top             =   83
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   113180675
               CurrentDate     =   42180
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   255
               Left            =   2760
               TabIndex        =   23
               Top             =   83
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   450
               _Version        =   393216
               CustomFormat    =   "yyyy-MM-dd"
               Format          =   113180675
               CurrentDate     =   42180
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "~"
               Height          =   180
               Left            =   2535
               TabIndex        =   28
               Top             =   120
               Width           =   90
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "�������"
               Height          =   180
               Left            =   120
               TabIndex        =   27
               Top             =   120
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "����ҽ��"
               Height          =   180
               Left            =   2520
               TabIndex        =   26
               Top             =   900
               Width           =   720
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "������"
               Height          =   180
               Left            =   120
               TabIndex        =   25
               Top             =   1290
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "�������"
               Height          =   180
               Left            =   120
               TabIndex        =   24
               Top             =   900
               Width           =   720
            End
         End
         Begin TabDlg.SSTab tabReport 
            Height          =   360
            Left            =   480
            TabIndex        =   7
            Top             =   120
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   635
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "��ǰ���˱���"
            TabPicture(0)   =   "frmReport.frx":6C4D
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).ControlCount=   0
            TabCaption(1)   =   "ȫԺ���˱���"
            TabPicture(1)   =   "frmReport.frx":6C69
            Tab(1).ControlEnabled=   0   'False
            Tab(1).ControlCount=   0
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfReportList 
            Height          =   3615
            Left            =   240
            TabIndex        =   6
            Top             =   2640
            Width           =   6735
            _cx             =   11880
            _cy             =   6376
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
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
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mMenuID
    conMenu_File_PrintSet = 101
    conMenu_File_Preview = 102
    conMenu_File_Print = 103
    conMenu_File_ViewState
    conMenu_File_ViewImage
    conMenu_File_Refresh
    conMenu_File_Exit
End Enum

Private mobjReport As Object
'Private mcnOracle As New ADODB.Connection
Private mrsRecord As ADODB.Recordset

Private mlngAdviceId As Long
Private mstrReportId As String
Private mobjFtp As New clsFtp
Private mobjFtpInfo As tFtpInfo

Public mblnAutoView As Boolean  '�Ƿ��Զ����ģ�True--�Զ����ģ�������б���Զ���ǡ��Ѳ��ġ�
Public mblnRIS As Boolean       '�Ƿ�RIS�в鿴�����������ʾ�˳������ģ���ӡ�Ȱ�ť�����ò�ѯ��ʼʱ��Ϊһ��
Public mblnShow As Boolean      '�Ƿ���ʾ��ߵĹ�������

Private mblnIsNewReport As Boolean
Private mblnIsNoAskPrint As Boolean
Private mlngPatFrom As Long
Private mblnIsConfiging As Boolean
Private mblnFirst As Boolean
Private mlngViewReport As Long      '0--���ǩ���󼴿ɲ鿴���棬1--����ǩ���󼴿ɲ鿴����


Private Sub InitCommandBars()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrPopControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    Dim str3DFuncs() As String
    Dim blnShowCaption As Boolean
    
    Dim rsCollection As ADODB.Recordset
    Dim rsViewShare As ADODB.Recordset
    Dim rsShareCount As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    
    Dim i As Integer
    Dim i3DFunc As Integer
    Dim intTxtLen As Integer
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbrMain.VisualTheme = xtpThemeOffice2003
    Set cbrMain.Icons = gobjComLib.zlCommFun.GetPubIcons
    
    With cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbrMain.EnableCustomization False
    cbrMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbrMain.ActiveMenuBar.Visible = False
    
'---------------------����������------------------------------------------
    Set cbrToolBar = cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = True
    
    If mblnRIS = False Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_PrintSet, "��ӡ����", "��ӡ����", 181, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Preview, "Ԥ��", "��ӡԤ��", 102, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Print, "��ӡ", "��ӡ", 103, False)
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_ViewState, "����", "����", 2322, True)
    End If
    
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_ViewImage, "��Ƭ", "��Ƭ", 8111, True)
    Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Refresh, "ˢ��", "��ѯˢ������", 791, True)
    
    If mblnRIS = False Then
        Set cbrControl = CreateModuleMenu(cbrToolBar.Controls, xtpControlButton, conMenu_File_Exit, "�˳�", "�˳�", 2613, True)
    End If
End Sub

''Ϊ�˲�ѯClob�������ݣ��軻��oledb���ӷ�ʽ
'Public Function ConnectOracle(ByVal strUser As String, ByVal strPassW As String, ByVal strServer As String) As Boolean
'On Error GoTo ErrH
'    ConnectOracle = False
'
'    '�ж�����״̬
'    If mcnOracle.State = adStateOpen Then mcnOracle.Close
'
'    mcnOracle.ConnectionString = "Provider=OraOLEDB.Oracle.1;User ID=" & strUser & ";password=" & strPassW & ";Data Source=" & strServer & ";Persist Security Info=False"
'    mcnOracle.CursorLocation = adUseClient
'
'    mcnOracle.Open
'
'    If mcnOracle.State = adStateOpen Then
'        ConnectOracle = True
'    Else
'        ConnectOracle = False
'    End If
'
'    Exit Function
'ErrH:
'    MsgBox err.Description, vbCritical, "ϵͳ��Ϣ"
'    err.Clear
'End Function

Public Sub InitEdit()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    '��������
    strSql = " Select Distinct A.ID,A.����,A.����,b.��������" & _
                " From ���ű� A,��������˵�� B " & _
                " Where B.����ID = A.ID " & _
                " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                " And (B.�������� IN('�ٴ�','���','���'))" & _
                " Order by A.����"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, Caption)
    
    cbo��������.Clear
    cbo��������.AddItem ""
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cbo��������.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        If cbo��������.ListCount > 0 And cbo��������.ListIndex = -1 Then cbo��������.ListIndex = 0
    End If
    
    cboDept.Clear
    cboDept.AddItem ""
    rsTmp.Filter = "��������='���'"
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    End If
    
    If cboKind.ListCount > 0 And cboKind.ListIndex = -1 Then cboKind.ListIndex = 0
    
    dtpEnd.Value = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    dtpStart = Format(Now - 3, "yyyy-mm-dd HH:MM")
End Sub

Private Sub InitDoctors(ByVal lng����ID As Long)
'���ܣ���ȡ��ǰ���������а�����������Ա
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    
    strSql = "Select /*+ RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b, ��Ա����˵�� c" & vbNewLine & _
                " Where a.��Աid = b.Id And b.Id = c.��Աid And c.��Ա���� = 'ҽ��' And" & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null) and a.����id = [1] " & vbNewLine & _
                " Order By ���� "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, Caption, lng����ID)
    
    cboDoctor.Clear
    cboDoctor.AddItem ""
    If Not rsTmp.EOF Then
        Do Until rsTmp.EOF
            cboDoctor.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ID = UserInfo.ID Then cboDoctor.ListIndex = cboDoctor.NewIndex
            rsTmp.MoveNext
        Loop
        
        If cboDoctor.ListCount > 0 And cboDoctor.ListIndex = -1 Then cboDoctor.ListIndex = 0
    End If
End Sub

'Private Function InitOracleConn() As Boolean
'    Dim strUser As String, strPwd As String, strServer As String
'
'On Error GoTo Errorhand
'
'    strUser = UCase(gcnOracle.Properties(23))
'    strPwd = gcnOracle.Properties(24)
'    strServer = UCase(Split(gcnOracle.Properties(8), "=")(2))
'    gstrDBUser = strUser
'
'    'Ϊ�˲�ѯClob�������ݣ��軻��oledb���ӷ�ʽ
'    If Not ConnectOracle(strUser, strPwd, strServer) Then Exit Function
'
'    InitOracleConn = True
'
'    Exit Function
'Errorhand:
'    MsgBox err.Description, vbExclamation, gstrSysName
'End Function

Public Sub ShowMe(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional ByVal blnAutoView As Boolean = True, Optional objParent As Object, Optional blnShowModal As Boolean = False)
    
On Error GoTo Errorhand
    
    'ֱ�Ӵ�PACS���洰��ʱ��Ĭ����ʾ��ߵ��б�ͱ����ӡ�����ģ��˳��Ȱ�ť
    mblnShow = True
    mblnRIS = False
    
    mblnAutoView = blnAutoView
    
    Call RefreshForm(lngAdviceID, strReportId, objParent)
    
    Call Show(IIf(blnShowModal, 1, 0), objParent)
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function LoadImageFromXml(ByVal strXml As String, ByVal strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'���ܣ�����xml�ĵ�����ͼ�����ص�ָ��λ�ã�����������ͼ���ļ���
'���أ�
    Dim objXml As New DOMDocument
    Dim objXmlNodes As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeAttribute As IXMLDOMNode
    Dim strImgSVG As String
    Dim objImgFileName As New Collection
    Dim objSvg As New ZLSvgProcess.zlSvgConvert
    Dim objPic As StdPicture
    
On Error GoTo Errorhand
    
    If objXml.loadXML(strXml) = False Then
        strMsg = "�������ݼ���ʧ�ܣ�"
        Exit Function
    End If
    
    Set objXmlNodes = objXml.selectNodes("*//image")
    
    If objXmlNodes.length <= 0 Then
        Set LoadImageFromXml = objImgFileName
        strMsg = "�˱���û��ͼ��"
        Exit Function
    End If
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '��FTP�ϻ�ȡͼ���ļ��󷵻�ͼ��
            strImgSVG = objSvg.DecodeBase64(GetFtpImgSVG(objXmlNodeAttribute.Text, strMsg))
            
            If objSvg.IsSvgContext(strImgSVG) Then
                '��ͼ��ŵ�ָ��Ŀ¼��
                Set objPic = objSvg.ContextToPic(strImgSVG)
                Call SavePicture(objPic, strTmpImgDir & objXmlNodeAttribute.Text & ".jpg")
                
                '��ͼ�����Ʒ���ͼ�񼯺���
                Call objImgFileName.Add(objXmlNodeAttribute.Text)
            End If
        End If
    Next
    
    Set LoadImageFromXml = objImgFileName
    
    Exit Function
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function AddImageInfoToXml(ByVal strXml As String, ByVal strDocId As String) As String
'���ܣ�����xml�ĵ�����ͼ����Ϣ�����ĵ���
'���أ�
    Dim objXml As New DOMDocument
       
    Dim objXmlNodes As IXMLDOMNodeList
    Dim objXmlNode As IXMLDOMNode
    Dim objXmlNodeAttribute As IXMLDOMNode
    Dim strImgSVG As String
    
On Error GoTo Errorhand
    
    If objXml.loadXML(strXml) = False Then
        MsgBox "�������ݼ���ʧ�ܣ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Set objXmlNodes = objXml.selectNodes("*//image")
    
    If objXmlNodes.length <= 0 Then
        AddImageInfoToXml = strXml
        Exit Function
    End If
    
    '��ʼ��FTP�����Ϣ
    Call InitFtpInfo(strDocId)
    
    For Each objXmlNode In objXmlNodes
        Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("key")
        
        If Not objXmlNodeAttribute Is Nothing Then
            '��FTP�ϻ�ȡͼ���ļ��󷵻�ͼ��
            strImgSVG = GetFtpImgSVG(objXmlNodeAttribute.Text)
            
            Set objXmlNodeAttribute = objXmlNode.Attributes.getNamedItem("img")
            '��ͼ����Ϣд��xml
            objXmlNodeAttribute.Text = strImgSVG
        End If
    Next
    
    AddImageInfoToXml = objXml.xml
    
    Exit Function
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Function InitFtpInfo(ByVal strDocId As String) As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    strSql = "Select 'ReportImages/' || to_Char(b.����ʱ��,'YYYYMMDD') || '/' || b.id || '/' As URL," & _
            "a.�豸��, a.FTP�û���, a.FTP����, a.IP��ַ,'/'||a.FtpĿ¼||'/' As Root " & _
            "From Ӱ���豸Ŀ¼ a, Ӱ�񱨸��¼ b where a.�豸�� = b.�豸�� And b.id = [1]"
    
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡFTP��Ϣ", strDocId)
    
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    mobjFtpInfo.FtpDir = Nvl(rsTmp("Root"))
    mobjFtpInfo.FtpIP = Nvl(rsTmp("IP��ַ"))
    mobjFtpInfo.FtpPswd = Nvl(rsTmp("FTP����"))
    mobjFtpInfo.FTPUser = Nvl(rsTmp("FTP�û���"))
    mobjFtpInfo.DiviceId = Trim(Nvl(rsTmp("�豸��")))
    
    mobjFtpInfo.SubDir = Nvl(rsTmp("URL"))
    mobjFtpInfo.DestMainDir = IIf(Len(App.Path) > 3, App.Path & "\TmpReportImage\", App.Path & "TmpReportImage\")
    
    InitFtpInfo = True
End Function

Private Function ConnFtp() As Boolean
    If mobjFtp.hConnection = 0 Then
        '����FTP�洢�豸
        If mobjFtp.FuncFtpConnect(mobjFtpInfo.FtpIP, mobjFtpInfo.FTPUser, mobjFtpInfo.FtpPswd) = 0 Then
            Exit Function
        End If
    End If
    
    ConnFtp = True
End Function

'��FTP�ϻ�ȡSVG��ʽͼ��
Private Function GetFtpImgSVG(ByVal strKey As String, Optional ByRef strMsg As String = "") As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strLocalFileName As String
    Dim strVirtualPath As String
    
    If strKey = "" Then Exit Function
    
    strLocalFileName = Replace(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir & strKey, "/", "\")
    strVirtualPath = Replace(mobjFtpInfo.FtpDir & mobjFtpInfo.SubDir, "\", "/")
    
    '��������·��
    If Not objFSO.FolderExists(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir) Then
        Call MkLocalDir(mobjFtpInfo.DestMainDir & mobjFtpInfo.SubDir)
    End If
    
    '������ش�����ɾ��
    If objFSO.FileExists(strLocalFileName) Then Call objFSO.DeleteFile(strLocalFileName, True)
    
    '����FTP
    If ConnFtp() = False Then
        strMsg = "FTP�����������ӣ������������á�"
        Exit Function
    End If
    
    If mobjFtp.FuncDownloadFile(strVirtualPath, strLocalFileName, objFSO.GetFileName(strLocalFileName)) <> 0 Then
        strMsg = "ͼ�����ݴ�FTP�������ϻ�ȡʧ�ܣ�"
        Exit Function
    End If
    
    '���غ��ȡ
    GetFtpImgSVG = GetFileContent(strLocalFileName)
End Function

Private Sub InitReportEditor()
    zlDocEditor.FooterVisible = False
    zlDocEditor.HeaderVisible = False
    zlDocEditor.HMenuVisible = False
    zlDocEditor.PageNoVisible = False
    zlDocEditor.ThumbilsVisible = False
    zlDocEditor.VMenuVisible = False
    zlDocEditor.ZoomPercent = 100
    zlDocEditor.CanEdit = False
    zlDocEditor.CanInput = False
    zlDocEditor.TableGridVisible = False
    zlDocEditor.InitOCX hWnd
    rtxtReport.Locked = True
End Sub

Private Sub InitReportList()
'��ʼ����ʽ�����б�
    With vsfReportList
        .ColHidden(0) = True
        .ColHidden(1) = True
        .ColHidden(2) = True
        .AutoSize 0, .Cols - 1
        .ScrollBars = flexScrollBarBoth
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
    End With
End Sub

Private Sub LoadReportContent(ByVal strContent As String, ByVal strDocId As String)
    Dim strXml As String
    
    If strContent = "" Then Exit Sub
    
    '����xml�ĵ�����ͼ����Ϣ�����ĵ���
    strXml = AddImageInfoToXml(strContent, strDocId)
    strXml = Replace(strXml, "��", "��")
    strXml = Replace(strXml, "�P", "��")
    strXml = Replace(strXml, "�H", "��")
    
    zlDocEditor.OpenWithXML strXml
    zlDocEditor.FinalShowMode = True
End Sub

Public Sub RefreshReportContent(ByVal strReportId As String)
    Call ViewReportContent(strReportId)

'    If Trim(strReportId) <> "" Then
'        Call ViewReportContent(strReportId)
'    Else
'        Call LoadReport("")
'    End If
    
End Sub

Private Sub cboDept_Click()
    Call GetFilterData
End Sub

Private Sub cboDept_DropDown()
     On Error GoTo errHandle
    Call SendMessage(cboDept.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cboDoctor_Click()
    Call GetFilterData
End Sub

Private Sub cboDoctor_DropDown()
    On Error GoTo errHandle
    Call SendMessage(cboDoctor.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cboKind_Click()
    txtKind.Text = ""
End Sub

Private Sub cboKind_DropDown()
    On Error GoTo errHandle
    Call SendMessage(cboKind.hWnd, &H160, 100, 0)
errHandle:
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 And cbo��������.Text <> "" Then
        InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
    Else
        cboDoctor.Clear
    End If
    
    Call GetFilterData
End Sub

Private Sub Menu_File_Preview(ByVal Control As XtremeCommandBars.ICommandBarControl, ByVal lngAdviceID As Long)
On Error GoTo ErrHand
    If mobjReport Is Nothing Then Set mobjReport = CreateObject("zlRichEPR.cDockReport")        '���Ӳ�������
    
    If mblnIsNewReport Then
        Call zlDocEditor.PrintPreview(False, False, False, False, True)
    Else
        If Not mobjReport Is Nothing Then
            mobjReport.zlRefresh 0, 0
            mobjReport.zlRefresh lngAdviceID, UserInfo.����ID
            mobjReport.zlExecuteCommandBars Control
        End If
    End If
    
    Exit Sub
    
ErrHand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Menu_File_Print(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objReport As Object
    
    If mblnIsNewReport Then
        Call zlDocEditor.PrintPages
'    Else
'        Set objReport = CreateObject("zlRichEPR.cDockReport")       '���Ӳ�������
'
'        If Not objReport Is Nothing Then
'            objReport.zlRefresh 0, 0, , , , 1258
'            objReport.zlRefresh mlngAdviceId, UserInfo.����ID, , , True, 1258
'            objReport.zlExecuteCommandBars Control
'        End If
    End If
End Sub


Private Sub cbo��������_DropDown()
     On Error GoTo errHandle
    Call SendMessage(cbo��������.hWnd, &H160, 150, 0)
errHandle:
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error GoTo Errorhand
    Dim lngAdviceID As Long
    
    If Control.ID = conMenu_File_ViewImage Or Control.ID = conMenu_File_Preview Then
        If vsfReportList.RowSel <= 0 Then
            MsgBox "����ѡ����Ҫ�����ļ��", vbExclamation, gstrSysName
            Exit Sub
        End If
        lngAdviceID = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "ҽ��ID"))
    End If
    Select Case Control.ID
        Case conMenu_File_PrintSet  '��ӡ����
            Call zlPrintSet
            
        Case conMenu_File_Preview   'Ԥ��
            Call Menu_File_Preview(Control, lngAdviceID)
                
        Case conMenu_File_Print '��ӡ
            Call Menu_File_Print(Control)
                
        Case conMenu_File_ViewState  '����
            If vsfReportList.RowSel <= 0 Then Exit Sub
            Call UpdateReportViewState(vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "����ID")))
                
        Case conMenu_File_ViewImage '��Ƭ
            If vsfReportList.RowSel <= 0 Then Exit Sub
            Call mdlPublic.ViewImage(lngAdviceID, Me)
        
        Case conMenu_File_Refresh   '��ѯ��ˢ������
            Call LoadReport(GetFilter)
        
        Case conMenu_File_Exit  '�˳�
            Unload Me
        
    End Select
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cbrMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
On Error GoTo Errorhand
    Select Case Control.ID
        Case mMenuID.conMenu_File_ViewState
            If vsfReportList.Rows <= 1 Then Exit Sub
            Control.Caption = IIf(vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "����")) = "��", "�Ѳ���", "����")
            Control.Enabled = Not mblnAutoView And Not vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "����")) = "��"
            If Control.Enabled Then Control.Enabled = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "ҽ��ID")) = mlngAdviceId
            
        Case mMenuID.conMenu_File_Preview, mMenuID.conMenu_File_Print
            Control.Enabled = vsfReportList.Rows > 1
                    
    End Select
    
    Control.Enabled = Not Control.Enabled
    Control.Enabled = Not Control.Enabled
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Function GetFilter() As String
    Dim strPatFrom As String
    Dim strFilter As String
    Dim i As Integer
    Dim intPatFromCount As Integer
    
On Error GoTo Errorhand
    
    If cbo��������.Text <> "" Then
        strFilter = strFilter & "��������id = " & cbo��������.ItemData(cbo��������.ListIndex)
    End If
    
    If cboDoctor.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "������ = '" & Split(cboDoctor.Text, "-")(1) & "'"
    End If
    
    If cboDept.Text <> "" Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        strFilter = strFilter & "ִ�п���id = '" & cboDept.ItemData(cboDept.ListIndex) & "'"
    End If
    
    If cboKind.Text <> "" And Trim(txtKind.Text) <> "" And tabReport.Tab = 1 Then
        If strFilter <> "" Then strFilter = strFilter & " and "
        
        If cboKind.Text = "����" Then
            strFilter = strFilter & cboKind.Text & " like '" & Trim(txtKind.Text) & "%'"
        Else
            strFilter = strFilter & cboKind.Text & " = '" & Trim(txtKind.Text) & "'"
        End If
    End If
    
    For i = 0 To chkPatFrom.Count - 1
        If chkPatFrom(i).Value = 0 Then
            intPatFromCount = intPatFromCount + 1
        End If
    Next
    
    If intPatFromCount <> chkPatFrom.Count Then
        If chkPatFrom(0).Value = 0 Then
            strPatFrom = "��Դ <> '��'"
        End If
    
        If chkPatFrom(1).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "��Դ <> 'ס'"
        End If
    
        If chkPatFrom(2).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "��Դ <> '��'"
        End If
    
        If chkPatFrom(3).Value = 0 Then
            If strPatFrom <> "" Then strPatFrom = strPatFrom + " and "
            strPatFrom = strPatFrom & "��Դ <> '��'"
        End If
        
        If strPatFrom <> "" Then
            If strFilter <> "" Then strFilter = strFilter & " and "
            strFilter = strFilter & strPatFrom
        End If
    End If
    
    
    GetFilter = strFilter
    
    Exit Function
Errorhand:
    GetFilter = ""
    MsgBox err.Description, vbExclamation, gstrSysName
End Function

Private Sub LoadReport(strFilter As String, Optional ByVal blnRefreshFormCall As Boolean = False)
'blnRefreshFormCall �Ƿ�RefreshForm���ã��ȼ��ڵ�һ�μ������ݣ�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strTmp As String
On Error GoTo Errorhand
    If mblnIsConfiging Then Exit Sub
    
    If mlngViewReport = 0 Then
        strTmp = "b.����״̬ in (3,4)"
    Else
        strTmp = "b.����״̬ = 4"
    End If
    
    If tabReport.Tab = 0 Then
        'RISʹ���Լ���һ�ײ�ѯSQL
        If mblnRIS = True Then
            strSql = "Select * from (Select 2 as ����,TO_CHAR(RAWTOHEX(b.��鱨��id)) ����ID,a.ִ�п���id, a.Ӱ�����,a.�������� As ���ʱ��, c.ҽ������, c.����ʱ��, " & _
                "  f.������ as ������,f.�������� as �����, c.Id As ҽ��id,b.����id " & _
                " From Ӱ�����¼ a, ����ҽ������ b, ����ҽ����¼ c, Ӱ�����¼ d, ����ҽ����¼ e,Ӱ�񱨸��¼ f" & _
                " Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And a.ҽ��id = f.ҽ��id And e.Id = " & mlngAdviceId & " And b.ҽ��id = c.Id " & _
                " And (c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null " & _
                " And b.��鱨��id is not null And b.��鱨��id =f.id And f.����״̬ in (2,3,4)" & _
                " Union All " & _
                "Select 1 as ����,TO_CHAR(b.����id) ����ID,a.ִ�п���id, a.Ӱ�����,a.�������� As ���ʱ��, c.ҽ������, c.����ʱ��, " & _
                " a.������,a.������ �����,c.Id As ҽ��id, b.����id " & _
                " From Ӱ�����¼ a, ����ҽ������ b, ����ҽ����¼ c, Ӱ�����¼ d, ����ҽ����¼ e,����ҽ������ f" & _
                " Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id = " & mlngAdviceId & " And b.ҽ��id = c.Id " & _
                " And a.ҽ��ID = f.ҽ��ID And (c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null " & _
                " And f.ִ�й��� > =4 And Nvl(Rawtohex(��鱨��id), ' ') = ' ' )  Order By ���ʱ�� Asc "

        Else
            strSql = "Select* from (Select distinct m.����, m.����ID, m.ִ�п���id, m.����, m.��ӡ, m.����,m.Ӱ�����, m.ҽ������,m.����ʱ�� , m.�Ա�, m.����,  " & _
                "Decode(m.������Դ,1,'��',2,'ס',3,'��','��') as ��Դ,   O.סԺ��, O.��Ժ���� as ����, " & _
                "m.������, m.������, m.������, n.�����, n.���￨��, n.���֤��,m.����id,m.ҽ��id, m.��ҳID, m.��������ID " & _
                "From (Select distinct 2 as ����,TO_CHAR(RAWTOHEX(b.id)) ����ID,a.ִ�п���id,decode(nvl(f.����״̬,0),0,'','��') ����, " & _
                "decode(nvl(b.�����ӡ,0),0,'','��') ��ӡ,a.����,a.�Ա�,a.����,a.Ӱ�����,c.ҽ������,c.������Դ, c.��ҳID, " & _
                "c.����ʱ�� as ����ʱ��,c.����ҽ�� ������,b.������ ������,b.�������� ������,c.����id, c.Id As ҽ��id,c.��������ID " & _
                "From Ӱ�����¼ A, Ӱ�񱨸��¼ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E,����ҽ������ F " & _
                "Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And b.ҽ��id = c.Id And b.ҽ��id=f.ҽ��id and b.id=f.��鱨��id and " & _
                "(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null and " & strTmp & " and e.id = " & mlngAdviceId & " Union All " & _
                "Select distinct 1 as ����,TO_CHAR(b.����id) ����ID,a.ִ�п���id,decode(nvl(b.����״̬,0),0,'','��') ����, " & _
                "decode(nvl(a.�����ӡ,0),0,'','��') ��ӡ,a.����,a.�Ա�,a.����,a.Ӱ�����,c.ҽ������,c.������Դ,c.��ҳID, " & _
                "c.����ʱ�� as ����ʱ��,c.����ҽ�� ������,a.������,a.������ ������,c.����id, c.Id As ҽ��id,c.��������ID " & _
                "From Ӱ�����¼ A, ����ҽ������ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E " & _
                "Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And b.ҽ��id = c.Id And b.����ID Is Not Null And " & _
                "(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null and nvl(a.������,' ')<>' ' and e.id = " & mlngAdviceId & ") m,������Ϣ n,������ҳ o,����ҽ������ p " & _
                "where m.����id = n.����id and m.����id=o.����id(+) and m.��ҳID=o.��ҳID(+) and m.ҽ��id = p.ҽ��id " & _
                IIf(mblnFirst Or blnRefreshFormCall, "", "and p.����ʱ�� between to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss') and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')") & ") " & _
                IIf(strFilter = "", "", " where " & strFilter)
        End If
    Else
        strSql = "Select * from (select distinct 2 as ����,TO_CHAR(RAWTOHEX(f.id)) ����ID, b.ִ�п���id,decode(nvl(g.����״̬,0),0,'','��') ����," & _
                "decode(nvl(f.�����ӡ,0),0,'','��') ��ӡ,c.����,e.Ӱ�����,b.ҽ������,b.����ʱ�� as ����ʱ��,c.�Ա�,c.����,Decode(b.������Դ,1,'��',2,'ס',3,'��','��') as ��Դ,d.סԺ��," & _
                "d.��Ժ���� as ����, b.����ҽ�� ������,f.������ as ������,f.�������� ������,c.�����, c.���￨��, c.���֤��,c.����id,b.id as ҽ��ID,b.��������id " & _
                "from ����ҽ������ A, ����ҽ����¼ b, ������Ϣ c, ������ҳ d, Ӱ�����¼ e,Ӱ�񱨸��¼ f, ����ҽ������ g " & _
                "where a.ҽ��id=b.id and b.����id=c.����id and b.����id=d.����id(+) and b.��ҳid=d.��ҳid(+) and b.id=e.ҽ��id and e.ҽ��id=f.ҽ��id and f.id=g.��鱨��id and b.id=g.ҽ��id " & _
                "and a.����ʱ�� between trunc(to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')) and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss') " & _
                "Union All " & _
                "select distinct 1 as ����,TO_CHAR(g.����id) ����ID, b.ִ�п���id,decode(nvl(g.����״̬,0),0,'','��') ����," & _
                "decode(nvl(e.�����ӡ,0),0,'','��') ��ӡ,c.����,e.Ӱ�����,b.ҽ������,b.����ʱ�� as ����ʱ��,c.�Ա�,c.����,Decode(b.������Դ,1,'��',2,'ס',3,'��','��') as ��Դ,d.סԺ��," & _
                "d.��Ժ���� as ����, b.����ҽ�� ������,e.������,e.������ ������,c.�����, c.���￨��, c.���֤��,c.����id,b.id as ҽ��ID,b.��������id " & _
                "from ����ҽ������ A, ����ҽ����¼ b, ������Ϣ c, ������ҳ d, Ӱ�����¼ e, ����ҽ������ g " & _
                "where a.ҽ��id=b.id and b.����id=c.����id and b.����id=d.����id(+) and b.��ҳid=d.��ҳid(+) and b.id=e.ҽ��id and e.ҽ��id=g.ҽ��id and g.����id is not null and b.id=g.ҽ��id " & _
                "and a.����ʱ�� between trunc(to_date('" & Format(dtpStart.Value, "yyyy-mm-dd 00:00:00") & "','yyyy-mm-dd hh24:mi:ss')) and to_date('" & Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59") & "','yyyy-mm-dd hh24:mi:ss')) " & _
                IIf(strFilter = "", "", " where " & strFilter)
    End If
    
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "��ȡ������Ϣ") ' GetRecordset(strSql)
    Debug.Print rsTemp.RecordCount
    If mblnFirst Or blnRefreshFormCall Then rsTemp.Filter = "ҽ��id=" & mlngAdviceId
    
    Call SetReportlistDataSource(rsTemp)
    Set mrsRecord = rsTemp
    
    Call InitReportList
    
    Exit Sub
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub GetFilterData()
    If mrsRecord Is Nothing Then
        Call LoadReport(GetFilter())
    Else
        mrsRecord.Filter = GetFilter()
        Call SetReportlistDataSource(mrsRecord)
    End If
End Sub

Private Sub chkPatFrom_Click(index As Integer)
    Call GetFilterData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Call LoadReport(GetFilter())
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Errorhand
    mblnIsConfiging = True
    
    Call InitReportEditor
    
    If mblnShow Then
        Call InitCommandBars
        Call InitEdit
    End If
    
    'û��Ȩ�ޣ�������RIS���ã�����ʾȫ������ҳ��
    If (InStr(gstrPrivs, VIEW_ALLREPORT) <= 0 Or mblnRIS = True) Then tabReport.TabVisible(1) = False
    
    picReportControl.Visible = Not mblnRIS
        
    mblnIsConfiging = False
    mblnFirst = True
    
    If Trim(mstrReportId) <> "" Then
        Call RefreshReportContent(mstrReportId)
    Else
        Call LoadReport("")
    End If
    
    mblnFirst = False
    
    Exit Sub
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    picBarControl.Left = 0
    picBarControl.Top = 0
    picBarControl.Width = ScaleWidth
    
    If Not mblnShow Then
        picVLine.Visible = False
        picReportList.Visible = False
        picBarControl.Visible = False
        picBarControl.Height = 0
    End If
    
    picBack.Left = 0
    picBack.Top = picBarControl.Height
    picBack.Width = ScaleWidth
    picBack.Height = ScaleHeight - picBarControl.Height
    
    picVLine.Top = ScaleTop
    picVLine.Height = picBack.ScaleHeight
    If picVLine.Left < 500 Then picVLine.Left = 500
    If picVLine.Left > picBack.ScaleWidth - 500 Then picVLine.Left = picBack.ScaleWidth - 500
    
    picReportList.Left = ScaleLeft
    picReportList.Width = picVLine.Left - picReportList.Left
    picReportList.Top = picBack.ScaleTop
    picReportList.Height = picBack.ScaleHeight - picReportList.Top
    
    picReportContent.Left = IIf(picReportList.Visible, picVLine.Left + picVLine.Width, 0)
    
    picReportContent.Width = picBack.ScaleWidth - IIf(picReportList.Visible, picReportContent.Left, 0)
    picReportContent.Top = picBack.ScaleTop
    picReportContent.Height = picBack.ScaleHeight - picReportContent.Top
    
    Call picReportList_Resize
End Sub

Private Sub picVLine_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next

    If Button = 1 Then
        picVLine.Left = picVLine.Left + x
        
        Form_Resize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errorhand
    
    Set mobjReport = Nothing
    
    If Not mobjFtp Is Nothing Then
        If mobjFtp.hConnection <> 0 Then mobjFtp.FuncFtpDisConnect
        Set mobjFtp = Nothing
    End If
    
    Set mrsRecord = Nothing
    
    Exit Sub
    
    zlDocEditor.ClearAll
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub optTime_Click(index As Integer)
On Error Resume Next
    dtpEnd.Value = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd HH:MM")
    
    If index = 0 Then
        dtpStart = Format(Now - 3, "yyyy-mm-dd HH:MM")
    ElseIf index = 1 Then
        dtpStart = Format(Now - 7, "yyyy-mm-dd HH:MM")
    ElseIf index = 2 Then
        dtpStart = Format(Now - 15, "yyyy-mm-dd HH:MM")
    Else
        dtpStart = Format(Now - 30, "yyyy-mm-dd HH:MM")
    End If
    
    Call LoadReport(GetFilter())
End Sub

Private Sub picReportContent_Resize()
    On Error Resume Next
    
    zlDocEditor.Left = 0
    zlDocEditor.Top = 0
    zlDocEditor.Width = picReportContent.Width
    zlDocEditor.Height = picReportContent.Height
    
    rtxtReport.Left = 0
    rtxtReport.Top = 0
    rtxtReport.Height = picReportContent.Height
    rtxtReport.Width = picReportContent.Width
End Sub

Private Sub picReportList_Resize()
    On Error Resume Next
    
    tabReport.Left = 20
    tabReport.Top = 20
    tabReport.Width = picReportList.ScaleWidth - 40
    
    picReportControl.Left = tabReport.Left
    picReportControl.Top = tabReport.Top + tabReport.Height
    picReportControl.Width = tabReport.Width
    
    vsfReportList.Left = picReportControl.Left
    vsfReportList.Top = picReportControl.Top + picReportControl.Height
    vsfReportList.Width = picReportList.ScaleWidth - 20
    
    If Not mblnShow Then
        picReportControl.Enabled = False
        tabReport.Enabled = False
    End If
	
	vsfReportList.Height = picReportList.ScaleHeight - picReportControl.Top - picReportControl.Height 
End Sub

Private Sub UpdateReportViewState(ByVal strDocId As String)
    Dim strSql As String
    
On Error GoTo Errorhand
    
    strSql = "Zl_Ӱ�񱨸���ļ�¼_Insert(" & mlngAdviceId & ", '" & strDocId & "')"
    Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
    
    If vsfReportList.RowSel <= 0 Then Exit Sub
    vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "����")) = "��"
    
    Exit Sub
    
Errorhand:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub tabReport_Click(PreviousTab As Integer)
    '����������ֻ��ȫԺ��ѯ�ſ���
    cboKind.Enabled = tabReport.Tab = 1
    txtKind.Enabled = cboKind.Enabled
    txtKind.Text = ""
    
    If cbo��������.ListCount > 0 Then cbo��������.ListIndex = 0
    If cboDept.ListCount > 0 Then cboDept.ListIndex = 0
    
    '�л�ҳ��ˢ������
    Call LoadReport(GetFilter)
End Sub

Private Sub txtKind_Change()
    Call GetFilterData
End Sub

Private Sub txtKind_KeyPress(KeyAscii As Integer)
    If cboKind.Text = "�����" Or cboKind.Text = "סԺ��" Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) <= 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsfReportList_SelChange()
    Dim strCurReportId As String
On Error GoTo ErrorHnad
    If vsfReportList.RowSel <= 0 Then
        If vsfReportList.Rows > 1 Then
            vsfReportList.RowSel = 1
        End If
        Exit Sub
    End If
    
    strCurReportId = vsfReportList.Cell(flexcpText, vsfReportList.RowSel, 1)
    
    Call ViewReportContent(strCurReportId)
     
    If mblnAutoView And vsfReportList.Cell(flexcpText, vsfReportList.RowSel, GetColNum(vsfReportList, "����")) <> "��" Then
        Call UpdateReportViewState(strCurReportId)
    End If
    
    Exit Sub
ErrorHnad:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub ViewReportContent(ByVal strDocId As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim intReportEditor As Integer
    Dim strContent As String
    Dim intLoop As Integer
    
    If strDocId = "" Then
        Call ClearReportContent
        Exit Sub
    End If
    
    If Len(strDocId) = 32 Then
        rtxtReport.Visible = False
        zlDocEditor.Visible = True
        mblnIsNewReport = True
        
        strSql = "Select Length(a.��������.GetClobVal()) as ContentLength From Ӱ�񱨸��¼ a Where a.ID = '" & strDocId & "'"
        Set rsTemp = GetRecordset(strSql)
        
        If rsTemp.BOF = False Then
            If rsTemp("ContentLength").Value > 2000 Then
                For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
                    strSql = "select to_char(substr(a.��������.GetClobVal()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                             " from Ӱ�񱨸��¼ a where a.ID = '" & strDocId & "'"
                             
                    Set rsTemp = GetRecordset(strSql)
                    
                    If rsTemp.EOF = False Then
                        strContent = strContent & Nvl(rsTemp("send_content").Value)
                    End If
                Next
            Else
                strSql = "Select a.��������.GetClobVal() as send_content From Ӱ�񱨸��¼ a Where a.ID = '" & strDocId & "'"
                
                Set rsTemp = GetRecordset(strSql)
                    
                If rsTemp.EOF = False Then
                    strContent = Nvl(rsTemp("send_content").Value)
                End If
            End If
        End If
        
        If strContent = "" Then
            MsgBox "�������ݲ����ڡ�"
            Exit Sub
        End If
        
        Call LoadReportContent(strContent, strDocId)
    Else
        rtxtReport.Text = ""
        rtxtReport.TextRTF = ""
        rtxtReport.Visible = True
        zlDocEditor.Visible = False
        mblnIsNewReport = False
        If vsfReportList.RowSel <= 0 Then Exit Sub
        
        '�ϰ汨��
        '�ж�ʹ�õ��Ӳ����༭������Pscs�༭��
        strSql = "select ����ֵ from ����ҽ������ a,Ӱ�����¼ b,Ӱ�����̲��� c " & _
                 "where a.ҽ��id = b.ҽ��id and b.ִ�п���id = c.����id and a.����id=" & strDocId & " and c.������='�鿴��ʷ����'"
        
        Set rsTemp = GetRecordset(strSql)
        
        If rsTemp.RecordCount > 0 Then intReportEditor = Nvl(rsTemp!����ֵ, "1")
        
        If intReportEditor = 0 Then
            Call LoadRichReportContent(strDocId)
        Else
            Call LoadPacsReportContent(strDocId)
        End If
    End If
End Sub

Private Sub LoadRichReportContent(ByVal strDocId As String)
    Dim tmpPath As String
    Dim objFSO As New Scripting.FileSystemObject
    Dim strBlobFile As String
    
    tmpPath = App.Path()
    
    If objFSO.FileExists(tmpPath & "\tmp.zip") Then Call objFSO.DeleteFile(tmpPath & "\tmp.zip")
    If objFSO.FileExists(tmpPath & "\tmp.rtf") Then Call objFSO.DeleteFile(tmpPath & "\tmp.rtf")
        
    strBlobFile = gobjComLib.Sys.ReadLob(100, 5, strDocId, Replace(tmpPath & "\tmp.zip", "\\", "\"))
    If objFSO.FileExists(strBlobFile) = False Then
        strBlobFile = gobjComLib.Sys.ReadLob(100, 5, strDocId, Replace(tmpPath & "\tmp.zip", "\\", "\"), 0, 1)
    End If
     
    
    If objFSO.FileExists(strBlobFile) Then
        '��ѹ��������ʾ
        Dim cc As New cUnzip
        
        cc.ZipFile = strBlobFile
        cc.UnzipFolder = tmpPath
        cc.Unzip
        Set cc = Nothing
        
        If objFSO.FileExists(tmpPath & "\tmp.rtf") Then
            Call rtxtReport.LoadFile(tmpPath & "\tmp.rtf")
        End If
    End If
End Sub

Private Sub LoadPacsReportContent(ByVal strDocId As String)
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext As String
    Dim blnShow As Boolean
    Dim strOffset As String
    Dim strTitle As String
    Dim strText As String
    Dim str������� As String
    Dim str������ As String
    Dim str���� As String
    
    strOffset = "  "
    
    '��ȡ������ⶨ��
    strSql = "Select ����ֵ,������ From ����ҽ������ a, Ӱ�����¼ b, Ӱ�����̲��� c Where a.ҽ��id = b.ҽ��id " & _
            " And b.ִ�п���id = c.����id And a.����id = " & Val(strDocId)
    Set rsTemp = GetRecordset(strSql)
    If Not rsTemp.EOF Then
        rsTemp.Filter = "������='�����������'"
        If Not rsTemp.EOF Then
            str������� = rsTemp!����ֵ
        Else
            str������� = "�������"
        End If
        
        rsTemp.Filter = "������='����������'"
        If Not rsTemp.EOF Then
            str������ = rsTemp!����ֵ
        Else
            str������ = "������"
        End If
        
        rsTemp.Filter = "������='��������'"
        If Not rsTemp.EOF Then
            str���� = rsTemp!����ֵ
        Else
            str���� = "����"
        End If
    End If
    
    '��ȡ���������
    strSql = "Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b " & _
             " Where a.�ļ�id = '" & strDocId & "' And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0  "
    Set rsTemp = GetRecordset(strSql)
    
    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs20 "
                
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!����
            Case "�������"
                strTitle = str�������
                strText = vbCrLf & strOffset & Nvl(rsTemp!����) & vbCrLf & vbCrLf
                blnShow = True
            Case "������"
                strTitle = str������
                strText = vbCrLf & strOffset & Nvl(rsTemp!����) & vbCrLf & vbCrLf
                blnShow = True
            Case "����"
                strTitle = str����
                strText = vbCrLf & strOffset & Nvl(rsTemp!����) & vbCrLf & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then
            strFormatContext = strFormatContext & "\b\cf2" & strTitle & "\par\b0\cf0" & Replace(strText, vbCrLf, "\par\cf0") & "\par"
        End If
        rsTemp.MoveNext
    Wend
    
    strFormatContext = strFormatContext & "}"
    
    rtxtReport.TextRTF = strFormatContext
End Sub

Public Function zlDocGetList(ByVal lngPatId As Long, Optional ByVal lngPageId As Long, Optional ByVal strRegNo As String) As Recordset
'����Pacs�ĵ��༭����д�ı����б�
'lngPatID:����ID
'lngPageID:��ҳID
'strRegNo:�Һŵ�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrH
    
    If Trim(strRegNo) <> "" Then
        strSql = "Select A.ID As ҽ��ID, RawToHex(B.��鱨��ID) As ����ID, C.�ĵ�����, C.����״̬, C.���༭ʱ��, C.���༭�� " & _
                 "From ����ҽ����¼ A, ����ҽ������ B, Ӱ�񱨸��¼ C " & _
                 "Where A.����ID=" & lngPatId & " And A.�Һŵ� = '" & strRegNo & "' And A.ID = B.ҽ��ID And B.��鱨��ID = C.ID"
    Else
        strSql = "Select A.ID As ҽ��ID, RawToHex(B.��鱨��ID) As ����ID, C.�ĵ�����, C.����״̬, C.���༭ʱ��, C.���༭�� " & _
                 "From ����ҽ����¼ A, ����ҽ������ B, Ӱ�񱨸��¼ C " & _
                 "Where A.����ID=" & lngPatId & " And A.��ҳID = " & lngPageId & " And A.ID = B.ҽ��ID And B.��鱨��ID = C.ID"
    End If
    
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "���ݲ��˻�ȡ�����б�") ' GetRecordset(strSql)
    
    Set zlDocGetList = rsTemp
    
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetList = Nothing
End Function

Public Function zlDocGetListWithAdvice(ByVal strAdviceId As String) As Recordset
'����Pacs�ĵ��༭����д�ı����б�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo ErrH
    
    If Trim(strAdviceId) = "" Then Exit Function
         
    strAdviceId = Replace(strAdviceId, "��", ",")
    
    strSql = "Select A.ҽ��ID, RawToHex(A.��鱨��ID) As ����ID, B.�ĵ�����, B.����״̬, B.���༭ʱ��, B.���༭�� " & _
             "From ����ҽ������ A, Ӱ�񱨸��¼ B, Table(Cast(f_Str2list('" & strAdviceId & "') As zlTools.t_Strlist)) C " & _
             "Where A.ҽ��ID = C.Column_Value And  A.��鱨��ID = B.ID "
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "����ҽ����ȡ�����б�") 'GetRecordset(strSql)
    
    Set zlDocGetListWithAdvice = rsTemp
    Exit Function
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
    Set zlDocGetListWithAdvice = Nothing
End Function

Public Sub zlDocViewStateUpdate(ByVal blnState As Boolean, ByVal lngAdviceID As Long)
'����pacs�ĵ��༭������Ĳ���״̬,���blnState=True����ʾ��Ӧ�����б���Ϊ���ģ�blnState=False,��ʾΪδ��
'lngPatID:����ID
'lngPageID:��ҳID
'strRegNo:�Һŵ�
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strUserName As String
    
On Error GoTo ErrH
    
    strSql = "Select RawToHex(B.��鱨��ID) As ����ID From ����ҽ������ B Where ҽ��ID = " & lngAdviceID
             
    Set rsTemp = GetRecordset(strSql)
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    If GetUserInfo Then strUserName = UserInfo.����
    
    Do While Not rsTemp.EOF
        If Nvl(rsTemp!����ID) <> "" Then
            If blnState Then
                strSql = "Zl_Ӱ�񱨸���ļ�¼_Insert(" & lngAdviceID & ", '" & Nvl(rsTemp!����ID) & "')"
            Else
                strSql = "Zl_Ӱ�񱨸���ļ�¼_Cancel(" & lngAdviceID & ", '" & Nvl(rsTemp!����ID) & "','" & strUserName & "')"
            End If
            
            Call gobjComLib.zlDatabase.ExecuteProcedure(strSql, gstrSysName)
        End If
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrH:
    If gobjComLib.ErrCenter() = 1 Then Resume
End Sub

Public Function GetReportImage(ByVal lngAdviceID As Long, ByVal strDocReportId As String, _
    strTmpImgDir As String, Optional ByRef strMsg As String = "") As Collection
'����:�����ҽ������Ӧ�����е�����ͼ�񣬱��浽ָ��Ŀ¼,��ȡPACS�����е�����ͼ��
'lngAdviceId:ҽ��id
'strTmpImgFolder:ͼ�񻺴�Ŀ¼
'objImgFileName:����ͼ���ļ�������

'˵�����˹���Ŀǰֻ�������ã�һ��ҽ����Ӧһ�ݱ���
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    Dim strContent As String
    Dim strDocId As String
    Dim intLoop As Integer
    Dim strFilter As String
    
    If (lngAdviceID <= 0) Then Exit Function
        
    If Trim(strDocReportId) = "" Then
        strFilter = "a.ҽ��ID=" & lngAdviceID & ""
    Else
        strFilter = "a.ID=HexToRaw('" & strDocReportId & "')"
    End If
    
    strSql = "Select Length(a.��������.GetClobVal()) as ContentLength, A.ID From Ӱ�񱨸��¼ a Where " & strFilter
    Set rsTemp = GetRecordset(strSql)
    
    If rsTemp.RecordCount <= 0 Then Exit Function

    strDocId = Nvl(rsTemp!ID)
    
    If rsTemp("ContentLength").Value > 2000 Then
        For intLoop = 1 To rsTemp("ContentLength").Value / 2000 + 1
            strSql = "select to_char(substr(a.��������.getclobval()," & CDbl(intLoop) * 2000 - 1999 & ",2000)) as send_content " & _
                     " from Ӱ�񱨸��¼ a where " & strFilter
                     
            Set rsTemp = GetRecordset(strSql)
            
            If rsTemp.EOF = False Then
                strContent = strContent & Nvl(rsTemp("send_content").Value)
            End If
        Next
    Else
        strSql = "Select a.��������.getclobval() as send_content From Ӱ�񱨸��¼ a Where " & strFilter
        
        Set rsTemp = GetRecordset(strSql)
            
        If rsTemp.EOF = False Then
            strContent = Nvl(rsTemp("send_content").Value)
        End If
    End If
    
    If strContent = "" Then
        strMsg = "�������ݲ����ڡ�"
        Exit Function
    End If
    
    '��ʼ��FTP�����Ϣ
    If InitFtpInfo(strDocId) = False Then
        strMsg = "��ȡͼ���FTP��Ϣʧ��"
        Exit Function
    End If
    
    Set GetReportImage = LoadImageFromXml(strContent, strTmpImgDir, strMsg)
End Function

Public Sub RefreshForm(ByVal lngAdviceID As Long, Optional ByVal strReportId As String = "", Optional objParent As Object)
    '���ܣ�ˢ�´���Ĳ���
    '������ lngAdviceId -- ҽ��ID
    '       strReportId -- ����ID
    '       blnAutoView -- �Ƿ��Զ�����
    '       objParent -- ������
    
    Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    mlngAdviceId = lngAdviceID
    mstrReportId = strReportId
    
    '�����RIS�鿴�����ÿ�ʼʱ��
    If mblnRIS = True Then
        dtpStart.Value = Format(Now - 365, "yyyy-mm-dd HH:MM")
    End If
    
    '��������˱���ID����ֻ��ʾһ�ݱ��棬������ʾ��ߵĹ��������Ͳ˵�
    If strReportId <> "" Then mblnShow = False
    
    strSql = "Select ����ֵ From Ӱ�����̲��� a, Ӱ�����¼ b " & _
             "Where a.����ID = b.ִ�п���id And a.������ = 'ҽ��վ�鿴����' And b.ҽ��id = [1]"
    Set rsTemp = gobjComLib.zlDatabase.OpenSQLRecord(strSql, "", lngAdviceID)
    
    If rsTemp.RecordCount > 0 Then
        mlngViewReport = Val(Nvl(rsTemp!����ֵ))
    Else
        mlngViewReport = 1
    End If
    
    Call LoadReport(GetFilter, True)
    
    Exit Sub
err:
    MsgBox err.Description, vbExclamation, gstrSysName
End Sub

Private Sub ClearReportContent()
    '���ܣ���ձ���༭��
    rtxtReport.Visible = True
    zlDocEditor.Visible = False
    rtxtReport.TextRTF = ""
    rtxtReport.Text = ""
End Sub

Private Sub SetReportlistDataSource(rsData As ADODB.Recordset)
    '���ܣ����ñ����б������Դ��ͬʱˢ�±���༭������
    '������rsData -- ����Դ
    
    Set vsfReportList.DataSource = rsData
    '����в�ѯ�����ǿ��ˢ��һ������
    If rsData.EOF = False Then
        Call vsfReportList_SelChange
    Else
        '��ձ༭��
        Call ClearReportContent
    End If
End Sub
