VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmQCContrast 
   Caption         =   "������Ŀ�ȶ�"
   ClientHeight    =   7275
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   Icon            =   "frmQCContrast.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   11760
   StartUpPosition =   3  '����ȱʡ
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   7110
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox CboItem 
      Height          =   300
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   210
      Width           =   3015
   End
   Begin VB.PictureBox PicChar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   240
      ScaleHeight     =   2415
      ScaleWidth      =   8955
      TabIndex        =   8
      Top             =   4140
      Width           =   8955
      Begin C1Chart2D8.Chart2D ChtThis 
         Height          =   1875
         Left            =   180
         TabIndex        =   9
         Top             =   150
         Width           =   8445
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   14896
         _ExtentY        =   3307
         _StockProps     =   0
         ControlProperties=   "frmQCContrast.frx":058A
      End
   End
   Begin VB.PictureBox PicData 
      BorderStyle     =   0  'None
      Height          =   2625
      Left            =   240
      ScaleHeight     =   2625
      ScaleWidth      =   8925
      TabIndex        =   7
      Top             =   1260
      Width           =   8925
      Begin VSFlex8Ctl.VSFlexGrid Vsf 
         Height          =   2010
         Left            =   300
         TabIndex        =   10
         Top             =   120
         Width           =   8295
         _cx             =   14631
         _cy             =   3545
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
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
         AutoResize      =   0   'False
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
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.ComboBox CboMachine 
      Height          =   300
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   210
      Width           =   3105
   End
   Begin VB.PictureBox PicCondition 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   300
      ScaleHeight     =   375
      ScaleWidth      =   11415
      TabIndex        =   0
      Top             =   720
      Width           =   11415
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   6570
         TabIndex        =   17
         Top             =   -60
         Width           =   2175
         Begin VB.OptionButton OptData 
            Caption         =   "�ȶ�ͼ"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   150
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton OptData 
            Caption         =   "ƫ��ͼ"
            Height          =   255
            Index           =   1
            Left            =   1170
            TabIndex        =   18
            Top             =   150
            Width           =   915
         End
      End
      Begin VB.OptionButton OptChart 
         Caption         =   "����ͼ"
         Height          =   195
         Index           =   1
         Left            =   10140
         TabIndex        =   14
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton OptChart 
         Caption         =   "����ͼ"
         Height          =   195
         Index           =   0
         Left            =   9060
         TabIndex        =   13
         Top             =   120
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.ComboBox CboLevel 
         Height          =   300
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   30
         Width           =   1125
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   4
         Top             =   30
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   133365763
         CurrentDate     =   39167
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   315
         Index           =   1
         Left            =   4860
         TabIndex        =   6
         Top             =   30
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   133365763
         CurrentDate     =   39167
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Left            =   4560
         TabIndex        =   5
         Top             =   90
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   2280
         TabIndex        =   3
         Top             =   90
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�ˮƽ"
         Height          =   180
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   6900
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCContrast.frx":0B0D
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15663
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
   Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
      Height          =   1560
      Left            =   9420
      TabIndex        =   16
      Top             =   1650
      Visible         =   0   'False
      Width           =   1635
      _cx             =   2884
      _cy             =   2752
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   14737632
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCContrast.frx":139F
      Left            =   975
      Top             =   330
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCContrast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String                 'Ȩ��

Const conPane_Condition = 201
Const conPane_Data = 202
Const conPane_Chart = 203

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar
Dim mlngMachineID As Long

Private Sub CboItem_Click()
    RefreshData
End Sub

Private Sub CboLevel_Click()
    If Me.CboItem.ListCount > 0 Then
        RefreshData
    End If
End Sub

Private Sub cboMachine_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngMachineID As Long
                                        
    If Me.CboMachine.ListCount = 0 Then Exit Sub            'û����ʱ�˳�
    lngMachineID = Me.CboMachine.ItemData(Me.CboMachine.ListIndex)
    
    gstrSql = "Select a.��Ŀid ,b.����, b.������ From ����������Ŀ a , ����������Ŀ b" & vbNewLine & _
              "Where a.��Ŀid = b.Id  And a.����id = [1] order by b.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, lngMachineID)
    Me.CboItem.Clear
    Do Until rsTmp.EOF
        With Me.CboItem
            .AddItem rsTmp("����") & "-" & rsTmp("������")
            .ItemData(.NewIndex) = rsTmp("��ĿId")
        End With
        rsTmp.MoveNext
    Loop
    
    If Me.CboItem.ListCount > 0 Then Me.CboItem.ListIndex = 0
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_PrintSet:
            Call zlPrintSet
        Case conMenu_File_Print
            With Me.ChtThis
                .PrintChart oc2dFormatBitmap, oc2dScaleToFit, 0, 0, 0, 0
            End With
        Case conMenu_File_BatPrint:
            Call zlRptPrint(1)
        Case conMenu_Edit_Save
            With Me.comDlg
                .CancelError = True
                .DialogTitle = "���Ϊ"
                .filter = "(ͼ���ļ�)|*.jpg"
                .FileName = Me.Caption & Format(Me.dtp����(1), "yyyyMMdd") & ".jpg"
                Err = 0: On Error Resume Next
                .ShowSave
                If Err <> 0 Then Exit Sub
                If .FileName = "" Then Exit Sub
                Me.ChtThis.SaveImageAsJpeg .FileName, 100, False, False, False
            End With
        Case conMenu_Edit_MarkMap
            Me.ChtThis.CopyToClipboard (oc2dFormatBitmap)
    
        Case conMenu_File_Exit:
            Unload Me
    
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
        Case conMenu_View_Refresh
            RefreshData
        Case conMenu_Help_Help:
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_Help_Web_Home:
            Call zlHomePage(Me.hWnd)
        Case conMenu_Help_Web_Mail:
            Call zlMailTo(Me.hWnd)
        Case conMenu_Help_About:
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else

            If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub


Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    If Me.Visible = False Then Exit Sub

    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Print, conMenu_File_BatPrint, conMenu_Edit_Save, conMenu_Edit_MarkMap:
            Control.Enabled = (Me.vfgRecord.Rows > Me.vfgRecord.FixedRows)
        Case conMenu_View_ToolBar_Button:
            Control.Checked = Me.cbsThis(2).Visible
        Case conMenu_View_ToolBar_Text:
            Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size:
            Control.Checked = Me.cbsThis.Options.LargeIcons
        Case conMenu_View_StatusBar:
            Control.Checked = Me.stbThis.Visible
    End Select
    
End Sub

Private Sub ChtThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim px As Long
    Dim py As Long
    Dim Series As Long
    Dim Point As Long
    Dim Distance As Long
    Dim Region As Long
    
    On Error Resume Next
    
    px = X / Screen.TwipsPerPixelX
    py = Y / Screen.TwipsPerPixelY
    
    If (Button = 0) Then
        With ChtThis
            Region = .ChartGroups(1).CoordToDataIndex(px, py, oc2dFocusXY, Series, Point, Distance)
            If (Series > 0 And Point > 0) And (Distance <= 5) Then
                If (Region = oc2dRegionInChartArea) Then
                    .ToolTipText = .ChartGroups(1).Data(Series, Point)
                End If
            Else
                .ToolTipText = ""
                .Footer.Text = ""
            End If
            .Refresh
        End With
    End If
End Sub


Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_Condition
            Item.Handle = Me.PicCondition.hWnd
        Case conPane_Data
            Item.Handle = Me.PicData.hWnd
        Case conPane_Chart
            Item.Handle = Me.PicChar.hWnd
    End Select
End Sub

Private Sub dtp����_Change(Index As Integer)
    
    If Me.dtp����(1).Value < Me.dtp����(0).Value Then Me.dtp����(0).Value = Me.dtp����(1).Value
    
    RefreshData
    
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
    mstrPrivs = gstrPrivs
    
    Me.PicCondition.BackColor = Me.BackColor
    Me.PicChar.BackColor = Me.BackColor
    Me.PicCondition.BackColor = Me.BackColor
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
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
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ����ͼ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "��ӡ�ʿؽ��(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "������ͼ(&S)..."): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "���ƿ���ͼ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
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
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��������")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "��������")
    cbrCustom.Handle = Me.CboMachine.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��Ŀ")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "��Ŀ")
    cbrCustom.Handle = Me.CboItem.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_Edit_MarkMap
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "���Ϊ"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    Call zlDatabase.ShowReportMenu(Me.cbsThis, glngSys, glngModul, mstrPrivs)
    
    '-----------------------------------------------------
    '����ͣ������
    Dim panThis As Pane, panChild As Pane
    
    With Me.dkpMan
        Set panThis = .CreatePane(conPane_Condition, 200, 400, DockTopOf, Nothing)
        panThis.Title = "����"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panThis = .CreatePane(conPane_Data, 400, 300, DockBottomOf, Nothing)
        panThis.Title = "�����б�"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panChild = .CreatePane(conPane_Chart, 400, 700, DockBottomOf, panThis)
        panChild.Title = "ͼ��"
        panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = True
    End With
    
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    
    '-----------------------------------------------------
    'װ���������
    Dim rsTmp As New ADODB.Recordset

    Me.dtp����(1).Value = Date: Me.dtp����(0).Value = DateAdd("m", -1, Date)
    
    With Me.CboLevel
        .AddItem "�ȶ�1"
        .ItemData(.NewIndex) = 1
        .AddItem "�ȶ�2"
        .ItemData(.NewIndex) = 2
        .AddItem "�ȶ�3"
        .ItemData(.NewIndex) = 3
        .AddItem "�ȶ�4"
        .ItemData(.NewIndex) = 4
        .AddItem "�ȶ�5"
        .ItemData(.NewIndex) = 5
        .ListIndex = 0
    End With
    
    
    Err = 0: On Error GoTo ErrHand
    
    With Me.ChtThis.ChartGroups(1).Data
        .NumSeries = 0
    End With
    
    gstrSql = "select ID ,����, ���� from �������� order by ���� "
    zlDatabase.OpenRecordset rsTmp, gstrSql, gstrSysName
        
    Me.CboMachine.Clear
    Do Until rsTmp.EOF
        With Me.CboMachine
            .AddItem rsTmp("����") & "-" & rsTmp("����")
            .ItemData(.NewIndex) = rsTmp("Id")
            If rsTmp("ID") = mlngMachineID Then .ListIndex = .NewIndex
        End With
        rsTmp.MoveNext
    Loop
    If Me.CboMachine.ListCount = 0 Then
        MsgBox "��δ����������ã�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    If Me.CboMachine.ListIndex < 0 Then Me.CboMachine.ListIndex = 0
    If Me.CboMachine.ListCount = 1 Then Me.CboMachine.Enabled = False
    
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim panThis As Pane
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    Set panThis = Me.dkpMan.FindPane(conPane_Condition)
    panThis.MinTrackSize.SetSize panThis.MinTrackSize.Width, 375 / Screen.TwipsPerPixelX
    panThis.MaxTrackSize.SetSize panThis.MinTrackSize.Width, 375 / Screen.TwipsPerPixelX
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    
    With Me.ChtThis
        '��1
        .ChartLabels(1).AttachCoord.X = .Header.Location.Left + (.ChartLabels(1).Location.Width / 2) - 150
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 30
        
        '��2
        .ChartLabels(2).AttachCoord.X = .Header.Location.Left + (.ChartLabels(2).Location.Width / 2) - 150
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        
        '��3
        .ChartLabels(3).AttachCoord.X = .Header.Location.Left + (.ChartLabels(3).Location.Width / 2) - 150
        .ChartLabels(3).AttachCoord.Y = .ChartLabels(2).Location.Top + .ChartLabels(2).Location.Height + 10
        
        '��3
        .ChartLabels(4).AttachCoord.X = .Header.Location.Left + (.ChartLabels(4).Location.Width / 2) - 150
        .ChartLabels(4).AttachCoord.Y = .ChartLabels(3).Location.Top + .ChartLabels(3).Location.Height + 10
    End With
End Sub

Private Sub OptChart_Click(Index As Integer)
    RefreshData
End Sub

Private Sub OptData_Click(Index As Integer)
    RefreshData
End Sub

Private Sub PicChar_Resize()
    On Error Resume Next
    With ChtThis
        .Top = 0
        .Left = 0
        .Width = PicChar.ScaleWidth
        .Height = PicChar.ScaleHeight
    End With
End Sub

Private Sub picData_Resize()
    On Error Resume Next
    With Vsf
        .Top = 0
        .Left = 0
        .Width = PicData.ScaleWidth
        .Height = PicData.ScaleHeight
    End With
End Sub

Private Sub RefreshData()
    '���ܣ�         ˢ������
    Dim rsMachine As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim intDateCount As Integer                 '��������
    Dim intLoop As Integer                      'ѭ��ʱʹ��
    Dim dateNow As Date                         '��ǰʱ������ѭ����ʹ��
    Dim intCol As Integer                       '��������
    Dim intRow As Integer                       '��������
    Dim intRowCount As Integer                  '������
    Dim strHearCaption As String                'ͷ����
    Dim intTmp As Integer                       '��ʱ��¼����
    Dim dblMax As Double                        '���ֵ
    Dim intCount As Integer
        
    If Me.CboItem.ListCount = 0 Then
        MsgBox "û���ҵ��ȶ���Ŀ��������ѡ��������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errH
    
    '��������
    intDateCount = DateDiff("d", Me.dtp����(0), Me.dtp����(1))
'    Me.Vsf.BackColor = &H80000005
    Me.Vsf.BackColorFixed = &HFDD6C6
    Me.Vsf.Cols = intDateCount + 2
    Me.Vsf.ColWidth(0) = 2500
    Me.Vsf.TextMatrix(1, 0) = Me.CboMachine.Text
    Me.Vsf.TextMatrix(0, 0) = "����"
    Me.Vsf.Rows = 1
    Me.Vsf.Rows = 2
    
    dateNow = Me.dtp����(0)
    For intLoop = 1 To intDateCount + 1
        With Me.Vsf
            .TextMatrix(0, intLoop) = Format(dateNow, "MMDD")
            .ColWidth(intLoop) = 700
            dateNow = DateAdd("d", 1, dateNow)
        End With
    Next
    
    
    '���ز�������
    gstrSql = "Select ����ʱ��, a.����id ,b.������" & vbNewLine & _
                " From ����걾��¼ a, ������ͨ��� b" & vbNewLine & _
                " Where  ����ʱ�� Between [1] And [2]" & vbNewLine & _
                "      And a.Id = b.����걾id And b.������Ŀid = [3] And Nvl(B.���ý��,0)=0 And �Ƿ��ʿ�Ʒ = [4] And a.����id = [5] " & vbNewLine & _
                " Order By ����ʱ��"

    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, CDate(Format(Me.dtp����(0), "yyyy-mm-dd 00:00:00")), _
                CDate(Format(Me.dtp����(1), "yyyy-mm-dd 23:59:59")), Me.CboItem.ItemData(Me.CboItem.ListIndex), _
                -Me.CboLevel.ItemData(Me.CboLevel.ListIndex), Me.CboMachine.ItemData(Me.CboMachine.ListIndex))
    intRow = 1
    
    Do Until rsData.EOF
        Me.Vsf.RowData(intRow) = Val(rsData("����Id"))
        Me.Vsf.TextMatrix(1, 0) = Me.CboMachine.Text
        intCol = FillVsfVal(Format(rsData("����ʱ��"), "MMDD"))
        If intCol <> 0 Then
            Me.Vsf.TextMatrix(intRow, intCol) = Nvl(rsData("������"))
        End If
        rsData.MoveNext
    Loop
    
    Me.Vsf.Cell(flexcpBackColor, intRow, 0, intRow, Me.Vsf.Cols - 1) = &HF2F9EE
    
    '������������
'    gstrSql = "Select Distinct b.id,b.����  From ����������Ŀ a,�������� b" & vbNewLine & _
'                "Where a.����id = b.Id And a.��Ŀid = [1] And b.id <> [2] "
'
'    Set rsMachine = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, Me.CboItem.ItemData(Me.CboItem.ListIndex), _
'                                        Me.CboMachine.ItemData(Me.CboMachine.ListIndex))
    gstrSql = "Select  distinct c.id,c.����  " & vbNewLine & _
                    " From ����걾��¼ a, ������ͨ��� b,�������� C " & vbNewLine & _
                    " Where  ����ʱ�� Between [1] And [2]" & vbNewLine & _
                    "      And a.Id = b.����걾id And b.������Ŀid = [3] And Nvl(B.���ý��,0)=0  And �Ƿ��ʿ�Ʒ = [4] And a.����id <> [5] " & vbNewLine & _
                    "      And a.����id = c.id "
    Set rsMachine = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, CDate(Format(Me.dtp����(0), "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.dtp����(1), "yyyy-mm-dd 23:59:59")), Me.CboItem.ItemData(Me.CboItem.ListIndex), _
                    -Me.CboLevel.ItemData(Me.CboLevel.ListIndex), Me.CboMachine.ItemData(Me.CboMachine.ListIndex))
    

    Me.Vsf.Rows = rsMachine.RecordCount * 2 + 2
    Do Until rsMachine.EOF
        If intRow = 1 Then
            intRow = intRow + 1
        Else
            intRow = intRow + 2
        End If
        
        If (intRow) / 2 Mod 2 = 0 Then
            Me.Vsf.Cell(flexcpBackColor, intRow, 0, intRow + 1, Me.Vsf.Cols - 1) = &HF2F9EE
        End If
        
        Me.Vsf.TextMatrix(intRow, 0) = rsMachine("����") & "(���)"
        Me.Vsf.TextMatrix(intRow + 1, 0) = rsMachine("����") & "(ƫ����)"
        '���ز�������
        gstrSql = "Select ����ʱ��, a.����id ,b.������" & vbNewLine & _
                    " From ����걾��¼ a, ������ͨ��� b" & vbNewLine & _
                    " Where  ����ʱ�� Between [1] And [2]" & vbNewLine & _
                    "      And a.Id = b.����걾id And Nvl(B.���ý��,0)=0 And b.������Ŀid = [3] And �Ƿ��ʿ�Ʒ = [4] And a.����id = [5] " & vbNewLine & _
                    " Order By ����ʱ��"

        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, CDate(Format(Me.dtp����(0), "yyyy-mm-dd 00:00:00")), _
                    CDate(Format(Me.dtp����(1), "yyyy-mm-dd 23:59:59")), Me.CboItem.ItemData(Me.CboItem.ListIndex), _
                    -Me.CboLevel.ItemData(Me.CboLevel.ListIndex), CLng(rsMachine("Id")))

        Do Until rsData.EOF
            Me.Vsf.RowData(intRow) = Val(rsData("����Id"))
            intCol = FillVsfVal(Format(rsData("����ʱ��"), "MMDD"))
            If intCol <> 0 Then
                Me.Vsf.TextMatrix(intRow, intCol) = rsData("������")
                If intRow >= 2 Then
                    If Val(Nvl(rsData("������"))) <> 0 And Val(Me.Vsf.TextMatrix(1, intCol)) <> 0 Then
                        Me.Vsf.RowData(intRow + 1) = 0
                        Me.Vsf.TextMatrix(intRow + 1, intCol) = Format((Val(rsData("������")) - Me.Vsf.TextMatrix(1, intCol)) / Me.Vsf.TextMatrix(1, intCol) * 100, "###0.00") & "%"
                    End If
                End If
            End If
            rsData.MoveNext
        Loop
        rsMachine.MoveNext
    Loop


    '''''''''''''''''''''''''''��ͼ'''''''''''''''''''''''''''
    Dim aryX() As Variant, aryY() As Variant
    
    
    If Me.Vsf.Rows <= 2 Then Exit Sub
    With Me.ChtThis
        '��ʾ����
        .Reset
        .IsBatched = True
        .Header.Text = "����������Ŀ�ȶ�ͼ" & vbCrLf & " " & vbCrLf & " "
        .Header.Adjust = oc2dAdjustCenter
        .Header.Font.Bold = True
        .Header.Font.Size = 16
        .IsBatched = False
        .ChartLabels.RemoveAll
        
        '��1
        .ChartLabels.Add
        .ChartLabels(1).AttachMethod = oc2dAttachCoord
        intCount = 26 - Len("��������: " & Me.CboMachine.Text)
        If intCount < 0 Then intCount = 0
        strHearCaption = "��������:  " & Me.CboMachine.Text & Space(intCount)
        strHearCaption = strHearCaption & "�ȶ�����:"
    
        For intLoop = 2 To Me.Vsf.Rows - 1
            If InStr(Me.Vsf.TextMatrix(intLoop, 0), "(���)") > 0 Then
                strHearCaption = strHearCaption & "  " & Replace(Me.Vsf.TextMatrix(intLoop, 0), "(���)", "")
            End If
        Next
        .ChartLabels(1).Text = strHearCaption
        .ChartLabels(1).AttachCoord.X = (.ChartLabels(1).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2)
        .ChartLabels(1).AttachCoord.Y = .Header.Location.Top + .Header.Location.Height - 20
        
        '��2
        .ChartLabels.Add
        .ChartLabels(2).AttachMethod = oc2dAttachCoord
        .ChartLabels(2).Adjust = oc2dAdjustRight
        strHearCaption = ""
        
        intCount = 28 - Len("�ȶ���Ŀ:  " & Me.CboItem.Text)
        If intCount < 0 Then intCount = 0
        strHearCaption = "�ȶ���Ŀ:  " & Me.CboItem.Text & Space(intCount)
        strHearCaption = strHearCaption & "��    ��:  " & Format(Me.dtp����(0), "yyyy��mm��dd��") & "��" & Format(Me.dtp����(1), "yyyy��mm��dd��")
        intCount = 40 - Len("��    ��:  " & Format(Me.dtp����(0), "yyyy��mm��dd��") & "��" & Format(Me.dtp����(1), "yyyy��mm��dd��"))
        If intCount < 0 Then intCount = 0
        strHearCaption = strHearCaption & Space(intCount)
        gstrSql = " Select �ȶԾ�ʾ��,�ȶ�ʧ���� From ������Ŀ where ������ĿID = [1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, Me.CboItem.ItemData(Me.CboItem.ListIndex))
        If rsData.EOF = False Then
            strHearCaption = strHearCaption & " ��ʾ��:" & rsData("�ȶԾ�ʾ��") & "      ʧ����:" & rsData("�ȶ�ʧ����")

        End If
        .ChartLabels(2).Text = strHearCaption
        .ChartLabels(2).AttachCoord.X = (.ChartLabels(2).Location.Width / 2) + (.Width / Screen.TwipsPerPixelX / 2) - (.ChartLabels(1).Location.Width / 2)
        .ChartLabels(2).AttachCoord.Y = .ChartLabels(1).Location.Top + .ChartLabels(1).Location.Height + 10
        

'        .IsBatched = False
    End With
    


    With Me.ChtThis


        With .ChartGroups(1).Data

            If Me.OptData(0).Value = True Then
                '�ȶ�ͼ
                .NumSeries = 0
                .NumSeries = (Me.Vsf.Rows - 2) / 2 + 1
                .NumPoints(1) = Me.Vsf.Cols

                ReDim aryX(Me.Vsf.Cols - 1)
                ReDim aryY(Me.Vsf.Cols - 1, (Me.Vsf.Rows - 2) / 2)

                'X����ֵ
                With Me.ChtThis.ChartArea.Axes("X")
                    .AnnotationMethod = oc2dAnnotateValueLabels   '������2��ʾֵ��ʾ
                    With .ValueLabels
                        .RemoveAll
                        For intLoop = 1 To Me.Vsf.Cols - 1
                            .Add intLoop, Me.Vsf.TextMatrix(0, intLoop)
                        Next
                    End With
    '                .MajorGrid.Spacing = 1
                    .Max = Me.Vsf.Cols '- 1
                End With

                With Me.ChtThis.ChartArea.Axes("Y")
    '                .MajorGrid.Spacing = 1
                    .Min = 0
                End With

                'X
                For intLoop = 0 To Me.Vsf.Cols - 1
                    aryX(intLoop) = intLoop
                Next

                'Y
                intTmp = 0
                For intRow = 0 To Me.Vsf.Rows - 1
                    If InStr(Me.Vsf.TextMatrix(intRow, 0), "���") > 0 Or intRow = 1 Then
                        aryY(0, intTmp) = 1E+308
                        For intLoop = 1 To Me.Vsf.Cols - 1
                            aryY(intLoop, intTmp) = IIf(Me.Vsf.TextMatrix(intRow, intLoop) = "", 1E+308, Val(Me.Vsf.TextMatrix(intRow, intLoop)))
                        Next
                        intTmp = intTmp + 1
                    End If
                Next

            Else
                '�ȶ�ͼ
                If Me.OptChart(0).Value = True Then
                    .NumSeries = 0
                    .NumSeries = (Me.Vsf.Rows - 2) / 2 + 5
                    .NumPoints(1) = Me.Vsf.Cols

                    ReDim aryX(Me.Vsf.Cols - 1)
                    ReDim aryY(Me.Vsf.Cols - 1, (Me.Vsf.Rows - 2) / 2 + 4)
                Else
                    .NumSeries = 0
                    .NumSeries = (Me.Vsf.Rows - 2) / 2
                    .NumPoints(1) = Me.Vsf.Cols

                    ReDim aryX(Me.Vsf.Cols - 1)
                    ReDim aryY(Me.Vsf.Cols - 1, (Me.Vsf.Rows - 2) / 2 - 1)
                End If


                With Me.ChtThis.ChartArea.Axes("Y")
    '                .MajorGrid.Spacing = 1
                    gstrSql = "Select �ȶ�ʧ����,�ȶԾ�ʾ�� From ������ĿĿ¼ a , ���鱨����Ŀ b , ������Ŀ c " & _
                               " Where a.ID = b.������Ŀid And b.������Ŀid = c.������ĿId And a.�����Ŀ = 0 And c.������ĿId = [1] "
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, Me.CboItem.ItemData(Me.CboItem.ListIndex))

                    If IsNull(rsData(1)) = True Then
                        MsgBox "�뵽������Ŀ����������<�ȶ�ʧ���ʡ��ȶԾ�ʾ��>��", vbInformation, gstrSysName
                        Exit Sub
                    End If

                    .Min = -Val(Nvl(rsData(0), 0))
                    .Max = Val(Nvl(rsData(0), 0))
                    .Origin = -Val(Nvl(rsData(0), 0))

                    .AnnotationMethod = oc2dAnnotateValueLabels
                    With .ValueLabels
                        .RemoveAll
                        .Add 0, Me.Vsf.TextMatrix(1, 0)
                        .Add -Val(Nvl(rsData(0), 0)), "-ʧ����(" & -Val(Nvl(rsData(0), 0)) & "%)"
                        .Add Val(Nvl(rsData(0), 0)), "ʧ����(" & Val(Nvl(rsData(0), 0)) & "%)"
                        .Add -Val(Nvl(rsData(1), 0)), "-��ʾ��(" & -Val(Nvl(rsData(1), 0)) & "%)"
                        .Add Val(Nvl(rsData(1), 0)), "��ʾ��(" & Val(Nvl(rsData(1), 0)) & "%)"
                    End With
                End With

                'X����ֵ
                With Me.ChtThis.ChartArea.Axes("X")
                    .AnnotationMethod = oc2dAnnotateValueLabels   '������2��ʾֵ��ʾ
                    With .ValueLabels
                        .RemoveAll
                        For intLoop = 1 To Me.Vsf.Cols - 1
                            .Add intLoop, Me.Vsf.TextMatrix(0, intLoop)
                        Next
                    End With
'                    .MajorGrid.Spacing = 5
                    .Max = Me.Vsf.Cols - 1
                End With

                'X
                For intLoop = 0 To Me.Vsf.Cols - 1
                    aryX(intLoop) = intLoop
                Next

                'Y
                intTmp = 0
                For intRow = 2 To Me.Vsf.Rows - 1
                    If InStr(Me.Vsf.TextMatrix(intRow, 0), "(ƫ����)") > 0 Then
                        aryY(0, intTmp) = 1E+308
                        For intLoop = 1 To Me.Vsf.Cols - 1
                            With Me.ChtThis.ChartArea.Axes("Y")
                                If Val(Replace(Me.Vsf.TextMatrix(intRow, intLoop), "%", "")) > .Max Then
                                    dblMax = .Max
                                ElseIf Val(Replace(Me.Vsf.TextMatrix(intRow, intLoop), "%", "")) < .Min Then
                                    dblMax = .Min
                                ElseIf Me.Vsf.TextMatrix(intRow, intLoop) = "" Then
                                    dblMax = 1E+308
                                Else
                                    dblMax = Val(Replace(Me.Vsf.TextMatrix(intRow, intLoop), "%", ""))
                                End If
                                aryY(intLoop, intTmp) = dblMax
                            End With

'                            aryY(intLoop, intTmp) = IIf(Me.Vsf.TextMatrix(intRow, intLoop) = "", 1E+308, Replace(Me.Vsf.TextMatrix(intRow, intLoop), "%", ""))
                        Next
                        intTmp = intTmp + 1
                    End If
                Next

                If Me.OptChart(0).Value = True Then
                    '����λ��
                    '������
                    intTmp = (Me.Vsf.Rows - 2) / 2

                    For intLoop = 0 To Me.Vsf.Cols - 1
                        aryY(intLoop, intTmp) = 0
                    Next
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Symbol.Shape = oc2dShapeNone
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Line.COLOR = vbYellow

                    'ʧ������С��
                    intTmp = (Me.Vsf.Rows - 2) / 2 + 1
                    For intLoop = 0 To Me.Vsf.Cols - 1
                        aryY(intLoop, intTmp) = -Val(Nvl(rsData(0), 0))
                    Next
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Symbol.Shape = oc2dShapeNone
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Line.COLOR = vbRed

                    'ʧ���������
                    intTmp = (Me.Vsf.Rows - 2) / 2 + 2
                    For intLoop = 0 To Me.Vsf.Cols - 1
                        aryY(intLoop, intTmp) = Val(Nvl(rsData(0), 0))
                    Next
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Symbol.Shape = oc2dShapeNone
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Line.COLOR = vbRed

                    '��ʾ����С��
                    intTmp = (Me.Vsf.Rows - 2) / 2 + 3
                    For intLoop = 0 To Me.Vsf.Cols - 1
                        aryY(intLoop, intTmp) = -Val(Nvl(rsData(1), 0))
                    Next
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Symbol.Shape = oc2dShapeNone
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Line.COLOR = vbGreen

                    '��ʾ�������
                    intTmp = (Me.Vsf.Rows - 2) / 2 + 4
                    For intLoop = 0 To Me.Vsf.Cols - 1
                        aryY(intLoop, intTmp) = Val(Nvl(rsData(1), 0))
                    Next
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Symbol.Shape = oc2dShapeNone
                    Me.ChtThis.ChartGroups(1).Styles(intTmp + 1).Line.COLOR = vbGreen
                End If
            End If
            Call .CopyXVectorIn(1, aryX)
            Call .CopyYArrayIn(aryY)

        End With

        '��־
        If Me.OptData(0).Value = True Then
            For intLoop = 1 To Me.Vsf.Rows - 1
                If InStr(Me.Vsf.TextMatrix(intLoop, 0), "���") > 0 Or intLoop = 1 Then
                    .ChartGroups(1).SeriesLabels.Add Me.Vsf.TextMatrix(intLoop, 0)
                End If
            Next
        Else
            For intLoop = 1 To Me.Vsf.Rows - 1
                If InStr(Me.Vsf.TextMatrix(intLoop, 0), "(ƫ����)") > 0 Then
                    .ChartGroups(1).SeriesLabels.Add Me.Vsf.TextMatrix(intLoop, 0)
                End If
            Next
        End If
        .Legend.Anchor = oc2dAnchorSouth
        .Legend.Orientation = oc2dOrientHorizontal

        '��ʾ��ͬ��ͼ��
        .ChartGroups(1).ChartType = IIf(Me.OptChart(0).Value = True, oc2dTypePlot, oc2dTypeBar)

        .IsBatched = False
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function FillVsfVal(strDate As String) As Integer
    '����:          �ҵ���Ӧ��ʱ����������
    '����:          ���ڸ�ʽΪ("MMDD")
    '����:          �ڼ���
    Dim intLoop As Integer
    
    For intLoop = 1 To Vsf.Cols - 1
        If strDate = Vsf.TextMatrix(0, intLoop) Then
            FillVsfVal = intLoop
        End If
    Next
    
End Function
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.vfgRecord.Rows <= Me.vfgRecord.FixedRows Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.Vsf
    objPrint.Title.Text = Mid(Me.CboMachine.Text, InStr(1, Me.CboMachine.Text, ",") + 1) & "�ʿؽ���嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Public Sub ShowMe(objfrm As Object, MachineID As Long)
    Dim intLoop As Integer
    mlngMachineID = MachineID
    
    Me.Show vbModal, objfrm
End Sub
