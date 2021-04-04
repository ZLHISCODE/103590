VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "ZLRICHEDITOR.OCX"
Begin VB.Form frmEPRFileContent 
   BackColor       =   &H80000003&
   BorderStyle     =   0  'None
   Caption         =   "病历文件提纲"
   ClientHeight    =   10440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   14400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picWave 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   600
      ScaleHeight     =   3255
      ScaleWidth      =   6150
      TabIndex        =   7
      Top             =   3210
      Visible         =   0   'False
      Width           =   6150
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   1875
         Left            =   90
         TabIndex        =   11
         Top             =   1230
         Width           =   3030
         _cx             =   5345
         _cy             =   3307
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         BackColorFixed  =   -2147483634
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEPRFileContent.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         WallPaperAlignment=   4
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   1935
         TabIndex        =   8
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1050
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   0
         ScaleHeight     =   900
         ScaleWidth      =   1575
         TabIndex        =   10
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.PictureBox picTab 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   5430
      ScaleHeight     =   2985
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   3420
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   4170
         Left            =   -210
         TabIndex        =   1
         Top             =   780
         Width           =   7500
         _cx             =   13229
         _cy             =   7355
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
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
         Rows            =   4
         Cols            =   6
         FixedRows       =   3
         FixedCols       =   0
         RowHeightMin    =   250
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
         AutoSizeMode    =   1
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
      Begin VB.Label lblSubEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注:##"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "一般护理记录单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2970
         TabIndex        =   3
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
   End
   Begin VB.PictureBox picRich 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   1785
      ScaleHeight     =   3150
      ScaleWidth      =   4830
      TabIndex        =   4
      Top             =   60
      Width           =   4830
      Begin zlRichEditor.Editor edtThis 
         Height          =   2580
         Left            =   150
         TabIndex        =   5
         Top             =   75
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4551
         WithViewButtonas=   0   'False
         ShowRuler       =   0   'False
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   45
      Top             =   45
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRFileContent.frx":0062
      Left            =   120
      Top             =   645
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmEPRFileContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum zlEnumCompendParentKind     '提纲父类型
    cprEmCPKFileDefine = 0              '文件定义内容
    cprEmCPKModelEssay = 1              '范文内容
End Enum

Private Enum FileType
    conPane_RichEpr = 1
    conPane_TendEpr = 2
    conPane_TablEpr = 3
    conPane_Infection = 4
    conPane_WaveEpr = 5 '专科体温单页面
End Enum

Private msinVStep As Single      '滚动条的步长
Private msinHStep As Single      '滚动条的步长
'-----------------------------------------------------
'窗体事件
'-----------------------------------------------------
Public Event DblClick()                                                 '返回双击操作事件
Private mObjTabEprView As cTableEPR
Private mobjInfection As Object

'-----------------------------------------------------
'临时变量

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Control.Enabled = edtThis.Selection.EndPos <> edtThis.Selection.StartPos
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ID_EDIT_COPY
        Me.edtThis.Copy
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_RichEpr
        Item.Handle = picRich.hWnd
    Case conPane_TendEpr
        Item.Handle = picTab.hWnd
    Case conPane_TablEpr
        Item.Handle = mObjTabEprView.zlGetForm.hWnd
    Case conPane_Infection
        Item.Handle = mobjInfection.zlGetForm.hWnd
    Case conPane_WaveEpr
        Item.Handle = picWave.hWnd
    End Select
End Sub

Private Sub Form_Load()
    cbsThis.ActiveMenuBar.Visible = False
    cbsThis.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    
    Dim Pane1 As Pane, pane2 As Pane, pane3 As Pane, Pane4 As Pane, Pane5 As Pane
    Set mObjTabEprView = New cTableEPR
    mObjTabEprView.InitTableEPR gcnOracle, glngSys, gstrDbOwner
        
    Set Pane1 = dkpMan.CreatePane(conPane_RichEpr, 1200, 200, DockTopOf, Nothing)
    Pane1.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set pane2 = dkpMan.CreatePane(conPane_TendEpr, 1200, 200, DockTopOf, Nothing)
    pane2.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane2.Close
    
    Set pane3 = dkpMan.CreatePane(conPane_TablEpr, 1200, 200, DockTopOf, Nothing)
    pane3.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    pane3.Close
    
    Set Pane4 = dkpMan.CreatePane(conPane_Infection, 1200, 200, DockTopOf, Nothing)
    Pane4.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane4.Close
    
    Set Pane5 = dkpMan.CreatePane(conPane_WaveEpr, 1200, 200, DockTopOf, Nothing)
    Pane5.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    Pane5.Close
    
    With dkpMan
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        .Options.UseSplitterTracker = True
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = False
    End With
    
    Set mobjInfection = DynamicCreate("zlDisReportCard.clsDisReportCard", "传染病报告卡", True)
    If Not mobjInfection Is Nothing Then
        mobjInfection.Init gcnOracle, glngSys
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Unload mObjTabEprView.zlGetForm
    Set mObjTabEprView = Nothing
    Unload mobjInfection.zlGetForm
    Set mobjInfection.zlGetForm = Nothing
    Set mobjInfection = Nothing
End Sub
Private Sub picRich_Resize()
    edtThis.Top = 0: edtThis.Left = 0
    edtThis.Width = picRich.ScaleWidth: edtThis.Height = picRich.ScaleHeight
End Sub

Private Sub picTab_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Err = 0: On Error Resume Next
    Me.lblTitle.Move Me.picTab.ScaleLeft, Me.picTab.ScaleTop + 120, Me.picTab.ScaleWidth
    Me.lblSubhead.Move Me.picTab.ScaleLeft + 210, Me.lblTitle.Top + Me.lblTitle.Height + 120
    Me.vfgThis.Move Me.picTab.ScaleLeft + 210, Me.lblSubhead.Top + Me.lblSubhead.Height + 45, Me.picTab.ScaleWidth - 210 * 2
    Me.vfgThis.Height = Me.picTab.ScaleHeight - Me.vfgThis.Top - 210 - lblSubEnd.Height - 45
    Me.lblSubEnd.Move lblSubhead.Left, Me.vfgThis.Top + Me.vfgThis.Height + 45
End Sub

Private Sub edtThis_DblClick(ViewMode As zlRichEditor.ViewModeEnum)
    RaiseEvent DblClick
End Sub

Private Sub edtThis_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, X As Single, Y As Single)
    Dim Popup As CommandBar
    Dim Control As CommandBarControl
    
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Popup.ShowPopup
    End With
End Sub

'-----------------------------------------------------
'窗体公共方法
'-----------------------------------------------------

Public Sub zlRefresh(ByVal lngParentId As Long, Optional bytParentKind As zlEnumCompendParentKind = cprEmCPKFileDefine)
    '功能：显示指定文件/范文的内容；
    Dim strTemp As String, strZipFile As String
    Dim rsTemp As New ADODB.Recordset
    Dim mEPRFileInfo As cEPRFileDefineInfo
    Dim lngCount As Long
    Dim blnCollegeWave As Boolean '是否是专科体温单
    Dim lngTop As Long
    
    dkpMan.FindPane(conPane_TendEpr).Close
    dkpMan.FindPane(conPane_TablEpr).Close
    dkpMan.FindPane(conPane_WaveEpr).Close
    dkpMan.FindPane(conPane_Infection).Close
    dkpMan.ShowPane conPane_RichEpr
    Me.edtThis.ReadOnly = False
    Me.edtThis.NewDoc
    If lngParentId = 0 Then Me.edtThis.ReadOnly = True: Exit Sub
    Me.edtThis.Freeze
        
    If bytParentKind = cprEmCPKFileDefine Then '病历文件管理
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select b.种类, b.保留,B.编号,B.子类, a.格式" & vbNewLine & _
                "From 病历页面格式 a, 病历文件列表 b" & vbNewLine & _
                "Where a.种类 = b.种类 And a.编号 = b.页面 And b.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
    
        '设置页面格式
        Set mEPRFileInfo = New cEPRFileDefineInfo
        mEPRFileInfo.格式 = "" & rsTemp!格式
        mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
        Set mEPRFileInfo = Nothing
        Me.edtThis.ResetWYSIWYG
        
        If Val("" & rsTemp!保留) < 0 And rsTemp!种类 <> 6 Then
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            
            If rsTemp!种类 = 3 Then '体温单
                dkpMan.FindPane(conPane_TendEpr).Close
                dkpMan.FindPane(conPane_RichEpr).Close
                dkpMan.FindPane(conPane_TablEpr).Close
                dkpMan.FindPane(conPane_Infection).Close
                dkpMan.ShowPane conPane_WaveEpr
                Me.picWave.Visible = True
                VsfData.Visible = False
                msinVStep = 0: msinHStep = 0
                If NVL(rsTemp!子类) = "1" Then '专科体温单
                    gstrSQL = _
                        " SELECT Id, 文件id, 父id, 对象序号, 对象类型, 对象标记, 对象属性, 内容行次, 内容文本, 是否换行, 要素名称, 要素表示" & vbNewLine & _
                        " FROM 病历文件结构" & vbNewLine & _
                        " WHERE 文件id = [1]" & vbNewLine & _
                        " START WITH 父id IS NULL" & vbNewLine & _
                        " CONNECT BY PRIOR Id = 父id"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
                    blnCollegeWave = True
                Else '标准体温单展示"示例样式"
                    blnCollegeWave = False
                    Set rsTemp = GetPipWaveStyle(lngParentId)
                End If
                If rsTemp.RecordCount > 0 Then
                    picDraw.AutoRedraw = True
                    Call DrawWaveStyle(picDraw, rsTemp, Not blnCollegeWave, lngTop)
                    rsTemp.Filter = "要素名称 ='婴儿体温单'"
                    If rsTemp.RecordCount > 0 Then
                        If Val(rsTemp!内容文本) = 1 Then
                            rsTemp.Filter = ""
                            Call ShowTabBaby(rsTemp, lngTop)
                        End If
                    End If
                    Call CalcScrollBarSize
                End If
            Else
                dkpMan.FindPane(conPane_TendEpr).Close
                dkpMan.FindPane(conPane_WaveEpr).Close
                dkpMan.FindPane(conPane_TablEpr).Close
                dkpMan.FindPane(conPane_Infection).Close
                dkpMan.ShowPane conPane_RichEpr
                With Me.edtThis
                    .Text = vbCrLf & Space(4) & "该文件为特殊格式病历，不能浏览样式..."
                    .SelectAll
                    .ForceEdit = True
                    .Selection.Font.Name = "宋体": .Selection.Font.Size = 10.5
                    .SelLength = 0
                    .ForceEdit = False
                End With
            End If
        ElseIf NVL(rsTemp!保留, 0) = 2 Then
            With Me.edtThis
                .Text = vbCrLf & Space(4) & "该文件为表格式病历，正在读取文件样式..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "宋体": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TablEpr
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_病历文件定义, lngParentId, False, 0)
            Call mObjTabEprView.zlRefreshDockfrm '刷新显示
        ElseIf NVL(rsTemp!保留, 0) = 4 Then
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.ShowPane conPane_Infection
            Call mobjInfection.zlRefresh(0, 0, 0, False)
        ElseIf rsTemp!种类 = 3 Then
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TendEpr
            
            Dim lngCurColor As Long, strCurFont As String, objFont As StdFont
            Me.lblTitle.Caption = "": Me.lblSubhead.Caption = "": Me.lblSubEnd.Caption = ""
            Me.vfgThis.Redraw = flexRDNone
            Me.vfgThis.Clear: Me.vfgThis.MergeCells = flexMergeFixedOnly: vfgThis.MergeCellsFixed = flexMergeRestrictAll
            Me.vfgThis.MergeRow(0) = True
            Me.vfgThis.MergeRow(1) = True
            Me.vfgThis.MergeRow(2) = True
            Me.picTab.Visible = True
'
            gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
                " Order By d.对象序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Select Case "" & !要素名称
                    Case "表头层数"
                        If Val("" & !内容文本) = 1 Then
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = True
                            Me.vfgThis.RowHidden(2) = True
                        ElseIf Val("" & !内容文本) = 2 Then
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = False
                            Me.vfgThis.RowHidden(2) = True
                        Else
                            Me.vfgThis.RowHidden(0) = False
                            Me.vfgThis.RowHidden(1) = False
                            Me.vfgThis.RowHidden(2) = False
                        End If
                    Case "总列数":  Me.vfgThis.Cols = Val("" & !内容文本)
                    Case "最小行高": Me.vfgThis.RowHeightMin = Val("" & !内容文本)
                    Case "文本字体"
                        strCurFont = "" & !内容文本
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                        End With
                        Set Me.vfgThis.Font = objFont
                        Set Me.lblSubhead.Font = Me.vfgThis.Font
                        Set Me.lblSubEnd.Font = Me.vfgThis.Font
                        
                    Case "文本颜色": Me.vfgThis.ForeColor = Val("" & !内容文本)
                    Case "表格颜色": Me.vfgThis.GridColor = Val("" & !内容文本): Me.vfgThis.GridColorFixed = Me.vfgThis.GridColor
                    
                    Case "标题文本": Me.lblTitle.Caption = "" & !内容文本
                    Case "标题字体"
                        strCurFont = "" & !内容文本
                        Set objFont = New StdFont
                        With objFont
                            .Name = Split(strCurFont, ",")(0)
                            .Size = Val(Split(strCurFont, ",")(1))
                            .Bold = False: .Italic = False
                            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                        End With
                        Set Me.lblTitle.Font = objFont
                        Me.lblTitle.AutoSize = False
                    End Select
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
                " Order By d.对象序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubhead.Caption = ""
                Do While Not .EOF
                    Me.lblSubhead.Caption = Me.lblSubhead.Caption & " " & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
                    .MoveNext
                Loop
                Me.lblSubhead.Caption = Trim(Me.lblSubhead.Caption)
            End With
            
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表下标签'" & _
                " Order By d.对象序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Me.lblSubEnd.Caption = ""
                Do While Not .EOF
                    Me.lblSubEnd.Caption = Me.lblSubEnd.Caption & " " & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
                    .MoveNext
                Loop
                Me.lblSubEnd.Caption = Trim(Me.lblSubEnd.Caption)
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
                " Order By d.对象序号"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    vfgThis.TextMatrix(!内容行次 - 1, !对象序号 - 1) = "" & !内容文本
                    vfgThis.FixedAlignment(!对象序号 - 1) = flexAlignCenterCenter
                    .MoveNext
                Loop
            End With
            '---------------------------------------------------
            gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位" & _
                " From 病历文件结构 d, 病历文件结构 p" & _
                " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
                " Order By d.对象序号, d.内容行次"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
            With rsTemp
                Do While Not .EOF
                    Me.vfgThis.ColWidth(!对象序号 - 1) = Val("" & !对象属性)
                    .MoveNext
                Loop
            End With
            vfgThis.AutoSizeMode = flexAutoSizeRowHeight
            vfgThis.AutoSize 0, vfgThis.Cols - 1
            Me.vfgThis.Redraw = flexRDDirect
                    
            '---------------------------------------------------
            Call picTab_Resize
        Else
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_RichEpr
            strZipFile = zlBlobRead(1, lngParentId)
            If Len(strZipFile) > 0 Then
                If gobjFSO.FileExists(strZipFile) Then
                    strTemp = zlFileUnzip(strZipFile)
                    If gobjFSO.FileExists(strTemp) Then
                        Me.edtThis.OpenDoc strTemp
                        gobjFSO.DeleteFile strTemp, True
                    End If
                    gobjFSO.DeleteFile strZipFile, True
                End If
            End If
        End If
    Else '病历范文内容
        gstrSQL = "Select c.Id, c.性质, a.格式" & vbNewLine & _
            "From 病历页面格式 a, 病历文件列表 b, 病历范文目录 c" & vbNewLine & _
            "Where c.文件id = b.Id And b.种类 = a.种类 And b.页面 = a.编号 And c.Id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngParentId)
        If rsTemp.RecordCount <= 0 Then Exit Sub
        If NVL(rsTemp!性质, 0) = 2 Then
            With Me.edtThis
                .Text = vbCrLf & Space(4) & "该文件为表格式病历，暂不支持浏览样式..."
                .SelectAll
                .ForceEdit = True
                .Selection.Font.Name = "宋体": .Selection.Font.Size = 10.5
                .SelLength = 0
                .ForceEdit = False
            End With
            dkpMan.FindPane(conPane_RichEpr).Close
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_TablEpr
            Call mObjTabEprView.InitOpenEPR(Me, cprEM_修改, cprET_全文示范编辑, lngParentId, False, 0)
            Call mObjTabEprView.zlRefreshDockfrm '刷新显示
        Else
            dkpMan.FindPane(conPane_TendEpr).Close
            dkpMan.FindPane(conPane_TablEpr).Close
            dkpMan.FindPane(conPane_WaveEpr).Close
            dkpMan.FindPane(conPane_Infection).Close
            dkpMan.ShowPane conPane_RichEpr
            Set mEPRFileInfo = New cEPRFileDefineInfo
            mEPRFileInfo.格式 = "" & rsTemp!格式
            mEPRFileInfo.SetFormat Me.edtThis, mEPRFileInfo.格式
            Set mEPRFileInfo = Nothing
            Me.edtThis.ResetWYSIWYG
            
            If Val("" & rsTemp!性质) = 0 Then
                strZipFile = zlBlobRead(3, lngParentId)
                If Len(strZipFile) > 0 Then
                    If gobjFSO.FileExists(strZipFile) Then
                        strTemp = zlFileUnzip(strZipFile)
                        If gobjFSO.FileExists(strTemp) Then
                            Me.edtThis.OpenDoc strTemp
                            gobjFSO.DeleteFile strTemp, True
                        End If
                        gobjFSO.DeleteFile strZipFile, True
                    End If
                End If
            Else
                Call InsertContent(lngParentId)
            End If
        End If
    End If
    
    '表头填写
    vfgThis.MergeCells = flexMergeFixedOnly
    vfgThis.MergeCellsFixed = flexMergeFree
    For lngCount = 0 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    Me.vfgThis.AutoSize 0, Me.vfgThis.Cols - 1
    
    Me.edtThis.UnFreeze
    edtThis.RefreshTargetDC
    Me.edtThis.ReadOnly = True
    Exit Sub
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub SetCaption身份证()
    If Not mobjInfection Is Nothing Then Call mobjInfection.SetCaption身份证
End Sub

Private Sub InsertContent(ByVal lngFileID As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim rsText As New ADODB.Recordset, strTSql As String
    Dim Elements As New cEPRElements
    Dim Diagnosises As New cEPRDiagnosises
    Dim aryProp() As String, intWCount As Integer
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String
    
    With Me.edtThis
        .Freeze: .ForceEdit = True: .SelStart = 1
        intWCount = (.PaperWidth - .MarginLeft - .MarginRight) / Me.TextWidth("┈") - 1
    End With
    
    gstrSQL = "Select Id, 内容文本 From 病历范文内容 Where 文件id = [1] And 对象类型 = 1 Order By 对象序号"
    strTSql = "Select Id, 对象类型, 对象属性, 内容文本, 是否换行, 要素名称, 诊治要素id, 替换域, 要素类型, 要素长度, 要素小数, 要素单位," & vbNewLine & _
            "       要素表示, 要素值域, 输入形态" & vbNewLine & _
            "From 病历范文内容" & vbNewLine & _
            "Where 文件id = [1] And 父id + 0 = [2]" & vbNewLine & _
            "Order By 对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    Do While Not rsTemp.EOF
        lngStart = Me.edtThis.SelStart
        strTmp = StrConv("" & Trim(rsTemp!内容文本), vbWide)
        strTmp = vbCrLf & "<" & strTmp & ">" & String(intWCount - Len(strTmp) - 1, "┈") & vbCrLf
        lngLen = Len(strTmp)
        Me.edtThis.Range(lngStart, lngStart) = strTmp
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Protected = False
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Hidden = False
        Me.edtThis.Range(lngStart, lngStart + lngLen).Font.ForeColor = &HFFC0C0
        Me.edtThis.Range(lngStart + lngLen, lngStart + lngLen).Selected
        
        Set rsText = zlDatabase.OpenSQLRecord(strTSql, Me.Caption, lngFileID, CLng(rsTemp!ID))
        Do While Not rsText.EOF
            lngStart = Me.edtThis.SelStart
            Select Case rsText!对象类型
            Case 2, 3, 5 '文本,表格,图形
                Select Case rsText!对象类型
                Case 2: strTmp = "" & rsText!内容文本 & IIf(Val("" & rsText!是否换行) = 1, vbCrLf, "")
                Case 3: strTmp = vbCrLf & "□" & vbCrLf
                Case 5: strTmp = vbCrLf & "□" & vbCrLf
                End Select
                lngLen = Len(strTmp)
                Me.edtThis.Range(lngStart, lngStart) = strTmp
                Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Protected = False
                Me.edtThis.Range(lngStart, lngStart + lngLen).Font.Hidden = False
                Me.edtThis.Range(lngStart + lngLen, lngStart + lngLen).Selected
            Case 4  '要素
                lngKey = Elements.Add
                With Elements("K" & lngKey)
                    .内容文本 = "" & rsText!内容文本
                    .要素名称 = "" & rsText!要素名称
                    .诊治要素ID = Val("" & rsText!诊治要素ID)
                    .替换域 = Val("" & rsText!替换域)
                    .要素类型 = Val("" & rsText!要素类型)
                    .要素长度 = Val("" & rsText!要素长度)
                    .要素小数 = Val("" & rsText!要素小数)
                    .要素单位 = "" & rsText!要素单位
                    .要素表示 = Val("" & rsText!要素表示)
                    .要素值域 = "" & rsText!要素值域
                    .输入形态 = Val("" & rsText!输入形态)
                    .是否换行 = Val("" & rsText!是否换行)
                    .InsertIntoEditor Me.edtThis, lngStart, , True
                End With
            Case 7  '诊断
                lngKey = Diagnosises.Add
                With Diagnosises("K" & lngKey)
                    .描述 = "" & rsText!内容文本
                    aryProp = Split("" & rsText!对象属性, ";")
                    .类型 = Val(aryProp(0))
                    .中医 = Val(aryProp(1))
                    .疾病id = Val(aryProp(2))
                    .诊断id = Val(aryProp(3))
                    .证候id = Val(aryProp(4))
                    .疑诊 = Val(aryProp(5))
                    .日期 = Format(aryProp(6), "yyyy-mm-dd hh:mm:ss")
                    .InsertIntoEditor Me.edtThis, lngStart, True
                End With
            End Select
            rsText.MoveNext
        Loop
        rsTemp.MoveNext
    Loop
    With Me.edtThis
        .ForceEdit = False: .SelStart = 1: .Modified = False: .UnFreeze
    End With
    Set Elements = Nothing
    Set Diagnosises = Nothing
End Sub


Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回： 调用成功返回TRUE；否则FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    vsb.Value = 0: hsb.Value = 0
    picDraw.Top = 0: picDraw.Left = 0
    hsb.Max = picDraw.Width - picWave.Width
    vsb.Max = picDraw.Height - picWave.Height
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    If hsb.Visible Then hsb.ZOrder 0
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    If vsb.Visible Then vsb.ZOrder 0
    
    With vsb
        .Height = picWave.Height
    End With
    
    With hsb
        .Width = picWave.Width - IIf(vsb.Visible = True, vsb.Width, 0)
    End With
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (picDraw.Width - picWave.Width + IIf(vsb.Visible = True, vsb.Width, 0)) / 10
    msinVStep = (picDraw.Height - picWave.Height + IIf(hsb.Visible = True, hsb.Height, 0)) / 10
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then
        hsb.Max = 10
        hsb.LargeChange = 10 / Int((Round((picDraw.Width - picWave.Width + IIf(vsb.Visible = True, vsb.Width, 0)) / picWave.Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 10
        vsb.LargeChange = 10 / Int((Round((picDraw.Height - picWave.Height + IIf(hsb.Visible = True, hsb.Height, 0)) / picWave.Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
End Function

Private Sub picWave_Resize()
    With vsb
        .Left = picWave.Width - .Width
        .Top = 0
        .Height = picWave.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picWave.Height - .Height
        .Width = picWave.Width - vsb.Width
    End With
    
    Call CalcScrollBarSize
End Sub

Private Sub vsb_Change()
    picDraw.Top = -1 * vsb.Value * msinVStep
    VsfData.Top = (picDraw.Height - VsfData.Height) + -1 * vsb.Value * msinVStep
End Sub

Private Sub hsb_Change()
    picDraw.Left = -1 * hsb.Value * msinHStep
    VsfData.Left = -1 * hsb.Value * msinHStep
End Sub

Private Function ShowTabBaby(ByVal rsTmp As ADODB.Recordset, ByVal lngHeight As Long)
    Dim lngCurveRows As Long
    Dim lngMaxValue As Long, lngMinValue As Long
    Dim lngTotal As Long, lngCurveNull As Long
    Dim lngCurveRowHeight As Long
    Dim lngTabBabyRowHeight As Long
    Dim lngRow As Long, lngDay As Long
    Dim lngId  As Long, lngTabBabyTitleID As Long, lngTabBabyNameID As Long
    Dim strSQL  As String
    Dim strBabyTitle As String, strTitleBabyFont As String
    Dim intTitleBabyTitleNum As Integer, i As Integer
    Dim BlnBaby As Boolean
    Dim objFont As StdFont
    
    Dim rsCurve As New ADODB.Recordset
    
    rsTmp.Filter = "父ID=NULL And 对象序号=1 And 内容文本='格式定义'"
    If rsTmp.RecordCount > 0 Then
        lngId = rsTmp!ID
        rsTmp.Filter = "父ID=" & lngId
        Do While Not rsTmp.EOF
            Select Case "" & rsTmp!要素名称
            Case "天数"
                lngDay = Val("" & rsTmp!内容文本)
            Case "婴儿标题文本"
                strBabyTitle = "" & rsTmp!内容文本
            Case "婴儿标题字体"
                strTitleBabyFont = "" & rsTmp!内容文本
            Case "婴儿表格高度"
                lngTabBabyRowHeight = Val("" & rsTmp!内容文本)
            Case "表头层数"
                intTitleBabyTitleNum = Val("" & rsTmp!内容文本)
            Case "婴儿体温单"
                BlnBaby = Val("" & rsTmp!内容文本)
            Case "总列数"
                VsfData.Cols = Val("" & rsTmp!内容文本)
            End Select
            rsTmp.MoveNext
        Loop
    End If
    If Not BlnBaby Then VsfData.Visible = False: Exit Function
    
    rsTmp.Filter = "父ID=NULL And 对象序号=4 And 内容文本='婴儿体温单表头项目'"
    Do While Not rsTmp.EOF
        lngTabBabyTitleID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    rsTmp.Filter = "父ID=NULL And 对象序号=3 And 内容文本='表格项目定义'"
    Do While Not rsTmp.EOF
        lngTabBabyNameID = Val("" & rsTmp!ID)
        rsTmp.MoveNext
    Loop
    
    
    With VsfData
        .Top = Me.ScaleX(lngHeight, vbPixels, vbTwips) + 200
        .Left = 0
        .Rows = .FixedRows + lngDay + 1
        .Width = picDraw.Width
        .Height = lngTabBabyRowHeight * (VsfData.Rows + 2)
        
        Select Case intTitleBabyTitleNum
            Case 1
                .RowHidden(2) = True
                .RowHidden(3) = True
            Case 2
                .RowHidden(3) = True
        End Select
        
        rsTmp.Filter = "父ID= " & lngTabBabyNameID
        rsTmp.Sort = "对象序号"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .ColWidth(Val(rsTmp!对象序号) - 1) = Split(rsTmp!对象属性, "`")(0)
                rsTmp.MoveNext
            Loop
        End If
        rsTmp.Filter = "父ID= " & lngTabBabyTitleID
        rsTmp.Sort = "对象序号"
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .TextMatrix((Val(rsTmp!内容行次)), Val(rsTmp!对象序号) - 1) = NVL(rsTmp!内容文本)
                rsTmp.MoveNext
            Loop
        End If
        .Cell(flexcpText, 0, 0, 0, .Cols - 1) = strBabyTitle
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        .CellBorderRange 1, 0, .Rows - 1, .Cols - 1, vbBlack, 1, 1, 1, 1, 1, 1
        .MergeCellsFixed = flexMergeFree
        .MergeCol(-1) = True
        .MergeRow(-1) = True
        
        Set objFont = New StdFont
        With objFont
            .Name = Split(strTitleBabyFont, ",")(0)
            .Size = Val(Split(strTitleBabyFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strTitleBabyFont, "粗") > 0 Then .Bold = True
            If InStr(1, strTitleBabyFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 0, .FixedCols, 0, .Cols - 1) = objFont
        .ROWHEIGHT(0) = objFont.Size * 20 + 150
        For i = 4 To .Rows - 1
        .ROWHEIGHT(i) = lngTabBabyRowHeight
        VsfData.Redraw = True
        Next
        
    End With
    picDraw.Height = picDraw.Height + VsfData.Height
    VsfData.Visible = True
    
End Function


