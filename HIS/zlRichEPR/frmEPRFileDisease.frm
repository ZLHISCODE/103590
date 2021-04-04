VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEPRFileDisease 
   Caption         =   "疾病证明报告前提"
   ClientHeight    =   7995
   ClientLeft      =   2835
   ClientTop       =   3825
   ClientWidth     =   11880
   Icon            =   "frmEPRFileDisease.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1845
      Index           =   4
      Left            =   1050
      ScaleHeight     =   1845
      ScaleWidth      =   3690
      TabIndex        =   12
      Top             =   4950
      Width           =   3690
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1425
         Left            =   390
         TabIndex        =   15
         Top             =   270
         Width           =   1950
         _cx             =   3440
         _cy             =   2514
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEPRFileDisease.frx":058A
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3420
      Index           =   3
      Left            =   285
      ScaleHeight     =   3420
      ScaleWidth      =   6525
      TabIndex        =   11
      Top             =   645
      Width           =   6525
      Begin VB.Frame fra 
         Height          =   750
         Left            =   0
         TabIndex        =   16
         Top             =   -90
         Width           =   5790
         Begin VB.Image imgNote 
            Height          =   480
            Left            =   60
            Picture         =   "frmEPRFileDisease.frx":05EC
            Top             =   150
            Width           =   480
         End
         Begin VB.Label lblMeasure 
            AutoSize        =   -1  'True
            Caption         =   "当发生以下诊断时，应按规定填写本文件，并向相关部门报告。"
            Height          =   180
            Left            =   630
            TabIndex        =   17
            Top             =   315
            Width           =   5040
         End
      End
      Begin XtremeSuiteControls.TabControl tbcPage 
         Height          =   1800
         Left            =   150
         TabIndex        =   13
         Top             =   735
         Width           =   2790
         _Version        =   589884
         _ExtentX        =   4921
         _ExtentY        =   3175
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3000
      Index           =   2
      Left            =   8340
      ScaleHeight     =   3000
      ScaleWidth      =   4230
      TabIndex        =   3
      Top             =   3420
      Width           =   4230
      Begin VB.ComboBox cboReport 
         Height          =   300
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   18
         ToolTipText     =   "退格或删除键清空"
         Top             =   2505
         Width           =   1155
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "添加(&A)"
         Height          =   350
         Index           =   0
         Left            =   2640
         TabIndex        =   9
         Top             =   2475
         Width           =   1200
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   60
         TabIndex        =   7
         Top             =   2025
         Width           =   2640
      End
      Begin VB.CommandButton cmdFind 
         Height          =   300
         Left            =   2685
         Picture         =   "frmEPRFileDisease.frx":0EB6
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   2010
         Width           =   360
      End
      Begin VB.CommandButton cmdSel 
         Height          =   300
         Index           =   0
         Left            =   3120
         Picture         =   "frmEPRFileDisease.frx":1440
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "选中所有的项目"
         Top             =   2010
         Width           =   360
      End
      Begin VB.CommandButton cmdSel 
         Height          =   300
         Index           =   1
         Left            =   3480
         Picture         =   "frmEPRFileDisease.frx":19CA
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "清除所有选择"
         Top             =   2010
         Width           =   360
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1845
         Left            =   30
         TabIndex        =   8
         Top             =   15
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3254
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblReport 
         Caption         =   "报告病种"
         Height          =   240
         Left            =   60
         TabIndex        =   19
         Top             =   2550
         Width           =   780
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   1470
      Index           =   0
      Left            =   420
      ScaleHeight     =   1470
      ScaleWidth      =   3030
      TabIndex        =   2
      Top             =   4290
      Width           =   3030
      Begin VSFlex8Ctl.VSFlexGrid vgdItems 
         Height          =   1170
         Left            =   45
         TabIndex        =   10
         Top             =   120
         Width           =   2475
         _cx             =   4366
         _cy             =   2064
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEPRFileDisease.frx":1F54
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
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   2625
      Index           =   1
      Left            =   8130
      ScaleHeight     =   2625
      ScaleWidth      =   4410
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   555
      Width           =   4410
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   2040
         Left            =   345
         TabIndex        =   1
         Tag             =   "1000"
         Top             =   255
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   3598
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imgList"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6135
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileDisease.frx":1FB6
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEPRFileDisease.frx":2550
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   7635
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRFileDisease.frx":2AEA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17066
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmEPRFileDisease.frx":337E
      Left            =   4155
      Top             =   -45
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRFileDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
Private Const conColumn_ID = 0
Private Const conColumn_编码 = 1
Private Const conColumn_名称 = 2
Private Const conColumn_报告病种 = 3

Private mlngFileID As Long        '病历文件ID
Private mblnOk As Boolean
Private mblnDeleteAsk As Boolean
Private mblnMustReport As Boolean
Private lngCount As Long

'######################################################################################################################

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long) As Boolean
    '******************************************************************************************************************
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '******************************************************************************************************************
    mlngFileID = lngFileID
    
    If ExecuteCommand("初始控件") = False Then GoTo EndPoint
    If InitDate = False Then GoTo EndPoint
    
    Call ExecuteCommand("刷新数据")
    
    
    Me.Show vbModal, frmParent
    
    ShowMe = mblnOk
    
    Exit Function

EndPoint:
    Unload Me
End Function
Private Function InitDate() As Boolean
Dim rs As ADODB.Recordset
    On Error GoTo errHand
    mblnDeleteAsk = True
    
    gstrSQL = "Select 种类, 编号, 名称,保留, 通用 From 病历文件列表 Where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rs.BOF Then
         MsgBox "文件丢失(可能被其他用户删除)！", vbInformation, gstrSysName
         Exit Function
    Else
        Me.Caption = "[" & rs!编号 & "-" & rs!名称 & "]的疾病证明报告前提"
        mblnMustReport = NVL(rs!保留, 1) = 4
    End If
    
    gstrSQL = "select 编码,简码,名称 from 传染病目录 order by 编码"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rs Is Nothing Then
        If rs.RecordCount > 0 Then
            Do While Not rs.EOF
                cboReport.AddItem NVL(rs!简码) & "-" & NVL(rs!名称)
                rs.MoveNext
            Loop
        End If
    End If

    zlControl.CboSetWidth cboReport.hWnd, 6000
    InitDate = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
Private Sub InitDockPannel()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    Dim objPane As Pane

    Set objPane = dkpMain.CreatePane(1, 100, 100, DockLeftOf, Nothing)
    objPane.Title = "内容"
    objPane.Options = PaneNoCaption
    
    Set objPane = dkpMain.CreatePane(2, 300, 100, DockRightOf, Nothing)
    objPane.Title = "分类"
    objPane.Options = PaneNoCaption

    Set objPane = dkpMain.CreatePane(3, 300, 100, DockBottomOf, objPane)
    objPane.Title = "明细"
    objPane.Options = PaneNoCaption
    
    dkpMain.SetCommandBars cbsMain
    
    Call DockPannelInit(dkpMain)

End Sub

Private Function InitTabControl() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    With tbcPage
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameSingleLine
            .ShowIcons = True
            .DisableLunaColors = False
        End With

        .InsertItem 0, "疾病编码(ICD-10)", picPane(0).hWnd, 0
        .InsertItem 1, "诊断编码", picPane(4).hWnd, 0
        
        .Item(0).Selected = True
        
    End With
    
    InitTabControl = True
    
End Function

Private Function ExecuteCommand(ByVal strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
Dim intLoop As Integer
Dim rs As New ADODB.Recordset
Dim rsSQL As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim strTmp As String
Dim lngTMP As Long
Dim objItem As ListItem
Dim objNode As Node
    
    On Error GoTo errHand
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        Call InitCommandBar
        Call InitDockPannel

        lvwItems.ListItems.Clear
        With lvwItems.ColumnHeaders
            .Clear
            .Add , "_编码", "编码", 1000
            .Add , "_名称", "名称", 2300
            .Add , "_简码", "简码", 600
        End With
        With lvwItems
            .SortKey = .ColumnHeaders("_编码").Index - 1
            .SortOrder = lvwAscending
        End With
    
        Call InitTabControl
            
        
    '--------------------------------------------------------------------------------------------------------------
    Case "刷新数据"

        Call ExecuteCommand("读取疾病前提")
        Call ExecuteCommand("读取诊断前提")
            
    '--------------------------------------------------------------------------------------------------------------
    Case "读取疾病前提"

        gstrSQL = "Select Id, 编码, 名称, p.报告病种 From 疾病编码目录 i, 疾病报告前提 p Where i.Id = p.疾病id And p.文件id = [1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set vgdItems.DataSource = rs
        With vgdItems
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            .ColHidden(conColumn_ID) = True
            .ColWidth(conColumn_编码) = 1000
            .ColWidth(conColumn_名称) = 3650
            .ColWidth(conColumn_报告病种) = 1800
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case "读取诊断前提"
    
        gstrSQL = "Select Id, 编码, 名称, p.报告病种 From 疾病诊断目录 i, 疾病报告前提 p Where i.Id = p.诊断id And p.文件id = [1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        Set vsf.DataSource = rs
        With vsf
            For lngCount = 0 To .Cols - 1
                .FixedAlignment(lngCount) = flexAlignCenterCenter
                .ColAlignment(lngCount) = flexAlignLeftCenter
            Next
            .ColHidden(conColumn_ID) = True
            .ColWidth(conColumn_编码) = 1000
            .ColWidth(conColumn_名称) = 3650
            .ColWidth(conColumn_报告病种) = 1800
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case "读取疾病分类"
        
        gstrSQL = "Select Id, 上级id, 序号, 名称 From 疾病编码分类 Where 类别 = 'D'  And (撤档时间 is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))  Start With 上级id Is Null Connect By Prior Id = 上级id Order By Level, 序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        On Error GoTo 0
        With rsTemp
            Me.tvwClass.Nodes.Clear
            Do While Not .EOF
                If IsNull(!上级ID) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !名称, "close")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, !名称, "close")
                End If
                objNode.ExpandedImage = "expend"
                .MoveNext
            Loop
            If tvwClass.Nodes.Count > 0 Then
                tvwClass.Nodes(1).Expanded = True
                tvwClass.Nodes(1).Selected = True
                Call tvwClass_NodeClick(tvwClass.Nodes(1))
            End If
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "读取疾病明细"
        
        gstrSQL = "Select i.Id, i.编码, i.名称, i.简码" & _
                " From 疾病编码目录 i, (Select 疾病id From 疾病报告前提 Where 文件id = [2] And 疾病id Is Not Null) s" & _
                " Where i.类别 = 'D' And i.分类id = [1] And i.Id = s.疾病id(+) And s.疾病id Is Null And Nvl(i.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Mid(tvwClass.SelectedItem.Key, 2)), mlngFileID)
        With rs
        
            lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = lvwItems.ListItems.Add(, "_" & rs!ID, rs!编码)
                objItem.SubItems(lvwItems.ColumnHeaders("_名称").Index - 1) = rs!名称
                objItem.SubItems(lvwItems.ColumnHeaders("_简码").Index - 1) = rs!简码
                .MoveNext
            Loop
        End With
        Me.lvwItems.Tag = "0"
        
    '--------------------------------------------------------------------------------------------------------------
    Case "读取诊断分类"
                
        gstrSQL = "Select Id, 上级id, 编码, 名称 From 疾病诊断分类 Start With 上级id Is Null Connect By Prior Id = 上级id"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        On Error GoTo 0
        With rs
            Me.tvwClass.Nodes.Clear
            Do While Not .EOF
                If IsNull(!上级ID) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, !名称, "close")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, !名称, "close")
                End If
                objNode.ExpandedImage = "expend"
                .MoveNext
            Loop
            If tvwClass.Nodes.Count > 0 Then
                tvwClass.Nodes(1).Expanded = True
                tvwClass.Nodes(1).Selected = True
                Call tvwClass_NodeClick(tvwClass.Nodes(1))
            End If
        End With
    '--------------------------------------------------------------------------------------------------------------
    Case "读取诊断明细"
                
        gstrSQL = "Select i.Id, i.编码, i.名称, zlSpellcode(i.名称) As 简码" & _
                " From 疾病诊断目录 i,疾病诊断属类 j,(Select 诊断id From 疾病报告前提 Where 文件id = [2] And 诊断id Is Not Null) s" & _
                " Where i.id = j.诊断id And j.分类id = [1] And i.Id = s.诊断id(+) And s.诊断id Is Null And Nvl(i.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(Mid(tvwClass.SelectedItem.Key, 2)), mlngFileID)
        With rs
            lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = lvwItems.ListItems.Add(, "_" & rs!ID, rs!编码)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_名称").Index - 1) = rs!名称
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_简码").Index - 1) = rs!简码
                .MoveNext
            Loop
        End With
        Me.lvwItems.Tag = "0"
    
    '--------------------------------------------------------------------------------------------------------------
    Case "删除疾病前提"
        
        With vgdItems
            If Val(.TextMatrix(.Row, conColumn_ID)) = 0 Then MsgBox "已经删除完成！", vbInformation, gstrSysName: Exit Function
            If mblnDeleteAsk Then
                If MsgBox("真的删除该疾病诊断吗？" & vbCrLf & "——" & .TextMatrix(.Row, conColumn_名称), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
            gstrSQL = "Zl_疾病报告前提_Delete(" & mlngFileID & "," & Val(.TextMatrix(.Row, conColumn_ID)) & ",Null)"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            mblnOk = True
            .RemoveItem .Row
        End With
        
    '--------------------------------------------------------------------------------------------------------------
    Case "删除诊断前提"
    
        With vsf
            If Val(.TextMatrix(.Row, conColumn_ID)) = 0 Then MsgBox "已经删除完成！", vbInformation, gstrSysName: Exit Function
            If mblnDeleteAsk Then
                If MsgBox("真的删除该疾病诊断吗？" & vbCrLf & "——" & .TextMatrix(.Row, conColumn_名称), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
            gstrSQL = "Zl_疾病报告前提_Delete(" & mlngFileID & ",Null," & Val(.TextMatrix(.Row, conColumn_ID)) & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            mblnOk = True
            .RemoveItem .Row
        End With
        
    End Select


    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    Call CommandBarInit(cbsMain)

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.Options.LargeIcons = True
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份

    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Edit_Delete, "删除(&D)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_View_Option, "删除提醒(&M)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_Help_Help, "帮助(&H)", True)
    Set objControl = NewToolBar(objBar, xtpControlButton, conMenu_File_Exit, "退出(&X)")
    
    '------------------------------------------------------------------------------------------------------------------
    '命令的快键绑定:公共部份主界面已处理

    With cbsMain.KeyBindings
        .Add 0, vbKeyF1, conMenu_Help_Help                  '帮助
    End With

End Function
Private Sub cboReport_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        cboReport.ListIndex = -1
    End If
End Sub

Private Sub cboReport_KeyPress(KeyAscii As Integer)
    Call zlControl.CboSetIndex(cboReport.hWnd, zlControl.CboMatchIndex(cboReport.hWnd, KeyAscii))
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)

    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        mblnDeleteAsk = Not mblnDeleteAsk
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete
        
        If tbcPage.Selected.Index = 0 Then
            Call ExecuteCommand("删除疾病前提")
        Else
            Call ExecuteCommand("删除诊断前提")
        End If

    '--------------------------------------------------------------------------------------------------------------
    Case Else

         '与业务无关的功能，公共的功能
        Call CommandBarExecutePublic(Control, Me)

    End Select
End Sub


Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Select Case Control.ID
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Option
        Control.Checked = mblnDeleteAsk
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_Edit_Delete

        If tbcPage.Selected.Index = 0 Then
            Control.Enabled = (Val(vgdItems.TextMatrix(vgdItems.Row, conColumn_ID)) > 0)
        Else
            Control.Enabled = (Val(vsf.TextMatrix(vsf.Row, conColumn_ID)) > 0)
        End If
        
    '--------------------------------------------------------------------------------------------------------------
    Case Else

        '与业务无关的功能，公共的功能
        Call CommandBarUpdatePublic(Control, Me)

    End Select
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub cmdEdit_Click(Index As Integer)
Dim strTemp As String, strReport As String
Dim objItem As ListItem
    
    
    If cboReport.Text = "" Then
        If mblnMustReport Then '必须录入报告病种
            MsgBox "必须选择报告病种，请检查！", vbInformation, gstrSysName
            If cboReport.Enabled Then
                cboReport.SetFocus
                SendKeys "{F4}"
            End If
            Exit Sub
        End If
    Else
        strReport = Split(cboReport.Text, "-")(1)
    End If
    
    If tbcPage.Selected.Index = 0 Then        '添加
    
        strTemp = ""
        
        For Each objItem In lvwItems.ListItems
            If objItem.Checked Then strTemp = strTemp & ";" & Mid(objItem.Key, 2)
        Next
        
        If strTemp = "" Then MsgBox "没有选择疾病诊断项目！", vbInformation, gstrSysName: Exit Sub
        
        If Len(strTemp) > 4000 Then MsgBox "一次选择了太多的疾病诊断项目！", vbInformation, gstrSysName: Exit Sub
        
        gstrSQL = "Zl_疾病报告前提_Append(" & mlngFileID & ",'" & Mid(strTemp, 2) & "',Null,'" & strReport & "')"
        
        Err = 0
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call ExecuteCommand("读取疾病前提")
    Else
    
        strTemp = ""
        
        For Each objItem In lvwItems.ListItems
            If objItem.Checked Then strTemp = strTemp & ";" & Mid(objItem.Key, 2)
        Next
        
        If strTemp = "" Then MsgBox "没有选择疾病诊断项目！", vbInformation, gstrSysName: Exit Sub
        If Len(strTemp) > 4000 Then MsgBox "一次选择了太多的疾病诊断项目！", vbInformation, gstrSysName: Exit Sub
        
        gstrSQL = "Zl_疾病报告前提_Append(" & mlngFileID & ",Null,'" & Mid(strTemp, 2) & "','" & strReport & "')"
        
        Err = 0
        On Error GoTo errHand
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Call ExecuteCommand("读取诊断前提")
        
    End If
    
    mblnOk = True: Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub cmdFind_Click()
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
    If Trim(Me.txtFind.Text) = "" Then MsgBox "没有输入查找内容！", vbInformation, gstrSysName: Exit Sub
    
    If tbcPage.Selected.Index = 0 Then
        gstrSQL = "Select i.Id, i.编码, i.名称, i.简码" & _
                " From 疾病编码目录 i, (Select 疾病id From 疾病报告前提 Where 文件id = [3] And 疾病id Is Not Null) s" & _
                " Where i.类别 = 'D' And (i.编码 like [1] or i.名称 like [2] or i.简码 like [2]) And i.Id = s.疾病id(+) And s.疾病id Is Null And Nvl(i.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
    Else
    
        gstrSQL = "Select i.Id, i.编码, i.名称, zlSpellCode(i.名称) As 简码" & _
                " From 疾病诊断目录 i, (Select 诊断id From 疾病报告前提 Where 文件id = [3] And 诊断id Is Not Null) s" & _
                " Where (i.编码 like [1] or i.名称 like [2] or zlSpellCode(i.名称) like [2]) And i.Id = s.诊断id(+) And s.诊断id Is Null And Nvl(i.撤档时间, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'YYYY-MM-DD')"
    End If
    
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(Me.txtFind.Text) & "%", gstrMatch & Trim(Me.txtFind.Text) & "%", mlngFileID)
    With rsTemp
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_名称").Index - 1) = !名称
            objItem.SubItems(Me.lvwItems.ColumnHeaders("_简码").Index - 1) = !简码
            .MoveNext
        Loop
    End With
    If lvwItems.ListItems.Count = 0 Then
        If Val(cmdFind.Tag) = 0 Then MsgBox "没有匹配的疾病诊断项目！", vbInformation, gstrSysName
        txtFind.SetFocus
    Else
        If tbcPage.Selected.Index = 0 Then
            vgdItems.SetFocus
        Else
            vsf.SetFocus
        End If
    End If
    lvwItems.Tag = "1"
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click(Index As Integer)
Dim objItem As ListItem
    For Each objItem In Me.lvwItems.ListItems
        objItem.Checked = (Index = 0)
    Next
    Me.lvwItems.SetFocus
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(3).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    Case 3
        Item.Handle = picPane(2).hWnd
    End Select
End Sub

Private Sub Form_Activate()
    Me.vgdItems.SetFocus
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Call SetPaneRange(dkpMain, 2, 100, 15, 300, 300)
    
    dkpMain.RecalcLayout
    
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwItems
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub


Private Sub picPane_Resize(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    Case 0
        vgdItems.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 1
        tvwClass.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    Case 2
        lvwItems.Move 0, 0, picPane(Index).Width, picPane(Index).Height - txtFind.Height - cmdEdit(0).Height - 90
        txtFind.Move 0, lvwItems.Top + lvwItems.Height + 45, picPane(Index).Width - cmdFind.Width - cmdSel(0).Width - cmdSel(1).Width - 45
        cmdFind.Move txtFind.Left + txtFind.Width + 15, txtFind.Top
        cmdSel(0).Move cmdFind.Left + cmdFind.Width + 15, cmdFind.Top
        cmdSel(1).Move cmdSel(0).Left + cmdSel(0).Width + 15, txtFind.Top
        
        lblReport.Move 0, cmdFind.Top + cmdFind.Height + 90
        cboReport.Move lblReport.Left + lblReport.Width + 15, cmdFind.Top + cmdFind.Height + 60, txtFind.Width - lblReport.Width - 15
        cmdEdit(0).Move picPane(Index).Width - cmdEdit(0).Width - 15, cmdFind.Top + cmdFind.Height + 30, picPane(Index).Width - lblReport.Width - cboReport.Width - 45
    Case 3
        
        fra.Move 0, -75, picPane(Index).Width
        tbcPage.Move 0, fra.Top + fra.Height - 75, picPane(Index).Width, picPane(Index).Height - tbcPage.Top
        
    Case 4
        vsf.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub

Private Sub tbcPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    
    If Item.Index = 0 Then
        Call ExecuteCommand("读取疾病分类")
    Else
        Call ExecuteCommand("读取诊断分类")
    End If
    
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If tbcPage.Selected.Index = 0 Then
        Call ExecuteCommand("读取疾病明细")
    Else
        Call ExecuteCommand("读取诊断明细")
    End If
    
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click: Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub vgdItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then Call cmdEdit_Click(1)
End Sub


