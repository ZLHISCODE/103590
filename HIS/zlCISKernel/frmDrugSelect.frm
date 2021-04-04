VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDrugSelect 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "药品选择器"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   Icon            =   "frmDrugSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   11670
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   10292
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox picVsf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   3600
      ScaleHeight     =   5895
      ScaleWidth      =   8055
      TabIndex        =   1
      Top             =   120
      Width           =   8055
      Begin VSFlex8Ctl.VSFlexGrid vsItem 
         Height          =   5835
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7980
         _cx             =   14076
         _cy             =   10292
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDrugSelect.frx":058A
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
         ExplorerBar     =   3
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   6120
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":0617
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":0BB1
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":114B
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":16E5
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":1C7F
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugSelect.frx":2219
            Key             =   "方案"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdEsc 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   180
      Left            =   0
      TabIndex        =   3
      Top             =   -900
      Width           =   90
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   10000
      Y1              =   5460
      Y2              =   5460
   End
End
Attribute VB_Name = "frmDrugSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr药品类型 As String, mlng项目id   As Long
Private mbytOK As Byte
Private mrsItem As ADODB.Recordset
Private mstr分类Tmp As String

Public Function ShowSelect(frmParent As Object, bytOK As Byte, Optional ByVal str药品类型 As String, Optional ByVal lng项目id As Long) As ADODB.Recordset
'功能：显示药品选择器
'参数：str药品类型=用于定位分类
'      lng项目id=用于定位项目


    mstr药品类型 = str药品类型
    mlng项目id = lng项目id
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    bytOK = mbytOK
    Set ShowSelect = IIF(bytOK = 1, mrsItem, Nothing)
End Function





Private Sub cmdEsc_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    InitDockPannel '初始化拖动控件
    Call FillTree
    
    mstr分类Tmp = ""
    mbytOK = 0
End Sub



'设置布局控件
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False '实时拖动
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    Exit Function
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Function



'InitDockPannel初始区域划分
Private Sub InitDockPannel()
    Dim objPane As Pane
On Error GoTo errH
    Set objPane = dkpMain.CreatePane(1, 200, 500, DockLeftOf, objPane)
    objPane.Title = "药品分类目录"
    objPane.Options = PaneNoCaption
    Set objPane = dkpMain.CreatePane(2, 550, 500, DockRightOf, objPane)
    objPane.Title = "详情"
    objPane.Options = PaneNoCaption

    Call DockPannelInit(dkpMain)
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub


Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
On Error GoTo errH
    Select Case Item.ID
        Case 1
            Item.Handle = tvw_s.hwnd
        Case 2
            Item.Handle = picVsf.hwnd

    End Select
    Exit Sub
errH:
    MsgBox err.Description, vbInformation, gstrSysName
    err.Clear
End Sub




Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node
    
    On Error GoTo errH

    strSQL = _
        " Select 0 as 级,类型,-类型 as ID,-Null as 上级ID,类型||'' as 编码," & _
        " 类型||'.'||Decode(类型,1,'西成药',2,'中成药',3,'中草药',4,'中药配方') as 名称" & _
        " From 诊疗分类目录 Where 类型 in (1,2,3,4) And 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Group by 类型"
    strSQL = strSQL & " Union ALL " & _
        " Select Level as 级,类型,ID,Nvl(上级ID,-类型) as 上级ID,编码,名称 From 诊疗分类目录" & _
        " Where 类型 in (1,2,3,4) And 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Order by 级,编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
        
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, "[" & rsTmp!编码 & "]" & rsTmp!名称, "Close")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub Form_Resize()
    On Error Resume Next
    Call PicVsf_Resize
End Sub

Private Sub PicVsf_Resize()
    vsItem.Top = 10: vsItem.Left = 10: vsItem.Width = picVsf.Width - 10: vsItem.Height = picVsf.Height - 10
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstr分类Tmp Then Exit Sub
    Call FillList
    mstr分类Tmp = Node.Key
End Sub


Private Function FillList() As Boolean
    Dim strSub As String, strSQL As String
    Dim objNode As Node, str类别 As String
    Dim i As Long
    
    On Error GoTo errH
    Set objNode = tvw_s.SelectedItem '可能为Nothing
    If objNode Is Nothing Then Exit Function
    If Not mrsItem Is Nothing Then mrsItem.Filter = ""
    
    '显示下级的项目
    If Val(Mid(objNode.Key, 2)) < 0 Then
        strSub = " And A.分类ID IN(" & _
            " Select ID From 诊疗分类目录 Where 类型=[1] And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " )"
    Else
        strSub = " And A.分类ID IN(" & _
            " Select ID From 诊疗分类目录 Where 撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD')" & _
            " Start With ID=[3] Connect by Prior ID=上级ID)"
    End If
    
    '树形中的类型确定类别
    If Val(objNode.Tag) > 0 Then str类别 = Choose(Val(objNode.Tag), "5", "6", "7", "8", "", "9", "4")
    If str类别 <> "" Then strSub = strSub & " And A.类别=[2]"
    
    strSQL = "Select a.Id As 诊疗项目id, b.Id As 收费细目id,decode(a.类别,'5','西成药','6','中成药','7','中草药','8','配方') as 类别, a.名称, b.规格, a.计算单位, d.药品剂型,C.住院单位 as 总量单位" & _
                " From 诊疗项目目录 A, 收费项目目录 B, 药品规格 C, 药品特性 D" & _
                " Where c.药品id= b.Id(+)   And a.Id =c.药名id (+) And c.药名id = d.药名id(+) and (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null)" & strSub
    
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(objNode.Tag), str类别, Val(Mid(objNode.Key, 2)))
    
    '绑定数据
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    
    '可能统计常用项目时设置为了0行0列
    If vsItem.FixedRows = 0 Then
        vsItem.Rows = 2
        vsItem.FixedRows = 1
    End If
    If vsItem.FixedCols = 0 Then
        vsItem.Cols = 2
        vsItem.FixedCols = 1
    End If
    
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '列属性调整
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColAlignment(i) = 1
        
        If vsItem.TextMatrix(0, i) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '记录原始列号,用于处理列顺序
    Next
    vsItem.Redraw = flexRDBuffered
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub vsItem_DblClick()
    If vsItem.Row = -1 Then Exit Sub
    If vsItem.Row >= vsItem.FixedRows Then
        If mrsItem.RecordCount = 1 Then
            mbytOK = 1
        Else
            mbytOK = 0
        End If

        Unload Me
    End If
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        mrsItem.Filter = IIF(vsItem.TextMatrix(NewRow, GetCol("类别")) = "配方", "诊疗项目id =" & Val(vsItem.TextMatrix(NewRow, GetCol("诊疗项目id"))), "诊疗项目id =" & Val(vsItem.TextMatrix(NewRow, GetCol("诊疗项目id"))) & " And 收费细目id =" & Val(vsItem.TextMatrix(NewRow, GetCol("收费细目id"))))
    End If
End Sub


Private Sub vsItem_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsItem.ColDataType(Col) = flexDTBoolean Then Cancel = True
End Sub

Private Function GetCol(ByVal strName As String) As Long
    Dim i As Long
    For i = 1 To vsItem.Cols - 1
        If UCase(vsItem.TextMatrix(0, i)) = UCase(strName) Then
            GetCol = i: Exit Function
        End If
    Next
End Function



