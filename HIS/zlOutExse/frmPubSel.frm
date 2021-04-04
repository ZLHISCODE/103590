VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPubSel 
   AutoRedraw      =   -1  'True
   Caption         =   "选择器"
   ClientHeight    =   5784
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8676
   Icon            =   "frmPubSel.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8676
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList img 
      Left            =   120
      Top             =   1800
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":058A
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":6DEC
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":D64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPubSel.frx":13EB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   4560
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3045
      _ExtentX        =   5376
      _ExtentY        =   8043
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3210
      Left            =   2145
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3216
      ScaleWidth      =   48
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   540
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   4560
      Left            =   3120
      TabIndex        =   7
      Top             =   600
      Width           =   5445
      _cx             =   9604
      _cy             =   8043
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.8
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
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPubSel.frx":1A712
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
   Begin VB.PictureBox picInfo 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      ScaleHeight     =   456
      ScaleWidth      =   8676
      TabIndex        =   5
      Top             =   0
      Width           =   8670
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请选择一个项目,然后点击确定"
         Height          =   180
         Left            =   180
         TabIndex        =   6
         Top             =   157
         Width           =   2430
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   516
      ScaleWidth      =   8676
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5268
      Width           =   8670
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   7335
         TabIndex        =   2
         Top             =   105
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6210
         TabIndex        =   1
         Top             =   105
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmPubSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mfrmParent As Object
Private mstrKey As String
Private mstrPrivs As String
Private mint险类 As Integer
Private mint病人来源 As Integer
Private mstr输入 As String
Private mint中药形态 As Integer
Private mrsItem As ADODB.Recordset
Private mblnOK As Boolean
Private mstrLike As String
Private mstrSaveTag As String
Private mlng病人ID As Long
Private mlng药房id As Long
Private mlng开单科室ID As Long
Private mlngX As Long
Private mlngY As Long
Private mlngH As Long

Public Function ShowSelect(ByVal frmParent As Object, ByVal strPrivs As String, _
                                           ByVal int病人来源 As Integer, ByVal int险类 As Integer, _
                                           ByVal lng病人ID As Long, ByVal lng药房ID As Long, _
                                           ByVal lng开单科室ID As Long, ByRef blnCancel As Boolean, _
                                           Optional ByVal str输入 As String, _
                                           Optional ByVal int中药形态 As Integer = -1, _
                                           Optional ByVal lngX As Long, _
                                           Optional ByVal lngY As Long, _
                                           Optional ByVal lngH As Long) As ADODB.Recordset
'功能：多功能选择器
'参数：
'入参:int病人来源=指病人来源,1-门诊,2-住院
    '     bln药房单位=是否按药房单位显示库存和价格
    '     str输入=输入匹配的内容,如果没有则为选择器方式,否则为列表方式
    '     int中药形态:-1表示不区分中药形态,0-只显示散装形态的中药,1-只显示饮片形态的中药;2-只显示免煎形态的中药
'出参:blnCancel-是否为用户取消操作
'返回：取消=Nothing,选择=SQL源的单行记录集

    Set mfrmParent = frmParent
    
    mstrPrivs = strPrivs: mstr输入 = str输入
    mlng药房id = lng药房ID: mlng病人ID = lng病人ID
    mint病人来源 = int病人来源: mint险类 = int险类:
    mlngX = lngX: mlngY = lngY: mlngH = lngH
    mlng开单科室ID = lng开单科室ID: mint中药形态 = int中药形态

    mstrSaveTag = IIf(mstr输入 <> "", 1, 0)

    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        If mrsItem Is Nothing Then Exit Function
        Set ShowSelect = mrsItem
        If ShowSelect.RecordCount = 0 Then Set ShowSelect = Nothing
    End If
    blnCancel = Not mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub cmdCancel_Click()
    Set mrsItem = Nothing '取消标志
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long
    If mrsItem Is Nothing Then Exit Sub
    If mrsItem.RecordCount < 1 Then Exit Sub
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    If mrsItem.RecordCount > 1 Then
        With vsItem
            If Val(.TextMatrix(.Row, .ColIndex("id"))) > 0 Then lngID = Val(.TextMatrix(.Row, .ColIndex("id")))
        End With
        mrsItem.Filter = "id=" & lngID
    End If
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsItem.Visible Then
        If vsItem.Row = 0 And tvw_s.Visible = True Then
            tvw_s.SetFocus
        Else
            vsItem.SetFocus
        End If
    Else
        tvw_s.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    If KeyCode = 13 And cmdOK.Enabled Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape And cmdCancel.Enabled Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long
    Dim vRect As RECT, strIDs As String, i As Long
    Dim lngUpH As Long, lngDnH As Long

    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    
    mblnOK = False
    mstrLike = gstrLike
    mstrKey = ""
    If mstr输入 = "" Then
        '读取类别失败,已提示,非取消退出
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '无类别,提示,非取消退出
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "没有设置相关收费项目类别,请先到收费项目管理中设置。", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    Else
        tvw_s.Visible = False
        pic.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False

        '填充匹配数据
        Call FillList(strIDs)
        If mrsItem Is Nothing Then
            Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount = 1 Then
            '只有一个项目时,直接返回
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount > 0 Then
            '多行是同一个项目时,直接返回
            If mstr输入 <> "" Then
                If UBound(Split(strIDs, ",")) = 1 Then
                    mblnOK = True: Unload Me: Exit Sub
                End If
            End If
            
            vsItem.Appearance = flexFlat
            Call zlControl.FormSetCaption(Me, False, False)
            Me.Left = mlngX: Me.Height = 3240
            lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
            If mlngY + mlngH + Me.Height > lngScrH Then
                Me.Top = mlngY - Me.Height
            Else
                Me.Top = mlngY + mlngH
            End If
            
            Call Form_Resize
        Else
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mstr输入 = "" Then
        
         tvw_s.Left = 0
        tvw_s.Top = picInfo.Top + picInfo.Height + 30
        tvw_s.Height = Me.ScaleHeight - picCmd.Height - tvw_s.Top
        
        pic.Top = tvw_s.Top
        pic.Left = tvw_s.Left + tvw_s.Width
        pic.Height = tvw_s.Height
         
        vsItem.Top = tvw_s.Top
        vsItem.Left = pic.Left + pic.Width
        vsItem.Width = Me.ScaleWidth - tvw_s.Width - pic.Width
        vsItem.Height = tvw_s.Height
       
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.5 < 4100 Then
            cmdCancel.Left = 4100
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 200
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    Else
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        vsItem.Height = Me.ScaleHeight
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColPosition
    Call SaveColWidth
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strTmp As String
    Dim str类型 As String
 
    strSQL = _
    "Select 0 As 级, To_Number('99999999' || 类型) As ID, -null As 上级id, '中草药' As 名称" & vbNewLine & _
    "From 诊疗分类目录" & vbNewLine & _
    "Where 类型 =3 And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & vbNewLine & _
    "Group By 类型"
    strSQL = strSQL & " Union ALL " & _
    "Select Level As 级,-id As ID, Nvl(-上级id, To_Number('99999999' || 类型)) As 上级id, '[' || 编码 || ']' || 名称 As 名称" & vbNewLine & _
    "From 诊疗分类目录" & vbNewLine & _
    "Where  类型=3 And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & vbNewLine & _
    "Start With 上级id Is Null" & vbNewLine & _
    "Connect By Prior ID = 上级id"

    strSQL = strSQL & " Order by 级,名称"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
    tvw_s.Visible = True
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!名称, "Close")
        End If
        objNode.Tag = 3 '存放分类类型:0-非药品和卫材,1-西成药,2-中成药,3-中草药,7-卫生材料
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
        'tvw_s.Nodes(1).Selected = True
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

Private Sub FillList(Optional strIDs As String)
'功能：根据当前界面条件装入诊疗项目目录
'参数：blnClass=是否重建分类卡(应在树形项目改变时才重建)
'          strIDs=读取的项目ID集,用于判断输入时是否别名不同的同一个收费项目
    Dim objTab As MSComctlLib.Tab
    Dim objNode As Node, objItem As ListItem
    Dim arrClass As Variant, strClass As String
    Dim strInput As String
    Dim str分类ID As String
    Dim lng药房ID As Long, strStock As String
    Dim str药房单位 As String, str药房包装 As String
    Dim strMain As String, strSQL As String
    Dim strTmp As String, strSQLItem As String
    Dim i As Long
    Dim strWherePriceGrade As String
    Dim cllTemp As Collection, bln显示库存 As Boolean
    Dim bln检查药品库存 As Boolean
    
    strIDs = ""
    Set objNode = tvw_s.SelectedItem '输入匹配时,为Nothing

    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    Me.Refresh
    

    On Error GoTo errH
    Screen.MousePointer = 11
    '获取草药信息记录集
    Set mrsItem = GetChineDrugRecordset(mstr输入)

    '绑定数据
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '列属性调整
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.COLS - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.COLS - 1
        vsItem.ColKey(i) = UCase(Trim(vsItem.TextMatrix(0, i)))
        If InStr("单价,库存", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If UCase(Trim(vsItem.TextMatrix(0, i))) Like "*ID" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf Trim(vsItem.TextMatrix(0, i)) = "换算系数" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '记录原始列号,用于处理列顺序
    Next
    
    '恢复列顺序:应放在排序处理之前
    Call RestoreColPosition
    Call RestoreColWidth
    '排序处理:先排序,以便后面处理行号
    Call RestoreColSort
    
     '调用服务，给网格控件中增加“库存”列信息
    If mint中药形态 = 0 And mlng药房id > 0 Then
        If mrsItem.RecordCount > 0 Then
            bln显示库存 = InStr(1, mstrPrivs, ";显示库存;") > 0
            bln检查药品库存 = InStr(mstrPrivs, ";不检查库存;") = 0
            Call gobjPublicExpense.zlLoadStockFromService(vsItem, mrsItem, 0, 0, mlng药房id, 0, bln显示库存, _
                    bln检查药品库存, False, False)
        End If
     End If

    '卡片相关数据计算
    '------------------------------------------------------------------------
    With vsItem
        For i = 1 To vsItem.Rows - 1
            .TextMatrix(i, 0) = i
            .RowHeight(i) = vsItem.RowHeightMin
            '收集项目ID:只收集最多2个
            If mstr输入 <> "" Then
                If UBound(Split(strIDs, ",")) < 2 And Val(.TextMatrix(i, .ColIndex("ID"))) > 0 Then
                    If InStr(strIDs & ",", "," & Val(.TextMatrix(i, .ColIndex("ID"))) & ",") = 0 Then
                        strIDs = strIDs & "," & Val(.TextMatrix(i, .ColIndex("ID")))
                    End If
                End If
            End If
        Next
    End With

    '行号列宽度
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    
    '选择器外挂插件调用
    If CreatePlugIn(0) Then
        On Error Resume Next
        Call gobjPlugIn.AfterSelectorReady(99, "中药选择器", vsItem, mfrmParent)
        Call zlPlugInErrH(Err, "AfterSelectorReady")
        Err.Clear: On Error GoTo errH
    End If
    
    vsItem.Redraw = flexRDDirect

    Call Form_Resize
    
    If Val(vsItem.TextMatrix(1, vsItem.ColIndex("id"))) = 0 Then mrsItem.Filter = "id=-1"
    If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst

    Screen.MousePointer = 0
    Exit Sub
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Sub

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If mstrKey = Node.Key Then Exit Sub
    mstrKey = Node.Key
    Call FillList
    
End Sub

Private Function GetSubTree(ByVal objNode As Node) As String
'功能：返回一个结点的子树结点的Key(含该结点)
    Dim strKeys As String
    Dim objTmp As Node
    
    strKeys = "," & Mid(objNode.Key, 2) & strKeys
    Set objTmp = objNode.Child
    Do While Not objTmp Is Nothing
        If objTmp.Children > 0 Then
            strKeys = "," & GetSubTree(objTmp) & strKeys
        Else
            strKeys = "," & Mid(objTmp.Key, 2) & strKeys
        End If
        Set objTmp = objTmp.Next
    Loop
    GetSubTree = Mid(strKeys, 2)
End Function

Private Function GetChineDrugRecordset(ByVal strInput As String) As ADODB.Recordset
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取草药信息记录集
    '入参：strInput-要查找的值
    '出参：
    '返回：记录集
    '------------------------------------------------------------------------------------------------------------------------
    Dim str特性 As String, str规格 As String, str特准项目 As String
    Dim strSQL As String, str分类ID As String
    Dim str撤档时间 As String, strWhere As String
    Dim lng分类id As Long
    Dim objNode As Node
    Dim bln显示下级 As Boolean
    
    On Error GoTo errHandle

    '特殊药品权限
    str特性 = ""
    If InStr(mstrPrivs, ";麻醉药品记帐;") = 0 Then str特性 = str特性 & " And E.毒理分类<>'麻醉药'"
    If InStr(mstrPrivs, ";毒性药品记帐;") = 0 Then str特性 = str特性 & " And E.毒理分类<>'毒性药'"
    If InStr(mstrPrivs, ";贵重药品记帐;") = 0 Then str特性 = str特性 & " And E.价值分类 Not IN('贵重','昂贵')"
    bln显示下级 = False
    Set objNode = tvw_s.SelectedItem '输入匹配时,为Nothing
    If mstr输入 = "" Then
        lng分类id = -1 * Val(Mid(objNode.Key, 2))
         '树形中的分类ID
        If bln显示下级 Then
            '显示下级的项目
            If Mid(objNode.Key, 2) = "99999999" & objNode.Tag Then
                str分类ID = " And A.分类ID IN(Select ID From 诊疗分类目录 Where 类型=3)"
            Else
                str分类ID = " And A.分类ID IN(Select ID From 诊疗分类目录 Start With ID=[9]Connect by Prior ID=上级ID)"
            End If
        Else
                str分类ID = " And A.分类ID=[9]"
        End If
    Else
        If Len(mstr输入) < 2 Then mstrLike = "" '优化
    End If
    
    If mint中药形态 = 0 Then
        str规格 = _
        "   And Nvl(C.中药形态,0) = [6] And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([7],3)" & _
        "   And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) "
    Else
         str规格 = " And Exists(Select 1 From 药品规格 C Where C.药名ID=E.药名ID And Nvl(C.中药形态,0) = [6])"
    End If
    
    str撤档时间 = "" & _
        "   And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) " & _
        "   And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        "   And A.服务对象 IN([7],3)"
    
    str特准项目 = ""
    If mint中药形态 = 0 Then
        If mint险类 <> 0 Then
            '刘兴洪:24862
            If zl_Check特准项目(gclsInsure, mint险类, mlng病人ID, False) Then str特准项目 = Get保险特准项目(mlng病人ID, "D.ID")
        End If
    End If
        
    If strInput <> "" Then
        strWhere = " And (A.编码 Like [1] And B.码类=[3] Or B.名称 Like [2] And B.码类=[3] Or B.简码 Like upper([2]) And B.码类 IN([3],3))"
        If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strWhere = " And (A.编码 Like [1] And B.码类=[3] Or B.简码 Like Upper([2]) And B.码类=3)"
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strWhere = " And B.简码 Like Upper([2]) And B.码类=[3]"
        ElseIf zlCommFun.IsCharChinese(strInput) Then
            strWhere = " And B.名称 Like [2] And B.码类=[3]"
        End If
         '非散装时按品种显示，且不显示库存
        strSQL = "" & _
        "   Select  distinct A.ID,A.编码,A.名称,A.计算单位" & _
        "   From 诊疗项目目录 A,诊疗项目别名 B" & _
        "   Where A.ID=B.诊疗项目ID  And A.类别='7' " & str撤档时间 & strWhere
        
        If mint中药形态 = 0 Then
            '散装才显示到规格级,保持原来不变
            strSQL = _
            " Select distinct  A.ID as 药名ID,C.药品ID as ID,C.药品ID,D.编码,A.名称,D.规格,A.计算单位 as 剂量单位," & _
                    IIf(gbln药房单位, "C." & gstr药房单位, "D.计算单位") & " as 单位,D.产地,D.费用类型,d.执行科室 AS 执行科室_ID," & IIf(mint险类 <> 0, "N.名称 医保大类,", "") & _
            "       Decode(D.是否变价,1,'时价',LTrim(To_Char(Sum(F.现价)" & _
                    IIf(gbln药房单位, "*Nvl(C." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as 单价," & _
            "       NULL as 库存," & IIf(gbln药房单位, "C." & gstr药房包装, "1 ") & " as 换算系数,7 as 类别id" & _
            " From 药品特性 E,药品规格 C,收费项目目录 D,收费价目 F, " & vbNewLine & _
                        IIf(mint险类 <> 0, "保险支付项目 M,保险支付大类 N,", "") & vbNewLine & _
            "          (" & strSQL & ") A " & vbNewLine & _
            " Where   A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID  " & vbNewLine & _
            "        And D.ID=F.收费细目ID " & vbNewLine & _
                     IIf(mint险类 <> 0, " And C.药品ID=M.收费细目ID(+) And M.险类(+)=[5] And M.大类ID=N.ID(+)" & vbNewLine, "") & _
            "        And exists(Select 1 From 收费执行科室 A1 Where A1.收费细目ID=C.药品ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine & _
            "        And Sysdate Between F.执行日期 and Nvl(F.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
                     str规格 & str特性 & str特准项目 & _
            " Group by A.ID,C.药品ID,A.计算单位,D.编码,A.名称,D.规格,D.产地,D.费用类型,d.执行科室,D.是否变价," & IIf(mint险类 <> 0, "N.名称,", "") & _
                    IIf(gbln药房单位, "C.门诊单位,C.门诊包装", "D.计算单位") & _
            " Order by D.编码"
        Else
             '非散装时按品种显示，且不显示库存
            strSQL = strSQL & _
            "        And exists(Select 1 From 诊疗执行科室 A1 Where A1.诊疗项目ID=A.ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine
            strSQL = _
                " Select Distinct A.ID,A.ID as 药名ID,A.编码,A.名称,A.计算单位 as 单位" & _
                " From 药品特性 E,(" & strSQL & ") A" & _
                " Where A.ID=E.药名ID  " & _
                "         And Exists(Select 1 From 药品规格 C Where C.药名ID=E.药名ID And Nvl(C.中药形态,0) = [6])" & _
                "         And Rownum<=100" & _
                " Order by A.编码"
        End If
    Else
        If mint中药形态 = 0 Then
            '散装才显示到规格级,保持原来不变
            strSQL = "" & _
            "  Select 药品ID As Id,药名ID,上级ID,编码,名称,规格,剂量单位,单位,产地,费用类型,执行科室_ID,单价,库存,药品ID,换算系数,7 as 类别id " & _
            "  From ( " & _
            " Select A.ID,A.ID as 药名ID,A.分类ID as 上级ID,D.编码,D.名称,D.规格,A.计算单位 as 剂量单位," & _
                        IIf(gbln药房单位, " C." & gstr药房单位, "D.计算单位") & " as 单位,D.产地,D.费用类型,d.执行科室 as 执行科室_ID" & IIf(mint险类 = 0, "", ",N.名称 医保大类") & "," & _
            "           Decode(D.是否变价,1,'时价',LTrim(To_Char(Sum(F.现价)" & _
                        IIf(gbln药房单位, "*Nvl(C." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "'))) as 单价," & _
            "           NULL as 库存,C.药品ID, " & IIf(gbln药房单位, "C." & gstr药房包装, "1 ") & " as 换算系数" & _
            " From  诊疗项目目录 A,药品特性 E,药品规格 C,收费项目目录 D,收费价目 F" & _
                        IIf(mint险类 = 0, "", ",保险支付项目 M,保险支付大类 N") & _
            " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID =F.收费细目ID And A.类别='7'  " & _
                    IIf(mint险类 = 0, "", "       And C.药品ID=M.收费细目ID(+) And   M.险类(+)=" & mint险类 & " And M.大类ID=N.ID(+)") & _
            "        And exists(Select 1 From 收费执行科室 A1 Where A1.收费细目ID=C.药品ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine & _
            "       And Sysdate Between F.执行日期 and Nvl(F.终止日期,TO_DATE('3000-01-01','YYYY-MM-DD'))" & _
            "       And D.服务对象 IN(" & mint病人来源 & ",3)" & str特准项目 & str规格 & str撤档时间 & str分类ID & _
            " Group by A.ID,A.计算单位 ,A.分类ID,D.编码,D.名称,D.规格,D.产地,D.费用类型,d.执行科室" & IIf(mint险类 = 0, "", ",N.名称") & ",D.是否变价,C.药品ID," & _
                 IIf(gbln药房单位, "C.门诊单位,C.门诊包装", "D.计算单位") & _
            ")"
        Else
            '非散装时按品种显示，且不显示库存
            strSQL = "" & _
            "Select Distinct A.ID,ID as 药名ID,A.分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,E.处方职务 as 处方职务ID" & _
            " From 诊疗项目目录 A,药品特性 E" & _
            " Where A.ID=E.药名ID" & str特性 & str撤档时间 & str规格 & str分类ID & _
            "        And exists(Select 1 From 诊疗执行科室 A1 Where A1.诊疗项目ID=A.ID And A1.执行科室ID=[4]   And (A1.病人来源 is NULL Or A1.病人来源=[7]) and (A1.开单科室ID is null or A1.开单科室ID=[8])  ) " & vbNewLine
        End If
    End If
    Set GetChineDrugRecordset = zlDatabase.OpenSQLRecord(strSQL, "中草药", strInput & "%", mstrLike & strInput & "%", gbytCode + 1, mlng药房id, mint险类, mint中药形态, mint病人来源, mlng开单科室ID, lng分类id)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveColPosition(Optional ByVal strType As String)
'功能：保存列顺序:列号,顺序|...
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        For i = 0 To .COLS - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'功能：保存列宽度
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Not gblnMyStyle Then Exit Sub
    If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'功能：恢复列宽度
'说明：应放在恢复列序之后
    Dim strType As String
    
    If Not gblnMyStyle Then Exit Sub
    
    If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub


Private Sub RestoreColPosition()
'功能：恢复列顺序
'说明：应放在排序处理之前
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Not gblnMyStyle Then Exit Sub
    
    With vsItem
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,改变后相关列号也改变
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .COLS - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .COLS - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub RestoreColSort()
'功能：排序处理
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .COLS - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .COLS - 1) = 7
        If gblnMyStyle Then
            If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
            strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '因为可能调整列顺序,所以查找真实的排序列
                For i = 0 To .COLS - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .COLS - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = img.ListImages(3).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = img.ListImages(4).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        cmdOK.Enabled = Val(vsItem.TextMatrix(NewRow, 1)) <> 0
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .COLS - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = img.ListImages(3).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = img.ListImages(4).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '因为可能列顺序改变,所以保存原始列号
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    '强制编码列按字符串排序
    If vsItem.TextMatrix(0, Col) = "编码" Then
        If Order = 1 Then Order = 7
        If Order = 2 Then Order = 8
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        Call vsItem_KeyPress(13)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then cmdOK_Click
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub


