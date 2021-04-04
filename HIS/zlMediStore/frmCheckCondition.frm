VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "盘点条件设置"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "基本(&1)"
      TabPicture(0)   =   "frmCheckCondition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl剂型"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lbl盘点方式"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl库房"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Lvw剂型"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dtpDate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkNoNum"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkNum"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Chk剂型"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Cbo盘点方式"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbo库房"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chk忽略盘点时间"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "药品分类(&2)"
      TabPicture(1)   =   "frmCheckCondition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tvw分类"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "库房货位(&3)"
      TabPicture(2)   =   "frmCheckCondition.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chk货位"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "vsfStock"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CheckBox chk货位 
         Caption         =   "仅显示当前库房已分配的货位"
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox chk忽略盘点时间 
         Caption         =   "始终以当前库存作为帐面数量"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   4800
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   420
         Width           =   3045
      End
      Begin VB.ComboBox Cbo盘点方式 
         Height          =   300
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3180
         Width           =   3045
      End
      Begin VB.CheckBox Chk剂型 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3555
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   810
         Width           =   675
      End
      Begin VB.CheckBox chkNum 
         Caption         =   "盘无库存记录药品"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   4020
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkNoNum 
         Caption         =   "仅盘无数量，但有库存金额或差价的药品"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   4425
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1140
         TabIndex        =   7
         Top             =   3600
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   162136067
         CurrentDate     =   36901
      End
      Begin MSComctlLib.ListView Lvw剂型 
         Height          =   2085
         Left            =   420
         TabIndex        =   10
         Top             =   1020
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   3678
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.TreeView tvw分类 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   8705
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Width           =   4575
         _cx             =   8070
         _cy             =   8070
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckCondition.frx":0054
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
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
      Begin VB.Label lbl库房 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "库房(&S)"
         Height          =   180
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Lbl盘点方式 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "方式(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   13
         Top             =   3240
         Width           =   630
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间(&T)"
         Height          =   180
         Left            =   480
         TabIndex        =   12
         Top             =   3660
         Width           =   630
      End
      Begin VB.Label Lbl剂型 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "剂型(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   480
         TabIndex        =   11
         Top             =   810
         Width           =   630
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5160
      TabIndex        =   2
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   0
      Top             =   405
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5160
      Top             =   1230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCondition.frx":00C4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Private mblnBootUp As Boolean
Private mstr剂型 As String
Private mlng库房ID As Long
Private mlng盘点方式 As Integer
Private mstr盘点时间 As String
Private mint盘无库存药品 As Integer
Private mstr库房货位  As String
Private mbln中药库房 As Boolean                         '如果是中药房或中药库，则该变量为真
Private mfrmMain As Form
Private mblnCheckNoNum  As Boolean
Private mstr分类ID As String
Private mbln忽略盘点时间 As Boolean
Private mstrSelRows As String
Private mint盘点时间范围 As Integer         '用来模块参数设置的盘点时间范围

Private Sub CheckItem(ByVal str编码 As String, ByVal intChecked As Integer)
    Dim lngRows As Long
    
    With vsfStock
        If str编码 = "所有" Then
            '当前是根节点，所有库房设置为全选或全不选
            For lngRows = 2 To .rows - 1
                .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = intChecked
            Next
        Else
            '当前是子节点，需要递归处理下级子节点
            For lngRows = 2 To .rows - 1
                If .TextMatrix(lngRows, .ColIndex("上级")) = str编码 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = intChecked
                    
                    Call CheckItem(.TextMatrix(lngRows, .ColIndex("编码")), intChecked)
                End If
            Next
        End If
    End With
End Sub
Public Function GetCondition(FrmMain As Form, ByRef str剂型编码 As String, _
    ByRef lng库房ID As Long, ByRef 盘点方式 As Integer, ByRef str盘点时间 As String, _
    ByRef int盘无库存药品 As Integer, ByRef str库房货位 As String, ByRef bln盘无库存有金额药品 As Boolean, _
    ByRef str分类ID As String, ByRef bln忽略盘点时间 As Boolean) As Boolean
    
    mstr剂型 = ""
    mlng库房ID = 0
    mlng盘点方式 = 0
    mstr盘点时间 = ""
    mint盘无库存药品 = 0
    mstr库房货位 = "所有"
    mblnSelect = False
    mblnCheckNoNum = False
    mbln忽略盘点时间 = False
    
    Set mfrmMain = FrmMain
    Me.Show vbModal, FrmMain
    GetCondition = mblnSelect
    
    str剂型编码 = mstr剂型
    lng库房ID = mlng库房ID
    盘点方式 = mlng盘点方式
    str盘点时间 = mstr盘点时间
    int盘无库存药品 = mint盘无库存药品
    str库房货位 = mstr库房货位
    bln盘无库存有金额药品 = mblnCheckNoNum
    str分类ID = mstr分类ID
    bln忽略盘点时间 = mbln忽略盘点时间
    
End Function

Private Sub GetSubItem(ByVal str上级编码 As String, ByVal rsData As ADODB.Recordset)
    '用递归算法找树表的子项目
    Dim rsClone As ADODB.Recordset
    
    Set rsClone = rsData.Clone
    
    rsClone.Filter = "上级='" & str上级编码 & "'"
    rsClone.Sort = "名称"
    
    '没找到下一级时一定要退出
    If rsClone.RecordCount = 0 Then Exit Sub
    
    With vsfStock
        .Redraw = flexRDNone
        
        Do While Not rsClone.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("级数")) = rsClone!级数 + 1
            .TextMatrix(.rows - 1, .ColIndex("上级")) = rsClone!上级
            .TextMatrix(.rows - 1, .ColIndex("名称")) = rsClone!名称
            .TextMatrix(.rows - 1, .ColIndex("选择")) = 0
            .TextMatrix(.rows - 1, .ColIndex("编码")) = rsClone!编码
            
            '找下一级的项目
            Call GetSubItem(rsClone!编码, rsData)
            
            rsClone.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub Load库房货位()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim intLevel As Integer
    
    On Error GoTo errHandle
    If chk货位.Value = 1 Then
        gstrSQL = "Select a.编码, a.名称 From 药品库房货位 A " & _
            " Where a.库房id = [1] And Exists (Select 1 From 药品货位对照 B Where b.库房id = a.库房id And b.货位id = a.Id) " & _
            " Order By 名称 "
    Else
        gstrSQL = "Select 编码, 名称 From  药品库房货位 Where 库房id = [1] Order By 名称 "
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有药品库房货位", Val(cbo库房.ItemData(cbo库房.ListIndex)))
    
    If rsData.RecordCount = 0 Then
        vsfStock.rows = 1
        Exit Sub
    End If
    
    With vsfStock
        .Redraw = flexRDNone
        .rows = 1
        
        Do While Not rsData.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("选择")) = 0
            .TextMatrix(.rows - 1, .ColIndex("编码")) = rsData!编码
            .TextMatrix(.rows - 1, .ColIndex("名称")) = rsData!名称
            
            rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo库房_Click()
    Dim blnEXIST As Boolean
    Dim rsTemp As New ADODB.Recordset
    '提取该库房现有剂型，供用户选择
    mbln中药库房 = False
    On Error GoTo errHandle
    gstrSQL = "Select 1 From 部门性质说明 " & _
             " Where 工作性质 Like '中药%' And 部门ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", Me.cbo库房.ItemData(cbo库房.ListIndex))

    If Not rsTemp.EOF Then mbln中药库房 = True
    
    gstrSQL = "Select Distinct J.编码,J.名称 " & _
             " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
             " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
             " And A.执行科室ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", Me.cbo库房.ItemData(cbo库房.ListIndex))
    Lvw剂型.ListItems.Clear
    
    With rsTemp
        Do While Not .EOF
            If blnEXIST = False Then
                blnEXIST = (!名称 = "方剂")
            End If
            Lvw剂型.ListItems.Add , "K" & !编码, !名称, , 1
            .MoveNext
        Loop
        If mbln中药库房 And blnEXIST = False Then
            Lvw剂型.ListItems.Add , "KK1", "方剂", , 1
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkNoNum_Click()
    chkNum.Enabled = chkNoNum.Value = 0 '不勾选“仅盘无数量，但有库存金额或差价的药品”时，“盘无库存记录药品”才可用
End Sub

Private Sub chkNum_Click()
    chkNoNum.Enabled = chkNum.Value = 0 '不勾选“盘无库存记录药品”时，“仅盘无数量，但有库存金额或差价的药品”才可用
End Sub

Private Sub chk货位_Click()
    Load库房货位
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    
    '取得剂型（不选剂型则不提取药品，需要手工录入）
    mstr剂型 = ""
    
    If Chk剂型.Value = 1 Then
        mstr剂型 = "所有"
    Else
        intItems = Me.Lvw剂型.ListItems.count
        For intItem = 1 To intItems
            If Lvw剂型.ListItems(intItem).Checked Then
                mstr剂型 = mstr剂型 & "," & Lvw剂型.ListItems(intItem).Text
            End If
        Next
    
        If mstr剂型 <> "" Then mstr剂型 = Mid(mstr剂型, 2)
    End If

    '取得药品分类（不选分类表示所有分类）
    mstr分类ID = ""
    For intItem = 1 To tvw分类.Nodes.count
        If tvw分类.Nodes(intItem).Key = "Root" And tvw分类.Nodes(intItem).Checked = True Then
            mstr分类ID = ""
            Exit For
        ElseIf tvw分类.Nodes(intItem).Key <> "Root" And _
            tvw分类.Nodes(intItem).Key <> "_中成药" And _
            tvw分类.Nodes(intItem).Key <> "_中草药" And _
            tvw分类.Nodes(intItem).Key <> "_西成药" And _
            tvw分类.Nodes(intItem).Checked Then
            mstr分类ID = mstr分类ID & "," & Mid(tvw分类.Nodes(intItem).Key, 2)
        End If
    Next

    If mstr分类ID <> "" Then
        mstr分类ID = Mid(mstr分类ID, 2)
    End If
    
    mlng库房ID = cbo库房.ItemData(cbo库房.ListIndex)
    mlng盘点方式 = Cbo盘点方式.ItemData(Cbo盘点方式.ListIndex)
    mstr盘点时间 = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    mblnSelect = True
    mint盘无库存药品 = chkNum.Value
    mbln忽略盘点时间 = (chk忽略盘点时间.Value = 1)
    
    '取得库房货位（不选库房表示不考虑存储库房）
    mstr库房货位 = ""
    With vsfStock
        For intItem = 1 To .rows - 1
            If .Cell(flexcpChecked, intItem, .ColIndex("选择")) = flexChecked Then
                mstr库房货位 = .TextMatrix(intItem, .ColIndex("名称")) & "," & mstr库房货位
            End If
        Next
    End With
    
'    If mstr库房货位 <> "" Then
'        mstr库房货位 = Mid(mstr库房货位, 2)
'    End If
    
    mblnCheckNoNum = chkNoNum.Value
    
    frmNewCheckCard.txtStock.Caption = cbo库房.Text
    frmNewCheckCard.txtStock.Tag = mlng库房ID
    frmNewCheckCard.txtCheckDate = mstr盘点时间
'    frmCheckCard.CmdSave.Enabled = False
'    frmCheckCard.cmdCancel.Enabled = False
    
    Unload Me
End Sub



Private Sub Command1_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnBootUp = False Then
        Unload Me
        Exit Sub
    End If
    
    Call Load库房货位
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim blnSelectStock As String
    Dim objnode As Node
    
    On Error GoTo errHandle
    blnSelectStock = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品盘点管理", "库房", "0")
    mint盘点时间范围 = Val(zlDatabase.GetPara("盘点时间范围设置", glngSys, 1307, 30))
    dtpDate.MinDate = CDate(Format(DateAdd("d", -mint盘点时间范围, Date), "yyyy-mm-dd") & " 00:00:00")
    '药品材质权限控制
    
    dtpDate.Value = Format(Sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    mblnBootUp = False

    With Cbo盘点方式
        .Clear
        .AddItem "每日"
        .ItemData(.NewIndex) = 1
        .AddItem "每周"
        .ItemData(.NewIndex) = 2
        .AddItem "每月"
        .ItemData(.NewIndex) = 3
        .AddItem "每季度"
        .ItemData(.NewIndex) = 4
        .AddItem "忽略盘点方式"
        .ItemData(.NewIndex) = 5
        .ListIndex = 0
    End With
    
    With mfrmMain.cboStock
        cbo库房.Clear
        For i = 0 To .ListCount - 1
            cbo库房.AddItem .List(i)
            cbo库房.ItemData(cbo库房.NewIndex) = .ItemData(i)
        Next
        cbo库房.ListIndex = .ListIndex
    End With
        
    If zlStr.IsHavePrivs(gstrprivs, "所有库房") Then
        If blnSelectStock = "0" Then
            cbo库房.Enabled = False
        Else
            cbo库房.Enabled = True
        End If
    Else
        cbo库房.Enabled = False
    End If
    
    '库房货位
    Load库房货位
    
    '药品分类
    gstrSQL = "Select Level as 层,ID,上级ID,名称,DECODE(类型,1,'西成药',2,'中成药','中草药') As 材质 " & _
        " From 诊疗分类目录" & _
        " Where 类型 in (1,2,3)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Order by Level,编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取药品分类")

    Set objnode = tvw分类.Nodes.Add(, , "Root", "所有用途", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_西成药", "西成药", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_中成药", "中成药", 1)

    Do While Not rsTemp.EOF
        If rsTemp!层 = 1 Then
            Set objnode = tvw分类.Nodes.Add("_" & rsTemp!材质, 4, "_" & rsTemp!id, rsTemp!名称, 1)
        Else
            Set objnode = tvw分类.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!id, rsTemp!名称, 1)
        End If
        rsTemp.MoveNext
    Loop
    tvw分类.Nodes("Root").Selected = True
    tvw分类.Nodes("Root").Expanded = True
    
    mblnBootUp = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Chk剂型_Click()
    If Chk剂型.Value = 2 Then Exit Sub
    Call SetSelect(Lvw剂型, Chk剂型.Value)
End Sub

Private Sub Lvw剂型_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Call ItemCheck(Lvw剂型, Item, Chk剂型)
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

Private Sub tvw分类_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode tvw分类, Node, Node.Checked
End Sub
Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Sub SetParentNode(ByVal objMyTreeView As TreeView, ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If objMyTreeView.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = objMyTreeView.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If objMyTreeView.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode objMyTreeView, Node, blnCheck
        End If
    End If
End Sub

Private Sub vsfStock_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    Dim lngRow As Long
    Dim strFlag As String

    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("选择") Then Exit Sub
        
        mstrSelRows = ""

        For lngRow = 1 To .rows - 1
            If .IsSelected(lngRow) Then
                mstrSelRows = IIf(mstrSelRows = "", "", mstrSelRows & ",") & lngRow
            End If
        Next
    End With
End Sub


Private Sub vsfStock_Click()
    Dim IntCheck As Integer
    Dim lngRow As Long
    Dim lngRows As Long
    
    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("选择") Then Exit Sub
        
        lngRow = .MouseRow
        
        IntCheck = .Cell(flexcpChecked, lngRow, .ColIndex("选择"))

        If InStr(1, mstrSelRows, ",") > 0 Then
            '多选处理
            For lngRows = 1 To .rows - 1
                If InStr(1, "," & mstrSelRows & ",", "," & lngRows & ",") > 0 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = IntCheck
                End If
            Next
            
            mstrSelRows = ""
        Else
            '单选时处理当前节点
            .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = IntCheck
        End If
        
    End With
End Sub
