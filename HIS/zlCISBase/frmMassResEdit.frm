VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMassResEdit 
   BorderStyle     =   0  'None
   Caption         =   "质控品信息"
   ClientHeight    =   7890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox chkHide 
      Caption         =   "隐藏中文名"
      Height          =   195
      Left            =   7335
      TabIndex        =   21
      Top             =   98
      Value           =   1  'Checked
      Width           =   1350
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   4605
      Left            =   3240
      TabIndex        =   17
      Top             =   330
      Width           =   5460
      _cx             =   9631
      _cy             =   8123
      Appearance      =   2
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
      Rows            =   3
      Cols            =   3
      FixedRows       =   2
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
      AutoSizeMouse   =   -1  'True
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
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7845
      ScaleMode       =   0  'User
      ScaleWidth      =   3225
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3225
      Begin VB.ComboBox cbo校准物 
         Height          =   300
         Left            =   405
         TabIndex        =   28
         Top             =   4785
         Width           =   2745
      End
      Begin VB.ComboBox cbo试剂 
         Height          =   300
         Left            =   405
         TabIndex        =   26
         Top             =   4200
         Width           =   2745
      End
      Begin VB.TextBox txt取值序列 
         Height          =   300
         Left            =   405
         TabIndex        =   23
         Top             =   3090
         Width           =   2745
      End
      Begin VB.TextBox txt序列值 
         Height          =   300
         Left            =   405
         TabIndex        =   22
         Top             =   3645
         Width           =   2745
      End
      Begin VB.CheckBox chk非定值 
         Caption         =   "非定值 (非定值不能预设值)"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Width           =   2925
      End
      Begin VB.ComboBox cbo水平 
         Height          =   300
         Left            =   2310
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   825
      End
      Begin VB.ComboBox cbo方法 
         Height          =   300
         Left            =   330
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   7665
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox txt浓度 
         Height          =   300
         Left            =   600
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1710
      End
      Begin VB.TextBox txt批号 
         Height          =   300
         Left            =   600
         MaxLength       =   10
         TabIndex        =   2
         Top             =   405
         Width           =   1155
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   600
         MaxLength       =   40
         TabIndex        =   4
         Top             =   735
         Width           =   2520
      End
      Begin VB.TextBox txt标本号 
         Height          =   300
         Left            =   390
         MaxLength       =   40
         TabIndex        =   16
         ToolTipText     =   "可以指定该质控物使用的标本号，标本号之间以,分隔"
         Top             =   2460
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   300
         Left            =   390
         TabIndex        =   12
         Top             =   1905
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   101711875
         CurrentDate     =   39064
      End
      Begin MSComCtl2.DTPicker dtp结束日期 
         Height          =   300
         Left            =   1890
         TabIndex        =   14
         Top             =   1905
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   101711875
         CurrentDate     =   39429
      End
      Begin VB.Label lbl校准物 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "校准物:"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   4545
         Width           =   630
      End
      Begin VB.Label lbl试剂来源 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "试剂:"
         Height          =   180
         Left            =   210
         TabIndex        =   27
         Top             =   3960
         Width           =   450
      End
      Begin VB.Label lbl取值序列 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取值序列:"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label lbl序列值 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "取值序列对应数字:"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   3435
         Width           =   1530
      End
      Begin VB.Label lbl基本信息 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "控制品信息:"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   150
         Width           =   990
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMassResEdit.frx":0000
         ForeColor       =   &H00008000&
         Height          =   2340
         Left            =   360
         TabIndex        =   18
         Top             =   5235
         Width           =   2865
         WordWrap        =   -1  'True
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   45
         Picture         =   "frmMassResEdit.frx":0153
         Top             =   5205
         Width           =   240
      End
      Begin VB.Label lbl结束日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   180
         Left            =   1680
         TabIndex        =   13
         Top             =   1965
         Width           =   180
      End
      Begin VB.Label lbl名称 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         Height          =   180
         Left            =   195
         TabIndex        =   3
         Top             =   795
         Width           =   360
      End
      Begin VB.Label lbl批号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "批号"
         Height          =   180
         Left            =   195
         TabIndex        =   1
         Top             =   465
         Width           =   360
      End
      Begin VB.Label lbl方法 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "方法"
         Height          =   180
         Left            =   60
         TabIndex        =   8
         Top             =   7710
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lbl浓度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "浓度"
         Height          =   180
         Left            =   195
         TabIndex        =   5
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label lbl开始日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "使用日期范围:"
         Height          =   180
         Left            =   195
         TabIndex        =   11
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label lbl标本号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "对应标本号:"
         Height          =   180
         Left            =   195
         TabIndex        =   15
         Top             =   2250
         Width           =   990
      End
   End
   Begin VB.Label lbl质控项目 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "质控检测项目: "
      Height          =   180
      Left            =   3240
      TabIndex        =   20
      Top             =   105
      Width           =   1260
   End
End
Attribute VB_Name = "frmMassResEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngResId As Long          '当前显示的物品id
Private mlngDevId As Long          '当前显示的仪器id

Private Enum mCol
    ID = 0: 选择: 中文名: 英文名: 单位: 取值序列: 参考靶值: 参考SD: 预设均值: 预设SD: 项目qc码: 方法qc码: 结果类型: 质控取值: 方法: 序列值
End Enum
Private mstr方法 As String '用于表格中的选择功能
Private mlng仪器类型 As Long '0-普通仪器 1-微物生 2-酶标仪
Dim lngCount As Long
Dim lngLastID As Long
Private mblnEditRow As Boolean  '是否修改了序列值

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置参考值列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    Dim strLists As String, strValue As String
    
    With Me.vfgList
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 2: .FixedRows = 2: .Cols = 16: .FixedCols = 0
        End If
        If .Cols < 16 Then .Cols = 16
        .MergeCells = flexMergeFixedOnly: .MergeRow(0) = True: .MergeCol(mCol.方法) = True: .MergeCol(mCol.质控取值) = True
       
        .TextMatrix(0, mCol.选择) = "项目": .TextMatrix(0, mCol.中文名) = "项目": .TextMatrix(0, mCol.英文名) = "项目"
        .TextMatrix(0, mCol.单位) = "项目": .TextMatrix(1, mCol.取值序列) = "项目"
        .TextMatrix(1, mCol.选择) = "": .TextMatrix(1, mCol.中文名) = "中文名": .TextMatrix(1, mCol.英文名) = "英文名"
        .TextMatrix(1, mCol.单位) = "单位": .TextMatrix(1, mCol.取值序列) = "取值序列"
        .TextMatrix(0, mCol.参考靶值) = "参考控制值": .TextMatrix(0, mCol.参考SD) = "参考控制值"
        .TextMatrix(1, mCol.参考靶值) = "靶值": .TextMatrix(1, mCol.参考SD) = "SD"
        .TextMatrix(0, mCol.预设均值) = "预设控制值": .TextMatrix(0, mCol.预设SD) = "预设控制值"
        .TextMatrix(1, mCol.预设均值) = "均值": .TextMatrix(1, mCol.预设SD) = "SD"
        .TextMatrix(0, mCol.项目qc码) = "对应QC码": .TextMatrix(0, mCol.方法qc码) = "对应QC码"
        .TextMatrix(1, mCol.项目qc码) = "项目码": .TextMatrix(1, mCol.方法qc码) = "方法码"
        .TextMatrix(0, mCol.方法) = "方法": .TextMatrix(1, mCol.方法) = "方法"
        .TextMatrix(0, mCol.结果类型) = "结果类型": .TextMatrix(1, mCol.结果类型) = "结果类型"
        .TextMatrix(0, mCol.序列值) = "序列值": .TextMatrix(1, mCol.序列值) = "序列值"
        
        .TextMatrix(0, mCol.质控取值) = "质控取值": .TextMatrix(1, mCol.质控取值) = "质控取值"
        
        
        .ColWidth(mCol.中文名) = IIf(Me.chkHide.Value = vbChecked, 0, 2000)
        .ColWidth(mCol.英文名) = 800: .ColWidth(mCol.单位) = 800: .ColWidth(mCol.取值序列) = 0
        .ColWidth(mCol.参考靶值) = 700: .ColWidth(mCol.参考SD) = 700
        .ColWidth(mCol.预设均值) = 700: .ColWidth(mCol.预设SD) = 700
        .ColWidth(mCol.项目qc码) = 800: .ColWidth(mCol.方法qc码) = 800
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.选择) = 270
        .ColWidth(mCol.结果类型) = 0: .ColWidth(mCol.序列值) = 0
        
        .ColWidth(mCol.质控取值) = 0
        .ColHidden(mCol.质控取值) = False
        If mlng仪器类型 = 2 Then .ColWidth(mCol.质控取值) = 900
        
        .ColComboList(mCol.方法) = mstr方法
        .ColComboList(mCol.质控取值) = "|[OD]|[SCO]"
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, mCol.选择)) = 1 Then
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked
            Else
                .Cell(flexcpChecked, lngCount, mCol.选择) = flexUnchecked
            End If
            .TextMatrix(lngCount, mCol.选择) = ""
            If Trim(.TextMatrix(lngCount, mCol.取值序列)) <> "" And Trim(.TextMatrix(lngCount, mCol.质控取值)) = "" Then
                strLists = Trim(.TextMatrix(lngCount, mCol.取值序列))
                strValue = Trim(.TextMatrix(lngCount, mCol.参考靶值))
                If Val(strValue) = Int(Val(strValue)) And Val(strValue) > 0 And Val(strValue) <= UBound(Split(strLists, ";")) + 1 Then
                    .TextMatrix(lngCount, mCol.参考靶值) = Split(strLists, ";")(strValue - 1)
                Else
                    .TextMatrix(lngCount, mCol.参考靶值) = "": .TextMatrix(lngCount, mCol.参考SD) = ""
                End If
                strValue = Trim(.TextMatrix(lngCount, mCol.预设均值))
                If Val(strValue) = Int(Val(strValue)) And Val(strValue) > 0 And Val(strValue) <= UBound(Split(strLists, ";")) + 1 Then
                    .TextMatrix(lngCount, mCol.预设均值) = Split(strLists, ";")(strValue - 1)
                Else
                    .TextMatrix(lngCount, mCol.预设均值) = "": .TextMatrix(lngCount, mCol.预设SD) = ""
                End If
            End If
            If Left(.TextMatrix(lngCount, mCol.参考靶值), 1) = "." Then .TextMatrix(lngCount, mCol.参考靶值) = "0" & .TextMatrix(lngCount, mCol.参考靶值)
            If Left(.TextMatrix(lngCount, mCol.参考SD), 1) = "." Then .TextMatrix(lngCount, mCol.参考SD) = "0" & .TextMatrix(lngCount, mCol.参考SD)
            If Left(.TextMatrix(lngCount, mCol.预设均值), 1) = "." Then .TextMatrix(lngCount, mCol.预设均值) = "0" & .TextMatrix(lngCount, mCol.预设均值)
            If Left(.TextMatrix(lngCount, mCol.预设SD), 1) = "." Then .TextMatrix(lngCount, mCol.预设SD) = "0" & .TextMatrix(lngCount, mCol.预设SD)
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngResID As Long, lngDevId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    mlngResId = lngResID
    
    '清除此前项目的显示
    Me.txt名称.Text = "": Me.txt批号.Text = ""
    Me.txt浓度.Text = "": Me.cbo水平.Clear: Me.cbo方法.ListIndex = -1
    Me.dtp开始日期.Value = Now(): Me.dtp结束日期.Value = DateAdd("m", 13, Now()) - 1
    Me.txt标本号 = "": Me.chk非定值.Value = vbUnchecked
    '--------------------------------------------------
    ' 2009-09-03增加
    Me.cbo试剂.Text = "": Me.cbo校准物.Text = ""
    '--------------------------------------------------
    Err = 0: On Error GoTo errHand
    If lngDevId <> 0 And mlngDevId <> lngDevId Then
        mlngDevId = lngDevId
        gstrSql = "Select 微生物 From 检验仪器 where id=[1] "
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngDevId)
        Do Until rsTemp.EOF
            mlng仪器类型 = Val("" & rsTemp!微生物)
            rsTemp.MoveNext
        Loop
         
    End If
    
    If lngResID = 0 Then
        Call setListFormat:        zlRefresh = True: Exit Function
    End If
    '获取指定项目的信息
    gstrSql = "Select D.质控水平数, R.名称, R.批号, R.非定值, R.浓度, R.水平, R.方法, R.开始日期, R.结束日期, R.标本号, D.微生物, R.试剂, R.校准物" & vbNewLine & _
            "From 检验质控品 R, 检验仪器 D" & vbNewLine & _
            "Where R.仪器id  = D.ID And R.ID  = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt名称.Text = "" & !名称
            Me.txt批号.Text = "" & !批号
            Me.txt浓度.Text = "" & !浓度
            For lngCount = 0 To Me.cbo方法.ListCount - 1
                If Me.cbo方法.List(lngCount) = "" & !方法 Then Me.cbo方法.ListIndex = lngCount: Exit For
            Next
            For lngCount = 1 To Val("" & !质控水平数)
                Me.cbo水平.AddItem "水平" & lngCount
                If lngCount = Val("" & !水平) Then Me.cbo水平.ListIndex = Me.cbo水平.NewIndex
            Next
            
            If Not IsNull(!开始日期) Then Me.dtp开始日期.Value = !开始日期
            If Not IsNull(!结束日期) Then Me.dtp结束日期.Value = !结束日期
            Me.txt标本号.Text = "" & !标本号
            Me.chk非定值.Value = IIf(Val("" & !非定值) = 1, vbChecked, vbUnchecked)
            '--------------------------------------------------
            ' 2009-09-03增加
            Me.cbo试剂.Text = Trim("" & !试剂)
            Me.cbo校准物.Text = Trim("" & !校准物)
            '--------------------------------------------------
        End If
    End With
    
    gstrSql = "Select I.ID, 1 As 选择, I.中文名, I.英文名, I.单位, Decode(P.结果类型, 1, '',P.取值序列) As 取值序列," & vbNewLine & _
            "       K.靶值 As 参考靶值, K.Sd As 参考sd, X.均值 As 预设均值, X.Sd As 预设sd, K.项目qc码, K.方法qc码,P.结果类型, K.质控取值, K.方法, '' as 序列值" & vbNewLine & _
            "From 诊治所见项目 I, 检验项目 P, 检验质控品项目 K," & vbNewLine & _
            "     (Select X.质控品id, X.项目id, X.均值, X.Sd" & vbNewLine & _
            "       From 检验质控品 R, 检验质控均值 X" & vbNewLine & _
            "       Where R.ID = X.质控品id(+) And R.开始日期 = X.开始日期(+) And R.ID = [1]) X" & vbNewLine & _
            "Where I.ID = P.诊治项目id And I.ID = K.项目id And K.质控品id = X.质控品id(+) And K.项目id = X.项目id(+) And" & vbNewLine & _
            "      K.质控品id = [1]" & vbNewLine & _
            "Order By I.编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    Call Refresh对照(lngResID)
    Call Show参考(vfgList.Row)
    zlRefresh = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngResID As Long, lngDevId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngResId-增加的参照项目，或者指定编辑的项目
    '       lngDevId-当前增加物品的所属设备id
    Dim rsTemp As New ADODB.Recordset
    Dim str试剂 As String, str校准物 As String
    mlngDevId = lngDevId
    
    str试剂 = Trim(Me.cbo试剂.Text)
    str校准物 = Trim(Me.cbo校准物.Text)
    
    Err = 0: On Error GoTo errHand
    gstrSql = "Select 名称 From 质控试剂来源"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Me.cbo试剂.Clear: Me.cbo校准物.Clear
    Me.cbo试剂.AddItem "": cbo校准物.AddItem ""
    Do Until rsTemp.EOF
        Me.cbo试剂.AddItem Trim("" & rsTemp!名称)
        Me.cbo校准物.AddItem Trim("" & rsTemp!名称)
        rsTemp.MoveNext
    Loop
    
    cbo试剂.Text = str试剂: cbo校准物.Text = str校准物
    
    If blnAdd Then
        gstrSql = "Select 质控水平数 From 检验仪器 Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId)
        Me.cbo水平.Clear
        With rsTemp
            If .RecordCount > 0 Then
                For lngCount = 1 To Val("" & !质控水平数)
                    Me.cbo水平.AddItem "水平" & lngCount
                Next
            End If
        End With
        
        '清除并设置备注值
        Me.txt名称.Text = "": Me.txt批号.Text = ""
        Me.txt浓度.Text = ""
        Me.dtp开始日期.Value = Now(): Me.dtp结束日期.Value = DateAdd("m", 13, Now()) - 1
        Me.txt标本号 = ""
        Me.chk非定值.Value = vbUnchecked
        
        Me.cbo试剂.Text = "": Me.cbo校准物.Text = ""
        
    End If
    
'    gstrSql = "Select I.ID, Decode(K.项目id, Null, 0, 1) As 选择, I.中文名, I.英文名, I.单位, I.取值序列, K.靶值 As 参考靶值," & vbNewLine & _
'            "       K.Sd As 参考sd, X.均值 As 预设均值, X.Sd As 预设sd, K.项目qc码, K.方法qc码, K.方法" & vbNewLine & _
'            "From (Select I.ID, I.编码, I.中文名, I.英文名, I.单位, Decode(P.结果类型, 3, P.取值序列, '') As 取值序列" & vbNewLine & _
'            "       From 检验仪器项目 L, 诊治所见项目 I, 检验项目 P" & vbNewLine & _
'            "       Where L.项目id = I.ID And I.ID = P.诊治项目id And L.仪器id = [2] And" & vbNewLine & _
'            "             (P.结果类型 = 1 Or P.结果类型 = 3 And P.取值序列 Is Not Null)) I," & vbNewLine & _
'            "     (Select 质控品id, 项目id, 靶值, Sd, 项目qc码, 方法qc码, 方法 From 检验质控品项目 Where 质控品id = [1]) K," & vbNewLine & _
'            "     (Select X.质控品id, X.项目id, X.均值, X.Sd" & vbNewLine & _
'            "       From 检验质控品 R, 检验质控均值 X" & vbNewLine & _
'            "       Where R.ID = X.质控品id(+) And R.开始日期 = X.开始日期(+) And R.ID = [1]) X" & vbNewLine & _
'            "Where I.ID = K.项目id(+) And K.质控品id = X.质控品id(+) And K.项目id = X.项目id(+)" & vbNewLine & _
'            "Order By I.编码"

    gstrSql = "Select I.ID, Decode(K.项目id, Null, 0, 1) As 选择, I.中文名, I.英文名, I.单位, I.取值序列, K.靶值 As 参考靶值," & vbNewLine & _
            "       K.Sd As 参考sd, X.均值 As 预设均值, X.Sd As 预设sd, K.项目qc码, K.方法qc码 ,I.结果类型, K.质控取值, K.方法,'' as 序列值" & vbNewLine & _
            "From (Select I.ID, I.编码, I.中文名, I.英文名, I.单位, Decode(P.结果类型, 1,'', P.取值序列) As 取值序列, P.结果类型" & vbNewLine & _
            "       From 检验仪器项目 L, 诊治所见项目 I, 检验项目 P" & vbNewLine & _
            "       Where L.项目id = I.ID And I.ID = P.诊治项目id And L.仪器id = [2] And ( P.结果类型=1 Or P.取值序列 is Not null)) I," & vbNewLine & _
            "     (Select 质控品id, 项目id, 靶值, Sd, 项目qc码, 方法qc码, 方法, 质控取值 From 检验质控品项目 Where 质控品id = [1]) K," & vbNewLine & _
            "     (Select X.质控品id, X.项目id, X.均值, X.Sd" & vbNewLine & _
            "       From 检验质控品 R, 检验质控均值 X" & vbNewLine & _
            "       Where R.ID = X.质控品id(+) And R.开始日期 = X.开始日期(+) And R.ID = [1]) X" & vbNewLine & _
            "Where I.ID = K.项目id(+) And K.质控品id = X.质控品id(+) And K.项目id = X.项目id(+)" & vbNewLine & _
            "Order By I.编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID, lngDevId)
    Set Me.vfgList.DataSource = rsTemp: Call setListFormat(True)
    Call Refresh对照(lngResID)
    Me.Tag = IIf(blnAdd, "增加", "修改"): Call Form_Resize
    Me.txt批号.SetFocus
    Call Show参考(vfgList.Row)
    zlEditStart = True: Exit Function

errHand:
    
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngResId, mlngDevId)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long
    Dim strLists As String, strItems As String
    Dim strValList As String, strCurValue As String, lngValCount As Long
    Dim dblValue As Double
         
    If ActiveControl = txt序列值 And mblnEditRow = True Then
        If Chk序列值(txt序列值.Text) Then
            With vfgList
                If .TextMatrix(.Row, mCol.序列值) <> txt序列值 Then
                    .TextMatrix(.Row, mCol.序列值) = txt序列值
                End If
                mblnEditRow = False
            End With
        Else
            Exit Function
        End If
    End If
    strLists = ""
    With Me.vfgList
        For lngCount = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mCol.选择) = flexChecked Then
                dblValue = Val(.TextMatrix(lngCount, mCol.参考靶值))
                If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "行“参考靶值”错误(太大或精度太高）！", vbInformation, gstrSysName
                    .SetFocus: zlEditSave = 0: Exit Function
                End If
                dblValue = Val(.TextMatrix(lngCount, mCol.参考SD))
                If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                    MsgBox "第" & lngCount & "行“参考SD”错误(太大或精度太高）！", vbInformation, gstrSysName
                    .SetFocus: zlEditSave = 0: Exit Function
                End If
                If Me.chk非定值.Value = vbUnchecked Then
                    dblValue = Val(.TextMatrix(lngCount, mCol.预设均值))
                    If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                        MsgBox "第" & lngCount & "行“预设均值”错误(太大或精度太高）！", vbInformation, gstrSysName
                        .SetFocus: zlEditSave = 0: Exit Function
                    End If
                    dblValue = Val(.TextMatrix(lngCount, mCol.预设SD))
                    If dblValue > 999999 Or Val(dblValue * 10000) - Int(Val(dblValue * 10000)) > 0 Then
                        MsgBox "第" & lngCount & "行“预设SD”错误(太大或精度太高）！", vbInformation, gstrSysName
                        .SetFocus: zlEditSave = 0: Exit Function
                    End If
                End If
                
                strItems = .TextMatrix(lngCount, mCol.ID)
                If Trim(.TextMatrix(lngCount, mCol.取值序列)) = "" Or Trim(.TextMatrix(lngCount, mCol.质控取值)) <> "" Then
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.参考靶值))
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.参考SD))
                    If Me.chk非定值.Value = vbUnchecked Then
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.预设均值))
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.预设SD))
                    Else
                        strItems = strItems & ";;"
                    End If
                Else
                    strValList = Trim(Me.vfgList.TextMatrix(lngCount, mCol.取值序列))
                    
                    strCurValue = Trim(Me.vfgList.TextMatrix(lngCount, mCol.参考靶值))
                    strItems = strItems & ";"
                    For lngValCount = 0 To UBound(Split(strValList, ";"))
                        If strCurValue = Split(strValList, ";")(lngValCount) Then strItems = strItems & lngValCount + 1: Exit For
                    Next
                    strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.参考SD))
                    If Me.chk非定值.Value = vbUnchecked Then
                        strCurValue = Trim(Me.vfgList.TextMatrix(lngCount, mCol.预设均值))
                        strItems = strItems & ";"
                        For lngValCount = 0 To UBound(Split(strValList, ";"))
                            If strCurValue = Split(strValList, ";")(lngValCount) Then strItems = strItems & lngValCount + 1: Exit For
                        Next
                        strItems = strItems & ";" & Val(.TextMatrix(lngCount, mCol.预设SD))
                    Else
                        strItems = strItems & ";;"
                    End If
                
                End If
                strItems = strItems & ";" & Left(Trim(.TextMatrix(lngCount, mCol.项目qc码)), 8)
                strItems = strItems & ";" & Left(Trim(.TextMatrix(lngCount, mCol.方法qc码)), 8)
                strItems = strItems & ";" & Trim(.TextMatrix(lngCount, mCol.方法))
                strItems = strItems & ";" & Replace(Trim(.TextMatrix(lngCount, mCol.取值序列)), ";", "；")
                strItems = strItems & ";" & Replace(Trim(.TextMatrix(lngCount, mCol.序列值)), ";", "；")
                strItems = strItems & ";" & IIf(mlng仪器类型 = 2, Trim(.TextMatrix(lngCount, mCol.质控取值)), "")
                strLists = strLists & "|" & strItems
            End If
        Next
    End With
    If strLists <> "" Then strLists = Mid(strLists, 2)
    
    '一般特性检查
    If Trim(Me.txt批号.Text) = "" Then
        MsgBox "请输入批号！", vbInformation, gstrSysName
        Me.txt批号.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt批号.Text), vbFromUnicode)) > Me.txt批号.MaxLength Then
        MsgBox "批号超长（最多" & Me.txt批号.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt批号.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt浓度.Text), vbFromUnicode)) > Me.txt浓度.MaxLength Then
        MsgBox "浓度超长（最多" & Me.txt浓度.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt浓度.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo水平.ListIndex = -1 Then
        MsgBox "必须指定浓度的水平标记！(如无法指定，则可能是未设置仪器的质控水平数)", vbInformation, gstrSysName
        Me.cbo水平.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt标本号.Text), vbFromUnicode)) > Me.txt标本号.MaxLength Then
        MsgBox "标本号超长（最多" & Me.txt标本号.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt标本号.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '数据保存语句组织
    
    gstrSql = "'" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt批号.Text) & "'"
    gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp开始日期.Value, "yyyy-MM-dd") & "','yyyy-mm-dd')"
    gstrSql = gstrSql & ",To_Date('" & Format(Me.dtp结束日期.Value, "yyyy-MM-dd") & "','yyyy-mm-dd')"
    gstrSql = gstrSql & ",'" & Trim(Me.txt浓度.Text) & "'," & Me.cbo水平.ListIndex + 1 & ",'" & Trim(Me.cbo方法.Text) & "'," & mlngDevId
    gstrSql = gstrSql & "," & IIf(Me.chk非定值.Value = vbChecked, 1, 0) & ",'" & Trim(Me.txt标本号.Text) & "'"
    
    If Me.Tag = "增加" Then
        lngNewId = zldatabase.GetNextId("检验质控品")
        gstrSql = "Zl_检验质控品_Edit(1," & lngNewId & "," & gstrSql & ",'" & strLists & "','" & Trim(cbo试剂.Text) & "','" & Trim(cbo校准物.Text) & "')"
    Else
        gstrSql = "Zl_检验质控品_Edit(2," & mlngResId & "," & gstrSql & ",'" & strLists & "','" & Trim(cbo试剂.Text) & "','" & Trim(cbo校准物.Text) & "')"
    End If
    
    Err = 0: On Error GoTo errHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngResId = lngNewId
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngResId: Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

Private Sub Refresh对照(ByVal lngResID As Long)
    '将定性项目的定量对照，填到vfgList的对应列
    Dim rsQua As New ADODB.Recordset '存序列值
    Dim lngItemID As Long
    Dim iRow As Integer
    '定性项目和定量项目，取出序列值
    On Error GoTo errHand
    With Me.vfgList
        If Me.vfgList.Rows > 2 Then
            For iRow = 2 To Me.vfgList.Rows - 1
                If Val(vfgList.TextMatrix(iRow, mCol.结果类型)) <> 1 Then
                    lngItemID = Val(vfgList.TextMatrix(iRow, mCol.ID))
                    gstrSql = "Select 取值序列,序列值 From 检验质控品项目 Where 质控品ID=[1] And 项目ID=[2]"
                    Set rsQua = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngResID, lngItemID)
                    Do Until rsQua.EOF
                        If "" & rsQua!取值序列 <> "" Then vfgList.TextMatrix(iRow, mCol.取值序列) = "" & rsQua!取值序列
                        vfgList.TextMatrix(iRow, mCol.序列值) = "" & rsQua!序列值
                        rsQua.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Show参考(ByVal intRow As Integer)

    mblnEditRow = False
    lbl取值序列.Visible = False: txt取值序列.Visible = False
    lbl序列值.Visible = False: txt序列值.Visible = False
    txt取值序列.Text = "": txt序列值 = ""
    With vfgList
        If intRow > 1 And intRow < .Rows And .Cols >= 15 Then
            
            If Val(.TextMatrix(intRow, mCol.结果类型)) <= 1 Then
                txt取值序列.Text = "": txt序列值 = ""
            Else
                
                txt取值序列.Text = .TextMatrix(intRow, mCol.取值序列)
                txt序列值 = .TextMatrix(intRow, mCol.序列值)
                lbl取值序列.Visible = True: txt取值序列.Visible = True
                lbl序列值.Visible = True: txt序列值.Visible = True
            End If
        End If
    End With
End Sub


'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cbo方法_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo水平_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo水平_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkHide_Click()
    If Me.chkHide.Value = vbChecked Then
        Me.vfgList.ColWidth(mCol.中文名) = 0: Me.vfgList.ColHidden(mCol.中文名) = True
    Else
        Me.vfgList.ColWidth(mCol.中文名) = 2000: Me.vfgList.ColHidden(mCol.中文名) = False
    End If
End Sub

Private Sub chk非定值_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk非定值_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub dtp结束日期_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub dtp开始日期_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mlngResId = 0: mlngDevId = 0
    If Val(zldatabase.GetPara("隐藏中文名", glngSys, 1062, 1)) = 0 Then
        Me.chkHide.Value = vbUnchecked
    Else
        Me.chkHide.Value = vbChecked
    End If
    
    Err = 0: On Error GoTo errHand
    '字段长度限制
    gstrSql = "Select 名称, 批号, 浓度, 方法, 标本号 From 检验质控品 Where ID = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngResId)
    With rsTemp
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt批号.MaxLength = .Fields("批号").DefinedSize
        Me.txt浓度.MaxLength = .Fields("浓度").DefinedSize
        Me.txt标本号.MaxLength = .Fields("标本号").DefinedSize
    End With
    
    '质控检验方法
    mstr方法 = "|"
    gstrSql = "Select 名称 From 质控检验方法 Order By 编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo方法.Clear
        Do While Not .EOF
            Me.cbo方法.AddItem "" & Trim(!名称)
            mstr方法 = mstr方法 & "|" & Trim(!名称)
            .MoveNext
        Loop
        If Me.cbo方法.ListCount > 0 Then Me.cbo方法.ListIndex = 0
    End With
    
    Call setListFormat
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picEdit.Height = Me.ScaleHeight
    Me.chkHide.Left = Me.ScaleWidth - Me.chkHide.Width - 90
    With Me.vfgList
        .Width = Me.ScaleWidth - .Left - 90
        .Height = Me.ScaleHeight - .Top - 90
    End With
    If Me.Tag <> "" Then
        Me.picEdit.Enabled = True: Me.picEdit.BackColor = RGB(250, 250, 250)
        Me.vfgList.Editable = flexEDKbd: Me.vfgList.FocusRect = flexFocusHeavy
    Else
        Me.picEdit.Enabled = False: Me.picEdit.BackColor = Me.BackColor
        Me.vfgList.Editable = flexEDNone: Me.vfgList.FocusRect = flexFocusNone
    End If
    Me.chk非定值.BackColor = Me.picEdit.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.chkHide.Value = vbChecked Then
        Call zldatabase.SetPara("隐藏中文名", 1, glngSys, 1062)
    Else
        Call zldatabase.SetPara("隐藏中文名", 0, glngSys, 1062)
    End If
End Sub

Private Sub txt标本号_GotFocus()
    Me.txt标本号.SelStart = 0: Me.txt标本号.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt标本号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        If InStr(1, ",-", Chr(KeyAscii)) > 0 Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt序列值_KeyPress(KeyAscii As Integer)
    mblnEditRow = True
End Sub

Private Function Chk序列值(ByVal str序列值 As String) As Boolean

    Dim var值 As Variant
    Dim var序列 As Variant
    Dim i As Integer
    
    If str序列值 <> "" Then
        var序列 = Split(str序列值, ";")
        var值 = Split(str序列值, ";")
    
        If UBound(var序列) <> UBound(var值) Then
            MsgBox "序列值的项目格式和取值序列的格式不一致，请重新设置！", vbQuestion, gstrSysName
            Exit Function
        End If
        For i = LBound(var值) To UBound(var值)
            If Not IsNumeric(var值(i)) Then
                MsgBox "序列值中，分号中间的内容应为数字，请重新设置！", vbQuestion, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    Chk序列值 = True
End Function
Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt浓度_GotFocus()
    Me.txt浓度.SelStart = 0: Me.txt浓度.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt浓度_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt批号_GotFocus()
    Me.txt批号.SelStart = 0: Me.txt批号.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt批号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt序列值_Validate(Cancel As Boolean)
    If Not Chk序列值(txt序列值.Text) Then
        Cancel = True
    Else
        With vfgList
            If .TextMatrix(.Row, mCol.序列值) <> txt序列值 Then
                .TextMatrix(.Row, mCol.序列值) = txt序列值
            End If
            mblnEditRow = False
        End With
    End If
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strLists As String, strValue As String
    
    If Col <> mCol.参考靶值 And Col <> mCol.预设均值 Then Exit Sub
    If Trim(Me.vfgList.TextMatrix(Row, Col)) = "" Then Exit Sub
      
    strLists = Trim(Me.vfgList.TextMatrix(Row, mCol.取值序列))
    strValue = Trim(Me.vfgList.TextMatrix(Row, Col))
    If Trim(Me.vfgList.TextMatrix(Row, mCol.质控取值)) <> "" Then strLists = ""
    
    If strLists = "" Then Exit Sub
    For lngCount = 0 To UBound(Split(strLists, ";"))
        If strValue = Split(strLists, ";")(lngCount) Then Exit Sub
    Next
    Me.vfgList.TextMatrix(Row, Col) = ""
    
    strValue = "该项目为半定量项目，" & IIf(Col = mCol.参考靶值, "靶值", "均值") & "设置需符合取值序列(" & strLists & ")要求！"
    MsgBox strValue, vbInformation, gstrSysName
End Sub

Private Sub vfgList_DblClick()
    If Me.vfgList.MouseRow < Me.vfgList.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    With Me.vfgList
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked Then
            .Cell(flexcpChecked, .Row, mCol.选择) = flexUnchecked
        Else
            .Cell(flexcpChecked, .Row, mCol.选择) = flexChecked
        End If
    End With
End Sub

Private Sub vfgList_EnterCell()
     vfgList.Select vfgList.Row, vfgList.Col
End Sub

Private Sub vfgList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Or Me.vfgList.Col > mCol.英文名 Then Exit Sub
    Call vfgList_DblClick
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col < mCol.参考靶值 Then Exit Sub
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22, vbKeyReturn: Exit Sub
    Case Else
        Select Case Col
        Case mCol.参考靶值, mCol.预设均值
            With Me.vfgList
                If Trim(.TextMatrix(.Row, mCol.取值序列)) <> "" Then Exit Sub
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
            End With
        Case mCol.参考SD, mCol.预设SD
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
        Case mCol.项目qc码, mCol.方法qc码
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
'        Case mCol.方法
'            Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub vfgList_LeaveCell()
    'Call vfgList.Select(vfgList.Row, vfgList.Col)
End Sub

Private Sub vfgList_SelChange()
    If mblnEditRow Then
        
    End If
    Call Show参考(vfgList.Row)
End Sub

Private Sub vfgList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < mCol.参考靶值 Then Cancel = True: Exit Sub
    If Row < Me.vfgList.FixedRows Then Cancel = True: Exit Sub
    If Me.chk非定值.Value = vbChecked And (Col = mCol.预设均值 Or Col = mCol.预设SD) Then Cancel = True: Exit Sub
End Sub
