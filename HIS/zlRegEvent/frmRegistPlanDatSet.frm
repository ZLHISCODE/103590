VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanDatSet 
   AutoRedraw      =   -1  'True
   Caption         =   "挂号安排时段设置"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   Icon            =   "frmregistplandatset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7350
   ScaleWidth      =   10410
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra应用于 
      Caption         =   "应用于(&B)"
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   6650
      Width           =   7095
      Begin VB.OptionButton opt所有 
         Caption         =   "应用于所有"
         Height          =   255
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "应用与本科室"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt本医生 
         Caption         =   "应用于本医生"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8640
      TabIndex        =   25
      Top             =   6840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7320
      TabIndex        =   24
      Top             =   6840
      Width           =   1100
   End
   Begin VB.Frame fraInfo 
      Caption         =   "基本信息"
      Height          =   1380
      Left            =   120
      TabIndex        =   6
      Top             =   75
      Width           =   10095
      Begin VB.ComboBox cbo号类 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   26
         Text            =   "cbo号类"
         Top             =   307
         Width           =   1155
      End
      Begin VB.TextBox txt限约 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   14
         Top             =   307
         Width           =   1215
      End
      Begin VB.TextBox txt限号 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4965
         MaxLength       =   5
         TabIndex        =   13
         Top             =   307
         Width           =   975
      End
      Begin VB.CheckBox chk序号控制 
         Caption         =   "序号控制"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   330
         Width           =   1095
      End
      Begin VB.CheckBox chk病案 
         Caption         =   "挂号时必须建病案"
         Enabled         =   0   'False
         Height          =   195
         Left            =   8040
         TabIndex        =   5
         Top             =   360
         Width           =   1845
      End
      Begin VB.ComboBox cbo科室 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         TabIndex        =   2
         Text            =   "cbo科室"
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboDoctor 
         Enabled         =   0   'False
         Height          =   300
         Left            =   6720
         TabIndex        =   4
         Top             =   705
         Width           =   2115
      End
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3360
         TabIndex        =   3
         Text            =   "cboItem"
         Top             =   705
         Width           =   2235
      End
      Begin VB.TextBox txt号别 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   5
         TabIndex        =   0
         Top             =   307
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "限约"
         Height          =   180
         Left            =   6240
         TabIndex        =   16
         Top             =   367
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "限号"
         Height          =   180
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "号类"
         Height          =   180
         Left            =   3000
         TabIndex        =   12
         Top             =   367
         Width           =   360
      End
      Begin VB.Label lbl医生 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "院内医生"
         Height          =   180
         Left            =   5940
         TabIndex        =   10
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "项目"
         Height          =   180
         Left            =   3000
         TabIndex        =   9
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   765
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   367
         Width           =   390
      End
   End
   Begin VB.Frame fraDate 
      Caption         =   "时段设置"
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   10215
      Begin VB.PictureBox picTime 
         BorderStyle     =   0  'None
         Height          =   4665
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   9945
         TabIndex        =   17
         Top             =   240
         Width           =   9945
         Begin VB.CommandButton cmdOtherCalc 
            Caption         =   "其他辅助计划(&R)"
            Height          =   345
            Left            =   3735
            TabIndex        =   32
            Top             =   30
            Width           =   1860
         End
         Begin VB.CommandButton cmd设置时段 
            Caption         =   "辅助计算(&F)"
            Height          =   350
            Left            =   2520
            TabIndex        =   31
            ToolTipText     =   "点击重新计算时段"
            Top             =   35
            Width           =   1150
         End
         Begin VB.TextBox txtTimeOut 
            Height          =   300
            Left            =   1560
            TabIndex        =   29
            Text            =   "10"
            Top             =   60
            Width           =   500
         End
         Begin MSComCtl2.UpDown udTime 
            Height          =   345
            Left            =   2160
            TabIndex        =   28
            Top             =   38
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComctlLib.TabStrip tbWeekTime 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   1
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTime 
            Height          =   3825
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   9765
            _cx             =   17224
            _cy             =   6747
            Appearance      =   1
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
            GridColor       =   12632256
            GridColorFixed  =   0
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   2
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmregistplandatset.frx":000C
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
            Editable        =   2
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
            Begin VB.CommandButton cmd删除 
               Caption         =   "删"
               Height          =   255
               Left            =   7305
               TabIndex        =   33
               Top             =   2025
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton cmd预约 
               Caption         =   "预"
               Height          =   255
               Left            =   7320
               TabIndex        =   27
               Top             =   1560
               Visible         =   0   'False
               Width           =   375
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "时间间隔(分)"
            Height          =   180
            Left            =   360
            TabIndex        =   30
            Top             =   120
            Width           =   1080
         End
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "院内医生"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "含外援医生"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanDatSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit '要求变量声明
'
Public Enum ViewMode
     NewItem
     ViewItem '查看
     Edit '编辑
End Enum
Private mViewMode         As ViewMode    '页面显示模式
Private mlng安排ID        As Long        '安排ID
Private mlngPre安排ID     As Long
Private mrsTime          As ADODB.Recordset
Private mrs限号          As ADODB.Recordset
Private mrs上班时间段     As ADODB.Recordset
Private mrs安排          As ADODB.Recordset
Private mblnCellChange   As Boolean
Private mstrKey         As String
Private mblnChange      As Boolean
Private mblnReload      As Boolean '在挂号安排管理页面调用 ShowMe以后 是否需要刷新
Private mstr限制修改 As String '在某一天或者多天的安排限制更改
'对外上班时间
Private Type t_上班时间
  dat_上午上班 As Date
  dat_上午下班 As Date
  dat_下午上班 As Date
  dat_下午下班 As Date
End Type
Private t_时间 As t_上班时间
Private Const strMaskKey As String = "09:00-09:00"
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1

Private Sub chk序号控制_Click()
    cmdOtherCalc.Visible = chk序号控制.Value = 1
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
    cmdOK.Enabled = False
    zlCommFun.ShowFlash "正在保存挂号安排时段数据,请稍候……"
    If SaveDate() = True Then
        '************************
        '如果保存成功需要重新对
        '挂号安排时段进行提取
        '************************
        Call InitData
        mblnChange = False
        mblnReload = True
        'If tbWeekTime.Tabs.Count > 0 Then tbWeekTime.Tabs(1).Selected = True
    End If
    zlCommFun.StopFlash
   cmdOK.Enabled = True
End Sub


Private Sub cmdOtherCalc_Click()
    Dim str安排 As String
    
    If chk序号控制.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    str安排 = Replace(Split(tbWeekTime.SelectedItem.Caption & "(", "(")(1), ")", "")
    Call mfrmOtherCalc.zlShowMe(Me, str安排, Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing
End Sub

Private Sub cmd删除_Click()
    Call DeleteSelectPain
End Sub

Private Sub cmd设置时段_Click()
'对挂号安排时段进行设置
    Dim str星期         As String
    
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrsTime.Filter = "星期='" & str星期 & "'"
    If mrsTime.RecordCount > 0 Then
      '****************************************************************
      '在已有挂号安排时段的情况下
      '提示操作员 是否需要重新计算时段
      '****************************************************************
        If MsgBox("此安排在" & str星期 & "已经存在时段 " & vbCrLf & "是否重新计算时段?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
            mrsTime.Filter = 0
            Exit Sub
        End If
    End If
    Select Case chk序号控制.Value = 1
    Case True:
        Set专家号时段
        setVsFlexBgColor (True)
    Case False:
        Set普通号时段
        setVsFlexBgColor
    End Select
    
    mblnChange = True
End Sub
Private Sub Set普通号时段()
    Dim strSQL      As String
    Dim str星期     As String
    Dim str时段     As String
    Dim lng限号     As Long
    Dim lng限约     As Long
    Dim lng间隔     As Long
    Dim dblDatCount As Long '总时间间隔
    Dim dat时点     As Date '每个时间段的
    Dim bln全天     As Boolean  '是否是全天都允许挂号 如果是全天则分为上午和下午
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs上班时间段 Is Nothing Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs限号.Filter = "星期='" & str星期 & "'"
    If mrs限号.RecordCount = 0 Then
        MsgBox "当前号别在" & str星期 & ",没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
        Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
    End If
    lng限号 = Nvl(mrs限号!限号数, 0): lng限约 = Nvl(mrs限号!限约数, 0)
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str星期 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt限号.Text = lng限号
    Me.txt限约.Text = lng限约
    If lng限约 = 0 Then lng限约 = lng限号 '如果对预约没有限制则认为最大限约数和限号数相同
    str时段 = Nvl(mrs安排(str星期).Value)
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
     
    '*********************************
    '分时段具体处理 分为全天和非全天
    '全天分为上午和下午
    '*********************************
  
    lng间隔 = Val(txtTimeOut.Text)
   
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
    End With
    '*************************************
    '普通号
    '*************************************
    With vsTime
        .Cols = 8: .FixedCols = 0
        .Rows = 1: .FixedRows = 1
        For i = 0 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "时间段"
        Next
        For i = 1 To .Cols - 1 Step 2
           .TextMatrix(0, i) = "预约人数"
        Next
        lngRow = 1: lngCol = -1
        j = 1: lngStart = 1
        Do While Not mrs上班时间段.EOF
            If blnExit Then Exit Do
            dat时点 = CDate(Nvl(mrs上班时间段!上班, "00:00:00"))
            For i = j To lng限号
                If lngStart > lng限号 Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                
                lngCol = lngCol + 1
                If lngCol * 2 > .Cols - 2 Then lngRow = lngRow + 1: lngCol = 0
                strData = IIf(lng限约 >= i, 1, 0)
                strTime = Format(dat时点, "HH:mm") & "-" & _
                      IIf(Format(DateAdd("n", lng间隔, dat时点), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                      Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
               
                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(lngRow, lngCol * 2) = strTime
                .TextMatrix(lngRow, lngCol * 2 + 1) = strData
                lngStart = lngStart + 1
                dat时点 = DateAdd("n", lng间隔, dat时点)
            Next
            mrs上班时间段.MoveNext
        Loop
       
 
         For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
Private Sub Set专家号时段()
    Dim strSQL      As String
    Dim str星期     As String
    Dim str时段     As String
    Dim lng限号     As Long
    Dim lng限约     As Long
    Dim lng间隔     As Long
    Dim dblDatCount As Long '总时间间隔
    Dim dat时点     As Date '每个时间段的
    Dim str时点     As String
    Dim bln全天     As Boolean  '是否是全天都允许挂号 如果是全天则分为上午和下午
    Dim datStart    As Date
    Dim datEnd      As Date
    Dim i           As Long
    Dim j           As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim strData     As String
    Dim strTime     As String
    Dim strList()   As String
    Dim blnExit     As Boolean
    Dim lngIndex    As Long
    Dim lngStart    As Long
    On Error GoTo Hd
    If mrs上班时间段 Is Nothing Then Exit Sub
    If mrs限号 Is Nothing Then
        strSQL = _
        "Select 安排id, 限制项目 as 星期 , 限号数, 限约数 From 挂号安排限制 Where 安排id = [1]"
        Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt号别.Tag))
        If mrsTime.RecordCount = 0 Then
            MsgBox "当前号别没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
            Set mrs限号 = Nothing
            Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
        End If
    End If
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    mrs限号.Filter = "星期='" & str星期 & "'"
    If mrs限号.RecordCount = 0 Then
        MsgBox "当前号别在" & str星期 & ",没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
        Exit Sub '如果挂号安排中没有设置此天的信息 就不允许设置
    End If
    lng限号 = Nvl(mrs限号!限号数, 0): lng限约 = Nvl(mrs限号!限约数, 0)
    If lng限号 = 0 Then
        MsgBox "当前号别在" & str星期 & ",没有对挂号数进行限制,无法设置时段,请检查!", vbOKOnly, Me.Caption
        Exit Sub
    End If
    Me.txt限号.Text = lng限号
    Me.txt限约.Text = lng限约
    lng限约 = lng限号
    str时段 = Nvl(mrs安排(str星期).Value)
    mrs上班时间段.Filter = "时间段='" & str时段 & "'"
    
'*************************************************************
'时间间隔根据 设置的间隔
'*************************************************************
      lng间隔 = Val(Me.txtTimeOut.Text)
     ' dat时点 = CDate(Nvl(mrs上班时间段!开始时间, "00:00:00"))
        
      With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
      End With
    '*************************************
    '专家号
    '序号填充规则
    '根据 时间段表中的 上下班时间来判断
    '其中 全天这种情况  分为上午和下午
    '*************************************
    
    With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         j = 1
         lngStart = 1
         Do While Not mrs上班时间段.EOF
            If blnExit Then Exit Do
             
            dat时点 = CDate(Nvl(mrs上班时间段!上班, "00:00:00"))
             For i = j To lng限约
                If lngStart > lng限约 Then
                    blnExit = True
                    Exit For
                End If
              
                If Format(dat时点, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                 End If
                lngCol = lngCol + 1
                If str时点 <> Format(dat时点, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
                If lngCol = 1 Then
                     If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                     str时点 = Format(dat时点, "HH") & ":00"
                     vsTime.TextMatrix(lngRow - 1, 0) = str时点
                     vsTime.TextMatrix(lngRow, 0) = str时点
                
                End If
                strData = lngStart
                lngStart = lngStart + 1
                strTime = Format(dat时点, "HH:mm") & "-" & _
                           IIf(Format(DateAdd("n", lng间隔, dat时点), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
                           Format(CDate(Nvl(mrs上班时间段!下班, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
    
                If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
                vsTime.TextMatrix(lngRow - 1, lngCol) = strData
                vsTime.TextMatrix(lngRow, lngCol) = strTime
                '是第一项时 填写 开始时间到首行
                
                dat时点 = DateAdd("n", lng间隔, dat时点)
             Next
             mrs上班时间段.MoveNext
         Loop
'         '***********************
'         '序号填充
'         '**********************
'         For i = 1 To lng限约
'            If Format(dat时点, "dd:mm:ss") >= Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "dd:mm:ss") Then Exit For
'            lngCol = lngCol + 1
'            If str时点 <> Format(dat时点, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
'            If lngCol = 1 Then
'                 If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                 str时点 = Format(dat时点, "HH") & ":00"
'                 vsTime.TextMatrix(lngRow - 1, 0) = str时点
'                 vsTime.TextMatrix(lngRow, 0) = str时点
'
'            End If
'            strData = i
'            strTime = Format(dat时点, "HH:mm") & "-" & _
'                       IIf(DateAdd("n", lng间隔, dat时点) > CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), _
'                       Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
'
'            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'            '是第一项时 填写 开始时间到首行
'
'            dat时点 = DateAdd("n", lng间隔, dat时点)
'         Next
'         If bln全天 Then
'             mrs上班时间段.Filter = "时间段='下午'"
'            dat时点 = CDate(Nvl(mrs上班时间段!开始时间, "00:00:00"))
'         End If
'         j = i
'         For i = j To lng限约
'            If Format(dat时点, "dd:mm:ss") >= CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")) Then Exit For
'            lngCol = lngCol + 1
'            If lngCol > vsTime.Cols - 1 Then lngRow = lngRow + 2: lngCol = 1
'            strData = i
'            strTime = Format(dat时点, "HH:mm") & "-" & _
'                       IIf(DateAdd("n", lng间隔, dat时点) > CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), _
'                       Format(CDate(Nvl(mrs上班时间段!终止时间, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng间隔, dat时点), "HH:mm"))
'            If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'            If lngRow < 0 Then vsTime.Rows = vsTime.Rows + 2: lngRow = lngRow + 2
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'
'            '是第一项时 填写 开始时间到首行
'            If lngCol = 1 Then
'                 vsTime.TextMatrix(lngRow - 1, 0) = Format(dat时点, "HH:mm")
'                 vsTime.TextMatrix(lngRow, 0) = Format(dat时点, "HH:mm")
'            End If
'            dat时点 = DateAdd("n", lng间隔, dat时点)
'         Next
         For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
         Next
         .ColWidth(0) = 1200
         .FixedAlignment(0) = flexAlignRightTop
         .ColAlignment(0) = flexAlignRightTop
         If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
         End If
         .Redraw = flexRDBuffered
    End With
     
Exit Sub
Hd:
    If ErrCenter() = 1 Then
         Resume
    End If
    SaveErrLog
End Sub

Private Sub cmd预约_Click()
    Dim intFirstRow As Integer '问题号:57488
    Dim intSecondRow As Integer '问题号:57488
    '问题号:57488
    With vsTime
        intFirstRow = .Row: intSecondRow = intFirstRow + 1
        If .Row Mod 2 = 1 Then
            intFirstRow = .Row - 1
            intSecondRow = intFirstRow + 1
        End If
    End With
    
    '对时间段能否预约进行设置
    If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Then Exit Sub
    If mViewMode = ViewMode.ViewItem Or vsTime.TextMatrix(vsTime.MouseRow, vsTime.MouseCol) = "" Then Exit Sub
    
    With vsTime
        If .CellForeColor = vbBlue Then
            .Cell(flexcpForeColor, intFirstRow, .Col, intSecondRow, .Col) = &H80000008
            .Cell(flexcpFontBold, intFirstRow, .Col, intSecondRow, .Col) = False
         Else
            .Cell(flexcpForeColor, intFirstRow, .Col, intSecondRow, .Col) = vbBlue
            .Cell(flexcpFontBold, intFirstRow, .Col, intSecondRow, .Col) = True
        End If
    End With
    mblnChange = True
End Sub

Private Sub Form_Activate()
    Me.Icon = frmRegistPlan.Icon
End Sub

Private Sub Form_Load()
    Init时间段
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  '********************************************
  '首先设置 窗体的最小宽度和最小高度
  '********************************************
  If Me.Width < 701 * Screen.TwipsPerPixelX Then Me.Width = 701 * Screen.TwipsPerPixelX
  If Me.Height < 511 * Screen.TwipsPerPixelY Then Me.Height = 511 * Screen.TwipsPerPixelY
  '********************************************
  '挂号安排基本信息 位置不移动移动
  '仅移动 时段设置
  '********************************************
  With fraDate
     .Width = Me.ScaleWidth - 2 * .Left
     .Height = Me.ScaleHeight - Me.fraInfo.Top - Me.fraInfo.Height - 65 * Screen.TwipsPerPixelY
  End With
  
  With picTime
     .Width = fraDate.Width - 2 * .Left
     .Height = fraDate.Height - .Top * 2
  End With
  With Me.tbWeekTime
    .Width = picTime.ScaleWidth - 2 * .Left
  End With
  With Me.vsTime
    .Width = picTime.ScaleWidth - 2 * .Left
    .Height = picTime.ScaleHeight - .Top - cmd设置时段.Top
  End With
  '-------------------------------------------
  '应用于 位置的调整
  '-------------------------------------------
  With Me.fra应用于
       .Left = .Left
       .Top = Me.fraDate.Top + Me.fraDate.Height + 5 * Screen.TwipsPerPixelY
   
  End With
  
  '********************************************
  '确定按钮和取消按钮的移动
  '********************************************
  
  With Me.cmdCancel
       .Left = Me.ScaleWidth - 40 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
  With Me.cmdOK
       .Left = cmdCancel.Left - 20 * Screen.TwipsPerPixelX - .Width
       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
     mlngPre安排ID = -1
     mblnChange = False
     Set mrsTime = Nothing
     mstr限制修改 = ""
     Set mrs限号 = Nothing
     Set mrs上班时间段 = Nothing
     Set mrs安排 = Nothing
End Sub
Private Sub mfrmOtherCalc_zlRefreshCon(ByVal VarTimes As Variant)
    Dim strTemp  As String, varData As Variant, varTemp As Variant
    Dim i As Long, int分钟 As Integer, dtStart As Date, dtEnd As Date
    Dim lngRow As Long, lng序号 As Long, dtTemp As Date, j As Long
    Dim lng限号数 As Long, lng限约数 As Long, str星期 As String
    Dim lng已挂最大序号 As Long '问题号:51427
    Dim lngCol As Long '问题号:54127
    Dim lng安排ID As Long '问题号:54127
    Dim K As Long '问题好:54127
    
    If chk序号控制.Value <> 1 Then Exit Sub
    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng安排ID = Val(txt号别.Tag) '问题号:51427
    
    If Get限号数(str星期, lng限号数, lng限约数) = False Then Exit Sub
    
    'VarTiems
    '       "时间间隔"
    '       "分段间隔":时间(如:8:00～9:00),2;时间2,间隔;....
    If VarTimes("时间间隔") <> "" Then
        txtTimeOut.Text = Val(VarTimes("时间间隔"))
        Call cmd设置时段_Click
        Exit Sub
    End If
    strTemp = VarTimes("分段间隔")
    If strTemp = "" Then Exit Sub
    
    '问题号:51427
    lng已挂最大序号 = ExistsBooking(lng安排ID, str星期)
    If lng已挂最大序号 <> -1 Then
         If MsgBox("该安排下已有被挂出去的号,只能修改黑色字体显示的时段" & vbCrLf & "您确定要继续修改吗?", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
         End If
    End If
    
    varData = Split(strTemp, ";")
    lngRow = -2: lng序号 = 1: lngCol = 1
    '问题号:51427
    For i = 0 To vsTime.Rows - 1
        For j = 0 To vsTime.Cols - 1
            If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                If CLng(vsTime.TextMatrix(i, j)) = lng已挂最大序号 Then
                    lngRow = i: lngCol = j
                End If
            End If
        Next
    Next
    
    '初始化vsTime控件
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = IIf(lngRow = -2, 2, lngRow + 2): lngRow = IIf(lngRow = 0 And lngCol = 1, -2, lngRow): i = 0: .FixedCols = 1
        .FixedRows = 0
        If lngRow = -2 Then
            .Rows = 0
            .Rows = 2
        End If
    lng序号 = IIf(lng已挂最大序号 = -1, 1, lng已挂最大序号 + 1)
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int分钟 = Val(varTemp(1))
        varTemp = Split(varTemp(0), "～")
        dtStart = CDate(varTemp(0))
        dtEnd = CDate(varTemp(1))
        
        '同一时间段中没有挂出的号序
        If dtStart = IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            j = IIf(lngCol = 1 And lngRow = -2, 0, lngCol) + 1
            '清空没挂出的选项
            For K = j To .Cols - 1
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), K) = ""
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, K) = ""
            Next
            .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = Format(dtStart, "HH:00")
            .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0) = Format(dtStart, "HH:00")
            If lngCol = 1 Then
                dtStart = .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, 0)
            Else
                dtStart = Split(.TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, lngCol), "-")(1)
            End If
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int分钟 * 1 / 24 / 60, "HH:MM")
                If dtTemp > dtEnd Or lng序号 > lng限号数 Then Exit Do
                .TextMatrix(IIf(lngRow = -2, 0, lngRow), j) = lng序号
                .TextMatrix(IIf(lngRow = -2, 0, lngRow) + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng序号 = lng序号 + 1
                j = j + 1
            Loop
        dtStart = "00:00:00"
        End If
        '不同时间段没有被挂出的号序
        If dtStart > IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) Then
            If IIf(.TextMatrix(IIf(lngRow = -2, 0, lngRow), 0) = "", "00:00:00", .TextMatrix(IIf(lngRow = -2, 0, lngRow), 0)) <> Format(dtStart, "HH:00") Then
                 If lng序号 > 1 Then
                     lngRow = IIf(lngRow = -2, 0, lngRow)
                 End If
                 lngRow = lngRow + 2
                .Rows = .Rows + 2
                .TextMatrix(lngRow, 0) = Format(dtStart, "HH:00")
                .TextMatrix(lngRow + 1, 0) = Format(dtStart, "HH:00")
            End If
            j = 1
            Do While True
                If j > .Cols - 1 Then .Cols = .Cols + 1
                dtTemp = Format(dtStart + int分钟 * 1 / 24 / 60, "HH:MM")
                If dtTemp > dtEnd Or lng序号 > lng限号数 Then Exit Do
                .TextMatrix(lngRow, j) = lng序号
                .TextMatrix(lngRow + 1, j) = Format(dtStart, "HH:MM") & "-" & Format(dtTemp, "HH:MM")
                dtStart = dtTemp: lng序号 = lng序号 + 1
                j = j + 1
            Loop
        End If
    Next
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignCenterCenter
        .ColWidth(i) = 1200
    Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
    If .Rows > 0 Then
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
    End If
    .Redraw = flexRDBuffered
    End With
    Call setVsFlexBgColor(True)
End Sub

Private Sub tbWeekTime_Click()
    Dim i       As Integer
    Dim j As Long '问题号:51427
    
    If mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2) Then Exit Sub
    If mblnChange Then
        mblnChange = False
        If MsgBox("当前挂号安排在" & mstrKey & "的时段已改变!是否保存?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption) = vbYes Then
            cmdOK_Click
'         For i = 1 To tbWeekTime.Tabs.Count
'            If tbWeekTime.Tabs(i).Key = "K" & mstrKey Then
'                tbWeekTime.Tabs(i).Selected = True
'                Exit For
'            End If
'         Next
        End If
    End If
    mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2)
     If mstr限制修改 <> "" Then
        vsTime.Editable = flexEDKbdMouse: cmd设置时段.Enabled = True
        If InStr(mstr限制修改, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone: cmd设置时段.Enabled = False
    End If
    Select Case mViewMode
        Case ViewMode.ViewItem:
             Call LoadTimePlan(mlng安排ID, Me.chk序号控制.Value = 1)
        Case ViewMode.Edit:
            cmd预约.Visible = False
            cmd删除.Visible = False
            Call LoadEditTimePlan(mlng安排ID, Me.chk序号控制.Value = 1)
    End Select
     setVsFlexBgColor (Me.chk序号控制.Value = 1)
End Sub


 

Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
   
    '限制非数字输入
    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
    If txtTimeOut.Text = "" And KeyAscii = Asc(0) Then KeyAscii = 0
End Sub

Private Sub txtTimeOut_Validate(Cancel As Boolean)
    If Val(txtTimeOut.Text) < 1 Then Cancel = True
End Sub

 

Private Sub udTime_DownClick()
    If Val(txtTimeOut.Text) < 2 Then Exit Sub
    txtTimeOut.Text = Val(txtTimeOut.Text) - 1
End Sub

Private Sub udTime_UpClick()
  txtTimeOut.Text = Val(txtTimeOut.Text) + 1
End Sub


 
 
'Private Sub vsTime_Click()
'  Select Case mViewMode
'    Case ViewMode.Edit, ViewMode.NewItem:
'       If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Or (chk序号控制.Value = 0 And vsTime.MouseRow < 1) Then Exit Sub
'       Select Case chk序号控制.Value = 1
'            Case True:
'            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            Case False:
'            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'       End Select
'        If vsTime.MouseRow < 0 Or vsTime.MouseCol < 1 Then Exit Sub
'
'        If chk序号控制.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
'            cmd预约.Left = vsTime.MouseCol * 1200 + 20
'            cmd预约.Top = vsTime.MouseRow * 400 + 20
'            cmd预约.Visible = True
'        End If
'
'    Case ViewMode.ViewItem:
'         vsTime.Editable = flexEDNone
'  End Select
'End Sub

Public Function ShowMe(lng安排ID As Long, mode As ViewMode) As Boolean
    mViewMode = mode: mlng安排ID = lng安排ID
    If InitData() = False Then
        '加载挂号安排基本信息
         Exit Function
    End If
    Select Case mViewMode
         Case ViewMode.ViewItem:
                vsTime.Editable = flexEDNone
                Me.txtTimeOut.Enabled = False
                Me.cmd设置时段.Enabled = False
               '查看
              Call LoadTimePlan(mlng安排ID, chk序号控制.Value = 1, False)
         Case ViewMode.Edit
              If LoadEditTimePlan(mlng安排ID, chk序号控制.Value = 1, False) = False Then
               Exit Function
              End If
    End Select
    setVsFlexBgColor (chk序号控制.Value = 1)
    Me.Show 1
    ShowMe = mblnReload
End Function
'------------------------------------------------------------------------
'页面调用过程与方法
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL          As String
    Dim lng安排ID       As Long
    If mlng安排ID = -1 Then Exit Function
     lng安排ID = mlng安排ID
     On Error GoTo Hd
     strSQL = " " & _
        "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
        "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,nvl(A.默认时段间隔,5) As 默认时段间隔, " & _
        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
        "   From 挂号安排 A,收费项目目录 B,部门表 D " & _
        "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
        "         And A.Id=[1]"
         Set mrs安排 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
         
         If mrs安排.EOF Then
              ShowMsgbox "未找到指定的号别,请检查!"
             Exit Function
        End If
        strSQL = "Select 限制项目,限号数,  限约数,限制项目 as 星期 From  挂号安排限制 where 安排ID=[1]  Order BY 限制项目      "
        Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
        cbo号类.Text = Nvl(mrs安排!号类)
        txt号别.Tag = Nvl(mrs安排!安排ID)
        txtTimeOut.Tag = Val(Nvl(mrs安排!默认时段间隔, 0))
        txtTimeOut.Text = txtTimeOut.Tag
        txt号别.Text = Nvl(mrs安排!号码)
        cbo科室.Text = Nvl(mrs安排!科室)
        cboItem.Text = Nvl(mrs安排!项目)
        cboDoctor.Text = Nvl(mrs安排!医生姓名)
        chk病案.Value = IIf(Val(Nvl(mrs安排!病案必须)) = 1, 1, 0)
        chk序号控制.Value = IIf(Val(Nvl(mrs安排!序号控制)) = 1, 1, -0):  chk序号控制.Tag = chk序号控制.Value
        '问题号:51429
        Call chk序号控制_Click
        strSQL = "" & _
        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
        "               限制数量,是否预约" & _
        "   From  挂号安排时段 " & _
        "   Where 安排ID=[1]" & _
        "   Order by 排序,时点,序号"
        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
        mstr限制修改 = Get已约限制(mlng安排ID)
       InitData = True
Exit Function
Hd:
     If ErrCenter() = 1 Then Resume
     SaveErrLog
End Function

 
Private Function LoadEditTimePlan(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
    Optional bln计划 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str星期          As String
    Dim i                As Long
    Dim j                As Long
    Dim r                As Integer
    Dim lngRow           As Long
    Dim lngCol           As Integer
    Dim str时点          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
    Dim lng已挂最大序号 As Long '问题号:51427
     
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsTime Is Nothing Then
        mlngPre安排ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre安排ID = -1
    End If
    If mlngPre安排ID <> lng安排ID Then
        mlngPre安排ID = lng安排ID
        tbWeekTime.Tabs.Clear
        With tbWeekTime
            If Not mrs限号.EOF Then
                mrs限号.Filter = "星期='周一'"
                If mrs限号.RecordCount > 0 Then
                '限号数,  限约数,限制项目
                    If Nvl(mrs限号!限号数, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K周一", "周一" & IIf(Nvl(mrs安排!周一) = "", "", "(" & Nvl(mrs安排!周一) & ")")
                    End If
                End If
                mrs限号.Filter = "星期='周二'"
                If mrs限号.RecordCount > 0 Then
                   If Nvl(mrs限号!限号数, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K周二", "周二" & IIf(Nvl(mrs安排!周二) = "", "", "(" & Nvl(mrs安排!周二) & ")")
                    End If
                End If
                mrs限号.Filter = "星期='周三'"
                If mrs限号.RecordCount > 0 Then
                     If Nvl(mrs限号!限号数, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                        "K周三", "周三" & IIf(Nvl(mrs安排!周三) = "", "", "(" & Nvl(mrs安排!周三) & ")")
                    End If
                 End If
                 
                mrs限号.Filter = "星期='周四'"
                If mrs限号.RecordCount > 0 Then
                  If Nvl(mrs限号!限号数, 0) > 0 Then
                    tbWeekTime.Tabs.Add , _
                      "K周四", "周四" & IIf(Nvl(mrs安排!周四) = "", "", "(" & Nvl(mrs安排!周四) & ")")
                  End If
                End If
                mrs限号.Filter = "星期='周五'"
                If mrs限号.RecordCount > 0 Then
                     If Nvl(mrs限号!限号数, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K周五", "周五" & IIf(Nvl(mrs安排!周五) = "", "", "(" & Nvl(mrs安排!周五) & ")")
                     End If
                End If
                
                mrs限号.Filter = "星期='周六'"
                If mrs限号.RecordCount > 0 Then
                   If Nvl(mrs限号!限号数, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                          "K周六", "周六" & IIf(Nvl(mrs安排!周六) = "", "", "(" & Nvl(mrs安排!周六) & ")")
                   End If
                End If
                mrs限号.Filter = "星期='周日'"
                If mrs限号.RecordCount > 0 Then
                    If Nvl(mrs限号!限号数, 0) > 0 Then
                        tbWeekTime.Tabs.Add , _
                            "K周日", "周日" & IIf(Nvl(mrs安排!周日) = "", "", "(" & Nvl(mrs安排!周日) & ")")
                    End If
                End If
                mrs限号.Filter = 0
            End If
            .Visible = tbWeekTime.Tabs.Count <> 0
            If .Tabs.Count > 0 Then
                .Tabs(1).Selected = True
            Else
                MsgBox "该安排没有设置对应的限号数和限约数,请检查!", vbOKOnly, Me.Caption
                Exit Function
            End If
           
        End With
    End If
    str星期 = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "星期='" & str星期 & "'"
    mrs限号.Filter = "星期='" & str星期 & "'"
    txt限号.Text = ""
    txt限约.Text = ""
    If mrs限号.RecordCount <> 0 Then
        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
    End If
     str时点 = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln序号控制 Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "时间段"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "预约人数"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mrsTime!限制数量))
                strTime = mrsTime!时间范围
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
            LoadEditTimePlan = True
            Exit Function
        End If
        .Cols = 7: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        i = 1: r = -1
        lngRow = -1: lngCol = 1
        '******************************************
        With vsTime
         .Cols = 2
         lngRow = -1: lngCol = 0
         '***********************
         '序号填充
         '**********************
         r = mrsTime.RecordCount
         For i = 1 To r
            If mrsTime.EOF Then Exit For
            lngCol = lngCol + 1
            If str时点 <> Nvl(mrsTime!时点) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                str时点 = Nvl(mrsTime!时点)
                If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
                vsTime.TextMatrix(lngRow - 1, 0) = str时点
                vsTime.TextMatrix(lngRow, 0) = str时点
             End If
            strData = mrsTime!序号
            strTime = mrsTime!时间范围
            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
            'If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
            vsTime.TextMatrix(lngRow, lngCol) = strTime
            '是第一项时 填写 开始时间到首行
            If lngCol = 1 Then
            End If
            If Val(Nvl(mrsTime!是否预约)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            mrsTime.MoveNext
         Next
         
         End With
        '******************************************
'        Do While Not mrsTime.EOF
'            If i = 1 Then
'                r = r + 2
'                str时点 = Nvl(mrsTime!时点)
'                If r > .Rows - 1 Then .Rows = .Rows + 2
'                .TextMatrix(r, 0) = str时点
'                .TextMatrix(r - 1, 0) = str时点
'            End If
'            i = i + 1
'            strData = mrsTime!序号
'            strTime = mrsTime!时间范围
'            If i >= .Cols - 1 Then i = 1
'            If r > .Rows - 1 Then .Rows = .Rows + 2
'            .TextMatrix(r, i) = strTime
'            .TextMatrix(r - 1, i) = strData
'
'        Loop
        
        
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        If .Rows > 0 Then
            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        End If
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    
    '设置不可删除项颜色
    '问题号:51427
    lng已挂最大序号 = ExistsBooking(mlng安排ID, Mid(tbWeekTime.SelectedItem.Key, 2))
    If chk序号控制.Value = 1 Then
        '问题号:51427
        For i = 0 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1
                If IsNumeric(vsTime.TextMatrix(i, j)) = True Then
                    If CLng(vsTime.TextMatrix(i, j)) <= lng已挂最大序号 Then
                        vsTime.Cell(flexcpForeColor, i, j) = &HC0C0C0
                        vsTime.Cell(flexcpForeColor, i + 1, j) = &HC0C0C0
                    End If
                End If
            Next
        Next
    End If
    
    LoadEditTimePlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
 
 
 
Private Sub LoadEditTimePlantext(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
    Optional bln计划 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str星期          As String
    Dim i                As Long
    Dim r                As Integer
    Dim str时点          As String
    Dim strTime          As String
    Dim strData          As String
    Dim strKey           As String
     
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsTime Is Nothing Then
        mlngPre安排ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre安排ID = -1
    End If
    If mlngPre安排ID <> lng安排ID Then
        mlngPre安排ID = lng安排ID
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!星期) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!星期), Nvl(mrsTime!星期)
                    strTime = Nvl(mrsTime!星期)
                End If
                .MoveNext
            Loop
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
    End If
    str星期 = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "星期='" & str星期 & "'"
    mrs限号.Filter = "星期='" & str星期 & "'"
    txt限号.Text = ""
    txt限约.Text = ""
    If mrs限号.RecordCount <> 0 Then
        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
    End If
     str时点 = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 400: .RowHeightMin = 400
        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln序号控制 Then
             .Cols = 8: .FixedCols = 0
             .Rows = 1: .FixedRows = 1
             For i = 0 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "时间段"
             Next
             For i = 1 To .Cols - 1 Step 2
                .TextMatrix(0, i) = "预约人数"
             Next
             
             r = 1: i = -1
            Do While Not mrsTime.EOF
                If i * 2 > .Cols - 2 Then r = r + 1: i = -1
                i = i + 1
                strData = Val(Nvl(mrsTime!限制数量))
                strTime = mrsTime!时间范围
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If str时点 <> Nvl(mrsTime!时点) Then
                r = r + 2
                str时点 = Nvl(mrsTime!时点)
                If r > .Rows - 1 Then .Rows = .Rows + 2
                .TextMatrix(r, 0) = str时点
                .TextMatrix(r - 1, 0) = str时点
                i = 0
            End If
            i = i + 1
            strData = mrsTime!序号
            strTime = mrsTime!时间范围
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            .TextMatrix(r - 1, i) = strData
            If Val(Nvl(mrsTime!是否预约)) = 1 Then
                 
                .Cell(flexcpForeColor, r - 1, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r - 1, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeCol(0) = True
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
 
Private Sub LoadTimePlan(ByVal lng安排ID As Long, ByVal bln序号控制 As Boolean, _
    Optional bln计划 As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:
    '入参:
    '编制:
    '日期:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL           As String
    Dim rsTemp           As ADODB.Recordset
    Dim str星期          As String
    Dim i                As Long
    Dim r                As Integer
    Dim str时点          As String
    Dim strTime          As String
    Dim strKey           As String
    On Error GoTo errHandle
    '加载该挂号项目的的停用时间信息
    If mrsTime Is Nothing Then
         mlngPre安排ID = -1
    ElseIf mrsTime.State <> 1 Then
         mlngPre安排ID = -1
    End If
    If mlngPre安排ID <> lng安排ID Then
        mlngPre安排ID = lng安排ID
'        strSQL = "" & _
'        "   Select decode(星期,'周日',1,'周一',2,'周二',3,'周三',4,'周四',5,'周五',6,7) as 排序,星期,to_char(开始时间,'HH24')||':00' as 时点,序号,to_char(开始时间,'hh24:mi')||'-' ||to_char(结束时间,'hh24:mi') as 时间范围, " & _
'        "               限制数量,是否预约" & _
'        "   From  挂号安排时段 " & _
'        "   Where 安排ID=[1]" & _
'        "   Order by 排序,时点,序号"
'        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
        tbWeekTime.Tabs.Clear
        With mrsTime
            strTime = ""
            Do While Not .EOF
                If strTime <> Nvl(mrsTime!星期) Then
                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!星期), Nvl(mrsTime!星期)
                    strTime = Nvl(mrsTime!星期)
                End If
                .MoveNext
            Loop
            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
            If tbWeekTime.Tabs.Count > 0 Then
                tbWeekTime.Tabs(1).Selected = True
            End If
            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
        End With
        If tbWeekTime.Tabs.Count = 0 Then
            MsgBox "该安排没有设置对应的时段,请检查!"
        End If
    End If
    str星期 = "": strTime = ""
    If Not tbWeekTime.SelectedItem Is Nothing Then
        str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    End If
    mrsTime.Filter = "星期='" & str星期 & "'"
    mrs限号.Filter = "星期='" & str星期 & "'"
    txt限号.Text = ""
    txt限约.Text = ""
    If mrs限号.RecordCount <> 0 Then
        Me.txt限号.Text = Nvl(mrs限号!限号数, 0)
        Me.txt限约.Text = Nvl(mrs限号!限约数, 0)
    End If
     str时点 = ""
    With vsTime
        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
        .RowHeightMax = 800: .RowHeightMin = 800
        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
        .FixedRows = 0
        If Not bln序号控制 Then
             .Cols = 8: .FixedCols = 0
             r = 0: i = 0
            Do While Not mrsTime.EOF
                i = i + 1
                If i > .Cols - 1 Then r = r + 1: i = 0
                strTime = "预约" & Val(Nvl(mrsTime!限制数量)) & "人" & vbCrLf & vbCrLf
                strTime = strTime & mrsTime!时间范围
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i) = strTime
                mrsTime.MoveNext
            Loop
            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
            Next
            .Redraw = flexRDBuffered
             Exit Sub
        End If
        Do While Not mrsTime.EOF
            If str时点 <> Nvl(mrsTime!时点) Then
                r = r + 1
                str时点 = Nvl(mrsTime!时点)
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, 0) = str时点
                i = 0
            End If
            i = i + 1
            strTime = mrsTime!序号 & vbCrLf & vbCrLf
            strTime = strTime & mrsTime!时间范围
            If i > .Cols - 1 Then .Cols = .Cols + 1
            If r > .Rows - 1 Then .Rows = .Rows + 1
            .TextMatrix(r, i) = strTime
            If Val(Nvl(mrsTime!是否预约)) = 1 Then
                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
                .Cell(flexcpFontBold, r, i, r, i) = True
            End If
            mrsTime.MoveNext
        Loop
        For i = 1 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 1200
        Next
        .ColWidth(0) = 1200
        .FixedAlignment(0) = flexAlignRightTop
        .ColAlignment(0) = flexAlignRightTop
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
    
Private Sub vsTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
 If vsTime.Row < 0 Or vsTime.Col < 0 Or (chk序号控制.Value = 0 And vsTime.Row < 1) Then cmd预约.Visible = False: mblnCellChange = False: Exit Sub
 
    Call SetCtrlMove
 
    Select Case mViewMode
    Case ViewMode.Edit, ViewMode.NewItem:
       Select Case chk序号控制.Value = 1
            Case True:
            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '设置日期掩码格式
            '******************************************
            If vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
            Case False:
            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
            '******************************************
            '设置日期掩码格式
            '******************************************
            If NewCol Mod 2 = 0 And vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
       End Select
        If vsTime.Row < 0 Or vsTime.Col < 1 Then Exit Sub
        
        If chk序号控制.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
            mblnCellChange = True
        Else
           mblnCellChange = False
        End If
        
    Case ViewMode.ViewItem:
         mblnCellChange = False
         vsTime.Editable = flexEDNone
  End Select
   If mstr限制修改 <> "" Then
        vsTime.Editable = flexEDKbdMouse
        If InStr(mstr限制修改, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone
        
   End If
End Sub



Private Sub vsTime_KeyDown(KeyCode As Integer, Shift As Integer)
    '问题号:51429
        If cmd删除.Visible = False Then Exit Sub
        If KeyCode = 46 Then '快捷键Delete
            Call DeleteSelectPain
        End If
End Sub
Private Sub SetCtrlMove()
    Dim blnDel As Boolean
    With vsTime
        If chk序号控制.Value = 1 Then
            cmd删除.Left = .CellLeft + .CellWidth - cmd删除.Width
            If .Row Mod 2 <> 0 Then
                cmd删除.Top = .CellTop - .CellHeight - 15
            Else
                cmd删除.Top = .CellTop + 15
            End If
            cmd预约.Left = .CellLeft + 15
            cmd预约.Top = cmd删除.Top
            If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
            Else
                blnDel = True
            End If
            blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> ""
            cmd删除.Visible = blnDel And chk序号控制.Value = 1
            cmd预约.Visible = Val(txt限约.Text) <> 0 And InStr(mstr限制修改, mstrKey) = 0
        Else
            cmd预约.Left = .CellLeft + 15
            cmd预约.Top = .CellTop + 15
            cmd预约.Visible = False ' Val(txt限约.Text) <> 0 And InStr(mstr限制修改, mstrKey) = 0
        End If
    End With
End Sub


Private Sub vsTime_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    '**************************************************************
    '当操作员 拖动滚动条时 把 预约按钮 隐藏
    '**************************************************************
    Me.cmd预约.Visible = False
     '问题号:51429
    Me.cmd删除.Visible = False
End Sub
Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If mViewMode = ViewItem Then Exit Sub
    Select Case chk序号控制.Value = 1
        Case True:
            '******************************************
            '专家号时 控制输入
            '******************************************
            If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
        Case False:
            '******************************************
            '普通号时 控制输入
            '******************************************
            If Col Mod 2 = 0 Then
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
            Else
                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
            End If
            
    End Select
   
 
End Sub
 
Private Function validateVsFlex() As Boolean
    '***************************************
    '验证用户对挂号安排时段的修改
    '***************************************
     Dim i          As Long
     Dim j          As Long
     Dim lng预约    As Long
     Dim lng限约    As Long
     Dim lng限号    As Long
     Dim str星期    As String
     If tbWeekTime.SelectedItem Is Nothing Then Exit Function
      str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
     lng限号 = Val(txt限号.Text)
     lng限约 = Val(txt限约.Text)
     If lng限约 = 0 Then lng限约 = lng限号
     Select Case chk序号控制.Value = 1
     Case True:
     '*************************************
     '专家号检查限约数是否大于限号数
     '*************************************
        With vsTime
            For i = 0 To .Rows - 1 Step 2
                For j = 1 To .Cols - 1
                    If .Cell(flexcpForeColor, i, j, i, j) = vbBlue And .TextMatrix(i, j) <> "" Then
                        lng预约 = lng预约 + 1
                    End If
                Next
            Next
        End With
     Case False:
     '*************************************
     '普通号检查限约数是否大于限号数
     '*************************************
        With vsTime
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1 Step 2
                    If .TextMatrix(i, j) <> "" Then
                        lng预约 = lng预约 + Val(.TextMatrix(i, j))
                    End If
                Next
            Next
        End With
     End Select
     If lng预约 > lng限约 Then
        MsgBox "在" & str星期 & "设置的预约数" & lng预约 & "大于了" & IIf(lng限号 = lng限约, "限号数" & lng限约, "限约数" & lng限约) & ",请检查!", vbOKOnly, Me.Caption
        Exit Function
     End If
    validateVsFlex = True
    Exit Function
End Function

Private Function SaveDate() As Boolean
    '*********************************
    '对挂号安排时段进行保存
    '*********************************
    Dim strSQL      As String
    Dim cllSQL      As Collection
    Dim i           As Long
    Dim j           As Long
    Dim blnTrans    As Boolean
    Dim lng安排ID   As Long
    Dim str星期     As String
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim bln预约     As Boolean
    Dim lng限制     As Long '挂号安排时段的限制数量
    Dim bln专家号   As Boolean
    Dim lng序号     As Long
    Dim lngType     As Long
    Dim lng生效计划ID As Long '问题号:51427
    Dim cll计划SQL As Collection '问题号:51427
    
    If validateVsFlex() = False Then Exit Function '进行数据的验证
    
    
    lng安排ID = Val(txt号别.Tag)
    str星期 = mstrKey
    bln专家号 = chk序号控制.Value = 1
    
    Set cllSQL = New Collection
    '****************************************************
    'CREATE OR REPLACE Procedure Zl_挂号安排时段_Delete(
    '安排id_In 挂号安排时段.安排id%Type,
    '星期_In   挂号安排时段.星期 %Type)
    '**********删除以前对此星期安排的时段*****************

    strSQL = "Zl_挂号安排时段_Delete(" & lng安排ID & ",'" & str星期 & "')"
    zlAddArray cllSQL, strSQL
    
   
    Select Case bln专家号
    Case True:
       lng序号 = 0
       For i = 1 To vsTime.Rows - 1 Step 2
            For j = 1 To vsTime.Cols - 1
               If vsTime.TextMatrix(i, j) = "" Then Exit For
               str开始时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
               str结束时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
               lng限制 = 1
               lng序号 = lng序号 + 1
               bln预约 = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
               strSQL = GetInsertSql(lng安排ID, lng序号, str开始时间, str结束时间, 1, bln预约, str星期)
               zlAddArray cllSQL, strSQL
            Next
       Next
    Case False:
        lng序号 = 0
        For i = 1 To vsTime.Rows - 1
            For j = 0 To vsTime.Cols - 1 Step 2
               If vsTime.TextMatrix(i, j) <> "" Then
                str开始时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
                str结束时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
                lng限制 = Val(vsTime.TextMatrix(i, j + 1))
                lng序号 = lng序号 + 1
                bln预约 = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
                strSQL = GetInsertSql(lng安排ID, lng序号, str开始时间, str结束时间, lng限制, bln预约, str星期)
                zlAddArray cllSQL, strSQL
               End If
            Next
        Next
    End Select
     
    If opt本医生.Value Then
        lngType = 1
    ElseIf opt科室.Value Then
        lngType = 2
    ElseIf opt所有.Value Then
        lngType = 3
    End If
    If lngType <> 0 Then
        '--type_in
        '--1-应用与本人
        '--2-应用与本科室
        '--3 or others -应用于所有
       'CREATE OR REPLACE Procedure zl_挂号安排时段_批量应用
       strSQL = "zl_挂号安排时段_批量应用("
       '安排Id_in 挂号安排时段.安排Id%Type,
       strSQL = strSQL & lng安排ID & ","
       'Type_In Number:=1
       strSQL = strSQL & lngType & ")"
       zlAddArray cllSQL, strSQL
    End If
     
  On Error GoTo Errhand
    gcnOracle.BeginTrans
    
    For i = 1 To cllSQL.Count
        strSQL = cllSQL(i)
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
 '问题号:51427
' '跟新挂号计划时段信息
'    lng生效计划ID = 是否有生效计划(lng安排ID, zlDatabase.Currentdate)
'
'    If lng生效计划ID > 0 Then '存在生效的计划，则需要同步跟新挂号计划时段的信息
'        Set cll计划SQL = New Collection
'        strSQL = "Zl_挂号计划时段_Delete(" & lng生效计划ID & ",'" & str星期 & "')"
'        zlAddArray cll计划SQL, strSQL
'
'        Select Case bln专家号
'        Case True:
'           lng序号 = 0
'           For i = 1 To vsTime.Rows - 1 Step 2
'                For j = 1 To vsTime.Cols - 1
'                   If vsTime.TextMatrix(i, j) = "" Then Exit For
'                   str开始时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
'                   str结束时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
'                   lng限制 = 1
'                   lng序号 = lng序号 + 1
'                   bln预约 = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
'                   strSQL = GetInsert计划Sql(lng生效计划ID, lng序号, str开始时间, str结束时间, 1, bln预约, str星期)
'                   zlAddArray cll计划SQL, strSQL
'                Next
'           Next
'        Case False:
'            lng序号 = 0
'            For i = 1 To vsTime.Rows - 1
'                For j = 0 To vsTime.Cols - 1 Step 2
'                   If vsTime.TextMatrix(i, j) <> "" Then
'                    str开始时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(0))
'                    str结束时间 = ConvertToDate(Split(vsTime.TextMatrix(i, j), "-")(1))
'                    lng限制 = Val(vsTime.TextMatrix(i, j + 1))
'                    lng序号 = lng序号 + 1
'                    bln预约 = vsTime.Cell(flexcpForeColor, i, j, i, j) = vbBlue
'                    strSQL = GetInsert计划Sql(lng生效计划ID, lng序号, str开始时间, str结束时间, lng限制, bln预约, str星期)
'                    zlAddArray cll计划SQL, strSQL
'                   End If
'                Next
'            Next
'        End Select
'        For i = 1 To cll计划SQL.Count
'            strSQL = cll计划SQL(i)
'            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
'        Next
'    End If
    gcnOracle.CommitTrans
    SaveDate = True
 Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    Call ErrCenter
    SaveErrLog
End Function

Private Function GetInsertSql(ByVal lngID As Long, ByVal lng序号 As Long, ByVal str开始时间 As String, _
        ByVal str结束时间 As String, ByVal lng限制数量 As Long, ByVal bln是否预约 As Boolean, ByVal str星期 As String)
    '根据提供的信息生成sql语句
    Dim strSQL      As String
   '********************************************************
    '    'CREATE OR REPLACE Procedure Zl_挂号安排时段_Insert
    '    (
    '    安排id_In   挂号安排时段.安排id%Type,
    '    序号_In     挂号安排时段.序号%Type,
    '    开始时间_In 挂号安排时段.开始时间%Type,
    '    结束时间_In 挂号安排时段.结束时间%Type,
    '    限制数量_In 挂号安排时段.限制数量%Type,
    '    是否预约_In 挂号安排时段.是否预约%Type,
    '    星期_In     挂号安排时段.星期%Type
    '    )
    '********************************************************
    strSQL = "  Zl_挂号安排时段_Insert("
     '安排id_In   挂号安排时段.安排id%Type,
    strSQL = strSQL & lngID & ","
     '序号_In     挂号安排时段.序号%Type,
    strSQL = strSQL & lng序号 & ","
     '开始时间_In 挂号安排时段.开始时间%Type,
     strSQL = strSQL & str开始时间 & ","
      '结束时间_In 挂号安排时段.结束时间%Type,
    strSQL = strSQL & str结束时间 & ","
      '限制数量_In 挂号安排时段.限制数量%Type,
    strSQL = strSQL & lng限制数量 & ","
     '是否预约_In 挂号安排时段.是否预约%Type,
    strSQL = strSQL & IIf(bln是否预约, 1, 0) & ","
     '星期_In     挂号安排时段.星期%Type
    strSQL = strSQL & "'" & str星期 & "')"
    GetInsertSql = strSQL
End Function

                             

Private Function ConvertToDate(ByVal strDate As String, Optional ByVal haveYear = False) As String
    '**********************************************************
    '把字符串转换成oracle数据库能够识别的日期
    '**********************************************************
    Select Case haveYear
    Case True:
        ConvertToDate = "To_Date('" & strDate & "', 'YYYY-MM-DD HH24:MI:SS')"
    Case False:
        ConvertToDate = "To_Date('" & strDate & "', 'HH24:MI:SS')"
    End Select
End Function



Private Sub vsTime_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng限号   As Long
  Dim lng限约   As Long
  Dim lng预约数 As Long
  If mViewMode = ViewItem Then Exit Sub

  '*************************************
  '时间进行验证 输入了时间范围
  '**************************************
  If vsTime.Editable = flexEDKbdMouse And vsTime.ColEditMask(vsTime.Col) = strMaskKey Then
    Validate时段 Row, Col, Cancel
    If Not Cancel Then mblnChange = True
    Exit Sub
  End If
  '****************************************
  '在普通号 分时段 对输入的限制预约数进行限制
  '****************************************
   If chk序号控制.Value = 0 And vsTime.ColEditMask(vsTime.Col) <> strMaskKey And vsTime.Editable = flexEDKbdMouse Then
        If vsTime.EditText = "" Then vsTime.EditText = "0"
        mblnChange = True
   End If
End Sub

Private Sub Validate时段(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim i         As Long
  Dim j         As Long
  Dim lng限号   As Long
  Dim lng限约   As Long
  Dim lng预约数 As Long
   
  Dim str时段()  As String
  If mViewMode = ViewItem Then Exit Sub
  
  '*************************************
  '验证时段
  '**************************************
  str时段 = Split(vsTime.EditText, "-")
  If UBound(str时段) <> 1 Then Cancel = True: Exit Sub
   If Not IsDate(str时段(0)) Then Cancel = True: Exit Sub
   If Not IsDate(str时段(1)) Then Cancel = True: Exit Sub
   If CDate(str时段(0)) >= CDate(str时段(1)) Then
        MsgBox "开始时间必须小于结束时间!请检查!", vbOKOnly, Me.Caption
        Cancel = True
   End If
   
End Sub

Private Sub setVsFlexBgColor(Optional ByVal bln序号控制 As Boolean = False)
    '**************************************************************
    '对时间段设置间隔背景
    '**************************************************************
     Dim i           As Long
     If (bln序号控制 And vsTime.Rows = 0) Or (bln序号控制 = False And vsTime.Rows = 1) Then Exit Sub
     For i = IIf(bln序号控制, 0, 1) To vsTime.Rows - 1 Step 2
            vsTime.Cell(flexcpBackColor, i, IIf(bln序号控制, 1, 0), i, vsTime.Cols - 1) = &HE0E0D3
     Next
End Sub
 



Private Sub Init时间段()
  '--------------------------------
  '功能:获取上下班时间段
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("上午上下班时间", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_上午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午上班 = CDate("08:00:00")
    End If
   
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_上午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_上午下班 = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("下午上下班时间", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_时间.dat_下午上班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午上班 = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_时间.dat_下午下班 = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_时间.dat_下午下班 = CDate("1900-01-01 18:00:00")
    End If
    With t_时间
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
         If .dat_上午上班 > .dat_上午下班 Then
            .dat_上午下班 = DateAdd("d", 1, .dat_上午下班)
         End If
    End With
    strSQL = _
    "       Select 时间段, 上班, 下班 " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select 时间段,To_Date('1900-01-01 ' || To_Char(开始时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 开始时间," & vbNewLine & _
    "                               To_Date(Decode(Sign(开始时间 - 终止时间), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(终止时间, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As 终止时间," & _
    "                               Sign(开始时间 - 终止时间) As 隔天, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午上班时间, " & vbNewLine & _
    "                                To_Date('" & Format(t_时间.dat_上午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 上午下班时间, " & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午上班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午上班时间," & vbNewLine & _
    "                                 To_Date('" & Format(t_时间.dat_下午下班, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As 下午下班时间"
    strSQL = strSQL & vbNewLine & _
    "                       From 时间段 )" & vbNewLine & _
    "           Select 时间段, '无' As 标签, 0 As 标志, 开始时间 As 上班, 终止时间 As 下班, 开始时间, 终止时间," & _
    "                  上午上班时间 As 上班时间, 上午下班时间 As 下班时间" & vbNewLine & _
    "            From Tb  Where (开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) And " & _
    "                      (开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select 时间段, '有-上午' As 标签, 1 As 标志, Decode(Sign(上午上班时间 - 开始时间), 1, 上午上班时间, 开始时间) As 上班, " & vbNewLine & _
    "                        Decode(Sign(终止时间 - 上午下班时间), 1, 上午下班时间, 终止时间) As 下班, 开始时间, 终止时间, " & _
    "                        上午上班时间 As 上班时间, 上午下班时间 As 下班时间 " & vbNewLine & _
    "           From Tb a Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 上午下班时间 Or 终止时间 <= 上午上班时间) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select 时间段, '有-下午' As 标签, 1 As 标志, Decode(Sign(下午上班时间 - 开始时间), 1, 下午上班时间, 开始时间) As 上班, " & _
    "                   Decode(Sign(终止时间 - 下午下班时间), 1, 下午下班时间, 终止时间) As 下班, 开始时间, 终止时间, 下午上班时间 As 上班时间, 下午下班时间 As 下班时间 " & vbNewLine & _
    "         From Tb a   Where 时间段 Not In (Select 时间段 From Tb Where 开始时间 >= 下午下班时间 Or 终止时间 <= 下午上班时间)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By 时间段,上班"
     Set mrs上班时间段 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
Private Function Get已约限制(ByVal lng安排ID As Long) As String
    '获取不能修改的安排星期
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    Dim strTmp  As String
    strSQL = "Select Decode(To_Char(A.预约时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7'," & _
    "                             '周六') As 日期 " & vbCrLf & _
    "          From 病人挂号记录 A,挂号安排　B " & vbCrLf & _
    "        Where  A.号别=B.号码 And A.记录状态 = 1 And b.ID = [1] And A.发生时间 > A.登记时间 And A.预约时间 Is Not Null"
    
    If gint预约天数 = 0 Then
        strSQL = strSQL & " And A.预约时间 > Sysdate "
    Else
        strSQL = strSQL & " And A.预约时间 Between Sysdate And Sysdate+" & gint预约天数
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID)
    If rsTmp.EOF Then Exit Function
    
    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!日期)) < 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!日期)
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Get已约限制 = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function Get限号数(ByVal str星期 As String, ByRef lng限号数 As Long, ByRef lng限约数 As Long) As Boolean
    Dim strSQL As String
    If mrs限号 Is Nothing Then
        strSQL = _
        "Select 安排id, 限制项目 as 星期 , 限号数, 限约数 From 挂号安排限制 Where 安排id = [1]"
        Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt号别.Tag))
        If mrs限号.RecordCount = 0 Then
            MsgBox "当前号别没有对应的挂号安排限制" & vbCrLf & "请到挂号安排中设置!", vbOKOnly, Me.Caption
            Set mrs限号 = Nothing
            Exit Function
        End If
    End If
    mrs限号.Filter = "星期='" & str星期 & "'"
    If mrs限号.RecordCount <> 0 Then
        lng限号数 = Val(Nvl(mrs限号!限号数))
        lng限约数 = Val(Nvl(mrs限号!限约数))
        Get限号数 = True
    End If
End Function
'
Private Function 是否有生效计划(lng安排ID As Long, dat日期 As Date) As Long
    '**************************************************************
    '检查该安排下是否有生效的计划
    '参数：lng安排ID-安排ID；dat日期-当前日期
    '返回值:有返回生效计划ID,无返回-1
    '**************************************************************
    Dim strSQL As String
    Dim rs生效计划 As Recordset
    On Error GoTo errH
     strSQL = "" & _
     "      Select A.ID From 挂号安排计划 A" & _
     "      Where A.安排ID=[1] And [2] between Nvl(A.生效时间, [2]) And A.失效时间 And A.审核时间 is Not Null" & _
     "      order By A.生效时间 Desc"
    Set rs生效计划 = zlDatabase.OpenSQLRecord(strSQL, "检查指定排班下是否有生效计划", lng安排ID, dat日期)
    
    If rs生效计划 Is Nothing Then 是否有生效计划 = -1: Exit Function
    If rs生效计划.RecordCount = 0 Then 是否有生效计划 = -1: Exit Function
    rs生效计划.MoveFirst
    是否有生效计划 = rs生效计划!ID
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetInsert计划Sql(ByVal lngID As Long, ByVal lng序号 As Long, ByVal str开始时间 As String, _
        ByVal str结束时间 As String, ByVal lng限制数量 As Long, ByVal bln是否预约 As Boolean, ByVal str星期 As String)
    '根据提供的信息生成计划sql语句
    '问题号:51427
    Dim strSQL      As String
   '********************************************************
    '    'CREATE OR REPLACE Procedure Zl_挂号计划时段_Insert
    '    (
    '    计划ID_In   挂号计划时段.计划ID%Type,
    '    序号_In     挂号计划时段.序号%Type,
    '    开始时间_In 挂号计划时段.开始时间%Type,
    '    结束时间_In 挂号计划时段.结束时间%Type,
    '    限制数量_In 挂号计划时段.限制数量%Type,
    '    是否预约_In 挂号计划时段.是否预约%Type,
    '    星期_In     挂号计划时段.星期%Type
    '    )
    '********************************************************
    strSQL = "  Zl_挂号计划时段_Insert("
     '计划ID_In   挂号计划时段.计划ID%Type,
    strSQL = strSQL & lngID & ","
     '序号_In     挂号计划时段.序号%Type,
    strSQL = strSQL & lng序号 & ","
     '开始时间_In 挂号计划时段.开始时间%Type,
     strSQL = strSQL & str开始时间 & ","
      '结束时间_In 挂号计划时段.结束时间%Type,
    strSQL = strSQL & str结束时间 & ","
      '限制数量_In 挂号计划时段.限制数量%Type,
    strSQL = strSQL & lng限制数量 & ","
     '是否预约_In 挂号计划时段.是否预约%Type,
    strSQL = strSQL & IIf(bln是否预约, 1, 0) & ","
     '星期_In     挂号计划时段.星期%Type
    strSQL = strSQL & "'" & str星期 & "')"
    GetInsert计划Sql = strSQL
End Function

Private Function ExistsBooking(ByVal lng安排ID As String, str星期 As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定号别是否存在预约挂号单
    '入参:str号别-号别;str星期-星期几的安排
    '返回:存在,返回最大挂号序号,不存在返回-1
    '编制:
    '日期:2012-04-26 10:32:02
    '问题号:51657
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "" & _
    "   Select max(号序) as 号序  From 病人挂号记录 A, 挂号安排 B " & _
    "   Where A.号别 = B.号码 " & _
    "       And 记录状态 = 1 and b.id=[1] " & _
    "       And Decode(To_Char(A.发生时间, 'D'), '1', '周日', '2','周一', '3', '周二', '4', '周三', '5', '周四', '6','周五', '7', '周六', Null) =[2]" & _
    "       And A.发生时间 >= Trunc(Sysdate)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng安排ID, str星期)
    ExistsBooking = CLng(Nvl(rsTmp!号序, "-1"))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub DeleteSelectPain()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除选中的时段序号
    '编制:王吉
    '日期:2012-07-12 10:32:02
    '问题号:51429
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str星期 As String
    Dim lng安排ID As Long
    Dim lng最大序号 As Long
    Dim lng当前序号行号 As Long
    Dim lng当前号序 As Long
    Dim blnDel As Boolean
    Dim i As Long
    Dim j As Long
    
    If chk序号控制.Value <> 1 Then Exit Sub
    If vsTime.TextMatrix(vsTime.Row, vsTime.Col) = "" Then Exit Sub
    cmd删除.Visible = False
    cmd预约.Visible = False
    str星期 = Mid(tbWeekTime.SelectedItem.Key, 2)
    lng安排ID = Val(txt号别.Tag)
    lng最大序号 = ExistsBooking(lng安排ID, str星期)
    
    '检查是否是从最后开始删除
    With vsTime
'         For i = 0 To vsTime.Rows - 1
'            For j = 0 To vsTime.Cols - 1
'                If IsNumeric(.TextMatrix(i, j)) = True Then
'                    If lng最大序号 < IIf(.TextMatrix(i, j) = "", "0", .TextMatrix(i, j)) Then
'                        lng最大序号 = .TextMatrix(i, j)
'                    End If
'                End If
'            Next
'         Next

'         If lng最大序号 <> CLng(IIf(.TextMatrix(lng当前序号行号, .Col) = "", "0", .TextMatrix(lng当前序号行号, .Col))) Then
'                MsgBox "只能从最后的号序开始删除！", vbInformation, Me.Caption
'                Exit Sub
'         End If
     If .Row Mod 2 = 0 Then
            lng当前序号行号 = .Row
         Else
            lng当前序号行号 = .Row - 1
     End If
     lng当前号序 = Val(.TextMatrix(lng当前序号行号, .Col))
   
    '检查是否该号别已经被挂出
     If lng最大序号 >= lng当前号序 Then
                MsgBox lng最大序号 & "号已经被挂出,只能删除该号以后的序号！", vbInformation, Me.Caption
                Exit Sub
     End If
     
     SetVsTime lng当前序号行号, .Col
     '清空该序号信息
     
'     .TextMatrix(lng当前序号行号, .Col) = ""
'     .TextMatrix(lng当前序号行号 + 1, .Col) = ""
    End With
End Sub


Public Sub SetVsTime(lngRow As Long, lngCol As Long)
    Dim i As Long
    Dim j As Long
    Dim lng当前序号 As Long
    
    With vsTime
         lng当前序号 = Val(.TextMatrix(lngRow, .Col))
         .TextMatrix(lngRow, .Col) = ""
         .TextMatrix(lngRow + 1, .Col) = ""
         For i = lngRow + 2 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                        .TextMatrix(i, j) = lng当前序号
                         lng当前序号 = lng当前序号 + 1
                    End If
            Next
         Next
    End With
End Sub
