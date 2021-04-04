VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmFinanceSuperviseRollingCurtainEdit 
   Caption         =   "财务收款单"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFinanceSuperviseRollingCurtainEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   11775
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   660
      ScaleHeight     =   1905
      ScaleWidth      =   2685
      TabIndex        =   6
      Top             =   3195
      Width           =   2685
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   870
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   1860
         _cx             =   3281
         _cy             =   1535
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinanceSuperviseRollingCurtainEdit.frx":6852
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
   Begin VB.PictureBox picDown 
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
      Height          =   2340
      Left            =   120
      ScaleHeight     =   2340
      ScaleWidth      =   11490
      TabIndex        =   23
      Top             =   5160
      Width           =   11490
      Begin VB.CommandButton cmdCashMoney 
         Caption         =   "点钞(&D)"
         Height          =   350
         Left            =   -15
         TabIndex        =   32
         Top             =   1875
         Width           =   1100
      End
      Begin VB.TextBox txtActual 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4800
         MaxLength       =   16
         TabIndex        =   18
         Top             =   810
         Width           =   2625
      End
      Begin VB.TextBox txtLendTotal 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   8640
         TabIndex        =   15
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtBorrowTotal 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   4800
         TabIndex        =   13
         Top             =   405
         Width           =   2625
      End
      Begin VB.TextBox txtPrepay 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   975
         TabIndex        =   11
         Top             =   390
         Width           =   2625
      End
      Begin VB.TextBox txtMemo 
         Height          =   350
         Left            =   975
         MaxLength       =   500
         TabIndex        =   9
         Top             =   -15
         Width           =   10305
      End
      Begin VB.TextBox txtTime 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   350
         Left            =   975
         TabIndex        =   24
         Top             =   1245
         Width           =   2625
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   9975
         TabIndex        =   21
         Top             =   1875
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   8775
         TabIndex        =   20
         Top             =   1875
         Width           =   1100
      End
      Begin VB.Label lblRemainMoney 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   8655
         TabIndex        =   31
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label lblSupposeMoney 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   975
         TabIndex        =   30
         Top             =   810
         Width           =   2625
      End
      Begin VB.Label lblLendTotal 
         AutoSize        =   -1  'True
         Caption         =   "借出合计"
         Height          =   210
         Left            =   7800
         TabIndex        =   14
         Top             =   450
         Width           =   840
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "收款时间"
         Height          =   210
         Left            =   45
         TabIndex        =   25
         Top             =   1275
         Width           =   840
      End
      Begin VB.Label lblRemain 
         AutoSize        =   -1  'True
         Caption         =   "本次暂存"
         Height          =   210
         Left            =   7815
         TabIndex        =   19
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblActual 
         AutoSize        =   -1  'True
         Caption         =   "现金实收"
         Height          =   210
         Left            =   3960
         TabIndex        =   17
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblSuppose 
         AutoSize        =   -1  'True
         Caption         =   "现金应收"
         Height          =   210
         Left            =   60
         TabIndex        =   16
         Top             =   885
         Width           =   840
      End
      Begin VB.Label lblBorrowTotal 
         AutoSize        =   -1  'True
         Caption         =   "借款合计"
         Height          =   210
         Left            =   3960
         TabIndex        =   12
         Top             =   450
         Width           =   840
      End
      Begin VB.Label lblPrepay 
         AutoSize        =   -1  'True
         Caption         =   "冲预交"
         Height          =   210
         Left            =   255
         TabIndex        =   10
         Top             =   450
         Width           =   630
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "摘要"
         Height          =   210
         Left            =   465
         TabIndex        =   8
         Top             =   45
         Width           =   420
      End
      Begin VB.Line linMain 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   10440
         Y1              =   1650
         Y2              =   1650
      End
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   75
      ScaleHeight     =   825
      ScaleWidth      =   11175
      TabIndex        =   22
      Top             =   510
      Width           =   11175
      Begin VB.TextBox txtGroups 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Left            =   3735
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   450
         Width           =   2490
      End
      Begin VB.ComboBox cboNO 
         Height          =   330
         Left            =   8925
         TabIndex        =   26
         Top             =   75
         Width           =   2040
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   3660
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   450
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   473
         Width           =   1785
      End
      Begin VB.Label lblGroups 
         AutoSize        =   -1  'True
         Caption         =   "人员分组"
         Height          =   240
         Left            =   2775
         TabIndex        =   29
         Top             =   510
         Width           =   960
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   210
         Left            =   8565
         TabIndex        =   27
         Top             =   135
         Width           =   210
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "缴款部门"
         Height          =   210
         Left            =   2700
         TabIndex        =   2
         Top             =   525
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "缴款人"
         Height          =   210
         Left            =   135
         TabIndex        =   0
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.PictureBox picRollingCurtain 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   2640
      ScaleHeight     =   1740
      ScaleWidth      =   8370
      TabIndex        =   4
      Top             =   1650
      Width           =   8370
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   33
         Top             =   60
         Width           =   210
         Begin VB.Image imgColPlan 
            Height          =   195
            Left            =   0
            Picture         =   "frmFinanceSuperviseRollingCurtainEdit.frx":68CC
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRollingCurtain 
         Height          =   930
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   10740
         _cx             =   18944
         _cy             =   1640
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   12632256
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFinanceSuperviseRollingCurtainEdit.frx":6E1A
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmFinanceSuperviseRollingCurtainEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Private mlngModule As Long, mstrPrivs As String
Private mstr轧帐IDs As String, mlngGroupID As Long
Private mlng缴款人ID As Long, mstr缴款人 As String
Private Enum mPaneIndex
    EM_PN_表头信息 = 1
    EM_PN_轧帐信息 = 2
    EM_PN_结算信息 = 3
    EM_PN_表尾信息 = 4
End Enum
Private mblnOK As Boolean
Private mblnNotBrush As Boolean
Private mrsBalance As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnChange  As Boolean
Private Sub LoadBalance(ByVal blnReOpenRecord As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算信息
    '入参:blnReOpenRecord-重新打开记录集
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-10 14:50:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String, bytType As Byte, i As Long
   Dim lng轧帐ID As Long, blnSel As Boolean, str轧帐IDs As String
   Dim str结算方式 As String, lngRow As Long, blnFind As Boolean
   Dim dblTotal(0 To 3) As Double
   
    On Error GoTo errHandle
    
    If mrsBalance Is Nothing Or blnReOpenRecord Then
        strSQL = "" & _
        "   Select /*+ rule */ decode(nvl(M.性质,0),1,1,2,2,3,10,4,11,4) as 序号,A.ID as 收缴ID,  " & _
        "           b.结算方式,b.金额,b.结算号,b.余额,b.结算号," & _
        "           a.冲预交款 as 冲预交,A.借入合计 as 借款合计,A.借出合计 " & _
        "   From 人员收缴记录 A, 人员收缴明细 B,结算方式 M, Table(f_Num2list([2])) J" & _
        "   Where a.Id = b.收缴id  And A.记录性质=[1]  And A.ID=J.Column_Value and B.结算方式=M.名称(+) and nvl(金额,0)<>0 " & _
        "   Order by 序号,结算方式"
        bytType = IIf(mlngGroupID <> 0, 3, 1)
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytType, mstr轧帐IDs)
    End If
    For i = 0 To 3
        dblTotal(i) = 0
    Next
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            lng轧帐ID = Val(.TextMatrix(i, .ColIndex("ID")))
            blnSel = Val(.TextMatrix(i, .ColIndex("选择"))) <> 0
            If blnSel And lng轧帐ID <> 0 Then
                '需要汇总
                str轧帐IDs = str轧帐IDs & "," & lng轧帐ID
                dblTotal(0) = dblTotal(0) + Val(.TextMatrix(i, .ColIndex("冲预交款")))
                dblTotal(1) = dblTotal(1) + Val(.TextMatrix(i, .ColIndex("借入合计")))
                dblTotal(2) = dblTotal(2) + Val(.TextMatrix(i, .ColIndex("借出合计")))
            End If
        Next
    End With
 
    With vsBalance
        .Redraw = flexRDNone
        .Clear 1: .Rows = 2
        If str轧帐IDs = "" Then GoTo goEnd:
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            str结算方式 = NVL(mrsBalance!结算方式)
            If Val(NVL(mrsBalance!金额)) <> 0 _
                And InStr(str轧帐IDs & ",", "," & NVL(mrsBalance!收缴ID) & ",") > 0 Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If str结算方式 = Trim(.TextMatrix(i, .ColIndex("结算方式"))) Then
                        blnFind = True: lngRow = i: Exit For
                    End If
                Next
                If blnFind = False Then
                    If .TextMatrix(.Rows - 1, .ColIndex("结算方式")) = "" Then
                         lngRow = .Rows - 1
                    Else
                        .Rows = .Rows + 1: lngRow = .Rows - 1
                    End If
                End If
                .TextMatrix(lngRow, .ColIndex("序号")) = NVL(mrsBalance!序号)
                .TextMatrix(lngRow, .ColIndex("结算方式")) = str结算方式
                .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(.TextMatrix(lngRow, .ColIndex("金额"))) + Val(NVL(mrsBalance!金额)), gstrDec)
                
                If InStr(Mid(str轧帐IDs, 2), ",") > 0 Then
                    '一次收取多个时，需要重新录入
                    .TextMatrix(lngRow, .ColIndex("结算号码")) = ""
                Else
                    '只针对一个轧帐记录收取时，提取原结算号码
                    .TextMatrix(lngRow, .ColIndex("结算号码")) = NVL(mrsBalance!结算号)
                End If
                If Val(NVL(mrsBalance!序号)) = 1 Then
                    '现金合计
                    dblTotal(3) = dblTotal(3) + Val(NVL(mrsBalance!金额))
                End If
            End If
            mrsBalance.MoveNext
        Loop
goEnd:
        .Cell(flexcpSort, 1, .ColIndex("序号"), .Rows - 1, .ColIndex("序号")) = flexSortNumericAscending
        .Redraw = flexRDBuffered
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .ColWidth(.ColIndex("结算号码")) = 3000
    End With
    '恢复列设置
    'zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "结算方式列表", False
    '加载合计数据
    txtPrepay.Text = Format(dblTotal(0), "##0.00;-##0.00;;")
    txtBorrowTotal.Text = Format(dblTotal(1), "##0.00;-##0.00;;")
    txtLendTotal.Text = Format(dblTotal(2), "##0.00;-##0.00;;")
    lblSupposeMoney.Caption = Format(dblTotal(3), "##0.00;-##0.00;;")
    txtActual.Text = Format(dblTotal(3), "##0.00;-##0.00;;")
    lblRemainMoney.Caption = Format(0, "##0.00;-##0.00;0.00;0.00")
    txtActual.Enabled = dblTotal(3) <> 0 And mlngGroupID = 0
    txtActual.BackColor = IIf(txtActual.Enabled, &H80000005, txtLendTotal.BackColor)
    
    Exit Sub
errHandle:
    vsBalance.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function LoadGroup() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载财务组信息
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-24 12:16:41
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim strSQL As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If mlngGroupID = 0 Then LoadGroup = True: Exit Function
    '读取财务组
    strSQL = " " & _
    "   Select a.Id As 编码, a.组名称, a.简码, b.姓名 As 组负责人,A.说明" & _
    "   From 财务缴款分组 A, 人员表 B " & _
    "   Where a.负责人id = b.Id and A.ID=[1] " & _
    "   Order By a.组名称"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取财务组信息", mlngGroupID)
    If rsTemp.RecordCount <> 0 Then
        txtGroups.Text = NVL(rsTemp!组名称)
    End If
    LoadGroup = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Function zlShowMe(ByVal frmMain As Object, _
    ByVal lngModule As String, ByVal strPrivs As String, _
    ByVal str缴款人 As String, ByVal lng缴款人ID As Long, ByVal str轧帐IDs As String, _
    Optional ByVal lngGroupID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain-调用的主窗体
    '       lngModule-模块号
    '       strPrivs-权限串
    '       str轧帐IDs-本次要收款的轧帐IDS数据
    '       lngGroupID>0:针对财务组收款(即财务组ID)
    '返回:收款成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-10 14:08:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnOK = False: mlngModule = lngModule: mstrPrivs = strPrivs
    mstr轧帐IDs = str轧帐IDs: mlngGroupID = lngGroupID
    mstr缴款人 = str缴款人: mlng缴款人ID = lng缴款人ID
    
    Call ClearData
    Call SetCtrlEnable
    txtName.Text = mstr缴款人
    'If LoadDept = False Then Unload Me: Exit Function
    If LoadGroup = False Then Unload Me: Exit Function
    If LoadCollectData = False Then Unload Me: Exit Function
    mblnChange = False
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlShowMe = mblnOK
End Function

Private Function LoadDept() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载缴款人部门信息
    '编制:刘兴洪
    '日期:2013-09-11 14:05:08
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
        
    strSQL = "" & _
    "   Select Distinct a.Id, a.编码, a.名称,b.缺省" & vbNewLine & _
    "   From 部门表 a, 部门人员 b" & vbNewLine & _
    "   Where a.Id = b.部门id And b.人员ID=[1] " & vbNewLine & _
     "              And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
    "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
    "   Order By a.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng缴款人ID)
    With cboDept
        .Clear
        Do While Not rsTemp.EOF
            .AddItem NVL(rsTemp!编码) & "-" & rsTemp!名称
            .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!缺省)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 And .ListCount <> 0 Then .ListIndex = 0
    End With


    LoadDept = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载轧帐数据
    '编制:刘兴洪
    '日期:2013-09-11 17:34:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
      Dim i As Long, strHead As String, varData As Variant
   Dim lngWidth As Long
    strHead = "过滤,选择,ID,轧帐单号,开始时间,终止时间,轧帐人,轧帐时间,收款员,收款部门,冲预交款,借入合计,借出合计,小组收款人,小组收款时间,轧帐说明"
    varData = Split(strHead, ",")
    With vsRollingCurtain
        .Clear: .Rows = 2: .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = varData(i)
            .ColKey(i) = varData(i)
            If .ColKey(i) = "过滤" Then .TextMatrix(0, i) = ""
            If .ColKey(i) = "过滤" Or .ColKey(i) = "选择" Or .ColKey(i) = "ID" Or .ColKey(i) = "收款部门" Then .ColData(i) = "-1|1"
            If .ColKey(i) = "轧帐人" Or .ColKey(i) = "冲预交款" Or .ColKey(i) = "借入合计" Or .ColKey(i) = "借出合计" Then .ColHidden(i) = True
            If .ColKey(i) = "轧帐单号" Or .ColKey(i) = "开始时间" Or .ColKey(i) = "终止时间" Or .ColKey(i) = "轧帐时间" Then .ColData(i) = "1|0"
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "选择" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
        
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        lngWidth = .ColWidth(.ColIndex("选择"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
        .ColWidth(.ColIndex("选择")) = lngWidth
        .Editable = flexEDKbdMouse
    End With
    
    With vsBalance
           Set .Font = Me.Font
           .Clear 1
           .Cols = 4: .Rows = 2
           .FixedRows = 1
           .TextMatrix(0, 0) = "序号"
           .TextMatrix(0, 1) = "结算方式"
           .TextMatrix(0, 2) = "金额"
           .TextMatrix(0, 3) = "结算号码"
           For i = 0 To .Cols - 1
               .ColKey(i) = .TextMatrix(0, i)
               If i = .ColIndex("金额") Then
                   .ColAlignment(i) = flexAlignRightCenter
               Else
                   .ColAlignment(i) = flexAlignLeftCenter
               End If
               .FixedAlignment(i) = flexAlignCenterCenter
           Next
           .ColHidden(.ColIndex("序号")) = True
           .AutoSizeMode = flexAutoSizeColWidth
           .AutoResize = True
           Call .AutoSize(0, .Cols - 1)
           
           .ColWidth(.ColIndex("结算号码")) = .ColWidth(.ColIndex("结算号码")) * 3
           .ExtendLastCol = False
           'zl_vsGrid_Para_Restore mlngModule, vsBalance, Me.Name, "结算方式列表", False
           .Editable = flexEDKbdMouse
       End With
End Sub
Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面数据
    '编制:刘兴洪
    '日期:2013-10-10 14:21:53
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitGrid
    txtMemo.Text = ""
    txtPrepay.Text = ""
    txtBorrowTotal.Text = ""
    txtLendTotal.Text = ""
    lblSupposeMoney.Caption = ""
    txtActual.Text = ""
    lblRemainMoney.Caption = ""
End Sub
Public Function LoadCollectData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-26 11:38:15
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lng轧帐ID As Long, bytType As Byte, i As Long, lngWidth As Long
 
    On Error GoTo errHandle
    txtTime.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    If mstr轧帐IDs = "" Then
        MsgBox "你当未选择任何轧帐记录，不能进行收款操作员!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
     
    strSQL = "" & _
    "   Select /*+ rule */-1 as 选择,a.Id,a.No As 轧帐单号, a.开始时间, a.终止时间, a.登记人 As 轧帐人, a.登记时间 As 轧帐时间,  " & _
    "         a.收款员 ,b.名称 As 收款部门, " & _
    "         ltrim(to_char(a.冲预交款,'9999999999990.00')) as 冲预交款, " & _
    "         ltrim(to_char(a.借入合计,'9999999999990.00')) as 借入合计, " & _
    "         ltrim(to_char(a.借出合计,'9999999999990.00')) as 借出合计," & _
    "         a.小组收款人, To_Char(a.小组收款时间, 'yyyy-mm-dd hh24:mi:ss') As 小组收款时间, " & _
    "         a.摘要 As 轧帐说明" & _
    "  From 人员收缴记录 A, 部门表 B, Table(f_Num2list([2])) J " & _
    "  Where a.收款部门id = b.Id(+) And A.ID=J.Column_Value And a.记录性质 = [1] " & _
    "               And A.作废时间 is Null and A.财务收款ID is null  " & _
    "  Order by 登记时间 desc,轧帐单号 desc,小组收款时间 desc"

    bytType = IIf(mlngGroupID <> 0, 3, 1)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, bytType, mstr轧帐IDs)
    With vsRollingCurtain
        mblnNotBrush = True
        .Clear 1: .Rows = 2
        .FixedRows = 1
        .Redraw = flexRDNone
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, .ColIndex("选择")) = -1
            .TextMatrix(.Rows - 1, .ColIndex("ID")) = NVL(rsTemp!ID)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐单号")) = NVL(rsTemp!轧帐单号)
            .TextMatrix(.Rows - 1, .ColIndex("开始时间")) = NVL(rsTemp!开始时间)
            .TextMatrix(.Rows - 1, .ColIndex("终止时间")) = NVL(rsTemp!终止时间)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐人")) = NVL(rsTemp!轧帐人)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐时间")) = NVL(rsTemp!轧帐时间)
            .TextMatrix(.Rows - 1, .ColIndex("收款员")) = NVL(rsTemp!收款员)
            .TextMatrix(.Rows - 1, .ColIndex("冲预交款")) = NVL(rsTemp!冲预交款)
            .TextMatrix(.Rows - 1, .ColIndex("借入合计")) = NVL(rsTemp!借入合计)
            .TextMatrix(.Rows - 1, .ColIndex("借出合计")) = NVL(rsTemp!借出合计)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款人")) = NVL(rsTemp!小组收款人)
            .TextMatrix(.Rows - 1, .ColIndex("小组收款时间")) = NVL(rsTemp!小组收款时间)
            .TextMatrix(.Rows - 1, .ColIndex("轧帐说明")) = NVL(rsTemp!轧帐说明)
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        For i = 0 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            If .ColKey(i) = "收款部门" Then .ColHidden(i) = True
            If .ColKey(i) Like "*ID" Or .ColKey(i) = "收款员" Then .ColHidden(i) = True
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*时间" Or .ColKey(i) = "轧帐单号" Then
                .ColAlignment(i) = flexAlignCenterCenter
            ElseIf .ColKey(i) Like "*合计" Or .ColKey(i) = "冲预交款" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) = "选择" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColDataType(i) = flexDTBoolean
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        If .Rows > 2 Then .Rows = .Rows - 1
        .Row = 1
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(1, .Cols - 1)
        lngWidth = .ColWidth(.ColIndex("选择"))
        zl_vsGrid_Para_Restore mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False
         .ColWidth(.ColIndex("选择")) = lngWidth
        If .Enabled And .Visible Then .SetFocus
        .Redraw = flexRDBuffered
    End With
    '加载结算数据
    Call LoadBalance(True)
    LoadCollectData = True
    mblnNotBrush = False
    Exit Function
errHandle:
    vsBalance.Redraw = flexRDBuffered
    vsRollingCurtain.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Function
Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化区域
    '编制:刘兴洪
    '日期:2013-10-10 11:51:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane
    Dim lngHeight As Long, lngTemp As Long
    lngHeight = 1740 / Screen.TwipsPerPixelY
    With dkpMan
        lngTemp = 825 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(EM_PN_表头信息, 100, lngTemp, DockLeftOf, Nothing)
        objPane.Title = "表头": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngTemp: objPane.MaxTrackSize.Height = lngTemp
        objPane.Handle = picTop.hWnd
        
        Set objPane = .CreatePane(EM_PN_轧帐信息, 100, lngHeight, DockBottomOf, objPane)
        objPane.Title = "轧帐信息": objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngHeight
        objPane.Handle = picRollingCurtain.hWnd
        Set objPane = .CreatePane(EM_PN_结算信息, 400, 400, DockBottomOf, objPane)
        objPane.Title = "结算明细"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        objPane.MinTrackSize.Height = 1000 / Screen.TwipsPerPixelY
        
        lngTemp = 2340 / Screen.TwipsPerPixelY
        Set objPane = .CreatePane(EM_PN_表尾信息, 100, lngTemp, DockBottomOf, objPane)
        objPane.Title = "表尾": objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable
        objPane.MinTrackSize.Height = lngTemp: objPane.MaxTrackSize.Height = lngTemp
        objPane.Handle = picDown.hWnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
    End With
End Function
 

'Private Sub cboDept_Click()
'    mblnChange = True
'End Sub

'Private Sub cboDept_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then vsRollingCurtain.SetFocus
'End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdCashMoney_Click()
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:现金点钞
    '编制:刘兴洪
    '日期:2013-09-13 16:08:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    Dim objCash As New clsChargeBill
    objCash.CheckCash Me, dblMoney
    Set objCash = Nothing
End Sub

Private Sub cmdOK_Click()
    Dim str轧帐IDs As String, str未选轧帐IDs As String
    Dim strNO As String
    str轧帐IDs = GetSelRollingCurtainIds(str未选轧帐IDs)
    If isValied(str轧帐IDs) = False Then Exit Sub
    If SaveData(str轧帐IDs, strNO) = False Then Exit Sub
    mblnOK = True
    '打印收据
    cboNO.AddItem strNO
    Call BillPrint(strNO)
    mblnChange = False
    If str未选轧帐IDs = "" Then Unload Me: Exit Sub
     '重新加载数据
     mstr轧帐IDs = str未选轧帐IDs
    Call LoadCollectData
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
'    If cboDept.ListCount = 1 And cboDept.Enabled And cboDept.Visible Then
'        cboDept.SetFocus
'    End If
End Sub

Private Sub Form_Load()
    Call InitPanel
    RestoreWinState Me, App.ProductName
    mblnFirst = True
End Sub
Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Width < 12015 * 0.8 Then Width = 12015 * 0.8
    If Height < 9075 * 0.8 Then Height = 9075 * 0.8
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        End If
    End If
    
    Err = 0: On Error Resume Next
    Call SaveWinState(Me, App.ProductName)
    Set mrsBalance = Nothing
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Top = .ScaleTop
        vsBalance.Left = .ScaleLeft
        vsBalance.Width = .ScaleWidth
        vsBalance.Height = .ScaleHeight
    End With
End Sub
Private Sub picDown_Resize()
    Dim lngSplit As Long
    Err = 0: On Error Resume Next
    lngSplit = (picDown.ScaleWidth - 600) \ 3
    With picDown
        txtMemo.Width = .ScaleWidth - txtMemo.Left - 50
        linMain.X2 = .ScaleWidth
        txtPrepay.Width = lngSplit - txtPrepay.Left
        lblSupposeMoney.Width = txtPrepay.Width
        txtTime.Width = txtPrepay.Width
        
        lblBorrowTotal.Left = lngSplit + 300
        txtBorrowTotal.Left = lblBorrowTotal.Left + lblBorrowTotal.Width
        txtBorrowTotal.Width = lngSplit * 2 + 300 - txtBorrowTotal.Left
        lblActual.Left = lblBorrowTotal.Left
        txtActual.Left = txtBorrowTotal.Left
        txtActual.Width = txtBorrowTotal.Width
        
        lblLendTotal.Left = lngSplit * 2 + 600
        txtLendTotal.Left = lblLendTotal.Left + lblLendTotal.Width
        txtLendTotal.Width = .ScaleWidth - txtLendTotal.Left - 50
        lblRemain.Left = lblLendTotal.Left
        lblRemainMoney.Left = txtLendTotal.Left
        lblRemainMoney.Width = txtLendTotal.Width
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Sub picRollingCurtain_Resize()
    Err = 0: On Error Resume Next
    With picRollingCurtain
        vsRollingCurtain.Top = .ScaleTop
        vsRollingCurtain.Left = .ScaleLeft
        vsRollingCurtain.Width = .ScaleWidth
        vsRollingCurtain.Height = .ScaleHeight
    End With
End Sub

 

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    cboNO.Left = picTop.ScaleWidth - cboNO.Width - 50
    lblNO.Left = cboNO.Left - lblNO.Width - 10
End Sub

Private Sub txtActual_Change()
    lblRemainMoney.Caption = Format(Val(lblSupposeMoney.Caption) - Val(txtActual.Text), "0.00")
    If Val(lblRemainMoney.Caption) <> 0 Then
        lblRemainMoney.ForeColor = vbRed
    Else
        lblRemainMoney.ForeColor = txtActual.ForeColor
    End If
    mblnChange = True
End Sub

Private Sub txtActual_GotFocus()
    zlCommFun.OpenIme False
    zlControl.TxtSelAll txtActual
End Sub

Private Sub txtActual_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtMemo_Change()
    mblnChange = True
End Sub

Private Sub txtMemo_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtMemo
End Sub
Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub txtMemo_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtMemo_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsRollingCurtain_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("选择")
            Call LoadBalance(False)
            mblnChange = True
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("选择")
            Cancel = True
        Case Else
            Exit Sub
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng轧帐ID As Long
    With vsRollingCurtain
        Select Case Col
        Case .ColIndex("选择")
            lng轧帐ID = Val(.TextMatrix(Row, .ColIndex("ID")))
           ' If lng轧帐ID = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsRollingCurtain_GotFocus()
    Call zl_VsGridGotFocus(vsRollingCurtain)
End Sub
Private Sub vsRollingCurtain_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: vsBalance.SetFocus
End Sub

Private Sub vsRollingCurtain_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsRollingCurtain)
    vsRollingCurtain.Tag = "0"
End Sub
Private Sub vsRollingCurtain_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsRollingCurtain_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsRollingCurtain, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsRollingCurtain_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsBalance_GotFocus()
    Call zl_VsGridGotFocus(vsBalance)
End Sub
Private Sub vsBalance_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsBalance)
End Sub
Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    'zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "结算方式列表", False, zlCheckPrivs(mstrPrivs, "参数设置")
End Sub
Private Sub vsBalance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call zl_VsGridRowChange(vsBalance, OldRow, NewRow, OldCol, NewCol)
End Sub
Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    'zl_vsGrid_Para_Save mlngModule, vsBalance, Me.Name, "结算方式列表", False, zlCheckPrivs(mstrPrivs, "参数设置")
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBalance
        Select Case Col
        Case .ColIndex("结算号码")
            If .TextMatrix(Row, .ColIndex("结算方式")) Like "*冲预交*" _
                Or .TextMatrix(Row, Col) Like "*借款合计*" _
                Or .TextMatrix(Row, Col) Like "*借款合计*" Then
                Cancel = True: Exit Sub
            End If
        Case Else
            Cancel = True: Exit Sub
        End Select
    End With
End Sub
Private Sub vsBalance_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsBalance
        If .Col = .Cols - 1 And .Row = .Rows - 1 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
    End With
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("结算方式"), vsBalance.Cols - 1, False)
End Sub
Private Sub vsBalance_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call zlVsMoveGridCell(vsBalance, vsBalance.ColIndex("结算方式"), vsBalance.Cols - 1, False)
End Sub
Private Sub vsBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Or KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub
Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsBalance
        If Row <= 1 Then Exit Sub
            VsFlxGridCheckKeyPress vsBalance, Row, Col, KeyAscii, m文本式
            If KeyAscii = Asc("'") Or KeyAscii = Asc("|") Or KeyAscii = Asc(",") Then KeyAscii = 0
    End With
End Sub
Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer
    '数据验证
    With vsBalance
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        Select Case Col
        Case .ColIndex("结算号码")
            If zlCommFun.ActualLen(strKey) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
            If InStr(1, strKey, "'") > 0 Or InStr(1, strKey, "|") > 0 Or InStr(1, strKey, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                Cancel = True: Exit Sub
            End If
        Case Else
        End Select
    End With
End Sub
Private Sub SetCtrlEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件Enabled属性
    '编制:刘兴洪
    '日期:2013-10-10 17:08:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txtName.Enabled = False
    txtGroups.Visible = mlngGroupID <> 0
    lblGroups.Visible = mlngGroupID <> 0
 End Sub
Private Function isValied(ByVal strSel轧帐IDs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存前的数据合法性检查
    '入参:strSel轧帐IDs-选中的轧帐IDs(多个用逗号分离)
    '返回:数据合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-10 17:31:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strTemp As String, rsTemp As ADODB.Recordset, strSQL As String
    Dim str轧帐IDs As String, lng轧帐ID As Long, dblMoney As Double, blnFind As Boolean
    
    On Error GoTo errHandle
    If strSel轧帐IDs = "" Then
        MsgBox "收款时,必须至少选中一条轧帐记录,不能进行收款!", vbInformation, gstrSysName
        If vsRollingCurtain.Visible And vsRollingCurtain.Enabled Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(strSel轧帐IDs) > 4000 Then
        MsgBox "收款时,选择的轧帐记录过多,请取消部分针对轧帐记录的收款!", vbInformation, gstrSysName
        If vsRollingCurtain.Visible And vsRollingCurtain.Enabled Then vsRollingCurtain.SetFocus
        Exit Function
    End If
    '问题号:110281,焦博,2017/08/15,把轧账说明的上限从50个字符调整为500个字符
    If zlCommFun.ActualLen(txtMemo.Text) > 500 Then
        MsgBox "摘要超长,最多只能输入250个汉子或500个字符", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        MsgBox "摘要中不能包含单引号!", vbInformation, gstrSysName
        If txtMemo.Visible And txtMemo.Enabled Then txtMemo.SetFocus
        Exit Function
    End If
'    If cboDept.ListIndex < 0 Then
'        MsgBox "未选择缴款部门!", vbInformation, gstrSysName
'        If cboDept.Visible And cboDept.Enabled Then cboDept.SetFocus
'        Exit Function
'    End If
    If Val(txtActual.Text) > Val(lblSupposeMoney.Caption) Then
        MsgBox "现金实收金额不能大于现金应收金额!", vbInformation, gstrSysName
        If txtActual.Visible And txtActual.Enabled Then txtActual.SetFocus
        Exit Function
    End If
    
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = .TextMatrix(i, .ColIndex("结算号码"))
            If zlCommFun.ActualLen(strTemp) > 10 Then
                MsgBox "结算号码超长,最多只能输入10个字符或5个汉字", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("结算号码")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
            If InStr(1, strTemp, "'") > 0 Or InStr(1, strTemp, "|") > 0 Or InStr(1, strTemp, ",") > 0 Then
                MsgBox "结算号码中不能包含特殊字符:',| ", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("结算号码")
                If Not .RowIsVisible(.Row) Or Not .ColIsVisible(.Col) = True Then
                    .TopRow = .Row: .LeftCol = .Col
                End If
                If .Visible And .Enabled Then .SetFocus
                Exit Function
            End If
        Next
    End With
    '总金额检查
    strSQL = "" & _
       "   Select  b.结算方式,sum(b.金额) as 金额 " & _
       "   From  人员收缴明细 B, Table(f_Num2list([1])) J" & _
       "   Where  B.收缴ID=J.Column_Value " & _
       "   Group by  B.结算方式"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr轧帐IDs)
    With vsBalance
        For i = 1 To .Rows - 1
            strTemp = Trim(.TextMatrix(i, .ColIndex("结算方式")))
            rsTemp.Filter = "结算方式='" & strTemp & "'"
'            If rsTemp.EOF And Val(txtPrepay.Text) = 0 And Val(txtBorrowTotal.Text) = 0 And Val(txtLendTotal.Text) = 0 Then
'               If MsgBox(" 在结算明细列表中" & strTemp & "的结算方式 " & vbCrLf & _
'                "在选中的轧帐记录中不存在,可能是因为并发原因造成的," & vbCrLf & _
'                "为了保证数据的一致性,你需要重新提取数据," & vbCrLf & _
'                "你是否要重新提取数据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
'                Call LoadCollectData
'               End If
'                If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
'                Exit Function
'            End If
            If Not rsTemp.EOF Then
                dblMoney = Val(NVL(rsTemp!金额))
                If dblMoney <> Val(.TextMatrix(i, .ColIndex("金额"))) Then
                   If MsgBox(" 在结算明细列表中" & strTemp & "的合计数与 " & vbCrLf & _
                    "选中的轧帐记录的合计数不一致,可能是因为并发原因造成的," & vbCrLf & _
                    "为了保证数据的一致性,你需要重新提取数据," & vbCrLf & _
                    "你是否要重新提取数据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    Call LoadCollectData
                   End If
                    If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
                    Exit Function
                End If
            End If
        Next
        rsTemp.Filter = 0
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            strTemp = NVL(rsTemp!结算方式)
            dblMoney = Val(NVL(rsTemp!金额))
            If dblMoney <> 0 Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If strTemp = Trim(.TextMatrix(i, .ColIndex("结算方式"))) Then
                        blnFind = True: Exit For
                    End If
                Next
                If Not blnFind Then
                    If MsgBox(" 在结算明细列表中不存在" & strTemp & "的结算方式 " & vbCrLf & _
                     "可能是因为并发原因造成的, 为了保证数据的一致性,你需要重新提取数据," & vbCrLf & _
                     "你是否要重新提取数据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                     Call LoadCollectData
                    End If
                     If vsRollingCurtain.Enabled And vsRollingCurtain.Visible Then vsRollingCurtain.SetFocus
                     Exit Function
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData(ByVal str轧帐IDs As String, ByRef strNO As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据保存
    '入参:str轧帐IDs-选中的轧帐IDs(多个用逗号分隔)
    '出参:strNO-保存成功后,返回的收款单据号
    '返回:保存成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-10 18:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, lngID As String, str结算信息 As String, str结算方式 As String
    Dim i As Long
    
    On Error GoTo errHandle
    With vsBalance
        For i = 1 To .Rows - 1
            str结算方式 = .TextMatrix(i, .ColIndex("结算方式"))
            If str结算方式 <> "" And Trim(.TextMatrix(i, .ColIndex("结算号码"))) <> "" Then
                str结算信息 = str结算信息 & "|" & str结算方式 & "," & Trim(.TextMatrix(i, .ColIndex("结算号码")))
            End If
        Next
    End With
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 2)
    If zlCommFun.ActualLen(str结算信息) > 4000 Then
        MsgBox "在结算明细信息中输入的结算号码与结算方式超长了,最多只能为4000个字符", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    lngID = zlDatabase.GetNextId("人员收缴记录")
    strNO = zlDatabase.GetNextNo(140)
    'Zl_财务收款记录_Insert
    strSQL = "Zl_财务收款记录_Insert("
    '  Id_In         In 人员收缴记录.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  No_In         In 人员收缴记录.No%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  收款员_In     In 人员收缴记录.收款员%Type,
    strSQL = strSQL & "'" & mstr缴款人 & "',"
    '  收款部门id_In In 人员收缴记录.收款部门id%Type,
    strSQL = strSQL & "Null,"
'    strSQL = strSQL & cboDept.ItemData(cboDept.ListIndex) & ","
    '  缴款组id_In   In 人员收缴记录.缴款组id%Type,
    strSQL = strSQL & "" & IIf(mlngGroupID = 0, "NULL", mlngGroupID) & ","
    '  暂存金额_In   In 人员收缴记录.冲预交款%Type,
    strSQL = strSQL & "" & IIf(mlngGroupID <> 0, "0", Val(Replace(lblRemainMoney.Caption, ",", ""))) & ","
    '  摘要_In       In 人员收缴记录.摘要%Type,
    strSQL = strSQL & "" & IIf(Trim(txtMemo.Text) = "", "NULL", "'" & Trim(txtMemo.Text) & "'") & ","
    '  登记人_In     In 人员收缴记录.登记人%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  登记时间_In   In 人员收缴记录.登记时间%Type,
    strSQL = strSQL & "sysdate,"
    '  轧帐ids_In    In Varchar2,轧帐IDs_In:轧帐ID1,轧帐ID2,...
    strSQL = strSQL & "'" & str轧帐IDs & "',"
    '  结算号码_In   In Varchar2:结算方式,结算号码|结算方式,结算号码,..
    strSQL = strSQL & "'" & str结算信息 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSelRollingCurtainIds(Optional ByRef strNotSelRollingCurtainIDs As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前选中的轧帐ID
    '出参:strNotSelRollingCurtainIDs-未选中的轧帐IDs(用逗号分离
    '返回:选中的轧帐ID
    '编制:刘兴洪
    '日期:2013-10-10 17:40:05
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim str轧帐IDs As String, lng轧帐ID As Long, i As Long
    On Error GoTo errHandle
   strNotSelRollingCurtainIDs = ""
    With vsRollingCurtain
        For i = 1 To .Rows - 1
            lng轧帐ID = Val(.TextMatrix(i, .ColIndex("ID")))
            If lng轧帐ID <> 0 And Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Then
                str轧帐IDs = str轧帐IDs & "," & lng轧帐ID
            ElseIf lng轧帐ID <> 0 Then
                strNotSelRollingCurtainIDs = strNotSelRollingCurtainIDs & "," & lng轧帐ID
            End If
        Next
    End With
    If strNotSelRollingCurtainIDs <> "" Then strNotSelRollingCurtainIDs = Mid(strNotSelRollingCurtainIDs, 2)
    If str轧帐IDs <> "" Then str轧帐IDs = Mid(str轧帐IDs, 2)
    GetSelRollingCurtainIds = str轧帐IDs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub BillPrint(ByVal strNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:收款收据打印
    '编制:刘兴洪
    '日期:2013-09-11 11:55:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
    blnPrint = False
    If Not zlStr.IsHavePrivs(mstrPrivs, "收款收据打印") Then Exit Sub
    Select Case Val(zlDatabase.GetPara("收款收据打印方式", glngSys, mlngModule))     '使用医生站的相关参数
    Case 0    '不打印
        Exit Sub
    Case 1    '自助动打印
        blnPrint = True
    Case 2    '选择打印
        If MsgBox("你是否要打印缴款收据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            blnPrint = True
        End If
    End Select
    If blnPrint = False Then Exit Sub
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1500", Me, "NO=" & strNO, "记录性质=4", 2)
End Sub

Private Sub imgColPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgPlan.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsRollingCurtain, lngLeft, lngTop, imgColPlan.Height)
    zl_vsGrid_Para_Save mlngModule, vsRollingCurtain, Me.Name, "轧帐信息列表", False, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub picImgPlan_Click()
    Call imgColPlan_Click
End Sub
