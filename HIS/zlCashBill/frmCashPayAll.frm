VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmCashPayAll 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增缴款记录"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCashPayAll.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtGroups 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   360
      Left            =   6810
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   735
      Width           =   2490
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   9360
      TabIndex        =   17
      Top             =   8805
      Width           =   9360
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   9330
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印设置(&S)"
         Height          =   420
         Left            =   1650
         TabIndex        =   18
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   7755
         TabIndex        =   14
         Top             =   120
         Width           =   1530
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   6225
         TabIndex        =   13
         Top             =   120
         Width           =   1530
      End
   End
   Begin VB.ComboBox cboTimes 
      Height          =   360
      Left            =   5190
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1257
      Width           =   1500
   End
   Begin VB.ComboBox cboType 
      Height          =   360
      Left            =   3880
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1257
      Width           =   1335
   End
   Begin VB.ComboBox cbo缴款部门 
      Height          =   360
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   750
      Width           =   2250
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   420
      Left            =   6810
      TabIndex        =   8
      ToolTipText     =   "热键：F5"
      Top             =   1227
      Width           =   1530
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   750
      Width           =   1530
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1250
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   168951811
      CurrentDate     =   36904
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5490
      Left            =   80
      TabIndex        =   21
      Top             =   3240
      Width           =   10005
      Begin VB.TextBox txtRquareEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1620
      End
      Begin VB.TextBox txtRquareEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1635
      End
      Begin VB.TextBox txtLoanEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3615
         Width           =   1635
      End
      Begin VB.TextBox txtLoanEdit 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3615
         Width           =   1620
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   2
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4545
         Width           =   8115
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   30
         Left            =   -120
         TabIndex        =   25
         Top             =   120
         Width           =   9390
      End
      Begin VB.TextBox txtItem 
         Height          =   360
         Index           =   3
         Left            =   1095
         MaxLength       =   50
         TabIndex        =   12
         Top             =   4950
         Width           =   5010
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   4
         Left            =   7215
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4950
         Width           =   1995
      End
      Begin VB.TextBox txtItem 
         BackColor       =   &H8000000F&
         Height          =   360
         Index           =   1
         Left            =   1095
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3165
         Width           =   4980
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshIncome 
         Height          =   3855
         Left            =   6180
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   600
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   6800
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483643
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin ZL9BillEdit.BillEdit mshCash 
         Height          =   2490
         Left            =   30
         TabIndex        =   10
         Top             =   600
         Width           =   6060
         _ExtentX        =   10689
         _ExtentY        =   4392
         Enabled         =   -1  'True
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   3
         RowHeight0      =   360
         RowHeightMin    =   300
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lblLoan 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "发消费卡"
         Height          =   240
         Index           =   3
         Left            =   45
         TabIndex        =   36
         Top             =   4110
         Width           =   960
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "充值"
         Height          =   240
         Index           =   2
         Left            =   3870
         TabIndex        =   38
         Top             =   4110
         Width           =   480
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "借出"
         Height          =   240
         Index           =   1
         Left            =   3870
         TabIndex        =   34
         Top             =   3675
         Width           =   480
      End
      Begin VB.Label lblLoan 
         AutoSize        =   -1  'True
         Caption         =   "借款"
         Height          =   240
         Index           =   0
         Left            =   540
         TabIndex        =   32
         Top             =   3675
         Width           =   480
      End
      Begin VB.Label lblItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "缴款合计"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   31
         Top             =   4605
         Width           =   960
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结算明细："
         Height          =   240
         Index           =   2
         Left            =   75
         TabIndex        =   30
         Tag             =   "结算明细："
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "摘要(&D)"
         Height          =   240
         Index           =   6
         Left            =   210
         TabIndex        =   11
         Top             =   5010
         Width           =   840
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "经手人"
         Height          =   240
         Index           =   7
         Left            =   6450
         TabIndex        =   29
         Top             =   5010
         Width           =   720
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "预交款"
         Height          =   240
         Index           =   4
         Left            =   330
         TabIndex        =   28
         Top             =   3225
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收入明细："
         Height          =   240
         Index           =   3
         Left            =   6225
         TabIndex        =   27
         Tag             =   "收入明细："
         Top             =   285
         Width           =   1200
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshRec 
      Height          =   1560
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1680
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   2752
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   0
      MergeCells      =   3
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.Label lblGroups 
      AutoSize        =   -1  'True
      Caption         =   "人员分组"
      Height          =   240
      Left            =   5820
      TabIndex        =   40
      Top             =   780
      Width           =   960
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "部门"
      Height          =   240
      Index           =   8
      Left            =   2880
      TabIndex        =   16
      Top             =   810
      Width           =   480
   End
   Begin VB.Label lblTimePeriod 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴款登记卡"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   2707
      TabIndex        =   0
      Top             =   195
      Width           =   2250
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "缴款人"
      Height          =   240
      Index           =   0
      Left            =   390
      TabIndex        =   1
      Top             =   810
      Width           =   720
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "截止时间"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   1317
      Width           =   960
   End
End
Attribute VB_Name = "frmCashPayAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gPayMoneyEdit
    PM_全额缴款 = 0
    PM_按日缴款 = 1
End Enum
Private mEditType As gPayMoneyEdit   '0-全额缴款;1-按日缴款

Private mblnOK  As Boolean
Private mstr缴款人 As String, mlng缴款人ID As Long
Private mstrLast As String '记录本次提取成功的截止时间
Private mrsDetail As ADODB.Recordset
Private mrsTimes As ADODB.Recordset '当前缴款日期的选择的收费类型的缴款次数
Private mlng组ID As Long    '-1时,不分组:暂不用
Private Const CONDFormat = "yyyy-MM-dd HH:mm:ss"
Private Const CONRecHead = "类别|850|1,次数|600|1,开始时间|2600|4,终止时间|2600|4"

Public Function ShowMe(ByVal strUser As String, ByVal lngUserID As Long, frmParent As Object, _
    Optional ByVal EditType As gPayMoneyEdit = PM_全额缴款) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口(显示或缴款)
    '入参:strUser-缴款人员
    '       lngUserID-缴款人员ID
    '       frmParent-调用的主窗体
    '       EditType-调用功能(0-全额缴款;1-按日缴款)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-29 14:16:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType
    mstr缴款人 = strUser: mlng缴款人ID = lngUserID
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function
Private Sub zlSetDefaultDate()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省时间
    '返回:
    '编制:刘兴洪
    '日期:2009-10-14 14:26:23
    '问题号:25752
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mEditType = PM_全额缴款 Then
        ''全额缴款
        strSQL = "Select Max(终止时间) as 终止时间 From  收费清点记录 Where 收款员=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人)
        If IsNull(rsTemp!终止时间) Then
            dtpDate.Value = zlDatabase.Currentdate
        Else
            dtpDate.Value = CDate(Format(rsTemp!终止时间, CONDFormat))
        End If
    Else
        dtpDate.Value = CDate(Format(zlDatabase.Currentdate, dtpDate.CustomFormat))
        Call dtpDate_Change
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitFace()
'功能：初始化界面显示
    Set mrsDetail = New ADODB.Recordset
    mstrLast = "" '标记为空
    
    If Not Visible Then
        If mEditType = PM_全额缴款 Then '全额缴款
            lblTimePeriod.Visible = False
            cboType.Visible = False
            cboTimes.Visible = False
            
            lblItem(1).Caption = "截止时间"
            dtpDate.CustomFormat = CONDFormat
            cmdRefresh.Left = dtpDate.Left + dtpDate.Width + 100
            
            lblTimePeriod.Visible = True
            mshRec.Visible = False
            fraMain.Top = lblTimePeriod.Top + lblTimePeriod.Height
            Me.Height = Me.Height - (mshRec.Height - lblTimePeriod.Height)
            
        Else '按日缴款
            lblTimePeriod.Visible = True
            cboType.Visible = True
            cboTimes.Visible = True
            
            lblItem(1).Caption = "缴款日期"
            dtpDate.CustomFormat = "yyyy-MM-dd 00:00:00"
            dtpDate.Width = txtItem(0).Width
            cboType.Left = dtpDate.Left + dtpDate.Width + 100
            cboTimes.Left = cboType.Left + cboType.Width + 100
            cmdRefresh.Left = cboTimes.Left + cboTimes.Width + 100
            
            lblTimePeriod.Visible = False
            mshRec.Visible = True
        End If
    End If
    
    lblItem(2).Caption = lblItem(2).Tag
    lblItem(3).Caption = lblItem(3).Tag
    txtItem(1).Text = ""
    txtItem(2).Text = ""
    
    With mshCash
        .AllowAddRow = False
        .Font.Size = 12
        .TxtEditFont.Size = 12
        .Active = False
        
        .ClearBill
        .RowHeight(0) = .RowHeightMin
        .TextMatrix(0, 0) = "结算方式"
        .TextMatrix(0, 1) = "金额"
        .TextMatrix(0, 2) = "结算号"
        
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        
        .MsfObj.ColAlignmentFixed(0) = 4
        .MsfObj.ColAlignmentFixed(1) = 4
        .MsfObj.ColAlignmentFixed(2) = 4
        
        .ColWidth(0) = 1300
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        
        .ColData(2) = 4
    End With
    
    With mshIncome
        .Clear
        .Rows = 2
        .TextMatrix(0, 0) = "收入项目"
        .TextMatrix(0, 1) = "金额"
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColWidth(0) = 1350
        .ColWidth(1) = 1350
    End With
End Sub

Private Sub cboType_Click()
    Call LoadCashTimes(mstr缴款人, dtpDate.Value, Val(cboType.ItemData(cboType.ListIndex)))
End Sub

Private Sub cmdRefresh_Click()
'功能：重新提取当前缴款人员的缴款明细
    Dim rsTmp As New ADODB.Recordset, strIF As String
    Dim strSQL As String, strSub As String, i As Long, bytFlag As Byte
    Dim datBegin As Date, datEnd As Date, strTable As String, strWhere As String
    Dim cur缴款合计 As Currency, cur结算合计 As Currency
    Dim cur收入合计 As Currency, cur预交合计 As Currency
    Dim dbl借款 As Double, dbl借出 As Double
    Dim dbl卡面销售 As Double, dbl退卡额 As Double, dbl充值额 As Double
    
    If dtpDate.Value > zlDatabase.Currentdate Then
        MsgBox "缴款截止时间不应越过当前系统时间。", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    
    Call InitFace
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    If mEditType = PM_按日缴款 Then
        bytFlag = Val(cboType.ItemData(cboType.ListIndex))  '0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡',6,'消费卡')
        If bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5 Then strIF = " And 记录性质 = " & IIf(bytFlag = 3, 1, bytFlag)
    End If
    
    '问题:42376:执行状态<>9
    '获取缴款员上次缴款截止时间
    If mEditType = PM_按日缴款 Then
        '按日缴款
        If cboTimes.ListIndex = 0 Then
            strSQL = "Select Min(开始时间) 开始时间,Max(终止时间) 终止时间 From 收费清点记录 Where 收款员=[1] And 日期=[2]"
            If cboType.ListIndex > 0 Then strSQL = strSQL & " And 性质 = [3]"
            
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, dtpDate.Value, bytFlag)
            If rsTmp.RecordCount > 0 Then
                If Not IsNull(rsTmp!开始时间) Then datBegin = DateAdd("s", -1, rsTmp!开始时间)
                If Not IsNull(rsTmp!终止时间) Then datEnd = rsTmp!终止时间
            End If
            If datBegin = CDate(0) Then datBegin = CDate(Format(DateAdd("d", -1, dtpDate.Value), "yyyy-MM-dd ") & "23:59:59")
            If datEnd = CDate(0) Then datEnd = CDate(Format(dtpDate.Value, "yyyy-MM-dd ") & "23:59:59")
        Else    '必定有cboType.ListIndex > 0
            mrsTimes.Filter = "次数=" & cboTimes.ListIndex
            datBegin = mrsTimes!开始时间
            datEnd = mrsTimes!终止时间
        End If
        If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
            MsgBox "上次缴款时间:" & Format(datBegin, CONDFormat) & vbCrLf & "处于最近一次历史数据转出之前,要缴款的部门数据已转入后备! " & _
                vbCrLf & "请与系统管理员联系转入数据,或者首次缴款采用手工缴款方式!", vbInformation, gstrSysName
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        
    Else
        '全额缴款
        strSQL = "Select Max(截止时间) as 截止时间 From 人员缴款记录 Where 收款员=[1] And 截止时间 is Not NULL"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人)
        If rsTmp.RecordCount > 0 Then
            If Not IsNull(rsTmp!截止时间) Then datBegin = rsTmp!截止时间
        End If
        If datBegin = CDate(0) Then
            datBegin = CDate(Format("1990-01-01 00:00:00", CONDFormat))
            
            If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
                If MsgBox("当前收款员未缴过款,如果在上次历史数据转移之前存在收款数据,当前数据可能不完整。你确定要继续吗？", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Else
            If zlDatabase.DateMoved(Format(datBegin, CONDFormat), , , Me.Caption) Then
                MsgBox "上次缴款时间:" & Format(datBegin, CONDFormat) & vbCrLf & "处于最近一次历史数据转出之前,要缴款的部门数据已转入后备表! " & _
                    vbCrLf & "请与系统管理员联系转入数据,或者首次缴款采用手工缴款方式!", vbInformation, gstrSysName
                Screen.MousePointer = 0
                Exit Sub
            End If
        End If
        datEnd = dtpDate.Value
        lblTimePeriod.Caption = "时间范围:" & Format(datBegin, CONDFormat) & "~" & Format(datEnd, CONDFormat)
    End If
    
        
    
    '获取该缴款员本次各种结算方式的应缴金额
    '-----------------------------------------------------------------------------------------------
    '收费部份：收费、挂号、发卡、收费冲预交(单独显示)
    strSQL = ""
    If mEditType = PM_全额缴款 Or _
        mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
           strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.结帐ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=1"
        '0-全额缴款;1-按日缴款
        'bytFlag: 0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡')
       If bytFlag = 5 And mEditType = PM_按日缴款 Then
            '按日缴款且为就只统计就诊卡
            strTable = "" & _
            " Select Distinct 结帐ID" & _
            " From 住院费用记录 A" & _
            " Where Nvl(记帐费用,0)=0 And 记录状态<>0 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") "
        ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_按日缴款 Then
            '按日缴款,但非就诊卡和全部缴款
            strTable = "" & _
            " Select Distinct 结帐ID" & _
            " From 门诊费用记录 A" & _
            " Where Nvl(记帐费用,0)=0  and nvl(费用状态,0)<>1 And 记录状态<>0 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
            "           And Not Exists(" & strSub & ") "
        Else
            strTable = "" & _
            " Select Distinct 结帐ID" & _
            " From 住院费用记录 A" & _
            " Where Nvl(记帐费用,0)=0 And 记录状态<>0 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") " & _
            " Union  " & _
            " Select Distinct 结帐ID" & _
            " From 门诊费用记录 A" & _
            " Where Nvl(记帐费用,0)=0 And 记录状态<>0 and nvl(费用状态,0)<>1 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
            "        And Not Exists(" & strSub & ") "
        End If
        
        strSQL = _
        " Select Decode(Mod(B.记录性质,10),1,'[冲预交款]',B.结算方式) as 结算方式,Sum(B.冲预交) as 金额" & _
        " From ( " & strTable & ") A,病人预交记录 B" & _
        " Where A.结帐ID=B.结帐ID And nvl(B.校对标志,0) =0" & _
        " Group by Decode(Mod(B.记录性质,10),1,'[冲预交款]',B.结算方式)"
    End If
        
    '结帐部份：结帐补款、结帐冲预交(单独显示)
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=2"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Decode(Mod(B.记录性质,10),1,'[冲预交款]',B.结算方式) as 结算方式,Sum(B.冲预交) as 金额" & _
            " From 病人结帐记录 A,病人预交记录 B" & _
            " Where A.ID=B.结帐ID And A.结算状态 Is Null And A.操作员姓名||''=[1] And A.收费时间>[2] And A.收费时间<=[3] And Not Exists(" & strSub & ")" & _
            " Group by Decode(Mod(B.记录性质,10),1,'[冲预交款]',B.结算方式)"
    End If
    
    '收预交部份：直接收预交
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=3"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            "   Select 结算方式,Sum(金额) as 金额" & _
            "   From 病人预交记录 A" & _
            "   Where 记录性质=1 And 操作员姓名||''=[1] And 收款时间>[2] And 收款时间<=[3] And Not Exists(" & strSub & ") " & _
            "   Group by 结算方式"
    End If
    
    '消费卡:
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 6) Then
        '0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡',6,'消费卡')
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=6"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.结算方式,Sum(A.实收金额) as 金额" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (1, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And Not Exists(" & strSub & ") " & _
            " Group by 结算方式"
        
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=5"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.结算方式, Sum(A.实收金额) as 金额" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (2, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And Not Exists(" & strSub & ") " & _
            " Group by 结算方式"
    End If
    
 
     
    '借款部分统计
    '-----------------------------------------------------------------------------------------------
    '直接扣减现金
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=4"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
             "Select A.结算方式, Sum(nvl(a.借款金额,0)) as 金额" & _
            " From 人员借款记录 A" & _
            " Where  A.借款人||''=[1] And A.借出时间>[2] And A.借出时间<=[3] And A.取消时间 is NULL And Not Exists(" & strSub & ")" & _
            " Group by A.结算方式"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
             "Select A.结算方式,-1* Sum(nvl(a.借款金额,0)) as 金额" & _
            " From 人员借款记录 A" & _
            " Where  A.借出人||''=[1] And A.借出时间>[2] And A.借出时间<=[3] And A.取消时间 is NULL And Not Exists(" & strSub & ")" & _
            " Group by A.结算方式"
    End If
    
    strSQL = "Select 结算方式,Sum(金额) as 金额 From (" & strSQL & ") Group by 结算方式 Having Sum(金额)<>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
    If Not rsTmp.EOF Then
        With mshCash
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Nvl(rsTmp!结算方式)
                .TextMatrix(i, 1) = Format(Nvl(rsTmp!金额, 0), "0.00")
                cur结算合计 = cur结算合计 + Nvl(rsTmp!金额, 0)
                If Nvl(rsTmp!结算方式) <> "[冲预交款]" Then
                    cur缴款合计 = cur缴款合计 + Nvl(rsTmp!金额, 0)
                End If
                rsTmp.MoveNext
            Next
        End With
    End If
  
    
    '获取该缴款员本次缴款对应的收入金额
    '-----------------------------------------------------------------------------------------------
    '收费部份：收费,挂号,发卡,收费冲预交
    strSQL = ""
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.结帐ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=1"
        strWhere = " Where Nvl(记帐费用,0)=0  and nvl(执行状态,0)<>9 And 记录状态<>0 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
        "        And Not Exists(" & strSub & ")  "
        
        '0-全额缴款;1-按日缴款
        'bytFlag: 0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡')
        If bytFlag = 5 And mEditType = PM_按日缴款 Then
             '按日缴款且为就只统计就诊卡
             strTable = "住院费用记录"
         ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_按日缴款 Then
             '按日缴款,但非就诊卡和全部缴款
             strTable = "门诊费用记录"
         Else
            strTable = " ( " & _
             " Select 收入项目ID,Sum(结帐金额) as 结帐金额" & _
             " From 住院费用记录 A" & _
             " Where Nvl(记帐费用,0)=0 And 记录状态<>0 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
             "        And Not Exists(" & strSub & ")  " & _
             " Group by 收入项目ID  " & _
             " Union ALL " & _
             " Select 收入项目ID,Sum(结帐金额) as 结帐金额" & _
             " From 门诊费用记录 A" & _
             " Where Nvl(记帐费用,0)=0 And 记录状态<>0 and nvl(A.费用状态,0)<>1 And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3]" & strIF & _
             "        And Not Exists(" & strSub & ")  " & _
             " Group by 收入项目ID ) "
             strWhere = ""
         End If
        strSQL = _
        " Select 收入项目ID,Sum(结帐金额) as 金额" & _
        " From " & strTable & " A" & _
                 strWhere & _
        " Group by 收入项目ID"
    End If
        
    '结帐部份：结帐补款,结帐冲预交
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=2"
        
        'bytFlag: 0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡')
        '可能存在门诊结帐的情况,因此，必需全关联
        strTable = "" & _
        " Select B.收入项目ID,Sum(B.结帐金额) as 结帐金额" & _
        " From 病人结帐记录 A,门诊费用记录 B" & _
        " Where B.记帐费用=1 And A.结算状态 Is Null And A.ID=B.结帐ID And A.操作员姓名||''=[1] And A.收费时间>[2] And A.收费时间<=[3]" & _
        "       And Not Exists(" & strSub & ") " & _
        " Group by B.收入项目ID " & _
        " Union ALL  " & _
        " Select B.收入项目ID,Sum(B.结帐金额) as 结帐金额" & _
        " From 病人结帐记录 A,住院费用记录 B" & _
        " Where B.记帐费用=1 And A.结算状态 Is Null And A.ID=B.结帐ID And A.操作员姓名||''=[1] And A.收费时间>[2] And A.收费时间<=[3]" & _
        "       And Not Exists(" & strSub & ") " & _
        " Group by B.收入项目ID"
        
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select A.收入项目ID,Sum(A.结帐金额) as 金额" & _
            " From  (" & strTable & ") A" & _
            " Group by A.收入项目ID"
    End If
    
    If strSQL <> "" Then
        strSQL = "Select B.编码,B.名称,Sum(A.金额) as 金额 From (" & strSQL & ") A,收入项目 B" & _
            " Where A.收入项目ID=B.ID Group by B.编码,B.名称 Having Sum(A.金额)<>0 Order by B.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then
            With mshIncome
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, 0) = Nvl(rsTmp!名称)
                    .TextMatrix(i, 1) = Format(Nvl(rsTmp!金额, 0), "0.00")
                    cur收入合计 = cur收入合计 + Nvl(rsTmp!金额, 0)
                    rsTmp.MoveNext
                Next
            End With
        End If
    End If
    
    '获取该缴款员本次缴款对应的预交收款
    '-----------------------------------------------------------------------------------------------
    '收预交部份：直接收预交
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=3"
        strSQL = "Select Sum(金额) as 金额" & _
            " From 病人预交记录 A" & _
            " Where 记录性质=1 And 操作员姓名||''=[1] And 收款时间>[2] And 收款时间<=[3] And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then cur预交合计 = Nvl(rsTmp!金额, 0)
    End If
    
   '消费卡:
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 6) Then
        '0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡',6,'消费卡')
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=6"
        strSQL = "" & _
            " Select  Sum(A.实收金额) as 金额" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (1, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And Not Exists(" & strSub & ") "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl卡面销售 = Nvl(rsTmp!金额, 0)
        
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=5"
        strSQL = "" & _
            " Select  Sum(A.实收金额) as 金额" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (2, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And Not Exists(" & strSub & ") "
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl充值额 = Nvl(rsTmp!金额, 0)
    End If
        
        
    '借款部分统计
    '-----------------------------------------------------------------------------------------------
    '直接扣减现金
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=4"
        strSQL = "Select Sum(借款金额) as 金额" & _
            " From 人员借款记录 A" & _
            " Where  借款人||''=[1] And 借出时间>[2] And 借出时间<=[3] And 取消时间 is NULL And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl借款 = Nvl(rsTmp!金额, 0)
        
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=4"
        strSQL = "Select Sum(借款金额) as 金额" & _
            " From 人员借款记录 A" & _
            " Where  借出人||''=[1] And 借出时间>[2] And 借出时间<=[3] And 取消时间 is NULL And Not Exists(" & strSub & ")"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
        If Not rsTmp.EOF Then dbl借出 = Nvl(rsTmp!金额, 0)
    End If
    
    
    
    '获取该缴款员本次缴款对照
    '-----------------------------------------------------------------------------------------------
    '收费部份：收费,挂号,发卡,收费冲预交
    strSQL = ""
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 3 Or bytFlag = 4 Or bytFlag = 5) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.结帐ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=1"
        '0-全额缴款;1-按日缴款
        '问题:44344
        'bytFlag: 0-全部,Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡')
        strWhere = " where Nvl(记帐费用,0)=0 And 记录状态<>0 and nvl(执行状态,0)<>9   And 操作员姓名||''=[1] And 登记时间>[2] And 登记时间<=[3] And Not Exists(" & strSub & ")" & strIF
        If bytFlag = 5 And mEditType = PM_按日缴款 Then
             '按日缴款且为就只统计就诊卡
             strTable = "住院费用记录"
         ElseIf InStr(1, "05", bytFlag) = 0 And mEditType = PM_按日缴款 Then
             '按日缴款,但非就诊卡和全部缴款
             strTable = "门诊费用记录"
         Else
            strTable = " ( " & _
             " Select Distinct 结帐ID" & _
             " From 住院费用记录 A" & _
               strWhere & _
             "  " & _
             " Union ALL " & _
             " Select Distinct 结帐ID " & _
             " From 门诊费用记录 A" & _
               strWhere & _
             " ) "
             strWhere = ""
         End If
        
        strSQL = _
            " Select Distinct 1 as 性质,结帐ID as 记录ID" & _
            " From " & strTable & " A" & _
            "  " & strWhere
    End If
        
    '结帐部份：结帐补款,结帐冲预交
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 2) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=2"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 2 as 性质,ID as 记录ID" & _
            " From 病人结帐记录 A" & _
            " Where 操作员姓名||''=[1] And 收费时间>[2] And 收费时间<=[3] And 结算状态 Is Null And Not Exists(" & strSub & ")"
    End If
    
    '收预交部份：直接收预交
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 1) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=3"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 3 as 性质,ID as 记录ID" & _
            " From 病人预交记录 A" & _
            " Where 记录性质=1 And 操作员姓名||''=[1] And 收款时间>[2] And 收款时间<=[3] And Not Exists(" & strSub & ")"
    End If
    
    '-消费卡
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 And (bytFlag = 0 Or bytFlag = 6) Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=6"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 6 as 性质,ID as 记录ID" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (1, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And Not Exists(" & strSub & ")"
        
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=5"
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 5 as 性质,ID as 记录ID" & _
            " From 病人卡结算记录 A, 病人卡结算记录 B" & _
            " Where a.交易序号 = b.交易序号(+) And a.记录性质 In (2, 3) And b.记录性质(+) = 3 And a.操作员姓名||''=[1] And a.登记时间>[2] And a.登记时间<=[3] And   Not Exists(" & strSub & ")  "
        
    End If
 
    '获取借款记录部分:0-全额缴款;1-按日缴款
    '   根据按日缴款的上班时间，或者全额缴款的截止时间，统计相应时间范围内的"人员借款记录"， _
    '以没有取消的借出时间为准，并以现金结算方式汇总作为缴款结算
    If mEditType = PM_全额缴款 Or mEditType = PM_按日缴款 Then
        strSub = "Select Y.记录ID From 人员缴款记录 X,人员缴款对照 Y Where Y.记录ID=A.ID And X.收款员=[1] And X.ID=Y.单据ID And Y.性质=4"
        '借款
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 4 as 性质,ID as 记录ID" & _
            " From  人员借款记录 A" & _
            " Where 借款人||''=[1] And 借出时间>[2] And 借出时间<=[3] and 取消时间 is NULL And Not Exists(" & strSub & ")"
        
        '借出
        strSQL = strSQL & IIf(strSQL = "", "", " Union ALL ") & _
            " Select Distinct 4 as 性质,ID as 记录ID" & _
            " From  人员借款记录 A" & _
            " Where 借出人||''=[1] And 借出时间>[2] And 借出时间<=[3] And 取消时间 is NULL And Not Exists(" & strSub & ")"
    End If
    
    strSQL = "Select /*+ Rule*/ 性质,记录ID From (" & strSQL & ")"
    Set mrsDetail = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr缴款人, datBegin, datEnd)
    
    '显示合计信息
    '-----------------------------------------------------------------------------------------------
    lblItem(2).Caption = lblItem(2).Tag & Format(cur结算合计, "0.00")
    lblItem(3).Caption = lblItem(3).Tag & Format(cur收入合计, "0.00")
    txtItem(1).Text = Format(cur预交合计, "0.00")
    
    txtLoanEdit(0).Text = Format(dbl借款, "0.00")
    txtLoanEdit(1).Text = Format(dbl借出, "0.00")
    
    txtRquareEdit(0).Text = Format(dbl卡面销售 - dbl退卡额, "0.00")
    txtRquareEdit(1).Text = Format(dbl充值额, "0.00")
    
    If cur缴款合计 <> 0 Then
        txtItem(2).Text = Format(cur缴款合计, "0.00元") & " （" & zlCommFun.UppeMoney(cur缴款合计) & "）"
    Else
        txtItem(2).Text = Format(cur缴款合计, "0.00元")
    End If
    
    '标记成功
    mshCash.Active = True
    mshCash.Row = 1: mshCash.Col = 2
    mstrLast = "To_Date('" & Format(datEnd, CONDFormat) & "','YYYY-MM-DD HH24:MI:SS')"
    Screen.MousePointer = 0
    mshCash.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL() As Variant, i As Long, k As Long
    Dim strDate As String, lng单据ID As Long, lngID As Long, blnTrans As Boolean
    Dim dblSumMoney As Double
    
    If InStr(txtItem(3).Text, "'") > 0 Then
        MsgBox "摘要信息中包含非法的字符。", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txtItem(3).Text) > txtItem(3).MaxLength Then
        MsgBox "摘要信息中包含的内容太多，最多允许 " & txtItem(3).MaxLength \ 2 & " 个汉字或 " & txtItem(3).MaxLength & " 个字符。", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    
    With mshCash
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 And .TextMatrix(i, 0) <> "[冲预交款]" Then
                If InStr(.TextMatrix(i, 2), "'") > 0 Then
                    MsgBox "结算方式 [" & .TextMatrix(i, 0) & "] 的结算号中包含非法字符。", vbInformation, gstrSysName
                    .Row = i: .Col = 2: .SetFocus: Exit Sub
                ElseIf zlCommFun.ActualLen(.TextMatrix(i, 2)) > 10 Then
                    MsgBox "结算方式 [" & .TextMatrix(i, 0) & "] 的结算号过长，最多允许10个字符。", vbInformation, gstrSysName
                    .Row = i: .Col = 2: .SetFocus: Exit Sub
                End If
                dblSumMoney = dblSumMoney + Val(.TextMatrix(i, 1))
'                If Val(.TextMatrix(i, 1)) < 0 Then
'                    MsgBox "结算方式 [" & .TextMatrix(i, 0) & "] 的金额不能为负数。", vbInformation, gstrSysName
'                    .Row = i: .Col = 1: .SetFocus: Exit Sub
'                End If
                k = k + 1
            End If
        Next
    End With
    
    '刘兴洪 问题:????吉林省人民医院   日期:2010-12-06 11:09:29
    '       实际情况中是总额也会出现负数的情况的，比如退了一笔支票，收的钱没有退的多的时候。
    '    '刘兴洪:25694,
    '    If dblSumMoney < 0 Then
    '        MsgBox "结算总额[" & Format(dblSumMoney, "####0.00;-####0.00;0;0") & "]不能为负数,请检查。", vbInformation, gstrSysName
    '        mshCash.SetFocus: Exit Sub
    '    End If

    
    If k = 0 Or mstrLast = "" Or mrsDetail.State = 0 Then
        MsgBox "没有提取有效的缴款金额！", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    If mrsDetail.RecordCount = 0 Then
        MsgBox "没有提取有效的缴款金额！", vbInformation, gstrSysName
        dtpDate.SetFocus: Exit Sub
    End If
    
    '产生SQL语句
    arrSQL = Array()
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, CONDFormat) & "','YYYY-MM-DD HH24:MI:SS')"
    lng单据ID = zlDatabase.GetNextId("人员缴款记录")
    With mshCash
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 And .TextMatrix(i, 0) <> "[冲预交款]" Then
                If lngID = 0 Then
                    lngID = lng单据ID
                Else
                    lngID = zlDatabase.GetNextId("人员缴款记录")
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_人员缴款记录_Insert(" & lngID & "," & lng单据ID & "," & strDate & "," & _
                    "'" & mstr缴款人 & "','" & UserInfo.姓名 & "','" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 1)) & "," & _
                    "'" & .TextMatrix(i, 2) & "','" & txtItem(3).Text & "'," & mstrLast & "," & cbo缴款部门.ItemData(cbo缴款部门.ListIndex) & ")"
            End If
        Next
    End With
    If mrsDetail.RecordCount <> 0 Then mrsDetail.MoveFirst
    For i = 1 To mrsDetail.RecordCount
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_人员缴款对照_Insert(" & lng单据ID & "," & mrsDetail!性质 & "," & mrsDetail!记录ID & ")"
        mrsDetail.MoveNext
    Next
    
    '保存缴款记录
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    '打印票据
    Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me, "单据ID=" & lng单据ID, 2)
    
    Screen.MousePointer = 0
    mstrLast = "" '标记为可以关闭
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdPrint_Click()
    ReportPrintSet gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1500", Me
End Sub

Private Sub dtpDate_Change()
    If mEditType = PM_按日缴款 Then
        Call LoadCashType(mstr缴款人, dtpDate.Value)
        Call LoadRec(mstr缴款人, dtpDate.Value)
    End If
End Sub

Private Sub LoadCashType(ByVal strOperator As String, ByVal datThis As Date)
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
 
    strSQL = "Select Distinct 性质, Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡',6, '消费卡') 性质说明" & vbNewLine & _
            "From 收费清点记录" & vbNewLine & _
            "Where 收款员 = [1] And 日期 = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis)
    With cboType
        .Clear
        .AddItem "全部类别"
        .ItemData(.NewIndex) = 0
        .ListIndex = 0  '触发click事件
        For i = 1 To rsTmp.RecordCount
            .AddItem rsTmp!性质说明
            .ItemData(.NewIndex) = rsTmp!性质
            rsTmp.MoveNext
        Next
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRec(ByVal strOperator As String, ByVal datThis As Date)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    With mshRec
        .Redraw = False
        Call zlControl.MshSetFormat(mshRec, CONRecHead, Me.Caption, , , True)
        
        strSQL = "Select Decode(性质, 1, '预交款', 2, '结帐', 3, '收费', 4, '挂号', 5, '就诊卡',6,'消费卡') 性质," & vbNewLine & _
                "       Row_Number() Over(Partition By 性质 Order By 开始时间) 次数," & vbNewLine & _
                "       开始时间, 终止时间" & vbNewLine & _
                "From 收费清点记录" & vbNewLine & _
                "Where 收款员 = [1] And 日期 = [2]" & vbNewLine & _
                "Order By 性质, 次数"

        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis)
        If rsTmp.RecordCount > 0 Then
            .Rows = .FixedRows + rsTmp.RecordCount
            .MergeCol(0) = True
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = rsTmp!性质
                .TextMatrix(i, 1) = rsTmp!次数
                .TextMatrix(i, 2) = Format(rsTmp!开始时间, CONDFormat)
                .TextMatrix(i, 3) = Format(rsTmp!终止时间, CONDFormat)
                
                rsTmp.MoveNext
            Next
        Else
            .Rows = .FixedRows + 1
        End If
        
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCashTimes(ByVal strOperator As String, ByVal datThis As Date, ByVal bytFlag As Byte)
    Dim strSQL As String, i As Long
    
    With cboTimes
        .Clear
        .AddItem "全部次数"
        .ListIndex = .NewIndex  '触发click事件
        
        If bytFlag <> 0 Then
            strSQL = "Select Rownum 次数, 开始时间,终止时间 From " & _
            "(Select 开始时间,终止时间 From 收费清点记录 Where 收款员=[1] And 日期 = [2] And 性质 = [3] Order By 开始时间)"
        
            On Error GoTo errH
            Set mrsTimes = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strOperator, datThis, bytFlag)
            For i = 1 To mrsTimes.RecordCount
                .AddItem "第" & i & "次缴款"
                mrsTimes.MoveNext
            Next
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyF5 Then
        Call cmdRefresh_Click
    ElseIf KeyCode = 13 Then
        If Not ActiveControl Is mshCash Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    
    mblnOK = False
    
    txtItem(0).Text = mstr缴款人
    Set rsTmp = GetPersonnelDept(mlng缴款人ID)
    Call zlControl.CboAddData(cbo缴款部门, rsTmp, True)
    If cbo缴款部门.ListCount > 0 Then cbo缴款部门.ListIndex = 0
    txtItem(4).Text = UserInfo.姓名
    
    Call InitFace
    Call zlSetDefaultDate
    Call LoadGroups
   
End Sub
Private Sub LoadGroups()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载指定人员分组信息
    '编制:刘兴洪
    '日期:2010-11-29 14:30:43
    '问题:33633
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    '问题:33633
    gstrSQL = "" & _
    "   Select A.组名称,A.ID From 财务缴款分组 A ,缴款成员组成 B " & _
    "   Where A.ID=B.组ID And B.成员ID=[1] and A.删除日期>=sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng缴款人ID)
    txtGroups.Text = ""
    mlng组ID = -1
    txtGroups.Visible = True: lblGroups.Visible = True
    If Not rsTemp.EOF Then
        txtGroups.Text = Nvl(rsTemp!组名称)
        mlng组ID = Val(Nvl(rsTemp!ID))
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mstrLast <> "" Then
        If MsgBox("确实要放弃缴款吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    mstr缴款人 = ""
    mstrLast = ""
    Set mrsDetail = Nothing
End Sub

Private Sub mshCash_EnterCell(Row As Long, Col As Long)
    If mshCash.TextMatrix(Row, 0) = "[冲预交款]" Then
        mshCash.ColData(2) = 0
    Else
        mshCash.ColData(2) = 4
    End If
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtItem(Index))
End Sub

Private Sub txtItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txtItem(Index).Locked Then
        glngTXTProc = GetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txtItem(Index).Locked Then
        Call SetWindowLong(txtItem(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub
