VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBalanceAnalize 
   Caption         =   "医保考核指标统计表"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "frmBalanceAnalize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10215
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd统计 
      Caption         =   "统计"
      Height          =   350
      Left            =   5100
      TabIndex        =   4
      Top             =   90
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp开始日期 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   38785
   End
   Begin VB.CommandButton cmd退出 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   8730
      TabIndex        =   6
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "输出&EXCEL"
      Height          =   350
      Left            =   150
      TabIndex        =   7
      Top             =   6480
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   5865
      Left            =   0
      TabIndex        =   5
      Top             =   510
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   10345
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmBalanceAnalize.frx":0ECA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtp结束日期 
      Height          =   315
      Left            =   3630
      TabIndex        =   3
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   38785
   End
   Begin VB.Label lbl结束日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2760
      TabIndex        =   2
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lbl开始日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   157
      Width           =   720
   End
End
Attribute VB_Name = "frmBalanceAnalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
'统计指定时间范围内，所有病人的费用情况

Public Sub ShowME(ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub cmd统计_Click()
    Dim rsTmp As New ADODB.Recordset
    Call InitTable
    
    '统计病人结算数据
    gstrSQL = "" & _
             " Select B.病人ID,A.结帐ID,B.姓名,B.住院号, " & _
             "        To_char(C.入院日期,'yyyy-MM-dd') As 入院日期, " & _
             "        to_char(C.出院日期,'yyyy-MM-dd') As 出院日期,C.住院天数 As 住院床日, " & _
             "        trim(to_char(A.费用总额,'9000990.00')) AS 费用总额,trim(to_char(A.药品费,'9000990.00')) AS 药品费,trim(to_char(A.非目录内药品费,'9000990.00')) AS 非目录内药品费, " & _
             "        trim(to_char(A.药品费/A.费用总额*100,'9990.00'))||'%' As 药品比例, " & _
             "        trim(to_char(A.非目录内药品费/A.费用总额*100,'9990.00'))||'%' As 非目录内药品比例 " & _
             " From ( " & _
             "      Select A.病人ID,A.主页ID,B.ID As 结帐ID, " & _
             "             sum(Nvl(A.实收金额,0)) 费用总额,"
    gstrSQL = gstrSQL & "Sum(DECODE(A.收费类别,'5',Nvl(A.实收金额,0),'6',Nvl(A.实收金额,0),'7',Nvl(A.实收金额,0),0)) As 药品费, " & _
             "             Sum(DECODE(Nvl(A.统筹金额,0),0,DECODE(A.收费类别,'5',Nvl(A.实收金额,0),'6',Nvl(A.实收金额,0),'7',Nvl(A.实收金额,0),0),0)) As 非目录内药品费 " & _
             "      From 住院费用记录 A,病人结帐记录 B " & _
             "      Where A.结帐ID=B.ID And Nvl(A.实收金额,0)<>0 And Nvl(A.记录状态,0)<>0 And Nvl(A.附加标志,0)<>9 " & _
             "      And B.收费时间 Between to_date('" & Format(dtp开始日期.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') And to_date('" & Format(dtp结束日期.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') " & _
             "      Having sum(Nvl(A.实收金额,0))<>0 " & _
             "      Group By A.病人ID,A.主页ID,B.Id " & _
             " ) A,病人信息 B,病案主页 C,保险帐户 D " & _
             " Where A.病人ID =B.病人ID And B.病人ID=C.病人ID And A.主页ID=C.主页ID And A.病人ID=D.病人ID And D.险类=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "统计病人结算数据", mintInsure)
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    With mshList
        Set .DataSource = rsTmp
        .ColWidth(0) = 1000
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 800
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 7
    End With
End Sub

Private Sub Form_Load()
    Me.dtp结束日期.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    Me.dtp开始日期.Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    Call InitTable
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With cmd退出
        .Left = Me.ScaleWidth - .Width - 150
        .Top = Me.ScaleHeight - .Height - 150
    End With
    cmdExcel.Top = cmd退出.Top
    
    With mshList
        .Height = cmd退出.Top - 600
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub InitTable()
    With mshList
        .Clear
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, 0) = "病人ID"
        .TextMatrix(0, 1) = "结帐ID"
        .TextMatrix(0, 2) = "姓名"
        .TextMatrix(0, 3) = "住院号"
        .TextMatrix(0, 4) = "入院日期"
        .TextMatrix(0, 5) = "出院日期"
        .TextMatrix(0, 6) = "住院床日"
        .TextMatrix(0, 7) = "费用总额"
        .TextMatrix(0, 8) = "药品费"
        .TextMatrix(0, 9) = "非目录内药品费"
        .TextMatrix(0, 10) = "药品比例"
        .TextMatrix(0, 11) = "非目录内药品比例"
    End With
End Sub
