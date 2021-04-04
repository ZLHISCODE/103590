VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm药品批量发药 
   Caption         =   "批量发药"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   Icon            =   "Frm药品批量发药.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10365
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra多处方 
      Height          =   675
      Left            =   120
      TabIndex        =   16
      Top             =   520
      Visible         =   0   'False
      Width           =   9825
      Begin VB.CommandButton cmd过滤 
         Caption         =   "过滤"
         Height          =   375
         Left            =   8520
         TabIndex        =   24
         Top             =   210
         Width           =   1095
      End
      Begin VB.CheckBox chk住院 
         Caption         =   "住院"
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk门诊 
         Caption         =   "门诊"
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chk记账 
         Caption         =   "记账"
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox chk收费 
         Caption         =   "收费"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CommandButton cmd病人科室 
         Height          =   250
         Left            =   8145
         Picture         =   "Frm药品批量发药.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   250
         Width           =   270
      End
      Begin VB.TextBox txt病人科室 
         Height          =   300
         Left            =   6240
         TabIndex        =   35
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl病人科室 
         Caption         =   "病人科室"
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   270
         Width           =   735
      End
      Begin VB.Label lbl门诊标志 
         AutoSize        =   -1  'True
         Caption         =   "门诊标志"
         Height          =   180
         Left            =   2640
         TabIndex        =   20
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl收费类型 
         AutoSize        =   -1  'True
         Caption         =   "单据类型"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.CheckBox chk科室 
      Caption         =   "按开单科室选择"
      Height          =   255
      Index           =   1
      Left            =   8760
      TabIndex        =   36
      Top             =   300
      Width           =   1575
   End
   Begin VB.CheckBox chk科室 
      Caption         =   "按病人科室选择"
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   33
      Top             =   50
      Width           =   1575
   End
   Begin VB.ComboBox Cbo药房 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   180
      Width           =   1560
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      TabIndex        =   6
      Top             =   6030
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   2580
      TabIndex        =   5
      Top             =   6030
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton CmdPrintSet 
      Caption         =   "设置(&S)"
      Height          =   350
      Left            =   1350
      TabIndex        =   4
      Top             =   6030
      Visible         =   0   'False
      Width           =   1100
   End
   Begin TabDlg.SSTab TabShow 
      Height          =   2685
      Left            =   100
      TabIndex        =   1
      Top             =   3120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "单据明细(&D)"
      TabPicture(0)   =   "Frm药品批量发药.frx":0E44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Msf待发明细"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "药品汇总(&T)"
      TabPicture(1)   =   "Frm药品批量发药.frx":0E60
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Msf待发汇总"
      Tab(1).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发明细 
         Height          =   2175
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   9720
         _ExtentX        =   17145
         _ExtentY        =   3836
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发汇总 
         Height          =   2265
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3995
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   -2147483625
         GridColorFixed  =   0
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7920
      TabIndex        =   3
      Top             =   6030
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf待发列表 
      Height          =   1755
      Left            =   100
      TabIndex        =   0
      Top             =   1300
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   3096
      _Version        =   393216
      FixedCols       =   0
      ForeColorSel    =   -2147483640
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6720
      TabIndex        =   2
      Top             =   6030
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker Dtp开始Date 
      Height          =   300
      Left            =   3360
      TabIndex        =   10
      Top             =   180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   123207683
      CurrentDate     =   37007
   End
   Begin MSComCtl2.DTPicker Dtp结束Date 
      Height          =   300
      Left            =   6480
      TabIndex        =   11
      Top             =   180
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      Format          =   123207683
      CurrentDate     =   37007
   End
   Begin VB.CheckBox chk药房 
      Caption         =   "药房"
      Height          =   180
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Value           =   1  'Checked
      Width           =   735
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   6540
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "Frm药品批量发药.frx":0E7C
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13203
            Text            =   "未输入任何处方"
            TextSave        =   "未输入任何处方"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VB.Frame fra单处方 
      Height          =   675
      Left            =   120
      TabIndex        =   26
      Top             =   520
      Width           =   9825
      Begin VB.TextBox TxtNo 
         Height          =   300
         Left            =   840
         TabIndex        =   31
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox Txt姓名 
         Height          =   300
         Left            =   6480
         MaxLength       =   12
         TabIndex        =   28
         Top             =   240
         Width           =   1845
      End
      Begin VB.TextBox txt医保号 
         Height          =   300
         Left            =   3720
         TabIndex        =   27
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label LblNo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "处方号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6000
         TabIndex        =   30
         Top             =   270
         Width           =   360
      End
      Begin VB.Label lbl医保号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3060
         TabIndex        =   29
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Label lbl药房 
      AutoSize        =   -1  'True
      Caption         =   "药房"
      Height          =   180
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Lbl开始Date 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "开始日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2520
      TabIndex        =   13
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl结束Date 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "Frm药品批量发药"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--外部传递参数--
Private mblnModify As Boolean
Private strPrivs As String
Private strUnit As String
Private mint服务对象 As Integer                     '药房的服务对象：1-门诊病人;2-住院病人;3-门诊和住院
Private mlng药房ID As Long                          '药房
Private IntSendAfterDosage As Integer               '允许未配药发药
Private Int允许未审核处方发药 As Integer            '允许未审核处方发药
Private mint允许未收费处方发药 As Integer           '允许未收费处方发药
Private IntCheckStock As Integer                    '库存检测
Private Int校验处方 As Integer                      '校验处方
Private Str窗口 As String                           '发药窗口
Private int金额保留位数 As Integer                  '费用金额保留位数
Private int审核划价单 As Integer                    '执行后自动审核划价单
Private mbln发其他药房处方 As Boolean               '发其他药房处方
Private mbln发药前收费或审核 As Boolean             '发药前必须要收费或者审核
Private mbln发生时间过滤 As Boolean                 '药品医嘱按发生时间(首次时间)过滤：0-按产生单据时间过滤；1-按发生时间过滤
Private mint金额显示 As Integer                     '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
Private mstrOpr As String
Private mblnLoadDrug As Boolean
Private mblnConPacker As Boolean                    '门诊自动发药是否链接
Private mbln启用审方 As Boolean

'--本程序使用变量--
Private RecBill As New ADODB.Recordset              '单据记录
Private RecTotal As New ADODB.Recordset             '汇总数据
Private BlnStartUp As Boolean
Private LngListRow As Long                          '待发列表
Private LngDetailRow As Long                        '待发明细
Private LngTotalRow As Long                         '待发汇总
Private StrBillNo As String                         '汇总单据号
Private strID As String                             '汇总ID

Private mrsApplyforcredit As Recordset                  '用于记录存在销帐申请的单据

Private LngBillCount As Long
Public str配药人 As String
Public str核查人 As String

Private rs序号 As ADODB.Recordset

Private rs处方来源部门 As ADODB.Recordset            '记录所有待发药处方的来源部门

Private rs待发汇总明细 As ADODB.Recordset            '记录待发汇总的记录，实际是按单据号的明细记录
Private mstr汇总单据 As String

Private mstrDeptNode As String          '当前药房的站点
Private mobjDrugMAC As Object
Private mobjPlugIn As Object             '外挂接口对象

'单据操作控制
Private Type Type_BillControl
    bln是否控制 As Boolean
    int时间限制 As Integer
    bln他人单据 As Boolean
    dbl金额上限 As Double
End Type
Private myBillControl As Type_BillControl

Public Property Get In_DrugMAC() As Object
    Set In_DrugMAC = mobjDrugMAC
End Property
Public Property Set In_DrugMAC(ByVal objVal As Object)
    Set mobjDrugMAC = objVal
End Property
Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Public Property Get In_自动发药() As Boolean
    In_自动发药 = mblnConPacker
End Property

Public Property Let In_自动发药(ByVal vNewValue As Boolean)
    mblnConPacker = vNewValue
End Property

Public Property Get In_启用发药() As Boolean
    In_启用发药 = mblnLoadDrug
End Property

Public Property Let In_启用发药(ByVal vNewValue As Boolean)
    mblnLoadDrug = vNewValue
End Property
Public Property Get In_窗口() As String
    In_窗口 = mstrOpr
End Property

Public Property Let In_窗口(ByVal vNewValue As String)
    mstrOpr = vNewValue
End Property

Public Property Get In_发其他药房处方() As Boolean
    In_发其他药房处方 = mbln发其他药房处方
End Property

Public Property Let In_发其他药房处方(ByVal vNewValue As Boolean)
    mbln发其他药房处方 = vNewValue
End Property
Public Property Get In_权限() As String
    In_权限 = strPrivs
End Property

Public Property Let In_权限(ByVal vNewValue As String)
    strPrivs = vNewValue
End Property

Public Property Get In_服务对象() As Integer
    In_服务对象 = mint服务对象
End Property
Public Property Let In_服务对象(ByVal vNewValue As Integer)
    mint服务对象 = vNewValue
End Property

Public Property Get In_校验处方() As Integer
    In_校验处方 = Int校验处方
End Property

Public Property Let In_校验处方(ByVal vNewValue As Integer)
    Int校验处方 = vNewValue
End Property

Public Property Get In_库存检查() As Integer
    In_库存检查 = IntCheckStock
End Property

Public Property Let In_库存检查(ByVal vNewValue As Integer)
    IntCheckStock = vNewValue
End Property

Public Property Get In_药房ID() As Long
    In_药房ID = mlng药房ID
End Property

Public Property Let In_药房ID(ByVal vNewValue As Long)
    mlng药房ID = vNewValue
End Property

Public Property Get In_发药窗口() As String
    In_发药窗口 = Str窗口
End Property

Public Property Let In_发药窗口(ByVal vNewValue As String)
    Str窗口 = vNewValue
End Property

Public Property Get In_允许未配药发药() As Integer
    In_允许未配药发药 = IntSendAfterDosage
End Property

Public Property Let In_允许未配药发药(ByVal vNewValue As Integer)
    IntSendAfterDosage = vNewValue
End Property

Public Property Get IN_允许未审核发药() As Integer
    IN_允许未审核发药 = Int允许未审核处方发药
End Property

Public Property Let IN_允许未审核发药(ByVal vNewValue As Integer)
    Int允许未审核处方发药 = vNewValue
End Property

Public Property Get IN_允许未收费发药() As Integer
    IN_允许未收费发药 = mint允许未收费处方发药
End Property

Public Property Let IN_允许未收费发药(ByVal vNewValue As Integer)
    mint允许未收费处方发药 = vNewValue
End Property

Public Property Get In_金额保留位数() As Integer
    In_金额保留位数 = int金额保留位数
End Property

Public Property Let In_金额保留位数(ByVal vNewValue As Integer)
    int金额保留位数 = vNewValue
End Property

Public Property Get IN_审核划价单() As Integer
    IN_审核划价单 = int审核划价单
End Property

Public Property Let IN_审核划价单(ByVal vNewValue As Integer)
    int审核划价单 = vNewValue
End Property

Private Function CheckBillOperate() As Boolean
    Dim n, i As Integer
    Dim Dbl金额 As Double
    
    For n = 1 To Msf待发列表.rows - 1
        If Msf待发列表.TextMatrix(n, 2) <> "" Then
            Msf待发列表.Row = n
            Call Msf待发列表_EnterCell
            DoEvents
            
            Dbl金额 = 0
            
            For i = 1 To Msf待发明细.rows - 2
                Dbl金额 = Dbl金额 + Val(Msf待发明细.TextMatrix(i, 7))
            Next
            
            If CheckBillControl(3, Val(Msf待发列表.RowData(n)), Msf待发列表.TextMatrix(n, 2), Dbl金额) = False Then
                Exit Function
            End If
        End If
    Next
    
    CheckBillOperate = True
End Function

Private Function CheckDrugStock() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim n As Integer
    Dim lngTemp库房id As Long
    
    If mbln发其他药房处方 = False Then
        lngTemp库房id = Cbo药房.ItemData(Cbo药房.ListIndex)
    Else
        lngTemp库房id = mlng药房ID
    End If
    
    If Msf待发列表.TextMatrix(1, 2) <> "" Then
        For n = 1 To Msf待发汇总.rows - 2
            If MediWork_CheckStorageStock(lngTemp库房id, Val(Msf待发汇总.TextMatrix(n, 9))) = False Then
                MsgBox Msf待发汇总.TextMatrix(n, 1) & "未设置存储库房，不能发药！", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    CheckDrugStock = True
End Function

Private Sub GetRecipe(ByVal intType As Integer, ByVal txtInput As TextBox)
    'intType：1－处方号；2－医保号；3－病人姓名
    Dim RecRecord As New ADODB.Recordset
    Dim intYear As Integer, strYear As String
    Dim intRow As Integer
    Dim strNo As String, IntBill As Integer, ArrTmp, strTmp As String
    Dim strCon As String
    Dim strBeginDate As String
    Dim strEndDate As String
    Dim strInput As String
    Dim strSqlFrom As String
    Dim lng药房ID As Long
    Dim strsql As String
    Dim int记录性质 As Integer
    Dim int门诊标志 As Integer
    
    On Error GoTo errHandle
    
    If Trim(txtInput.Text) = "" Then Exit Sub
    If intType <> 3 Then
        strInput = Trim(UCase(txtInput.Text))
    Else
        strInput = Trim(txtInput.Text)
    End If
    
    If Me.Cbo药房.ListIndex = -1 Then Exit Sub
        
    strBeginDate = Format(Dtp开始Date.Value, "yyyy-MM-dd hh:mm:ss")
    strEndDate = Format(Dtp结束Date.Value, "yyyy-MM-dd hh:mm:ss")
        
    If intType = 1 Then
        strCon = " And C.No=[1] "
    ElseIf intType = 2 Then
        strCon = " And D.医保号=[1] "
    Else
        strCon = " And B.姓名 Like [2] "
    End If
    
    If mbln发其他药房处方 Then
        strCon = strCon & " And C.记录状态=1 And C.填制日期 Between To_Date([4],'yyyy-MM-dd hh24:mi:ss') And To_Date([5] ,'yyyy-MM-dd hh24:mi:ss') "
    Else
        strCon = strCon & " And Mod(C.记录状态,3)=1 And C.填制日期 Between To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') And To_Date([5] ,'yyyy-MM-dd hh24:mi:ss') "
    End If
    
    '批量发药不支持刷卡消费，所以提取的必须是已收费或已审核的处方
    gstrSQL = " Select Distinct S.名称 As 药房, Decode(C.单据,8,'收费',9,'记帐') 类型,C.No,C.单据,A.已收费,Decode(A.配药人,Null,'','部门发药','',A.配药人) 配药人,P.名称 科室,B.姓名,B.标识号 住院号,'' 床号," & _
             " B.开单人 开单医生,B.操作员姓名 填制人,To_Char(C.填制日期,'yyyy-MM-dd') 填制日期, S.ID As 药房id,B.记录性质,B.门诊标志,A.病人ID, d.病人类型 " & _
             " From 未发药品记录 A,门诊费用记录 B,药品收发记录 C,部门表 P,部门表 S, 病人信息 D " & IIf(mbln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & _
             " Where C.费用ID=B.ID And B.开单部门ID+0=P.ID And Nvl(C.库房ID,0)+0=S.ID and Nvl(A.库房ID,0)=Nvl(C.库房ID,0)  And A.No=C.No " & IIf(mbln启用审方, " and b.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 And ((b.医嘱序号 is null or nvl(q.审查结果,0) = 1) or not Exists(select 1 from 处方审查记录 Q where q.病人id = d.病人id and q.提交科室id = p.id and q.发药药房id = b.执行部门id And q.id = k.审方id))", "") & _
             IIf(mbln发其他药房处方, "", IIf(Str窗口 = "", "", " And (C.发药窗口 In (Select * From Table(Cast(f_Str2list([7]) As Zltools.t_Strlist))) Or C.发药窗口 Is NULL)")) & _
             " And C.审核人 Is Null And Nvl(B.费用状态,0)<>1 " & _
             " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) " & _
             " And C.单据=A.单据 and nvl(C.发药方式,-999)<>-1 And A.病人id=D.病人id(+) " & strCon
    
    If Me.chk收费.Value = 1 And Me.chk记账.Value = 1 Then
        If mbln发药前收费或审核 = True Then
            gstrSQL = gstrSQL & " And A.单据 In(8,9) And A.已收费=1 "
        ElseIf mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And (C.单据=8 And A.已收费=1 Or C.单据=9) "
        ElseIf Int允许未审核处方发药 = False Then
            gstrSQL = gstrSQL & " And (C.单据=9 And A.已收费=1 Or C.单据=8) "
        Else
            gstrSQL = gstrSQL & " And A.单据 In(8,9) "
        End If
    ElseIf Me.chk收费.Value = 1 Then
        If mbln发药前收费或审核 = True Or mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And A.单据=8 And A.已收费=1 "
        Else
            gstrSQL = gstrSQL & " And A.单据=8 "
        End If
    ElseIf Me.chk记账.Value = 1 Then
        If mbln发药前收费或审核 = True Or mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And A.单据=9 And A.已收费=1 "
        Else
            gstrSQL = gstrSQL & " And A.单据=9 "
        End If
    End If
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (P.站点 = [8] Or P.站点 Is Null) "
    End If
    
    If mbln发其他药房处方 = True Then
        If chk药房.Value = 1 Then
            gstrSQL = gstrSQL & " And C.库房ID+0=[3] "
        Else
            gstrSQL = gstrSQL & " And C.库房ID+0<>[6] "
        End If
    Else
        gstrSQL = gstrSQL & " And (C.库房ID+0=[3] OR C.库房ID IS NULL)"
    End If
    
    If mbln发生时间过滤 = True Then
        strsql = gstrSQL & " And B.医嘱序号 Is Null"
        gstrSQL = Replace(gstrSQL, "C.填制日期", "B.发生时间") & " And B.医嘱序号 Is Not Null"
        gstrSQL = strsql & " Union All " & gstrSQL
    End If
    
    If mint服务对象 = 3 Then
        strsql = Replace(gstrSQL, "'' 床号", "B.床号")
        strsql = Replace(strsql, "门诊费用记录", "住院费用记录")
        strsql = Replace(strsql, "And Nvl(B.费用状态,0)<>1", "")
        gstrSQL = gstrSQL & " Union All " & strsql
    ElseIf mint服务对象 = 2 Then
        gstrSQL = Replace(gstrSQL, "'' 床号", "B.床号")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.费用状态,0)<>1", "")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput, strInput & "%", Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex), strBeginDate, strEndDate, mlng药房ID, Str窗口, mstrDeptNode)
    
    If RecBill.EOF Then
        
        If intType = 1 Then
            If CheckBillExist(strInput, mlng药房ID) = False Then GoTo ExitSub
        End If
        
        MsgBox "未找到指定处方，或指定的处方未收费或未审核，请重新输入！", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    If RecBill.RecordCount > 1 Then
        strTmp = Frm单据选择.ShowMe(Me, RecBill)
        If strTmp = "" Then GoTo ExitSub
        
        ArrTmp = Split(strTmp, ";")
        strNo = ArrTmp(0)
        IntBill = ArrTmp(1)
        lng药房ID = ArrTmp(2)
                
        RecBill.MoveFirst
'        RecBill.Find "单据=" & IntBill
        RecBill.Filter = "单据=" & IntBill & " And No='" & strNo & "'"
        
        int记录性质 = RecBill!记录性质
        int门诊标志 = RecBill!门诊标志
    Else
        strNo = RecBill!NO
        IntBill = RecBill!单据
        lng药房ID = RecBill!药房ID
        int记录性质 = RecBill!记录性质
        int门诊标志 = RecBill!门诊标志
    End If
    
    Me.TxtNo = strNo
    Me.TxtNo.Tag = IntBill
    Me.Txt姓名.Tag = IntBill
    
    '如果已存在该单据，则退出
    If SetLocateBill(False) Then
        MsgBox "该处方已经输入，请重输！", vbInformation, gstrSysName
        GoTo ExitSub
    End If
    
    '如果当前输入处方的科室与已录入的处方的科室不同，则给予提示
    If CheckSource(IntBill, strNo, lng药房ID) = False Then Exit Sub
    If WriteSendListData(0) = False Then GoTo ExitSub
    
    LngBillCount = LngBillCount + 1
    Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
    '定位到刚才输入的处方单
    Call SetLocateBill
    
    With Msf待发列表
        CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
    End With
    
    mblnModify = True
'    If tabShow.Tab = 1 Then Call RefreshData
    Call RefreshData(lng药房ID)
    With TxtNo
        .SelStart = 0
        .SelLength = Len(txtInput)
    End With
    Exit Sub
    
ExitSub:
    With txtInput
        .SelStart = 0
        .SelLength = Len(txtInput)
        .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function CheckBillExist(ByVal strNo As String, ByVal lng药房ID As Long) As Boolean
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select A.配药人,A.审核人,C.操作员姓名 填制人 " & _
            " From 药品收发记录 A, 门诊费用记录 C " & _
            " Where A.费用id = C.ID And mod(A.记录状态,3)=1 And Rownum=1 " & _
            " And A.No=[1] And A.单据 in (8,9,10)"
    
    If mbln发其他药房处方 = True Then
        If chk药房.Value = 1 Then
            gstrSQL = gstrSQL & " And A.库房ID+0=[2] "
        Else
            gstrSQL = gstrSQL & " And A.库房ID+0<>[2] "
        End If
    Else
        gstrSQL = gstrSQL & " And (A.库房ID+0=[2] OR A.库房ID IS NULL)"
    End If
    
    gstrSQL = gstrSQL & "Union All" & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, lng药房ID)
        
    With rstemp
        If .EOF Then CheckBillExist = False: MsgBox "该处方[" & strNo & "]不存在！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            CheckBillExist = False
            MsgBox "该处方[" & strNo & "]已被其它操作员发药！", vbInformation, gstrSysName: Exit Function
        End If
    End With

    CheckBillExist = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub IniControl()
    LngBillCount = 0
    Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
    
    '初始化
    strID = ""
    StrBillNo = ""
    TxtNo = ""
    Txt姓名 = ""
    
    With Msf待发汇总
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf待发列表
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    With Msf待发明细
        .Clear
        .rows = 2
        .RowData(1) = 0
    End With
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    CmdOK.Enabled = False
    If Me.chk科室(0).Value = 0 And Me.chk科室(1).Value = 0 Then TxtNo.SetFocus
End Sub

Private Sub InitRecSum()
    '初始化汇总数据集
    Set rs待发汇总明细 = New ADODB.Recordset
    With rs待发汇总明细
        If .State = 1 Then .Close
        .Fields.Append "单据号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "单价", adDouble, 18, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .Fields.Append "实收金额", adDouble, 18, adFldIsNullable
        .Fields.Append "药品id", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Ini药房()
    Dim rstemp As ADODB.Recordset
    Dim n As Integer
    Dim lngTemp As Long
    
    On Error GoTo errHandle

    Me.Cbo药房.Enabled = mbln发其他药房处方
    
    If zlStr.IsHavePrivs(strPrivs, "所有药房") Or mbln发其他药房处方 Then
        gstrSQL = "(Select Distinct 部门ID From 部门性质说明 Where 工作性质 Like '%药房')"
    Else
        gstrSQL = "(Select distinct A.部门ID From 部门人员 A,部门性质说明 B " & _
                 " Where A.人员ID=[1] And A.部门ID=B.部门ID And B.工作性质 Like '%药房')"
    End If
    gstrSQL = " Select Distinct P.ID,P.名称 From 部门表 P " & _
             " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In " & gstrSQL & _
             " And (P.撤档时间 Is Null Or P.撤档时间=To_Date('3000-01-01','yyyy-MM-dd'))"
    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, glngUserId)
    
    With Me.Cbo药房
        Do While Not rstemp.EOF
            If Not mbln发其他药房处方 Then
                .AddItem rstemp!名称
                n = .NewIndex
                .ItemData(n) = rstemp!Id
                
                If lngTemp = 0 Then
                    If rstemp!Id = mlng药房ID Then
                        lngTemp = n
                    End If
                End If
            Else
                If rstemp!Id <> mlng药房ID Then
                    .AddItem rstemp!名称
                    n = .NewIndex
                    .ItemData(n) = rstemp!Id
                Else
                    Me.Caption = "发其他药房处方(当前药房：" & rstemp!名称 & ")"
                End If
            End If
            
            rstemp.MoveNext
        Loop
        
        .ListIndex = lngTemp
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub SetFormat(Optional ByVal IntStyle As Integer = 1)
    Dim intCol As Integer
    '设置各列表控件的格式

    Select Case IntStyle
    Case 1
        With Msf待发列表
            .rows = 2
            .Cols = 15
            
            .TextMatrix(0, 0) = "药房"
            .TextMatrix(0, 1) = "类型"
            .TextMatrix(0, 2) = "NO"
            .TextMatrix(0, 3) = "科室"
            .TextMatrix(0, 4) = "姓名"
            .TextMatrix(0, 5) = "住院号"
            .TextMatrix(0, 6) = "床号"
            .TextMatrix(0, 7) = "收费员"
            .TextMatrix(0, 8) = "开单医生"
            .TextMatrix(0, 9) = "开单日期"
            .TextMatrix(0, 10) = "药房ID"
            .TextMatrix(0, 11) = "记录性质"
            .TextMatrix(0, 12) = "门诊标志"
            .TextMatrix(0, 13) = "已收费"
            .TextMatrix(0, 14) = "病人ID"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = IIf(mbln发其他药房处方 = True And chk药房.Value = 0, 1500, 0)
                .ColWidth(1) = 500
                .ColWidth(2) = 1000
                .ColWidth(3) = 1200
                .ColWidth(4) = 1000
                .ColWidth(5) = 1000
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                .ColWidth(9) = 1200
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
                .ColWidth(13) = 0
                .ColWidth(14) = 0
                
                .Row = 1
                Call RestoreFlexState(Msf待发列表, Me.Name)
                If glngSys \ 100 <> 1 Then
                    .ColWidth(3) = 0
                    .ColWidth(5) = 0
                    .ColWidth(6) = 0
                End If
                .ColWidth(8) = IIf(Int校验处方 = 1, 0, 1000)
            End If
        End With
    Case 2
        With Msf待发明细
            .rows = 2
            .Cols = 8
    
            .TextMatrix(0, 0) = "药品名称"
            .TextMatrix(0, 1) = "商品名"
            .TextMatrix(0, 2) = "规格"
            .TextMatrix(0, 3) = "单位"
            .TextMatrix(0, 4) = "单价"
            .TextMatrix(0, 5) = "数量"
            .TextMatrix(0, 6) = "应收金额"
            .TextMatrix(0, 7) = "实收金额"
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
    
            If BlnStartUp = False Then
                .ColWidth(0) = 2000
                .ColWidth(2) = 1500
                .ColWidth(3) = 500
                .ColWidth(4) = 800
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 1000
                
                .Row = 1
                Call RestoreFlexState(Msf待发明细, Me.Name)
                If gint药品名称显示 = 2 Then
                    If .ColWidth(1) = 0 Then .ColWidth(1) = 2000
                Else
                    .ColWidth(1) = 0
                End If
                
                If mint金额显示 = 0 Then
                    .ColWidth(7) = 0
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                ElseIf mint金额显示 = 1 Then
                    .ColWidth(6) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                Else
                    If .ColWidth(6) <= 0 Then .ColWidth(6) = 1000
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                End If
                
            End If
        End With
    Case 3
        With Msf待发汇总
            .rows = 2
            .Cols = 11
    
            .TextMatrix(0, 0) = "序号"
            .TextMatrix(0, 1) = "药品名称"
            .TextMatrix(0, 2) = "商品名"
            .TextMatrix(0, 3) = "规格"
            .TextMatrix(0, 4) = "单位"
            .TextMatrix(0, 5) = "单价"
            .TextMatrix(0, 6) = "数量"
            .TextMatrix(0, 7) = "应收金额"
'            .TextMatrix(0, 8) = "药品id"
'            .TextMatrix(0, 9) = "批次"
            .TextMatrix(0, 8) = "实收金额"
            .TextMatrix(0, 9) = "药品id"
            .TextMatrix(0, 10) = "批次"

            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            If BlnStartUp = False Then
                .ColWidth(0) = 500
                .ColWidth(1) = 2000
                .ColWidth(3) = 1500
                .ColWidth(4) = 500
                .ColWidth(5) = 800
                .ColWidth(6) = 800
                .ColWidth(7) = 1000
                .ColWidth(8) = 1000
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                .Row = 1
                Call RestoreFlexState(Msf待发汇总, Me.Name)
                
                If gint药品名称显示 = 2 Then
                    If .ColWidth(2) = 0 Then .ColWidth(2) = 2000
                Else
                    .ColWidth(2) = 0
                End If
                
                If mint金额显示 = 0 Then
                    .ColWidth(8) = 0
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                ElseIf mint金额显示 = 1 Then
                    .ColWidth(7) = 0
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                Else
                    If .ColWidth(7) <= 0 Then .ColWidth(7) = 1000
                    If .ColWidth(8) <= 0 Then .ColWidth(8) = 1000
                End If
            End If
        End With
    End Select
End Sub



Private Sub chk记账_Click()
    If Me.chk记账.Value = 0 Then
        Me.chk收费.Value = 1
    End If
End Sub

Private Sub chk科室_Click(index As Integer)
    Dim i As Integer
    
    
    If index = 0 Then
        If Me.chk科室(0).Value = 1 Then
            If Me.chk科室(1).Value = 1 Then
                Me.chk科室(1).Value = 0
            End If
            Me.lbl病人科室.Caption = "病人科室"
        End If
        
    Else
        If Me.chk科室(1).Value = 1 Then
            If Me.chk科室(0).Value = 1 Then
                Me.chk科室(0).Value = 0
            End If
            Me.lbl病人科室.Caption = "开单科室"
        End If
    End If
    
    If Me.chk科室(0).Value = 0 And Me.chk科室(1).Value = 0 Then
        Me.fra多处方.Visible = False
        Me.fra单处方.Visible = True

    Else
        Me.fra多处方.Visible = True
        Me.fra单处方.Visible = False
    End If



    If rs待发汇总明细.RecordCount > 0 Then
        Set RecTotal = Nothing
        With Msf待发列表
            .Clear
            .rows = 2
            Call SetFormat(1)
        End With

        With Msf待发明细
            .Clear
            .rows = 2
            Call SetFormat(2)
        End With

        With Msf待发汇总
            .Clear
            .rows = 2
            Call SetFormat(3)
        End With

        rs待发汇总明细.MoveLast

        For i = 0 To rs待发汇总明细.RecordCount - 1
            ''''删除当前行
            rs待发汇总明细.Delete adAffectCurrent
            ''''向前移动指针
            rs待发汇总明细.MovePrevious
        Next

        Me.stbThis.Panels(2).Text = "未输入任何处方"
    End If
End Sub

Private Sub chk门诊_Click()
    If Me.chk住院.Enabled = False Then
        Me.chk门诊.Value = 1
    Else
        If Me.chk门诊.Value = 0 Then
            Me.chk住院.Value = 1
        End If
    End If
End Sub

Private Sub chk收费_Click()
    If Me.chk收费.Value = 0 Then
        Me.chk记账.Value = 1
    End If
End Sub

Private Sub chk药房_Click()
    IniControl
    
    If chk药房.Value = 1 Then
        Cbo药房.Enabled = True
        Msf待发列表.ColWidth(0) = 0
    Else
        Cbo药房.Enabled = False
        If Msf待发列表.ColWidth(0) = 0 Then
            Msf待发列表.ColWidth(0) = 1500
        End If
    End If
End Sub

Private Sub chk住院_Click()
    If Me.chk门诊.Enabled = False Then
        Me.chk住院.Value = 1
     Else
        If Me.chk住院.Value = 0 Then
            Me.chk门诊.Value = 1
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    '启用电子签名时检查用户是否注册
    If gblnESign处方发药 = True Then
        If Not gobjESign.CheckCertificate(gstrDbUser) Then
            Exit Sub
        End If
    End If
    
'    Call RefreshData
     '检查单据是否是当天的单据
    If mbln发其他药房处方 Then
        If CheckDate = False Then Exit Sub
    End If
    
    If CheckDrugStock = False Then Exit Sub
    If CheckStock = False Then Exit Sub
    If Not CheckCorrelation Then Exit Sub
    If Not CheckBillOperate Then Exit Sub
    If SendBill = False Then Exit Sub
    
    IniControl
End Sub

Private Function CheckDate() As Boolean
'用于发其他药房处方时，检查是否是当天的单据
    Dim i As Integer
    Dim dateCur As Date
    
    dateCur = Sys.Currentdate
    With Msf待发列表
        For i = 1 To .rows - 1
            If .TextMatrix(i, 2) <> "" Then
                If Format(.TextMatrix(i, 9), "YYYY-MM-DD") < Format(dateCur, "YYYY-MM-DD") Then
                    If MsgBox("        代发非当天单据，会删除汇总数据重新汇总，" & vbCrLf & "如果已经出了报表的可能需要重新出报表，是否继续操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        CheckDate = False
                    Else
                        CheckDate = True
                    End If
                    Exit Function
                End If
            End If
        Next
    End With
    
    CheckDate = True
End Function

Private Sub cmdPrint_Click()
    Dim HisPrint As New zlPrint1Grd
    Dim HisRow As New zlTabAppRow
    Dim ArrayNo, IntArray As Integer
    Dim LngSelectRow As Long, intCol As Integer
    
    On Error Resume Next
    '取消表格的选择状态
    With Msf待发汇总
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
    End With
    
    HisPrint.Title = "药品汇总"
    Set HisRow = New zlTabAppRow
    HisRow.Add "日期:" & Format(Sys.Currentdate, "yyyy年MM月dd日")
    HisPrint.UnderAppRows.Add HisRow
    
    ArrayNo = Split(StrBillNo, ";")
    
    Set HisRow = New zlTabAppRow
    HisRow.Add "单据号:"
    HisPrint.BelowAppRows.Add HisRow
    For IntArray = 0 To UBound(ArrayNo)
        Set HisRow = New zlTabAppRow
        HisRow.Add Space(10) & ArrayNo(IntArray)
        HisPrint.BelowAppRows.Add HisRow
    Next
    
    Set HisPrint.Body = Msf待发汇总
    Select Case zlPrintAsk(HisPrint)
    Case 1
        zlPrintOrView1Grd HisPrint, 1
    Case 2
        zlPrintOrView1Grd HisPrint, 2
    Case 3
        zlPrintOrView1Grd HisPrint, 3
    End Select
    
    '恢复表格的选择状态
    With Msf待发汇总
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub CmdPrintSet_Click()
    zlPrintSet
End Sub

Private Sub cmd病人科室_Click()
'检查','检验','治疗','手术','营养'
    If Me.lbl病人科室.Caption = "病人科室" Then
        If Select部门(Me, Me.txt病人科室, "", "临床", False, mint服务对象) = False Then
            Exit Sub
        End If
    Else
        If Select部门(Me, Me.txt病人科室, "", "检查,检验,治疗,手术,营养", False, mint服务对象) = False Then
            Exit Sub
        End If
    End If
    
End Sub

Private Sub cmd过滤_Click()
    Dim strBeginDate As String
    Dim strEndDate As String
    Dim strCon As String
    Dim strsql As String
    Dim i As Integer
    
    On Error GoTo errHandle
    If Me.chk记账.Value = 0 And Me.chk收费.Value = 0 Then
        MsgBox "请选择收费类型！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If Me.chk门诊.Value = 0 And Me.chk住院.Value = 0 Then
        MsgBox "请选择门诊标志！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If Trim(Me.txt病人科室.Tag) = "" Then
        MsgBox "请选择病人科室！", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
        
    strBeginDate = Format(Dtp开始Date.Value, "yyyy-MM-dd hh:mm:ss")
    strEndDate = Format(Dtp结束Date.Value, "yyyy-MM-dd hh:mm:ss")
        
    If mbln发其他药房处方 Then
        strCon = strCon & " And C.记录状态=1 And C.填制日期 Between To_Date([3],'yyyy-MM-dd hh24:mi:ss') And To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') "
    Else
        strCon = strCon & " And Mod(C.记录状态,3)=1 And C.填制日期 Between To_Date([3] ,'yyyy-MM-dd hh24:mi:ss') And To_Date([4] ,'yyyy-MM-dd hh24:mi:ss') "
    End If
    
    '批量发药不支持刷卡消费，所以提取的必须是已收费或已审核的处方
    gstrSQL = " Select /*+ Rule*/ Distinct S.名称 As 药房, Decode(C.单据,8,'收费',9,'记帐') 类型,C.No,C.单据,A.已收费,Decode(A.配药人,Null,'','部门发药','',A.配药人) 配药人,P.名称 科室,B.姓名,B.标识号 住院号,'' 床号," & _
             " B.开单人 开单医生,B.操作员姓名 填制人,To_Char(C.填制日期,'yyyy-MM-dd') 填制日期, S.ID As 药房id,B.记录性质,B.门诊标志,A.病人ID, d.病人类型 " & _
             " From 未发药品记录 A,门诊费用记录 B,药品收发记录 C,部门表 P,部门表 S, 病人信息 D " & IIf(Str窗口 = "", "", ",Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)) E ") & IIf(mbln启用审方, ",处方审查记录 Q,处方审查明细 K ", "") & _
             " Where C.费用ID=B.ID And B.开单部门ID+0=P.ID And Nvl(C.库房ID,0)+0=S.ID and Nvl(A.库房ID,0)=Nvl(C.库房ID,0)  And A.No=C.No " & IIf(mbln启用审方, " and b.医嘱序号=k.医嘱id(+) and Q.id(+)=K.审方id and K.最后提交(+)=1 And (b.医嘱序号 is null or nvl(q.审查结果,0) = 1)", "") & _
             IIf(Str窗口 = "", "", " And (C.发药窗口=E.Column_Value Or C.发药窗口 Is NULL)") & _
            IIf(IntSendAfterDosage = 0, " And C.配药人 is not null And C.配药日期 is not null", "") & _
             " And C.审核人 Is Null  And Nvl(B.费用状态,0)<>1 " & _
             " and Not Exists(select 1 from 药品收发记录 F where F.单据=C.单据 and F.库房id=C.库房id and F.no=C.no and 发药方式=-1) " & _
             " And C.单据=A.单据 and nvl(C.发药方式,-999)<>-1 And A.病人id=D.病人id(+) " & strCon
    
    If Me.chk收费.Value = 1 And Me.chk记账.Value = 1 Then
        If mbln发药前收费或审核 = True Then
            gstrSQL = gstrSQL & " And A.单据 In(8,9) And A.已收费=1 "
        ElseIf mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And (C.单据=8 And A.已收费=1 Or C.单据=9) "
        ElseIf Int允许未审核处方发药 = False Then
            gstrSQL = gstrSQL & " And (C.单据=9 And A.已收费=1 Or C.单据=8) "
        Else
            gstrSQL = gstrSQL & " And A.单据 In(8,9) "
        End If
    ElseIf Me.chk收费.Value = 1 Then
        If mbln发药前收费或审核 = True Or mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And A.单据=8 And A.已收费=1 "
        Else
            gstrSQL = gstrSQL & " And A.单据=8 "
        End If
    ElseIf Me.chk记账.Value = 1 Then
        If mbln发药前收费或审核 = True Or mint允许未收费处方发药 = False Then
            gstrSQL = gstrSQL & " And A.单据=9 And A.已收费=1 "
        Else
            gstrSQL = gstrSQL & " And A.单据=9 "
        End If
    End If
    
    If Me.chk门诊.Value <> 1 Or Me.chk住院.Value <> 1 Then
        If Me.chk门诊.Value = 1 Then
            gstrSQL = gstrSQL & " And (B.记录性质=1 or (b.记录性质=2 and (B.门诊标志=1 or B.门诊标志=4)))"
        Else
            gstrSQL = gstrSQL & " And (b.记录性质=2 and (B.门诊标志<>1 and B.门诊标志<>4))"
        End If
    End If
    
    If mstrDeptNode <> "" Then
        gstrSQL = gstrSQL & " And (P.站点 = [7] Or P.站点 Is Null) "
    End If
    
    If mbln发其他药房处方 = True Then
        If chk药房.Value = 1 Then
            gstrSQL = gstrSQL & " And C.库房ID+0=[2] "
        Else
            gstrSQL = gstrSQL & " And C.库房ID+0<>[5] "
        End If
    Else
        gstrSQL = gstrSQL & " And (C.库房ID+0=[2] OR C.库房ID IS NULL)"
    End If
    
    If Me.lbl病人科室.Caption = "病人科室" Then
        gstrSQL = gstrSQL & " And B.病人科室id in (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) "
    Else
        gstrSQL = gstrSQL & " And B.开单部门ID in (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist))) "
    End If
    
    If Me.chk门诊.Value = 1 And Me.chk住院.Value = 1 Then
        strsql = Replace(gstrSQL, "'' 床号", "B.床号")
        strsql = Replace(strsql, "门诊费用记录", "住院费用记录")
        strsql = Replace(strsql, "And Nvl(B.费用状态,0)<>1", "")
        gstrSQL = gstrSQL & " Union All " & strsql
    ElseIf Me.chk住院.Value = 1 Then
        gstrSQL = Replace(gstrSQL, "'' 床号", "B.床号")
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
        gstrSQL = Replace(gstrSQL, "And Nvl(B.费用状态,0)<>1", "")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Me.txt病人科室.Tag, Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex), strBeginDate, strEndDate, mlng药房ID, Str窗口, mstrDeptNode)
    
    '清空当前列表
    Set RecTotal = Nothing
    With Msf待发列表
        .Clear
        .rows = 2
        Call SetFormat(1)
    End With
    
    With Msf待发明细
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    With Msf待发汇总
        .Clear
        .rows = 2
        Call SetFormat(3)
    End With
    
    If rs待发汇总明细.RecordCount > 0 Then
        rs待发汇总明细.MoveLast
        
        For i = 0 To rs待发汇总明细.RecordCount - 1
            ''''删除当前行
            rs待发汇总明细.Delete adAffectCurrent
            ''''向前移动指针
            rs待发汇总明细.MovePrevious
        Next
        
        Me.stbThis.Panels(2).Text = "未输入任何处方"
    End If
    
    If Not RecBill.EOF Then
        Call InitRec
        '将单据信息与明细序号写入内部映射记录集中
'        With rs序号
'            If RecBill.RecordCount <> 0 Then
'                Do While Not RecBill.EOF
'                    .AddNew
'                    !单据标识 = RecBill!NO & "|" & RecBill!单据
'                    !序号 = RecBill!序号
'                    !记录性质 = RecBill!记录性质
'                    !门诊标志 = RecBill!门诊标志
'                    .Update
'                    RecBill.MoveNext
'                Loop
'            End If
'            RecBill.MoveFirst
'        End With
        
        If WriteSendListData(1) = True Then
            
            Me.stbThis.Panels(2).Text = "已输入" & RecBill.RecordCount & "张处方"
            
            Call Msf待发列表_EnterCell
            
            
            
'            '定位到刚才输入的处方单
'            Call SetLocateBill
            
            With Msf待发列表
                CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            End With
            
            mblnModify = True
            Call RefreshData(Me.Cbo药房.ItemData(Me.Cbo药房.ListIndex))
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
    
    If mbln发其他药房处方 = True Then
        chk药房.Visible = True
        lbl药房.Visible = False
       
        chk药房.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "代发药药房选择", "1"))
    
    Else
        chk药房.Visible = False
        lbl药房.Visible = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim dateCurDate As Date
    
    BlnStartUp = False
    LngBillCount = 0
    
    strID = ""
    StrBillNo = ""
    
    dateCurDate = Sys.Currentdate()
    Me.Dtp开始Date.Value = Format(dateCurDate, "yyyy-MM-dd 00:00:00")
    Me.Dtp结束Date.Value = Format(dateCurDate, "yyyy-MM-dd 23:59:59")
    
    If mbln发其他药房处方 Then
        Dtp开始Date.MinDate = dateCurDate - 30
        Dtp结束Date.MinDate = dateCurDate - 30
    End If
    
    mbln发药前收费或审核 = (gtype_UserSysParms.P163_项目执行前必须先收费或先记帐审核 = 1)
    mbln发生时间过滤 = (Val(zldatabase.GetPara("药品医嘱按发生时间过滤", glngSys, 1341, 0)) = 1)
    mint金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
    mbln启用审方 = ((gtype_UserSysParms.P240_药房处方审查 = 1 Or gtype_UserSysParms.P240_药房处方审查 = 3) And gtype_UserSysParms.P241_处方审查时机 = 2)
    
    
    Call SetFormat(1)
    Call SetFormat(2)
    Call SetFormat(3)
    
    Call InitRec
    
    Call Ini药房
   
    If Me.Cbo药房.ListCount = 0 Then
        MsgBox "没有设置其他药房！", vbInformation, gstrSysName
        Unload Me
    End If
     
    If mint服务对象 = 3 Then
        Me.chk门诊.Value = 1
        Me.chk住院.Value = 1
    End If
    
    If mint服务对象 <> 1 And mint服务对象 <> 3 Then
        Me.chk门诊.Enabled = False
        Me.chk住院.Value = 1
    End If
    
    If mint服务对象 <> 2 And mint服务对象 <> 3 Then
        Me.chk住院.Enabled = False
        Me.chk门诊.Value = 1
    End If
    
    BlnStartUp = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Width < 9495 Then Me.Width = 9495
    If Me.Height < 6705 Then Me.Height = 6705
    
    With CmdHelp
        .Top = Me.ScaleHeight - .Height - 100 - Me.stbThis.Height
    End With
    
    With CmdPrintSet
        .Top = CmdHelp.Top
        .Left = CmdHelp.Left + CmdHelp.Width + 100
    End With
    
    With CmdPrint
        .Top = CmdHelp.Top
        .Left = CmdPrintSet.Left + CmdPrintSet.Width + 100
    End With
    
    With CmdCancel
        .Top = CmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 100
    End With
    
    With CmdOK
        .Top = CmdHelp.Top
        .Left = CmdCancel.Left - .Width - 100
    End With
    
    With Msf待发列表
        .Height = (CmdOK.Top - 200 - .Top) / 2
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With TabShow
        .Top = Msf待发列表.Top + Msf待发列表.Height + 100
        .Height = CmdOK.Top - 100 - .Top
        .Width = Msf待发列表.Width
    End With
    
    With Msf待发汇总
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    
    With Msf待发明细
        .Left = 50
        .Height = TabShow.Height - .Top - 80
        .Width = TabShow.Width - .Left - 50
    End With
    
    With fra单处方
        .Left = 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
    
    With fra多处方
        .Left = 100
        .Width = Me.ScaleWidth - .Left - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveFlexState(Msf待发汇总, Me.Name)
    Call SaveFlexState(Msf待发列表, Me.Name)
    Call SaveFlexState(Msf待发明细, Me.Name)
    
    If mbln发其他药房处方 = True Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "代发药药房选择", chk药房.Value)
    End If
End Sub

Private Sub Msf待发汇总_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发汇总
        .Redraw = False
        
        
        LngSelectRow = .Row     '保存当前选中行
        If LngTotalRow > 0 And LngTotalRow < .rows Then
            .Row = LngTotalRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngTotalRow = LngSelectRow
        .Row = LngTotalRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发汇总_GotFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发汇总_LostFocus()
    With Msf待发汇总
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发列表_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发列表
        .Redraw = False


        LngSelectRow = .Row     '保存当前选中行
        If LngListRow > 0 And LngListRow < .rows Then
            .Row = LngListRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H80000005
                If intCol <> 4 Then
                    .CellForeColor = &H80000008
                End If
            Next
            .Col = 0
        End If

        LngListRow = LngSelectRow
'        If LngSelectRow = 0 Then
'            LngListRow = 1
'        End If
        .Row = LngListRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
            .Col = intCol
            .CellBackColor = &H8000000D
            If intCol <> 4 Then
                .CellForeColor = &H80000005
            End If
        Next
        .Col = 0
        .Redraw = True
        
        If Trim(.TextMatrix(.Row, 2)) = "" Then
            With Msf待发明细
                .Clear
                .rows = 2
                Call SetFormat(2)
            End With
            Exit Sub
        End If
        
        '显示单据明细
        Call ReadBillData(.RowData(.Row), .TextMatrix(.Row, 2), Val(.TextMatrix(.Row, 10)), Val(.TextMatrix(.Row, 11)), Val(.TextMatrix(.Row, 12)))
    End With
End Sub

Private Sub Msf待发列表_GotFocus()
    With Msf待发列表
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发列表_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng单据 As Long, strNo As String
    Dim int记录性质 As Integer
    Dim int门诊标志 As Integer
    If KeyCode = vbKeyDelete Then
        If Msf待发列表.TextMatrix(Msf待发列表.Row, 2) = "" Then Exit Sub
        With rs待发汇总明细
            If .RecordCount > 0 Then
                .MoveFirst
                Do While Not .EOF
                    .Find "单据号='" & Msf待发列表.TextMatrix(Msf待发列表.Row, 2) & "'"
                    If Not .EOF Then .Delete
                    If Not .EOF Then .MoveNext
                Loop
            End If
        End With
        With rs处方来源部门
            If .RecordCount > 0 Then
                .MoveFirst
                .Find "来源部门='" & Msf待发列表.TextMatrix(Msf待发列表.Row, 3) & "'"
                If Not .EOF Then .Delete
            End If
        End With
        With Msf待发列表
            lng单据 = Val(.RowData(.Row))
            strNo = .TextMatrix(.Row, 2)
            If .rows - 1 = 1 Then
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
                .TextMatrix(1, 7) = ""
                .TextMatrix(1, 8) = ""
                .TextMatrix(1, 9) = ""
                .TextMatrix(1, 10) = ""
                .TextMatrix(1, 11) = ""
                .TextMatrix(1, 12) = ""
                .RowData(1) = 0
            Else
                If Trim(.TextMatrix(.Row, 2)) <> "" Then .RemoveItem .Row: LngBillCount = LngBillCount - 1
            End If
            
            CmdOK.Enabled = (.RowData(IIf(.rows - 1 = 1, 1, .rows - 2)) <> 0)
            Me.stbThis.Panels(2).Text = IIf(LngBillCount = 0, "未输入任何处方", "已输入" & LngBillCount & "张处方")
            'Call RefreshData
        
            '删除该单据
            With rs序号
                If .RecordCount <> 0 Then .MoveFirst
                .Find "单据标识='" & strNo & "|" & lng单据 & "'"
                If Not .EOF Then .Delete
            End With
            
            If rs序号.RecordCount = 0 Then InitRec
        End With
        
        Msf待发列表_EnterCell
        mblnModify = True
'        If tabShow.Tab = 1 Then Call RefreshData
        Call WriteTotalDataToBill
    End If
End Sub

Private Sub Msf待发列表_LostFocus()
    With Msf待发列表
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf待发明细_EnterCell()
    Dim LngSelectRow As Long
    Dim intCol As Integer
    With Msf待发明细
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngDetailRow > 0 And LngDetailRow < .rows Then
            .Row = LngDetailRow       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngDetailRow = LngSelectRow
        .Row = LngDetailRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
    End With
End Sub

Private Sub Msf待发明细_GotFocus()
    With Msf待发明细
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf待发明细_LostFocus()
    With Msf待发明细
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub


Private Sub tabShow_Click(PreviousTab As Integer)
    Select Case TabShow.Tab
    Case 0
        Msf待发明细.ZOrder
        Msf待发明细_EnterCell
    Case 1
'        Call RefreshData
        WriteTotalDataToBill
        Msf待发汇总.ZOrder
        Msf待发汇总_EnterCell
    End Select
End Sub

Private Sub TxtNo_GotFocus()
    GetFocus TxtNo
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
       
    '--如果不满八位,则按规则产生--
    Me.TxtNo = UCase(LTrim(Me.TxtNo))
    Me.TxtNo.Text = GetFullNO(Me.TxtNo.Text, 13)
    
    If mstrDeptNode = "" Or Cbo药房.Tag <> mstrDeptNode Then
        mstrDeptNode = GetDeptStationNode(Val(Cbo药房.ItemData(Cbo药房.ListIndex)))
        Cbo药房.Tag = mstrDeptNode
    End If
    
    Call GetRecipe(1, TxtNo)
End Sub

Private Function CheckSource(ByVal Int单据 As Integer, ByVal strNo As String, ByVal lng药房ID As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim bln重复部门 As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "Select B.编码 as 编码,B.名称 as 来源部门 From 药品收发记录 A,部门表 B Where A.对方部门id=B.id and No=[1] And 单据=[2] " & _
          " And Mod(记录状态,3)=1 And 审核人 Is Null And (库房ID+0=[3] Or 库房ID Is NULL) And Rownum<2"
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, Int单据, lng药房ID)
    
    If rs.RecordCount = 0 Then
        CheckSource = False
        Exit Function
    End If
    
    With rs处方来源部门
        If .RecordCount = 0 Then
            .AddNew
            !编码 = rs!编码
            !来源部门 = rs!来源部门
            CheckSource = True
        Else
            .MoveFirst
            For n = 1 To .RecordCount
                If !编码 = rs!编码 Then
                    bln重复部门 = True
                    Exit For
                End If
                .MoveNext
            Next
            If Not bln重复部门 Then
                If MsgBox("当前处方的开单科室是[" & rs!编码 & "]" & rs!来源部门 & "，你确定要加入该处方吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                Else
                    .AddNew
                    !编码 = rs!编码
                    !来源部门 = rs!来源部门
                    CheckSource = True
                End If
            Else
                CheckSource = True
            End If
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReadBillData(ByVal BillStyle As Integer, ByVal BillNo As String, ByVal lng药房ID As Long, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Boolean
    Dim IntStyle As Integer
    Dim str序号 As String
    Dim str明细单位串 As String
    '--读取单据内容--
    'BillStyle-单据类型;BIllNO-单据号
    '单位显示根据服务对象来（门诊：门诊单位；住院或住院门诊：住院单位；其它；售价单位）
'    On Error Resume Next
'    err = 0
    On Error GoTo errHandle
    ReadBillData = False
    
    strUnit = GetUnit(lng药房ID, BillStyle, BillNo, int门诊标志)
    Select Case strUnit
    Case "售价单位"
        str明细单位串 = "C.计算单位 单位,B.零售价 单价,B.实际数量*Nvl(B.付数,1) 数量"
    Case "门诊单位"
        str明细单位串 = "D.门诊单位 单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装) 单价,B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装)*Nvl(B.付数,1) 数量"
    Case "住院单位"
        str明细单位串 = "D.住院单位 单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装) 单价,B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装)*Nvl(B.付数,1) 数量"
    Case "药库单位"
        str明细单位串 = "D.药库单位 单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装) 单价,B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装)*Nvl(B.付数,1) 数量"
    End Select
    str明细单位串 = str明细单位串 & ",B.零售金额 金额,Nvl(B.付数, 1) * B.实际数量 / (Nvl(F.付数, 1) * F.数次) * F.实收金额 As 实收金额 "
    
    gstrSQL = " SELECT DISTINCT F.序号,F.病人ID,'['||C.编码||']'|| " & IIf(gint药品名称显示 = 1, "Nvl(A.名称,C.名称)", "C.名称") & " As 品名,A.名称 AS 商品名, " & _
        " DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格," & _
        str明细单位串 & _
        " FROM 药品收发记录 B,药品规格 D,收费项目目录 C,收费项目别名 A,门诊费用记录 F" & _
        " WHERE B.药品ID=D.药品ID AND D.药品ID=C.ID And B.费用ID=F.ID" & _
        " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 " & _
        " AND MOD(B.记录状态,3)=1 AND B.NO=[1] AND B.单据=[2] " & _
        " AND (B.库房ID+0=[3] OR B.库房ID IS NULL) " & _
        " And 审核人 Is Null" & _
        " Order by F.序号"
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    Set RecBill = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, BillNo, BillStyle, lng药房ID)
    
    With RecBill
        str序号 = ""
        Do While Not .EOF
            str序号 = str序号 & "," & !序号
            .MoveNext
        Loop
        If str序号 <> "" Then str序号 = Mid(str序号, 2)
        .MoveFirst
    End With
    
    '将单据信息与明细序号写入内部映射记录集中
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        .Find "单据标识='" & BillNo & "|" & BillStyle & "'"
        If str序号 <> "" Then
            If .EOF Then
                .AddNew
                !单据标识 = BillNo & "|" & BillStyle
                !序号 = str序号
                !记录性质 = int记录性质
                !门诊标志 = int门诊标志
                .Update
            End If
        End If
    End With
    
    If WriteDataToBill() = False Then Exit Function

    If err <> 0 Then
        MsgBox "读取处方时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    ReadBillData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBill(ByVal intRow As Integer, ByVal IntBillStyle As Integer, ByVal strNo As String, ByVal lng药房ID As Long, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer) As Integer
    Dim RecCheck As New ADODB.Recordset

    '--根据将要执行的操作，判断是否允许--
    '返回:
    '0-允许操作
    '1-未配药
    '2-已配药
    '3-已发药
    '4-已删除
    '5-未发药
    On Error GoTo errHandle
    gstrSQL = " Select A.配药人,A.审核人,nvl(B.已收费,0) 已收费, C.操作员姓名 填制人 " & _
            " From 药品收发记录 A,未发药品记录 B, 门诊费用记录 C " & _
            " Where A.No=B.No And A.单据=B.单据 And A.费用id = C.ID And mod(A.记录状态,3)=1 And A.审核人 IS Null And Rownum=1 " & _
            " And A.No=[1] And A.单据=[2] And (A.库房ID+0=[3] Or A.库房ID Is NULL)"
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    Set RecCheck = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, IntBillStyle, lng药房ID)
        
    With RecCheck
        If .EOF Then CheckBill = 4: MsgBox "未找到处方[" & strNo & "],可能已经被其他操作员处理,操作被迫中止！", vbInformation, gstrSysName: Exit Function
        If Not IsNull(!审核人) Then
            CheckBill = 3: MsgBox "该处方[" & strNo & "]已被其它操作员发药，发药操作中止！", vbInformation, gstrSysName: Exit Function
        End If
        
        '更新已收费标志
        Msf待发列表.TextMatrix(intRow, 13) = !已收费
    End With

    CheckBill = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteSendListData(ByVal intType As Integer) As Boolean
    'intType：0-录入单处方方式；1-批量提取方式
    Dim RecCheck As New ADODB.Recordset
    Dim i As Integer
    Dim blnContinue As Boolean
    
    WriteSendListData = False
    
    mstr汇总单据 = ""
    Do While Not (RecBill.EOF)
        blnContinue = True
        
        '单处方方式，不满足条件的就提示和禁止，批量方式时不提取对应的处方也不提示
        If mbln发其他药房处方 = False And IntSendAfterDosage = 0 Then
            If IsNull(RecBill!配药人) Then
                If intType = 0 Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                Else
                    blnContinue = False
                End If
            End If
            If Trim(RecBill!配药人) = "" Then
                If intType = 0 Then
                    MsgBox "该处方还未配药，不能执行发药操作！", vbInformation, gstrSysName
                    Exit Function
                Else
                    blnContinue = False
                End If
            End If
        End If
        
        If blnContinue = True Then
            With Msf待发列表
                .Redraw = False
                .TextMatrix(.rows - 1, 0) = RecBill!药房
                .TextMatrix(.rows - 1, 1) = RecBill!类型
                .TextMatrix(.rows - 1, 2) = RecBill!NO
                .TextMatrix(.rows - 1, 3) = IIf(IsNull(RecBill!科室), "", RecBill!科室)
                .TextMatrix(.rows - 1, 4) = IIf(IsNull(RecBill!姓名), "", RecBill!姓名)
                .TextMatrix(.rows - 1, 5) = IIf(IsNull(RecBill!住院号), "", RecBill!住院号)
                .TextMatrix(.rows - 1, 6) = IIf(IsNull(RecBill!床号), "", RecBill!床号)
                .TextMatrix(.rows - 1, 7) = IIf(IsNull(RecBill!填制人), "", RecBill!填制人)
                .TextMatrix(.rows - 1, 8) = IIf(IsNull(RecBill!开单医生), "", RecBill!开单医生)
                .TextMatrix(.rows - 1, 9) = IIf(IsNull(RecBill!填制日期), "", RecBill!填制日期)
                .TextMatrix(.rows - 1, 10) = RecBill!药房ID
                .TextMatrix(.rows - 1, 11) = RecBill!记录性质
                .TextMatrix(.rows - 1, 12) = RecBill!门诊标志
                .TextMatrix(.rows - 1, 13) = RecBill!已收费
                .TextMatrix(.rows - 1, 14) = IIf(IsNull(RecBill!病人ID), "", RecBill!病人ID)
                .RowData(.rows - 1) = RecBill!单据
    '            str单据号 = RecBill!NO
                If chk科室(0).Value = 0 And chk科室(1).Value = 0 Then
                    mstr汇总单据 = RecBill!单据 & "," & RecBill!NO
                Else
                    mstr汇总单据 = IIf(mstr汇总单据 = "", "", mstr汇总单据 & "|") & RecBill!单据 & "," & RecBill!NO
                End If
                
                .Row = .rows - 1
                .Col = 4
                .CellForeColor = zldatabase.GetPatiColor(IIf(IsNull(RecBill!病人类型), "", RecBill!病人类型))
    
                .rows = .rows + 1
                .RowData(.rows - 1) = 0
                .Redraw = True
            End With
            WriteSendListData = True
        End If
        
        RecBill.MoveNext
    Loop
    
End Function

Private Function RefreshData(ByVal lng药房ID As Long) As Boolean
    Dim intRow As Integer, intRows As Integer
    Dim arrID
    Dim StrNoThis As String, IntBillThis As Integer
    Dim str汇总单位串 As String
    Dim strTemp As String
    
    If mblnModify = False Then Exit Function
    RefreshData = False
    On Error GoTo errHandle
    '清空汇总表格
    With Msf待发汇总
        .Clear
        .rows = 2
        SetFormat (3)
    End With
 
   
    '显示汇总数据
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lng药房ID, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lng药房ID, gint门诊药房)
    Else
        strUnit = GetSpecUnit(lng药房ID, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        str汇总单位串 = "C.计算单位 单位,B.零售价 单价,Sum(B.实际数量*Nvl(B.付数,1)) 数量"
    Case "门诊单位"
        str汇总单位串 = "D.门诊单位 单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装) 单价,Sum(B.实际数量/Decode(D.门诊包装,Null,1,0,1,D.门诊包装)*Nvl(B.付数,1)) 数量"
    Case "住院单位"
        str汇总单位串 = "D.住院单位 单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装) 单价,Sum(B.实际数量/Decode(D.住院包装,Null,1,0,1,D.住院包装)*Nvl(B.付数,1)) 数量"
    Case "药库单位"
        str汇总单位串 = "D.药库单位 单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装) 单价,Sum(B.实际数量/Decode(D.药库包装,Null,1,0,1,D.药库包装)*Nvl(B.付数,1)) 数量"
    End Select
    
    str汇总单位串 = str汇总单位串 & ",Sum(B.零售金额) 金额,Sum(Nvl(B.付数, 1) * B.实际数量 / (Nvl(B.费用付数, 1) * B.数次) * B.实收金额) As 实收金额 "

    gstrSQL = "Select A.No,A.药品id,A.批次,A.零售价,A.实际数量,A.付数,A.零售金额,C.付数 As 费用付数,C.数次,C.实收金额,A.产地 From 药品收发记录 A,门诊费用记录 C,Table(f_Str2list2([1], '|', ',')) B " & _
        " Where A.NO=B.C2 And A.单据=B.C1 And Mod(A.记录状态,3)=1 And A.审核人 Is Null And (A.库房ID+0=[2] Or A.库房ID Is NULL) And A.费用id = C.Id "
    gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    gstrSQL = "Select Distinct D.*,'['||D.编码||']'|| " & IIf(gint药品名称显示 = 1, "Nvl(A.名称,D.通用名称)", "D.通用名称") & " As 品名,A.名称 AS 商品名 " & _
             " From " & _
             "     (SELECT B.No,D.药品ID,C.编码,C.名称 通用名称,NVL(B.批次,0) 批次," & _
             "     DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)) 规格," & str汇总单位串 & _
             "     FROM (" & gstrSQL & ") B," & _
             "           药品规格 D,收费项目目录 C " & _
             "     WHERE B.药品ID+0=D.药品ID AND D.药品ID=C.ID" & _
             "     GROUP BY B.No,D.药品ID,C.编码,C.名称,NVL(B.批次,0)," & _
             "     DECODE(C.规格,NULL,B.产地,DECODE(B.产地,NULL,C.规格,C.规格||'|'||B.产地)),"
    Select Case strUnit
    Case "售价单位"
        gstrSQL = gstrSQL & "C.计算单位,B.零售价"
    Case "门诊单位"
        gstrSQL = gstrSQL & "D.门诊单位,B.零售价*Decode(D.门诊包装,Null,1,0,1,D.门诊包装)"
    Case "住院单位"
        gstrSQL = gstrSQL & "D.住院单位,B.零售价*Decode(D.住院包装,Null,1,0,1,D.住院包装)"
    Case "药库单位"
        gstrSQL = gstrSQL & "D.药库单位,B.零售价*Decode(D.药库包装,Null,1,0,1,D.药库包装)"
    End Select
    gstrSQL = gstrSQL & ") D,收费项目别名 A" & _
            " Where D.药品ID=A.收费细目ID(+) AND A.性质(+)=3"
    gstrSQL = gstrSQL & " Order By D.编码"
    
    
    If Len(mstr汇总单据) > 4000 Then
        For intRow = 0 To UBound(GetArrayByStr(mstr汇总单据, 3900, "|"))
            
            Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(GetArrayByStr(mstr汇总单据, 3900, "|")(intRow)), lng药房ID)
            Call WriteTotalDataToBill(intRow > 0, Not (intRow = UBound(GetArrayByStr(mstr汇总单据, 3900, "|"))))
        Next
    Else
        Set RecTotal = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr汇总单据, lng药房ID)
        Call WriteTotalDataToBill
    End If
    
    If err <> 0 Then
        MsgBox "显示汇总数据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mblnModify = False
    RefreshData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteTotalDataToBill(Optional ByVal blnFirst As Boolean, Optional ByVal blnLast As Boolean) As Boolean
    Dim dbl应收金额 As Double
    Dim dbl实收金额 As Double
    Dim str金额显示 As String
    
    On Error GoTo errHandle
    
    '将汇总数据装入
    If Not blnFirst Then
        With Msf待发汇总
            .Redraw = False
            .Clear
            .rows = 2
            Call SetFormat(3)
            .Redraw = True
        End With
    End If
    
    '填充单据内容
    
    If RecTotal.State = 0 Then Exit Function
    
    If RecTotal.RecordCount > 0 Then
'        If Not RecTotal.EOF Then Call InitRecSum
    
        Do While Not RecTotal.EOF
            With rs待发汇总明细
                .AddNew
                !单据号 = RecTotal!NO
                !药品名称 = RecTotal!品名
                !商品名 = IIf(IsNull(RecTotal!商品名), "", RecTotal!商品名)
                !编码 = RecTotal!编码
                !规格 = IIf(IsNull(RecTotal!规格), "", RecTotal!规格)
                !单位 = IIf(IsNull(RecTotal!单位), "", RecTotal!单位)
                !单价 = RecTotal!单价
                !数量 = RecTotal!数量
                !金额 = RecTotal!金额
                !实收金额 = RecTotal!实收金额
                !药品id = RecTotal!药品id
                !批次 = RecTotal!批次
            End With
            RecTotal.MoveNext
        Loop
    End If
    
    If blnLast Then Exit Function
    
    With rs待发汇总明细
        If .RecordCount <> 0 Then
            .Sort = "编码,批次"
            .MoveFirst
        End If
        Do While Not .EOF
            If Msf待发汇总.rows = 2 And Msf待发汇总.TextMatrix(1, 1) = "" Then
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 0) = Msf待发汇总.rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 1) = !药品名称
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 2) = !商品名
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 3) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 4) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 5) = Format(!单价, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = Format(!数量, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = Format(!金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = Format(!实收金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 9) = !药品id
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 10) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.rows - 1) = False
            ElseIf Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 9) <> !药品id Then
                Msf待发汇总.rows = Msf待发汇总.rows + 1
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 0) = Msf待发汇总.rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 1) = !药品名称
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 2) = !商品名
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 3) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 4) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 5) = Format(!单价, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = Format(!数量, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = Format(!金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = Format(!实收金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 9) = !药品id
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 10) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.rows - 1) = False
            ElseIf Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 10) <> !批次 Then
                Msf待发汇总.rows = Msf待发汇总.rows + 1
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 0) = Msf待发汇总.rows - 1
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 1) = !药品名称
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 2) = !商品名
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 3) = IIf(IsNull(!规格), "", !规格)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 4) = IIf(IsNull(!单位), "", !单位)
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 5) = Format(!单价, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = Format(!数量, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = Format(!金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = Format(!实收金额, "#####0.00;-#####0.00; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 9) = !药品id
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 10) = !批次
                Msf待发汇总.MergeRow(Msf待发汇总.rows - 1) = False
            Else
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = Format(CDbl(Val(Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6))) + !数量, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = Format(CDbl(Val(Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7))) + !金额, "#####0.00000;-#####0.00000; ;")
                Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = Format(CDbl(Val(Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8))) + !实收金额, "#####0.00000;-#####0.00000; ;")
            End If
            dbl应收金额 = dbl应收金额 + !金额
            dbl实收金额 = dbl实收金额 + !实收金额
            .MoveNext
        Loop
        
        '显示合计
        Msf待发汇总.rows = Msf待发汇总.rows + 1
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 0) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 1) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 2) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 3) = "合计"
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 4) = "合计"
        
        If mint金额显示 = 1 Then
            str金额显示 = "实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        ElseIf mint金额显示 = 2 Then
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;") & "    实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        Else
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;")
        End If
        
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 5) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 6) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 7) = str金额显示
        Msf待发汇总.TextMatrix(Msf待发汇总.rows - 1, 8) = str金额显示
        
        Msf待发汇总.MergeCells = flexMergeFree
        Msf待发汇总.MergeRow(Msf待发汇总.rows - 1) = True
    End With

    WriteTotalDataToBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteDataToBill() As Boolean
    Dim dbl应收金额 As Double
    Dim dbl实收金额 As Double
    Dim str金额显示 As String
    
    '--显示指定处方的明细--
    On Error Resume Next
    err = 0
    
    WriteDataToBill = False
    With Msf待发明细
        .Clear
        .rows = 2
        Call SetFormat(2)
    End With
    
    '填充单据内容
    With RecBill
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            Msf待发明细.MergeRow(.AbsolutePosition) = False
            Msf待发明细.TextMatrix(.AbsolutePosition, 0) = !品名
            Msf待发明细.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!商品名), "", !商品名)
            Msf待发明细.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!规格), "", !规格)
            Msf待发明细.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!单位), "", !单位)
            Msf待发明细.TextMatrix(.AbsolutePosition, 4) = Format(!单价, "#####0.00000;-#####0.00000; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 5) = Format(!数量, "#####0.00000;-#####0.00000; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 6) = Format(!金额, "#####0.00;-#####0.00; ;")
            Msf待发明细.TextMatrix(.AbsolutePosition, 7) = Format(!实收金额, "#####0.00;-#####0.00; ;")
            dbl应收金额 = dbl应收金额 + Val(!金额)
            dbl实收金额 = dbl实收金额 + Val(!实收金额)
            
            If .AbsolutePosition >= Msf待发明细.rows - 1 Then Msf待发明细.rows = Msf待发明细.rows + 1
            .MoveNext
        Loop
    End With
    With Msf待发明细
        .TextMatrix(.rows - 1, 0) = "合计"
        .TextMatrix(.rows - 1, 1) = "合计"
        .TextMatrix(.rows - 1, 2) = "合计"
        .TextMatrix(.rows - 1, 3) = "合计"
        
        If mint金额显示 = 1 Then
            str金额显示 = "实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        ElseIf mint金额显示 = 2 Then
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;") & "    实收金额：" & Format(dbl实收金额, "#####0.00;-#####0.00; ;")
        Else
            str金额显示 = "应收金额：" & Format(dbl应收金额, "#####0.00;-#####0.00; ;")
        End If
        
        .TextMatrix(.rows - 1, 4) = str金额显示
        .TextMatrix(.rows - 1, 5) = str金额显示
        .TextMatrix(.rows - 1, 6) = str金额显示
        .TextMatrix(.rows - 1, 7) = str金额显示
        .MergeCells = flexMergeFree
        .MergeRow(.rows - 1) = True
    End With
    
    If err <> 0 Then
        MsgBox "显示单据时，发生不可预知的错误！", vbInformation, gstrSysName
        Exit Function
    End If
    WriteDataToBill = True
End Function

Private Function SetLocateBill(Optional ByVal BlnEnterCell As Boolean = True) As Boolean
    Dim intRow As Integer
    
    SetLocateBill = False
    With Msf待发列表
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 2) = TxtNo And TxtNo.Tag = .RowData(intRow) Then
                .Row = intRow
                .TopRow = intRow
                SetLocateBill = True
                Exit For
            End If
        Next
    End With
    
    If BlnEnterCell Then Msf待发列表_EnterCell
End Function

Private Function CheckStock() As Boolean
    Dim RecCheckStock As New ADODB.Recordset
    Dim dblStock As Double
    Dim strSubSql As String
    Dim n As Integer
    Dim lngTemp库房id As Long
    Dim dblUsableStock As Double
    
    On Error GoTo errHandle
    If mbln发其他药房处方 = False Then
        lngTemp库房id = Cbo药房.ItemData(Cbo药房.ListIndex)
    Else
        lngTemp库房id = mlng药房ID
    End If
    
    IntCheckStock = MediWork_GetCheckStockRule(lngTemp库房id)
    
    '检查库存
    If IntCheckStock = 0 Then CheckStock = True: Exit Function
    
    '将库存数量转换为对应单位的实际数量
    Dim intUnit As Integer
    intUnit = Val(zldatabase.GetPara("药房属性", glngSys, 1341, 0))
    If intUnit = 0 Then
        strUnit = GetDrugUnit(lngTemp库房id, "", True)
    ElseIf intUnit = 1 Then
        strUnit = GetSpecUnit(lngTemp库房id, gint门诊药房)
    Else
        strUnit = GetSpecUnit(lngTemp库房id, gint住院药房)
    End If
    Select Case strUnit
    Case "售价单位"
        strSubSql = "/1"
    Case "门诊单位"
        strSubSql = "/Decode(B.门诊包装,Null,1,0,1,B.门诊包装)"
    Case "住院单位"
        strSubSql = "/Decode(B.住院包装,Null,1,0,1,B.住院包装)"
    Case "药库单位"
        strSubSql = "/Decode(B.药库包装,Null,1,0,1,B.药库包装)"
    End Select
    
    CheckStock = False
    If Msf待发列表.TextMatrix(1, 2) <> "" Then
        For n = 1 To Msf待发汇总.rows - 2
            gstrSQL = " Select nvl(可用数量,0)" & strSubSql & " AS 可用数量, nvl(实际数量,0)" & strSubSql & " AS 实际数量 " & _
                         " From 药品库存 A,药品规格 B" & _
                         " Where B.药品ID=A.药品ID And A.性质=1 And A.库房ID=[1] And A.药品ID=[2] And Nvl(A.批次,0)=[3]"
            Set RecCheckStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngTemp库房id, Val(Msf待发汇总.TextMatrix(n, 9)), Val(Msf待发汇总.TextMatrix(n, 10)))
                         
            With RecCheckStock
                If .EOF Then
                    dblStock = 0
                    dblUsableStock = 0
                Else
                    dblStock = !实际数量
                    dblUsableStock = !可用数量
                End If
                
                '如果是代发其他药房处方，则既要检查实际数量，也要检查可用数量
                If dblStock < Val(Msf待发汇总.TextMatrix(n, 6)) Or (mbln发其他药房处方 = True And dblUsableStock < Val(Msf待发汇总.TextMatrix(n, 6))) Then
                    If Msf待发汇总.TextMatrix(n, 10) <> 0 Then
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(Msf待发汇总.TextMatrix(n, 1) & "的批次库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf待发汇总.TextMatrix(n, 1) & "的批次库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                        End Select
                    Else
                        Select Case IntCheckStock
                        Case 1
                            If MsgBox(Msf待发汇总.TextMatrix(n, 1) & "的库存数不够，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                        Case 2
                            MsgBox Msf待发汇总.TextMatrix(n, 1) & "的库存数不够，不能继续发药！", vbInformation, gstrSysName: Exit Function
                        End Select
                    End If
                End If
            End With
        Next
    End If
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitApplyforcredit()
    '存在销帐申请的记录集
    Set mrsApplyforcredit = New ADODB.Recordset
    With mrsApplyforcredit
        If .State = 1 Then .Close
        
        .Fields.Append "费用ID", adDouble, 18, adFldIsNullable
        .Fields.Append "收发ID", adDouble, 18, adFldIsNullable              '药品收发ID
        .Fields.Append "标志", adDouble, 1, adFldIsNullable      '0-不允许该单据发药；1-允许该单据发药
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "销帐申请数量", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 10, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Function CheckNotAudited(ByRef rsData As ADODB.Recordset) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim bln销帐申请 As Boolean
    Dim bln允许发送 As Boolean
    Dim str明细单位串 As String
    
    On Error GoTo errHandle
    
    Call InitApplyforcredit
    
    CheckNotAudited = True
    bln销帐申请 = True
    
    '检测当前药房是否为住院药房，不是则退出此项检查
    gstrSQL = "Select *" & vbNewLine & _
            "From 部门表 A, 部门性质说明 B" & vbNewLine & _
            "Where a.Id = b.部门id And a.Id = [1] And (b.工作性质 Like '%药房' Or b.工作性质 Like '%药库') And b.服务对象 In (2, 3)"

    Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检测当前药房是否为住院药房", Cbo药房.ItemData(Cbo药房.ListIndex))
    If rsTmp.EOF Then Exit Function
    
    Select Case strUnit
    Case "门诊单位"
        str明细单位串 = "e.门诊包装 as 包装,e.门诊单位 as 单位"
    Case "住院单位"
        str明细单位串 = "e.住院包装 as 包装,e.住院单位 as 单位"
    Case "药库单位"
        str明细单位串 = "e.药库包装 as 包装,e.药库单位 as 单位"
    Case Else
        str明细单位串 = "e.门诊包装 as 包装,e.门诊单位 as 单位"
    End Select
    
    gstrSQL = " select A.*,b.姓名,b.性别,b.年龄,c.数量 As 销帐申请数量,decode(d.名称,null,f.名称,d.名称) as 药名,e.门诊包装 as 包装,e.门诊单位 as 单位" & vbNewLine & _
            "  from (select distinct id as 收发id,费用id, 药品id, 批号, 实际数量 from 药品收发记录 where No =  [1] and mod(记录状态, 3) = 1 and 审核日期 is null) A," & vbNewLine & _
            "住院费用记录 B,病人费用销帐 C,药品别名 D, 药品规格 E,收费项目目录 F" & vbNewLine & _
            " where a.费用id = b.Id And b.Id = c.费用id and a.药品id = d.药品id(+) and e.药品id = a.药品id and e.药品id = f.id"

    
    With rsData
        rsData.Sort = "单据,NO"
    
        Do While Not .EOF
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否存在销帐申请未审核的单据", rsData!NO)

            If rsTmp.RecordCount > 0 Then
                bln销帐申请 = False

                With mrsApplyforcredit
                    Do While Not rsTmp.EOF
                        .AddNew
                        
                        !标志 = 1
                        !NO = rsData!NO
                        !药品名称 = rsTmp!药名
                        !批号 = rsTmp!批号
                        !数量 = Format(rsTmp!实际数量 / rsTmp!包装, "#####0.0000;-#####0.0000; ;") & rsTmp!单位
                        !销帐申请数量 = Format(rsTmp!销帐申请数量 / rsTmp!包装, "#####0.0000;-#####0.0000; ;") & rsTmp!单位
                        !姓名 = rsTmp!姓名
                        !性别 = rsTmp!性别
                        !年龄 = rsTmp!年龄
                        !费用ID = rsTmp!费用ID
                        !收发ID = rsTmp!收发ID
                        
                        rsTmp.MoveNext
                    Loop
                End With

            End If

            .MoveNext
        Loop
    End With

    '对含有销帐申请的单据进行处理
    If bln销帐申请 = False Then
        Call frm部门发药销帐申请清单.ShowCard(Me, mrsApplyforcredit, bln允许发送, 1)

        '由子窗体返回用户是否继续执行操作，若【取消】则禁止继续发送
        CheckNotAudited = bln允许发送
        If CheckNotAudited = False Then Exit Function
        
        '修正取消发送的单据的执行状态
        mrsApplyforcredit.Filter = "标志 = 0"
        
        If mrsApplyforcredit.RecordCount > 0 Then
            Do While Not mrsApplyforcredit.EOF
                rsData.Filter = "No = '" & mrsApplyforcredit!NO & "'"
                If rsData.RecordCount > 0 Then
                    rsData.Delete
                    rsData.Update
                End If
                mrsApplyforcredit.MoveNext
            Loop
        End If

        rsData.Filter = ""
    End If
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendBill() As Boolean
    Dim intRow As Integer
    Dim StrDate As String
    Dim rsSendRecipeByNo As ADODB.Recordset
    Dim int门诊 As Integer
    Dim arrSql As Variant
    Dim i As Long
    Dim blnInTrans As Boolean
    Dim str签名记录 As String
    Dim strReturn As String
    Dim strNo As String
    Dim cur汇总发药号 As Currency
    Dim strReserve As String
    Dim lng药房ID As Long
    
    On Error GoTo ErrHand
    
    arrSql = Array()

    SendBill = False
    
    Set rsSendRecipeByNo = New ADODB.Recordset
    With rsSendRecipeByNo
        If .State = 1 Then .Close
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "单据", adDouble, 18, adFldIsNullable
        .Fields.Append "药房ID", adDouble, 18, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        .Fields.Append "已收费", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "填制日期", adDate, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    StrDate = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
    cur汇总发药号 = Val(zldatabase.GetNextNo(20))
    
    With Msf待发列表
        For intRow = 1 To .rows - 1
            If .RowData(intRow) <> 0 Then
                '检查处方，防止并发
                If CheckBill(intRow, .RowData(intRow), .TextMatrix(intRow, 2), Val(.TextMatrix(intRow, 10)), Val(.TextMatrix(intRow, 11)), Val(.TextMatrix(intRow, 12))) <> 0 Then
                    Exit Function
                End If
                
                '零差价管理
                If CheckPriceAdjustByNO(Val(.RowData(intRow)), Val(.TextMatrix(intRow, 10)), .TextMatrix(intRow, 2), IIf(mbln发其他药房处方, mlng药房ID, 0)) = False Then
                    Exit Function
                End If
                
                With rsSendRecipeByNo
                    .AddNew
                    !NO = Msf待发列表.TextMatrix(intRow, 2)
                    !单据 = Msf待发列表.RowData(intRow)
                    !药房ID = Val(Msf待发列表.TextMatrix(intRow, 10))
                    !记录性质 = Val(Msf待发列表.TextMatrix(intRow, 11))
                    !门诊标志 = Val(Msf待发列表.TextMatrix(intRow, 12))
                    !已收费 = Val(Msf待发列表.TextMatrix(intRow, 13))
                    !病人ID = Val(Msf待发列表.TextMatrix(intRow, 14))
                    !填制日期 = Msf待发列表.TextMatrix(intRow, 9)
                    .Update
                End With
            End If
        Next
    End With
    
    '检查[住院单据]是否存在销帐申请未审核的单据
    If CheckNotAudited(rsSendRecipeByNo) = False Then Exit Function

    '按处方号排序后批量发药
    rsSendRecipeByNo.Sort = "NO"
    rsSendRecipeByNo.MoveFirst
    For intRow = 1 To rsSendRecipeByNo.RecordCount
        '先检查或执行预调价
        Call AutoAdjustPrice_ByNO(rsSendRecipeByNo!单据, rsSendRecipeByNo!NO)
        
        If Val(rsSendRecipeByNo!记录性质) = 1 Or (Val(rsSendRecipeByNo!记录性质) = 2 And (Val(rsSendRecipeByNo!门诊标志) = 1 Or Val(rsSendRecipeByNo!门诊标志) = 4)) Then
            int门诊 = 1
        Else
            int门诊 = 2
        End If
        
        '批量发药暂时不支持刷卡消费
'        '未收费的划价单
'        If rsSendRecipeByNo!单据 = 8 And rsSendRecipeByNo!已收费 = 0 And mint允许未收费处方发药 = 0 Then
'            If Not gobjSquareCard Is Nothing And mbln发药前收费或审核 = True Then
'                '刷卡收费
'                If gobjSquareCard.zlSquareAffirm(Me, 1341, strPrivs, rsSendRecipeByNo!病人id, 0, False, 1, rsSendRecipeByNo!NO) = False Then
'                    Exit Function
'                End If
'            Else
'                MsgBox "该处方[" & rsSendRecipeByNo!NO & "]还未收费或收费程序错误，发药操作中止！", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
'
'        '未审核的记账划价单
'        If rsSendRecipeByNo!单据 = 9 And rsSendRecipeByNo!已收费 = 0 And Int允许未审核处方发药 = 0 Then
'            If Not gobjSquareCard Is Nothing And mbln发药前收费或审核 = True Then
'                '刷卡收费
'                If gobjSquareCard.zlSquareAffirm(Me, 1341, strPrivs, rsSendRecipeByNo!病人id, 0, False, 2, rsSendRecipeByNo!NO) = False Then
'                    Exit Function
'                End If
'            Else
'                MsgBox "该处方[" & rsSendRecipeByNo!NO & "]还未审核或审核程序错误，发药操作中止！", vbInformation, gstrSysName
'                Exit Function
'            End If
'        End If
        
        If mbln发其他药房处方 Then
            gstrSQL = "Zl_药品收发记录_更改库房("
            '现库房ID
            gstrSQL = gstrSQL & mlng药房ID
            '单据
            gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
            'NO
            gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
            '原库房ID
            gstrSQL = gstrSQL & "," & Val(rsSendRecipeByNo!药房ID)
            '门诊
            gstrSQL = gstrSQL & "," & int门诊
            '填制日期
            gstrSQL = gstrSQL & ",to_date('" & rsSendRecipeByNo!填制日期 & "','yyyy-MM-dd')"
            gstrSQL = gstrSQL & ")"

            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
        
        gstrSQL = "zl_药品收发记录_处方发药("
        '库房ID
        gstrSQL = gstrSQL & mlng药房ID
        '单据
        gstrSQL = gstrSQL & "," & rsSendRecipeByNo!单据
        'NO
        gstrSQL = gstrSQL & ",'" & rsSendRecipeByNo!NO & "'"
        '审核人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '配药人
        gstrSQL = gstrSQL & "," & IIf(IntSendAfterDosage = 0, "NULL", "'" & str配药人 & "'")
        '校验人
        gstrSQL = gstrSQL & ",NULL"
        '发药方式
        gstrSQL = gstrSQL & ",2"
        '发药时间
        gstrSQL = gstrSQL & ",to_date('" & StrDate & "','yyyy-MM-dd hh24:mi:ss')"
        '操作员编号
        gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
        '操作员姓名
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '金额保留位数
        gstrSQL = gstrSQL & "," & int金额保留位数
        '审核划价单
        gstrSQL = gstrSQL & "," & int审核划价单
        '是否门诊
        gstrSQL = gstrSQL & "," & int门诊
        '核查人
        gstrSQL = gstrSQL & ",'" & str核查人 & "'"
        '是否未取药
        gstrSQL = gstrSQL & ",NULL"
        '汇总发药号
        gstrSQL = gstrSQL & "," & cur汇总发药号
        
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        strNo = strNo & rsSendRecipeByNo!单据 & "," & rsSendRecipeByNo!NO & "|"
        rsSendRecipeByNo.MoveNext
    Next
    
    '调用发药前的外挂接口
    err.Clear
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        If mobjPlugIn.DrugBeforeSendByRecipe(mlng药房ID, strNo, strReserve) = False Then
            If err.Number <> 0 Then
                err.Clear: On Error GoTo 0
            Else
                Exit Function
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo ErrHand
    
    '先处理发药事务
    gcnOracle.BeginTrans
    blnInTrans = True
    
    '单独处理电子签名，放到业务处理前面，防止弹框造成数据死锁
    '如果已启用了电子签名，则需要对配药人进行电子签名处理
    If gblnESign处方发药 = True And gblnESignUserStoped = False Then
        rsSendRecipeByNo.MoveFirst
        For intRow = 1 To rsSendRecipeByNo.RecordCount
            '因为要先查询数据再更新药房，所以待发时用单据原来的药房来查询，同时需要传入现药房来产生签名原文
            lng药房ID = IIf(mbln发其他药房处方 = True, Val(rsSendRecipeByNo!药房ID), mlng药房ID)
            str签名记录 = ""
            If GetSignatureRecored(EsignTache.send, rsSendRecipeByNo!单据, rsSendRecipeByNo!NO, _
                    lng药房ID, str签名记录, 0, CDate(StrDate), gstrUserName, mlng药房ID) = False Then
                gcnOracle.RollbackTrans
                blnInTrans = False
                Exit Function
            End If
            
            If str签名记录 <> "" Then
                gstrSQL = "Zl_药品签名记录_Insert(" & str签名记录 & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gcnOracle.RollbackTrans
                blnInTrans = False
                MsgBox "对发药人电子签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
            
            rsSendRecipeByNo.MoveNext
        Next
    End If
    
    For i = 0 To UBound(arrSql)
        Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "RecipeWork_Abolish")
    Next
    
    gcnOracle.CommitTrans
    blnInTrans = False
    
    If TypeName(mobjDrugMAC) = "clsDrugPacker" Then
        If mblnConPacker And strNo <> "" And mblnLoadDrug And Not mbln发其他药房处方 Then
            Call mobjDrugMAC.DYEY_MZ_TransRecipeList(mstrOpr, UserInfo.用户编码, UserInfo.用户姓名, mlng药房ID, Mid(strNo, 1, Len(strNo) - 1), strReturn)
        End If
    ElseIf TypeName(mobjDrugMAC) = "clsDrugMachine" Then
        If mblnConPacker Then
            If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
            mobjDrugMAC.Operation gstrDbUser, Val("22-开始发药"), "1|" & Replace(strNo, "|", ";"), strReturn
        End If
    End If
        
    If MsgBox("你需要打印汇总清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_BILL_1341_2", "ZL8_BILL_1341_2"), Me, "库房=" & IIf(mbln发其他药房处方 = True, mlng药房ID, Cbo药房.ItemData(Cbo药房.ListIndex)), "发药方式=批量发药|2", "包装系数=" & IIf(strUnit = "门诊单位", "D.门诊包装", "D.住院包装"), "发药时间=" & StrDate, 2)
    End If
    
    '调用发药后的外挂接口
    If Not mobjPlugIn Is Nothing Then
        If Right(strNo, 1) = "|" Then strNo = Left(strNo, Len(strNo) - 1)
        On Error Resume Next
        mobjPlugIn.DrugSendByRecipe mlng药房ID, strNo, CDate(StrDate), strReserve
        err.Clear: On Error GoTo 0
    End If
    
    SendBill = True
    Exit Function
ErrHand:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckCorrelation() As Boolean
    Dim strNo As String, lng单据 As Long, str序号 As String
    '检查处方是否已结帐、检查该病人是否已出院，并对权限进行检查
    With rs序号
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !单据标识
            lng单据 = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            str序号 = nvl(!序号)
            If Not IsReceiptBalance_Charge(0, strPrivs, lng单据, strNo, str序号, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            If Not IsOutPatient(strPrivs, lng单据, strNo, Val(!记录性质), Val(!门诊标志)) Then Exit Function
            .MoveNext
        Loop
    End With
    
    CheckCorrelation = True
End Function

Private Sub InitRec()
    Set rs序号 = New ADODB.Recordset
    With rs序号
        If .State = 1 Then .Close
        .Fields.Append "单据标识", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "序号", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "记录性质", adDouble, 18, adFldIsNullable
        .Fields.Append "门诊标志", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rs处方来源部门 = New ADODB.Recordset
    With rs处方来源部门
        If .State = 1 Then .Close
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "来源部门", adLongVarChar, 100, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Call InitRecSum
    
End Sub

Private Sub txt病人科室_GotFocus()
    GetFocus txt病人科室
End Sub

Private Sub txt病人科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt病人科室.Text) = "" Then Exit Sub
    
    If Select部门(Me, Me.txt病人科室, Trim(txt病人科室.Text), "临床", False, mint服务对象) = False Then
        Exit Sub
    End If
End Sub

Private Function Select部门(ByVal FrmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln操作员 As Boolean = False, _
    Optional ByVal int服务对象 As Integer, _
    Optional strsql As String = "") As Boolean
    '------------------------------------------------------------------------------
    '功能:部门选择器
    '参数:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln操作员-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '返回:成功,返回true,否则返回False
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rstemp  As ADODB.Recordset
    Dim strComment As String
    
    On Error GoTo errHandle
    
    strTittle = "部门选择器"
    vRect = zlControl.GetControlRect(objCtl.hWnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strsql <> "" Then
    
        gstrSQL = strsql
    Else
        gstrSQL = "" & _
        "   Select /*+ Rule*/ distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
        "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
    
        If str工作性质 = "" And bln操作员 = False Then
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c," & _
            IIf(str工作性质 = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.名称=J.column_value ") & _
            "         AND a.id = c.部门id and" & IIf(int服务对象 <> 3, " c.服务对象=[4] ", " (c.服务对象=1 or c.服务对象=2 or c.服务对象=[4])") & _
            IIf(bln操作员 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) And (a.站点=[5] or a.站点 is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            If Mid(gtype_UserSysParms.Para_输入方式, 1, 1) = "1" Then strFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            If Mid(gtype_UserSysParms.Para_输入方式, 2, 1) = "1" Then strFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [3] "
        End If
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strsql = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.编码"
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strsql = "" Then
        '分上下级
        Set rstemp = zldatabase.ShowSQLMultiSelect(FrmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, int服务对象)
    Else
        Set rstemp = zldatabase.ShowSQLMultiSelect(FrmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, glngUserId, str工作性质, strKey, int服务对象, gstrNodeNo)
    End If
    
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If rstemp Is Nothing Then
        MsgBox "没有满足条件的部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    Call zlCtlSetFocus(objCtl, True)
    
    objCtl.Text = ""
    objCtl.Tag = ""
    
    For i = 1 To rstemp.RecordCount
        If i = 1 Then
            objCtl.Text = zlStr.nvl(rstemp!编码) & "-" & zlStr.nvl(rstemp!名称)
        ElseIf i = 2 Then
            objCtl.Text = objCtl.Text & "..."
        End If
        
        strComment = IIf(strComment = "", "", strComment & ",") & zlStr.nvl(rstemp!编码) & "-" & zlStr.nvl(rstemp!名称)
        
        '部门ID保存到Tag属性
        objCtl.Tag = IIf(objCtl.Tag = "", "", objCtl.Tag & ",") & Val(rstemp!Id)
        
        rstemp.MoveNext
    Next
    
    objCtl.ToolTipText = strComment
        
    zlCommFun.PressKey vbKeyTab
    
    Select部门 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Txt姓名_GotFocus()
    GetFocus Txt姓名
End Sub


Private Sub Txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(3, Txt姓名)
End Sub


Private Sub txt医保号_GotFocus()
    GetFocus txt医保号
End Sub

Private Sub txt医保号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call GetRecipe(2, txt医保号)
End Sub


