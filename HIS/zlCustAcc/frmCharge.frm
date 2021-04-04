VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCharge 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "记帐处理"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "frmCharge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6975
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraCancel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   8460
      TabIndex        =   7
      Top             =   5730
      Width           =   2115
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "取消(&C)"
         Height          =   420
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "热键:Esc"
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame fraOK 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   6540
      TabIndex        =   6
      Top             =   5640
      Width           =   2025
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00C0C0C0&
         Caption         =   "确定(&O)"
         Height          =   420
         Left            =   510
         TabIndex        =   8
         ToolTipText     =   "热键：F2"
         Top             =   330
         Width           =   1275
      End
   End
   Begin VB.Frame fra时间 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   7830
      TabIndex        =   5
      Top             =   4710
      Width           =   2265
      Begin MSMask.MaskEdBox txtDate 
         Height          =   360
         Left            =   240
         TabIndex        =   11
         Top             =   210
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         HideSelection   =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "yyyy-MM-dd hh:mm:ss"
         Mask            =   "####-##-## ##:##:##"
         PromptChar      =   "_"
      End
   End
   Begin VB.Frame fra开单人 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   5460
      TabIndex        =   4
      Top             =   4830
      Width           =   2265
      Begin VB.ComboBox cbo开单人 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   60
         TabIndex        =   10
         Top             =   120
         Width           =   2085
      End
   End
   Begin VB.Frame fra销 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   9600
      TabIndex        =   3
      Top             =   180
      Width           =   855
      Begin VB.CheckBox chk销 
         Caption         =   "销"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "热键:F8"
         Top             =   150
         Width           =   405
      End
   End
   Begin VB.Frame fraNO 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   8070
      TabIndex        =   1
      Top             =   180
      Width           =   1605
      Begin VB.ComboBox cboNO 
         ForeColor       =   &H00C00000&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   1425
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6615
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCharge.frx":08CA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11800
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   88
            Key             =   "病人余额"
            Object.ToolTipText     =   "病人余额"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   318
            MinWidth        =   71
            Key             =   "MedicareType"
            Object.ToolTipText     =   "医保大类"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":115E
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCharge.frx":1798
            Key             =   "WB"
            Object.ToolTipText     =   $"frmCharge.frx":1DD2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
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
   Begin VB.Frame fraForm 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6255
      Left            =   60
      TabIndex        =   13
      Top             =   90
      Width           =   11205
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   4755
         Width           =   1800
      End
      Begin VB.CheckBox chk附加 
         Caption         =   "附加手术"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   9045
         TabIndex        =   38
         Top             =   2498
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Index           =   0
         Left            =   7710
         TabIndex        =   37
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt实收金额 
         Height          =   300
         Index           =   0
         Left            =   6675
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txt应收金额 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txt标准单价 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   4965
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.TextBox txt计算单位 
         Height          =   300
         Index           =   0
         Left            =   3285
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txt主页ID 
         Height          =   300
         Left            =   3570
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txt标识号 
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txt病人ID 
         Height          =   300
         Left            =   390
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txt床号 
         Height          =   300
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txt病人科室 
         Height          =   300
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox txt实收 
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   5670
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chk加班 
         Caption         =   "加班(&A)"
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   4770
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txt年龄 
         Height          =   300
         Left            =   3570
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   4710
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton cmd细目选择 
         Caption         =   "…"
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "热键：Ctrl+Enter"
         Top             =   2468
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.ComboBox cbo性别 
         Height          =   300
         Left            =   2040
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.ComboBox cbo开单科室 
         Height          =   300
         Left            =   8490
         TabIndex        =   14
         Top             =   840
         Width           =   1830
      End
      Begin VB.TextBox txt收费项目 
         Height          =   300
         Index           =   0
         Left            =   1635
         TabIndex        =   39
         Top             =   2460
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.ComboBox cbo收费类别 
         Height          =   300
         Index           =   0
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2460
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtPatient 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   420
         TabIndex        =   18
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox txt病人病区 
         Height          =   300
         Left            =   4710
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1350
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox txt数次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   4230
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2460
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt应收 
         Height          =   300
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5670
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "记帐单"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   23
         Top             =   210
         Visible         =   0   'False
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'————————————————————————————————————————————————————————————————————
'入口参数：
Public Enum UseType
    Use住院 = 0
    Use科室分散 = 1
    Use医技科室 = 2
    Use门诊 = 3
End Enum
Public Enum InState
    sta执行 = 0
    sta查阅 = 1
    sta调整 = 2
    sta销帐 = 3
End Enum

'2.表单初始状态参数：
Public mlng记帐ID  As Long '记帐单ID
Public mbytUseType As UseType   '记帐单用途,0-普通记帐,1-按科室分散记帐,2-医技科室记帐,3-门诊记帐
Public mbytInState As InState   '0-执行,1-浏览,2-调整,3-销帐
Public mstrInNO As String       '所操作的单据号
Public mlngUnitID As Long '当前记帐病区,为0时表示所有病区
Public mlngDeptID As Long '当前记帐科室,为0时表示所有科室
Public mlng病人ID As Long  '科室分散记帐用
Public mstrPrivs As String
Public mblnViewCancel As Boolean '是否查看已退的单据(mbytInState=1时有效)

Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
'————————————————————————————————————————————————————————————————————
'数据对象
Private mrsMedAudit As ADODB.Recordset  '病人已审批的费用项目
Private mrsMedPayMode As ADODB.Recordset '所有可用的医疗付款方式
Private mrsClass As ADODB.Recordset '根据参数读取的当前可用的收费类别
Private mrsUnit As ADODB.Recordset '可选择的执行科室
Private mrsInfo As New ADODB.Recordset '病人信息
Private mrs开单科室 As ADODB.Recordset  '可选的开单科室
Private mrs开单人 As ADODB.Recordset    '可选医生和护士

'程序对象
Private mobjBill As ExpenseBill  '★★★费用单据对象★★★
Private mcolBillDetails As BillDetails '单据的收费细目集
Private mobjBillDetail As BillDetail   '单据的收费细目对象
Private mcolBillInComes As BillInComes '收费细目的收入项目集
Private mobjBillIncome As BillInCome   '收费细目的收入项目对象
Private mobjDetail As Detail           '单独的收费细目对象
Private mcolDetails As Details   '★★★单独的收费细目集合★★

'程序变量
Private mstrWarn As String '已经报过警并选择继续的类别
Private mrsWarn As ADODB.Recordset  '病区报警线

Private mlngRows As Long            '当前记帐单的收费行数
Private mintCurrentRow As Integer   '当前记帐单的行号
Private mblnCard As Boolean         '是否刷就诊卡

Private mblnNOMoved As Boolean '操作的单据是否在后备数据表中,不从外面传值,现判断
Private mcurModiMoney As Currency '修改单据时原单据的金额
Private mstrUnitIDs As String   '当前操作员的所有病区ID

Private mcurPreMoney As Currency       '记录修改时原单费用,以便正确获取实际剩余款
Private mblnOne As Boolean             '是否只有一个可用收费类别
Private mcur可用金额 As Currency       '当前病人可用的最大金额
Private mblnDo As Boolean   '在combobox的Click职件中判断是否执行,避免index=**时隐式执行
Private marrDr() As String '记录医生的"ID|科室ID|编号|姓名|简码"

Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
    实时监控 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private mstrFreeTable As String
Private mstrTitle As String '用于窗体个性化保存的窗体名

Private mobjPublicExpense As Object  '费用公共部件
Private mintPriceGradeStartType As Integer
Private mstrPriceGrade As String

Public Sub MainProc()
    Dim tmpBill As ExpenseBill
    Dim i As Long, lngPre As Long, strPre As String, strTmp As String

    If mbytUseType <> Use门诊 Then
        mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p记帐操作)
    End If
    gblnOK = False: mblnDo = True
    Load frmCharge
    Set mobjBill = New ExpenseBill
    
    '初始化单据的界面
    If InitFace = False Then
        Unload frmCharge
        Exit Sub
    End If
    
    If mbytUseType <> Use门诊 Then
        mstrUnitIDs = GetUserUnits
    Else
        mstrUnitIDs = ""
    End If
    
    '初始化单据数据
    '新增除外,是否已转入后备数据表中
    If mbytInState = sta查阅 Then
        mblnNOMoved = zlDatabase.NOMoved(mstrFreeTable, mstrInNO, , 2, Me.Caption)
    Else
        If Not (mbytInState = sta执行 And mstrInNO = "") Then  '修改,调整,销帐
            If zlDatabase.NOMoved(mstrFreeTable, mstrInNO, , 2, Me.Caption) Then
                If Not ReturnMovedExes(mstrInNO, 2, Me.Caption) Then Exit Sub
            End If
            mblnNOMoved = False
        End If
    End If
    
    '这里执行表示新增或修改
    If mbytInState = sta执行 Or mbytInState = sta调整 Then
        '对本程序要用到的一些辅助数据进行装入操作，比如费别、开单科室、执行科室
        If Not InitData Then
            Unload frmCharge
            Exit Sub
        End If
    End If
    If mbytInState <> sta执行 Then   '显示、调整、销帐单据(1,2,3)
        '这些处理都很简单，用不着再去构造类
        Call NewBill
        If Not ReadBill(mstrInNO) Then
            Unload frmCharge
            Exit Sub
        End If
        cboNO.Text = mstrInNO
    Else '新增
        '读取该单据的内容
        If mstrInNO <> "" Then '修改单据  如果是在后备表中，不会执行到这里，在前面已退出
            Call ImportBill(mstrInNO, mlngRows, mstrPriceGrade)
            If mobjBill.NO = "" Then
                MsgBox "不能正确读取单据内容！", vbInformation, gstrSysName
                Unload frmCharge
                Exit Sub
            Else
                mcurModiMoney = GetBillMoney(IIf(mbytUseType = Use门诊, 1, 2), mobjBill.NO) '要在读取病人信息前先读
                
                lngPre = mobjBill.开单部门ID
                strPre = mobjBill.开单人

                txtPatient.Text = "-" & mobjBill.病人ID
                Call txtPatient_KeyPress(13)
                
                If mbytUseType <> Use门诊 Then
                    Call ReCalcInsure '重新计算统筹金额
                End If
                
                '显示的是原单据号,保存的是新单据号
                cboNO.Text = mobjBill.NO
                txtDate.Text = Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss")
                chk加班.Value = mobjBill.加班标志

                mblnDo = False
                    cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, lngPre)
                    If cbo开单科室.ListIndex = -1 And lngPre <> 0 Then
                        strTmp = GET部门名称(lngPre)
                        If strTmp <> "" Then
                            cbo开单科室.AddItem strTmp
                            cbo开单科室.ListIndex = cbo开单科室.NewIndex
                            cbo开单科室.ItemData(cbo开单科室.NewIndex) = lngPre
                        End If
                    End If
                    
                    i = 0
                    If cbo开单科室.ListIndex <> -1 Then i = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                    Call FillDoctor(i)
                    Call cbo.SeekIndex(cbo开单人, strPre, , True)
                    If cbo开单人.ListIndex = -1 And strPre <> "" Then
                        cbo开单人.AddItem strPre
                        cbo开单人.ListIndex = cbo开单人.NewIndex
                    End If
                mblnDo = True
                mobjBill.开单部门ID = lngPre
                mobjBill.开单人 = strPre

                '修改时应保存当前操作员的名字
                mobjBill.操作员编号 = UserInfo.编号
                mobjBill.操作员姓名 = UserInfo.姓名
                
                Call zlControl.CboLocate(cboBaby, mobjBill.婴儿费, True)

                Call ShowDetails
                Call ShowMoney
                
                'byZT200302
                For i = 0 To mlngRows - 1
                    If mobjBill.Details("R" & i).Detail.变价 Then
                        txt数次(i).TabStop = False
                        txt数次(i).Locked = True
                        txt标准单价(i).TabStop = True
                        txt标准单价(i).Locked = False
                    Else
                        txt数次(i).TabStop = True
                        txt数次(i).Locked = False
                        txt标准单价(i).TabStop = False
                        txt标准单价(i).Locked = True
                    End If
                    chk附加(i).Enabled = mobjBill.Details("R" & i).收费类别 = "F" '手术
                    If chk附加(i).Enabled = False Then chk附加(i).Value = 0
                    
                    '执行科室!!!
                    If mobjBill.Details("R" & i).收费细目ID <> 0 Then Call Fill执行科室(i)
                    
                    If cbo执行科室(i).ListCount = 1 Then
                        cbo执行科室(i).TabStop = False
                    Else
                        cbo执行科室(i).TabStop = True
                    End If
                Next

                mcurPreMoney = CalcGridToTal
            End If
        Else
            Call NewBill
            If mbytUseType = Use科室分散 And mlng病人ID <> 0 Then
                txtPatient.Text = "-" & mlng病人ID
                Call txtPatient_KeyPress(13)
            End If
        End If
    End If
    
    '初始化成功
    If Not gfrmMain Is Nothing Then
        frmCharge.Show vbModal, gfrmMain
    ElseIf glngMain <> 0 Then
        zlCommFun.ShowChildWindow frmCharge.hwnd, glngMain
    End If
End Sub

Private Sub cbo开单科室_Validate(Cancel As Boolean)
    '强制要选中一个(第一个)
    If cbo开单科室.ListIndex = -1 And cbo开单科室.ListCount <> 0 Then cbo开单科室.ListIndex = 0
End Sub

Private Sub cbo开单人_Validate(Cancel As Boolean)
    If cbo开单人.Text <> "" Then
        If cbo.FindIndex(cbo开单人, zlStr.NeedName(cbo开单人.Text), True) = -1 Then cbo开单人.ListIndex = -1: cbo开单人.Text = ""
    End If
    If cbo开单人.Text = "" Then Call cbo开单人_KeyPress(vbKeyReturn)
    '当开单科室确定开单人时,可能此时不选开单人,先去调整开单科室后再来选
    If gbln开单人 And cbo开单人.ListIndex = -1 And txtPatient.Text <> "" And cbo开单人.ListCount > 0 Then Cancel = True
End Sub

Private Sub chk附加_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    
    If mbytInState = sta查阅 Then
        cmdCancel.SetFocus
    ElseIf mbytInState = sta调整 Then
        txtDate.SetFocus
    ElseIf mbytInState = sta销帐 Then
        cmdOK.SetFocus
    Else
        If mbytUseType = Use科室分散 And mobjBill.姓名 <> "" Then
            If cbo收费类别(0).ListIndex = -1 And cbo收费类别(0).Visible = True Then
                cbo收费类别(0).SetFocus
            Else
                If txt收费项目(0).TabStop = True Then
                    txt收费项目(0).SetFocus
                Else
                    SendKeys "{TAB}"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    mstrTitle = "记帐处理"
    Call CreatePublicExpenseObject
    Call RestoreWinState(Me, App.ProductName, mstrTitle)
End Sub

Public Sub CreatePublicExpenseObject()
    '功能:创建公共费用部件
    Err = 0: On Error Resume Next
    If mobjPublicExpense Is Nothing Then
        Set mobjPublicExpense = CreateObject("zlPublicExpense.clsPublicExpense")
        If Err <> 0 Then
            MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)创建失败，请与系统管理员联系！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    If mobjPublicExpense Is Nothing Then Exit Sub
    
    'zlInitCommon(ByVal lngSys As Long, _
     ByVal cnOracle As ADODB.Connection, Optional ByVal strDbUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关的系统号及相关连接
    '入参:lngSys-系统号
    '     cnOracle-数据库连接对象
    '     strDBUser-数据库所有者
    '返回:初始化成功,返回true,否则返回False
    If mobjPublicExpense.zlInitCommon(glngSys, gcnOracle, gstrDbUser) = False Then
         MsgBox "注意:" & vbCrLf & "   费用公共部件(zl9PublicExpense)初始化失败，请与系统管理员联系！", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    mintPriceGradeStartType = mobjPublicExpense.zlGetPriceGradeStartType()
    If mintPriceGradeStartType = 0 Then Exit Sub
    '读取站点价格等级
    Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrInNO = 0
    mlngUnitID = 0
    mlngDeptID = 0
    mlng病人ID = 0
    mintCurrentRow = 0
    mblnViewCancel = False
    Set mrs开单科室 = Nothing
    Set mrs开单人 = Nothing
    Set mrsInfo = Nothing
    Set mrsMedAudit = Nothing
    Set mrsMedPayMode = Nothing
    Set mrsWarn = Nothing
    Call SaveWinState(Me, App.ProductName, mstrTitle)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If ActiveControl Is cbo开单人 Then Call cbo开单人_KeyPress(vbKeyReturn)
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF6 '清除当前单据内容,进入新单状态
            If mbytInState = 0 Then
                If fraForm.Enabled Then '正常输入单据状态'(清除后当作是新病人单据)
                    mstrInNO = ""
                    txtPatient.Text = "": txt年龄.Text = "": txt床号.Text = "": mcur可用金额 = 0
                    Call NewBill
                    txtPatient.SetFocus
                ElseIf chk销.Value = Checked Then '退据单状态
                    chk销.Value = Unchecked
                    Call NewBill
                    Call SetDisible(True)
                    txtPatient.SetFocus
                ElseIf Not fraForm.Enabled Then '收取划价单费用状态
                    Call NewBill
                    Call SetDisible(True)
                    txtPatient.SetFocus
                End If
            End If
        Case vbKeyF7 '切换输入法
            If Not gbln简码切换 Then Exit Sub   '35242
            If sta.Panels("WB").Visible And sta.Panels("PY").Visible Then
                If sta.Panels("WB").Bevel = sbrRaised Then
                    Call sta_PanelClick(sta.Panels("WB"))
                Else
                    Call sta_PanelClick(sta.Panels("PY"))
                End If
            End If
        Case vbKeyF8 '退(自动激活事件)
            If chk销.Visible And fra销.Enabled And chk销.Enabled Then chk销.Value = IIf(chk销.Value = Checked, Unchecked, Checked)
        Case vbKeyReturn
            If Shift And vbCtrlMask = vbCtrlMask Then
                If ActiveControl.Name = "txt收费项目" Then
                    Call cmd细目选择_Click(ActiveControl.Index)
                End If
            End If
        Case vbKeyEscape, vbKeyX
            If KeyCode = vbKeyX And Shift <> 4 Then Exit Sub
            Call cmdCancel_Click
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub sta_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If gbln简码切换 = False Then Exit Sub
    If Panel.Bevel = sbrRaised And (Panel.Key = "PY" Or Panel.Key = "WB") Then
        '切换并保存简码匹配方式
        Panel.Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        If Panel.Key = "PY" Then
            sta.Panels("WB").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        Else
            sta.Panels("PY").Bevel = IIf(Panel.Bevel = sbrInset, sbrRaised, sbrInset)
        End If
        zlDatabase.SetPara "简码方式", IIf(sta.Panels("PY").Bevel = sbrInset And sta.Panels("WB").Bevel = sbrInset, 2, IIf(sta.Panels("WB").Bevel = sbrInset, 1, 0))
        gbytCode = Val(zlDatabase.GetPara("简码方式", , , 0))
    End If
End Sub

Private Sub txt标识号_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt病人ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt病人病区_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt病人科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt床号_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt计算单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt年龄_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt年龄.Text = mobjBill.年龄
End Sub

Private Sub txt床号_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt床号.Text = mobjBill.床号
End Sub

Private Sub txt病人ID_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt病人ID.Text = Format(mobjBill.病人ID, "#;;;")
End Sub

Private Sub txt标识号_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt标识号.Text = Format(mobjBill.标识号, "#;;;")
End Sub

Private Sub txt实收金额_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt应收金额_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt主页ID_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt主页ID.Text = Format(mobjBill.主页ID, "#;;;")
End Sub

Private Sub txt病人病区_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt病人病区.Text = mobjBill.病区
End Sub

Private Sub txt病人科室_Change()
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    txt病人科室.Text = mobjBill.科室
End Sub

Private Sub txt收费项目_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    If txt收费项目(Index).Locked = True Then
        txt收费项目(Index).Text = mobjBill.Details("R" & Index).收费名称
    End If
End Sub

Private Sub txt计算单位_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    If txt计算单位(Index).Locked = True Then
        txt计算单位(Index).Text = mobjBill.Details("R" & Index).计算单位
    End If
End Sub

Private Sub txt数次_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If Not fraForm.Enabled Then Exit Sub
    
    If txt数次(Index).Locked Then
        If mobjBill.Details("R" & Index).数次 = 0 Then
            txt数次(Index).Text = ""
        Else
            txt数次(Index).Text = mobjBill.Details("R" & Index).数次
        End If
    End If
End Sub

Private Sub txt标准单价_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    If txt标准单价(Index).Locked Then
        If mobjBill.Details("R" & Index).标准单价 <> 0 Then
            txt标准单价(Index).Text = Format(mobjBill.Details("R" & Index).标准单价, "0.0000")
        Else
            txt标准单价(Index).Text = ""
        End If
    End If
End Sub

Private Sub txt应收金额_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    If mobjBill.Details("R" & Index).应收金额 <> 0 Then
        txt应收金额(Index).Text = Format(mobjBill.Details("R" & Index).应收金额, gstrDec)
    Else
        txt应收金额(Index).Text = ""
    End If
End Sub

Private Sub txt实收金额_Change(Index As Integer)
    '如果fraForm不可用，那肯定是程序在改变
    If fraForm.Enabled = False Then Exit Sub
    
    If mobjBill.Details("R" & Index).实收金额 <> 0 Then
        txt实收金额(Index).Text = Format(mobjBill.Details("R" & Index).实收金额, gstrDec)
    Else
        txt实收金额(Index).Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    mblnCard = False
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    
    If txtPatient.Text = mobjBill.姓名 Then Exit Sub
    If Trim(txtPatient.Text) = "" Then
        '空串特殊处理
        txtPatient.Text = mobjBill.姓名
        Exit Sub
    End If
    
    If Input姓名() = False Then
        Cancel = True
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Integer
    
    mblnCard = False
    
    On Error Resume Next

    If txtPatient.Locked Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    mblnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, glngSys)

    If Trim(Me.txtPatient.Text) = "" And KeyAscii = 13 Then
        With frmPatiSelect
            If mbytUseType = Use住院 Or mbytUseType = Use科室分散 Then
                .mlngUnitID = mlngUnitID
            ElseIf mbytUseType = Use医技科室 Then
                .mlngUnitID = mlngDeptID
            Else
                KeyAscii = 0
                Exit Sub
            End If
            .mbytUseType = mbytUseType
            .mstrPrivs = mstrPrivs
            Set .mfrmParent = Me
            .Show 1, Me
            Me.Refresh
        End With
    End If

    If mblnCard And Len(txtPatient.Text) = gbytCardNOLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then '划卡或敲回车
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        Else
            If txtPatient.Text = mobjBill.姓名 Then
                If cbo开单科室.ListIndex = -1 Then
                    cbo开单科室.SetFocus
                Else
                    If cbo收费类别(0).ListIndex = -1 And cbo收费类别(0).Visible = True Then
                        cbo收费类别(0).SetFocus
                    Else
                        If txt收费项目(0).TabStop = True Then
                            txt收费项目(0).SetFocus
                        Else
                            SendKeys "{TAB}"
                        End If
                    End If
                End If
                Exit Sub
            End If
        End If
        KeyAscii = 0
        
        If Input姓名() = True Then
            '输入得到保存
            If cbo开单科室.ListIndex = -1 Then
                cbo开单科室.SetFocus
            Else
                If cbo收费类别(0).ListIndex = -1 And cbo收费类别(0).Visible = True Then
                    cbo收费类别(0).SetFocus
                Else
                    If txt收费项目(0).TabStop = True Then
                        txt收费项目(0).SetFocus
                    Else
                        SendKeys "{TAB}"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Function Input姓名() As Boolean
'功能：输入病人姓名
'参数：表示来源于键盘或鼠标
    '读取病人信息
    Dim blnReturn As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim strSQL As String
    Dim blnOutMsg As Boolean
    If Not (mbytInState = 0 And mbytUseType = 1 And sta.Panels(2) Like "上一张*") Then
        sta.Panels(2) = ""
    End If
    If mbytUseType = Use门诊 Then
       blnReturn = GetPatientOut(txtPatient.Text)
    Else
       blnReturn = GetPatientIn(txtPatient.Text, mblnCard, blnOutMsg)
    End If
    If Not blnReturn Then
        If mblnCard Then
            txtPatient.Text = ""
           If Not blnOutMsg Then MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation, gstrSysName
            Call ClearPatient
        Else
            If Not blnOutMsg Then MsgBox "不能读取病人信息！", vbInformation, gstrSysName
            If mstrInNO = "" Then
                strTemp = txtPatient.Text
                Call ClearPatient
                txtPatient.Text = strTemp
            End If
            zlControl.TxtSelAll txtPatient
        End If
        Exit Function
    Else
        '就诊卡密码检查
        If Mid(gstrCardPass, 6, 1) = "1" And mblnCard Then
            If Not zlCommFun.VerifyPassWord(Me, "" & mrsInfo!卡验证码, mrsInfo!姓名, mrsInfo!性别, "" & mrsInfo!年龄) Then
                txtPatient.Text = ""
                Call ClearPatient
                Exit Function
            End If
        End If
    
        '判断该病人的费别是否合适
        Call cbo.SeekIndex(cbo费别, IIf(IsNull(mrsInfo("费别")), "", mrsInfo("费别")), , True)
        If cbo费别.ListIndex = -1 Then
            txtPatient.Text = ""
            MsgBox "病人" & IIf(IsNull(mrsInfo("姓名")), "", mrsInfo("姓名")) & "的费别信息不再有效，不能记帐！", vbInformation, gstrSysName
            Call ClearPatient
            Exit Function
        End If
        
        '不再是上级程序传进来的病人
        If mbytUseType = Use科室分散 And mrsInfo!病人ID <> mlng病人ID Then mlng病人ID = 0

        If mbytUseType = Use门诊 Then
            '由挂号单得来时有执行部门并作为这里的开单部门
            If Not IsNull(mrsInfo("科室ID")) Then
                If IsNull(mrsInfo!姓名) Then
                    txtPatient.Text = ""
                    MsgBox "该病人挂号时没有登记档案,需要输入病人姓名！", vbInformation, gstrSysName
                    Call ClearPatient
                    Set mrsInfo = New ADODB.Recordset
                    Exit Function
                End If
                
                mobjBill.科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
                Set开单科室 IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
            End If
        Else
             '自动设置开单科室(同时设置记帐报警信息)
            mobjBill.科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
            Set开单科室 IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
        End If
        
        '病人预交款信息
        Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, CDbl(mcurModiMoney), Val("" & mrsInfo!险类) > 0)
        If rsTmp.State = adStateOpen Then
            sta.Panels(3).Text = "预交:" & Format(rsTmp!预交余额, "0.00")
            sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(rsTmp!费用余额, gstrDec)
            sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
            cmdOK.Tag = rsTmp!预交余额
            cmdCancel.Tag = rsTmp!费用余额
            mcur可用金额 = rsTmp!预交余额 - rsTmp!费用余额
        Else
            sta.Panels(3).Text = "预交:0.00/费用:" & gstrDec & "/剩余:0.00"
            cmdOK.Tag = 0
            cmdCancel.Tag = 0
            mcur可用金额 = 0
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        '刘兴洪:26952
        Dim cur余额 As Currency, curItemMoney As Currency, cur当日额 As Currency, curTotal As Currency
        cur余额 = mcur可用金额
        curItemMoney = 0
        '单据费用
        curTotal = CalcGridToTal
        
        '重新读取当日额
        cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
        If gbln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(2, Val(NVL(mrsInfo!病人ID)))
        
        
        gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名, Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, cur余额, cur当日额 - mcurModiMoney, curTotal, Val(NVL(mrsInfo!担保额)), "", "", mstrWarn, , , True)
        '返回:0;没有报警,继续
        '     1:报警提示后用户选择继续
        '     2:报警提示后用户选择中断
        '     3:报警提示必须中断
        '     4:强制记帐报警,继续
        '     5.报警提示后用户选择继续,但只允许保存存为划价单
        If gbytWarn = 2 Or gbytWarn = 3 Then
            Set mrsInfo = New ADODB.Recordset: txtPatient.Text = "":
            mlng病人ID = 0
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            Call ClearPatient: Exit Function
        End If
        '--------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        
        
        Call LoadPatientBaby(cboBaby, mrsInfo!病人ID, mrsInfo!主页ID)
                                
        '病人信息
        With mobjBill
            .姓名 = IIf(IsNull(mrsInfo!姓名), 0, mrsInfo!姓名)
            .病人ID = IIf(IsNull(mrsInfo!病人ID), 0, mrsInfo!病人ID)
            .主页ID = IIf(IsNull(mrsInfo!主页ID), 0, mrsInfo!主页ID)
            .标识号 = IIf(IsNull(mrsInfo!标识号), 0, mrsInfo!标识号)
            .床号 = "" & mrsInfo!床号
            .性别 = IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
            .年龄 = IIf(IsNull(mrsInfo!年龄), 0, mrsInfo!年龄)
            .费别 = IIf(IsNull(mrsInfo!费别), "", mrsInfo!费别)
            .担保额 = IIf(IsNull(mrsInfo!担保额), 0, mrsInfo!担保额)

            .病区ID = IIf(IsNull(mrsInfo!病区ID), 0, mrsInfo!病区ID)
            .科室ID = IIf(IsNull(mrsInfo!科室ID), 0, mrsInfo!科室ID)
            .病区 = IIf(IsNull(mrsInfo!病区), "", mrsInfo!病区)
            .科室 = IIf(IsNull(mrsInfo!科室), 0, mrsInfo!科室)

            If cbo开单科室.ListIndex <> -1 Then
                mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
            Else
                mobjBill.开单部门ID = 0
            End If
        End With
        
        If Not IsNull(mrsInfo!险类) Then
            MCPAR.负数记帐 = gclsInsure.GetCapability(support负数记帐, mrsInfo!病人ID, mrsInfo!险类)
            MCPAR.记帐上传 = gclsInsure.GetCapability(support记帐上传, mrsInfo!病人ID, mrsInfo!险类)
            MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, mrsInfo!病人ID, mrsInfo!险类)
            MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, mrsInfo!病人ID, mrsInfo!险类)
            MCPAR.实时监控 = gclsInsure.GetCapability(support实时监控, mrsInfo!病人ID, mrsInfo!险类)
        End If
        
        Call ShowPatient
        txtPatient.PasswordChar = ""

        If Not IsNull(mrsInfo!出院日期) And mbytUseType <> Use门诊 Then
            MsgBox "提醒您：" & vbCrLf & vbCrLf & "该病人已于 " & Format(mrsInfo!出院日期, "yyyy-MM-dd") & " 出院，现在对该病人强制进行记帐！", vbInformation, gstrSysName
            txtDate.Text = Format(mrsInfo!出院日期, "yyyy-MM-dd HH:mm:ss")
        Else
            txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        
        '读取价格等级
        If mintPriceGradeStartType >= 2 Then
            Call mobjPublicExpense.zlGetPriceGrade(gstrNodeNo, Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID)), _
                NVL(mrsInfo!医疗付款方式), , , mstrPriceGrade)
        End If
        
        If mbytInState = 0 And mobjBill.Details.Count > 0 Then
            '重新计算价格
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
    Input姓名 = True
End Function

Private Sub ClearPatient()
'功能：清除病人信息的显示
    With mobjBill
        .病人ID = 0
        .主页ID = 0

        .病区ID = 0
        .科室ID = 0
        .病区 = ""
        .科室 = ""

        .床号 = ""
        .标识号 = 0
        .姓名 = ""
        .性别 = ""
        .年龄 = ""
        .费别 = ""
        .担保额 = 0
    End With
    Call ShowPatient
End Sub

Private Sub ShowPatient()
'功能：显示病人信息
    With mobjBill
        txtPatient.Text = .姓名
        Call cbo.SeekIndex(cbo性别, .性别, , True)
        txt年龄.Text = .年龄
        Call cbo.SeekIndex(cbo费别, .费别, , True)
        txt床号.Text = Format(.床号, "#;;;")
        
        txt病人ID.Text = Format(.病人ID, "#;;;")
        txt主页ID.Text = Format(.主页ID, "#;;;")
        txt标识号.Text = Format(.标识号, "#;;;")
        txt病人病区.Text = .病区
        txt病人科室.Text = .科室
    End With
End Sub

Private Sub Set开单科室(ByVal lngID As Long)
'功能：设置开单科室
'注意：如果开单科室的Tag属性有设置的话，还要进行相应的处理
    If cbo开单科室.Tag <> "" Then
        Select Case cbo开单科室.Tag
            Case "C1" '病人所有科室
                cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, IIf(mobjBill.科室ID = 0, lngID, mobjBill.科室ID))
            Case "C2" '操作员所在科室
                cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID))
            Case Else '指定科室
                cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, Val(cbo开单科室.Tag))
                If cbo开单科室.ListIndex < 0 Then
                    cbo开单科室.AddItem GET部门名称(cbo开单科室.Tag, mrs开单科室), 0
                    cbo开单科室.ListIndex = 0
                End If
        End Select
        
        If cbo开单科室.ListCount > 0 And cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    Else
        cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, lngID)
    End If
    
    If cbo开单科室.ListIndex = -1 Then
        mobjBill.开单部门ID = 0
    Else
        mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    End If
End Sub

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo性别.ListIndex <> -1 Then mobjBill.性别 = Mid(cbo性别.Text, InStr(cbo性别.Text, "-") + 1)
        SendKeys "{TAB}"
    End If
    If cbo性别.Locked Then Exit Sub
    If SendMessage(cbo性别.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
End Sub

Private Sub txt年龄_Gotfocus()
    zlControl.TxtSelAll txt年龄
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        mobjBill.年龄 = txt年龄.Text
        SendKeys "{TAB}"
    End If
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep
End Sub

Private Sub cbo费别_Click()
    If cbo费别.ListIndex <> -1 And Not mobjBill Is Nothing Then
        mobjBill.费别 = zlStr.NeedName(cbo费别.Text)

        If mbytInState = sta执行 Then
            If mobjBill.Details.Count = 0 Then Exit Sub
            '重新计算价格
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If
    End If
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If cbo费别.Locked Then
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
        Exit Sub
    End If
    If KeyAscii = vbKeyReturn And cbo费别.ListIndex <> -1 Then
        mobjBill.费别 = zlStr.NeedName(cbo费别.Text)

        If mbytInState = sta执行 And mstrInNO <> "" Then
            '重新计算价格
            Call CalcMoneys
            Call ShowDetails
            Call ShowMoney
        End If

        SendKeys "{TAB}"
    End If
'    If SendMessage(cbo费别.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
    lngIdx = zlControl.CboMatchIndex(cbo费别.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo费别.ListIndex = lngIdx
End Sub

Private Sub cbo开单科室_KeyPress(KeyAscii As Integer)
    

   Dim lngIdx As Long, lng医生ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo开单科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If cbo开单人.ListIndex >= 0 Then lng医生ID = cbo开单人.ItemData(cbo开单人.ListIndex)
    If mrs开单科室 Is Nothing Then Call FillDept(lng医生ID)
    
    If zlSelectDept(Me, 0, cbo开单科室, mrs开单科室, cbo开单科室.Text) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub

'
'
'
'
'    Dim lngIdx As Long
'    If KeyAscii = 13 And cbo开单科室.ListIndex <> -1 Then
'        mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
'        SendKeys "{TAB}"
'        Exit Sub
'    End If
'    If cbo开单科室.Locked Then Exit Sub
'
'    If SendMessage(cbo开单科室.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
'    lngIdx = MatchIndex(cbo开单科室.hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo开单科室.ListIndex = lngIdx

    '强制要选中一个(第一个)
    If cbo开单科室.ListIndex = -1 And cbo开单科室.ListCount <> 0 Then cbo开单科室.ListIndex = 0
End Sub

Private Sub cbo开单科室_Click()
    Dim i As Long, strDoctor As String
    If Not mblnDo Then Exit Sub
       
    '定位医生
    cbo开单人.Clear
    If cbo开单科室.ListIndex <> -1 Then
        FillDoctor cbo开单科室.ItemData(cbo开单科室.ListIndex)
    End If

    '数据对象
    If mbytInState = 0 Then
        If cbo开单科室.ListIndex = -1 Then
            mobjBill.开单部门ID = 0
        Else
            mobjBill.开单部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
    End If
    
    '重新设置相关项目的执行科室
    'byZT200302
    If mbytInState = 0 And cbo开单科室.ListIndex <> -1 And cbo开单科室.Visible Then
        For i = 0 To mobjBill.Details.Count - 1
            With mobjBill.Details("R" & i)
                If .Detail.执行科室 = 6 Then '6-开单人科室
                    cbo执行科室(i).Clear
                    'Call ShowDetail(i)
                    .执行部门ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
                End If
            End With
        Next
    End If
End Sub

Private Function isCheck开单人Exists(ByVal str姓名 As String, Optional blnLocateItem As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查姓名是否在开单人下拉列表中.
    '入参:str姓名-姓名
    '     blnLocateItem:是否直接定位
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-07-20 17:53:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To cbo开单人.ListCount - 1
        If zlStr.NeedName(cbo开单人.List(i)) = str姓名 Then
            If blnLocateItem Then cbo开单人.ListIndex = i
            isCheck开单人Exists = True
            Exit Function
        End If
    Next
End Function


Private Sub cbo开单人_KeyPress(KeyAscii As Integer)
    Dim i As Integer, intIdx As Integer, strResult As String, iCount As Integer
    Dim strText As String, strFilter As String, rsTemp As ADODB.Recordset
        
    If KeyAscii = vbKeyReturn Then
        strText = UCase(cbo开单人.Text)
        If cbo开单人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo开单人.List(cbo开单人.ListIndex) Then
                Call zlControl.CboSetIndex(cbo开单人.hwnd, -1)
            Else
                zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        
        If strText = "" Then
            cbo开单人.ListIndex = -1
        Else
            intIdx = -1
          strFilter = IIf(gbln护士, "人员性质<>''", "人员性质<>'护士'")
            '刘兴洪:22383
            '先复制记录集
            Set rsTemp = zlDatabase.zlCopyDataStructure(mrs开单人)
            Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
            Dim strCompents As String '匹配串
            
            strCompents = Replace(gstrLike, "%", "*") & strText & "*"
            
            If IsNumeric(strText) Then
                intInputType = 0
            ElseIf zlCommFun.IsCharAlpha(strText) Then
                intInputType = 1
            Else
                intInputType = 2
            End If
            
            mrs开单人.Filter = strFilter: iCount = 0
            With mrs开单人
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not mrs开单人.EOF
                    Select Case intInputType
                    Case 0  '输入的是全数字
                        '1.编号输入值相等,主要输入如:12 匹配000012这种况,但如果输入的是01与编号01相等,则直接定位到01,则不定位在1上.
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                        '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
                        If NVL(!编号) = strText Then strResult = NVL(!姓名): iCount = 0: Exit Do
                        
                        '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等.因此如果存在此种情况,需要弹出选择器供选择
                        If Val(NVL(!编号)) = Val(strText) Then
                            If iCount = 0 Then strResult = NVL(!姓名)
                            iCount = iCount + 1
                        End If
                        '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
                         If Val(mrs开单人!编号) Like strText & "*" Then
                            If isCheck开单人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                         End If
                    Case 1  '输入的是全字母
                        '规则:
                        ' 1.输入的简码相等,则直接定位
                        ' 2.根据参数来匹配相同数据
                        
                        '1.输入的简码相等,则直接定位
                        If Trim(NVL(!简码)) = strText Then
                            If iCount = 0 Then strResult = NVL(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.根据参数来匹配相同数据
                        If Trim(NVL(!简码)) Like strCompents Then
                            If isCheck开单人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                        End If
                    Case Else  ' 2-其他
                        '规则:可能存在汉字等情况,或编号类似于N001简码可能有ZYK01这种情况
                        '1.编码\简码相等,直接定位
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        
                        '1.编码\简码相等,直接定位
                        If Trim(!编号) = strText Or Trim(!简码) = strText Or Trim(!姓名) = strText Then
                            If iCount = 0 Then strResult = NVL(!姓名)   '可能存在多个相同的多个
                            iCount = iCount + 1
                        End If
                        
                        '2.简码或编码或姓名 根据参数来匹配数(但编码只能左匹配)
                        If Trim(!编号) Like strText & "*" Or Trim(NVL(!简码)) Like strCompents Or Trim(NVL(!姓名)) Like strCompents Then
                            If isCheck开单人Exists(NVL(!姓名)) Then Call zlDatabase.zlInsertCurrRowData(mrs开单人, rsTemp)
                        End If
                    End Select
                    mrs开单人.MoveNext
                Loop
            End With
             If iCount > 1 Then strResult = ""
            If strResult = "" And rsTemp.RecordCount = 1 Then strResult = NVL(rsTemp!姓名)
            '刘兴洪:直接定位
            If strResult <> "" Then
                rsTemp.Close: Set rsTemp = Nothing
                If isCheck开单人Exists(strResult, True) Then zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            
            '需要检查是否有多条满足条件的记录
            If rsTemp.RecordCount <> 0 Then
                '先按某种方式进行排序
                Select Case intInputType
                Case 0 '输入全数字
                    rsTemp.Sort = "编号"
                Case 1 '输入全拼音
                    rsTemp.Sort = "简码"
                Case Else
                    '根据选择来定
'                    If gbyt开单人显示 = 1 Then '简码
'                        rsTemp.Sort = "简码"
'                    Else
                        rsTemp.Sort = "编号"
                  '  End If
                End Select
                '弹出选择器
                Dim rsReturn As ADODB.Recordset
                If zlDatabase.zlShowListSelect(Me, glngSys, 1133, cbo开单人, rsTemp, True, "", "缺省,职务,优先级别", rsReturn) Then
                    If Not rsReturn Is Nothing Then
                        If rsReturn.RecordCount <> 0 Then
                            '进行定位
                            If isCheck开单人Exists(NVL(rsReturn!姓名), True) Then
                                rsTemp.Close: Set rsTemp = Nothing
                                zlCommFun.PressKey vbKeyTab
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Else
                '未找到
                rsTemp.Close: Set rsTemp = Nothing
                KeyAscii = 0: zlControl.TxtSelAll cbo开单人: Exit Sub
            End If
            rsTemp.Close: Set rsTemp = Nothing
                         
            
'
'            For i = 0 To cbo开单人.ListCount - 1
'                If InStr(cbo开单人.List(i), UCase(strText)) > 0 Then
'                    If intIdx = -1 Then cbo开单人.ListIndex = i
'                    intIdx = i
'                End If
'                If IsNumeric(strText) Then
'                    If cbo开单人.ItemData(i) = CDbl(strText) Then
'                        If intIdx = -1 Then cbo开单人.ListIndex = i
'                        intIdx = i
'                    End If
'                End If
'            Next
        End If
        If cbo开单人.ListIndex = -1 Then
            cbo开单人.Text = ""
            mobjBill.开单人 = UserInfo.姓名
        Else
            mobjBill.开单人 = zlStr.NeedName(cbo开单人.Text)
            If intIdx <> cbo开单人.ListIndex Then SendKeys "{F4}": Exit Sub
            SendKeys "{TAB}"
        End If
    End If
End Sub

Private Sub cbo开单人_Click()
    If Not mblnDo Then Exit Sub
    
    If mbytInState = 0 Then
        '数据对象
        mobjBill.开单人 = IIf(cbo开单人.ListIndex = -1, "", zlStr.NeedName(cbo开单人.Text))
    End If
End Sub

Private Sub cboBaby_Click()
    mobjBill.婴儿费 = cboBaby.ItemData(cboBaby.ListIndex)
End Sub

Private Sub cboBaby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub chk销_Click()
    Dim i As Long
    
    mstrInNO = ""
    '改变可用性
    If chk销.Value = 1 Then
        chk销.ForeColor = &HFF&
        cboNO.Locked = False
        
        fraForm.Enabled = False
        fra开单人.Enabled = False
        fra时间.Enabled = False
    Else
        chk销.ForeColor = 0
        
        fraForm.Enabled = True
        fra开单人.Enabled = True
        fra时间.Enabled = True
        
        cboNO.Locked = True
    End If
        
    'btZY200302
    For i = 0 To mlngRows - 1
        cbo执行科室(i).Clear
    Next
    
    '初始化
    Call NewBill
    
    '扫尾处理
    If chk销.Value = 1 Then
        cboNO.SetFocus
    Else
        '读出数据
        Call cbo开单科室_Click
        If mbytUseType = 1 And mlng病人ID > 0 Then
            txtPatient.Text = "-" & mlng病人ID
            Call txtPatient_KeyPress(13)
        Else
            txtPatient.SetFocus
        End If
    End If
End Sub

Private Sub chk加班_Click()
    If mbytInState = sta查阅 Or chk销.Value = Checked Then Exit Sub
    If mbytInState = sta调整 Then Exit Sub
    If Not chk加班.Visible Then Exit Sub

    Dim blnAdd As Boolean

    blnAdd = OverTime(zlDatabase.Currentdate)
    If chk加班.Value = Unchecked And blnAdd Then
        If MsgBox("当前处于加班时间范围内,要取消加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Checked
        End If
    End If
    If chk加班.Value = Checked And Not blnAdd Then
        If MsgBox("当前不处于加班时间范围内,要执行加班加价吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            chk加班.Value = Unchecked
        End If
    End If
    mobjBill.加班标志 = IIf(chk加班.Value = Checked, 1, 0)
    
    '重新计算价格
    If Not mobjBill.Details.Count = 0 Then
        Call CalcMoneys
        Call ShowDetails
        Call ShowMoney
    End If
End Sub

Private Sub chk加班_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 0
    txtDate.SelLength = Len(txtDate.Text)
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsDate(txtDate.Text) Then
        mobjBill.登记时间 = CDate(txtDate.Text)
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtDate_LostFocus()
    txtDate.SelLength = 0
    If IsDate(txtDate.Text) Then mobjBill.登记时间 = CDate(txtDate.Text)
End Sub

Private Sub cboNO_GotFocus()
    cboNO.SelStart = 0
    cboNO.SelLength = Len(cboNO.Text)
    If chk销.Value = Checked Then
        cboNO.Locked = False
    Else
        cboNO.Locked = True
    End If
End Sub

Private Sub cboNO_KeyPress(KeyAscii As Integer)
    Dim blnRead As Boolean, strOper As String, vDate As Date
    
    If KeyAscii > 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress cboNO, KeyAscii, m文本式
 
    If KeyAscii = 13 And cboNO.Locked Then
        SendKeys "{TAB}"
    End If
    If KeyAscii = 13 And cboNO.Text <> "" And Not cboNO.Locked Then
        cboNO.Text = GetFullNO(cboNO.Text, 14)

        If chk销.Value = 1 Then
            '是否已转入后备数据表中
            If zlDatabase.NOMoved(mstrFreeTable, cboNO.Text, , 2, Me.Caption) Then
                If Not ReturnMovedExes(cboNO.Text, 2, Me.Caption) Then Exit Sub
                mblnNOMoved = False
            End If
            
             '单据权限
            If Not ReadBillInfo(IIf(mbytUseType = Use门诊, 1, 2), cboNO.Text, 2, strOper, vDate) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
            If mbytUseType = 0 And InStr(mstrPrivs, "所有操作员") <= 0 Then
                If UserInfo.姓名 <> strOper Then
                    MsgBox "你没有""所有操作员""权限,不能对" & strOper & "的单据进行销帐!", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                End If
            End If
            If Not BillOperCheck(5, strOper, vDate, "销帐", cboNO.Text) Then
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If
        
            If CheckExecute(cboNO.Text, mlng记帐ID, IIf(mbytUseType = Use门诊, 1, 2)) Then
                MsgBox "该记帐单内容已经全部执行" & vbCrLf & "或不是由本记帐单登记的，不能销帐！", vbInformation, gstrSysName
                cboNO.Text = "": cboNO.SetFocus: Exit Sub
            End If

            '是否已结帐
            'int来源-1-门诊;2-住院
            If HaveBilling(IIf(mbytUseType = Use门诊, 1, 2), cboNO.Text, False) <> 0 Then  'mlng记帐ID
                If BillExistInsure(cboNO.Text) <> 0 Then
                    MsgBox "该医保记帐单据包含已经结帐的内容,不能销帐！", vbInformation, gstrSysName
                    cboNO.Text = "": cboNO.SetFocus: Exit Sub
                Else
                    Select Case gbytBillOpt
                        Case 0
                        Case 1
                            If MsgBox("该记帐单据包含已经结帐的内容,要销帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                cboNO.Text = "": cboNO.SetFocus: Exit Sub
                            End If
                        Case 2
                            MsgBox "该记帐单据包含已经结帐的内容,不能销帐！", vbInformation, gstrSysName
                            cboNO.Text = "": cboNO.SetFocus: Exit Sub
                    End Select
                End If
            End If
            
            '是否存在重算冲减记录
            If CheckRecalcRecord(cboNO.Text) Then
                MsgBox "发现该记帐单据存在按费别重算的打折冲减记录!" & vbCrLf & _
                    "结帐前请按费别重算费用，否则病人将享受单据销帐前的打折优惠金额！", vbInformation, Me.Caption
            End If
            
            blnRead = ReadBill(cboNO.Text)
        End If

        If blnRead Then
            mstrInNO = cboNO.Text '确定时以mstrInNO为准
            cmdOK.SetFocus
        Else
            mstrInNO = "": cboNO.Text = "": cboNO.SetFocus
        End If
    End If
End Sub

Private Function CheckRecalcRecord(ByVal strNO As String) As Boolean
'功能：判断指定病人的指定单据是否存在按费别重算的冲减记录(数次为0的记录)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim bln住院 As Boolean
    bln住院 = (mstrFreeTable = "住院费用记录")
    
    Err = 0: On Error GoTo errH:
    strSQL = "Select Count(A.ID) Num" & vbNewLine & _
            "From " & mstrFreeTable & " A," & vbNewLine & _
            "     ( Select 病人id," & IIf(bln住院, " 主页id, 病人病区id,", "0 as 主页id,0 as 病人病区ID,") & " 病人科室id, 收费细目id, 收入项目id, 开单部门id, 执行部门id, 发生时间" & vbNewLine & _
            "       From " & mstrFreeTable & vbNewLine & _
            "       Where NO = [1] And 记帐费用 = 1" & vbNewLine & _
            "       Group By 病人id," & IIf(bln住院, " 主页id, 病人病区id,", "") & "病人科室id, 收费细目id, 收入项目id, 开单部门id, 执行部门id, 发生时间) B" & vbNewLine & _
            "Where A.记录性质 = 2 And A.数次 = 0 And A.病人id+0 = B.病人id " & _
                   IIf(bln住院, " And A.主页id = B.主页id And A.病人病区id + 0 = B.病人病区id ", "") & _
            "       And A.病人科室id + 0 = B.病人科室id And A.收费细目id + 0 = B.收费细目id And" & vbNewLine & _
            "      A.收入项目id + 0 = B.收入项目id And A.开单部门id + 0 = B.开单部门id And A.执行部门id + 0 = B.执行部门id And" & vbNewLine & _
            "      A.发生时间 = B.发生时间"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, strNO)
    If rsTmp.RecordCount > 0 Then CheckRecalcRecord = rsTmp!Num > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    If mbytInState = sta执行 Then
        If Not CheckBillisZero Then
            If MsgBox("确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim curTotal As Currency, cur当日额 As Currency
    Dim intInsure As Integer, i As Long
    Dim strInfo As String

    If mbytInState = sta销帐 Then '%%%
        '医保记帐作废上传(注意判断顺序)
        If mbytUseType <> Use门诊 Then
            intInsure = BillExistInsure(mstrInNO) '判断是否医保病人记的帐
            '去掉了医保连接匹配检查
        End If
        
        If mbytUseType = Use门诊 Then
            strSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '医保记帐作废上传
        If mbytUseType <> Use门诊 And intInsure <> 0 Then
            If MCPAR.记帐作废上传 And Not MCPAR.记帐完成后上传 Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        '医保记帐作废上传
        If mbytUseType <> Use门诊 And intInsure <> 0 Then
            If MCPAR.记帐作废上传 And MCPAR.记帐完成后上传 Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        gblnOK = True
        Unload Me: Exit Sub
    ElseIf mbytInState = sta调整 Then
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入合法的费用时间！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check发生时间(CDate(txtDate.Text), cboNO.Text)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
            
        If Not SaveModi() Then Exit Sub
        gblnOK = True: Unload Me: Exit Sub
    ElseIf chk销.Value = 0 Then '正常输入单据状态'%%%
        If mrsInfo.State = adStateClosed Then
            MsgBox "没有发现病人信息,请确定病人信息！", vbInformation, gstrSysName
            txtPatient.SetFocus: Exit Sub
        End If
        If cbo费别.ListIndex = -1 Or mobjBill.费别 = "" Then
            MsgBox "请选择病人费别！", vbInformation, gstrSysName
            If cbo费别.Visible = True Then cbo费别.SetFocus: Exit Sub
        End If
        If mobjBill.开单部门ID = 0 Then
            MsgBox "请确定开单科室！", vbInformation, gstrSysName
            cbo开单科室.SetFocus
            Exit Sub
        End If
        
        If mobjBill.开单人 = "" And gbln开单人 Then
            MsgBox "请输入开单人！", vbInformation, gstrSysName
            cbo开单人.SetFocus: Exit Sub
        End If
        
        strSQL = ""
        For i = 0 To mlngRows - 1
            If mobjBill.Details("R" & i).收费细目ID <> 0 Then
                strSQL = "有输入"
                
                If mobjBill.Details("R" & i).执行部门ID = 0 Then
                    MsgBox "该条收费项目的执行部门没设置！", vbInformation, gstrSysName
                    If cbo执行科室(i).Visible = True Then
                        cbo执行科室(i).SetFocus
                    Else
                        txt收费项目(i).SetFocus
                    End If
                    Exit Sub
                End If
            End If
        Next
        If strSQL = "" Then
            MsgBox "单据中没有任何内容,请正确输入单据内容！", vbInformation, gstrSysName
            txt收费项目(0).SetFocus: Exit Sub
        End If
        If Not IsDate(txtDate.Text) Then
            MsgBox "请输入正确的费用日期！", vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        strInfo = Check发生时间(CDate(txtDate.Text), mrsInfo!病人ID)
        If strInfo <> "" Then
            MsgBox strInfo, vbInformation, gstrSysName
            txtDate.SetFocus: Exit Sub
        End If
        
        If Not IsNull(mrsInfo!出院日期) Then
            If Format(txtDate.Text, txtDate.Format) > Format(mrsInfo!出院日期, txtDate.Format) Then
                MsgBox "强制对出院病人记帐时，费用时间不能大于病人出院时间:" & Format(mrsInfo!出院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        If Not IsNull(mrsInfo!险类) And Not IsNull(mrsInfo!入院日期) Then
            If Format(txtDate.Text, txtDate.Format) < Format(mrsInfo!入院日期, txtDate.Format) Then
                MsgBox "费用的发生时间不能小于医保病人的入院时间:" & Format(mrsInfo!入院日期, txtDate.Format), vbInformation, gstrSysName
                txtDate.SetFocus: Exit Sub
            End If
        End If
        
        '医保负数记帐检查    因为操作员可能先输单据,再确定病人,所以要再检查一次
        If InStr(mstrPrivsOpt, "诊疗负数记帐") > 0 And (mbytUseType = Use住院 Or mbytUseType = Use科室分散) Then      '至少有其中一种负数记帐权限,才可能是负数
            If Not IsNull(mrsInfo!险类) Then
                If Not MCPAR.负数记帐 Then
                    For i = 1 To mobjBill.Details.Count
                        If mobjBill.Details(i).数次 < 0 Then
                                MsgBox "单据中第 " & i & " 行是负数,本地医保不支持负数记帐！", vbInformation, gstrSysName
                                txtDate.SetFocus: Exit Sub
                        End If
                    Next
                End If
            End If
        End If
        '并发操作记帐权限检查
        If mbytUseType <> Use门诊 Then
            If Not PatiCanBilling(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0), mstrPrivsOpt) Then Exit Sub
            If zlPatiIS病案已编目(mrsInfo!病人ID, NVL(mrsInfo!主页ID, 0)) = True Then Exit Sub
            If zlIsAllowFeeChange(Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID))) = False Then Exit Sub             '问题:49501
        End If
        
        '医保费用项目审批检查
        If mbytUseType <> Use门诊 Then
            If Not IsNull(mrsInfo!险类) Then
                If Not mrsMedAudit Is Nothing Then
                    If Not CheckExamine(mobjBill.Details, mrsMedAudit, mrsInfo!险类) Then Exit Sub
                End If
                
                If MCPAR.实时监控 Then
                    If gclsInsure.CheckItem(mrsInfo!险类, 1, 2, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text))) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
                
        '记帐分类报警
        If mbytInState = sta执行 Then
            mrsWarn.Filter = ""
            If mrsWarn.RecordCount > 0 Then
                '单据费用
                curTotal = CalcGridToTal
                If curTotal > 0 Then
                    '病人预交款信息
                    Set rsTmp = GetMoneyInfo(mrsInfo!病人ID, CDbl(mcurModiMoney), Val("" & mrsInfo!险类) > 0)
                    If Not rsTmp Is Nothing Then
                        sta.Panels(3).Text = "预交:" & Format(rsTmp!预交余额, "0.00")
                        sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(rsTmp!费用余额, gstrDec)
                        sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(rsTmp!预交余额 - rsTmp!费用余额, "0.00")
                        cmdOK.Tag = rsTmp!预交余额
                        cmdCancel.Tag = rsTmp!费用余额
                        mcur可用金额 = rsTmp!预交余额 - rsTmp!费用余额
                    Else
                        sta.Panels(3).Text = "预交:0.00/费用:" & gstrDec & "/剩余:0.00"
                        cmdOK.Tag = 0
                        cmdCancel.Tag = 0
                        mcur可用金额 = 0
                    End If
                    
                    '重新读取当日额
                    cur当日额 = GetPatiDayMoney(mrsInfo!病人ID)
                                    
                    If gbln报警包含划价费用 Then mcur可用金额 = mcur可用金额 - GetPriceMoneyTotal(2, mrsInfo!病人ID)
                    
                    For i = 1 To mobjBill.Details.Count
                        gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名, Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, mcur可用金额, cur当日额 - mcurModiMoney, curTotal, mobjBill.担保额, mobjBill.Details(i).收费类别, mobjBill.Details(i).Detail.类别名称, mstrWarn)
                        If gbytWarn = 2 Or gbytWarn = 3 Then Exit Sub
                    Next
                End If
            End If
        End If
        
        '项目服务对象检查(主要因为多了门诊留观病人)
        If mbytUseType <> Use门诊 Then
            If Check服务对象 > 0 Then Exit Sub
        End If
        
        If Not SaveBill Then
            Exit Sub
        Else
            If mstrInNO = "" Then
                sta.Panels(2) = "上一张单据:" & mobjBill.NO
                Call NewBill
                mstrInNO = ""
                If mlng病人ID <> 0 And mbytUseType = 1 Then
                    txtPatient.Text = "-" & mlng病人ID
                    Call txtPatient_KeyPress(13)
                Else
                    txtPatient.SetFocus
                End If
            Else '修改
                gblnOK = True: Unload Me
            End If
        End If
    ElseIf chk销.Value = 1 Then '退单据状态
        If mstrInNO = "" Then
            MsgBox "没有读取单据内容,不能销帐！", vbInformation, gstrSysName
            cboNO.SetFocus: Exit Sub
        End If

        '医保记帐作废上传(注意判断顺序)
        If mbytUseType <> Use门诊 Then
            intInsure = BillExistInsure(mstrInNO) '判断是否医保病人记的帐
            '去掉了医保连接匹配检查
        End If

        If mbytUseType = Use门诊 Then
            strSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "','','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If

        On Error GoTo errH
        gcnOracle.BeginTrans
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        '医保记帐作废上传
        If mbytUseType <> Use门诊 And intInsure <> 0 Then
            If MCPAR.记帐作废上传 And Not MCPAR.记帐完成后上传 Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    gcnOracle.RollbackTrans: Exit Sub
                End If
            End If
        End If
        
        gcnOracle.CommitTrans
        
        '医保记帐作废上传
        If mbytUseType <> Use门诊 And intInsure <> 0 Then
            If MCPAR.记帐作废上传 And MCPAR.记帐完成后上传 Then
                If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                    MsgBox "单据""" & mstrInNO & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        End If
        
        On Error GoTo 0

        mstrInNO = "": cboNO.Text = ""
        txtPatient.Text = "": txt年龄.Text = ""
        mcur可用金额 = 0
        Call NewBill
        chk销.Value = 0
        txtPatient.SetFocus
    End If
    gblnOK = True
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo收费类别_GotFocus(Index As Integer)
    LocateItem Index, "收费类别"
    
    If cbo收费类别(Index).ListIndex = -1 And Index > 0 Then
        cbo收费类别(Index).ListIndex = cbo收费类别(Index - 1).ListIndex
    End If
End Sub

Private Sub txt收费项目_GotFocus(Index As Integer)
    LocateItem Index, "收费项目"
    zlControl.TxtSelAll txt收费项目(Index)
End Sub

Private Sub txt计算单位_GotFocus(Index As Integer)
    LocateItem Index, "计算单位"
    zlControl.TxtSelAll txt计算单位(Index)
End Sub

Private Sub txt数次_GotFocus(Index As Integer)
    LocateItem Index, "数次"
    zlControl.TxtSelAll txt数次(Index)
End Sub

Private Sub txt标准单价_GotFocus(Index As Integer)
    LocateItem Index, "标准单价"
    zlControl.TxtSelAll txt标准单价(Index)
End Sub

Private Sub txt实收金额_GotFocus(Index As Integer)
    LocateItem Index, "实收金额"
    zlControl.TxtSelAll txt实收金额(Index)
End Sub

Private Sub txt应收金额_GotFocus(Index As Integer)
    LocateItem Index, "应收金额"
    zlControl.TxtSelAll txt应收金额(Index)
End Sub

Private Sub cbo执行科室_GotFocus(Index As Integer)
    LocateItem Index, "执行科室"
End Sub

Private Sub chk附加_GotFocus(Index As Integer)
    LocateItem Index, "附加标志"
End Sub

Private Sub cbo收费类别_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngIdx As Long
    
    If cbo收费类别(Index).Locked Then Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        Call Input收费类别(Index)
        SendKeys "{TAB}"
        Exit Sub
    End If
    
'    If SendMessage(cbo收费类别(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
    lngIdx = zlControl.CboMatchIndex(cbo收费类别(Index).hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo收费类别(Index).ListIndex = lngIdx
End Sub

Private Sub cbo收费类别_Validate(Index As Integer, Cancel As Boolean)
    If cbo收费类别(Index).Locked Then Exit Sub
    
    Call Input收费类别(Index)
End Sub

Private Function Input收费类别(ByVal Index As Long) As Boolean
    
    If cbo收费类别(Index).ListIndex <> -1 Then
        If mobjBill.Details("R" & Index).收费类别 <> Chr(cbo收费类别(Index).ItemData(cbo收费类别(Index).ListIndex)) Then
            '一旦更改收费类别,则清除(如有)原有该项目内容
            ClearDetail Index
            'Call CalcMoneys
            Call ShowMoney
        End If
    Else
        ClearDetail Index
    End If
    
    Input收费类别 = True
End Function

Private Sub txt收费项目_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    '此项目确定,该收费细目对应的程序对象才生成
    strText = txt收费项目(Index).Text
    If strText <> "" Then
        If mobjBill.Details("R" & Index).收费类别 = "" Then
            sta.Panels(2) = "没有确定费用类别,请先输入类别！"
            If cbo收费类别(Index).Visible = True Then
                txt收费项目(Index).Text = ""
                If cbo收费类别(Index).Enabled Then cbo收费类别(Index).SetFocus
            End If
            Call Beep: Exit Sub
        End If
        Call GetDetails(txt收费项目(Index).hwnd, strText, mobjBill.Details("R" & Index).收费类别)
        If Set收费细目(Index) = False Then
            zlControl.TxtSelAll txt收费项目(Index)
            Exit Sub
        End If
    Else
        If mobjBill.Details("R" & Index).收费细目ID <> 0 Then
            Set收费细目Empty Index
        End If
        If Index = mlngRows - 1 Then
            cmdOK.SetFocus
        Else
            txt收费项目(Index + 1).SetFocus
        End If
        Exit Sub
    End If

    SendKeys "{TAB}"
End Sub

Private Sub txt收费项目_Validate(Index As Integer, Cancel As Boolean)
    Dim strText As String
    
    strText = txt收费项目(Index).Text
    
    If strText = mobjBill.Details("R" & Index).收费名称 Then Exit Sub
    If Trim(strText) = "" Then
        '空串特殊处理
        txt收费项目(Index).Text = mobjBill.Details("R" & Index).收费名称
        Exit Sub
    End If
    
    If strText <> "" Then
        If mobjBill.Details("R" & Index).收费类别 = "" Then
            sta.Panels(2) = "没有确定费用类别,请先输入类别！"
            Set mcolDetails = New Details
            Set收费细目Empty Index
            Exit Sub
        End If
        '进行输入判断
        Call GetDetails(txt收费项目(Index).hwnd, strText, mobjBill.Details("R" & Index).收费类别)
        If Set收费细目(Index) = False Then
            zlControl.TxtSelAll txt收费项目(Index)
            Cancel = True
        End If
    End If
End Sub

Private Function Set收费细目(ByVal Index As Integer) As Boolean
    Dim lngDoUnit As Long, curTotal As Currency
    Dim int病人来源 As Integer, curItemMoney As Currency
    
    If mcolDetails.Count = 0 Then
        sta.Panels(2) = "找不到相应的收费项目,请确定输入是否正确！"
        Call Beep: Exit Function
    ElseIf mcolDetails.Count = 1 Then
        '确定了收费细目
        Set mobjDetail = mcolDetails(1)
        
        '一些输入项目的合法性检查
        '保险支付项目对应检查
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) Then
                If Not CheckMediCareItem(mobjDetail.ID, mrsInfo!险类, mobjDetail.名称, mobjDetail.变价 = False, mstrPriceGrade) Then
                    txt收费项目(Index).Text = mobjBill.Details("R" & Index).Detail.名称
                    Exit Function
                End If
                
                '医保病人费用项目要求审批
                If mbytUseType <> Use门诊 Then
                    If mobjDetail.要求审批 And Not mrsMedAudit Is Nothing Then
                        mrsMedAudit.Filter = "项目ID=" & mobjDetail.ID
                        If mrsMedAudit.RecordCount = 0 Then
                            MsgBox "当前病人未被批准使用该项目！", vbInformation, gstrSysName
                            txt收费项目(Index).Text = "": Exit Function
                        ElseIf Not IsNull(mrsMedAudit!可用数量) Then
                            If mrsMedAudit!可用数量 <= 0 Then
                                MsgBox "当前病人使用[" & mobjDetail.名称 & "]已达到批准的使用限量" & mrsMedAudit!使用限量 & "。", vbInformation, gstrSysName
                                txt收费项目(Index).Text = "": Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If mobjDetail.ID = mobjBill.Details("R" & Index).Detail.ID Then
           '仍然是以前的那个，所以用不着再改变了
           txt收费项目(Index).Text = mobjDetail.名称
           Set收费细目 = True
           Exit Function
        End If
        
        '病人来源
        If mbytUseType = Use门诊 Then
            int病人来源 = 1
        Else
            If mrsInfo.State = 1 Then
                '读取病人时已根据权限限制是否留观病人
                If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
                    int病人来源 = 2
                ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
                    int病人来源 = 1
                End If
            Else
                int病人来源 = 2
            End If
        End If
        '病人科室
        lngDoUnit = mobjBill.科室ID
        If lngDoUnit = 0 Then lngDoUnit = Get开单科室ID
        lngDoUnit = Get收费执行科室ID(mobjDetail.ID, mobjDetail.执行科室, lngDoUnit, Get开单科室ID, int病人来源, mobjBill.病区ID)

        '加入或修改该收费细目行
        With mobjBill.Details("R" & Index)
            Set .Detail = mobjDetail
            Set .InComes = New BillInComes
            .附加标志 = 0
            .计算单位 = mobjDetail.计算单位
            .收费类别 = mobjDetail.类别
            .收费细目ID = mobjDetail.ID
            .收费名称 = mobjDetail.名称
            
            If txt数次(Index).Tag <> "" Then
                .数次 = Val(txt数次(Index).Tag)
            Else
                .数次 = 1
            End If
            .执行部门ID = lngDoUnit
            '计算单价和金额
            Call CalcMoney(Index)
            
            '项目的值的预设的
        End With
        
        '记帐分类报警(在已经算出该行费用但未显示前)
        If mbytInState = sta执行 Then
            mrsWarn.Filter = ""
            If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
                curTotal = GetBillTotal(mobjBill)
                If curTotal > 0 Then
                    If gbln报警包含划价费用 Then mcur可用金额 = mcur可用金额 - GetPriceMoneyTotal(2, mrsInfo!病人ID)
                    
                    '刘兴洪:24491
                    curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
                    gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名, Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, mcur可用金额, mrsInfo!当日额 - mcurModiMoney, curTotal, mobjBill.担保额, mobjDetail.类别, mobjDetail.类别名称, mstrWarn, , curItemMoney)
                    If gbytWarn = 2 Or gbytWarn = 3 Then
                        ClearDetail Index '删除刚刚想要加入的费用行
                        txt收费项目(Index).Text = ""
                        Exit Function
                    End If
                End If
            End If
        End If
        If mrsInfo.State = 1 And mbytUseType <> Use门诊 Then
            If Not IsNull(mrsInfo!险类) And MCPAR.实时监控 Then
                If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), Index)) = False Then
                    ClearDetail Index '删除刚刚想要加入的费用行
                    txt收费项目(Index).Text = ""
                    Exit Function
                End If
            End If
        End If
        

        Call ShowDetails(Index)
        Call ShowMoney
    End If

    If mobjBill.Details("R" & Index).Detail.变价 Then
        txt数次(Index).TabStop = gblnTime
        txt数次(Index).Locked = Not gblnTime
        txt标准单价(Index).TabStop = True
        txt标准单价(Index).Locked = False
    Else
        txt数次(Index).TabStop = True
        txt数次(Index).Locked = False
        txt标准单价(Index).TabStop = False
        txt标准单价(Index).Locked = True
    End If
    chk附加(Index).Enabled = mobjBill.Details("R" & Index).收费类别 = "F" '手术
    If chk附加(Index).Enabled = False Then chk附加(Index).Value = 0
    
    '执行科室!!!
    Call Fill执行科室(Index)
    
    If cbo执行科室(Index).ListCount = 1 Then
        cbo执行科室(Index).TabStop = False
    Else
        cbo执行科室(Index).TabStop = True
    End If
        
    'byZT200302
    'cbo执行科室(Index).SelLength = 0
    
    Set收费细目 = True
End Function

Private Function Set收费细目Empty(Index As Integer) As Boolean
    Dim lngDoUnit As Long
    

    mobjBill.Details.Remove "R" & Index
    mobjBill.Details.AddEmpty Index + 1
    mobjBill.Details("R" & Index).收费类别 = cbo收费类别(Index).Tag
    
    If Val(txt收费项目(Index).Tag) > 0 Then
        Call GetInputDetail(Val(txt收费项目(Index).Tag))
        Call Set收费细目(Index)
    Else
        Call ShowDetails(Index)
    End If
    Call ShowMoney

    txt数次(Index).TabStop = False
    txt数次(Index).Locked = False
    txt标准单价(Index).TabStop = False
    txt标准单价(Index).Locked = False
    chk附加(Index).Enabled = False
    chk附加(Index).Value = 0
    
    Set收费细目Empty = True
End Function

Private Sub cmd细目选择_Click(Index As Integer)
    Dim strSQL As String, str特准项目 As String
    Dim str类别 As String, lng项目id As Long
    Dim int病人来源 As Integer, int险类 As Integer
    
    Call LocateItem(Index, "细目选择")
    
    '收费类别
    str类别 = mobjBill.Details("R" & Index).收费类别
    If str类别 <> "" Then str类别 = "'" & str类别 & "'"
        
    '病人来源
    If mbytUseType = -1 Then '设计
        int病人来源 = 0
    ElseIf mbytUseType = Use门诊 Then
        int病人来源 = 1
    Else
        If mrsInfo.State = 1 Then
            '读取病人时已根据权限限制是否留观病人
            If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
                int病人来源 = 2
            ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
                int病人来源 = 1
            End If
        Else
            '未确定病人,不限制,在保存时检查
            If (InStr(mstrPrivsOpt, "门诊留观记帐") > 0 And gbln门诊留观) Or mbytUseType = 2 Then
                int病人来源 = 0
            Else
                int病人来源 = 2
            End If
        End If
    End If
    If mbytUseType <> -1 Then
        '医保病人特准项目
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) Then
                int险类 = mrsInfo!险类
                '刘兴洪:24862
                If zl_Check特准项目(gclsInsure, int险类, Val(NVL(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(NVL(mrsInfo!病人ID)), "A.ID")
            End If
        End If
    End If
    
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, str类别, , , str特准项目, mstrPriceGrade)
    Me.Refresh
    txt收费项目(Index).SetFocus
    
    If lng项目id <> 0 Then
        If lng项目id = mobjBill.Details("R" & Index).收费细目ID Then
            SendKeys "{TAB}": Exit Sub
        End If
        Call GetInputDetail(lng项目id)
        If Not Set收费细目(Index) Then Exit Sub
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txt数次_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If txt数次(Index).Locked = True Then Exit Sub
    If KeyAscii <> vbKeyReturn Then
        If InStr("-.0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    If txt数次(Index).Text = "" Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If IsNumeric(txt数次(Index).Text) = False Then
        MsgBox "请输入合法数量。", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt数次(Index)
        Exit Sub
    End If
    If Input数次(Index) = True Then
        SendKeys "{TAB}"
    Else
        zlControl.TxtSelAll txt数次(Index)
    End If
End Sub

Private Sub txt数次_Validate(Index As Integer, Cancel As Boolean)
    If txt数次(Index).Locked Then Exit Sub
    
    If Not IsNumeric(txt数次(Index).Text) Or Val(txt数次(Index).Text) > 100000 Then
        If mobjBill.Details("R" & Index).数次 = 0 Then
            txt数次(Index).Text = ""
        Else
            txt数次(Index).Text = mobjBill.Details("R" & Index).数次
        End If
        Exit Sub
    Else
        If CSng(txt数次(Index).Text) = mobjBill.Details("R" & Index).数次 Then Exit Sub
    End If
    
    If Input数次(Index) Then
        Cancel = True
    End If
End Sub

Private Function Input数次(ByVal Index As Long) As Boolean
    Dim sngPreTime  As Single, sngItemNum As Single
    Dim sngInput As Single, curTotal As Currency, curItemMoney As Currency
    Dim dbl结帐数量 As Double
    
    With mobjBill.Details("R" & Index)
        If Val(txt数次(Index).Text) > 100000 Then
            MsgBox "数次值输入过大！", vbInformation, gstrSysName
            txt数次(Index).Text = .数次
            Exit Function
        End If
        If .Detail.录入限量 > 0 And Val(txt数次(Index).Text) > .Detail.录入限量 Then
            If MsgBox("输入的数次超过了录入限量" & .Detail.录入限量 & ",是否继续?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                txt数次(Index).Text = .数次
                Exit Function
            End If
        End If
        '审批限量
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) And .Detail.要求审批 And Not mrsMedAudit Is Nothing Then
                mrsMedAudit.Filter = "项目ID=" & .收费细目ID
                If mrsMedAudit.RecordCount > 0 Then
                    If Not IsNull(mrsMedAudit!可用数量) Then
                        If Val(txt数次(Index).Text) > mrsMedAudit!可用数量 Then
                            MsgBox "输入的数次超过了批准的使用限量" & mrsMedAudit!可用数量 & "。", vbInformation, gstrSysName
                            txt数次(Index).Text = .数次
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If

        '最大金额检查
        If gcurMaxMoney > 0 Then
            If CSng(txt数次(Index).Text) * Val(txt标准单价(Index).Text) > gcurMaxMoney Then
                If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    txt数次(Index).Text = .数次
                    Exit Function
                End If
            End If
        End If
    End With
    
    
    sngInput = Format(Val(txt数次(Index).Text), "0.000")
    If sngInput < 0 Then
        '负数权限检查
        If (mbytUseType = Use住院 Or mbytUseType = Use科室分散) Then
            If InStr(mstrPrivsOpt, "诊疗负数记帐") = 0 Then
                MsgBox "你没有权限输入负数！", vbInformation, gstrSysName
                txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                Exit Function
            Else
                If mrsInfo.State = 1 Then
                    If Not IsNull(mrsInfo!险类) Then
                        If Not MCPAR.负数记帐 Then
                            MsgBox "本地医保不支持对医保病人进行负数记帐！", vbInformation, gstrSysName
                            txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        '问题:26951
         If InStr(1, mstrPrivsOpt, ";负数记帐不检查发生项目;") = 0 Then
             '对于负数冲销时不检查本次住院发生的项目数量,有此权限,允许录入病人未曾发生的费用项目进行冲销,否则检查本次住院发生的项目数量才能冲销
            '负数合法性检查
            sngItemNum = GetDetailNum(Index, dbl结帐数量)
            '32106
            If Abs(sngInput) > sngItemNum - dbl结帐数量 Then
                Select Case gbytBillOpt '对已结帐的记帐单据的操作权限:0-允许,1-提醒,2-禁止。
                Case 0  '允许
                    If Abs(sngInput) > sngItemNum Then
                        MsgBox "该项目冲销数量多于已有数量[" & sngItemNum & "]！", vbInformation, gstrSysName
                        txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                        Exit Function
                    End If
                Case 1   '提醒
                    If Abs(sngInput) > sngItemNum Then
                        MsgBox "该项目冲销数量多于已有数量[" & sngItemNum & "]！", vbInformation, gstrSysName
                        txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                        Exit Function
                    End If
                    If MsgBox("该项目冲销数量中包含了已结部分(未结:" & Round(sngItemNum - dbl结帐数量, 5) & "; 已结:" & Round(dbl结帐数量, 5) & ") 。" & vbCrLf & _
                        " 是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                        Exit Function
                    End If
                Case 2   '禁止
                    MsgBox "该项目冲销数量多于已有数量[" & sngItemNum & "]！", vbInformation, gstrSysName
                    txt数次(Index).Text = mobjBill.Details("R" & Index).数次
                    Exit Function
                End Select
            End If
         End If
    End If

    '记录更改前的数次以便取消计算
    sngPreTime = mobjBill.Details("R" & Index).数次
    '更改该行数次
    mobjBill.Details("R" & Index).数次 = sngInput
    Call CalcMoneys(Index)

    '记帐分类报警(在已经算出该行费用但未显示前)
    mrsWarn.Filter = ""
    If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
        curTotal = GetBillTotal(mobjBill)
        If curTotal > 0 Then
    
            If gbln报警包含划价费用 Then mcur可用金额 = mcur可用金额 - GetPriceMoneyTotal(2, mrsInfo!病人ID)
            '刘兴洪:24491
            curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
            
            gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名, Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, mcur可用金额, mrsInfo!当日额 - mcurModiMoney, curTotal, mobjBill.担保额, mobjBill.Details("R" & Index).收费类别, mobjBill.Details("R" & Index).Detail.类别名称, mstrWarn, , curItemMoney)
            If gbytWarn = 2 Or gbytWarn = 3 Then
                mobjBill.Details("R" & Index).数次 = sngPreTime
                txt数次(Index).Text = sngPreTime
                Call CalcMoneys(Index)
                Exit Function
            End If
        End If
    End If
    
    If mrsInfo.State = 1 And mbytUseType <> Use门诊 Then
        If Not IsNull(mrsInfo!险类) And MCPAR.实时监控 Then
            If gclsInsure.CheckItem(mrsInfo!险类, 1, 0, MakeDetailRecord(mobjBill, zlStr.NeedName(cbo开单人.Text), zlStr.NeedName(cbo开单科室.Text), Index)) = False Then
                mobjBill.Details("R" & Index).数次 = sngPreTime
                txt数次(Index).Text = sngPreTime
                Call CalcMoneys(Index)
                Exit Function
            End If
        End If
    End If

    Call ShowDetails(Index)
    Call ShowMoney
    Input数次 = True
End Function

Private Sub txt标准单价_KeyPress(Index As Integer, KeyAscii As Integer)
    If txt标准单价(Index).Locked Then Exit Sub
    
    If KeyAscii <> 13 Then
        If InStr(".0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    If Not IsNumeric(txt标准单价(Index).Text) Then
        MsgBox "请输入合法单价。", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt标准单价(Index)
        Exit Sub
    End If
    
    If Input标准单价(Index) Then
        SendKeys "{TAB}"
    Else
        zlControl.TxtSelAll txt标准单价(Index)
    End If
End Sub

Private Sub txt标准单价_Validate(Index As Integer, Cancel As Boolean)
    If txt标准单价(Index).Locked Then Exit Sub
    
    If Not IsNumeric(txt标准单价(Index).Text) Then
        txt标准单价(Index).Text = Format(mobjBill.Details("R" & Index).标准单价, "0,000")
        Exit Sub
    Else
        If CCur(txt标准单价(Index).Text) = mobjBill.Details("R" & Index).标准单价 Then Exit Sub
    End If
    
    If Not Input标准单价(Index) Then
        Cancel = True
    End If
End Sub

Private Function Input标准单价(ByVal Index As Long) As Boolean
    Dim strScope As String, curTotal As Currency
    Dim curPreMoney As Currency, curItemMoney As Currency
    
    '如果没有对应的收入项目,则无法计算
    If mobjBill.Details("R" & Index).Detail.变价 And mobjBill.Details("R" & Index).InComes.Count > 0 Then
        '单价不允许输入负数
        If Val(txt标准单价(Index).Text) < 0 Then
            MsgBox "项目价格不应该为负数，要冲销费用，请输入负的数量来实现！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '检查变价输入范围
        If Not (mobjBill.Details("R" & Index).InComes(1).现价 = 0 And mobjBill.Details("R" & Index).InComes(1).原价 = 0) Then
            strScope = CheckScope(mobjBill.Details("R" & Index).InComes(1).原价, mobjBill.Details("R" & Index).InComes(1).现价, CCur(txt标准单价(Index).Text))
            If strScope <> "" Then
                sta.Panels(2) = strScope
                Exit Function
            End If
        End If
        '最大金额检查
        If gcurMaxMoney > 0 Then
            If Val(txt标准单价(Index).Text) * Val(mobjBill.Details("R" & Index).数次) > gcurMaxMoney Then
                If MsgBox("当前金额超过了" & gcurMaxMoney & ",你确定要继续吗?", vbInformation + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
                    Exit Function
                End If
            End If
        End If

        curPreMoney = mobjBill.Details("R" & Index).InComes(1).标准单价

        mobjBill.Details("R" & Index).InComes(1).标准单价 = txt标准单价(Index).Text '这种收费细目只能对应一个收入项目
        Call CalcMoneys(Index)

        '记帐分类报警(在已经算出该行费用但未显示前)
        mrsWarn.Filter = ""
        If mrsInfo.State = 1 And mrsWarn.RecordCount > 0 Then
            curTotal = GetBillTotal(mobjBill)
            If curTotal > 0 Then
        
                If gbln报警包含划价费用 Then mcur可用金额 = mcur可用金额 - GetPriceMoneyTotal(2, mrsInfo!病人ID)
                '刘兴洪:24491
                curItemMoney = GetBillRowTotal(mobjBill.Details("R" & Index).InComes)
                gbytWarn = BillingWarn(mstrPrivsOpt, mrsInfo!姓名, Val("" & mrsInfo!病区ID), mrsInfo!适用病人, mrsWarn, mcur可用金额, mrsInfo!当日额 - mcurModiMoney, curTotal, mobjBill.担保额, mobjBill.Details("R" & Index).收费类别, mobjBill.Details("R" & Index).Detail.类别名称, mstrWarn)
                If gbytWarn = 2 Or gbytWarn = 3 Then
                    mobjBill.Details("R" & Index).InComes(1).标准单价 = curPreMoney
                    txt标准单价(Index).Text = Format(curPreMoney, "0.0000")
                    If Val(txt标准单价(Index).Text) = 0 Then txt标准单价(Index).Text = ""
                    
                    Call CalcMoneys(Index)
                    Exit Function
                End If
            End If
        End If

        Call ShowDetails(Index)
        Call ShowMoney
    Else
        txt标准单价(Index).Text = "0"
        sta.Panels(2) = "该项目设有设置对应的费目，所以无法计算费用！"
        Beep
    End If
    Input标准单价 = True
End Function

Private Sub cbo执行科室_KeyPress(Index As Integer, KeyAscii As Integer)
        

    Dim lngIdx As Long, lng医生ID As Long
    
    If KeyAscii <> 13 Then Exit Sub
    If cbo执行科室(Index).ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    Fill执行科室 Index, True
    If zlSelectDept(Me, 0, cbo执行科室(Index), mrsUnit, cbo执行科室(Index)) = False Then
        KeyAscii = 0: Exit Sub
    End If
    
'    Exit Sub
'    Dim lngIdx As Long
'    If KeyAscii = 13 Then
'        KeyAscii = 0 'byZT200302
'        If cbo执行科室(Index).ListIndex <> -1 Then
'            mobjBill.Details("R" & Index).执行部门ID = cbo执行科室(Index).ItemData(cbo执行科室(Index).ListIndex)
'        End If
'        SendKeys "{TAB}"
'        Exit Sub
'    End If
'
'    If cbo执行科室(Index).Locked Then Exit Sub
'    If SendMessage(cbo执行科室(Index).hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then SendKeys "{F4}"
'    lngIdx = MatchIndex(cbo执行科室(Index).hwnd, KeyAscii)
'    If lngIdx <> -2 Then cbo执行科室(Index).ListIndex = lngIdx
'
'    'byZT200302
'    If cbo执行科室(Index).ListIndex = -1 And cbo执行科室(Index).ListCount <> 0 Then cbo执行科室(Index).ListIndex = 0
End Sub

Private Sub cbo执行科室_Validate(Index As Integer, Cancel As Boolean)
    If cbo执行科室(Index).ListIndex = -1 And cbo执行科室(Index).ListCount <> 0 Then cbo执行科室(Index).ListIndex = 0
    
    If cbo执行科室(Index).ListIndex <> -1 Then
        mobjBill.Details("R" & Index).执行部门ID = cbo执行科室(Index).ItemData(cbo执行科室(Index).ListIndex)
    Else
        mobjBill.Details("R" & Index).执行部门ID = 0
    End If
    
End Sub

Private Sub chk附加_Click(Index As Integer)
'说明：可以全部为主要手术,但不能全部为附加手术
    Dim i As Long, strCheck As String, bytTime As Byte

    '新增的未处理行无效
    
    For i = 0 To mlngRows - 1
        If mobjBill.Details("R" & i).收费类别 = "F" And chk附加(i).Value = 0 And i <> Index Then bytTime = bytTime + 1
    Next
    If bytTime > 0 Then
        mobjBill.Details("R" & Index).附加标志 = chk附加(Index).Value
        Call CalcMoneys(Index)
        Call ShowDetails(Index)
        Call ShowMoney
    ElseIf chk附加(Index).Value = 1 Then
        chk附加(Index) = 0
        MsgBox "单据中必然有一个手术不是附加手术！", vbInformation, gstrSysName
        Exit Sub
    End If
    mobjBill.Details("R" & Index).附加标志 = chk附加(Index).Value
End Sub

Private Sub Fill执行科室(ByVal lngRow As Long, Optional blnNotLoad As Boolean = False)
'功能：根据单据列设置下拉列表框内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String, i As Long
    Dim lng科室ID As Long, lng病区ID As Long, int病人来源 As Integer
    
    If mbytInState <> sta执行 Then Exit Sub

    '病人来源
    If mbytUseType = Use门诊 Then
        int病人来源 = 1
    Else
        If mrsInfo.State = 1 Then
            '读取病人时已根据权限限制是否留观病人
            If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
                int病人来源 = 2
            ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
                int病人来源 = 1
            End If
        Else
            int病人来源 = 2
        End If
    End If
    
    '病人科室
    lng科室ID = mobjBill.科室ID
    If lng科室ID = 0 Then lng科室ID = Get开单科室ID
    
    If int病人来源 = 1 Then
        lng病区ID = lng科室ID
    Else
        lng病区ID = mobjBill.病区ID
        If lng病区ID = 0 Then lng病区ID = Get病区ID(lng科室ID)
    End If
    
    '0-不明确,1-病人科室,2-病人病区,3-开单人科室,4-指定科室
    Select Case mobjBill.Details("R" & lngRow).Detail.执行科室
        Case 0 '不明确
            mrsUnit.Filter = 0
        Case 1 '病人科室
            mrsUnit.Filter = "ID=" & lng科室ID & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
        Case 2 '病人病区
            mrsUnit.Filter = "ID=" & lng病区ID & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
        Case 3 '操作员所在科室
            mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
        Case 4 '指定科室
            strSQL = "" & _
            "   Select Nvl(A.开单科室ID,0) as 开单科室ID,A.执行科室ID" & _
            "   From 收费执行科室 A,部门表 C" & _
            "   Where A.收费细目ID=[1]　And A.执行科室ID+0=C.ID " & _
            "       And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
            "       And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) " & vbNewLine & _
            "       And (A.病人来源 is NULL Or A.病人来源=[3])" & _
            "       And (A.开单科室ID is NULL Or A.开单科室ID=[2])" & _
            " Order by Decode(A.病人来源,Null,2,1)" '默认科室优先
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details("R" & lngRow).收费细目ID, lng科室ID, int病人来源)
            If Not rsTmp.EOF Then
                For i = 1 To rsTmp.RecordCount
                    strTmp = strTmp & "ID=" & rsTmp!执行科室ID & " OR "
                    rsTmp.MoveNext
                Next
                strTmp = strTmp & "ID=" & mobjBill.Details("R" & lngRow).执行部门ID & " OR "
                strTmp = Left(strTmp, Len(strTmp) - 4)
                mrsUnit.Filter = strTmp
            Else
                mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
            End If
        Case 5 '院外执行(预留,程序暂未用)
        Case 6 '开单人科室
           mrsUnit.Filter = "ID=" & Get开单科室ID & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
    End Select
    If mrsUnit.EOF Then mrsUnit.Filter = "ID=" & IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID) & " Or ID=" & mobjBill.Details("R" & lngRow).执行部门ID
    If blnNotLoad = True Then Exit Sub
    
    With cbo执行科室(lngRow)
        .Clear
        For i = 1 To mrsUnit.RecordCount
            strTmp = IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
            If Not (SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strTmp) >= 0) Then
                .AddItem strTmp
                .ItemData(.NewIndex) = mrsUnit!ID
                
                '缺省为开单科室
                If lngRow = 1 Then
                    If mrsUnit!ID = mobjBill.开单部门ID Then .ListIndex = .NewIndex
                '设成与上一行执行科室一样
                ElseIf lngRow > 1 Then
                    If mrsUnit!ID = mobjBill.Details("R" & (lngRow - 1)).执行部门ID And mobjBill.Details("R" & lngRow).Detail.执行科室 = mobjBill.Details("R" & (lngRow - 1)).Detail.执行科室 Then
                        .ListIndex = .NewIndex
                    ElseIf mrsUnit!ID = mobjBill.开单部门ID And .ListIndex = -1 Then
                        .ListIndex = .NewIndex
                    End If
                End If
            End If
            mrsUnit.MoveNext
        Next
        
        If mobjBill.Details("R" & lngRow).Detail.执行科室 = 4 Then   '执行科室为指定科室的,缺省为操作员所在科室
            For i = 0 To .ListCount - 1
                If .ItemData(i) = UserInfo.部门ID Then .ListIndex = i: Exit For
            Next
        End If
        
        If .ListIndex = -1 Then '如果没有则取现有的执行科室
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mobjBill.Details("R" & lngRow).执行部门ID Then .ListIndex = i: Exit For
            Next
        End If
        
        If .ListIndex = -1 And .ListCount <> 0 Then .ListIndex = 0
        If .ListIndex <> -1 Then
            mobjBill.Details("R" & lngRow).执行部门ID = .ItemData(.ListIndex)
        Else
            mobjBill.Details("R" & lngRow).执行部门ID = 0
        End If
        
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub
Private Sub LocateItem(ByVal Index As Integer, ByVal Item As String)
    mintCurrentRow = Index
End Sub

Private Sub SetDisible(Optional bln As Boolean = False)
'界面设置为不可修改状态
    cboNO.Locked = Not bln
    txtPatient.Locked = Not bln
    cbo开单科室.Locked = Not bln
    cbo开单人.Locked = Not bln

    chk加班.Enabled = bln
    txtDate.Enabled = bln
    fraForm.Enabled = bln
End Sub

Private Function GetPatientIn(ByVal strInput As String, ByVal blnCard As Boolean, Optional blnOutMsg As Boolean = False) As Boolean
'功能：获取病人信息
'参数：blnCard=是否就诊卡刷卡
    Dim strSQL As String, strIF As String
    
    On Error GoTo errH
        
    'a.是否具有强制记帐权限
    If InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 And InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
        strIF = ""
    ElseIf InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)<>0)"
    ElseIf InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
        strIF = " And ((B.出院日期 is NULL And Nvl(B.状态,0)<>3) Or Nvl(X.费用余额,0)=0)"
    Else
        strIF = " And B.出院日期 is NULL And Nvl(B.状态,0)<>3"
    End If
    
    'b.是否可以记所有病区病人
     If (mbytUseType = Use住院 Or mbytUseType = Use科室分散) And InStr(mstrPrivs, "所有病区") <= 0 Then
        If InStr(1, mstrUnitIDs, ",") = 0 Then
            strIF = strIF & " And B.当前病区ID+0=[3]"
        Else
            strIF = strIF & " And B.当前病区ID+0 IN(Select * From Table(Cast(f_num2list([4]) As zlTools.t_numlist)))"
        End If
    End If
       
    'c.是否留观病人记帐权限
    If (InStr(mstrPrivsOpt, "门诊留观记帐") > 0 And gbln门诊留观) And (InStr(mstrPrivsOpt, "住院留观记帐") > 0 And gbln住院留观) Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1,2)"
    ElseIf InStr(mstrPrivsOpt, "门诊留观记帐") > 0 And gbln门诊留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,1)"
    ElseIf InStr(mstrPrivsOpt, "住院留观记帐") > 0 And gbln住院留观 Then
        strIF = strIF & " And Nvl(B.病人性质,0) IN(0,2)"
    Else
        strIF = strIF & " And Nvl(B.病人性质,0)=0"
    End If
    
    '费用标志-->审核标志:58629
    strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期,C.名称 as 病区,D.名称 as 科室," & _
            "   A.就诊卡号,A.卡验证码,A.住院号 as 标识号,B.出院病床 as 床号, " & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(B.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额, " & _
            "   Zl_Patiwarnscheme(B.病人id, B.主页id) As 适用病人,B.险类,Nvl(B.病人性质,0) as 病人性质,b.审核标志,B.病人类型" & _
            " From 病人信息 A,病案主页 B,部门表 C,部门表 D,病人余额 X " & _
            " Where A.病人ID=B.病人ID And A.住院次数=B.主页ID And B.当前病区ID=C.ID(+) And B.出院科室ID=D.ID(+) " & _
            " And Nvl(B.主页ID,0)<>0 And A.病人ID=X.病人ID(+) And A.停用时间 is NULL " & strIF
        If mbytUseType <> Use门诊 Then '问题:49501
            strSQL = strSQL & " And X.类型(+)=2 and X.性质(+)=1"
        Else
            strSQL = strSQL & " And X.类型(+)=1 and X.性质(+)=1"
        End If
    If blnCard Then '就诊卡号
        strInput = UCase(strInput)
        strSQL = strSQL & " And A.就诊卡号=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "/" And IsNumeric(Mid(strInput, 2)) Then '床位号
        If mlngUnitID = 0 Then '病区不确定、则不能通过床号确定病人
            Set mrsInfo = New ADODB.Recordset: Exit Function
        End If
        strSQL = _
            "Select A.病人ID,B.主页ID,B.当前病区ID as 病区ID,B.出院科室ID as 科室ID,B.入院日期,B.出院日期," & _
            "   A.就诊卡号,A.卡验证码,A.住院号 as 标识号,B.出院病床 as 床号," & _
            "   nvl(B.姓名,A.姓名) as 姓名,nvl(B.性别,A.性别) as 性别,A.年龄,B.费别,B.住院医师,B.医疗付款方式," & _
            "   A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,B.主页ID)) 担保额,zl_PatiDayCharge(A.病人ID) as 当日额," & _
            "   Zl_Patiwarnscheme(B.病人id, B.主页id) As 适用病人,B.险类,Nvl(B.病人性质,0) as 病人性质,b.审核标志,B.病人类型" & _
            " From 病人信息 A,病案主页 B,床位状况记录 C,病人余额 X" & _
            " Where A.病人ID=B.病人ID And A.住院次数=B.主页ID" & _
            "   And Nvl(B.主页ID,0)<>0 And A.病人ID=C.病人ID And A.病人ID=X.病人ID(+) And A.停用时间 is NULL " & _
            "   And C.病区ID=[3] And C.床号=[1] " & strIF
            
        If mbytUseType <> Use门诊 Then  '问题:49501
            strSQL = strSQL & " And X.类型(+)=2 and X.性质(+)=1"
        Else
            strSQL = strSQL & " And X.类型(+)=1 and X.性质(+)=1"
        End If
            
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
        strSQL = strSQL & " And A.门诊号=[1]"
    Else '当作姓名
        strSQL = strSQL & " And A.姓名=[2]"
    End If
        
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, mlngUnitID, mstrUnitIDs)
    
    txtPatient.ForeColor = Me.ForeColor
    If Not mrsInfo.EOF Then
        If zlPatiIS病案已编目(Val(NVL(mrsInfo!病人ID)), NVL(mrsInfo!主页ID, 0)) Then
            Set mrsInfo = Nothing
            Set mrsMedAudit = Nothing
            blnOutMsg = True
            Exit Function
        End If
        If mbytUseType <> Use门诊 Then
            '问题:49501
            If zlIsAllowFeeChange(Val(NVL(mrsInfo!病人ID)), Val(NVL(mrsInfo!主页ID)), Val(NVL(mrsInfo!审核标志))) = False Then
                Set mrsInfo = Nothing
                Set mrsMedAudit = Nothing
                blnOutMsg = True
                Exit Function
            End If
        End If
        
        txtPatient.ForeColor = zlDatabase.GetPatiColor(NVL(mrsInfo!病人类型))
        Set mrsMedAudit = GetAuditRecord(mrsInfo!病人ID, mrsInfo!主页ID)
        GetPatientIn = True: Exit Function
    Else
        Set mrsMedAudit = Nothing   '医保病人必须在院才检查费用审批
    End If
    
        
    '医技科室记帐：没有发现住院(在院或出院)病人,以门诊病人读
    If mbytUseType = 2 And InStr(mstrPrivsOpt, "出院未结强制记帐") > 0 And InStr(mstrPrivsOpt, "出院结清强制记帐") > 0 Then
        strSQL = _
            "Select A.病人ID,Nvl(A.住院次数,0) 主页ID,A.当前病区ID as 病区ID,A.当前科室ID as 科室ID," & _
            " A.出院时间 as 出院日期,A.就诊卡号,A.卡验证码,A.住院号,A.当前床号 as 床号,A.姓名,A.性别,A.年龄," & _
            " A.入院时间 as 入院日期,A.费别,A.担保人,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,null)) 担保额" & _
            ", Zl_Patiwarnscheme(A.病人id) As 适用病人,NULL as 住院医师,A.医疗付款方式," & _
            " zl_PatiDayCharge(A.病人ID) as 当日额,A.险类,-1 as 病人性质" & _
            " From 病人信息 A Where A.停用时间 is NULL "
        If blnCard Then '就诊卡号
            strSQL = strSQL & " And A.就诊卡号=[2]"
        ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
            strSQL = strSQL & " And A.病人ID=[1]"
        ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(医技记帐)
            strSQL = strSQL & " And A.门诊号=[1]"
        Else '当作姓名
            strSQL = strSQL & " And A.姓名=[2]"
        End If
        
        Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
        If Not mrsInfo.EOF Then
                If zlPatiIS病案已编目(Val(NVL(mrsInfo!病人ID)), NVL(mrsInfo!主页ID, 0)) Then
                    Set mrsInfo = Nothing
                    blnOutMsg = True
                    Exit Function
                End If
        
            GetPatientIn = True
        Else
            Set mrsInfo = New ADODB.Recordset
        End If
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function


Private Function GetPatientOut(ByVal strInput As String) As Boolean
'功能：获取病人信息
'说明：门诊功能在院,住院功能出院,都不能读取病人信息
'字段列表：病人ID,主页ID,病区ID,科室ID,病区,科室,出院日期,就诊卡号,标识号,床号,姓名,性别,年龄,费别,担保额
    Dim strSQL As String, strNO As String
    On Error GoTo errH
    
    strSQL = _
    " Select A.病人ID,Nvl(A.住院次数,0) 主页ID,0 as 病区ID,0 as 科室ID,'' as 病区,'' as 科室," & _
    "       A.入院时间 as 入院日期,A.出院时间 as 出院日期,A.就诊卡号,A.卡验证码,A.门诊号 as 标识号,'' as 床号,A.姓名,A.性别,A.年龄," & _
    "       A.费别,Decode(A.担保额,null,A.担保额,Zl_Patientsurety(A.病人ID,null)) 担保额,A.险类,A.医疗付款方式," & _
    "       zl_PatiDayCharge(A.病人ID) as 当日额, Zl_Patiwarnscheme(A.病人id) As 适用病人,-1 as 病人性质" & _
    " From 病人信息 A " & _
    " Where A.停用时间 is NULL And Nvl(当前科室ID,0)=0 "
    
    If mblnCard Then '就诊卡号
        strInput = UCase(strInput)
        strSQL = strSQL & " and A.就诊卡号=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then '病人ID
        strSQL = strSQL & " and A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strSQL = strSQL & " and A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "." And IsNumeric(Mid(strInput, 2)) Then '挂号单号(！最后为执行部门ID)
        strNO = GetFullNO(Mid(strInput, 2), 12)
        strSQL = _
        "Select A.病人ID,0 主页ID,0 as 病区ID,执行部门ID as 科室ID,'' as 病区,'' as 科室," & _
        "       B.入院时间 as 入院日期,B.出院时间 as 出院日期,B.就诊卡号,Nvl(A.标识号,B.门诊号) as 标识号,'' as 床号," & _
        "       A.姓名,A.性别,A.年龄,A.费别,B.担保额,B.险类,B.医疗付款方式,zl_PatiDayCharge(B.病人ID) as 当日额," & _
        "       Zl_Patiwarnscheme(A.病人id) As 适用病人,-1 as 病人性质" & _
        " From 门诊费用记录 A,病人信息 B" & _
        " Where A.病人ID=B.病人ID(+) And A.记录性质=4 And A.记录状态=1" & _
             zlGetRegEventsCons("加班标志", "A") & _
        "       And A.NO=[3]"
    Else
        strSQL = strSQL & " and A.姓名=[2]"
    End If
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput, strNO)
    
    If Not mrsInfo.EOF Then
        GetPatientOut = True
    Else
        Set mrsInfo = New ADODB.Recordset
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = New ADODB.Recordset
End Function

Private Sub CalcMoneys(Optional ByVal lngRow As Long = -1)
'功能：计算或重新计算指定行或所有行的金额
'参数：lngRow=指定行,为0表示计算所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long
    If mobjBill.Details.Count = 0 Then Exit Sub
    If lngRow = -1 Then
        For i = 0 To mlngRows - 1
            CalcMoney i
        Next
    Else
        CalcMoney lngRow
    End If
End Sub

Private Sub CalcMoney(ByVal lngRow As Long)
'功能：计算或重新计算指定行的金额
'参数：lngRow=指定行
'说明：1.ExpenseBill集合的索引对应单据的行号
'      2.变价只能对应一个收入项目:mobjBill.Details("R" & lngRow).InComes(1)
'      3.如果变价细目未计算出收入项目(第一次计算),则使用默认现价
'      4.如果变价细目已经计算出收入项目(按第2步),并手动更改(也可能未改)了单价,则按该单价计算。
    Dim i As Long, strInfo As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim curMoney As Currency '用户输入的变价金额
    Dim strWherePriceGrade As String

    On Error GoTo errH
    If mstrPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (b.价格等级 = [2]" & vbNewLine & _
            "            Or (b.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From 收费价目" & vbNewLine & _
            "                               Where b.收费细目Id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
            "                                     And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And b.价格等级 Is Null"
    End If
    strSQL = _
        " Select B.收入项目ID,C.名称,C.收据费目,B.现价,B.原价,B.加班加价率,B.附术收费率,B.缺省价格 " & _
        " From 收费项目目录 A,收费价目 B,收入项目 C " & _
        " Where B.收费细目ID=A.ID And C.ID=B.收入项目ID " & _
        "   And Sysdate Between B.执行日期 And Nvl(B.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
        "   And A.ID=[1]" & vbNewLine & _
        strWherePriceGrade
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjBill.Details("R" & lngRow).收费细目ID, mstrPriceGrade)
    If rsTmp.EOF Then
        '如果没有收入项目,则清除对应的程序对象
        Set mobjBill.Details("R" & lngRow).InComes = New BillInComes
        Exit Sub
    End If
    
    If mobjBill.Details("R" & lngRow).Detail.变价 Then
        If mobjBill.Details("R" & lngRow).InComes.Count = 0 Then '第一次计算金额取缺省值
            curMoney = Val(NVL(rsTmp!缺省价格))
        Else                        '获取操作员以前输入的变价金额
            curMoney = mobjBill.Details("R" & lngRow).InComes(1).标准单价
            '如果用户输入的变价不满足变价范围，则取缺省值
            If CheckScope(Val(NVL(rsTmp!原价)), Val(NVL(rsTmp!现价)), curMoney) <> "" Then
                curMoney = Val(NVL(rsTmp!缺省价格))
            End If
        End If
    End If

    '再清除原有记录
    Set mobjBill.Details("R" & lngRow).InComes = New BillInComes

    '填写现有费用记录
    For i = 1 To rsTmp.RecordCount
        Set mobjBillIncome = New BillInCome
        With mobjBillIncome
            .收入项目ID = rsTmp!收入项目ID
            .收入项目 = rsTmp!名称
            .收据费目 = NVL(rsTmp!收据费目)
            .原价 = Val(NVL(rsTmp!原价))
            .现价 = Val(NVL(rsTmp!现价))
            If mobjBill.Details("R" & lngRow).Detail.变价 Then
                .标准单价 = Format(curMoney, "0.0000")
                '强制把数次改成 1
                'mobjBill.Details("R" & lngRow).数次 = 1
            Else
                .标准单价 = Format(IIf(IsNull(rsTmp!现价), 0, rsTmp!现价), "0.0000")
            End If

            '应收金额=单价 *  数次
            .应收金额 = .标准单价 * mobjBill.Details("R" & lngRow).数次
            '附加手术费率用计算(所有收入项目)
            If mobjBill.Details("R" & lngRow).附加标志 = 1 And mobjBill.Details("R" & lngRow).收费类别 = "F" Then
                .应收金额 = .应收金额 * IIf(IsNull(rsTmp!附术收费率), 1, rsTmp!附术收费率 / 100)
            End If
            '加班费用率计算
            If mobjBill.加班标志 = 1 And mobjBill.Details("R" & lngRow).Detail.加班加价 Then
                .应收金额 = .应收金额 + .应收金额 * IIf(IsNull(rsTmp!加班加价率), 0, rsTmp!加班加价率 / 100)
            End If

            .应收金额 = CCur(Format(.应收金额, gstrDec))

            If mobjBill.Details("R" & lngRow).Detail.屏蔽费别 Or .应收金额 = 0 Then
                .实收金额 = .应收金额
            Else
                .实收金额 = CCur(Format(ActualMoney(mobjBill.费别, .收入项目ID, .应收金额), gstrDec))
            End If
            
            '获取项目保险信息,医保病人才处理,不需要连接医保
            If mrsInfo.State = 1 And mbytUseType <> Use门诊 Then
                If Not IsNull(mrsInfo!险类) Then
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details("R" & lngRow).收费细目ID, .实收金额, False, mrsInfo!险类, _
                        mobjBill.Details("R" & lngRow).摘要 & "||" & mobjBill.Details("R" & lngRow).数次)
                    If strInfo <> "" Then
                        mobjBill.Details("R" & lngRow).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details("R" & lngRow).保险大类ID = Val(Split(strInfo, ";")(1))
                        .统筹金额 = Val(Split(strInfo, ";")(2))
                        mobjBill.Details("R" & lngRow).保险编码 = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details("R" & lngRow).摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details("R" & lngRow).Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                End If
            End If
            
            mobjBill.Details("R" & lngRow).InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, , .统筹金额
        End With
        rsTmp.MoveNext
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub ShowDetails(Optional ByVal lngRow As Long = -1)
'功能：刷新显示指定行或所有行的内容
'参数：lngRow=指定行,为-1表示显示所有行
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, curTotal As Currency

    If lngRow = -1 Then
        For i = 0 To mlngRows - 1
            ShowDetail i
        Next
    Else
        ShowDetail lngRow
    End If

    curTotal = GetBillTotal(mobjBill)

    If IsNumeric(cmdOK.Tag) Then
        sta.Panels(3).Text = "预交:" & Format(cmdOK.Tag, "0.00")
        sta.Panels(3).Text = sta.Panels(3).Text & "/费用:" & Format(CCur(cmdCancel.Tag) + curTotal, gstrDec)
        sta.Panels(3).Text = sta.Panels(3).Text & "/剩余:" & Format(mcur可用金额 - curTotal, "0.00")
    End If
End Sub

Private Sub ShowDetail(ByVal lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行，也就是类中的序号-1
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, j As Long, curMoney As Currency
    Dim objBillDetail As BillDetail
    Dim strTmp As String
    

    If lngRow > mlngRows - 1 Then Exit Sub
    Set objBillDetail = mobjBill.Details("R" & lngRow)
    
    With objBillDetail
        
        If .收费类别 <> "" Then
            For i = 0 To cbo收费类别(lngRow).ListCount - 1
                If cbo收费类别(lngRow).ItemData(i) = Asc(.收费类别) Then
                    cbo收费类别(lngRow).ListIndex = i
                    Exit For
                End If
            Next
        End If
        '刷新单据行
        '项目"
        txt收费项目(lngRow).Text = .Detail.名称
        '单位"
        txt计算单位(lngRow).Text = .Detail.计算单位
        '数次在第一次显示时已默认设置为1
        If .数次 = 0 Then
            txt数次(lngRow).Text = ""
        Else
            txt数次(lngRow).Text = .数次
        End If
        '单价是该收费细目所有收入项目的合计
        '第一次计算时是在默认数次为1的基础上计算出来的
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).标准单价
            Next
        End If
        .标准单价 = curMoney
        txt标准单价(lngRow).Text = Format(curMoney, "0.0000")
        If Val(txt标准单价(lngRow).Text) = 0 Then txt标准单价(lngRow).Text = ""
        
        '应收金额是该收费细目所有收入项目的合计
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).应收金额
            Next
        End If
        .应收金额 = curMoney
        txt应收金额(lngRow).Text = Format(curMoney, gstrDec)
        If Val(txt应收金额(lngRow).Text) = 0 Then txt应收金额(lngRow).Text = ""
        
        '实收金额是该收费细目所有收入项目的合计
        curMoney = 0
        If .InComes.Count > 0 Then
            For j = 1 To .InComes.Count
                curMoney = curMoney + .InComes(j).实收金额
            Next
        End If
        .实收金额 = curMoney
        txt实收金额(lngRow).Text = Format(curMoney, gstrDec)
        If Val(txt实收金额(lngRow).Text) = 0 Then txt实收金额(lngRow).Text = ""
        
        '执行科室"
        If mbytInState = sta执行 Then
            mrsUnit.Filter = "ID=" & .执行部门ID
            If mrsUnit.RecordCount <> 0 Then
                'byZT200302
                Call cbo.SeekIndex(cbo执行科室(lngRow), mrsUnit!名称, , True)
                If cbo执行科室(lngRow).ListIndex = -1 Then
                    cbo执行科室(lngRow).AddItem IIf(zlIsShowDeptCode, mrsUnit!编码 & "-", "") & mrsUnit!名称
                    cbo执行科室(lngRow).ItemData(cbo执行科室(lngRow).NewIndex) = .执行部门ID
                    cbo执行科室(lngRow).ListIndex = cbo执行科室(lngRow).NewIndex
                End If
                'cbo执行科室(lngRow).Text = mrsUnit!编码 & "-" & mrsUnit!名称
            Else
                'byZT200302
                strTmp = GET部门名称(.执行部门ID, mrsUnit)
                If strTmp <> "" Then
                    Call cbo.SeekIndex(cbo执行科室(lngRow), strTmp, , True)
                    If cbo执行科室(lngRow).ListIndex = -1 Then
                        cbo执行科室(lngRow).AddItem strTmp
                        cbo执行科室(lngRow).ListIndex = cbo执行科室(lngRow).NewIndex
                    End If
                End If
                'cbo执行科室(lngRow).Text = Get部门名称(.执行部门ID)
            End If
        Else
            '浏览单据只(能)显示名称
            'byZT200302
            strTmp = GET部门名称(.执行部门ID, mrsUnit)
            If strTmp <> "" Then
                cbo执行科室(lngRow).AddItem GET部门名称(.执行部门ID, mrsUnit)
            End If
            'cbo执行科室(lngRow).Text = Get部门名称(.执行部门ID)
        End If
            '标志"
        If .收费类别 = "F" Then
            chk附加(lngRow).Value = .附加标志
        Else
            chk附加(lngRow).Enabled = False
            chk附加(lngRow).Value = 0
        End If
    End With
End Sub

Public Sub ShowMoney()
'功能：刷新显示收入项目费用区
    Dim i As Integer, j As Integer
    Dim curTotal As Currency, cur应收Total As Currency


    '产生汇总费目
    For i = 1 To mobjBill.Details.Count
        For j = 1 To mobjBill.Details(i).InComes.Count
            curTotal = curTotal + mobjBill.Details(i).InComes(j).实收金额
            cur应收Total = cur应收Total + mobjBill.Details(i).InComes(j).应收金额
        Next
    Next

    txt应收.Text = Format(cur应收Total, gstrDec)
    txt实收.Text = Format(curTotal, gstrDec)
End Sub

Private Sub GetInputDetail(ByVal lng项目id As Long)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngMediCareNO As Long
            
    Set mcolDetails = New Details
    
    If mrsInfo.State = 1 Then lngMediCareNO = Val("" & mrsInfo!险类)
    If lngMediCareNO > 0 Then
        strSQL = _
            " Select" & _
            " A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.服务对象,M.要求审批,A.录入限量" & _
            " From 收费项目目录 A,收费项目类别 B,保险支付项目 M " & _
            " Where A.类别=B.编码 And A.ID=[1] And A.ID=M.收费细目ID(+) And M.险类(+)=[2]"

    Else
        strSQL = _
            " Select" & _
            " A.ID,A.类别,B.名称 as 类别名称,A.编码,A.名称,A.规格,A.计算单位," & _
            " A.屏蔽费别,A.是否变价,A.加班加价,A.执行科室,A.费用类型,A.服务对象,0 as 要求审批,A.录入限量" & _
            " From 收费项目目录 A,收费项目类别 B" & _
            " Where A.类别=B.编码 And A.ID=[1]"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng项目id, lngMediCareNO)
    
    Set mobjDetail = New Detail
    With mobjDetail
        .ID = rsTmp!ID
        .类别 = rsTmp!类别
        .类别名称 = rsTmp!类别名称
        .编码 = rsTmp!编码
        .名称 = rsTmp!名称
        .规格 = NVL(rsTmp!规格)
        .计算单位 = NVL(rsTmp!计算单位)
        .变价 = NVL(rsTmp!是否变价, 0) = 1
        .加班加价 = NVL(rsTmp!加班加价, 0) = 1
        .屏蔽费别 = NVL(rsTmp!屏蔽费别, 0) = 1
        .执行科室 = NVL(rsTmp!执行科室, 0)
        .服务对象 = NVL(rsTmp!服务对象, 0)
        .类型 = NVL(rsTmp!费用类型)
        .要求审批 = NVL(rsTmp!要求审批, 0) = 1
        .录入限量 = Val("" & rsTmp!录入限量)
        
        mcolDetails.Add .ID, .类别, .类别名称, .名称, .编码, .简码, .别名, .规格, .计算单位, .说明, .屏蔽费别, .变价, .加班加价, .执行科室, .服务对象, .类型, , , .要求审批, .录入限量
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetDetails(ByVal lngHwnd As Long, ByVal str输入 As String, Optional ByVal str类别 As String)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim int病人来源 As Integer, str特准项目 As String
    Dim lng项目id As Long, int险类 As Integer

    Set mcolDetails = New Details
    
    str输入 = UCase(str输入)
    
    '病人来源
    If mbytUseType = Use门诊 Then
        int病人来源 = 1
    Else
        If mrsInfo.State = 1 Then
            '读取病人时已根据权限限制是否留观病人
            If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
                int病人来源 = 2
            ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
                int病人来源 = 1
            End If
        Else
            '未确定病人,不限制,在保存时检查
            If (InStr(mstrPrivsOpt, "门诊留观记帐") > 0 And gbln门诊留观) Or mbytUseType = 2 Then
                int病人来源 = 0
            Else
                int病人来源 = 2
            End If
        End If
    End If
    If mbytUseType <> -1 Then
        '医保病人特准项目
        If mrsInfo.State = 1 Then
            If Not IsNull(mrsInfo!险类) Then
                int险类 = mrsInfo!险类
                '刘兴洪:24862
                If zl_Check特准项目(gclsInsure, int险类, Val(NVL(mrsInfo!病人ID)), False) Then str特准项目 = Get保险特准项目(Val(NVL(mrsInfo!病人ID)), "A.ID")
                
            End If
        End If
    End If
    
    sta.Panels("MedicareType").Text = ""
    If str类别 <> "" Then str类别 = "'" & str类别 & "'"
    lng项目id = frmItemSelect.ShowSelect(Me, mstrPrivs, int病人来源, int险类, str类别, str输入, lngHwnd, str特准项目, mstrPriceGrade)
    If lng项目id <> 0 Then
        Call GetInputDetail(lng项目id)
        If int险类 <> 0 Then sta.Panels("MedicareType").Text = Get医保大类(lng项目id, int险类)
    Else
        zlControl.TxtSelAll txt收费项目(mintCurrentRow)
    End If
End Function

Private Sub NewBill()
'功能：初始化一张新的单据(程序对象)
    Dim lngRow As Long
    
    mcurModiMoney = 0
    
    '清除单据的临时信息
    mcurPreMoney = 0: sta.Panels(3).Text = ""
    cmdOK.Tag = "": cmdCancel.Tag = "": txt实收.Tag = ""
    
    txt实收.Text = gstrDec: txt应收.Text = gstrDec

    '记帐分类报警
    mstrWarn = ""
        
    cboNO.Text = ""
    chk加班.Value = IIf(OverTime(zlDatabase.Currentdate), Checked, Unchecked)
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    Call LoadPatientBaby(cboBaby, 0, 0)
    Call cbo开单科室_Click
    
    
    '以后：设置开单科室的缺省值
    Set mrsMedAudit = Nothing
    Set mrsInfo = New ADODB.Recordset
    Call GetEmptyBill(mlngRows)
    '清除病人显示信息
    Call ShowPatient
    Call ShowMoney
    
End Sub

Private Sub GetEmptyBill(ByVal Rows As Integer)
    Dim i As Integer
    
    
    Set mobjBill = New ExpenseBill
    
    For i = 0 To Rows - 1
        mobjBill.Details.AddEmpty i + 1
        mobjBill.Details("R" & i).收费类别 = cbo收费类别(i).Tag
        '设置各项收费项目
        If Val(txt收费项目(i).Tag) > 0 Then
            Call GetInputDetail(Val(txt收费项目(i).Tag))
            Call Set收费细目(i)
        Else
            ShowDetail i
        End If
    Next
    
    With mobjBill
        .记录性质 = 2
        .记录状态 = 1
        .门诊标志 = 2
        .划价人 = UserInfo.姓名
        .开单人 = zlStr.NeedName(cbo开单人.Text)
        .操作员编号 = UserInfo.编号
        .操作员姓名 = UserInfo.姓名
        .发生时间 = CDate(txtDate.Text)
        .加班标志 = chk加班.Value
        
        If cbo开单科室.ListIndex = -1 Then
            Set开单科室 0
        Else
            Set开单科室 cbo开单科室.ItemData(cbo开单科室.ListIndex)
        End If
    End With
End Sub

Private Sub ImportBill(strNO As String, ByVal Rows As Integer, _
    Optional ByVal strPriceGrade As String)
'功能：读取费用单据到单据对象中(目前忽略从属项目,当作独立项目)
'参数：
'      strNO=单据号
'返回：存放单据信息的单据对象
'说明：因为可能现时项目价格信息已作调整,所以费用相关内容重新计算
    Dim objBillDetail As New BillDetail
    Dim objBillIncome As New BillInCome
    Dim int序号 As Integer, blnDo As Boolean, i As Integer
    Dim cur单价 As Currency, cur实收 As Currency, cur应收 As Currency
    Dim rsTmp As ADODB.Recordset, strSQL As String, strWherePriceGrade As String
    
    On Error GoTo errH
    If strPriceGrade <> "" Then
        strWherePriceGrade = _
            "       And (d.价格等级 = [3]" & vbNewLine & _
            "            Or (d.价格等级 Is Null" & vbNewLine & _
            "                And Not Exists(Select 1" & vbNewLine & _
            "                               From 收费价目" & vbNewLine & _
            "                               Where d.收费细目Id = 收费细目id And 价格等级 = [3]" & vbNewLine & _
            "                                     And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
    Else
        strWherePriceGrade = " And d.价格等级 Is Null"
    End If
    '价格父号 is NULL:只取每个具体的收费细目ID行
    '收费价目关联:新计算价格,如果有多个价格,则一个收费细目ID行就会有多条序号相同的记录
    strSQL = "Select A.ID, A.记录性质, A.NO AS 单据号, A.实际票号, A.记录状态, A.序号, A.从属父号, A.价格父号, A.记帐单id, A.病人id," & _
                    IIf(mstrFreeTable = "住院费用记录", " A.多病人单,A.主页id, A.病人病区id, A.床号,", " 0 as 多病人单,0 as 主页id,0 as 病人病区id, A.付款方式 as 床号, ") & vbNewLine & _
            "       A.医嘱序号, A.门诊标志, A.记帐费用, A.姓名, A.性别, A.年龄, A.标识号," & vbNewLine & _
            "       A.病人科室id, A.费别, A.收费类别, A.收费细目id, A.计算单位, A.付数, A.发药窗口, A.数次, A.加班标志, A.附加标志," & vbNewLine & _
            "       A.婴儿费, A.收入项目id, A.收据费目, A.标准单价, A.应收金额, A.实收金额, A.划价人, A.开单部门id, A.开单人," & vbNewLine & _
            "       A.发生时间, A.登记时间, A.执行部门id, A.执行人, A.执行状态, A.执行时间, A.结论, A.操作员编号, A.操作员姓名," & vbNewLine & _
            "       A.结帐id, A.结帐金额, A.保险大类id, A.保险项目否, A.保险编码, A.统筹金额, A.是否上传, A.摘要, A.是否急诊," & vbNewLine & _
            "       A.费用类型, B.编码, B.规格, B.名称 收费名称, B.计算单位, B.加班加价, B.类别, B.屏蔽费别, B.说明, B.执行科室," & vbNewLine & _
            "       B.费用类型 原费用类型, B.是否变价, C.名称 As 类别名称, D.收入项目id As 现收入id, D.原价ID,D.收费细目ID,D.原价,D.现价,D.缺省价格,D.收入项目ID,D.加班加价率,D.附术收费率," & vbNewLine & _
            "       E.名称 As 收入项目, E.收据费目 As 现费目" & vbNewLine & _
            "From " & mstrFreeTable & " A, 收费项目目录 B, 收费项目类别 C, 收费价目 D, 收入项目 E" & vbNewLine & _
            "Where E.ID = D.收入项目id And D.收费细目id = A.收费细目id And A.收费类别 = C.编码 And A.收费细目id = B.ID And" & vbNewLine & _
            "      A.价格父号 Is Null And A.记录性质 = 2 And A.NO = [1] And A.记录状态 = 1 And A.执行状态 <> 1 And" & vbNewLine & _
            "      A.记帐单id = [2] And Sysdate Between D.执行日期 And Nvl(D.终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                    strWherePriceGrade & vbNewLine & _
            "Order By A.序号"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, mlng记帐ID, strPriceGrade)
    
    '没有记录就是空单子
    Call GetEmptyBill(Rows)
    If rsTmp.RecordCount <> 0 Then
        With rsTmp
            i = 1
            Do While Not .EOF
                
                If i = 1 Then
                    '处理单据主体
                    mobjBill.NO = !单据号
                    mobjBill.记录性质 = IIf(IsNull(!记录性质), 0, !记录性质)
                    mobjBill.记录状态 = IIf(IsNull(!记录状态), 0, !记录状态)
                    mobjBill.病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
                    mobjBill.主页ID = IIf(IsNull(!主页ID), 0, !主页ID)
                    mobjBill.病区ID = IIf(IsNull(!病人病区ID), 0, !病人病区ID)
                    mobjBill.科室ID = IIf(IsNull(!病人科室ID), 0, !病人科室ID)
                    mobjBill.姓名 = IIf(IsNull(!姓名), "", !姓名)
                    mobjBill.性别 = IIf(IsNull(!性别), "", !性别)
                    mobjBill.年龄 = IIf(IsNull(!年龄), "", !年龄)
                    mobjBill.标识号 = IIf(IsNull(!标识号), 0, !标识号)
                    mobjBill.床号 = "" & !床号
                    mobjBill.费别 = IIf(IsNull(!费别), "", !费别)
                    mobjBill.门诊标志 = IIf(IsNull(!门诊标志), 0, !门诊标志)
                    mobjBill.加班标志 = IIf(IsNull(!加班标志), 0, !加班标志)
                    mobjBill.婴儿费 = IIf(IsNull(!婴儿费), 0, !婴儿费)
                    mobjBill.开单部门ID = IIf(IsNull(!开单部门ID), 0, !开单部门ID)
                    mobjBill.划价人 = IIf(IsNull(!划价人), "", !划价人)
                    mobjBill.开单人 = IIf(IsNull(!开单人), "", !开单人)
                    mobjBill.操作员编号 = IIf(IsNull(!操作员编号), "", !操作员编号)
                    mobjBill.操作员姓名 = IIf(IsNull(!操作员姓名), "", !操作员姓名)
                    mobjBill.发生时间 = !发生时间
                    mobjBill.登记时间 = !登记时间
                    mobjBill.多病人单 = (IIf(IsNull(!多病人单), 0, !多病人单) = 1)
                End If
                
                '处理收费细目,两个类共同构成了一条收费细目
                Set objBillDetail = New BillDetail
                Set objBillDetail.Detail = New Detail
                
                If !序号 > mlngRows Then
                    MsgBox "本记帐单的收费个数被改小，已经不能完整显示单据内容。", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                objBillDetail.序号 = !序号
                objBillDetail.收费类别 = IIf(IsNull(!收费类别), "", !收费类别)
                objBillDetail.收费细目ID = IIf(IsNull(!收费细目ID), 0, !收费细目ID)
                objBillDetail.收费名称 = IIf(IsNull(!收费名称), "", !收费名称)
                objBillDetail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                objBillDetail.数次 = IIf(IsNull(!数次), 1, !数次)
                objBillDetail.附加标志 = IIf(IsNull(!附加标志), 0, !附加标志)
                objBillDetail.摘要 = IIf(IsNull(!摘要), "", !摘要)
                objBillDetail.执行部门ID = IIf(IsNull(!执行部门ID), 0, !执行部门ID)
                
                If cbo收费类别(!序号 - 1).Tag <> "" And cbo收费类别(!序号 - 1).Tag <> objBillDetail.收费类别 Then
                    MsgBox "定制记帐单第" & !序号 & "行的固定收费类别与单据原有内容不同。", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                If Val(txt收费项目(!序号 - 1).Tag) > 0 And Val(txt收费项目(!序号 - 1).Tag) <> objBillDetail.收费细目ID Then
                    MsgBox "定制记帐单第" & !序号 & "行的固定收费项目与单据原有内容不同。", vbExclamation, gstrSysName
                    mobjBill.NO = ""
                    Exit Sub
                End If
                
                
                objBillDetail.Detail.ID = !收费细目ID
                objBillDetail.Detail.编码 = !编码
                objBillDetail.Detail.变价 = (IIf(IsNull(!是否变价), 0, !是否变价) = 1)
                objBillDetail.Detail.规格 = IIf(IsNull(!规格), "", !规格)
                objBillDetail.Detail.计算单位 = IIf(IsNull(!计算单位), "", !计算单位)
                objBillDetail.Detail.加班加价 = (IIf(IsNull(!加班加价), 0, !加班加价) = 1)
                objBillDetail.Detail.类别 = IIf(IsNull(!类别), "", !类别)
                objBillDetail.Detail.类别名称 = IIf(IsNull(!类别名称), "", !类别名称)
                objBillDetail.Detail.名称 = IIf(IsNull(!收费名称), "", !收费名称)
                objBillDetail.Detail.屏蔽费别 = (IIf(IsNull(!屏蔽费别), 0, !屏蔽费别) = 1)
                objBillDetail.Detail.说明 = IIf(IsNull(!说明), "", !说明)
                objBillDetail.Detail.执行科室 = IIf(IsNull(!执行科室), 0, !执行科室)
                objBillDetail.Detail.类型 = IIf(IsNull(!费用类型), "" & !原费用类型, !费用类型)
                objBillDetail.Detail.要求审批 = 0
                    
                Set objBillDetail.InComes = New BillInComes
                cur单价 = 0: cur实收 = 0: cur应收 = 0
                
                Do
                    '！！按照现有的价格设置重新计算
                    If IIf(IsNull(!是否变价), 0, !是否变价) = 1 Then
                        If Abs(!标准单价) > Abs(IIf(IsNull(!现价), 0, !现价)) Then
                            objBillIncome.标准单价 = IIf(IsNull(!缺省价格), 0, !缺省价格)
                        Else
                            objBillIncome.标准单价 = !标准单价
                        End If
                    Else
                        objBillIncome.标准单价 = !现价
                    End If
                    objBillIncome.收入项目ID = IIf(IsNull(!现收入ID), 0, !现收入ID)
                    objBillIncome.收入项目 = IIf(IsNull(!收入项目), "", !收入项目)
                    objBillIncome.收据费目 = IIf(IsNull(!现费目), "", !现费目)
                    objBillIncome.现价 = IIf(IsNull(!现价), 0, !现价)
                    objBillIncome.原价 = IIf(IsNull(!原价), 0, !原价)
                    
                    '应收金额=单价*付次*数次
                    objBillIncome.应收金额 = objBillIncome.标准单价 * IIf(IsNull(!数次), 1, !数次)
                    
                    '附加手术费率用计算(所有收入项目)
                    If IIf(IsNull(!附加标志), 0, !附加标志) = 1 And IIf(IsNull(!收费类别), "", !收费类别) = "F" Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 * IIf(IsNull(!附术收费率), 1, !附术收费率 / 100)
                    End If
                    
                    '加班费用率计算
                    If IIf(IsNull(!加班标志), 0, !加班标志) = 1 And IIf(IsNull(!加班加价), 0, !加班加价) = 1 Then
                        objBillIncome.应收金额 = objBillIncome.应收金额 + objBillIncome.应收金额 * IIf(IsNull(!加班加价率), 0, !加班加价率 / 100)
                    End If
                    
                    '计算实收金额
                    If IIf(IsNull(!屏蔽费别), 0, !屏蔽费别) = 1 Then
                        objBillIncome.实收金额 = objBillIncome.应收金额
                    Else
                        objBillIncome.实收金额 = ActualMoney(mobjBill.费别, !现收入ID, objBillIncome.应收金额)
                    End If
                    
                    objBillIncome.实际票号 = IIf(IsNull(!实际票号), "", !实际票号)
                    
                    With objBillIncome
                        objBillDetail.InComes.Add .收入项目ID, .收入项目, .收据费目, .标准单价, .应收金额, .实收金额, .原价, .现价, "_" & .实收金额, .实际票号
                        cur单价 = cur单价 + .标准单价
                        cur实收 = cur实收 + .实收金额
                        cur应收 = cur应收 + .应收金额
                    End With
                    
                    
                    '判断下一条记录是否属于当前行
                    blnDo = False
                    int序号 = !序号
                    .MoveNext
                    If Not .EOF Then blnDo = (int序号 = !序号)
                    i = i + 1
                Loop While blnDo And Not .EOF
                
                '完成了一条收费细目的增加
                With objBillDetail
                    mobjBill.Details.Remove "R" & .序号 - 1 '新增前先把以前的空记录删除
                    mobjBill.Details.Add .Detail, .收费细目ID, .收费名称, .序号, .收费类别, .计算单位, .数次, cur单价, cur实收, cur应收, .附加标志, .执行部门ID, .InComes
                End With
            Loop
        End With
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
End Sub

Private Sub ClearDetail(ByVal lngRow As Long)
'功能：刷新显示指定行的内容
'参数：lngRow=指定行，也就是类中的序号-1
'说明：ExpenseBill集合的索引对应单据的行号
    Dim i As Long, j As Long, curMoney As Currency
    Dim objBillDetail As BillDetail
    

    If lngRow > mlngRows - 1 Then Exit Sub
    mobjBill.Details.Remove "R" & lngRow
    mobjBill.Details.AddEmpty lngRow + 1
    
    If cbo收费类别(lngRow).ListIndex <> -1 Then
        mobjBill.Details("R" & lngRow).收费类别 = Chr(cbo收费类别(lngRow).ItemData(cbo收费类别(lngRow).ListIndex))
    End If
    Call ShowDetail(lngRow)
End Sub

Private Function SaveBill() As Boolean
'功能:保存当前输入的记帐单据
'返回:保存是否成功
    Dim i As Integer, j As Integer, arrSQL As Variant
    Dim lngCurID As Long, strNO As String, strTmp As String, strSQL As String
    Dim intInsure As Integer, str消息 As String
    Dim lngNO As Long, lngParent As Long, lngParentNO As Long, lngChildNO As Long

    mobjBill.NO = zlDatabase.GetNextNo(14)
    mobjBill.发生时间 = CDate(txtDate.Text)
    mobjBill.登记时间 = zlDatabase.Currentdate

    gstrModiNO = mobjBill.NO
    arrSQL = Array()
    
    lngChildNO = mlngRows + 1 '价格的父号只能从此开始
    For lngParentNO = 1 To mlngRows
        Set mobjBillDetail = mobjBill.Details("R" & lngParentNO - 1)
        If mobjBillDetail.数次 <> 0 Then
            lngParent = 0
            For Each mobjBillIncome In mobjBillDetail.InComes
                lngParent = lngParent + 1
                If lngParent = 1 Then
                    '第一个收入项目做为主记录
                    lngNO = lngParentNO
                Else
                    lngNO = lngChildNO
                    '子序号要手工递增
                    lngChildNO = lngChildNO + 1
                End If
                
                If mbytUseType = Use门诊 Then
                    '单据主体
                    With mobjBill
                        strSQL = "zl_门诊记帐记录_INSERT('" & .NO & "'," & lngNO & "," & .病人ID & "," & .标识号 & "," & _
                            "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .费别 & "'," & .加班标志 & "," & .婴儿费 & "," & _
                            IIf(.科室ID = 0, .开单部门ID, .科室ID) & "," & .开单部门ID & ",'" & .开单人 & "',"
                    End With
                
                    '收费细目部份
                    With mobjBillDetail
                        strSQL = strSQL & "Null," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "'," & _
                             "1," & .数次 & "," & .附加标志 & "," & .执行部门ID & ","
                    End With
                
                    '收入项目部份
                    With mobjBillIncome
                        strSQL = strSQL & IIf(lngParent = 1, "Null", lngParentNO) & "," & .收入项目ID & "," & _
                            "'" & .收据费目 & "'," & .标准单价 & "," & .应收金额 & "," & .实收金额 & ","
                    End With
                                                
                    '其它部分
                    strSQL = strSQL & _
                        "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "'" & mstrInNO & "',0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',NULL," & mlng记帐ID & ",'" & _
                        mobjBillDetail.摘要 & "',Null,Null,Null,Null,Null,Null)"
                
                Else '住院记帐
                    '单据主体
                    With mobjBill
                        strSQL = "zl_住院记帐记录_INSERT('" & .NO & "'," & lngNO & "," & .病人ID & "," & IIf(.主页ID = 0, "NULL", .主页ID) & "," & _
                            IIf(.标识号 = 0, "NULL", .标识号) & "," & "'" & .姓名 & "','" & .性别 & "','" & .年龄 & "','" & .床号 & "','" & .费别 & "'," & _
                            IIf(.病区ID = 0, .开单部门ID, .病区ID) & "," & IIf(.科室ID = 0, .开单部门ID, .科室ID) & "," & .加班标志 & "," & .婴儿费 & "," & .开单部门ID & ",'" & .开单人 & "',"
                    End With
    
                    '收费细目部份
                    With mobjBillDetail
                        strSQL = strSQL & "Null," & .收费细目ID & ",'" & .收费类别 & "','" & .计算单位 & "',"
                        strSQL = strSQL & IIf(.保险项目否, 1, 0) & "," & IIf(.保险大类ID = 0, "NULL", .保险大类ID) & ",'" & .保险编码 & "',"
                        strSQL = strSQL & "1," & .数次 & "," & .附加标志 & "," & .执行部门ID & ","
                    End With
    
                    '收入项目部份
                    With mobjBillIncome
                        strSQL = strSQL & IIf(lngParent = 1, "Null", lngParentNO) & "," & .收入项目ID & "," & _
                            "'" & .收据费目 & "'," & .标准单价 & "," & .应收金额 & "," & .实收金额 & ","
                        strSQL = strSQL & .统筹金额 & ","
                    End With
    
                    '其它部分
                    strSQL = strSQL & _
                        "To_Date('" & Format(mobjBill.发生时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "To_Date('" & Format(mobjBill.登记时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                        "'" & mstrInNO & "',0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',0,NULL," & mlng记帐ID & ",'" & _
                        mobjBillDetail.摘要 & "',0,Null,Null,Null,Null,Null,Null,0,'" & mobjBillDetail.Detail.类型 & "')"
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = mobjBillDetail.收费细目ID & ";" & strSQL
            Next
        End If
    Next

    '修改前退除原单据
    If mstrInNO <> "" Then
        If mbytUseType <> Use门诊 Then
            '医保记帐作废上传(修改之前已经判断)
            intInsure = BillExistInsure(mstrInNO)
            '去掉了医保连接匹配检查
        End If
    
        If mbytUseType = Use门诊 Then
            strSQL = "zl_门诊记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSQL = "zl_住院记帐记录_DELETE('" & mstrInNO & "',NULL,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & strSQL
        End If
    End If

    If UBound(arrSQL) >= 0 Then
        '对SQL序列按收费细目ID排序
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If CLng(Mid(arrSQL(j), 1, InStr(arrSQL(j), ";") - 1)) < CLng(Mid(arrSQL(i), 1, InStr(arrSQL(i), ";") - 1)) Then
                    strTmp = CStr(arrSQL(j))
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next

        '执行SQL语句
        On Error GoTo errH
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ";") + 1), Me.Caption)
        Next
        
        '医保接口
        If mbytUseType <> Use门诊 Then
            '1.医保记帐作废上传
            If mstrInNO <> "" And intInsure <> 0 Then
                If MCPAR.记帐作废上传 And Not MCPAR.记帐完成后上传 Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        gcnOracle.RollbackTrans: Exit Function
                    End If
                End If
            End If
            
            '2.记帐实时上传
            If Not IsNull(mrsInfo!险类) Then
                If MCPAR.记帐上传 And Not MCPAR.记帐完成后上传 Then
                    str消息 = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , mrsInfo!险类) Then
                        gcnOracle.RollbackTrans
                        If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
        
        gcnOracle.CommitTrans

        '医保接口
        If mbytUseType <> Use门诊 Then
            '1.医保记帐作废上传
            If mstrInNO <> "" And intInsure <> 0 Then
                If MCPAR.记帐作废上传 And MCPAR.记帐完成后上传 Then
                    If Not gclsInsure.TranChargeDetail(2, mstrInNO, 2, 2, "", , intInsure) Then
                        MsgBox "单据""" & mstrInNO & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                    End If
                End If
            End If
            
            '2.记帐实时上传
            If Not IsNull(mrsInfo!险类) Then
                '医保传输费用明细
                If MCPAR.记帐上传 And MCPAR.记帐完成后上传 Then
                    str消息 = ""
                    If Not gclsInsure.TranChargeDetail(2, mobjBill.NO, 2, 1, str消息, , mrsInfo!险类) Then
                        If str消息 <> "" Then
                            MsgBox str消息, vbInformation, gstrSysName
                        Else
                            MsgBox "单据""" & mobjBill.NO & """的数据向医保传送失败，该单据已保存。", vbInformation, gstrSysName
                        End If
                    End If
                End If
            End If
        End If

        '加入单据历史记录(所有类型单据)
        For i = 0 To cboNO.ListCount - 1
            strNO = strNO & "," & cboNO.List(i)
        Next
        strNO = mobjBill.NO & strNO
        cboNO.Clear
        For i = 0 To UBound(Split(strNO, ","))
            cboNO.AddItem Split(strNO, ",")(i)
            If i = 9 Then Exit For '只显示10个
        Next
        
        '医保接口
        If str消息 <> "" Then MsgBox str消息, vbInformation, gstrSysName
    End If
    SaveBill = True
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Private Function ReadBill(strNO As String) As Integer
'功能：根据单据号读取一张单据并将其填入表格
'参数：strFullNo=单据号
    Dim rsTmp As ADODB.Recordset, rsPatiMoney As ADODB.Recordset
    Dim i As Integer, blnDeal As Boolean, strSQL As String
    Dim curTotal As Currency, cur应收Total As Currency
    Dim strFullNO As String, strTmp As String
    Dim bln住院 As Boolean
    
    On Error GoTo errH
    bln住院 = (mstrFreeTable <> "门诊费用记录")
    
    strFullNO = GetFullNO(strNO, 14)
    If mbytInState = sta销帐 Then
        '判断该张记帐单能否销帐，如果有一条记录执行就不允许
        strSQL = "select nvl(count(执行状态),0) as  总数,nvl(sum(decode(执行状态,1,1,0)),0) as 执行 " & _
                " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & _
                " Where  " & IIf(bln住院, " Nvl(多病人单,0)=0 And ", "") & " 记录状态=1 " & _
                "       And 记录性质=2 and 记帐单ID=[2] And NO=[1]"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng记帐ID)
        
        If rsTmp("总数") = 0 Then
            MsgBox "没有发现该单据,该单据可能已经作废！", vbInformation, gstrSysName
            Exit Function
        End If
        If rsTmp("执行") > 0 Then
            MsgBox "不能对该单据销帐,它已经执行了！", vbInformation, gstrSysName
            Exit Function
        End If
        rsTmp.Close
    End If
    
    '读取单据主体   '" & IIf(mblnNOMoved, iif(blnNOOnline
    
    strSQL = _
    " Select A.病人ID," & IIf(bln住院, " Nvl(A.主页ID,0)  as  主页ID,A.病人病区ID,B.名称 as 病人病区,A.床号,", "0  as  主页ID, 0 as 病人病区ID,'' as 病人病区, '' as 床号,") & _
    "       A.标识号,A.姓名,A.性别,A.年龄,A.费别," & _
    "       A.病人科室ID,C.名称 as 病人科室,A.开单部门ID," & _
    "       Nvl(A.加班标志,0) as 加班标志,Nvl(A.婴儿费,0) as 婴儿费," & _
    "       A.开单人,A.划价人,A.操作员姓名,A.发生时间,A.结帐ID" & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & IIf(bln住院, ",部门表 B", "") & ",部门表 C,人员表 D" & _
    " Where Rownum=1  " & _
            IIf(bln住院, " And Nvl(A.多病人单,0)=0 And  A.病人病区ID=B.ID(+) ", "") & _
    "       And A.记帐单ID=[2] And A.记录状态" & IIf(mblnViewCancel, "=2", " IN(1,3)") & _
    "       And A.记录性质=2 And A.NO=[1]  and A.病人科室ID=C.ID(+) " & _
    "       And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
    "       And Nvl(A.操作员姓名,A.划价人)=D.姓名"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng记帐ID)
    
    If rsTmp.EOF Then
        MsgBox "没有发现该单据,该单据可能已经作废！", vbInformation, gstrSysName
        Exit Function
    ElseIf mbytUseType <> Use门诊 Then
        '门诊记帐不需要对科室权限进行判断
        If mbytUseType = 0 Or mbytUseType = 1 Then
            If InStr(mstrPrivs, "所有病区") = 0 And mlngUnitID > 0 Then
                If InStr(1, "," & mstrUnitIDs & ",", "," & IIf(IsNull(rsTmp!病人病区ID), 0, rsTmp!病人病区ID) & ",") = 0 Then
                    MsgBox "你没有权限读取其它病区的单据！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        ElseIf mbytUseType = 2 Then
            If InStr(mstrPrivs, "所有科室") = 0 And mlngDeptID > 0 Then
                If IIf(IsNull(rsTmp!开单部门ID), 0, rsTmp!开单部门ID) <> mlngDeptID Then
                    MsgBox "你没有权限读取其它科室开单的单据！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If

    '单据头
    cboNO.Text = strFullNO                                   '单据号
    txtPatient.Text = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名) '姓名
    txt年龄.Text = IIf(IsNull(rsTmp!年龄), "", rsTmp!年龄) '年龄
    txt床号.Text = IIf(IsNull(rsTmp!床号), "", rsTmp!床号) '年龄
    txt标识号.Text = IIf(IsNull(rsTmp!标识号), "", rsTmp!标识号) '标识号
    txt病人ID.Text = IIf(IsNull(rsTmp!病人ID), "", rsTmp!病人ID) '病人ID
    txt主页ID.Text = IIf(IsNull(rsTmp!主页ID), "", rsTmp!主页ID) '主页ID
    txt病人病区.Text = IIf(IsNull(rsTmp!病人病区), "", rsTmp!病人病区) '病人病区
    txt病人科室.Text = IIf(IsNull(rsTmp!病人科室), "", rsTmp!病人科室) '病人科室

    '性别
    Call cbo.SeekIndex(cbo性别, IIf(IsNull(rsTmp!性别), "", rsTmp!性别), , True)
    If cbo性别.ListIndex = -1 And Not IsNull(rsTmp!性别) Then
        cbo性别.AddItem rsTmp!性别, 0
        cbo性别.ListIndex = 0
    End If
    
    '费别
    Call cbo.SeekIndex(cbo费别, IIf(IsNull(rsTmp!费别), "", rsTmp!费别), , True)
    If cbo费别.ListIndex = -1 And Not IsNull(rsTmp!费别) Then
        cbo费别.AddItem rsTmp!费别, 0
        cbo费别.ListIndex = 0
    End If
    
    txtDate.Text = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm:ss")
    chk加班.Value = IIf(IsNull(rsTmp!加班标志), 0, rsTmp!加班标志)
    Call LoadPatientBaby(cboBaby, rsTmp!病人ID, rsTmp!主页ID)
    Call zlControl.CboLocate(cboBaby, rsTmp!婴儿费, True)
    
    mblnDo = False
    
        '科室确定医生
        cbo开单科室.ListIndex = cbo.FindIndex(cbo开单科室, NVL(rsTmp!开单部门ID, 0))
        If cbo开单科室.ListIndex = -1 And Not IsNull(rsTmp!开单部门ID) Then
            cbo开单科室.AddItem GET部门名称(rsTmp!开单部门ID, mrs开单科室), 0
            cbo开单科室.ItemData(cbo开单科室.NewIndex) = rsTmp!开单部门ID
            cbo开单科室.ListIndex = cbo开单科室.NewIndex
        End If
        
        cbo开单人.Clear
        If cbo开单科室.ListIndex <> -1 Then
            Call FillDoctor(cbo开单科室.ItemData(cbo开单科室.ListIndex))
        End If
        Call cbo.SeekIndex(cbo开单人, NVL(rsTmp!开单人), , True)
        If cbo开单人.ListIndex = -1 And Not IsNull(rsTmp!开单人) Then
            cbo开单人.AddItem rsTmp!开单人, 0
            cbo开单人.ListIndex = cbo开单人.NewIndex
        End If
    
    mblnDo = True
    
    
    '病人费用信息
    If Not IsNull(rsTmp!病人ID) Then
        Set rsPatiMoney = GetMoneyInfo(rsTmp!病人ID)
        If Not rsPatiMoney Is Nothing Then
            sta.Panels(3).Text = "预交:" & Format(rsPatiMoney!预交余额, "0.00") & _
            "/费用:" & Format(rsPatiMoney!费用余额, gstrDec) & _
            "/剩余:" & Format(rsPatiMoney!预交余额 - rsPatiMoney!费用余额, "0.00")
        End If
    End If
    
    '读取单据收费细目
    strSQL = _
    " Select Decode(A.价格父号,NULL,A.序号,A.价格父号) as 序号," & _
    "       C.编码,C.名称 as 类别,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型,A.计算单位," & _
    "       Avg(A.数次) as 数次,Sum(A.标准单价) as 单价,Sum(A.应收金额) as 应收金额, " & _
    "       Sum(A.实收金额) as 实收金额,A.附加标志,A.执行部门ID,D.名称 as 执行部门 " & _
    " From " & IIf(mblnNOMoved, zlGetFullFieldsTable(mstrFreeTable), mstrFreeTable & " A") & ",收费项目目录 B,收费项目类别 C,部门表 D " & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.执行部门ID=D.ID " & _
    "       And A.记录状态" & IIf(mblnViewCancel, "=2", " IN(1,3)") & " And A.NO=[1]" & _
    "       " & IIf(mstrFreeTable = "住院费用记录", "And Nvl(A.多病人单,0)=0 ", "") & " And A.记录性质=2 And A.记帐单ID=[2]" & _
    " Group by Decode(A.价格父号,NULL,A.序号,A.价格父号),C.编码,C.名称," & _
    "       B.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,A.附加标志,A.执行部门ID,D.名称" & _
    " Order by Decode(A.价格父号,NULL,A.序号,A.价格父号)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFullNO, mlng记帐ID)
    If rsTmp.EOF Then Exit Function
    
    '单据体
    curTotal = 0
    cur应收Total = 0
    
    For i = 0 To mlngRows - 1
        blnDeal = False
        If Not rsTmp.EOF Then
            If rsTmp("序号") = i + 1 Then
                
                cbo收费类别(i).AddItem rsTmp!类别
                cbo收费类别(i).ListIndex = cbo收费类别(i).NewIndex
                txt收费项目(i).Text = rsTmp!名称
                txt计算单位(i).Text = IIf(IsNull(rsTmp!计算单位), "", rsTmp!计算单位)
                txt数次(i).Text = rsTmp!数次
                txt标准单价(i).Text = Format(rsTmp!单价, "0.0000")
                If Val(txt标准单价(i).Text) = 0 Then txt标准单价(i).Text = ""
                txt应收金额(i).Text = Format(rsTmp!应收金额, gstrDec)
                If Val(txt应收金额(i).Text) = 0 Then txt应收金额(i).Text = ""
                txt实收金额(i).Text = Format(rsTmp!实收金额, gstrDec)
                If Val(txt实收金额(i).Text) = 0 Then txt实收金额(i).Text = ""
                
                'byZT200302
                strTmp = rsTmp("执行部门")
                If strTmp <> "" Then
                    Call cbo.SeekIndex(cbo执行科室(i), strTmp, , True)
                    If cbo执行科室(i).ListIndex = -1 Then
                        cbo执行科室(i).AddItem rsTmp("执行部门")
                        cbo执行科室(i).ListIndex = cbo执行科室(i).NewIndex
                    End If
                End If
                'cbo执行科室(i).Text = rsTmp("执行部门")
                
                chk附加(i).Value = rsTmp!附加标志
                
                curTotal = curTotal + rsTmp("实收金额")
                cur应收Total = cur应收Total + rsTmp("应收金额")
                rsTmp.MoveNext
                blnDeal = True
            End If
        End If
        '没找到合适的值
        If blnDeal = False And Val(txt收费项目(i).Tag) <= 0 Then
            cbo收费类别(i).ListIndex = cbo.FindIndex(cbo收费类别(i), Val(cbo收费类别(i).Tag))
            txt收费项目(i).Text = ""
            txt计算单位(i).Text = ""
            txt数次(i).Text = ""
            txt标准单价(i).Text = ""
            txt应收金额(i).Text = ""
            txt实收金额(i).Text = ""
            cbo执行科室(i).ListIndex = -1
            chk附加(i).Value = 0
        End If
    Next
    If rsTmp.EOF = False Then
        MsgBox "本记帐单的收费数被改小，已经不能完整显示单据内容。", vbExclamation, gstrSysName
        Exit Function
    End If
    

    txt实收.Text = Format(curTotal, gstrDec)
    txt应收.Text = Format(cur应收Total, gstrDec)

    ReadBill = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetDetailNum(ByVal lngRow As Long, Optional dbl结帐数量 As Double) As Single
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人指定细目的总记帐数据(含本单据中)
    '入参:lngRow=当前单据行
    '出参:dbl结帐数量-返回当前已经结帐的数量
    '返回:
    '编制:刘兴洪
    '日期:2010-08-19 18:02:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim lngNum As Long, i As Long
    Dim strSQL As String

    If mrsInfo.State = 1 Then
        '当前单据中的数量
        For i = 0 To mlngRows - 1
            If i <> lngRow And mobjBill.Details("R" & i).收费细目ID = mobjBill.Details("R" & lngRow).收费细目ID Then
                lngNum = lngNum + mobjBill.Details("R" & i).数次
            End If
        Next
        dbl结帐数量 = 0
        '数据库中的数量
        strSQL = _
        " Select Sum(Nvl(付数,1)*数次) as NUM," & _
        "           Sum(decode(结帐ID,NULL,0,1)* Nvl(付数,1)*数次) as 结帐数量  " & _
        " From " & mstrFreeTable & _
        " Where 价格父号 is Null And 记帐费用=1 And 记录状态<>0" & _
        "       And 病人ID=[1] " & IIf(mstrFreeTable = "门诊费用记录", "", " And Nvl(主页ID,0)=[2]") & " And 收费细目ID+0=[3]"
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mrsInfo!病人ID), Val("" & mrsInfo!主页ID), mobjBill.Details("R" & lngRow).收费细目ID)
        If Not rsTmp.EOF Then
            lngNum = lngNum + NVL(rsTmp!Num, 0)
            dbl结帐数量 = Val(NVL(rsTmp!结帐数量))
        End If
        GetDetailNum = lngNum
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FillDoctor(Optional lng科室ID As Long, Optional strMask As String)
'功能：读取并填写医生列表
'参数：strMask=输入匹配的简码或编码
    Dim i As Integer, lngOldID As Long
    Dim str病人医生 As String, strPre As String
    
    If cbo开单人.ListIndex <> -1 Then strPre = cbo开单人.List(cbo开单人.ListIndex)
    cbo开单人.Clear
    
    Call GetDoctor(lng科室ID, gbln护士 And (gstr收费类别 = "" _
        Or gstr收费类别 Like "*'E'*" Or gstr收费类别 Like "*'M'*" Or gstr收费类别 Like "*'4'*"), mrs开单人, IIf(mbytUseType = Use门诊, 1, 2))

    If Not mrs开单人 Is Nothing Then
        If mrsInfo.State = 1 And mbytUseType <> Use门诊 Then If Not IsNull(mrsInfo!住院医师) Then str病人医生 = mrsInfo!住院医师
        
        i = IIf(mrs开单人.RecordCount = 0, 0, mrs开单人.RecordCount - 1)
        ReDim marrDr(i)
        
        For i = 1 To mrs开单人.RecordCount
            If lngOldID <> mrs开单人!ID Then
                If strMask = "" Then
                    cbo开单人.AddItem IIf(IsNull(mrs开单人!简码), "", mrs开单人!简码 & "-") & mrs开单人!姓名
                    cbo开单人.ItemData(cbo开单人.NewIndex) = Val(mrs开单人!编号)
                Else
                    If InStr(IIf(IsNull(mrs开单人!简码), "", mrs开单人!简码 & "-") & mrs开单人!姓名, UCase(strMask)) Then
                        cbo开单人.AddItem IIf(IsNull(mrs开单人!简码), "", mrs开单人!简码 & "-") & mrs开单人!姓名
                        cbo开单人.ItemData(cbo开单人.NewIndex) = Val(mrs开单人!编号)
                    ElseIf IsNumeric(strMask) Then
                        If CDbl(strMask) = CDbl(mrs开单人!编号) Then
                            cbo开单人.AddItem IIf(IsNull(mrs开单人!简码), "", mrs开单人!简码 & "-") & mrs开单人!姓名
                            cbo开单人.ItemData(cbo开单人.NewIndex) = Val(mrs开单人!编号)
                        End If
                    End If
                End If
                
                marrDr(cbo开单人.NewIndex) = mrs开单人!ID & "|" & mrs开单人!部门ID & "|" & IIf(IsNull(mrs开单人!编号), "", mrs开单人!编号) & "|" & mrs开单人!姓名 & "|" & IIf(IsNull(mrs开单人!简码), "", mrs开单人!简码) & "|" & mrs开单人!职务 & "|" & mrs开单人!人员性质
                
                If cbo开单人.List(cbo开单人.NewIndex) = strPre And cbo开单人.ListIndex = -1 Then
                    cbo开单人.ListIndex = cbo开单人.NewIndex
                End If
                If str病人医生 = mrs开单人!姓名 And cbo开单人.ListIndex = -1 Then
                    cbo开单人.ListIndex = cbo开单人.NewIndex
                End If
                lngOldID = mrs开单人!ID
            End If
            mrs开单人.MoveNext
        Next
        
        If cbo开单人.ListCount > 0 Then ReDim Preserve marrDr(cbo开单人.ListCount - 1)
        If cbo开单人.ListCount = 1 And cbo开单人.ListIndex = -1 Then cbo开单人.ListIndex = 0
    End If
End Sub
Private Function FillDept(Optional lng人员ID As Long) As Long
'功能：读取并显示科室
'参数：lng人员ID=只读取指定人员所在科室(包含非缺省的)
'返回：科室个数
    
    Dim strSQL As String, i As Long
    Dim lngDeptID As Long, lngOldDepID As Long
    Dim strDepts As String  '指定人员所属的多个部门
    
    On Local Error GoTo errH
            
    '记录原科室,用于重新定位
    If cbo开单科室.ListIndex <> -1 Then
        lngDeptID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    End If
    cbo开单科室.Clear
    
    If mrs开单科室 Is Nothing Then  '一定要在Form_Unload中设置nothing
    
         '可选开单科室(如果是医技科室,则包含门诊和住院的)
        If (InStr(mstrPrivsOpt, "门诊留观记帐") > 0 And gbln门诊留观) Or mbytUseType = 2 Then
            strSQL = "1,2,3"
        Else
            strSQL = "2,3"
        End If
        If mbytUseType = Use住院 Or mbytUseType = Use科室分散 Then
            strSQL = _
                "Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质 " & _
                " from 部门表 A,部门性质说明 B " & _
                " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                " and B.部门ID=A.ID and (B.服务对象 IN(" & strSQL & ") AND B.工作性质 IN('临床','手术')  or B.工作性质='产科')" & _
                " Order by A.编码"
        ElseIf mbytUseType = Use医技科室 Then
            '医技科室记帐
            If InStr(mstrPrivs, "所有科室") > 0 Then
                strSQL = _
                    "Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质 " & _
                    " from 部门表 A,部门性质说明 B " & _
                    " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                    " and B.部门ID=A.ID and (B.服务对象 IN(" & strSQL & ") AND B.工作性质 IN('检查','检验','手术','治疗','营养') Or b.工作性质='产科')" & _
                    " Order by A.编码"
            Else
                strSQL = _
                    "Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质 " & _
                    " from 部门表 A,部门性质说明 B " & _
                    " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                    " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                    " and B.部门ID=A.ID and (B.服务对象 IN(" & strSQL & ") AND B.工作性质 IN('检查','检验','手术','治疗','营养') Or b.工作性质='产科')" & _
                    " And A.ID=" & mlngDeptID & _
                    " Order by A.编码"
            End If
        ElseIf mbytUseType = Use门诊 Then
            strSQL = _
                " Select Distinct A.ID,A.编码,A.名称,A.简码,B.工作性质 " & _
                " from 部门表 A,部门性质说明 B " & _
                " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                " and B.部门ID=A.ID and (B.服务对象 IN(1,3) AND B.工作性质 IN('临床','手术') Or b.工作性质='产科')" & _
                " Order by A.编码"
        End If
        Set mrs开单科室 = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(mrs开单科室, strSQL, Me.Caption)
    End If
   
    If lng人员ID <> 0 Then
        If mrs开单人 Is Nothing Then Call FillDoctor
        mrs开单人.Filter = "ID=" & lng人员ID
        For i = 1 To mrs开单人.RecordCount
            strDepts = strDepts & " OR ID=" & mrs开单人!部门ID      'filter不支持in
            mrs开单人.MoveNext
        Next
        If strDepts <> "" Then
            mrs开单科室.Filter = Mid(strDepts, 4)
        Else
            mrs开单科室.Filter = "ID=0" '人员没有设置部门,不显示开单科室
        End If
    Else
        mrs开单科室.Filter = ""
    End If
    
    If Not mrs开单科室.EOF Then
        For i = 1 To mrs开单科室.RecordCount
            If lngOldDepID <> mrs开单科室!ID Then   '一个部门可能同时属于手术和临床,不加载相同的
                cbo开单科室.AddItem IIf(zlIsShowDeptCode, mrs开单科室!编码 & "-", "") & mrs开单科室!名称
                cbo开单科室.ItemData(cbo开单科室.ListCount - 1) = mrs开单科室!ID
                
                If mrs开单科室!ID = mlngDeptID Then cbo开单科室.ListIndex = cbo开单科室.NewIndex
                
                If mrs开单科室!ID = lngDeptID And cbo开单科室.ListIndex = -1 Then
                    cbo开单科室.ListIndex = cbo开单科室.NewIndex
                End If
                
                lngOldDepID = mrs开单科室!ID
            End If
            mrs开单科室.MoveNext
        Next
        If cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
    End If
    
    FillDept = mrs开单科室.RecordCount
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CalcGridToTal(Optional bln应收 As Boolean) As Currency
    Dim objTmpDetail As BillDetail
    Dim objTmpIncome As BillInCome
    Dim i As Integer, intCol As Integer
    If mobjBill.Details.Count > 0 Then
        For Each objTmpDetail In mobjBill.Details
            For Each objTmpIncome In objTmpDetail.InComes
                If bln应收 Then
                    CalcGridToTal = CalcGridToTal + objTmpIncome.应收金额
                Else
                    CalcGridToTal = CalcGridToTal + objTmpIncome.实收金额
                End If
            Next
        Next
    Else
        For i = 0 To mlngRows - 1
            CalcGridToTal = CalcGridToTal + Val(IIf(bln应收, txt应收金额(i).Text, txt实收金额(i).Text))
        Next
    End If
End Function

Private Function CheckBillisZero() As Boolean
'功能：判断单据所有的行是否数量都为0
    Dim i As Integer, j As Integer
    
    For i = 0 To mlngRows - 1
        If mobjBill.Details("R" & i).数次 = 0 Then j = j + 1
    Next
    
    CheckBillisZero = (mlngRows = j)
End Function

Private Function SaveModi() As Boolean
'功能：保存当前修改的费用单据
    Dim strSQL As String

    strSQL = "zl_病人费用记录_Update('" & cboNO.Text & "',2,'" & zlStr.NeedName(cbo开单人.Text) & "'," & _
        "To_Date('" & txtDate.Text & "','YYYY-MM-DD HH24:MI:SS'),NULL," & IIf(mbytUseType = Use门诊, 1, 2) & " )"
    On Error GoTo errH
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveModi = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Private Function InitFace() As Boolean
'功能：完成对单据界面的初始化
    Dim arrHead() As String, i As Long
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim sngTemp As Single, arrBaby As Variant
    Dim ctlTemp As Control, varTemp As Variant
    Dim lngIndex As Long
    Dim blnContainer As Boolean       '该控件是否放在一个单独的容器中
    
    On Error GoTo errHandle
    
    InitFace = False
    If mbytUseType = Use门诊 Then
        mstrFreeTable = "门诊费用记录"
    Else
        mstrFreeTable = "住院费用记录"
    End If
    '一、得到单据头
    strSQL = "select 名称,收费项目数,适用范围,宽度,高度,背景色 from 收费记帐单 where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记帐ID)
    '判断可用性
    If rsTmp.EOF Then
        If mstrInNO <> "" Then
            strSQL = "zl_收费记帐单_Normalize('" & mstrInNO & "')"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            MsgBox "你所选择的记帐单属于自定义记帐单，它已经被删除。" & vbCrLf & _
                "该单据已改为由普通记帐单处理，请重新刷新列表。", vbExclamation, gstrSysName
        Else
            MsgBox "你所选择的记帐单属于自定义记帐单，它已经被删除。" & vbCrLf & _
                "请重新进入本程序。", vbExclamation, gstrSysName
        End If
        Exit Function
    End If
    If mbytUseType = Use门诊 And Mid(rsTmp("适用范围"), 1, 1) <> "1" Then
        MsgBox "本记帐单不再支持门诊记帐，请重新刷新列表。", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = Use住院 And Mid(rsTmp("适用范围"), 2, 1) <> "1" Then
        MsgBox "本记帐单不再支持住院记帐，请重新刷新列表。", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = Use科室分散 And Mid(rsTmp("适用范围"), 3, 1) <> "1" Then
        MsgBox "本记帐单不再支持科室分散记帐，请重新刷新列表。", vbExclamation, gstrSysName
        Exit Function
    End If
    If mbytUseType = Use医技科室 And Mid(rsTmp("适用范围"), 4, 1) <> "1" Then
        MsgBox "本记帐单不再支持医技科室记帐，请重新刷新列表。", vbExclamation, gstrSysName
        Exit Function
    End If
    '改变窗口的大小
    sngTemp = Me.Width - Me.ScaleWidth   '得到窗口大小与客户区大小的差值
    Me.Width = rsTmp("宽度") + sngTemp
    sngTemp = Me.Height - Me.ScaleHeight '得到窗口大小与客户区大小的差值
    Me.Height = rsTmp("高度") + sta.Height + sngTemp
    fraForm.Left = 0: fraForm.Top = 0
    fraForm.Width = rsTmp("宽度"): fraForm.Height = rsTmp("高度")
    fraForm.BackColor = rsTmp("背景色")
    
    Me.Caption = "记帐处理" & " - " & rsTmp("名称")
    '得到控件数组
    mlngRows = rsTmp("收费项目数")
    For i = 1 To mlngRows - 1
        Load cbo收费类别(i): Set cbo收费类别(i).Container = fraForm
        Load txt收费项目(i): Set txt收费项目(i).Container = fraForm
        Load cmd细目选择(i): Set cmd细目选择(i).Container = fraForm
        Load txt计算单位(i): Set txt计算单位(i).Container = fraForm
        Load txt数次(i):     Set txt数次(i).Container = fraForm
        Load txt标准单价(i): Set txt标准单价(i).Container = fraForm
        Load txt应收金额(i): Set txt应收金额(i).Container = fraForm
        Load txt实收金额(i): Set txt实收金额(i).Container = fraForm
        Load cbo执行科室(i): Set cbo执行科室(i).Container = fraForm
        Load chk附加(i):     Set chk附加(i).Container = fraForm
    Next
    rsTmp.Close
    '二、得到单据体
    strSQL = "select 对应字段,序号,类型,定义值,顺序号,左边,顶边,宽度,高度,字体,前景色,背景色,是否显示,外形,边框线,透明" & _
        " from 收费记帐单定义 where 记帐ID=[1] order by 顺序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng记帐ID)
    
    Do Until rsTmp.EOF
        blnContainer = False
        lngIndex = IIf(IsNull(rsTmp("序号")), 0, rsTmp("序号")) - 1
        Select Case rsTmp("类型")
            Case "CheckBox"
                Select Case rsTmp("对应字段")
                    Case "附加标志"
                        Set ctlTemp = chk附加(lngIndex)
                    Case "加班标志"
                        Set ctlTemp = chk加班
                    Case "销"
                        Set ctlTemp = chk销
                        blnContainer = True
                End Select
                ctlTemp.Caption = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                ctlTemp.Height = rsTmp("高度")
                ctlTemp.ForeColor = rsTmp("前景色")
                ctlTemp.BackColor = rsTmp("背景色")
                ctlTemp.Appearance = rsTmp("外形")
            Case "ComboBox"
                Select Case rsTmp("对应字段")
                    Case "费别"
                        Set ctlTemp = cbo费别
                    Case "性别"
                        Set ctlTemp = cbo性别
                    Case "NO"
                        Set ctlTemp = cboNO
                        blnContainer = True
                    Case "开单人"
                        Set ctlTemp = cbo开单人
                        blnContainer = True
                    Case "开单部门"
                        Set ctlTemp = cbo开单科室
                        cbo开单科室.Tag = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                        If cbo开单科室.Tag <> "" Then
                            cbo开单科室.Locked = True
                            cbo开单科室.TabStop = True
                        End If
                    Case "收费类别"
                        Set ctlTemp = cbo收费类别(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                        If ctlTemp.Tag = "0" Then ctlTemp.Tag = ""
                        ctlTemp.Locked = ctlTemp.Tag <> ""
                    Case "执行部门"
                        Set ctlTemp = cbo执行科室(lngIndex)
                    Case "婴儿费"
                        Set ctlTemp = cboBaby
                            '婴儿费
                        Call LoadPatientBaby(ctlTemp, 0, 0)
                End Select
                ctlTemp.ForeColor = rsTmp("前景色")
                ctlTemp.BackColor = rsTmp("背景色")
            Case "CommandButton"
                Select Case rsTmp("对应字段")
                    Case "取消"
                        Set ctlTemp = cmdCancel
                        blnContainer = True
                    Case "确定"
                        Set ctlTemp = cmdOK
                        blnContainer = True
                    Case "细目选择"
                        Set ctlTemp = cmd细目选择(lngIndex)
                End Select
                ctlTemp.Caption = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                ctlTemp.Height = rsTmp("高度")
            Case "Label"
                Load lbl(lbl.UBound + 1)
                Set ctlTemp = lbl(lbl.UBound)
                ctlTemp.Caption = Replace(IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值")), "[单位名称]", gstr单位名称)
                ctlTemp.Appearance = rsTmp("外形")
                ctlTemp.BorderStyle = rsTmp("边框线")
                ctlTemp.BackStyle = rsTmp("透明")
                ctlTemp.ForeColor = rsTmp("前景色")
                ctlTemp.BackColor = rsTmp("背景色")
                ctlTemp.Height = rsTmp("高度")
            Case "TextBox"
                Select Case rsTmp("对应字段")
                    Case "姓名"
                        Set ctlTemp = txtPatient
                    Case "标识号"
                        Set ctlTemp = txt标识号
                    Case "病人ID"
                        Set ctlTemp = txt病人ID
                    Case "年龄"
                        Set ctlTemp = txt年龄
                    Case "床号"
                        Set ctlTemp = txt床号
                    Case "病人病区"
                        Set ctlTemp = txt病人病区
                    Case "病人科室"
                        Set ctlTemp = txt病人科室
                    Case "入院次数"
                        Set ctlTemp = txt主页ID
                    Case "实收合计"
                        Set ctlTemp = txt实收
                    Case "应收合计"
                        Set ctlTemp = txt应收
                    Case "收费细目"
                        Set ctlTemp = txt收费项目(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                        
                        If Val(ctlTemp.Tag) > 0 Then
                            '有明确的值
                            cbo收费类别(lngIndex).Locked = True
                            txt收费项目(lngIndex).Locked = True
                            txt收费项目(lngIndex).TabStop = False
                            cmd细目选择(lngIndex).Enabled = False
                        End If
                    Case "计算单位"
                        Set ctlTemp = txt计算单位(lngIndex)
                    Case "数次"
                        Set ctlTemp = txt数次(lngIndex)
                        ctlTemp.Tag = IIf(IsNull(rsTmp("定义值")), "", rsTmp("定义值"))
                        
                        If Val(ctlTemp.Tag) > 0 Then
                            '有明确的值
                            ctlTemp.Locked = True
                            ctlTemp.TabStop = False
                        End If
                    Case "标准单价"
                        Set ctlTemp = txt标准单价(lngIndex)
                    Case "实收金额"
                        Set ctlTemp = txt实收金额(lngIndex)
                    Case "应收金额"
                        Set ctlTemp = txt应收金额(lngIndex)
                    Case "发生时间"
                        Set ctlTemp = txtDate
                        blnContainer = True
                End Select
                ctlTemp.Height = rsTmp("高度")
                ctlTemp.ForeColor = CLng(rsTmp("前景色"))
                ctlTemp.BackColor = CLng(rsTmp("背景色"))
                ctlTemp.Appearance = rsTmp("外形")
                ctlTemp.BorderStyle = rsTmp("边框线")
        End Select
        If blnContainer = True Then
            ctlTemp.Left = 0
            ctlTemp.Top = 0
            ctlTemp.Container.Left = rsTmp("左边")
            ctlTemp.Container.Top = rsTmp("顶边")
            ctlTemp.Container.Width = rsTmp("宽度")
            ctlTemp.Container.Height = rsTmp("高度")
        Else
            ctlTemp.Left = rsTmp("左边")
            ctlTemp.Top = rsTmp("顶边")
        End If
        
        ctlTemp.Width = rsTmp("宽度")
        varTemp = Split(rsTmp("字体"), "|")
        ctlTemp.Font.Name = varTemp(0)
        ctlTemp.Font.Size = varTemp(1)
        ctlTemp.Font.Bold = varTemp(2) = "1"
        ctlTemp.Font.Italic = varTemp(3) = "1"
        ctlTemp.Font.Underline = varTemp(4) = "1"
        ctlTemp.Visible = rsTmp("是否显示") = 1
        ctlTemp.TabIndex = rsTmp("顺序号")
        rsTmp.MoveNext
    Loop
    

    '三、根据表单要完成的功能设置界面布局
    cboBaby.Enabled = mbytInState = sta执行
    Select Case mbytInState
        Case sta执行  '执行
            If mbytUseType <> Use门诊 And (InStr(mstrPrivsOpt, "住院销帐") = 0 Or mstrInNO <> "") Then
                fra销.Visible = False
                chk销.Visible = False
            End If
        Case sta查阅 '查阅
            fraNO.Enabled = False
            fra开单人.Enabled = False
            fra时间.Enabled = False
            fraForm.Enabled = False
            '为了使控件不可见
            fraOK.Visible = False
            
            If mblnViewCancel = False Then
                fra销.Visible = False
            Else
                fra销.Enabled = False
                chk销.ForeColor = &HFF&
            End If
            cmdCancel.Caption = "退出(&X)"
        Case sta调整 '调整
            fra销.Visible = False
            fraForm.Enabled = False
            fraNO.Enabled = False
        Case sta销帐 '销帐
            fra销.Visible = False
            fraForm.Enabled = False
            fra开单人.Enabled = False
            fra时间.Enabled = False
            fraNO.Enabled = False
    End Select
    
    '读取简码匹配方式
    sta.Panels("MedicareType").Visible = mbytInState = 0
    sta.Panels("PY").Visible = mbytInState = 0 And gbln简码切换 '35242
    sta.Panels("WB").Visible = mbytInState = 0 And gbln简码切换
    If mbytInState = 0 Then
        '简码匹配方式：0-拼音,1-五笔,2-两者
        If gbytCode = 0 Then
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrRaised
        ElseIf gbytCode = 1 Then
            sta.Panels("PY").Bevel = sbrRaised
            sta.Panels("WB").Bevel = sbrInset
        Else
            sta.Panels("PY").Bevel = sbrInset
            sta.Panels("WB").Bevel = sbrInset
        End If
    End If
    
    InitFace = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function InitData() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lngCount As Long, strSQL As String

    On Error GoTo errH

    '自动识别加班
    If mbytInState <> 2 And mstrInNO = "" Then
        If OverTime(zlDatabase.Currentdate) Then chk加班.Value = Checked
    End If

    '可选性别
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 性别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo性别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 Then cbo性别.ListIndex = cbo性别.NewIndex
            rsTmp.MoveNext
        Next
    End If

    '可选费别
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            If rsTmp!缺省 = 1 And cbo费别.ListIndex = -1 Then cbo费别.ListIndex = cbo费别.NewIndex
            rsTmp.MoveNext
        Next
    Else
        MsgBox "没有初始化费别，请先到费别管理中进行设置！", vbInformation, gstrSysName
        Exit Function
    End If

   
    If FillDept() = 0 Then  '在设置listindex=0时调用FillDoctor
        If mbytUseType = Use门诊 Then
            MsgBox "没有初始化门诊临床科室,请先到部门管理中设置！", vbInformation, gstrSysName
        Else
            MsgBox "没有初始化住院临床科室,请先到部门管理中设置！", vbInformation, gstrSysName
        End If
        Exit Function
    End If

    '可用收费类别
    If gstr收费类别 = "" Then
        strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码 Not In ('1','4','5','6','7') Order by 序号"
    Else
        strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码 In(" & gstr收费类别 & ")  And 编码 Not In('4','5','6','7') Order by 序号"
    End If
    Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsClass.EOF Then
        MsgBox "没有设置可用的收费类别（本记帐模式不支持药品记帐），" & vbCrLf & _
               "请先在本地参数中设置！", vbInformation, gstrSysName
        Exit Function
    End If
    For i = 0 To mlngRows - 1
        cbo收费类别(i).Clear
    Next
    
    lngCount = 1
    Do Until mrsClass.EOF
        For i = 0 To mlngRows - 1
            cbo收费类别(i).AddItem lngCount & "-" & mrsClass("类别")
            cbo收费类别(i).ItemData(cbo收费类别(i).NewIndex) = Asc(mrsClass("编码"))
            
            '等于预设值
            If mrsClass("编码") = cbo收费类别(i).Tag Then
                cbo收费类别(i).ListIndex = cbo收费类别(i).NewIndex
                cbo收费类别(i).TabStop = False
            End If
        Next
        lngCount = lngCount + 1
        mrsClass.MoveNext
    Loop

    mblnOne = (mrsClass.RecordCount = 1)

    '执行部门(所有门诊或住院)
    strSQL = _
        "Select Distinct A.ID,A.编码,A.简码,A.名称,B.工作性质,B.服务对象 " & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID and B.服务对象 IN(" & IIf(mbytUseType = Use门诊, "1", "2") & ",3) " & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        " Order by B.服务对象,A.编码"
    Set mrsUnit = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrsUnit.EOF Then
        MsgBox "没有初始化部门信息,单据无法处理执行部门。请先到部门管理中设置！", vbInformation, gstrSysName
        Exit Function
    End If

    '开单日期
    txtDate.Text = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")

    If mbytInState = 0 Then Set mrsWarn = GetUnitWarn
    Set mrsInfo = New ADODB.Recordset

    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt主页ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub ReCalcInsure()
'功能：修改单据时,重新计算统筹金额及更新相关信息
    Dim i As Integer, j As Integer
    Dim strInfo As String
    
    If mrsInfo.State = 1 Then
        If Not IsNull(mrsInfo!险类) Then
            For i = 1 To mobjBill.Details.Count
                For j = 1 To mobjBill.Details(i).InComes.Count
                    strInfo = gclsInsure.GetItemInsure(mobjBill.病人ID, mobjBill.Details(i).收费细目ID, mobjBill.Details(i).InComes(j).实收金额, False, mrsInfo!险类, _
                        mobjBill.Details(i).摘要 & "||" & mobjBill.Details(i).数次)
                    If strInfo <> "" Then
                        mobjBill.Details(i).保险项目否 = Val(Split(strInfo, ";")(0)) <> 0
                        mobjBill.Details(i).保险大类ID = Val(Split(strInfo, ";")(1))
                        mobjBill.Details(i).InComes(j).统筹金额 = Val(Split(strInfo, ";")(2))
                        mobjBill.Details(i).保险编码 = CStr(Split(strInfo, ";")(3))
                        
                        If UBound(Split(strInfo, ";")) >= 4 Then
                            If CStr(Split(strInfo, ";")(4)) <> "" Then mobjBill.Details(i).摘要 = CStr(Split(strInfo, ";")(4))
                            If UBound(Split(strInfo, ";")) >= 5 Then
                                If Split(strInfo, ";")(5) <> "" Then mobjBill.Details(i).Detail.类型 = Split(strInfo, ";")(5)
                            End If
                        End If
                    End If
                Next
            Next
        End If
    End If
End Sub

Private Function BillExistInsure(strNO As String) As Integer
'功能：判断指定的住院记帐单据是否对医保病人记的帐
'参数：strNO=记帐单据号
'返回：如果是则返回病人险类
'说明：1.只管住院医保病人,不管门诊病人的医技记帐
'      2.记帐表只返回第一个病人的险类,单据中也应该只有一种险类
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select B.险类 From 住院费用记录 A,病案主页 B" & _
        " Where A.记录性质=2 And A.记录状态 IN(1,3) And B.险类 is Not NULL" & _
        " And A.NO=[1] And A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    
    If Not rsTmp.EOF Then BillExistInsure = rsTmp!险类
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckMediCareItem(ByVal lng收费细目ID As Long, ByVal int险类 As Integer, ByVal str收费项目名称 As String, _
    ByVal bln定价 As Boolean, Optional ByVal strPriceGrade As String) As Boolean
'功能：判断收费项目是否设置了保险支付项目
    Dim rsTmp As ADODB.Recordset, strSQL As String, dbl价格 As Double, rs价格 As ADODB.Recordset
    Dim strWherePriceGrade As String
    
    CheckMediCareItem = True
    
    If gbyt医保对码检查 = 0 Then Exit Function
    If gclsInsure.GetCapability(support允许不设置医保项目, , int险类) Then
        Exit Function
    End If
    On Error GoTo errH

   '刘兴洪 问题:27286 定价的价格为零的不进行检查对码 日期:2010-01-07 15:13:45
    If bln定价 Then
        If strPriceGrade <> "" Then
            strWherePriceGrade = _
                "      And (b.价格等级 = [2]" & vbNewLine & _
                "          Or (b.价格等级 Is Null" & vbNewLine & _
                "              And Not Exists(Select 1" & vbNewLine & _
                "                             From 收费价目" & vbNewLine & _
                "                             Where b.收费细目id = 收费细目id And 价格等级 = [2]" & vbNewLine & _
                "                                   And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD')))))"
        Else
            strWherePriceGrade = " And b.价格等级 Is Null"
        End If
        strSQL = _
        " Select  B.现价 " & _
        " From 收费价目 B " & _
        " Where   ((Sysdate Between B.执行日期 and B.终止日期) Or (Sysdate>=B.执行日期 And B.终止日期 is NULL))" & _
        "       And B.收费细目ID=[1]" & vbNewLine & _
                strWherePriceGrade
        Set rs价格 = zlDatabase.OpenSQLRecord(strSQL, "获取当前价格", lng收费细目ID, strPriceGrade)
        If rs价格.EOF = False Then
            dbl价格 = Val(NVL(rs价格!现价))
        Else
            dbl价格 = 0
        End If
        If dbl价格 = 0 Then Exit Function
    End If
    
    strSQL = "Select 收费细目ID From 保险支付项目 Where 收费细目ID=[1] And 险类=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlInExse", lng收费细目ID, int险类)
        
    If rsTmp.RecordCount = 0 Then
        If gbyt医保对码检查 = 1 Then
            If MsgBox("没有设置""" & str收费项目名称 & """对应的保险项目,要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckMediCareItem = False
            End If
        ElseIf gbyt医保对码检查 = 2 Then
            MsgBox "没有设置""" & str收费项目名称 & """对应的保险项目!", vbInformation, gstrSysName
            CheckMediCareItem = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check服务对象() As Integer
'功能：检查当前病人的记帐费用项目的服务对象是否一致
'说明：因为加入了门诊留观病人,所以有此检查
'返回：不一致的费用行,为0时正常
    Dim i As Integer
    
    If mrsInfo.State = 0 Then Exit Function
    For i = 1 To mobjBill.Details.Count
        If mrsInfo!病人性质 = 0 Or mrsInfo!病人性质 = 2 Then
            '住院病人或住院留观病人,不能用只服务于门诊的项目
            If mobjBill.Details(i).Detail.服务对象 = 1 Then
                MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于门诊,该病人不能使用.", vbInformation, gstrSysName
                Check服务对象 = i: Exit Function
            End If
        ElseIf mrsInfo!病人性质 = 1 Or mrsInfo!病人性质 = -1 Then
            '门诊或出院病人(医技记帐)或门诊留观病人,不能用只服务于住院的项目
            If mobjBill.Details(i).Detail.服务对象 = 2 Then
                MsgBox "第 " & i & " 行项目""" & mobjBill.Details(i).Detail.名称 & """仅服务于住院,该病人不能使用.", vbInformation, gstrSysName
                Check服务对象 = i: Exit Function
            End If
        End If
    Next
End Function
Private Function Get开单科室ID() As Long
    If cbo开单科室.ListIndex <> -1 Then
        Get开单科室ID = cbo开单科室.ItemData(cbo开单科室.ListIndex)
    Else
        Get开单科室ID = IIf(mlngDeptID = 0, UserInfo.部门ID, mlngDeptID)
    End If
End Function
