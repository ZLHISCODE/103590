VERSION 5.00
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#2.0#0"; "zlIDKind.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSurety 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "住院担保信息管理"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frmSurety.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   7575
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicDeposit 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   720
      ScaleHeight     =   3090
      ScaleWidth      =   5790
      TabIndex        =   18
      Top             =   3360
      Width           =   5790
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
         Height          =   2265
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblDeposit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交总额："
         Height          =   180
         Left            =   45
         TabIndex        =   19
         Top             =   75
         Width           =   900
      End
   End
   Begin VB.Frame fraPati 
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   7350
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1350
         TabIndex        =   3
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton cmdPati 
         Height          =   300
         Left            =   2640
         Picture         =   "frmSurety.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F2)"
         Top             =   225
         Width           =   300
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         ToolTipText     =   "快捷键F4"
         Top             =   225
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   $"frmSurety.frx":0914
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblCur 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款方式："
         Height          =   180
         Left            =   5085
         TabIndex        =   22
         Top             =   285
         Width           =   1260
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别等级："
         Height          =   180
         Left            =   5085
         TabIndex        =   21
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblSex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         Height          =   180
         Left            =   2985
         TabIndex        =   5
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblAge 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄："
         Height          =   180
         Left            =   3960
         TabIndex        =   6
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号："
         Height          =   180
         Left            =   330
         TabIndex        =   7
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室："
         Height          =   180
         Left            =   2325
         TabIndex        =   8
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblBed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号："
         Height          =   180
         Left            =   3960
         TabIndex        =   9
         Top             =   630
         Width           =   540
      End
   End
   Begin VB.PictureBox picSurety 
      BorderStyle     =   0  'None
      Height          =   3900
      Left            =   120
      ScaleHeight     =   3900
      ScaleWidth      =   7425
      TabIndex        =   16
      Top             =   1080
      Width           =   7425
      Begin VB.Frame fraEdit 
         Caption         =   "信息输入"
         Height          =   1095
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   7335
         Begin VB.TextBox txt担保额 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   9
            TabIndex        =   28
            Top             =   360
            Width           =   1005
         End
         Begin VB.TextBox txt担保人 
            Height          =   300
            Left            =   840
            MaxLength       =   100
            TabIndex        =   27
            Top             =   360
            Width           =   1005
         End
         Begin VB.CheckBox chkUnlimit 
            Caption         =   "不限额度"
            Height          =   255
            Left            =   2760
            TabIndex        =   26
            ToolTipText     =   "不限担保额时必须设置担保时限"
            Top             =   720
            Width           =   1050
         End
         Begin VB.CheckBox chk临时担保 
            Caption         =   "临时担保"
            Height          =   255
            Left            =   840
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   720
            Width           =   1050
         End
         Begin VB.TextBox txtReason 
            Height          =   300
            Left            =   5040
            MaxLength       =   50
            TabIndex        =   24
            Top             =   720
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker dtp担保时限 
            Height          =   300
            Left            =   5040
            TabIndex        =   29
            Top             =   360
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   190906371
            CurrentDate     =   38915.6041666667
         End
         Begin VB.Label lbl担保额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "担保额"
            Height          =   180
            Left            =   2160
            TabIndex        =   33
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl担保人 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "担保人"
            Height          =   180
            Left            =   240
            TabIndex        =   32
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl担保时限 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "到期时间"
            Height          =   180
            Left            =   4140
            TabIndex        =   31
            ToolTipText     =   "在院病人才能使用时限担保"
            Top             =   450
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "担保原因"
            Height          =   180
            Left            =   4140
            TabIndex        =   30
            Top             =   780
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdAdd 
         Cancel          =   -1  'True
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "仅当最近一条担保记录到期或没有限制期限时才允许增加"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "修改(&M)"
         Height          =   350
         Left            =   1350
         TabIndex        =   11
         ToolTipText     =   "只允许修改最近一条担保记录"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   2450
         TabIndex        =   12
         ToolTipText     =   "只允许删除最近一条担保记录"
         Top             =   1200
         Width           =   1100
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&X)"
         Height          =   350
         Left            =   6000
         TabIndex        =   13
         ToolTipText     =   "(F9)退出"
         Top             =   1200
         Width           =   1100
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
         Height          =   2265
         Left            =   0
         TabIndex        =   14
         Top             =   1680
         Width           =   7305
         _ExtentX        =   12885
         _ExtentY        =   3995
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FixedCols       =   0
         RowHeightMin    =   250
         BackColorBkg    =   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   5145
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9499
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3775
            MinWidth        =   3775
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
   Begin XtremeSuiteControls.TabControl tbcPage 
      Height          =   3945
      Left            =   555
      TabIndex        =   17
      Top             =   1035
      Width           =   3795
      _Version        =   589884
      _ExtentX        =   6694
      _ExtentY        =   6959
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmSurety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long
Private mbln在院病人 As Boolean
Private mstrPrivs As String
Private mlng主页ID As Long      '在院病人为当前住院登记的主页ID
Private mlngModul As Long       '0-病人信息管理;1-病人入院管理

'参数
Private mblnSeekName As Boolean         '姓名模糊查找
Private mblnCheckPass As Boolean        '入院登记时刷卡输入密码


Private mrsInfo As New ADODB.Recordset
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mobjOneCardCom As Object
Private mblnDefaultPassInputCardNo As Boolean '缺省刷卡是否密文输入卡号
Private mblnNotClick As Boolean
Private mblnFirst As Boolean
Private mstr病人类型 As String

Private Enum E_COL_Deposit
    COL_日期 = 1
    COL_单据号 = 2
    COL_科室
    COL_金额
    COL_缴款金额
    COL_结算
    COL_收款人
End Enum

Private Enum E_COL_Surety
    COL_类别 = 1
    COL_担保人
    COL_担保额
    COL_临时担保
    COL_担保原因
    COL_登记时间
    COL_到期时间
    COL_删除标志
    COL_操作员姓名
    COL_操作员编号
    COL_删除操作员姓名
    COL_删除操作员编号
    COL_删除时间
End Enum
 
Public Function ShowMe(ByRef frmMain As Form, ByVal lngPatiid As Long, ByVal bln在院病人 As Boolean, ByVal strPrivs As String, _
    ByVal lngModul As Long, Optional blnSeekName As Boolean, Optional blnCheckPass As Boolean) As Boolean
'功能:担保信息管理
'   frmMain-主窗体对象
'   lngPatiId-病人ID
'   bln在院病人-T-在院病人
'   strPrivs-权限
'   lngModul= 1101-病人信息管理,1131-P病人入院管理
'   blnSeekName-按姓名模糊查找(模块=病人入院管理时传入此值)
    mlng病人ID = lngPatiid
    mbln在院病人 = bln在院病人
    mstrPrivs = strPrivs
    mlngModul = lngModul
    
    If mlngModul = P病人入院管理 Then
        mblnSeekName = blnSeekName  '
        mblnCheckPass = blnCheckPass
    End If
    
    Me.Show 1, frmMain
    ShowMe = True
End Function
  
Private Sub chkUnlimit_Click()
     '不限担保额不能是临时担保
    If chkUnlimit.Value = 1 And IsNull(dtp担保时限.Value) Then
        dtp担保时限.Value = DateAdd("d", 3, dtp担保时限.MinDate)
    End If
    chk临时担保.Enabled = Not (chkUnlimit.Value = 1)
    txt担保额.Enabled = Not (chkUnlimit.Value = 1)
    
    If chkUnlimit.Value = 1 Then
        txt担保额.Text = "999999999":  txt担保额.BackColor = vbInactiveCaptionText
    Else
        txt担保额.Text = "": txt担保额.BackColor = vbWhite
    End If
End Sub

Private Sub chk临时担保_Click()
    If chk临时担保.Value = 1 Then
        dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
        dtp担保时限.Value = Null
        chkUnlimit.Value = 0        '值改变时有隐式调用click事件
    End If
    chkUnlimit.Enabled = Not (chk临时担保.Value = 1) And mbln在院病人
    dtp担保时限.Enabled = Not (chk临时担保.Value = 1) And mbln在院病人
End Sub

Private Sub cmdDel_Click()
    Dim strJsonIn As String, strJsonOut As String
    Dim str登记时间 As String
    Dim str删除标志 As String
    Dim blnOk As Boolean
    
    If mlngModul = P病人入院管理 Then
        blnOk = True
        If mrsInfo Is Nothing Then
            blnOk = False
        ElseIf mrsInfo.State = adStateClosed Then
            blnOk = False
        End If
        
        If blnOk = False Then
            stbThis.Panels(1).Text = "没有确定要进行担保的病人!"
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    End If
    '问题21368 by lesfeng 2010-08-02
    str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
    If str删除标志 = "删除" Then
        MsgBox "此条担保记录已经为删除标记，不能进行删除标记操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("确实要进行标记此条担保记录为删除标记吗?" & vbCrLf & vbCrLf & "注意,删除标记后，当前担保将会不能恢复!" _
        , vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errH
    
    If Trim(msh.TextMatrix(msh.Row, GetColNum("登记时间"))) = "" Then
        str登记时间 = ""
    Else
        str登记时间 = Format(Trim(msh.TextMatrix(msh.Row, GetColNum("登记时间"))), "YYYY-MM-DD HH:MM:SS")
    End If
    '问题21368 by lesfeng 2010-08-02
    strJsonIn = GetJsonSurety(3, mlng病人ID, mlng主页ID, , , , , , str登记时间)
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "删除担保记录", P病人入院管理, False, , , , True)
    End If
    
    stbThis.Panels(1).Text = "删除操作成功!"
    Call LoadSurety
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdModify_Click()
    Dim strSQL As String, str担保人 As String, str到期时间 As String
    Dim str登记时间 As String
    Dim str删除标志 As String
    Dim blnOk As Boolean
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    '只能修改当前选中并且有效的担保记录
    
    If cmdModify.Caption = "修改(&M)" Then
        If mlngModul = P病人入院管理 Then
            blnOk = True
            If mrsInfo Is Nothing Then
                blnOk = False
            ElseIf mrsInfo.State = adStateClosed Then
                blnOk = False
            End If
            
            If blnOk = False Then
                stbThis.Panels(1).Text = "没有确定要进行担保的病人!"
                If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
                Exit Sub
            End If
        End If
    
        '提取修改信息
        If msh.TextMatrix(msh.Row, GetColNum("担保人")) = "" Then
            stbThis.Panels(1).Text = "没有可以修改的担保信息!"
            Exit Sub
        End If
        '问题21368 by lesfeng 2010-08-02
        str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
        If str删除标志 = "删除" Then
            MsgBox "此条担保记录已经为删除标记，不能进行修改操作！", vbInformation, gstrSysName
            Exit Sub
        End If
        cmdModify.Caption = "保存(&S)"
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
        cmdExit.Caption = "取消(&C)"
        fraEdit.Enabled = True
        
        With msh
            txt担保人.Text = Trim(.TextMatrix(.Row, GetColNum("担保人")))
            If .TextMatrix(.Row, GetColNum("担保额")) = "不限" Then
                chkUnlimit.Value = 1    '值不同时隐式调用click事件
                txt担保额.Text = "999999999"
            Else
                chkUnlimit.Value = 0
                txt担保额.Text = Val(.TextMatrix(.Row, GetColNum("担保额")))
            End If
            
            If IsDate(.TextMatrix(.Row, GetColNum("到期时间"))) Then
                dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm"
                dtp担保时限.Value = CDate(.TextMatrix(.Row, GetColNum("到期时间")))
            Else
                dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm" '如果不可见，下面句执行会出错
                dtp担保时限.Value = Null
            End If
            
            chk临时担保.Value = IIf(.TextMatrix(.Row, GetColNum("临时担保")) = "√", 1, 0)
            If txt担保人.Enabled Then txt担保人.SetFocus
            txt担保人.Tag = Trim(.TextMatrix(msh.Row, GetColNum("登记时间")))
        End With
    Else
    '保存修改结果
        '1.数据检查
        If Not Check担保信息 Then Exit Sub
        '先恢复界面按钮状态
        cmdModify.Caption = "修改(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "退出(&X)"
        fraEdit.Enabled = True      'SetCanEdit会再次设置
        
        str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
        str到期时间 = ""
        If Not IsNull(dtp担保时限.Value) Then str到期时间 = Format(dtp担保时限.Value, "YYYY-MM-DD hh:mm:ss")
        str登记时间 = Format(txt担保人.Tag, "YYYY-MM-DD hh:mm:ss")
        
        '长度检查
        If Not CheckLen(txt担保人, 64) Then Exit Sub
        
        '2.数据保存
        On Error GoTo errH
        strJsonIn = GetJsonSurety(2, mlng病人ID, mlng主页ID, str担保人, Val(txt担保额.Text), chk临时担保.Value, Trim(txtReason.Text), str到期时间, str登记时间)
        If strJsonIn <> "" Then
            Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "更新担保记录", P病人入院管理, False, , , , True)
        End If
    
        '3.数据刷新
        stbThis.Panels(1).Text = "修改结果已保存!"
        Call LoadSurety
        Call Init担保信息
        If cmdExit.Enabled Then cmdExit.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init担保信息()
    Dim Datsys As Date

    txt担保人.Text = ""
    chkUnlimit.Enabled = mbln在院病人
    chkUnlimit.Value = 0            '如果值有变化,则隐式调用click事件
    txt担保额.Text = ""
    txtReason.Text = ""
    
    dtp担保时限.Enabled = mbln在院病人
    dtp担保时限.CheckBox = True: dtp担保时限.CustomFormat = "yyyy-MM-dd HH:mm" '设置checkbox可见性
    If dtp担保时限.Enabled Then
        Datsys = zldatabase.Currentdate
        dtp担保时限.MinDate = Datsys
        dtp担保时限.Value = DateAdd("d", 3, Datsys)
    End If
    dtp担保时限.Value = Null
    
    chk临时担保.Enabled = True
    chk临时担保.Value = 0
    chkUnlimit.TabStop = True
End Sub

Public Sub InitFace()
    lblSex.Caption = "性别：": lblNO.Caption = "住院号：": lblBed.Caption = "床号："
    lblAge.Caption = "年龄：": lblDept.Caption = "科室：": lblDeposit.Caption = "预交总额："
    lblType.Caption = "费别等级：": lblCur.Caption = "医疗付款方式："
End Sub

Private Sub cmdPati_Click()
    If frmPatiSelect.ShowMe(Me) = True Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call txtPatient_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub dtp担保时限_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii = vbKeySpace Then
        If dtp担保时限.CheckBox Then
            KeyAscii = 0
            If IsNull(dtp担保时限.Value) Then
                dtp担保时限.Value = DateAdd("d", 3, zldatabase.Currentdate)
            Else
                dtp担保时限.Value = Null
            End If
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mlngModul = P病人入院管理 Then
        If mblnFirst = True Then Exit Sub
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        mblnFirst = True
    End If
End Sub

Private Sub Form_Load()
        
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    Me.Height = 5925
    If mlngModul = P病人信息管理 Then
        Call LoadSurety
        Call Init担保信息
        
        If InStr(mstrPrivs, "担保信息增加") <= 0 Then
            cmdAdd.Visible = False
        End If
        
        If InStr(mstrPrivs, "担保信息调整") <= 0 Then
            cmdModify.Visible = False
        End If
        
        If InStr(mstrPrivs, "担保信息删除") <= 0 Then
            cmdDel.Visible = False
        End If
        
        If InStr(mstrPrivs, "担保信息增加") <= 0 And InStr(mstrPrivs, "担保信息调整") And InStr(mstrPrivs, "担保信息删除") <= 0 Then
            fraEdit.Enabled = False
            Me.Caption = "担保信息查看(当前用户：" & UserInfo.姓名 & ")"
        End If
        
        Me.Caption = "担保信息管理"
        Me.fraPati.Visible = False
        Me.PicDeposit.Visible = False
        Me.Height = Me.Height - fraPati.Height
        Me.picSurety.Top = fraPati.Top
        Me.tbcPage.Visible = False
    Else
        mblnFirst = False
        Call InitTabPage
        Call zlCardSquareObject
        Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, mobjOneCardCom, "", txtPatient)
        Set mobjIDCard = New clsIDCard
        Set mobjICCard = New clsICCard
        Call mobjIDCard.SetParent(Me.hWnd)
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
        IDKind.Enabled = True
        
        If Not mobjOneCardCom Is Nothing Then
            IDKind.IDKindStr = mobjOneCardCom.zlGetIDKindStr(IDKind.IDKindStr)
        End If
        
        Call ClearWinInfor(True)
        
        fraEdit.Enabled = True
        If InStr(mstrPrivs, "办理登记") <= 0 And InStr(mstrPrivs, "接收预约") = 0 And InStr(mstrPrivs, "保险病人登记") <= 0 Then
            fraEdit.Enabled = False
            cmdAdd.Visible = False
            cmdModify.Visible = False
            cmdDel.Visible = False
            Me.Caption = "住院担保信息查看(当前用户：" & UserInfo.姓名 & ")"
        End If
        txt担保人.Enabled = fraEdit.Enabled
        txt担保人.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        txt担保额.Enabled = fraEdit.Enabled
        txt担保额.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        chk临时担保.Enabled = fraEdit.Enabled
        chkUnlimit.Enabled = fraEdit.Enabled
        txtReason.Enabled = fraEdit.Enabled
        txtReason.BackColor = IIf(fraEdit.Enabled, &H80000005, &H8000000F)
        If mlng病人ID > 0 Then
            txtPatient.Text = "-" & mlng病人ID
            Call txtPatient_KeyPress(vbKeyReturn)
        Else
            cmdAdd.Enabled = False
        End If
        Me.Caption = "住院担保信息管理"
    End If
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
End Sub

Private Sub ClearWinInfor(Optional ByVal blnClear As Boolean = False)
    Call InitFace
    Call LoadSurety(blnClear)
    Call LoadPrepay(blnClear)
    Call Init担保信息
End Sub

Private Sub InitTabPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化分页控件
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
        
    Set objItem = tbcPage.InsertItem(1, "担保信息", picSurety.hWnd, 0)
    objItem.Tag = 1
    
    Set objItem = tbcPage.InsertItem(2, "预交信息", PicDeposit.hWnd, 0)
    objItem.Tag = 2
    
    With tbcPage
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To msh.Cols - 1
        If msh.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
    GetColNum = -1
End Function

Private Function GetColNumList(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNumList = i: Exit Function
    Next
    GetColNumList = -1
End Function

Private Sub SetSuretyHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|类别,4,1000|担保人,4,800|担保额,7,1250|临时担保,4,850|担保原因,4,1800|登记时间,1,1800|到期时间,1,1800|删除标志,4,850|操作员姓名,4,1050|操作员编号,4,1050|删除操作员姓名,4,1050|删除操作员编号,4,1050|删除时间,1,1800"
    With msh
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        .FixedRows = 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub SetDepositHeader()
    Dim strHead As String, i As Long
    strHead = ",4,300|日期,4,1350|单据号,4,1110|科室,1,1200|金额,1,0|缴款金额,7,1600|结算,4,1000|收款人,1,1000"
    With mshList
        .Redraw = False
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(msh, App.ProductName & "\" & Me.Name)
        
        .ForeColor = &H80000003
        .RowHeight(0) = 320
        .Redraw = True
    End With
End Sub

Private Sub GetSuretyBalance()
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim dblMoney As Double
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim colPati As Collection
    
    On Error GoTo errH
    If mlngModul = P病人信息管理 Then
        If zl_PatiSvr_GetPatiInfo(mlng病人ID, Nothing, colPati, 2) Then
            If colPati(1)("_pati_dept_id") = 0 Then
                mlng主页ID = 0
            Else
                mlng主页ID = colPati(1)("_pati_pageid")
            End If
        End If
    End If
    strJsonIn = GetNode("pati_id", mlng病人ID, True)
    strJsonIn = strJsonIn & GetNode("pati_pageid", mlng主页ID)
    strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
    If CallService("Zl_Exsesvr_Getpatisurety", strJsonIn, strJsonOut, "获取有效担保额", mlngModul) Then
        dblMoney = Val(GetJsonNodeValue("output.guarantee_money") & "")
        If dblMoney > 0 Then
            stbThis.Panels(2).Text = "有效担保额:" & dblMoney
        Else
            stbThis.Panels(2).Text = "有效担保额:无"
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadSurety(Optional ByVal blnClear As Boolean = False)
    Dim Curdate As Date
    Dim i As Integer, lngRow As Integer, RowPageid As Integer
    Dim str删除标志 As String
    Dim lng病人ID As Long, lng主页id As Long
    Dim collist As Collection
    Dim strJsonIn As String
    Dim strJsonOut As String
    
    On Error GoTo errH
    If mrsInfo Is Nothing Then
        lng病人ID = mlng病人ID
        lng主页id = mlng主页ID
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = mlng病人ID
        lng主页id = mlng主页ID
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        lng主页id = Val(Nvl(mrsInfo!主页ID))
    End If
    stbThis.Panels(2).Text = ""
    If blnClear = True Then
        msh.Rows = 1
        msh.Rows = 2
        msh.RowData(1) = 0
        Call SetSuretyHeader
    Else
        Call SetSuretyHeader
        strJsonIn = GetNode("pati_id", lng病人ID, True)
        If mlngModul = P病人入院管理 Then strJsonIn = strJsonIn & GetNode("pati_pageid", lng主页id)
        strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
        
        If CallService("Zl_Exsesvr_Getpatisuretylist", strJsonIn, strJsonOut, "获取担保记录", mlngModul) Then
            Set collist = GetJsonListValue("output.item_list")
        Else
            Set collist = New Collection
        End If
        If collist.Count > 0 Then
            With msh
                .Redraw = False
                .Rows = collist.Count + 1
                .FixedRows = 1
                For i = 1 To collist.Count
                    .RowData(i) = lng病人ID
                    .TextMatrix(i, COL_类别) = collist(i)("_type") & ""
                    .TextMatrix(i, COL_担保人) = collist(i)("_guarantor") & ""
                    .TextMatrix(i, COL_担保额) = collist(i)("_garnt_amount") & ""
                    .TextMatrix(i, COL_临时担保) = IIf(Nvl(collist(i)("_garnt_prop"), 0) = 1, "√", "")
                    .TextMatrix(i, COL_担保原因) = Nvl(collist(i)("_garnt_reason"))
                    .TextMatrix(i, COL_登记时间) = collist(i)("_create_time") & ""
                    .TextMatrix(i, COL_到期时间) = Nvl(collist(i)("_due_time"))
                    .TextMatrix(i, COL_删除标志) = collist(i)("_is_del") & ""
                    .TextMatrix(i, COL_操作员姓名) = collist(i)("_operator_name") & ""
                    .TextMatrix(i, COL_操作员编号) = collist(i)("_operator_code") & ""
                    .TextMatrix(i, COL_删除操作员姓名) = collist(i)("_del_operator_name") & ""
                    .TextMatrix(i, COL_删除操作员编号) = collist(i)("_del_operator_code") & ""
                    .TextMatrix(i, COL_删除时间) = collist(i)("_del_time") & ""
                Next
                .Redraw = True
            End With
        Else
            msh.Rows = 1 '不清空表头
            msh.Rows = 2
            msh.FixedRows = 1
        End If
        
        Call GetSuretyBalance
        Curdate = zldatabase.Currentdate
        For lngRow = 1 To msh.Rows - 1
            If UBound(Split(Trim(msh.TextMatrix(lngRow, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(lngRow, GetColNum("类别"))), "次住院")(0), "第")(1))
            Else
                RowPageid = 0
            End If
            '问题21368 by lesfeng 2010-08-02
            str删除标志 = Trim(msh.TextMatrix(lngRow, GetColNum("删除标志")))
            
            If lng主页id = RowPageid And (Trim(msh.TextMatrix(lngRow, GetColNum("到期时间"))) = "" Or Trim(msh.TextMatrix(lngRow, GetColNum("到期时间"))) > Curdate) Then
                msh.Row = lngRow
                For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '问题21368 by lesfeng 2010-08-02
                    If str删除标志 = "" Then
                        msh.CellForeColor = &HC00000
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            Else
                 For i = 0 To msh.Cols - 1
                    msh.Col = i
                    '问题21368 by lesfeng 2010-08-02
                    If str删除标志 = "" Then
                    Else
                        msh.CellForeColor = &HFF&
                    End If
                Next
            End If
            
        Next lngRow
    End If
    msh.Row = 1
    msh.Col = 0: msh.ColSel = msh.Cols - 1
    Call msh_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPrepay(Optional ByVal blnClear As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示历史的预交数据
    '编制:刘鹏飞
    '日期:2013-03-11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lng主页id As Long
    Dim strJsonIn As String
    Dim strJsonOut As String
    Dim collist As Collection
    Dim i As Long
    
    If mrsInfo Is Nothing Then
        lng病人ID = mlng病人ID
        lng主页id = mlng主页ID
    ElseIf mrsInfo.State <> 1 Then
        lng病人ID = mlng病人ID
        lng主页id = mlng主页ID
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        lng主页id = Val(Nvl(mrsInfo!主页ID))
    End If
    
    On Error GoTo errHandle
    
    If blnClear = True Then
        mshList.Clear
        mshList.Rows = 2
        Call SetDepositHeader
    Else
        '所有历史缴款明细清单
        strJsonIn = GetNode("pati_id", lng病人ID, True)
        strJsonIn = strJsonIn & GetNode("pati_pageid", lng主页id)
        strJsonIn = strJsonIn & GetNode("type", 2)
        strJsonIn = GetNode("input", "{" & strJsonIn & "}", True, 1)
        
        If CallService("Zl_Exsesvr_Getdepositlist", strJsonIn, strJsonOut, "获取预交记录", P病人入院管理) Then
            Set collist = GetJsonListValue("output.item_list")
        Else
            Set collist = New Collection
        End If
        If collist.Count > 0 Then
            With mshList
                .Redraw = False
                .Rows = collist.Count + 1
                For i = 1 To collist.Count
                    .TextMatrix(i, COL_日期) = collist(i)("_create_date")
                    .TextMatrix(i, COL_单据号) = collist(i)("_bill_no")
                    .TextMatrix(i, COL_科室) = collist(i)("_dept_name") & ""
                    .TextMatrix(i, COL_金额) = collist(i)("_money")
                    .TextMatrix(i, COL_缴款金额) = Format(collist(i)("_money"), "##,##0.00")
                    .TextMatrix(i, COL_结算) = collist(i)("_blnc_mode")
                    .TextMatrix(i, COL_收款人) = collist(i)("_operator_name")
                Next
                .Redraw = True
            End With
        Else
            mshList.Clear
            mshList.Rows = 2
        End If
         Call SetDepositHeader
    End If
     
    If mshList.Rows > 1 Then
        mshList.Row = 1: mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check担保信息() As Boolean
    Check担保信息 = True
    If mlngModul = P病人入院管理 Then
        If mrsInfo Is Nothing Then
            Check担保信息 = False
        ElseIf mrsInfo.State = adStateClosed Then
            Check担保信息 = False
        End If
        
        If Check担保信息 = False Then
            stbThis.Panels(1).Text = "没有确定要进行担保的病人!"
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    If Trim(txt担保人.Text) = "" Then
        stbThis.Panels(1).Text = "请输入担保人姓名,担保人不能为空!"
        If txt担保人.Enabled Then txt担保人.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    If Not IsNumeric(txt担保额.Text) Then
        stbThis.Panels(1).Text = "请输入正确的担保额,担保额要求是数值!"
        If txt担保额.Enabled Then txt担保额.SetFocus
        Check担保信息 = False
        Exit Function
    ElseIf Val(txt担保额.Text) = 0 Then
        stbThis.Panels(1).Text = "请输入担保额,担保额不能为零!"
        If txt担保额.Enabled Then txt担保额.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
    If chk临时担保.Value = 1 Then
        If Not IsNull(dtp担保时限.Value) Or chkUnlimit.Value = 1 Then
            stbThis.Panels(1).Text = "临时担保不允许设置担保时限或不限担保额!"
            If chk临时担保.Enabled Then chk临时担保.SetFocus
            Check担保信息 = False
            Exit Function
        End If
    End If
    
    If zlCommFun.ActualLen(Trim(txtReason.Text)) > 50 Then
        stbThis.Panels(1).Text = "担保原因过长，最多允许 25 个汉字或 50 个字符。"
        txtReason.SetFocus
        Check担保信息 = False
        Exit Function
    End If
    
End Function

Private Sub cmdAdd_Click()
    Dim str担保人 As String, str到期时间 As String
    Dim strJsonIn As String, strJsonOut As String
    Dim i As Integer, Curdate As Date, bln未到期 As Boolean, bln临时 As Boolean, RowPageid As Integer
    Dim str删除标志 As String
    
    '1.数据检查
    If Not Check担保信息 Then Exit Sub
    
    Curdate = zldatabase.Currentdate
    
    For i = 1 To msh.Rows - 1 '判断本次住院未到期的担保记录，加以提示
         If Trim(msh.TextMatrix(i, GetColNum("类别"))) <> "" Then
            If UBound(Split(Trim(msh.TextMatrix(i, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(i, GetColNum("类别"))), "次住院")(0), "第")(1))
            Else
                RowPageid = 0
            End If
            If mlng主页ID = RowPageid Then
                '问题21368 by lesfeng 2010-08-02
                str删除标志 = Trim(msh.TextMatrix(i, GetColNum("删除标志")))
               If (Trim(Nvl(msh.TextMatrix(i, GetColNum("到期时间")))) = "" Or Nvl(msh.TextMatrix(i, GetColNum("到期时间"))) > Curdate) And str删除标志 = "" Then
                   bln临时 = Nvl(msh.TextMatrix(i, GetColNum("临时担保"))) = "√"
                   bln未到期 = True: Exit For
               End If
            End If
        End If
    Next
    
    If bln未到期 Then
        If MsgBox("尚有未到期的" & IIf(bln临时, "临时", "") & "担保记录，新增将会" & IIf(bln临时, "让之前的临时担保自动失效", "累计担保") & "，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
        
    str担保人 = Replace(Trim(txt担保人.Text), "'", "''")
    str到期时间 = ""
    If Not IsNull(dtp担保时限.Value) Then str到期时间 = Format(dtp担保时限.Value, "yyyy-MM-dd HH:mm:ss")
    
    '长度检查
    If Not CheckLen(txt担保人, 64) Then Exit Sub
    
    '2.数据保存
    On Error GoTo errH
    strJsonIn = GetJsonSurety(1, mlng病人ID, mlng主页ID, str担保人, Val(txt担保额.Text), chk临时担保.Value, Trim(txtReason.Text), str到期时间)
    If strJsonIn <> "" Then
        Call CallService("Zl_Exsesvr_Updatepatisurety", strJsonIn, strJsonOut, "新增担保记录", P病人入院管理, True, , , , True)
    End If
    
    '3.数据刷新
    stbThis.Panels(1).Text = "新增信息已保存!"
    Call LoadSurety
    Call Init担保信息
    
    If cmdExit.Enabled Then cmdExit.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    
    If cmdExit.Caption = "取消(&C)" Then
        cmdModify.Caption = "修改(&M)"
        cmdAdd.Enabled = True
        cmdDel.Enabled = True
        cmdExit.Caption = "退出(&X)"
        fraEdit.Enabled = True      'SetCanEdit会再次设置
       
        '刷新数据,考虑并发操作
        stbThis.Panels(1).Text = ""
        Call LoadSurety
        Call Init担保信息
    Else
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim obj As Control
    If mlngModul = P病人信息管理 Then
       If KeyCode = vbKeyEscape Then
            Call cmdExit_Click
        End If
    Else
        Select Case KeyCode
        Case vbKeyEscape
            Call cmdExit_Click
        Case vbKeyF2
            Call cmdPati_Click
        Case vbKeyF4
            If Shift = vbCtrlMask And IDKind.Enabled Then
                Dim intIndex As Integer
                intIndex = IDKind.GetKindIndex("IC卡号")
                If intIndex <= 0 Then Exit Sub
                 IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
            End If
        Case vbKeyF11
            If txtPatient.Enabled And Not txtPatient.Locked Then txtPatient.SetFocus
        Case vbKeyReturn
            Set obj = Me.ActiveControl
            If InStr(1, ",txt担保人,txt担保额,dtp担保时限,chk临时担保,chkUnlimit,txtReason,", "," & obj.Name & ",") > 0 Then
               ' Call zlCommFun.PressKey(vbKeyTab)
            End If
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With tbcPage
        .Left = fraPati.Left
        .Top = fraPati.Top + fraPati.Height
        .Width = fraPati.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height
    End With
    
    PicDeposit.Width = picSurety.Width
    PicDeposit.Height = picSurety.Height
    
    With msh
        .Width = picSurety.ScaleWidth
        .Height = picSurety.ScaleHeight - .Top
    End With
    
    With lblDeposit
        .Left = 60
        .Top = 60
    End With
    
    With mshList
        .Top = lblDeposit.Top + lblDeposit.Height + 60
        .Left = 0
        .Width = msh.Width
        .Height = PicDeposit.ScaleHeight - .Top
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 
    If cmdModify.Caption = "保存(&S)" Then
        If MsgBox("当前修改的信息未保存,确实要退出吗?", vbYesNo + vbDefaultButton2 + vbInformation, gstrSysName) = vbNo Then Cancel = 1
    End If
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    If mlngModul = P病人入院管理 Then
        If Not mobjIDCard Is Nothing Then
            Call mobjIDCard.SetEnabled(False)
            Set mobjIDCard = Nothing
        End If
        If Not mobjICCard Is Nothing Then
            Call mobjICCard.SetEnabled(False)
            Set mobjICCard = Nothing
        End If
        Call zlCardSquareObject(True)
        Set mrsInfo = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlOneCardComLib.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
               Set mobjICCard = New clsICCard
               Call mobjICCard.SetParent(Me.hWnd)
               Set mobjICCard.gcnOracle = gcnOracle
        End If
           If Not mobjICCard Is Nothing Then
               txtPatient.Text = mobjICCard.Read_Card()
               If txtPatient.Text <> "" Then
                   Call txtPatient_KeyPress(vbKeyReturn)
               End If
           End If
           Exit Sub
    End If
     
    lng卡类别ID = objCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If mobjOneCardCom.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub
 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    Call txtPatient_GotFocus
    txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlOneCardComLib.Card, objPatiInfor As zlOneCardComLib.clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Text <> "" Or txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNo As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("IC卡号")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strCardNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long, lngIndex As Long
    
    If txtPatient.Text = "" And Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        lngIndex = IDKind.GetKindIndex("身份证号")
        If lngIndex >= 0 Then IDKind.IDKind = lngIndex
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub msh_EnterCell()
    Dim str到期时间 As String
    Dim Datsys As Date, RowPageid As Integer
    Dim str删除标志 As String
        
    If mlngModul = P病人入院管理 Then
        If Val(msh.RowData(msh.Row)) <= 0 Then
            stbThis.Panels(1).Text = ""
            cmdModify.Enabled = False
            cmdDel.Enabled = False
            Exit Sub
        End If
    End If
    
   '当前行主页与病人主页不同时不允许修改删除,已过期不允许修改删除
    Datsys = zldatabase.Currentdate
    
    '问题21368 by lesfeng 2010-08-02
    str删除标志 = Trim(msh.TextMatrix(msh.Row, GetColNum("删除标志")))
    
    If cmdModify.Caption = "修改(&M)" Then
        If mlng主页ID = 0 And Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))) = "门诊" Then
            '问题21368 by lesfeng 2010-08-02
            If str删除标志 = "" Then
                cmdModify.Enabled = True
                cmdDel.Enabled = True
                stbThis.Panels(1).Text = "当前担保记录有效"
            Else
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "当前担保记录已经标记删除"
            End If
        Else
            If UBound(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))), "次住院")) > 0 Then '取出选中行主页ID
                RowPageid = Val(Split(Split(Trim(msh.TextMatrix(msh.Row, GetColNum("类别"))), "次住院")(0), "第")(1))
            Else
                RowPageid = 0
            End If
            If mlng主页ID <> RowPageid Then
                cmdModify.Enabled = False
                cmdDel.Enabled = False
                stbThis.Panels(1).Text = "当前担保记录非本次住院担保。"
            Else
                str到期时间 = Trim(msh.TextMatrix(msh.Row, GetColNum("到期时间")))
            
                If str到期时间 <> "" Then
                    If CDate(str到期时间) < Datsys Then
                         cmdModify.Enabled = False
                         cmdDel.Enabled = False
                        '问题21368 by lesfeng 2010-08-02
                         If str删除标志 = "" Then
                            stbThis.Panels(1).Text = "当前担保记录已过期"
                        Else
                            stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                        End If
                    Else
                        '问题21368 by lesfeng 2010-08-02
                        If str删除标志 = "" Then
                            cmdModify.Enabled = True
                            cmdDel.Enabled = True
                            stbThis.Panels(1).Text = "当前担保记录有效"
                        Else
                            cmdModify.Enabled = False
                            cmdDel.Enabled = False
                            stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                        End If
                    End If
                Else
                    '问题21368 by lesfeng 2010-08-02
                    If str删除标志 = "" Then
                        cmdModify.Enabled = True
                        cmdDel.Enabled = True
                        stbThis.Panels(1).Text = "当前担保记录有效"
                    Else
                        cmdModify.Enabled = False
                        cmdDel.Enabled = False
                        stbThis.Panels(1).Text = "当前担保记录已经标记删除"
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim dblMoney As Double, lngRow As Long
    
    If txtPatient.Locked Then Exit Sub
        
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        txtPatient.IMEMode = 0
    End If
    
    If txtPatient.Tag <> "" Then Exit Sub
    
    If Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13 Then
        If frmPatiSelect.ShowMe(Me) = False Then
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            Exit Sub
        End If
    End If
    Me.Refresh
    mstr病人类型 = ""
    txtPatient.ForeColor = &HFF0000
    
    '刷卡完毕或输入号码后回车
    If blnCard And Len(Me.txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Me.txtPatient.Text <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        
        '读取病人信息
        Call ClearWinInfor(True)
        
        If IDKind.GetCurCard.名称 Like "IC卡*" And IDKind.GetCurCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        
        If Not GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
            If blnCancel Then '取消输入
                Call zlControl.TxtSelAll(txtPatient): txtPatient.SetFocus: Exit Sub
            End If
            stbThis.Panels(1).Text = "未找到该病人，请检查输入内容!"
            If blnCard = True Then
                txtPatient.PasswordChar = "": txtPatient.Text = "": txtPatient.IMEMode = 0
            Else
                txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
            End If
            Set mrsInfo = New ADODB.Recordset
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Else
            '设置病人费用信息
            mlng病人ID = Val(Nvl(mrsInfo!病人ID, 0))
            mlng主页ID = Val(Nvl(mrsInfo!主页ID, 0))
            
            Call ClearWinInfor
            If mrsInfo!当前科室ID <> 0 Then
                lblBed.Caption = "床号：" & IIf(mrsInfo!床号 = 0, "家庭", mrsInfo!床号)
            End If
            
            lblNO.Caption = "住院号：" & IIf(mrsInfo!住院号 = 0, "", mrsInfo!住院号)
            lblDept.Caption = "科室：" & GET部门名称(mrsInfo!当前科室ID)
            
            lblType.Caption = "费别等级：" & mrsInfo!费别
'            lbl担保人.Caption = lbl担保人.Tag & mrsInfo!担保人
'            lbl担保金额.Caption = lbl担保金额.Tag & mrsInfo!担保额
'            chk担保temp.Value = mrsInfo!担保性质
            
            txtPatient.PasswordChar = "": txtPatient.IMEMode = 0
            txtPatient.Text = mrsInfo!姓名
            txtPatient.Tag = mrsInfo!病人ID
            '-----------------------------------------------------------------------------------------
            lblSex.Caption = "性别：" & IIf(IsNull(mrsInfo!性别), "", mrsInfo!性别)
            lblAge.Caption = "年龄：" & IIf(IsNull(mrsInfo!年龄), "", mrsInfo!年龄)
'            lbl家庭地址.Caption = lbl家庭地址.Tag & Nvl(mrsInfo!家庭地址)
            lblCur.Caption = "医疗付款方式：" & Nvl(mrsInfo!医疗付款方式)
            dblMoney = 0
            For lngRow = 1 To mshList.Rows - 1
                 dblMoney = Format(dblMoney + Val(mshList.TextMatrix(lngRow, GetColNumList("金额"))), "#0.00;-#0.00;0.00")
            Next
            lblDeposit.Caption = "预交总额：" & IIf(dblMoney = 0, "", dblMoney)
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If mrsInfo Is Nothing Then
            cmdAdd.Enabled = False
        ElseIf mrsInfo.State = adStateClosed Then
            cmdAdd.Enabled = False
        Else
            cmdAdd.Enabled = True
        End If
    End If
End Sub

Private Function GetPatient(ByVal objCard As zlOneCardComLib.Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '功能：读取病人信息
    '参数：strInput=[刷卡]|[A病人ID]|[B住院号]
    '说明：
    '     自动识别病人在院状态,读出(病人ID,主页ID,姓名,性别,年龄,住院号,床号,在院标志)
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng卡类别ID As Long, bln存在帐户 As Boolean, lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    Dim colPati As Collection
    Dim colPage As Collection
    Dim colItem As Collection
    Dim colOtherFind As Collection
    Dim strFields As String
    Dim strKeyName As String
    Dim strKeyValue As String
    
    blnCancel = False
    strWhere = ""
      
    If (blnCard And objCard.名称 Like "姓名*") _
        And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then   '刷卡或缺省的卡
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If mobjOneCardCom.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        lng病人ID = Val(Mid(strInput, 2))
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        strKeyName = "住院号": strKeyValue = Mid(strInput, 2)
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        strKeyName = "门诊号": strKeyValue = Mid(strInput, 2)
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名"
                If Not mblnSeekName Then
                    MsgBox "请刷卡或输入[-病人ID]、[+住院号]、[*门诊号]等方式提取病人的信息。", vbInformation, gstrSysName
                    txtPatient.Text = "": txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                Else
                    If zlCommFun.IsCharChinese(strInput) Then
                        If Len(strInput) < 2 Then
                            MsgBox "输入内容""" & strInput & """长度不足2位。", vbInformation, gstrSysName
                            txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                        End If
                    Else
                        If Len(strInput) < 4 Then
                            MsgBox "输入内容""" & strInput & """长度不足4位。", vbInformation, gstrSysName
                            txtPatient.SetFocus: Set mrsInfo = Nothing: Exit Function
                        End If
                    End If
                    On Error GoTo errH
                    Set colOtherFind = New Collection
                    colOtherFind.Add Array("姓名", strInput & "%")
                    colOtherFind.Add Array("查询住院状态", 1)
                    If zl_PatiSvr_GetPatiInfo(0, colOtherFind, colPati, 2) Then
                        strSQL = ""
                        For Each colItem In colPati
                            If strSQL <> "" Then strSQL = strSQL & "Union All " & vbNewLine
                            strSQL = strSQL & "Select " & colItem("_pati_id") & " As 病人ID," & _
                                    Nvl(colItem("_pati_pageid"), "-NULL") & " As 主页ID," & _
                                    "'" & colItem("_pati_name") & "' As 姓名," & _
                                    "'" & colItem("_pati_sex") & "' As 性别," & _
                                    "'" & colItem("_pati_age") & "' As 年龄," & _
                                    Nvl(colItem("_inpatient_num"), "-NULL") & " As 住院号," & _
                                    "'" & colItem("_pati_dept_name") & "' As 科室," & _
                                    "'" & colItem("_pati_birthdate") & "' As 出生日期," & _
                                    "'" & colItem("_pati_idcard") & "' As 身份证号," & _
                                    "'" & colItem("_pat_home_addr") & "' As 家庭地址," & _
                                    "'" & colItem("_adta_time") & "' As 入院时间 From Dual "
                        Next
                    End If
                    If strSQL <> "" Then
                        strPati = "Select A.病人ID As ID,A.* From (" & strSQL & ") A Order by A.入院时间 DESC,A.姓名"
                        vRect = zlControl.GetControlRect(txtPatient.hWnd)
                        Set rsTmp = zldatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True)
                    End If
                    If Not rsTmp Is Nothing Then
                        lng病人ID = rsTmp!病人ID
                    Else
                        Set mrsInfo = New ADODB.Recordset: Exit Function
                    End If
                End If
            Case "医保号"
                strKeyName = "医保号": strKeyValue = UCase(strInput)
            Case "IC卡号"
                strInput = UCase(strInput)
                If mobjOneCardCom.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0) And objCard.系统
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                 strKeyName = "门诊号": strKeyValue = strInput
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strKeyName = "住院号": strKeyValue = strInput
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    bln存在帐户 = objCard.是否存在帐户
                    If mobjOneCardCom.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If mobjOneCardCom.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                blnHavePassWord = True
        End Select
    End If
    
    If strKeyName <> "" And strKeyValue <> "" Then
        Set colOtherFind = New Collection
        colOtherFind.Add Array(strKeyName, strKeyValue)
        If Not zl_PatiSvr_GetPatiID(Nothing, colOtherFind, colPati) Then GoTo NotFoundPati:
        lng病人ID = colPati(1)("_pati_id")
    End If
    
    If lng病人ID <= 0 Then GoTo NotFoundPati:
    '查询病人信息
    If Not zl_PatiSvr_GetPatiInfo(lng病人ID, Nothing, colPati, 2) Then GoTo NotFoundPati:
    Set colPati = colPati(1)
             
    '查询病案主页
    strSQL = "{""input"":{""query_type"":1,""pati_pageids"":""" & lng病人ID & ":" & colPati("_pati_pageid") & """,""is_lastpage"":1,""is_babyinfo"":0,""is_transdeptinfo"":1,""is_ex"":1,""is_bed"":0}}"
    Call CallService("zl_CIsSvr_GetPatiPageInfo", strSQL, , Me.Caption, , False, , , , True)
    Set colPage = gobjService.GetJsonListValue("output.page_list[0]")
    If colPage Is Nothing Or colPati Is Nothing Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    
    strFields = "病人ID|adBigInt|18,主页ID|adBigInt|18,卡验证码||50,姓名||100,性别||4,年龄||20,病人类型||50," & _
            "当前科室ID|adBigInt|18,床号||10,住院号|adBigInt|18," & _
            "费别||10,担保人||100,担保额||20,担保性质|adSingle|1,医疗付款方式||20"
    Set mrsInfo = InitRS(strFields)
    With mrsInfo
        .AddNew
        !病人ID = colPati("_pati_id"): !主页ID = colPati("_pati_pageid"): !卡验证码 = colPati("_card_captcha")
        !姓名 = colPati("_pati_name"): !性别 = colPati("_pati_sex"): !年龄 = colPati("_pati_age")
        !病人类型 = colPati("_pati_type"): !当前科室ID = colPage("_pati_dept_id"): !床号 = Nvl(colPage("_pati_bed"))
        !住院号 = Nvl(colPage("_inpatient_num")): !费别 = colPage("_fee_category")
        !医疗付款方式 = colPage("_mdlpay_mode_name")
    End With
    '获取担保信息
    strSQL = "{""input"":{""pati_pageid"":0,""pati_id"":" & lng病人ID & ",""surety_prop"":1}}"
    Call CallService("Zl_Exsesvr_Getpatisurety", strSQL, , , , False, , , , True)
    If Val(gobjService.GetJsonNodeValue("output.guarantee_money") & "") <> 0 Then
        mrsInfo!担保性质 = Val(gobjService.GetJsonNodeValue("output.surety_prop") & "")
        mrsInfo!担保人 = gobjService.GetJsonNodeValue("output.entsurety") & ""
        mrsInfo!担保额 = Format(Val(gobjService.GetJsonNodeValue("output.guarantee_money") & ""), "0.00")
    End If
    mrsInfo.Update
    '需要处理其他
    If mblnCheckPass And (blnCard Or blnICCard) Then
        If Not blnHavePassWord Then
            strPassWord = Nvl(mrsInfo!卡验证码)
        End If
        If strPassWord <> "" Then
            If CreatePublicExpense() Then
                If gobjPublicExpense.zlVerifyPassWord(Me, strPassWord, mrsInfo!姓名, mrsInfo!性别, mrsInfo!年龄) = False Then
                     Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            End If
        End If
    End If
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function


Private Sub txtPatient_Change()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "" And Me.ActiveControl Is txtPatient)
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_GotFocus()
    txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
    If Not mobjIDCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjIDCard.SetEnabled(True)
    If Not mobjICCard Is Nothing And txtPatient.Text = "" And Not txtPatient.Locked Then Call mobjICCard.SetEnabled(True)
    txtPatient.Tag = ""
    Call IDKind.SetAutoReadCard(txtPatient.Text = "")
End Sub

Private Sub txtPatient_LostFocus()
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    '问题27379 by lesfeng 2010-01-18
    If mrsInfo.State = 1 Then
        mstr病人类型 = IIf(IsNull(mrsInfo!病人类型), "", mrsInfo!病人类型)
    End If
    If mstr病人类型 = "" Then
        If mrsInfo.State = 1 Then
            If GetOutPatient(mrsInfo!病人ID) Then
                txtPatient.ForeColor = vbRed
            Else
                txtPatient.ForeColor = &HFF0000
            End If
        Else
            txtPatient.ForeColor = &HFF0000
        End If
    Else
        txtPatient.ForeColor = ReadPatiColor(mstr病人类型, True)
    End If
End Sub

Private Function GetOutPatient(ByVal lngID As Long) As Boolean
'功能：判断门诊病人是否属于医保
    Dim int险类 As Integer
    Dim colPati As Collection
    
    GetOutPatient = False
    On Error GoTo errH
 
    If zl_PatiSvr_GetPatiInfo(lngID, Nothing, colPati) Then
        int险类 = Val(colPati(1)("_insurance_type") & "")
        GetOutPatient = int险类 <> 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        '问题27554 by lesfeng 2010-01-19 lngTXTProc 修改为glngTXTProc
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtReason_GotFocus()
    zlControl.TxtSelAll txtReason
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("'|?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txtReason, KeyAscii
    End If
End Sub

Private Sub txtReason_LostFocus()
    If gstrIme <> "不自动开启" Then Call OS.OpenImeByName
End Sub

Private Sub txt担保额_GotFocus()
    zlControl.TxtSelAll txt担保额
End Sub

Private Sub txt担保额_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
        If KeyAscii = vbKeyReturn Then
            chkUnlimit.TabStop = (txt担保额.Text = "")
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            KeyAscii = 0
        End If
    ElseIf KeyAscii = Asc(".") And InStr(txt担保额.Text, ".") > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt担保额_LostFocus()
    If IsNumeric(txt担保额.Text) Then
        If txt担保额.Text = "999999999" Then
            stbThis.Panels(1).Text = "不允许输入该值，该值表示无限担保．"
            If txt担保额.Enabled Then txt担保额.SetFocus
        Else
            txt担保额.Text = Format(txt担保额.Text, "0.00")
        End If
    Else
        txt担保额.Text = ""
    End If
    
    Call zlCommFun.OpenIme
End Sub

Private Sub txt担保人_GotFocus()
    zlControl.TxtSelAll txt担保人
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt担保人_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        CheckInputLen txt担保人, KeyAscii
    End If
End Sub

Private Sub txt担保人_LostFocus()
    If gstrIme <> "不自动开启" Then Call OS.OpenImeByName
End Sub

Private Function To_Date(ByVal dat日期 As Date) As String
'功能:将入参中的日期传换成ORACLE需要的日期格式串
    To_Date = "To_Date('" & Format(dat日期, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Private Sub zlCardSquareObject(Optional blnClosed As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '入参:blnClosed:关闭对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
   
    '只有:执行或退费时,才可能管结算卡的
    If blnClosed Then
       If Not mobjOneCardCom Is Nothing Then
            Call mobjOneCardCom.CloseWindows
            Set mobjOneCardCom = Nothing
        End If
        Exit Sub
    End If
    '创建对象
    '刘兴洪:增加结算卡的结算:执行或退费时
    Err = 0: On Error Resume Next
    Set mobjOneCardCom = CreateObject("zlOneCardComLib.clsOneCardComLib")
    If Err <> 0 Then
        Err = 0: On Error GoTo 0:      Exit Sub
    End If
    
    '安装了结算卡的部件
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    '功能:zlInitComponents (初始化接口部件)
    '    ByVal frmMain As Object, _
    '        ByVal lngModule As Long, ByVal lngSys As Long, ByVal strDBUser As String, _
    '        ByVal cnOracle As ADODB.Connection, _
    '        Optional blnDeviceSet As Boolean = False, _
    '        Optional strExpand As String
    '出参:
    '返回:   True:调用成功,False:调用失败
    '编制:刘兴洪
    '日期:2009-12-15 15:16:22
    'HIS调用说明.
    '   1.进入门诊收费时调用本接口
    '   2.进入住院结帐时调用本接口
    '   3.进入预交款时
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    If mobjOneCardCom.zlInitComponents(Me, mlngModul, glngSys, gstrDBUser, gcnOracle, False, strExpend) = False Then
         '初始部件不成功,则作为不存在处理
         Exit Sub
    End If
End Sub

Private Function GetJsonSurety(ByVal bytFunc As Byte, ByVal lngPatiid As Long, ByVal lngPageID As Long, Optional ByVal strSurety As String, _
   Optional ByVal dblAmount As Double, Optional ByVal bytType As Byte, Optional ByVal strReason As String, Optional ByVal strDueTime As String, _
   Optional ByVal strCreateTime As String) As String
'功能:获取担保JSON串
'参数:
'      bytFunc          功能ID 1-新增;2-更新;3-删除
'      lngPatiId        病人id
'      lngPageId        主页ID
'      strSurety        担保人
'      dblAmount        担保额
'      bytType          担保性质
'      strReason        担保原因
'      strDueTime       到期时间   格式化 "yyyy-MM-dd HH:mm:ss"
'      strCreateTime    登记时间   更新或删除时传入
    Dim strIn As String
    
    strIn = GetNode("func_id", bytFunc, True)
    strIn = strIn & GetNode("pati_id", lngPatiid)
    strIn = strIn & GetNode("pati_pageid", lngPageID)
    If bytFunc <> 3 Then
        strIn = strIn & GetNode("guarantor", strSurety)
        strIn = strIn & GetNode("garnt_amount", dblAmount)
        strIn = strIn & GetNode("garnt_prop", bytType)
        strIn = strIn & GetNode("garnt_reason", strReason)
        If strDueTime <> "" Then
            strIn = strIn & GetNode("due_time", strDueTime)
        End If
    End If

    strIn = strIn & GetNode("operator_code", UserInfo.编号)
    strIn = strIn & GetNode("operator_name", UserInfo.姓名)

    If bytFunc <> 1 Then
        strIn = strIn & GetNode("create_time", strCreateTime)
    End If
    strIn = GetNode("input", "{" & strIn & "}", True, 1)
    GetJsonSurety = strIn
End Function

