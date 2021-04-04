VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "ZLIDKIND.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChargeFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab sst1 
      Height          =   4452
      Left            =   120
      TabIndex        =   22
      Top             =   0
      Width           =   6276
      _ExtentX        =   11060
      _ExtentY        =   7858
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "基本(&0)"
      TabPicture(0)   =   "frmChargeFilter.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "收费项目(&1)"
      TabPicture(1)   =   "frmChargeFilter.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtInput(0)"
      Tab(1).Control(1)=   "ListFeeItem(0)"
      Tab(1).Control(2)=   "tlbOpt(0)"
      Tab(1).Control(3)=   "ils16"
      Tab(1).Control(4)=   "lbl收入项目(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "条件"
      TabPicture(2)   =   "frmChargeFilter.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmAuditStatus"
      Tab(2).Control(1)=   "cboAudit"
      Tab(2).Control(2)=   "cboApply"
      Tab(2).Control(3)=   "chkDate(1)"
      Tab(2).Control(4)=   "chkDate(0)"
      Tab(2).Control(5)=   "dtpApplyB"
      Tab(2).Control(6)=   "dtpAuditB"
      Tab(2).Control(7)=   "dtpApplyE"
      Tab(2).Control(8)=   "dtpAuditE"
      Tab(2).Control(9)=   "Label13"
      Tab(2).Control(10)=   "Label12"
      Tab(2).Control(11)=   "lblAuditDate"
      Tab(2).Control(12)=   "lblAudit"
      Tab(2).Control(13)=   "lblApplyDate"
      Tab(2).Control(14)=   "lblApply"
      Tab(2).ControlCount=   15
      Begin VB.Frame frmAuditStatus 
         Caption         =   "审核状态"
         Height          =   615
         Left            =   -74085
         TabIndex        =   61
         Top             =   1320
         Width           =   3225
         Begin VB.CheckBox chkAudit 
            Caption         =   "拒绝"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   64
            Top             =   240
            Width           =   915
         End
         Begin VB.CheckBox chkAudit 
            Caption         =   "申请"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   63
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkAudit 
            Caption         =   "通过"
            Height          =   255
            Index           =   1
            Left            =   1215
            TabIndex        =   62
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.ComboBox cboAudit 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74085
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2070
         Width           =   2055
      End
      Begin VB.ComboBox cboApply 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -74085
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   2055
      End
      Begin VB.CheckBox chkDate 
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   -69480
         TabIndex        =   47
         Top             =   2505
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox chkDate 
         Height          =   255
         Index           =   0
         Left            =   -69480
         TabIndex        =   46
         Top             =   923
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   18
         ToolTipText     =   "最多匹配100项搜索结果"
         Top             =   540
         Width           =   2160
      End
      Begin VB.ListBox ListFeeItem 
         Height          =   3210
         Index           =   0
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   19
         ToolTipText     =   "Ctrl+A全选,Ctrl+C全消,如果一个都未选则表示不限制"
         Top             =   900
         Width           =   4725
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   3972
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   5925
         Begin VB.TextBox txt标识号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3825
            MaxLength       =   18
            TabIndex        =   9
            Top             =   1350
            Width           =   1830
         End
         Begin VB.TextBox txt姓名 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   975
            MaxLength       =   64
            TabIndex        =   8
            Top             =   1350
            Width           =   1830
         End
         Begin VB.TextBox txtPatient 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   17
            Top             =   3240
            Width           =   3495
         End
         Begin VB.Frame fra来源 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   120
            TabIndex        =   52
            Top             =   3720
            Width           =   5535
            Begin VB.OptionButton opt病人 
               Caption         =   "门诊病人"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   55
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt病人 
               Caption         =   "住院病人"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   2370
               TabIndex        =   54
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt病人 
               Caption         =   "门诊病人和住院病人"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   2
               Left            =   3675
               TabIndex        =   53
               Top             =   0
               Value           =   -1  'True
               Width           =   1935
            End
            Begin VB.Label lblFil 
               Alignment       =   1  'Right Justify
               Caption         =   "病人来源"
               Height          =   180
               Left            =   0
               TabIndex        =   56
               Top             =   0
               Width           =   930
            End
         End
         Begin VB.TextBox txtFactEnd 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3828
            TabIndex        =   13
            Top             =   2100
            Width           =   1830
         End
         Begin VB.TextBox txtFactBegin 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   972
            TabIndex        =   12
            Top             =   2100
            Width           =   1830
         End
         Begin VB.TextBox txtNoEnd 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3828
            MaxLength       =   8
            TabIndex        =   11
            Top             =   1728
            Width           =   1830
         End
         Begin VB.TextBox txtNOBegin 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   972
            MaxLength       =   8
            TabIndex        =   10
            Top             =   1728
            Width           =   1830
         End
         Begin VB.ComboBox cbo操作员 
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   972
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   2832
            Width           =   1830
         End
         Begin VB.CheckBox chk退费 
            Caption         =   "退费记录"
            Height          =   210
            Left            =   4695
            TabIndex        =   5
            Top             =   555
            Width           =   1020
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            Left            =   972
            TabIndex        =   14
            Text            =   "cbo科室"
            Top             =   2472
            Width           =   1830
         End
         Begin VB.CheckBox chk收费 
            Caption         =   "收费记录"
            Height          =   210
            Left            =   4695
            TabIndex        =   3
            Top             =   270
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk医保 
            Caption         =   "医保收费"
            Height          =   195
            Left            =   3480
            TabIndex        =   2
            Top             =   278
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk普通 
            Caption         =   "普通收费"
            Height          =   195
            Left            =   3480
            TabIndex        =   4
            ToolTipText     =   "指不包括医保收费的其它所有收费"
            Top             =   563
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.ComboBox cbo付款方式 
            Height          =   276
            IMEMode         =   3  'DISABLE
            Left            =   3825
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1020
            Width           =   1830
         End
         Begin VB.ComboBox cbo费别 
            Height          =   276
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1020
            Width           =   1830
         End
         Begin VB.TextBox txt开单人 
            Height          =   300
            Left            =   3828
            TabIndex        =   15
            Top             =   2472
            Width           =   1830
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   975
            TabIndex        =   1
            Top             =   570
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   173342723
            CurrentDate     =   36588
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   975
            TabIndex        =   0
            Top             =   150
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   173342723
            CurrentDate     =   36588
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   960
            TabIndex        =   58
            Top             =   3240
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   529
            Appearance      =   2
            IDKindStr       =   "医|医保号|0|0|0|0|0|;身|身份证号|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|;就|就诊卡|0|0|0|0|0|"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "宋体"
            IDKind          =   -1
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            BackColor       =   -2147483633
         End
         Begin VB.Label lbl标识号 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊号"
            Height          =   180
            Left            =   3024
            TabIndex        =   60
            Top             =   1410
            Width           =   768
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "姓名"
            Height          =   180
            Left            =   540
            TabIndex        =   59
            Top             =   1410
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "身份识别"
            Height          =   180
            Left            =   120
            TabIndex        =   57
            Top             =   3312
            Width           =   720
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "票据号"
            Height          =   180
            Left            =   360
            TabIndex        =   35
            Top             =   2160
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   3288
            TabIndex        =   34
            Top             =   2160
            Width           =   180
         End
         Begin VB.Label lbl操作员 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "操作员"
            Height          =   180
            Left            =   360
            TabIndex        =   33
            Top             =   2892
            Width           =   540
         End
         Begin VB.Label lbl科室 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开单科室"
            Height          =   180
            Left            =   180
            TabIndex        =   32
            Top             =   2532
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "单据号"
            Height          =   180
            Left            =   360
            TabIndex        =   31
            Top             =   1788
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   3288
            TabIndex        =   30
            Top             =   1788
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结束时间"
            Height          =   180
            Left            =   180
            TabIndex        =   29
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开始时间"
            Height          =   180
            Left            =   180
            TabIndex        =   28
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lbl费别 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "费别"
            Height          =   180
            Left            =   540
            TabIndex        =   27
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "医疗付款"
            Height          =   180
            Left            =   3075
            TabIndex        =   26
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label lbl开单人 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开单人"
            Height          =   180
            Left            =   3252
            TabIndex        =   25
            Top             =   2532
            Width           =   540
         End
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   600
         Index           =   0
         Left            =   -74760
         TabIndex        =   36
         Top             =   1140
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1058
         ButtonWidth     =   1614
         ButtonHeight    =   1058
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "移除(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "移除当前选择的列表项"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "清空(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "清空列表项目"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "保存(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "保存选择的列表项目"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   -69960
         Top             =   660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":0054
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":03EE
               Key             =   "Insert"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":0788
               Key             =   "Clear"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmChargeFilter.frx":0B22
               Key             =   "Delete"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpApplyB 
         Height          =   300
         Left            =   -74085
         TabIndex        =   39
         Top             =   900
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpAuditB 
         Height          =   300
         Left            =   -74085
         TabIndex        =   44
         Top             =   2475
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpApplyE 
         Height          =   300
         Left            =   -71640
         TabIndex        =   40
         Top             =   900
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpAuditE 
         Height          =   300
         Left            =   -71640
         TabIndex        =   45
         Top             =   2475
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   173342723
         CurrentDate     =   36588
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   -71880
         TabIndex        =   51
         Top             =   2535
         Width           =   180
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   -71880
         TabIndex        =   50
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblAuditDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核时间"
         Height          =   180
         Left            =   -74880
         TabIndex        =   49
         Top             =   2535
         Width           =   720
      End
      Begin VB.Label lblAudit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   -74700
         TabIndex        =   48
         Top             =   2130
         Width           =   540
      End
      Begin VB.Label lblApplyDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请时间"
         Height          =   180
         Left            =   -74880
         TabIndex        =   42
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lblApply 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "申请人"
         Height          =   180
         Left            =   -74700
         TabIndex        =   41
         Top             =   555
         Width           =   540
      End
      Begin VB.Label lbl收入项目 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费项目(&F)"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   37
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   6525
      TabIndex        =   23
      Top             =   1650
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6525
      TabIndex        =   21
      Top             =   690
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6525
      TabIndex        =   20
      Top             =   270
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrPrivs As String 'IN
Public mblnApply As Boolean '查看退费申请单
Public mstrFilter As String 'IN/Out
Public mblnDateMoved As Boolean 'Out
Public mstrFeeItems As String 'out

Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnNotClick As Boolean
Private mblnUnChange  As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mrsDept As ADODB.Recordset

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
'    If KeyAscii >= 32 Then
'        lngIdx = zlControl.CboMatchIndex(cbo科室.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo科室.ListCount > 0 Then lngIdx = 0
'        cbo科室.ListIndex = lngIdx
'    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Set mrsDept = GetDepartments("'临床','手术'", gint病人来源 & ",3")
    If zlSelectDept(Me, 1120, cbo科室, mrsDept, cbo科室.Text, True, "所有科室") = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub chkAudit_Click(Index As Integer)
    If Not Visible Then Exit Sub
    
    If chkAudit(0).Value = 0 And chkAudit(1).Value = 0 And chkAudit(2).Value = 0 Then
        chkAudit(Index).Value = 1   '递归调用
        Exit Sub
    End If
    
    If Index = 1 Or Index = 2 Then
        cboAudit.Enabled = chkAudit(1).Value = 1 Or chkAudit(2).Value = 1
        dtpAuditB.Enabled = cboAudit.Enabled
        dtpAuditE.Enabled = cboAudit.Enabled
        chkDate(1).Enabled = cboAudit.Enabled
        
        If cboAudit.Enabled Then chkDate(1).Value = 1
    End If
End Sub

Private Sub chkDate_Click(Index As Integer)
    If Not Visible Then Exit Sub
    
    If chkDate(0).Value = 0 And (chkDate(1).Value = 0 Or chkDate(1).Enabled = False) Then
        If chkDate(1).Enabled = False Then
            chkDate(0).Value = 1
            Exit Sub
        Else
            chkDate(1 - Index).Value = 1 '递归调用
        End If
    End If
        
    If Index = 0 Then
        dtpApplyB.Enabled = chkDate(Index).Value = 1
        dtpApplyE.Enabled = dtpApplyB.Enabled
    Else
        dtpAuditB.Enabled = chkDate(Index).Value = 1
        dtpAuditE.Enabled = dtpAuditB.Enabled
    End If
End Sub

Private Sub chk普通_Click()
    If chk医保.Value = 0 And chk普通.Value = 0 Then
        chk普通.Value = 1
    End If
End Sub

Private Sub chk退费_Click()
    If chk收费.Value = 0 And chk退费.Value = 0 Then
        chk退费.Value = 1
    End If
End Sub

Private Sub chk收费_Click()
    If chk收费.Value = 0 And chk退费.Value = 0 Then
        chk收费.Value = 1
    End If
End Sub

Private Sub chk医保_Click()
    If chk医保.Value = 0 And chk普通.Value = 0 Then
        chk医保.Value = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If gbln退费申请模式 And mblnApply Then
        If chkDate(0).Value = 0 And chkDate(1).Value = 0 Then
            MsgBox "请选择时间范围！", vbInformation, gstrSysName
            chkDate(0).SetFocus: Exit Sub
        End If
        fra来源.Visible = False
    Else
        If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
            If txtNoEnd.Text < txtNOBegin.Text Then
                MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
                txtNoEnd.SetFocus: Exit Sub
            End If
        End If
        fra来源.Visible = True
        If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
            If txtFactEnd.Text < txtFactBegin.Text Then
                MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
                txtFactEnd.SetFocus: Exit Sub
            End If
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    If gbln退费申请模式 And mblnApply Then
        fra来源.Visible = False '33789
        cboApply.SetFocus
    Else
        fra来源.Visible = True '33789
        dtpBegin.SetFocus
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If sst1.Tab = 1 Then
            txtInput(sst1.Tab - 1).SetFocus
        Else
            KeyCode = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf Shift = 2 Then
        If sst1.Tab = 1 Then
            Dim i As Integer, Index As Integer
            
            Index = sst1.Tab - 1
            If UCase(Chr(KeyCode)) = "A" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = True
                Next
            ElseIf UCase(Chr(KeyCode)) = "C" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = False
                Next
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
End Sub



Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngOldID As Long, strListFeeItem As String
    Dim Curdate As Date, Index As Integer, arrItem As Variant
    
    On Error GoTo errH
    gblnOK = False
    
    Curdate = zlDatabase.Currentdate
    '47928
    InitIDKind
    
    If gbln退费申请模式 And mblnApply Then
        sst1.TabVisible(0) = False
        sst1.TabVisible(1) = False
        sst1.TabVisible(2) = True
        
        cboApply.Clear
        
        cboApply.AddItem "所有申请人"
        Set rsTmp = GetPersonnel("门诊收费员", True)
        For i = 1 To rsTmp.RecordCount
            cboApply.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            
            If rsTmp!ID = UserInfo.ID Then cboApply.ListIndex = cboApply.NewIndex
            rsTmp.MoveNext
        Next
        
        cboAudit.AddItem "所有审核人"
        strSQL = "Select Distinct D.ID, D.简码, D.姓名" & vbNewLine & _
                "From 人员表 D,上机人员表 C, Zluserroles B, zlRoleGrant A" & vbNewLine & _
                "Where A.系统 = [1] And A.序号 = 1121 And A.功能 = '退费审核' And A.角色 = B.角色 And B.用户 = C.用户名 And C.人员id = D.ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys)
        For i = 1 To rsTmp.RecordCount
            cboAudit.AddItem rsTmp!简码 & "-" & rsTmp!姓名
            
            If rsTmp!ID = UserInfo.ID Then cboAudit.ListIndex = cboAudit.NewIndex
            rsTmp.MoveNext
        Next
        
        If cboApply.ListIndex = -1 And cboApply.ListCount > 0 Then cboApply.ListIndex = 0
        If cboAudit.ListIndex = -1 And cboAudit.ListCount > 0 Then cboAudit.ListIndex = 0
        
        
        dtpApplyB.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpApplyE.MaxDate = dtpApplyB.MaxDate
        
        dtpAuditB.MaxDate = dtpApplyB.MaxDate
        dtpAuditE.MaxDate = dtpApplyB.MaxDate
        
        dtpApplyB.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpApplyE.Value = dtpApplyB.MaxDate
        
        dtpAuditB.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpAuditE.Value = dtpApplyB.MaxDate
        
        chkAudit(0).Value = 1
        chkAudit(1).Value = 0
        chkAudit(2).Value = 0
        
        chkDate(0).Value = 1
        chkDate(1).Value = 1
        
    Else
        sst1.TabVisible(0) = True
        sst1.TabVisible(1) = True
        sst1.TabVisible(2) = False
        
        If glngSys Like "8??" Then
            lbl科室.Visible = False
            cbo科室.Visible = False
        End If
        
        txtNOBegin.Text = ""
        txtNoEnd.Text = ""
        txtFactBegin.Text = ""
        txtFactEnd.Text = ""
        txtPatient.Text = ""
        chk收费.Value = 1
        chk退费.Value = 0
        
        chk医保.Value = 1
        chk普通.Value = 1
        
        lbl标识号.Caption = IIf(gint病人来源 = 1, "门诊号", "住院号")
        
        dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
        dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = dtpBegin.MaxDate
        
        If InStr(mstrPrivs, "显示开单人") = 0 Then
            lbl开单人.Visible = False
            txt开单人.Visible = False
        Else
            lbl开单人.Visible = True
            txt开单人.Visible = True
        End If
        
        '操作员
        cbo操作员.Clear
        If InStr(mstrPrivs, "所有操作员") > 0 Then
            cbo操作员.AddItem "所有收费员"
            Set rsTmp = GetPersonnel("门诊收费员", True)
            For i = 1 To rsTmp.RecordCount
                cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
                cbo操作员.ItemData(cbo操作员.NewIndex) = rsTmp!ID
                If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
                rsTmp.MoveNext
            Next
        Else
            cbo操作员.AddItem UserInfo.简码 & "-" & UserInfo.姓名
            cbo操作员.ItemData(cbo操作员.NewIndex) = UserInfo.ID
        End If
        If cbo操作员.ListIndex = -1 And cbo操作员.ListCount > 0 Then cbo操作员.ListIndex = 0
        
        '开单科室'@
        cbo科室.Clear
        cbo科室.AddItem "所有科室"
        cbo科室.ListIndex = 0
        Set mrsDept = GetDepartments("'临床','手术'", gint病人来源 & ",3")
        For i = 1 To mrsDept.RecordCount
            If lngOldID <> mrsDept!ID Then
                cbo科室.AddItem mrsDept!编码 & "-" & mrsDept!名称
                cbo科室.ItemData(cbo科室.NewIndex) = mrsDept!ID
                lngOldID = mrsDept!ID
            End If
            mrsDept.MoveNext
        Next
        
        cbo费别.Clear
        cbo费别.AddItem "所有费别"
        cbo费别.ListIndex = 0
        
        strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Where Nvl(服务对象,3) IN(1,3) Order by 编码"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
                rsTmp.MoveNext
            Next
        End If
        
        '医疗付款方式,默认为空表示所有
        cbo付款方式.Clear
        cbo付款方式.AddItem ""
        cbo付款方式.ListIndex = 0
        strSQL = "Select 编码,名称,Nvl(缺省标志,0) as 缺省 From 医疗付款方式 Order by 编码"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo付款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
        
        If InStr(1, mstrPrivs, "明细项目过滤") = 0 Then
            sst1.TabVisible(1) = False
        Else
            For Index = 0 To 0  '将来可能会加收入项目条件
                strListFeeItem = ""
                ListFeeItem(Index).Clear
                
                Call GetRegisterItem(g私有模块, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "收费项目列表", "收入项目列表"), strListFeeItem)
                If strListFeeItem <> "" Then
                    arrItem = Split(strListFeeItem, ";")
                    
                    For i = 0 To UBound(arrItem)
                        ListFeeItem(Index).AddItem Split(arrItem(i), ",")(0)
                        ListFeeItem(Index).ItemData(ListFeeItem(Index).NewIndex) = Val(Split(arrItem(i), ",")(1))
                        ListFeeItem(Index).Selected(ListFeeItem(Index).NewIndex) = IIf(Val(Split(arrItem(i), ",")(2)) = 1, True, False)
                    Next
                End If
            Next
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    If gbln退费申请模式 And mblnApply Then
        sst1.Height = dtpAuditB.Top + dtpAuditB.Height * 2
        Me.Height = sst1.Height + dtpAuditB.Height * 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mblnApply = False
    mstrFilter = ""
    mstrFeeItems = ""
    If Not mrsDept Is Nothing Then Set mrsDept = Nothing
End Sub

Private Sub opt病人_Click(Index As Integer)
    lbl标识号.Caption = IIf(opt病人(0).Value, "门诊号", IIf(opt病人(1).Value, "住院号", "门诊/住院号"))
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    If Me.Visible = False Then Exit Sub
    
    If gbln退费申请模式 And mblnApply Then
        If cboApply.Visible And cboApply.Enabled Then Call cboApply.SetFocus
    Else
        If sst1.Tab = 0 Then
            txtPatient.SetFocus
        Else
            txtInput(0).SetFocus
        End If
    End If
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub


Private Sub tlbOpt_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Delete"
            If ListFeeItem(Index).ListIndex >= 0 Then
                Call ListFeeItem(Index).RemoveItem(ListFeeItem(Index).ListIndex)
            End If
        Case "Clear"
            ListFeeItem(Index).Clear
        Case "Save"
            Dim strTmp As String, i As Long
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    strTmp = strTmp & ";" & .List(i) & "," & .ItemData(i) & "," & IIf(.Selected(i), 1, 0)
                Next
            End With
            strTmp = Mid(strTmp, 2)
            Call SaveRegisterItem(g私有模块, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "收费项目列表", "收入项目列表"), strTmp)
    End Select
End Sub

Private Sub ListFeeItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If ListFeeItem(Index).ListIndex >= 0 Then
            ListFeeItem(Index).RemoveItem ListFeeItem(Index).ListIndex
        End If
    End If
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInput(Index))
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strSQL As String, strInput As String, strMatch As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strInput = UCase(Trim(txtInput(Index).Text))
        If strInput = "" Then Exit Sub
        strMatch = IIf(Len(strInput) < 3, "", gstrLike)
        
        If Index = 0 Then
        '收费项目
            If zlCommFun.IsNumOrChar(strInput) Then
                strIF = " And (A.编码 like [1] Or B.简码 like [1] And B.码类 in(3," & gbytCode + 1 & "))"
            Else
                strIF = " And B.名称 like [1]"
            End If
            strSQL = "Select Distinct A.ID, A.编码, B.名称 ,A.规格, A.产地, A.计算单位 " & _
                  " From 收费项目目录 A,收费项目别名 B Where A.id=B.收费细目ID " & strIF & _
                  " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
                  " And rownum<101 Order by 名称"
        Else
        '收入项目
            If zlCommFun.IsNumOrChar(strInput) Then
                If IsNumeric(strInput) Then
                    strIF = " And 编码 like [1]"
                Else
                    strIF = " And 简码 like [1]"
                End If
            Else
                strIF = " And 名称 like [1]"
            End If
            
            strSQL = "Select ID, 编码, 名称 From 收入项目 Where 末级=1 " & strIF & _
                " And rownum<101 Order by 名称"
        End If
        
        On Error GoTo errH
        vRect = zlControl.GetControlRect(txtInput(Index).hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "项目选择", 1, "", "请选择", False, False, True, vRect.Left, vRect.Top, txtInput(Index).Height, blnCancel, False, True, strMatch & strInput & "%")
        If Not rsTmp Is Nothing Then
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = rsTmp!ID Then
                        txtInput(Index).SetFocus
                        txtInput(Index).SelStart = 0
                        txtInput(Index).SelLength = Len(txtInput(Index).Text)
                        Exit Sub
                    End If
                Next
                If .ListCount < 100 Then
                    If Index = 0 Then
                        .AddItem rsTmp!编码 & "-" & rsTmp!名称 & "(" & rsTmp!规格 & ")"
                    Else
                        .AddItem rsTmp!编码 & "-" & rsTmp!名称
                    End If
                    .ItemData(.NewIndex) = rsTmp!ID
                    .Selected(.NewIndex) = True
                Else
                    MsgBox "出于性能考虑,搜索项目最多只允许添加100项!", vbInformation, gstrSysName
                End If
            End With
        End If
        
        txtInput(Index).SetFocus
        txtInput(Index).SelStart = 0
        txtInput(Index).SelLength = Len(txtInput(Index).Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo费别.hWnd, KeyAscii)
        If lngIdx = -1 And cbo费别.ListCount > 0 Then lngIdx = 0
        cbo费别.ListIndex = lngIdx
    End If
End Sub

Private Sub MakeFilter()
    Dim strSQL As String, Index As Integer, i As Long, strIDs As String
    Dim strSQLtmp As String
    Dim strNotAudit As String '未审核
    Dim strAudit As String '审核或拒绝
    Dim strAuditStatus As String '状态
    
    If gbln退费申请模式 And mblnApply Then
    
        mstrFilter = ""
        If cboApply.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.申请人=[1]"
        If dtpApplyB.Enabled Then mstrFilter = mstrFilter & " And A.申请时间 Between [2] And [3]"

        'cboAudit.Enabled:53990
        If cboAudit.ListIndex > 0 And cboAudit.Enabled Then strAudit = " And  A.审核人=[4] "
        If dtpAuditB.Enabled Then strAudit = strAudit & " And A.审核时间 Between [5] And [6] "

        '54391
        strAuditStatus = ""
        If chkAudit(0).Value = 1 Then '申请
             strNotAudit = " And A.审核人 is Null "
             strAuditStatus = strAuditStatus & "," & "0"
        End If
        If chkAudit(1).Value = 1 Or chkAudit(2).Value = 1 Then '审核或拒绝
            strAudit = strAudit & " And A.审核人 is Not Null"
            If chkAudit(1).Value = 1 Then strAuditStatus = strAuditStatus & "," & "1"
            If chkAudit(2).Value = 1 Then strAuditStatus = strAuditStatus & "," & "2"
        End If
        
        If strAudit <> "" Then
            strAudit = IIf(strNotAudit <> "", " OR (1 = 1 " & strAudit & " )", strAudit)
        End If
        mstrFilter = mstrFilter & strNotAudit & strAudit
        
        '限制审核状态
        If strAuditStatus = "" Then
            strAuditStatus = "0": chkAudit(0).Value = 1
        Else
            strAuditStatus = Mid(strAuditStatus, 2)
        End If
        mstrFilter = mstrFilter & " And Nvl(A.状态,0) in(" & strAuditStatus & ")"
    Else
        mstrFilter = " And a.登记时间 Between [1] And [2] "
        
        '仅显示退费时才涉及后备数据表,划价单没有转到后备数据表
        If chk收费.Value = 1 Then
            mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
        Else
            mblnDateMoved = False
            '必须要这句,否则退出重新进来时,保留了上次的值
        End If
        
        If cbo费别.ListIndex > 0 Then
            strSQL = "Select Distinct NO From 门诊费用记录 Where 记录性质=1 And 记录状态<>0 And 费别=[3]" & mstrFilter
            mstrFilter = mstrFilter & " And a.NO IN(" & strSQL & ")"
            '可能一张单据的多行有多种费别,包含所筛选的费别的NO都应该出来,所以不能用下面这种方式
            'mstrFilter = mstrFilter & " And 费别='" & txt费别.Text & "'"
        End If
        
        strSQL = ""
        If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
            strSQL = " And a.NO Between [4] And [5]"
        ElseIf txtNOBegin.Text <> "" Then
            strSQL = " And a.NO=[4]"
        End If
        If strSQL <> "" Then
            mstrFilter = mstrFilter & _
                " And a.结帐id In (Select Nvl(c.结帐id, b.结帐id)" & vbNewLine & _
                "          From 门诊费用记录 A, 病人预交记录 B, 病人预交记录 C" & vbNewLine & _
                "          Where a.结帐id = b.结帐id And b.结算序号 = c.结算序号(+) And Mod(a.记录性质, 10) = 1" & vbNewLine & _
                strSQL & ")"
        End If
        
        '门诊病人收费单的付款方式此处记录的是费别编码
        If cbo付款方式.ListIndex <> -1 And cbo付款方式.Text <> "" Then
            mstrFilter = mstrFilter & " And a.付款方式=[6]"   '更改问题:33789时,取消的.(问题未登记,与测试人员(王玲说了的))
        End If
        
        If cbo操作员.ListIndex = -1 Then
            mstrFilter = mstrFilter & " And a.操作员姓名||''=[7]"
        ElseIf cbo操作员.ItemData(cbo操作员.ListIndex) > 0 Then
            mstrFilter = mstrFilter & " And a.操作员姓名||''=[7]"
        End If
        
        
        If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
            If Val(Nvl(mrsInfo!ID)) = mlngPrePatient Then
                mstrFilter = mstrFilter & " And a.病人ID=[19]"
            End If
        End If
    
       If txt姓名.Text <> "" Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtPatient.Text, 1))) > 0 Then
                mstrFilter = mstrFilter & " And Upper(a.姓名) Like [8]"
            Else
                mstrFilter = mstrFilter & " And a.姓名 Like [8]"
            End If
        End If
            
        If txt标识号.Text <> "" Then mstrFilter = mstrFilter & " And a.标识号=[9]"
        
        strSQL = ""
        If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
            '无需根据票据号判断,直接根据单据的登记时间判断
            strSQLtmp = IIf(txtFactEnd.Text = "", " =[10] ", " Between [10] And [11] ")
            strSQL = "Select A.NO" & _
            " From 票据打印内容 A,票据使用明细 B" & _
            " Where A.数据性质=1 And A.ID=B.打印ID And B.票种=1 And B.性质=1" & _
            " And B.号码 " & strSQLtmp
        End If
        If strSQL <> "" Then
            mstrFilter = mstrFilter & _
                " And a.结帐id In (Select Nvl(c.结帐id, b.结帐id)" & vbNewLine & _
                "          From 门诊费用记录 A, 病人预交记录 B, 病人预交记录 C" & vbNewLine & _
                "          Where a.结帐id = b.结帐id And b.结算序号 = c.结算序号(+) And Mod(a.记录性质, 10) = 1" & vbNewLine & _
                "                And a.NO IN(" & strSQL & ")" & vbNewLine & _
                            ")"
        End If
        
        '药店固定为所有科室
        If Not glngSys Like "8??" Then
            If cbo科室.ListIndex <> 0 Then
                mstrFilter = mstrFilter & " And a.开单部门ID+0=[12]"
            End If
            If txt开单人.Text <> "" Then
                mstrFilter = mstrFilter & " And a.开单人=[17]"
            End If
        End If
        
        If InStr(1, mstrPrivs, "明细项目过滤") > 0 Then
            For Index = 0 To 0      '将来可能会加收入项目条件
                strIDs = ""
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    If ListFeeItem(Index).Selected(i) Then
                        strIDs = strIDs & "," & ListFeeItem(Index).ItemData(i)
                    End If
                Next
                If strIDs <> "" Then
                    strIDs = Mid(strIDs, 2)
                    If Index = 0 Then
                        mstrFeeItems = strIDs
                        mstrFilter = mstrFilter & " And Instr(','||[18]||',',','||a.收费细目ID||',')>0"
                    'Else
                        'mstrIncomeItems = strIDs
                        'mstrFilter = mstrFilter & " And Instr(','||[10]||',',','||a.收入项目ID||',')>0"
                    End If
                End If
            Next
        End If
    End If
End Sub
 

Private Sub txt标识号_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Beep: Exit Sub
        End If
    End If
End Sub

 
Private Sub txt开单人_GotFocus()
    zlControl.TxtSelAll txt开单人
End Sub

Private Sub txt开单人_Validate(Cancel As Boolean)
    Dim strDoctor As String
    strDoctor = UCase(Trim(txt开单人.Text))
    If strDoctor <> "" Then
        If zlCommFun.IsNumOrChar(strDoctor) Then
            strDoctor = GetDoctorName(strDoctor)
        End If
    End If
    txt开单人.Text = strDoctor
End Sub

Private Function GetDoctorName(ByVal strCode As String) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIF As String, lngDept As Long, blnCancel As Boolean, vRect As RECT
    
    If zlCommFun.IsCharAlpha(strCode) Then
        strIF = " And 简码 Like [1]"
        strCode = strCode & "%"
    Else
        strIF = " And (简码 = [1] Or 编号 = [1])"
    End If
    If cbo科室.ListIndex > 0 Then
        strIF = strIF & " And B.部门ID = [2]"
        lngDept = cbo科室.ItemData(cbo科室.ListIndex)
    End If
    
    strSQL = "Select Distinct A.Id,A.姓名" & vbNewLine & _
            "From 人员表 A, 部门人员 B, 人员性质说明 C, 部门性质说明 D" & vbNewLine & _
            "Where A.ID = B.人员id And A.ID = C.人员id And B.部门id = D.部门id And C.人员性质 In ('医生', '护士') And D.服务对象 In (" & gint病人来源 & ", 3) And" & vbNewLine & _
            "      D.工作性质 In ('临床','手术')" & strIF

    vRect = zlControl.GetControlRect(txt开单人.hWnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "选择医生", 1, "", "请选择医生", False, False, True, vRect.Left, vRect.Top, txt开单人.Height, blnCancel, False, True, strCode, lngDept)
    If Not rsTmp Is Nothing Then
        GetDoctorName = rsTmp!姓名
    End If
End Function

 

 

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function
'获取默认IDKind索引
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
End Function
                
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    If mlngPrePatient Then txtPatient.PasswordChar = ""
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
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
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If Trim(txtPatient.Text) <> "" Then Call FindPati(objCard, False, Trim(txtPatient.Text))
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    '问题:60010
    If txtPatient.Locked Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    If objCard.名称 Like "身份证*" And objCard.系统 Then
        txtPatient.Text = objPatiInfor.身份证号
    Else
        txtPatient.Text = objPatiInfor.卡号
    End If
    If Trim(txtPatient.Text) <> "" Then Call FindPati(objCard, False, Trim(txtPatient.Text))
End Sub
Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '返回：查找成功,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    
    strSQL = ""
    mlngPrePatient = 0
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then   '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If lng病人ID <= 0 Then lng病人ID = 0
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And B.住院号=[2]" & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!姓名 Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtPatient.Text = ""
                        Set mrsInfo = Nothing: Exit Function
                    Else
                       strSQL = strSQL & " And  B.姓名 Like [3]"
                       
                    End If
                Else
                    strSQL = strSQL & " And B.病人ID=[2]"
                    strInput = "-" & Val(mrsInfo!病人ID)
                End If
            Case "医保号"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:26982
                    strSQL = strSQL & " And B.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.医保号=[1]" & str非在院
                End If
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                ' strSQL = strSQL & " And B.身份证号=[1] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.门诊号=[1]" & str非在院
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.住院号=[1]" & str非在院
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    If lng病人ID = 0 Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID <= 0 Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select distinct  B.病人id As ID, Decode(sign(nvl(ylkxx.病人id,0)),0,'','√') as 三方账户, B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSQL = strSQL & vbNewLine & "      A.名称 险类名称"
    strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A,医疗卡类别 YLK,病人医疗卡信息 YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) and b.病人id=ylkxx.病人id(+) and ylkxx.状态(+)=0 and  ylkxx.卡类别id=ylk.id(+)  and ylk.是否自制(+)=0 And B.停用时间 Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
    
    If mrsInfo Is Nothing Then GoTo ClearPati:
    If mrsInfo.State <> 1 Then GoTo ClearPati:
    If mrsInfo.RecordCount = 0 Then GoTo ClearPati:
    If Val(Nvl(mrsInfo!ID)) = 0 Then GoTo ClearPati:
    
    txtPatient.Text = Nvl(mrsInfo!姓名)
    Me.txtPatient.Tag = Nvl(mrsInfo!ID)
    mlngPrePatient = Val(Nvl(mrsInfo!ID))
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    GetPatient = True
    Exit Function
ClearPati:
    txtPatient.Text = ""
    txtPatient.PasswordChar = ""
    Set mrsInfo = Nothing
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub


Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then IDKind.SetAutoReadCard True
End Sub


Private Sub txtPatient_LostFocus()
    IDKind.SetAutoReadCard False
End Sub

 

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If IsCardType(IDKind, "姓名") Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii = 13 Then
            KeyAscii = 0
            If Val(txtPatient.Tag) <> 0 Then    '存在
                 zlCommFun.PressKey vbKeyTab: Exit Sub
            End If
        End If
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-09-03 09:32:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not GetPatient(objCard, strInput, blnCard) Then Exit Sub
End Sub

