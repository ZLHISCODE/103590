VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmPACSFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤条件"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPACSFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox PicButton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5460
      TabIndex        =   42
      Top             =   4500
      Width           =   5460
      Begin VB.ComboBox cboSchemaName 
         Height          =   330
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdDelSchema 
         Caption         =   "删除方案(&D)"
         Height          =   375
         Left            =   2520
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveSchema 
         Caption         =   "保存方案(&S)"
         Height          =   375
         Left            =   3840
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkmore 
         Caption         =   "更多(&M)"
         Height          =   375
         Left            =   1200
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefault 
         Caption         =   "缺省(&F)"
         Height          =   375
         Left            =   105
         TabIndex        =   28
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Left            =   4200
         TabIndex        =   39
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   375
         Left            =   3120
         TabIndex        =   37
         ToolTipText     =   "F2"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label labSchema 
         Caption         =   "查询方案："
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picmore 
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
      Height          =   2745
      Left            =   -120
      ScaleHeight     =   2745
      ScaleWidth      =   5745
      TabIndex        =   46
      Top             =   4920
      Visible         =   0   'False
      Width           =   5745
      Begin VB.TextBox txt报告内容 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   31
         Top             =   405
         Width           =   3690
      End
      Begin VB.TextBox Txt影像诊断 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   30
         Top             =   30
         Width           =   3690
      End
      Begin VB.Frame Frame1 
         Caption         =   "PACS报告查询"
         Height          =   1455
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   5250
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   3690
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   34
            Top             =   600
            Width           =   3690
         End
         Begin VB.TextBox txtPacsRpt 
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   35
            Top             =   960
            Width           =   3690
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "检查所见"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   390
            Width           =   975
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "或 诊断意见"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lblPacsRpt 
            Caption         =   "或 建议"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   1110
            Width           =   1095
         End
      End
      Begin VB.TextBox txt随访 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1680
         TabIndex        =   32
         Top             =   795
         Width           =   3690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报告内容"
         Height          =   210
         Left            =   240
         TabIndex        =   55
         Top             =   465
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "疾病诊断"
         Height          =   210
         Left            =   240
         TabIndex        =   52
         Top             =   90
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "随访"
         Height          =   210
         Left            =   240
         TabIndex        =   51
         Top             =   855
         Width           =   420
      End
   End
   Begin VB.Frame Frabase 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   -120
      TabIndex        =   36
      Top             =   0
      Width           =   5685
      Begin VB.OptionButton optFindType 
         Caption         =   "安排"
         Height          =   300
         Index           =   4
         Left            =   1500
         TabIndex        =   65
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboPartGroup 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   1980
      End
      Begin VB.ComboBox cboAgeType 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":000C
         Left            =   4545
         List            =   "frmPACSFilter.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1920
         Width           =   825
      End
      Begin VB.TextBox txtEndAge 
         Height          =   330
         Left            =   3855
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txtBeginAge 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2400
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox cboSex 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0038
         Left            =   800
         List            =   "frmPACSFilter.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   1065
      End
      Begin VB.ComboBox cboAgeWhere 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":003C
         Left            =   3090
         List            =   "frmPACSFilter.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1920
         Width           =   705
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "半年"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "三月"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4425
         TabIndex        =   12
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "二月"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3945
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "一月"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3450
         TabIndex        =   10
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "二周"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2970
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         Caption         =   "一周"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2475
         TabIndex        =   8
         Top             =   1320
         Width           =   510
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "两天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   500
      End
      Begin VB.CommandButton cmdDayCfg 
         BackColor       =   &H00C0FFC0&
         Caption         =   "今天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cboYinYangXing 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":006E
         Left            =   3885
         List            =   "frmPACSFilter.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3975
         Width           =   1530
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "采图"
         Height          =   300
         Index           =   3
         Left            =   3000
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.ComboBox cboDiagDOC 
         Height          =   330
         Left            =   3885
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3615
         Width           =   1530
      End
      Begin VB.ComboBox cbo检查技师 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3135
         Width           =   1530
      End
      Begin VB.ComboBox cboAuditing 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3975
         Width           =   1530
      End
      Begin VB.ComboBox cboModality 
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0095
         Left            =   1275
         List            =   "frmPACSFilter.frx":0097
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2400
         Width           =   1980
      End
      Begin VB.ComboBox cbo质量 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "frmPACSFilter.frx":0099
         Left            =   3885
         List            =   "frmPACSFilter.frx":00A6
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3135
         Width           =   1530
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "报到"
         Height          =   300
         Index           =   2
         Left            =   2230
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.ComboBox cboPart 
         Height          =   330
         Left            =   3360
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2760
         Width           =   2055
      End
      Begin VB.OptionButton optFindType 
         Caption         =   "申请"
         Height          =   300
         Index           =   1
         Left            =   720
         TabIndex        =   0
         Top             =   240
         Width           =   840
      End
      Begin VB.ComboBox cboDept 
         Height          =   330
         Left            =   1275
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   3600
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3240
         TabIndex        =   4
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   104726531
         CurrentDate     =   38082.9993055556
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   104726531
         CurrentDate     =   38082
      End
      Begin MSComctlLib.Slider sldDays 
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   529
         _Version        =   393216
         Min             =   1
         Max             =   180
         SelStart        =   1
         TickFrequency   =   7
         Value           =   1
      End
      Begin VB.Line Line1 
         X1              =   1680
         X2              =   5400
         Y1              =   1755
         Y2              =   1755
      End
      Begin VB.Label Label8 
         Caption         =   "近"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1650
         TabIndex        =   64
         Top             =   1395
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "按(                             )时间检索"
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   270
         Width           =   5055
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   210
         Left            =   315
         TabIndex        =   58
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   210
         Left            =   1920
         TabIndex        =   57
         Top             =   1980
         Width           =   420
      End
      Begin VB.Label labYinYangXing 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "阴 阳 性"
         Height          =   210
         Left            =   2955
         TabIndex        =   56
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label labDiagDOC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断医生"
         Height          =   210
         Left            =   2955
         TabIndex        =   54
         Top             =   3675
         Width           =   840
      End
      Begin VB.Label lab检查技师 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查技师"
         Height          =   210
         Left            =   315
         TabIndex        =   53
         Top             =   3195
         Width           =   840
      End
      Begin VB.Label labAuditing 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核医生"
         Height          =   210
         Left            =   315
         TabIndex        =   45
         Top             =   4035
         Width           =   840
      End
      Begin VB.Label labModality 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "类    别"
         Height          =   210
         Left            =   315
         TabIndex        =   44
         Top             =   2460
         Width           =   840
      End
      Begin VB.Label lab质量 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "影像质量"
         Height          =   210
         Left            =   2955
         TabIndex        =   43
         Top             =   3195
         Width           =   840
      End
      Begin VB.Label labPartGroup 
         AutoSize        =   -1  'True
         Caption         =   "检查部位"
         Height          =   210
         Left            =   315
         TabIndex        =   41
         Top             =   2820
         Width           =   840
      End
      Begin VB.Label labDept 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人科室"
         Height          =   210
         Left            =   315
         TabIndex        =   40
         Top             =   3660
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
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
         Left            =   2760
         TabIndex        =   38
         Top             =   630
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmPACSFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngModul As Long      '模块号调用
Public mBeforeDays As Integer '查询天数
Public mDept As Long
Public mblnOK As Boolean '确定退出
Private mrsStudyPart As ADODB.Recordset
Private mrsPartGroup As ADODB.Recordset


Private Sub LoadSex()
'********************************************
'
'读取病人性别
'
'********************************************
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select 名称 From 性别"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取病人性别")
    CboSex.Clear
    CboSex.AddItem "全部"
        
    With Me.CboSex
        Do While Not rsTmp.EOF
            .AddItem zlCommFun.SpellCode(Nvl(rsTmp("名称"))) & "-" & Nvl(rsTmp("名称"))
            rsTmp.MoveNext
        Loop
    End With
    
    CboSex.ListIndex = 0
 
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cboAgeWhere_Click()
    txtBeginAge.Enabled = IIf(cboAgeWhere.Text = "~ ", True, False)
    If txtBeginAge.Enabled Then
        txtBeginAge.BackColor = &HFFFFFF
    Else
        txtBeginAge.Text = ""
        txtBeginAge.BackColor = &HE0E0E0
    End If
End Sub

Private Sub cboModality_Click()
On Error GoTo ErrHandle
    Call FilterGroupPart(cboModality.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cboPartGroup_Click()
On Error GoTo ErrHandle
    Call FilterStudyPart(cboPartGroup.Text)
Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub chkmore_Click()
    If chkmore.value = 1 Then
        Me.Height = Picmore.Top + Picmore.Height + PicButton.Height + 400
        Picmore.Visible = True
    Else
        Me.Height = Frabase.Top + Frabase.Height + PicButton.Height + 400
        Picmore.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Me.Hide
End Sub

Private Sub cmdDayCfg_Click(Index As Integer)
    dtpEnd.value = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59") 'zlDatabase.Currentdate
    
    Select Case Index
        Case 0
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd"))
            sldDays.value = 1
        Case 1
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 1
            sldDays.value = 2
        Case 2
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 6
            sldDays.value = 7
        Case 3
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 13
            sldDays.value = 14
        Case 4
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 29
            sldDays.value = 30
        Case 5
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 59
            sldDays.value = 60
        Case 6
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 89
            sldDays.value = 90
        Case 7
            dtpBegin.value = CDate(Format(dtpEnd.value, "yyyy-mm-dd")) - 179
            sldDays.value = 180
    End Select
End Sub

Private Sub cmdDefault_Click()
    Call Form_Load
    Call Form_Activate
End Sub



Private Sub cmdOK_Click()

    If dtpEnd.value < dtpBegin.value Then
        MsgBoxD Me, "开始时间不能大于截止时间，请检查！", vbInformation, gstrSysName
        dtpEnd.SetFocus
        Exit Sub
    End If
    
    mblnOK = True
    Me.Hide
End Sub

Private Sub GetQuerySchemaCfg(ByRef strSchemaFormat As String, ByRef strSchemaField As String)
    Dim strSchema As String
    Dim intBeginAge As Integer
    Dim intEndAge As Integer
    
    If Not dtpBegin Is Nothing Then
        If optFindType(1).value Then strSchema = strSchema & GetFieldFormatStr("发送时间", "病人医嘱发送", ">=", dtpBegin.value, "And", "")
        If optFindType(2).value Then strSchema = strSchema & GetFieldFormatStr("首次时间", "病人医嘱发送", ">=", dtpBegin.value, "And", "")
        If optFindType(3).value Then strSchema = strSchema & GetFieldFormatStr("接收日期", "影像检查记录", ">=", dtpBegin.value, "And", "")
    End If
        
    If Not dtpEnd Is Nothing Then
        If optFindType(1).value Then strSchema = strSchema & GetFieldFormatStr("发送时间", "病人医嘱发送", "<=", dtpEnd.value, "And", "")
        If optFindType(2).value Then strSchema = strSchema & GetFieldFormatStr("首次时间", "病人医嘱发送", "<=", dtpEnd.value, "And", "")
        If optFindType(3).value Then strSchema = strSchema & GetFieldFormatStr("接收日期", "影像检查记录", "<=", dtpEnd.value, "And", "")
    End If
    
    If CboSex.ListIndex <> 0 Then strSchema = strSchema & GetFieldFormatStr("性别", "病人信息", "<=", NeedName(CboSex.Text), "And", "")
    
    
    Select Case NeedName(cboAgeType.Text)
        Case "岁"
            intBeginAge = Val(txtBeginAge.Text) * 365
            intEndAge = Val(txtEndAge.Text) * 365
        Case "月"
            intBeginAge = Val(txtBeginAge.Text) * 30
            intEndAge = Val(txtEndAge.Text) * 30
        Case "周"
            intBeginAge = Val(txtBeginAge.Text) * 7
            intEndAge = Val(txtEndAge.Text) * 7
        Case "天"
            intBeginAge = Val(txtBeginAge.Text) * 1
            intEndAge = Val(txtEndAge.Text) * 1
    End Select
        
    If Trim(txtBeginAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", ">=", CStr(intBeginAge), "And", "")
        End If
    End If
    
    If Trim(txtEndAge.Text) <> "" Then
        If Trim(cboAgeWhere.Text) = "~" Then
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", "<=", CStr(intEndAge), "And", "")
        Else
            strSchema = strSchema & GetFieldFormatStr("年龄", "病人信息", Trim(cboAgeWhere.Text), CStr(intEndAge), "And", "")
        End If
    End If
    
    If cboDept.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("病人科室ID+0", "病人医嘱记录", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
    
    If cboPart.ListIndex <> 0 Then
        strSchema = strSchema & GetFieldFormatStr("病人科室ID+0", "病人医嘱记录", "=", cboDept.ItemData(cboDept.ListIndex), "And", "")
    End If
        
End Sub

Private Function GetFieldFormatStr(strFieldName As String, strTabName As String, _
    strWhere As String, strData As String, strLink As String, strQueryType As String, Optional strBracket As String) As String
    Dim strResult As String
            
    On Error GoTo ErrHandle
        strResult = "<" & strFieldName & ">#B=" & strBracket & "#F=" & strTabName & "#W=" & strWhere & "#D=" & strData & "#L=" & strLink & "#T=" & strQueryType & "</" & strFieldName & ">"
        
        GetFieldFormatStr = strResult & vbNewLine
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub GetQueryTimeType()
    'If optFindType(1).value Then GetQueryTimeType = "发送时间"
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
    'If KeyAscii = vbKeyF2 Then cmdOK_Click
End Sub
Private Sub Form_Activate()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Txt影像诊断.Text = ""
    txt报告内容.Text = ""
    'cboYinYangXing.ListIndex = 0
    'cbo质量.ListIndex = 0
    'cboSex.ListIndex = 0
    'cboAgeWhere.ListIndex = 4
    'cboAgeType.ListIndex = 0
    txt随访.Text = ""
    
    
    '判断和提取PACS报告内容的三个标题
    If pReport_CheckViewName = "" Or pReport_ResultName = "" Or pReport_AdviceName = "" Then
        
        pReport_CheckViewName = "检查所见"
        pReport_ResultName = "诊断意见"
        pReport_AdviceName = "建议"
        
        strSQL = "select ID ,科室ID,参数名,参数值 from 影像流程参数 where 科室ID = [1] " _
            & " and (参数名 = '检查所见名称' or 参数名 = '诊断意见名称' or 参数名 = '建议名称') "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mDept)
        
        If rsTemp.EOF = False Then
            Select Case rsTemp!参数名
                Case "检查所见名称"
                    pReport_CheckViewName = Nvl(rsTemp!参数值, "检查所见")
                Case "诊断意见名称"
                    pReport_ResultName = Nvl(rsTemp!参数值, "诊断意见")
                Case "建议名称"
                    pReport_AdviceName = Nvl(rsTemp!参数值, "建议")
            End Select
        End If
    End If
    txtPacsRpt(0).Text = ""
    txtPacsRpt(1).Text = ""
    txtPacsRpt(2).Text = ""
    lblPacsRpt(0).Caption = pReport_CheckViewName
    lblPacsRpt(1).Caption = "或 " & pReport_ResultName
    lblPacsRpt(2).Caption = "或 " & pReport_AdviceName
        
    dtpBegin.SetFocus
End Sub
Private Sub LoadDept()
'功能：根据病人来源读取病人科室
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    strSQL = "Select Distinct A.ID,A.编码,A.名称,B.服务对象" & _
        " From 部门表 A,部门性质说明 B" & _
        " Where A.ID=B.部门ID And B.工作性质 IN('临床','手术')" & _
        " And (A.撤档时间 is NULL Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cboDept.Clear
    cboDept.AddItem "所有科室"
    cboDept.ListIndex = 0
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        rsTmp.MoveNext
    Next

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub InitModality()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select 简码,名称 from 诊疗检查类型"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "诊疗检查类型")
    
    cboModality.Clear
    cboModality.AddItem "所有类别"
    cboModality.ListIndex = 0
    
    Do Until rsTemp.EOF
        cboModality.AddItem rsTemp!简码 & "-" & rsTemp!名称
        rsTemp.MoveNext
    Loop
    
End Sub

Private Function InitPart() As Boolean
    Dim strSQL As String
    Dim strGroup As String
    Dim strStudyPart As String
    Dim strRead As String
    Dim strGroupPY  As String
    Dim strFilter As String
    
    strFilter = ""
    If mlngModul = G_LNG_PATHOLSYS_NUM Then
        strFilter = " where 类型='病理' or upper(类型)='DG' or upper(类型)='BL'"
    End If
    
    '读取部位分组
    strSQL = "select Distinct 分组 || '(' || 类型 || ')' as 分组, 类型 from 诊疗检查部位 " & strFilter & " order by 类型"
    Set mrsPartGroup = zlDatabase.OpenSQLRecord(strSQL, "提取部位")
    
    cboPartGroup.Clear
    cboPartGroup.AddItem "所有分组"
    cboPartGroup.ListIndex = 0
    
    While Not mrsPartGroup.EOF
        strGroup = Nvl(mrsPartGroup("分组"))
        
         '截取分组名称 并转换为拼音简码
        strGroupPY = zlCommFun.SpellCode(Mid(strGroup, InStr(strGroup, "-") + 1, InStrRev(strGroup, "(") - InStr(strGroup, "-") - 1))
        
        strGroup = IIf(InStr(strGroup, "-") = 0, "-" & strGroup, strGroup)
        
        strGroup = strGroupPY & Mid(strGroup, InStr(strGroup, "-"), Len(strGroup))
        
        cboPartGroup.AddItem strGroup
        mrsPartGroup.MoveNext
    Wend
    

    
    '读取检查部位
    strSQL = "Select Distinct 名称, 分组 || '(' || 类型 || ')'  as 分组, 编码 From 诊疗检查部位 " & strFilter & " order by 编码"
    Set mrsStudyPart = zlDatabase.OpenSQLRecord(strSQL, "提取部位")
    
    cboPart.Clear
    cboPart.AddItem "所有部位"
    cboPart.ListIndex = 0
    
    strRead = ""
    
    With Me.cboPart
        Do While Not mrsStudyPart.EOF
            strStudyPart = zlCommFun.SpellCode(Nvl(mrsStudyPart("名称"))) & "-" & Nvl(mrsStudyPart("名称"))
            
            If InStr(strRead, strStudyPart & ";") <= 0 Then
                .AddItem strStudyPart
                
                strRead = strRead & strStudyPart & ";"
            End If
            
            mrsStudyPart.MoveNext
        Loop
    End With
    

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub FilterStudyPart(ByVal strGroupName As String)
    Dim strStudyPart As String
    Dim strRead As String
    Dim strSQL As String
    Dim strGroup As String
    Dim strType As String
    Dim rsGroup As ADODB.Recordset
    
     If mrsStudyPart Is Nothing Then Exit Sub
    
    If strGroupName <> "所有分组" Then
        '因加载时将分组前缀的编号 替换成了拼音简码，现需要
        strGroup = Mid(strGroupName, InStr(strGroupName, "-") + 1, InStrRev(strGroupName, "(") - InStr(strGroupName, "-") - 1)
        strType = Mid(strGroupName, InStr(strGroupName, "(") + 1, InStrRev(strGroupName, ")") - InStr(strGroupName, "(") - 1)
    
        strSQL = "Select Distinct  分组 From 诊疗检查部位 where 类型 = '" & strType & "' and 分组 like '%" & strGroup & "%'"
        Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, "得到分组信息")

        strGroup = IIf(rsGroup.RecordCount < 1, "", Nvl(rsGroup!分组) & "(" & strType & ")")
    End If
    
    mrsStudyPart.Filter = IIf(strGroupName = "所有分组", "", "分组='" & strGroup & "'")
    
    cboPart.Clear
    cboPart.AddItem "所有部位"
    cboPart.ListIndex = 0
    
    With Me.cboPart
        Do While Not mrsStudyPart.EOF
            strStudyPart = zlCommFun.SpellCode(Nvl(mrsStudyPart("名称"))) & "-" & Nvl(mrsStudyPart("名称"))
            
            If InStr(strRead, strStudyPart & ";") <= 0 Then
                .AddItem strStudyPart
                
                strRead = strRead & strStudyPart & ";"
            End If
            
            mrsStudyPart.MoveNext
        Loop
    End With

End Sub

Private Sub FilterGroupPart(ByVal strTypeName As String)
'过滤部位分组
    Dim strGroupPart As String
    Dim strRead As String
    Dim strType As String
    Dim strGroupPY As String
    
    If mrsPartGroup Is Nothing Then Exit Sub
    
    strType = Mid(strTypeName, InStr(strTypeName, "-") + 1, Len(strTypeName))

    mrsPartGroup.Filter = IIf(strTypeName = "所有类别", "", "类型='" & strType & "'")
    
    cboPartGroup.Clear
    cboPartGroup.AddItem "所有分组"
    cboPartGroup.ListIndex = 0
    
    With Me.cboPartGroup
        Do While Not mrsPartGroup.EOF
            strGroupPart = Nvl(mrsPartGroup("分组"))
            
            '截取分组名称 并转换为拼音简码
            strGroupPY = zlCommFun.SpellCode(Mid(strGroupPart, InStr(strGroupPart, "-") + 1, InStrRev(strGroupPart, "(") - InStr(strGroupPart, "-") - 1))
            
            strGroupPart = IIf(InStr(strGroupPart, "-") = 0, "-" & strGroupPart, strGroupPart)
            
            strGroupPart = strGroupPY & Mid(strGroupPart, InStr(strGroupPart, "-"), Len(strGroupPart))
        
            .AddItem strGroupPart
            
            mrsPartGroup.MoveNext
        Loop
    End With
    
    
End Sub




Private Sub InitDoc()
Dim rsTmp As ADODB.Recordset
    cboDiagDOC.Clear: cboAuditing.Clear: cbo检查技师.Clear
    cboDiagDOC.AddItem "所有医生": cboAuditing.AddItem "所有医生": cbo检查技师.AddItem "所有医生"
    cboDiagDOC.ListIndex = 0: cboAuditing.ListIndex = 0: cbo检查技师.ListIndex = 0
    On Error GoTo errH
    gstrSQL = "select distinct A.简码,A.姓名 from 人员表 A,部门人员 B where B.部门ID=[1] AND A.ID=B.人员ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "提取本科室医生", mDept)
    If rsTmp Is Nothing Then Exit Sub
    Do While Not rsTmp.EOF
        cboDiagDOC.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cboAuditing.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        cbo检查技师.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optAdviceTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub optCheckTime_Click()
    Me.dtpBegin.SetFocus
End Sub

Private Sub Form_Load()
Dim curDate As Date
Dim int时间类型 As Integer

    curDate = zlDatabase.Currentdate
    dtpEnd.value = Format(curDate, "yyyy-MM-dd 23:59")
    dtpEnd.Tag = dtpEnd.value
    dtpBegin.value = Format(dtpEnd.value - mBeforeDays, "yyyy-MM-dd 00:00")
    
    int时间类型 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", 1))
    If int时间类型 = 1 Then
        optFindType(1).value = True
    ElseIf int时间类型 = 2 Then
        optFindType(2).value = True
    ElseIf int时间类型 = 3 Then
        optFindType(3).value = True
    Else
        optFindType(4).value = True
    End If
    
    '调整位置时，增加Me.Visible条件，是让窗口只在第一次加载时，才进行调整
    '病理
    If mlngModul = G_LNG_PATHOLSYS_NUM And Not Me.Visible Then
        optFindType(1).Caption = "开嘱"
        optFindType(3).Caption = "病理申请"
        
        
        labModality.Visible = False
        cboModality.Visible = False
        
        lab检查技师.Visible = False
        cbo检查技师.Visible = False
        
        lab质量.Visible = False
        cbo质量.Visible = False
        
        labPartGroup.Top = labPartGroup.Top - cboModality.Height
        cboPartGroup.Top = cboPartGroup.Top - cboModality.Height
        cboPart.Top = cboPart.Top - cboModality.Height
        
        labDept.Top = labDept.Top - cbo检查技师.Height - cboModality.Height
        cboDept.Top = cboDept.Top - cbo检查技师.Height - cboModality.Height
        
        labDiagDOC.Top = labDiagDOC.Top - cbo检查技师.Height - cboModality.Height
        cboDiagDOC.Top = cboDiagDOC.Top - cbo检查技师.Height - cboModality.Height
        
        labAuditing.Top = labAuditing.Top - cbo检查技师.Height - cboModality.Height
        cboAuditing.Top = cboAuditing.Top - cbo检查技师.Height - cboModality.Height
        
        labYinYangXing.Top = labYinYangXing.Top - cbo检查技师.Height - cboModality.Height
        cboYinYangXing.Top = cboYinYangXing.Top - cbo检查技师.Height - cboModality.Height
        
        Frabase.Height = Frabase.Height - cbo检查技师.Height - cboModality.Height
        Picmore.Top = Picmore.Top - cbo检查技师.Height - cboModality.Height
        
        Me.Height = Me.Height - cbo检查技师.Height - cboModality.Height
        
    '采集
    ElseIf mlngModul = G_LNG_VIDEOSTATION_MODULE And Not Me.Visible Then
        labModality.Visible = False
        cboModality.Visible = False
        
        labPartGroup.Top = labPartGroup.Top - cboModality.Height
        cboPartGroup.Top = cboPartGroup.Top - cboModality.Height
        cboPart.Top = cboPart.Top - cboModality.Height
        
        labDept.Top = labDept.Top - cboModality.Height
        cboDept.Top = cboDept.Top - cboModality.Height
        
        labDiagDOC.Top = labDiagDOC.Top - cboModality.Height
        cboDiagDOC.Top = cboDiagDOC.Top - cboModality.Height
        
        labAuditing.Top = labAuditing.Top - cboModality.Height
        cboAuditing.Top = cboAuditing.Top - cboModality.Height
        
        labYinYangXing.Top = labYinYangXing.Top - cboModality.Height
        cboYinYangXing.Top = cboYinYangXing.Top - cboModality.Height
        
        Frabase.Height = Frabase.Height - cboModality.Height
        Picmore.Top = Picmore.Top - cboModality.Height
        
        lab检查技师.Top = lab检查技师.Top - cboModality.Height
        cbo检查技师.Top = cbo检查技师.Top - cboModality.Height
        
        lab质量.Top = lab质量.Top - cboModality.Height
        cbo质量.Top = cbo质量.Top - cboModality.Height
        
        Me.Height = Me.Height - cboModality.Height
    '医技
    Else
        '......
    End If
    
    '只在窗体第一次加载时执行，避免点击缺省功能按钮后，在重复进行加载
    If Not Me.Visible Then
        LoadDept
        LoadSex
        InitModality
        InitPart
        InitDoc
    End If
    
    
    txtBeginAge.Text = ""
    txtEndAge.Text = ""
    
    cboModality.Text = "所有类别"
    cboPartGroup.Text = "所有分组"
    cboPart.Text = "所有部位"
    cboDept.Text = "所有科室"
    cboDiagDOC.Text = "所有医生"
    cboAuditing.Text = "所有医生"
    cbo检查技师.Text = "所有医生"
    
    cboYinYangXing.ListIndex = 0
    cbo质量.ListIndex = 0
    CboSex.ListIndex = 0
    cboAgeWhere.ListIndex = 0
    cboAgeType.ListIndex = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Dim int时间类型 As Integer
    '保存过滤时间类型
    If optFindType(1).value = True Then
        int时间类型 = 1
    ElseIf optFindType(2).value = True Then
        int时间类型 = 2
    ElseIf optFindType(3).value = True Then
        int时间类型 = 3
    Else
        int时间类型 = 4
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "过滤时间类型", int时间类型
End Sub



Private Sub sldDays_Scroll()
    '设置查询时间范围
    dtpBegin.value = Format(CDate(dtpEnd.value - (sldDays.value - 1)), "yyyy-mm-dd 00:00")
End Sub
