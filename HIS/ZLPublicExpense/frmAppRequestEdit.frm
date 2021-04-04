VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmAppRequestEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "预约登记"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   Icon            =   "frmAppRequestEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   5250
      TabIndex        =   17
      Top             =   4530
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   6450
      TabIndex        =   18
      Top             =   4530
      Width           =   1100
   End
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Height          =   3195
      Left            =   30
      ScaleHeight     =   3135
      ScaleWidth      =   7830
      TabIndex        =   27
      Top             =   1200
      Width           =   7890
      Begin VB.CommandButton cmdAll 
         Height          =   315
         Left            =   3810
         Picture         =   "frmAppRequestEdit.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "所有号源"
         Top             =   60
         Width           =   315
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   1050
         TabIndex        =   15
         Top             =   2280
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93323267
         CurrentDate     =   42398
      End
      Begin VB.TextBox txtStyle 
         Height          =   330
         Index           =   5
         Left            =   4995
         TabIndex        =   16
         Top             =   2280
         Width           =   585
      End
      Begin VB.TextBox txt登记时间 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4995
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2700
         Width           =   2700
      End
      Begin VB.TextBox txt登记人 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1815
      End
      Begin VB.ComboBox cboNote 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   14
         Top             =   1860
         Width           =   6675
      End
      Begin VB.Frame Frame2 
         Caption         =   "复诊信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   165
         TabIndex        =   31
         Top             =   555
         Width           =   7560
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   4
            Left            =   3316
            TabIndex        =   36
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(4)"
            BuddyDispid     =   196612
            BuddyIndex      =   4
            OrigLeft        =   3315
            OrigTop         =   675
            OrigRight       =   3570
            OrigBottom      =   1005
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   0
            Left            =   1111
            TabIndex        =   35
            Top             =   690
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(0)"
            BuddyDispid     =   196612
            BuddyIndex      =   0
            OrigLeft        =   1110
            OrigTop         =   690
            OrigRight       =   1365
            OrigBottom      =   1020
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   3
            Left            =   5746
            TabIndex        =   34
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(3)"
            BuddyDispid     =   196612
            BuddyIndex      =   3
            OrigLeft        =   5745
            OrigTop         =   225
            OrigRight       =   6000
            OrigBottom      =   555
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   2
            Left            =   3316
            TabIndex        =   33
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(2)"
            BuddyDispid     =   196612
            BuddyIndex      =   2
            OrigLeft        =   3315
            OrigTop         =   255
            OrigRight       =   3570
            OrigBottom      =   585
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   4
            Left            =   3000
            TabIndex        =   13
            Top             =   690
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   0
            Left            =   795
            TabIndex        =   12
            Top             =   690
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   3
            Left            =   5430
            TabIndex        =   10
            Top             =   285
            Width           =   315
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   2
            Left            =   3000
            TabIndex        =   8
            Top             =   285
            Width           =   315
         End
         Begin MSComCtl2.UpDown udStyle 
            Height          =   330
            Index           =   1
            Left            =   1111
            TabIndex        =   32
            Top             =   285
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   582
            _Version        =   393216
            Value           =   1
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtStyle(1)"
            BuddyDispid     =   196612
            BuddyIndex      =   1
            OrigLeft        =   1111
            OrigTop         =   285
            OrigRight       =   1366
            OrigBottom      =   615
            Max             =   99
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtStyle 
            Height          =   330
            Index           =   1
            Left            =   795
            TabIndex        =   6
            Top             =   285
            Width           =   315
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "按      个疗程(每个疗程      天)复诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   300
            TabIndex        =   11
            Top             =   713
            Width           =   5085
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "按      天后复诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   4950
            TabIndex        =   9
            Top             =   308
            Width           =   2160
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "按      周后复诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   7
            Top             =   308
            Width           =   2175
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "按      个月后复诊"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   285
            TabIndex        =   5
            Top             =   308
            Width           =   2340
         End
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5415
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   60
         Width           =   2310
      End
      Begin VB.ComboBox cboArrangeNo 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   60
         Width           =   3240
      End
      Begin MSComCtl2.UpDown udStyle 
         Height          =   330
         Index           =   5
         Left            =   5580
         TabIndex        =   44
         Top             =   2280
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtStyle(5)"
         BuddyDispid     =   196612
         BuddyIndex      =   5
         OrigLeft        =   3315
         OrigTop         =   675
         OrigRight       =   3570
         OrigBottom      =   1005
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "登记时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   43
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "登记人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   330
         TabIndex        =   41
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "连续提醒         天"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4110
         TabIndex        =   39
         Top             =   2340
         Width           =   1995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "提醒时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   38
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "复诊说明"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4980
         TabIndex        =   30
         Top             =   120
         Width           =   420
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.TextBox txtBirth 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   750
      Width           =   2670
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   24
      Top             =   615
      Width           =   15000
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   2
      Top             =   180
      Width           =   2670
   End
   Begin VB.TextBox txtClinic 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   750
      Width           =   3195
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6345
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   1590
   End
   Begin VB.TextBox txtGender 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   4740
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   930
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   330
      Left            =   660
      TabIndex        =   1
      Top             =   180
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      Appearance      =   2
      IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   11.25
      FontName        =   "宋体"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "出生日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   26
      Top             =   810
      Width           =   840
   End
   Begin VB.Label lblClinic 
      AutoSize        =   -1  'True
      Caption         =   "门诊号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4080
      TabIndex        =   23
      Top             =   810
      Width           =   630
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5895
      TabIndex        =   22
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4290
      TabIndex        =   21
      Top             =   240
      Width           =   420
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   195
      TabIndex        =   20
      Top             =   240
      Width           =   420
   End
End
Attribute VB_Name = "frmAppRequestEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset, mintInsure As Integer
Private mstrYBPati As String, mstrPassWord As String
Private mrsPlan As ADODB.Recordset, mintIDKind As Integer
Private mrsExtra As ADODB.Recordset
Private mlng病人ID As Long

Private Sub cboArrangeNo_Click()
    If mrsPlan Is Nothing Then Exit Sub
    If cboArrangeNo.ItemData(cboArrangeNo.ListIndex) = 0 Then
        mrsPlan.Filter = "排序=" & cboArrangeNo.ListIndex + 1
        If Not mrsPlan.EOF Then
            txtDept.Text = Nvl(mrsPlan!科室)
        End If
    Else
        mrsExtra.Filter = "ID=" & cboArrangeNo.ItemData(cboArrangeNo.ListIndex)
        If Not mrsExtra.EOF Then
            txtDept.Text = Nvl(mrsExtra!科室)
        End If
    End If
End Sub

Private Sub cboArrangeNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call gobjCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboNote_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmdALL_Click()
    Dim strSql As String
    Dim vRect As RECT
    Dim i As Integer
    Dim rsExtra As ADODB.Recordset
    strSql = "Select a.Id, a.号码, a.科室, a.项目, a.医生姓名, a.医生id, a.科室id" & vbNewLine & _
            "From (Select a.Id, a.号码, b.名称 As 科室, c.名称 As 项目, a.医生姓名, a.医生id, a.科室id" & vbNewLine & _
            "       From 临床出诊号源 A, 部门表 B, 收费项目目录 C" & vbNewLine & _
            "       Where a.科室id = b.Id And Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And a.项目id = c.Id And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (b.站点 = '1' Or b.站点 Is Null) And Exists (Select 1 From 临床出诊安排 M,临床出诊表 N Where M.号源ID = A.ID And M.出诊ID = N.ID And N.发布时间 Is Not Null)" & vbNewLine & _
            "       Order By Decode(a.医生id, [1], 1, 0) Desc, a.医生id, a.科室id, a.号码) A"
    vRect = GetControlRect(cboArrangeNo.hWnd)
    Set rsExtra = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "其他号源选择", False, "", "其他号源选择", _
                                                False, False, True, vRect.Left, vRect.Top - 300, 600, False, True, False, UserInfo.ID)
    If rsExtra Is Nothing Then Exit Sub
    Set mrsExtra = rsExtra
    mrsPlan.Filter = "ID=" & Val(Nvl(mrsExtra!ID))
    If Not mrsPlan.EOF Then
        cboArrangeNo.ListIndex = Val(mrsPlan!排序) - 1
        mrsPlan.Filter = ""
        Exit Sub
    End If
    mrsPlan.Filter = ""
    For i = 0 To cboArrangeNo.ListCount - 1
        If cboArrangeNo.ItemData(i) = Val(Nvl(mrsExtra!ID)) Then
            cboArrangeNo.ListIndex = i
            Exit Sub
        End If
    Next i
    cboArrangeNo.AddItem mrsExtra!号码 & "-" & mrsExtra!项目 & "(" & mrsExtra!医生姓名 & ")"
    cboArrangeNo.ItemData(cboArrangeNo.NewIndex) = Val(Nvl(mrsExtra!ID))
    cboArrangeNo.ListIndex = cboArrangeNo.NewIndex
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub ReadBill(frmParent As Object, lng消息ID As Long)
    On Error GoTo errHandle
    Dim strSql As String, rsTemp As ADODB.Recordset
    cmdOK.Visible = False
    cmdCancel.Caption = "退出(&X)"
    picMain.Enabled = False
    txtPatient.Locked = True
'    IDKind.Locked = True
    
    strSql = "Select b.姓名, b.门诊号, b.性别, b.年龄, b.出生日期, c.号码, d.名称 As 项目, e.名称 As 科室, a.通知原因 As 原因, a.开始时间 As 提醒时间, a.终止时间, a.数量, a.复诊方式, a.登记人," & vbNewLine & _
            "       a.登记时间, a.医生姓名" & vbNewLine & _
            "From 病人服务信息记录 A, 病人信息 B, 临床出诊号源 C, 收费项目目录 D, 部门表 E" & vbNewLine & _
            "Where a.Id = [1] And a.病人id = b.病人id And a.号源id = c.Id And c.科室id = e.Id And c.项目id = d.Id"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng消息ID)
    If rsTemp.EOF Then
        MsgBox "不能读取消息,查看失败!", vbInformation, gstrSysName
        Exit Sub
    End If
    txtPatient.Text = Nvl(rsTemp!姓名)
    txtClinic.Text = Nvl(rsTemp!门诊号)
    txtGender.Text = Nvl(rsTemp!性别)
    txtAge.Text = Nvl(rsTemp!年龄)
    txtBirth.Text = Format(Nvl(rsTemp!出生日期), "yyyy-MM-dd hh:mm:ss")
    txtDept.Text = Nvl(rsTemp!科室)
    cboArrangeNo.AddItem rsTemp!号码 & "-" & rsTemp!项目 & "(" & rsTemp!医生姓名 & ")"
    cboArrangeNo.ListIndex = cboArrangeNo.NewIndex
    Select Case Val(Nvl(rsTemp!复诊方式))
    Case 1
        optStyle(0).Value = 1
        txtStyle(0).Text = Nvl(rsTemp!数量)
        txtStyle(4).Text = CInt(DateDiff("d", Format(Nvl(rsTemp!登记时间), "yyyy-MM-dd hh:mm:ss"), Format(Nvl(rsTemp!终止时间), "yyyy-MM-dd hh:mm:ss")) / Val(Nvl(rsTemp!数量, "1")))
    Case 2
        optStyle(1).Value = 1
        txtStyle(1).Text = Nvl(rsTemp!数量)
    Case 3
        optStyle(2).Value = 1
        txtStyle(2).Text = Nvl(rsTemp!数量)
    Case 4
        optStyle(3).Value = 1
        txtStyle(3).Text = Nvl(rsTemp!数量)
    End Select
    cboNote.Text = Nvl(rsTemp!原因)
    dtpDate.Value = Format(Nvl(rsTemp!提醒时间), "yyyy-MM-dd hh:mm:ss")
    txtStyle(5).Text = DateDiff("d", Format(Nvl(rsTemp!提醒时间), "yyyy-MM-dd hh:mm:ss"), Format(Nvl(rsTemp!终止时间), "yyyy-MM-dd hh:mm:ss"))
    txt登记人.Text = Nvl(rsTemp!登记人)
    txt登记时间.Text = Nvl(rsTemp!登记时间)
    
    Me.Show vbModal, frmParent
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    Dim strSql As String, byt复诊方式 As Byte
    Dim i As Integer, lngNum As Long
    Dim blnFind As Boolean
    
    If mrsInfo Is Nothing Then
        MsgBox "不能确定病人,请先选择病人!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    blnFind = False
    For i = 0 To 3
        If optStyle(i).Value = True Then
            byt复诊方式 = i + 1: lngNum = Val(txtStyle(i).Text)
            blnFind = True
        End If
    Next i
    If blnFind = False Then
        MsgBox "请选择一种复诊方式!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cboArrangeNo.ListIndex = -1 Then
        MsgBox "请选择一个号源!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If gobjCommFun.ActualLen(Me.cboNote.Text) > 100 Then
        MsgBox "输入的复诊说明超长!(最多允许输入50个汉字或100个字符)", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Format(dtpDate.Value, "yyyy-mm-dd hh:mm:ss") < Format(gdatRegistTime, "yyyy-mm-dd hh:mm:ss") Then
        MsgBox "预约登记的时间" & Format(dtpDate.Value, "yyyy-mm-dd hh:mm:ss") & "不是出诊表排班模式生效的时间,不能登记!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If cboArrangeNo.ItemData(cboArrangeNo.ListIndex) = 0 Then
        mrsPlan.Filter = "排序=" & cboArrangeNo.ListIndex + 1
        If mrsPlan.EOF Then
            MsgBox "不能确定当前选择的号源,无法进行登记!", vbInformation, gstrSysName
            Exit Sub
        End If
        
        strSql = "zl_病人预约登记_Insert("
        strSql = strSql & mrsInfo!病人ID & ","
        strSql = strSql & mrsPlan!ID & ","
        strSql = strSql & byt复诊方式 & ","
        strSql = strSql & lngNum & ",'"
        strSql = strSql & cboNote.Text & "',"
        strSql = strSql & "To_Date('" & dtpDate.Value & "','yyyy-mm-dd hh24:mi:ss'),"
        strSql = strSql & Val(txtStyle(5).Text) & ",'"
        strSql = strSql & UserInfo.姓名 & "','"
        strSql = strSql & UserInfo.编号 & "')"
    Else
        strSql = "zl_病人预约登记_Insert("
        strSql = strSql & mrsInfo!病人ID & ","
        strSql = strSql & cboArrangeNo.ItemData(cboArrangeNo.ListIndex) & ","
        strSql = strSql & byt复诊方式 & ","
        strSql = strSql & lngNum & ",'"
        strSql = strSql & cboNote.Text & "',"
        strSql = strSql & "To_Date('" & dtpDate.Value & "','yyyy-mm-dd hh24:mi:ss'),"
        strSql = strSql & Val(txtStyle(5).Text) & ",'"
        strSql = strSql & UserInfo.姓名 & "','"
        strSql = strSql & UserInfo.编号 & "')"
    End If
    
    Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    Unload Me
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '系统IC卡
        If Not mobjICCard Is Nothing Then
           txtPatient.Text = mobjICCard.Read_Card()
           If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        Call GetPatient(objCard, txtPatient.Text, True)
    End If
    
End Sub

Public Sub ShowMe(frmMain As Object, Optional lng病人ID As Long)
    mlng病人ID = lng病人ID
    Call LoadRegPlans
    dtpDate.Value = gobjDatabase.Currentdate
    Me.Show vbModal, frmMain
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible And cmdOK.Visible = True Then
        txtPatient.SetFocus
    Else
        If cmdOK.Visible = False And cmdCancel.Visible And cmdCancel.Enabled Then cmdCancel.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitIDKind
    IDKind.RaisEffect picMain, -1
    If mlng病人ID <> 0 Then
        Call GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False)
    End If
End Sub

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, strTemp As String
    Dim lngCardID As Long
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, 1260, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0;医|医保号|0;身|身份证号|0;门|门诊号|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim strValues As String, strArray() As String
    Set mrsInfo = Nothing
    Set mobjICCard = Nothing
    Set mobjIDCard = Nothing
    If cmdOK.Visible Then
        Call GetRegInFor(g私有模块, Me.Name, "strValues", strValues)
        If strValues = "" Then strValues = "1,14|1|2|14"
        If optStyle(0).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = txtStyle(0).Text & "," & txtStyle(4).Text & "|" & strArray(1) & "|" & strArray(2) & "|" & strArray(3)
        End If
        If optStyle(1).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & txtStyle(1).Text & "|" & strArray(2) & "|" & strArray(3)
        End If
        If optStyle(2).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & strArray(1) & "|" & txtStyle(2).Text & "|" & strArray(3)
        End If
        If optStyle(3).Value = True Then
            strArray = Split(strValues & "|||", "|")
            strValues = strArray(0) & "|" & strArray(1) & "|" & strArray(2) & "|" & txtStyle(3).Text
        End If
        
        Call SaveRegInFor(g私有模块, Me.Name, "strValues", strValues)
    End If
    mlng病人ID = 0
End Sub


Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.卡号
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
    txtPatient.Text = strCardNO
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    IDKind.IDKind = IDKind.GetKindIndex("身份证号")
    txtPatient.Text = strID
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub optStyle_Click(Index As Integer)
    Dim strValues As String, strArray() As String
    Call GetRegInFor(g私有模块, Me.Name, "strValues", strValues)
    If strValues = "" Then strValues = "1,14|1|2|14"
    strArray = Split(strValues & "|||", "|")
    If optStyle(0).Value = True Then
        txtStyle(0).Enabled = True
        udStyle(0).Enabled = True
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = True
        udStyle(4).Enabled = True
        If txtStyle(0).Text = "" Then
            txtStyle(0).Text = Split(strArray(0) & ",", ",")(0)
            txtStyle(4).Text = Split(strArray(0) & ",", ",")(1)
        End If
    End If
    If optStyle(1).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = True
        udStyle(1).Enabled = True
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(1).Text = "" Then
            txtStyle(1).Text = strArray(1)
        End If
    End If
    If optStyle(2).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = True
        udStyle(2).Enabled = True
        txtStyle(3).Enabled = False
        udStyle(3).Enabled = False
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(2).Text = "" Then
            txtStyle(2).Text = strArray(2)
        End If
    End If
    If optStyle(3).Value = True Then
        txtStyle(0).Enabled = False
        udStyle(0).Enabled = False
        txtStyle(1).Enabled = False
        udStyle(1).Enabled = False
        txtStyle(2).Enabled = False
        udStyle(2).Enabled = False
        txtStyle(3).Enabled = True
        udStyle(3).Enabled = True
        txtStyle(4).Enabled = False
        udStyle(4).Enabled = False
        If txtStyle(3).Text = "" Then
            txtStyle(3).Text = strArray(3)
        End If
    End If
    Call CaclDate
End Sub

Private Sub CaclDate()
    Dim intDays As Integer
    Dim strTemp As String
    If cmdOK.Visible = False Then Exit Sub
    strTemp = cboNote.Text
    cboNote.Clear
    If optStyle(0).Value = True Then
        intDays = Val(txtStyle(0).Text) * Val(txtStyle(4).Text)
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(0).Text) & "个疗程后复诊"
    End If
    If optStyle(1).Value = True Then
        intDays = Val(txtStyle(1).Text)
        dtpDate.Value = DateAdd("m", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(1).Text) & "个月后复诊"
    End If
    If optStyle(2).Value = True Then
        intDays = Val(txtStyle(2).Text) * 7
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(2).Text) & "周后复诊"
    End If
    If optStyle(3).Value = True Then
        intDays = Val(txtStyle(3).Text)
        dtpDate.Value = DateAdd("d", intDays, gobjDatabase.Currentdate)
        cboNote.AddItem Val(txtStyle(3).Text) & "天后复诊"
    End If
    cboNote.Text = strTemp
End Sub

Private Sub optStyle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    gobjCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub LoadRegPlans()
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errH
    strSql = "Select Rownum As 排序, a.Id, a.号码, a.科室, a.项目, a.医生姓名, a.医生id, a.科室id" & vbNewLine & _
            "From (Select a.Id, a.号码, b.名称 As 科室, c.名称 As 项目, a.医生姓名, a.医生id, a.科室id" & vbNewLine & _
            "       From 临床出诊号源 A, 部门表 B, 收费项目目录 C" & vbNewLine & _
            "       Where (a.医生id = [1] Or (a.医生id Is Null And a.科室id = [2])) And a.科室id = b.Id And a.项目id = c.Id And (c.撤档时间 Is Null Or c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
            "             (b.站点 = '1' Or b.站点 Is Null) And Exists (Select 1 From 临床出诊安排 M,临床出诊表 N Where M.号源ID = A.ID And M.出诊ID = N.ID And N.发布时间 Is Not Null)" & vbNewLine & _
            "       Order By Decode(a.医生id, [1], 1, 0) Desc, a.医生id, a.科室id, a.号码) A"
    Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID, UserInfo.部门ID)
    cboArrangeNo.Clear
    Do While Not mrsPlan.EOF
        cboArrangeNo.AddItem mrsPlan!号码 & "-" & mrsPlan!项目 & "(" & mrsPlan!医生姓名 & ")"
        cboArrangeNo.ItemData(cboArrangeNo.NewIndex) = 0
        mrsPlan.MoveNext
    Loop
    If cboArrangeNo.ListCount <> 0 Then cboArrangeNo.ListIndex = 0
    Exit Sub
errH:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    Dim rsTmp As ADODB.Recordset
    Dim cur余额 As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '结算模式(0-先结算后诊疗或1-先诊疗后结算)|挂号费收取方式(0-现收或1-记帐)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure, strAdvance)
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng病人ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    
    txtPatient.Text = "-" & lng病人ID
    Call txtPatient_Validate(False)    '其中的Setfocus调用使本事件(txtPatient_KeyPress)执行完后,不会再次自动执行txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str病人类型, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True
    
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    
    '医保验证
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    strKind = IDKind.GetCurCard.名称
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "挂号单"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean)
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '
    '         blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String, rsFeeType As ADODB.Recordset
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim bln医保号 As Boolean
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '非法卡类别

    strInputInfo = strInput
    
    On Error GoTo errH
    bln医保号 = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard

    strSql = "Select  A.病人ID,A.门诊号,A.住院号,A.就诊卡号,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号,A.其他证件,A.身份,A.职业,A.民族,A.病人类型, " & _
             "A.国籍,A.籍贯,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.户口地址, " & _
             "A.户口地址邮编,A.Email,A.QQ,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质,A.就诊时间,A.就诊状态, " & _
             "A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间,A.在院,A.IC卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号, " & _
             "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式 From 病人信息 A,保险类别 B  Where A.险类 = B.序号(+) And A.停用时间 is NULL  "

   
    If blnCard And objCard.名称 Like "姓名*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
'        Else
'            lng卡类别ID = gCurSendCard.lng卡类别ID
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0

        If lng病人ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSql = strSql & " And A.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSql = strSql & " And A.病人ID=[2]" & _
        IIf(mstrYBPati <> "", "", str非在院)
    ElseIf blnInputIDCard Then  '单独的身份证识别
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        strInput = "-" & lng病人ID
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = _
                    " Select distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A " & _
                    " Where Rownum <101 And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院
                    
                strPati = strPati & " Union ALL " & _
                        "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by 排序ID,姓名"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%")
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '当作新病人
                        txtPatient.Text = ""
                        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '以病人ID读取
                        strInput = rsTmp!病人ID
                        strSql = strSql & " And A.病人ID=[1]"
                    End If
                Else '取消选择
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "医保号"
                strInput = UCase(strInput)
                bln医保号 = True
                strSql = strSql & " And A.医保号=[1]" & str非在院
                
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
                 
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.门诊号=[1]" & str非在院
             Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strSql = strSql & " And A.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "病人身份验证失败！", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        txtPatient.Text = Nvl(mrsInfo!姓名) '会调用Change事件
        txtPatient.BackColor = &H80000005
        '在调用txtPatient_Change事件后在门诊号和病人姓名都为空的情况下 无法识别该病人信息 出现错误
        '对这类数据库数据错误不再进行后续的处理
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(mintInsure) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        txtGender.Text = Nvl(mrsInfo!性别)
        txtBirth.Text = Format(Nvl(mrsInfo!出生日期), "yyyy-MM-dd hh:mm")
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        txtAge.Text = Nvl(mrsInfo!年龄)
        txtClinic.Text = Nvl(mrsInfo!门诊号)
        txt登记人.Text = UserInfo.姓名
        txt登记时间.Text = Format(gobjDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
        
        If cboArrangeNo.Enabled And cboArrangeNo.Visible Then cboArrangeNo.SetFocus
    Else
NewPati:
        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    txtGender.Text = ""
    txtAge.Text = ""
    txtBirth.Text = ""
    optStyle(0).Value = 1
    mintInsure = 0
    Set mrsInfo = Nothing
End Sub

Private Sub txtStyle_Change(Index As Integer)
    Call CaclDate
End Sub

Private Sub txtStyle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Index <> 5 Then
            If cboNote.Visible And cboNote.Enabled Then cboNote.SetFocus
        Else
            Call gobjCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub
