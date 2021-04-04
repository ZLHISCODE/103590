VERSION 5.00
Begin VB.Form frm医疗保险申报单 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医疗保险申报单"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   Icon            =   "frm医疗保险申报单.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "包干结算"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   5
      Left            =   4950
      TabIndex        =   55
      Top             =   4080
      Width           =   4665
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   3270
         TabIndex        =   65
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   59
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   57
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt统筹基金 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   3270
         TabIndex        =   61
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt大额统筹 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   63
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   2490
         TabIndex        =   64
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   58
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   5
         Left            =   270
         TabIndex        =   56
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   2475
         TabIndex        =   60
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl大额统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   5
         Left            =   255
         TabIndex        =   62
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "日包干住院"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   4
      Left            =   180
      TabIndex        =   24
      Top             =   4080
      Width           =   4635
      Begin VB.TextBox txt住院天数 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   28
         Top             =   690
         Width           =   585
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   26
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   1020
         TabIndex        =   30
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   3270
         TabIndex        =   32
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl住院天数 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院天数"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   27
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   240
         TabIndex        =   29
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   4
         Left            =   2460
         TabIndex        =   31
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "重症住院"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   3
      Left            =   4950
      TabIndex        =   44
      Top             =   2370
      Width           =   4665
      Begin VB.TextBox txt大额统筹 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   52
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt统筹基金 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   3270
         TabIndex        =   50
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   46
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   1050
         TabIndex        =   48
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   3270
         TabIndex        =   54
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl大额统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   255
         TabIndex        =   51
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   2475
         TabIndex        =   49
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   45
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   270
         TabIndex        =   47
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   3
         Left            =   2490
         TabIndex        =   53
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "控制线住院"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   2
      Left            =   4950
      TabIndex        =   33
      Top             =   720
      Width           =   4635
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   43
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   37
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   35
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt统筹基金 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   3240
         TabIndex        =   39
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt大额统筹 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   2
         Left            =   1020
         TabIndex        =   41
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   2460
         TabIndex        =   42
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   2445
         TabIndex        =   38
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl大额统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   40
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "特殊门诊"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   2370
      Width           =   4635
      Begin VB.TextBox txt大额统筹 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   21
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt统筹基金 
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   3225
         TabIndex        =   19
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   15
         Top             =   300
         Width           =   585
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   1020
         TabIndex        =   17
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   3225
         TabIndex        =   23
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl大额统筹 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "大额统筹"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   20
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl统筹基金 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "统筹基金"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2430
         TabIndex        =   18
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   1
         Left            =   2430
         TabIndex        =   22
         Top             =   1140
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "普通门诊"
      Enabled         =   0   'False
      Height          =   1575
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   720
      Width           =   4635
      Begin VB.TextBox txt医疗补助 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   12
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox txt个人帐户 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   10
         Top             =   690
         Width           =   1155
      End
      Begin VB.TextBox txt就诊人次 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   0
         Left            =   1020
         TabIndex        =   8
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lbl医疗补助 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医疗补助"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lbl个人帐户 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lbl就诊人次 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊人次"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd申报 
      Caption         =   "申报(&O)"
      Height          =   350
      Left            =   6360
      TabIndex        =   5
      Top             =   210
      Width           =   1100
   End
   Begin VB.CommandButton cmd取数 
      Caption         =   "取数(&D)"
      Height          =   350
      Left            =   5220
      TabIndex        =   4
      Top             =   210
      Width           =   1100
   End
   Begin VB.ComboBox cbo保险类别 
      Height          =   300
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1665
   End
   Begin VB.ComboBox cbo期号 
      Height          =   300
      Left            =   690
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   1665
   End
   Begin VB.Label lbl保险类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "保险类别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2610
      TabIndex        =   2
      Top             =   300
      Width           =   720
   End
   Begin VB.Label lbl期号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "期号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   270
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frm医疗保险申报单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngID As Long              '0-新增;非零表示查阅
Private mblnOK As Boolean           '编辑成功

Private Enum 分类
    普通门诊
    特殊门诊
    控制线住院
    重症住院
    日包干住院
    包干结算
End Enum
'1、医疗保险申报清单中，门诊人次是指的普通门诊的人次？包干结算就诊人次是指普通门诊中选择的单病种的部分数据
'   a、控制线（清算=1），重症（清算=2），按日包干（清算=4），包干（清算=6）
'   b、普通门诊选择了单病种的就是门诊包干

Public Function ShowME(ByVal lngID As Long) As Boolean
    mblnOK = False
    mlngID = lngID
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmd取数_Click()
    Dim int保险类别 As Integer
    Dim lng总天数 As Long, lng住院天数 As Long
    Dim str期号 As String, str开始日期 As String, str结束日期 As String, str上期结束日期 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If mlngID <> 0 Then
        '查阅模式
        Unload Me
        Exit Sub
    End If
    
    '清空
    Call ClearCons
    
    str期号 = Me.cbo期号.Text
    int保险类别 = Me.cbo保险类别.ItemData(Me.cbo保险类别.ListIndex)
    str开始日期 = Mid(str期号, 1, 4) & "-" & Mid(str期号, 5, 2) & "-01 00:00:00"
    gstrSQL = " SELECT last_day(to_date('" & Mid(str开始日期, 1, 10) & "','yyyy-MM-dd')) from dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取月度最后一天")
    str结束日期 = Format(rsTemp.Fields(0).Value, "yyyy-MM-dd") & " 23:59:59"
    str上期结束日期 = Format(DateAdd("d", -1, str开始日期), "yyyy-MM-dd")
    
    '根据设定的条件取数
    '并发症保存的是保险类别
    '1、普通门诊（当期只收费，人次=1；当期收费且退费，人次=1；不允许跨月退费，因此不考虑）
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 门诊人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.医疗类别='11' AND A.性质=1 AND B.单病种编码_结算 IS NULL " & _
             " AND A.记录ID=C.结帐ID And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "普通门诊", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(普通门诊).Text = Format(rsTemp!门诊人次, "#0;-#0; ;")
    Me.txt个人帐户(普通门诊).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt医疗补助(普通门诊).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    
    '2、特殊门诊
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT B.医保号) AS 门诊人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金, " & _
             "        NVL(SUM(DECODE(C.结算方式,'大病统筹',NVL(C.冲预交,0),0)),0) AS 大额统筹, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,保险帐户 B,病人预交记录 C " & _
             " WHERE A.病人ID=B.病人ID And A.记录ID=C.结帐ID AND A.医疗类别='18' AND A.性质=1 " & _
             " And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "特殊门诊", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(特殊门诊).Text = Format(rsTemp!门诊人次, "#0;-#0; ;")
    Me.txt个人帐户(特殊门诊).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt统筹基金(特殊门诊).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    Me.txt大额统筹(特殊门诊).Text = Format(rsTemp!大额统筹, "#0.00;-#0.00; ;")
    Me.txt医疗补助(特殊门诊).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    
    '3、控制线住院
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 住院人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金, " & _
             "        NVL(SUM(DECODE(C.结算方式,'大病统筹',NVL(C.冲预交,0),0)),0) AS 大额统筹, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID AND A.性质=2 " & _
             " AND B.清算方式=1 And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "控制线住院", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(控制线住院).Text = Format(rsTemp!住院人次, "#0;-#0; ;")
    Me.txt个人帐户(控制线住院).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt统筹基金(控制线住院).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    Me.txt大额统筹(控制线住院).Text = Format(rsTemp!大额统筹, "#0.00;-#0.00; ;")
    Me.txt医疗补助(控制线住院).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    
    '4、重症住院
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 住院人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金, " & _
             "        NVL(SUM(DECODE(C.结算方式,'大病统筹',NVL(C.冲预交,0),0)),0) AS 大额统筹, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID AND A.性质=2 " & _
             " AND B.清算方式=2 And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "重症住院", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(重症住院).Text = Format(rsTemp!住院人次, "#0;-#0; ;")
    Me.txt个人帐户(重症住院).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt统筹基金(重症住院).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    Me.txt大额统筹(重症住院).Text = Format(rsTemp!大额统筹, "#0.00;-#0.00; ;")
    Me.txt医疗补助(重症住院).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    
    '5、日包干住院
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 住院人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID AND A.性质=2 " & _
             " AND B.清算方式=4 And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "日包干住院", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(日包干住院).Text = Format(rsTemp!住院人次, "#0;-#0; ;")
    Me.txt个人帐户(日包干住院).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt医疗补助(日包干住院).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    '计算住院天数的原则（算出不算入：8月31号入院未出院的，住院天数为零天；如果是31号入，出院，则住院天数为1天）：
    '1、按期号来算:当月最后一天减去入院时间
    '2、该病人当期出院:出院时间减上期最后一天
    '3、当期入出院的:出院时间减去入院时间
    '4、当期仍未出院:当月最后一天减去上期最后一天
    gstrSQL = " SELECT DISTINCT" & _
             "      A.就诊流水号,C.入院日期,C.出院日期,TO_CHAR(C.入院日期,'YYYYMM') AS 入院期号 " & _
             "  FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病案主页 C  " & _
             "  WHERE A.记录ID=B.结帐ID AND A.病人ID=C.病人ID And A.主页ID=C.主页ID AND A.性质=2  " & _
             "  AND B.清算方式=4 And A.并发症=[1] And A.险类=[2]" & _
             "  AND A.结算时间 BETWEEN [3] AND [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "日包干住院天数", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    With rsTemp
        Do While Not .EOF
            If !入院期号 <> str期号 Then
                '当期以前入的院
                If Not IsNull(!出院日期) Then
                    '当期出院
                    lng住院天数 = DateDiff("d", str上期结束日期, !出院日期)
                Else
                    lng住院天数 = DateDiff("d", str上期结束日期, str结束日期)
                End If
            Else
                '当期入院
                If Not IsNull(!出院日期) Then
                    '当期出院
                    lng住院天数 = DateDiff("d", !入院日期, !出院日期)
                Else
                    lng住院天数 = DateDiff("d", !入院日期, str结束日期)
                End If
            End If
            If lng住院天数 = 0 And Not IsNull(!出院日期) Then lng住院天数 = 1
            lng总天数 = lng总天数 + lng住院天数
        Loop
    End With
    Me.txt住院天数(日包干住院).Text = Format(lng总天数, "#0;-#0; ;")
    
    '6、包干结算（普通门诊中选择了单病种的，加上住院结算中清算方式=6的）
    gstrSQL = "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 包干人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金, " & _
             "        NVL(SUM(DECODE(C.结算方式,'大病统筹',NVL(C.冲预交,0),0)),0) AS 大额统筹, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.医疗类别='11' AND A.性质=1 AND B.单病种编码_结算 IS Not NULL " & _
             " AND A.记录ID=C.结帐ID And A.并发症=" & int保险类别 & " And A.险类=" & TYPE_贵阳市 & _
             " AND A.结算时间 BETWEEN TO_DATE('" & str开始日期 & "','YYYY-MM-DD HH24:MI:SS') " & _
             " AND TO_DATE('" & str结束日期 & "','YYYY-MM-DD HH24:MI:SS') "
    gstrSQL = gstrSQL & _
             "UNION " & _
             "SELECT  " & _
             "        COUNT(DISTINCT A.就诊流水号) AS 包干人次, " & _
             "        NVL(SUM(DECODE(C.结算方式,'个人帐户',NVL(C.冲预交,0),0)),0) AS 个人帐户, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医保基金',NVL(C.冲预交,0),0)),0) AS 统筹基金, " & _
             "        NVL(SUM(DECODE(C.结算方式,'大病统筹',NVL(C.冲预交,0),0)),0) AS 大额统筹, " & _
             "        NVL(SUM(DECODE(C.结算方式,'医疗补助',NVL(C.冲预交,0),0)),0) AS 医疗补助 " & _
             " FROM 保险结算记录 A,ZLGYYB.结算附加信息 B,病人预交记录 C " & _
             " WHERE A.记录ID=B.结帐ID AND A.记录ID=C.结帐ID AND A.性质=2 " & _
             " AND B.清算方式=6 And A.并发症=[1] And A.险类=[2]" & _
             " AND A.结算时间 BETWEEN [3] AND [4]"
    gstrSQL = " SELECT SUM(包干人次) AS 包干人次,SUM(个人帐户) AS 个人帐户,SUM(统筹基金) AS 统筹基金," & _
              "       SUM(大额统筹) AS 大额统筹,SUM(医疗补助) AS 医疗补助" & _
              " FROM (" & gstrSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "包干结算", int保险类别, TYPE_贵阳市, CDate(str开始日期), CDate(str结束日期))
    Me.txt就诊人次(包干结算).Text = Format(rsTemp!包干人次, "#0;-#0; ;")
    Me.txt个人帐户(包干结算).Text = Format(rsTemp!个人帐户, "#0.00;-#0.00; ;")
    Me.txt统筹基金(包干结算).Text = Format(rsTemp!统筹基金, "#0.00;-#0.00; ;")
    Me.txt大额统筹(包干结算).Text = Format(rsTemp!大额统筹, "#0.00;-#0.00; ;")
    Me.txt医疗补助(包干结算).Text = Format(rsTemp!医疗补助, "#0.00;-#0.00; ;")
    
    Me.Tag = 1
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call ClearCons
End Sub

Private Sub cmd申报_Click()
    Dim str流水号 As String
    On Error GoTo errHand
    
    If Val(Me.Tag) = 0 Then
        MsgBox "请指定条件后点“取数”按钮！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnGYYB.BeginTrans
    '对XML DomDocument对象进行初始化
    If InitXML = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    '住院虚拟结算只要求传入个人编码，正式结算时才要求传入磁卡数据及密码
    Call InsertChild(mdomInput.documentElement, "PERIOD", cbo期号.Text)
    Call InsertChild(mdomInput.documentElement, "OPERATOR", gstrUserName)
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
    Call InsertChild(mdomInput.documentElement, "INSURETYPE", cbo保险类别.ItemData(cbo保险类别.ListIndex))
    Call InsertChild(mdomInput.documentElement, "MZPSNS", Val(txt就诊人次(普通门诊).Text))                ' 门诊就诊人次
    Call InsertChild(mdomInput.documentElement, "MZACCT", Val(txt个人帐户(普通门诊).Text))
    Call InsertChild(mdomInput.documentElement, "MZFUND3", Val(txt医疗补助(普通门诊).Text))
    Call InsertChild(mdomInput.documentElement, "TMPSNS", Val(txt就诊人次(特殊门诊).Text))
    Call InsertChild(mdomInput.documentElement, "TMACCT", Val(txt个人帐户(特殊门诊).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND1", Val(txt统筹基金(特殊门诊).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND2", Val(txt大额统筹(特殊门诊).Text))
    Call InsertChild(mdomInput.documentElement, "TMFUND3", Val(txt医疗补助(特殊门诊).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1PSNS", Val(txt就诊人次(控制线住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1ACCT", Val(txt个人帐户(控制线住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND1", Val(txt统筹基金(控制线住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND2", Val(txt大额统筹(控制线住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY1FUND3", Val(txt医疗补助(控制线住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2PSNS", Val(txt就诊人次(重症住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2ACCT", Val(txt个人帐户(重症住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND1", Val(txt统筹基金(重症住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND2", Val(txt大额统筹(重症住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY2FUND3", Val(txt医疗补助(重症住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3PSNS", Val(txt就诊人次(日包干住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3DAYS", Val(txt住院天数(日包干住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3ACCT", Val(txt个人帐户(日包干住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY3FUND3", Val(txt医疗补助(日包干住院).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4PSNS", Val(txt就诊人次(包干结算).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4ACCT", Val(txt个人帐户(包干结算).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND1", Val(txt统筹基金(包干结算).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND2", Val(txt大额统筹(包干结算).Text))
    Call InsertChild(mdomInput.documentElement, "ZY4FUND3", Val(txt医疗补助(包干结算).Text))
    '调用接口
    If CommRecServer("APPRECM") = False Then
        gcnGYYB.RollbackTrans
        Exit Sub
    End If
    str流水号 = GetElemnetValue("APPNO")
    
    '产生数据
    mlngID = GetNextID("清算单", gcnGYYB)
    gstrSQL = "ZL_清算单_INSERT(" & mlngID & ",0,'" & Me.cbo期号.Text & "'," & Me.cbo保险类别.ItemData(cbo保险类别.ListIndex) & "," & _
        "'" & Me.cbo保险类别.Text & "','" & gstrUserName & "',sysdate,'" & str流水号 & "',NULL)"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    
    gstrSQL = "ZL_基本医疗清算明细_INSERT(" & mlngID & "," & Val(txt就诊人次(普通门诊).Text) & "," & Val(txt个人帐户(普通门诊).Text) & "," & Val(txt医疗补助(普通门诊).Text) & "," & _
            Val(txt就诊人次(特殊门诊).Text) & "," & Val(txt个人帐户(特殊门诊).Text) & "," & Val(txt统筹基金(特殊门诊).Text) & "," & Val(txt大额统筹(特殊门诊).Text) & "," & Val(txt医疗补助(特殊门诊).Text) & "," & _
            Val(txt就诊人次(控制线住院).Text) & "," & Val(txt个人帐户(控制线住院).Text) & "," & Val(txt统筹基金(控制线住院).Text) & "," & Val(txt大额统筹(控制线住院).Text) & "," & Val(txt医疗补助(控制线住院).Text) & "," & _
            Val(txt就诊人次(重症住院).Text) & "," & Val(txt个人帐户(重症住院).Text) & "," & Val(txt统筹基金(重症住院).Text) & "," & Val(txt大额统筹(重症住院).Text) & "," & Val(txt医疗补助(重症住院).Text) & "," & _
            Val(txt就诊人次(日包干住院).Text) & "," & Val(txt住院天数(日包干住院).Text) & "," & Val(txt个人帐户(日包干住院).Text) & "," & Val(txt医疗补助(日包干住院).Text) & "," & _
            Val(txt就诊人次(包干结算).Text) & "," & Val(txt个人帐户(包干结算).Text) & "," & Val(txt统筹基金(包干结算).Text) & "," & Val(txt大额统筹(包干结算).Text) & "," & Val(txt医疗补助(包干结算).Text) & ")"
    gcnGYYB.Execute gstrSQL, , adCmdStoredProc
    gcnGYYB.CommitTrans
    
    mblnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnGYYB.RollbackTrans
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim curDate As Date
    Dim str上月 As String, str本月 As String
    Dim rsData As New ADODB.Recordset
    
    If mlngID = 0 Then
        With cbo保险类别
            .Clear
            .AddItem "企业职工基本医疗保险"
            .ItemData(.NewIndex) = 1
            .AddItem "企业离休医疗保险"
            .ItemData(.NewIndex) = 2
            .AddItem "机关事业单位医疗保险"
            .ItemData(.NewIndex) = 3
            .AddItem "居民"
            .ItemData(.NewIndex) = 6
            .ListIndex = 0
        End With
        
        '缺省只装入上月、本月供申报
        curDate = zlDatabase.Currentdate()
        str上月 = Format(DateAdd("m", -1, curDate), "yyyyMM")
        str本月 = Format(curDate, "yyyyMM")
        With cbo期号
            .Clear
            .AddItem str上月
            .AddItem str本月
            .ListIndex = 0
        End With
        Exit Sub
    End If
    
    '读取申报单数据
    gstrSQL = "SELECT  " & _
             "        A.ID, A.期号, A.保险类别, A.操作员, A.日期 ,B.门诊人次, B.门诊个人帐户, B.门诊医疗补助, B.特殊门诊人次, B.特殊门诊个人帐户, B.特殊门诊基本统筹, B.特殊门诊大额统筹,  " & _
             "        B.特殊门诊医疗补助, B.控制线住院人次, B.控制线住院个人帐户, B.控制线住院基本统筹, B.控制线住院大额统筹, B.控制线住院医疗补助,  " & _
             "        B.重症住院人次, B.重症住院个人帐户, B.重症住院基本统筹, B.重症住院大额统筹, B.重症住院医疗补助, B.日包干住院人次, B.日包干住院天数,  " & _
             "        B.日包干住院个人帐户, 日包干住院医疗补助, B.包干结算人次, B.包干结算个人帐户, B.包干结算基本统筹, B.包干结算大额统筹, B.包干结算医疗补助, A.清算流水号, A.处理情况 " & _
             " FROM 清算单 A, 基本医疗清算明细 B " & _
             " WHERE A.ID=B.清算单ID AND A.ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "读取申报单数据", mlngID)
    
    '填数
    With rsData
        Me.cbo保险类别.AddItem !保险类别
        Me.cbo保险类别.ListIndex = 0
        Me.cbo期号.AddItem !期号
        Me.cbo期号.ListIndex = 0
        
        Me.txt就诊人次(普通门诊).Text = Format(Nvl(!门诊人次, 0), "#0;-#0; ;")
        Me.txt个人帐户(普通门诊).Text = Format(Nvl(!门诊个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(普通门诊).Text = Format(Nvl(!门诊医疗补助, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(特殊门诊).Text = Format(Nvl(!特殊门诊人次, 0), "#0;-#0; ;")
        Me.txt个人帐户(特殊门诊).Text = Format(Nvl(!特殊门诊个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(特殊门诊).Text = Format(Nvl(!特殊门诊基本统筹, 0), "#0.00;-#0.00; ;")
        Me.txt大额统筹(特殊门诊).Text = Format(Nvl(!特殊门诊大额统筹, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(特殊门诊).Text = Format(Nvl(!特殊门诊医疗补助, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(控制线住院).Text = Format(Nvl(!控制线住院人次, 0), "#0;-#0; ;")
        Me.txt个人帐户(控制线住院).Text = Format(Nvl(!控制线住院个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(控制线住院).Text = Format(Nvl(!控制线住院基本统筹, 0), "#0.00;-#0.00; ;")
        Me.txt大额统筹(控制线住院).Text = Format(Nvl(!控制线住院大额统筹, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(控制线住院).Text = Format(Nvl(!控制线住院医疗补助, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(重症住院).Text = Format(Nvl(!重症住院人次, 0), "#0;-#0; ;")
        Me.txt个人帐户(重症住院).Text = Format(Nvl(!重症住院个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(重症住院).Text = Format(Nvl(!重症住院基本统筹, 0), "#0.00;-#0.00; ;")
        Me.txt大额统筹(重症住院).Text = Format(Nvl(!重症住院大额统筹, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(重症住院).Text = Format(Nvl(!重症住院医疗补助, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(日包干住院).Text = Format(Nvl(!日包干住院人次, 0), "#0;-#0; ;")
        Me.txt住院天数(日包干住院).Text = Format(Nvl(!日包干住院天数, 0), "#0;-#0; ;")
        Me.txt个人帐户(日包干住院).Text = Format(Nvl(!日包干住院个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(日包干住院).Text = Format(Nvl(!日包干住院医疗补助, 0), "#0.00;-#0.00; ;")
        
        Me.txt就诊人次(包干结算).Text = Format(Nvl(!包干结算人次, 0), "#0;-#0; ;")
        Me.txt个人帐户(包干结算).Text = Format(Nvl(!包干结算个人帐户, 0), "#0.00;-#0.00; ;")
        Me.txt统筹基金(包干结算).Text = Format(Nvl(!包干结算基本统筹, 0), "#0.00;-#0.00; ;")
        Me.txt大额统筹(包干结算).Text = Format(Nvl(!包干结算大额统筹, 0), "#0.00;-#0.00; ;")
        Me.txt医疗补助(包干结算).Text = Format(Nvl(!包干结算医疗补助, 0), "#0.00;-#0.00; ;")
    End With
    
    '设置控件状态
    Me.cbo保险类别.Enabled = False
    Me.cbo期号.Enabled = False
    
    cmd申报.Visible = False
    cmd取数.Caption = "退出(&X)"
End Sub

Private Sub ClearCons()
    Me.Tag = ""
    Me.txt就诊人次(普通门诊).Text = ""
    Me.txt个人帐户(普通门诊).Text = ""
    Me.txt医疗补助(普通门诊).Text = ""
    
    Me.txt就诊人次(特殊门诊).Text = ""
    Me.txt个人帐户(特殊门诊).Text = ""
    Me.txt统筹基金(特殊门诊).Text = ""
    Me.txt大额统筹(特殊门诊).Text = ""
    Me.txt医疗补助(特殊门诊).Text = ""
    
    Me.txt就诊人次(控制线住院).Text = ""
    Me.txt个人帐户(控制线住院).Text = ""
    Me.txt统筹基金(控制线住院).Text = ""
    Me.txt大额统筹(控制线住院).Text = ""
    Me.txt医疗补助(控制线住院).Text = ""
    
    Me.txt就诊人次(重症住院).Text = ""
    Me.txt个人帐户(重症住院).Text = ""
    Me.txt统筹基金(重症住院).Text = ""
    Me.txt大额统筹(重症住院).Text = ""
    Me.txt医疗补助(重症住院).Text = ""
    
    Me.txt就诊人次(日包干住院).Text = ""
    Me.txt住院天数(日包干住院).Text = ""
    Me.txt个人帐户(日包干住院).Text = ""
    Me.txt医疗补助(日包干住院).Text = ""
    
    Me.txt就诊人次(包干结算).Text = ""
    Me.txt个人帐户(包干结算).Text = ""
    Me.txt统筹基金(包干结算).Text = ""
    Me.txt大额统筹(包干结算).Text = ""
    Me.txt医疗补助(包干结算).Text = ""
End Sub
