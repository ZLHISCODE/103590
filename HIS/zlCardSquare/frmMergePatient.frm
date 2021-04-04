VERSION 5.00
Begin VB.Form frmMergePatient 
   Caption         =   "病人合并"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmMergePatient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdTurn 
      Caption         =   "交换(&T)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   270
      Width           =   1100
   End
   Begin VB.CommandButton cmdPati 
      Height          =   330
      Left            =   4950
      Picture         =   "frmMergePatient.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "选择病人(F2)"
      Top             =   15
      Width           =   420
   End
   Begin VB.Frame fra 
      Caption         =   "要保留的病人信息      "
      Height          =   4725
      Index           =   1
      Left            =   3330
      TabIndex        =   6
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt状态 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   270
         Width           =   2025
      End
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   525
         Width           =   2025
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   780
         Width           =   2025
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2025
      End
      Begin VB.TextBox txt出生日期 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2025
      End
      Begin VB.TextBox txt国籍 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2025
      End
      Begin VB.TextBox txt民族 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2025
      End
      Begin VB.TextBox txt学历 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2025
      End
      Begin VB.TextBox txt婚姻状况 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2025
      End
      Begin VB.TextBox txt职业 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2025
      End
      Begin VB.TextBox txt身份 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2025
      End
      Begin VB.TextBox txt身份证号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2025
      End
      Begin VB.TextBox txt出生地点 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2025
      End
      Begin VB.TextBox txt家庭地址 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2025
      End
      Begin VB.TextBox txt科室 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3855
         Width           =   2025
      End
      Begin VB.TextBox txt床位 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   4110
         Width           =   2025
      End
      Begin VB.TextBox txt住院次数 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1005
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2025
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   74
         Top             =   780
         Width           =   450
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   73
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   72
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   71
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   70
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   69
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label lbl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   68
         Top             =   2325
         Width           =   810
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   67
         Top             =   2835
         Width           =   450
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   66
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   65
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   64
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   63
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   62
         Top             =   3855
         Width           =   450
      End
      Begin VB.Label lbl床位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   61
         Top             =   4110
         Width           =   450
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   60
         Top             =   4380
         Width           =   810
      End
      Begin VB.Label lbl状态 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   59
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   315
         TabIndex        =   58
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "被合并的病人信息"
      Height          =   4725
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt住院次数 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2025
      End
      Begin VB.TextBox txt床位 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4110
         Width           =   2025
      End
      Begin VB.TextBox txt科室 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3855
         Width           =   2025
      End
      Begin VB.TextBox txt家庭地址 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2025
      End
      Begin VB.TextBox txt出生地点 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2025
      End
      Begin VB.TextBox txt身份证号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2025
      End
      Begin VB.TextBox txt身份 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2025
      End
      Begin VB.TextBox txt职业 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2025
      End
      Begin VB.TextBox txt婚姻状况 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2025
      End
      Begin VB.TextBox txt学历 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2025
      End
      Begin VB.TextBox txt民族 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2025
      End
      Begin VB.TextBox txt国籍 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2025
      End
      Begin VB.TextBox txt出生日期 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2025
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2025
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   780
         Width           =   2025
      End
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   525
         Width           =   2025
      End
      Begin VB.TextBox txt状态 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   23
         Top             =   525
         Width           =   630
      End
      Begin VB.Label lbl状态 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   22
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   4380
         Width           =   810
      End
      Begin VB.Label lbl床位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   20
         Top             =   4110
         Width           =   450
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   19
         Top             =   3855
         Width           =   450
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   3345
         Width           =   810
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   15
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   14
         Top             =   2835
         Width           =   450
      End
      Begin VB.Label lbl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   2325
         Width           =   810
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   12
         Top             =   2055
         Width           =   450
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   10
         Top             =   1545
         Width           =   450
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   1290
         Width           =   810
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   8
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   7
         Top             =   780
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "合并(&M)"
      Height          =   350
      Left            =   6675
      TabIndex        =   2
      Top             =   795
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6675
      TabIndex        =   3
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6675
      TabIndex        =   4
      Top             =   4095
      Width           =   1100
   End
End
Attribute VB_Name = "frmMergePatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long
Private mlng被合并病人ID As Long
Private mblnOK As Boolean
Public Function zlShowPatiMerge(ByVal frmMain As Object, _
    ByVal lng被合并病人ID As Long, ByRef lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:病人合并
    '入参:   lng病人ID-合并的病人ID
    '           lng被合病人ID-被合并的病人ID
    '出参:
    '           lng病人ID-合并后的病人ID
    '出参:strOutput-应答数据
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-08-21 15:43:37
    '说明:
    '问题:52913
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng被合并病人ID = lng被合并病人ID
    mlng病人ID = lng病人ID
    mblnOK = False
    Me.Show 1, frmMain
    lng病人ID = mlng病人ID
    zlShowPatiMerge = mblnOK
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdPati_Click()
    frmPatiSel.mstrPrivs = ""
    frmPatiSel.Show 1, Me
    
    If frmPatiSel.mlng病人ID <> 0 Then
        If ExistFeeInsurePatient(frmPatiSel.mlng病人ID) Then
            MsgBox "该医保病人存在未结费用,请先结清后再合并！", vbExclamation, gstrSysName: Exit Sub
        End If
    
        Call ShowPatiInfo(frmPatiSel.mlng病人ID, 1)
    End If
End Sub

Private Sub cmdTurn_Click()
    Dim lngTmp As Long
    
    If Val(fra(1).Tag) = 0 Then
        If glngSys Like "8??" Then
            MsgBox "没有设置要保留的客户,请先选择一个客户！", vbInformation, gstrSysName
        Else
            MsgBox "没有设置要保留的病人,请先选择一个病人！", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
    If Val(fra(1).Tag) = Val(fra(0).Tag) Then
        If glngSys Like "8??" Then
            MsgBox "你选择了同一个客户,请选择其它客户！", vbInformation, gstrSysName
        Else
            MsgBox "你选择了同一个病人,请选择其它病人！", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
    
    lngTmp = fra(0).Tag
    Call ShowPatiInfo(CLng(fra(1).Tag), 0)
    Call ShowPatiInfo(lngTmp, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            cmdPati_Click
    End Select
End Sub

Private Sub Form_Load()
    If glngSys Like "8??" Then
        Caption = "客户合并"
        lbl住院号(0).Visible = False
        lbl住院号(1).Visible = False
        txt住院号(0).Visible = False
        txt住院号(1).Visible = False
        
        lbl科室(0).Visible = False
        lbl科室(1).Visible = False
        lbl床位(0).Visible = False
        lbl床位(1).Visible = False
        lbl住院次数(0).Visible = False
        lbl住院次数(1).Visible = False
    
        txt科室(0).Visible = False
        txt科室(1).Visible = False
        txt床位(0).Visible = False
        txt床位(1).Visible = False
        txt住院次数(0).Visible = False
        txt住院次数(1).Visible = False
        
        fra(0).Caption = "被合并的客户信息"
        fra(1).Caption = "要保留的客户信息"
        
        fra(0).Height = fra(0).Height - 750
        fra(1).Height = fra(1).Height - 750
        Me.Height = Me.Height - 750
        
        cmdExit.Top = cmdExit.Top - 750
        cmdHelp.Top = cmdHelp.Top - 750
    End If
    
    fra(1).Tag = ""
    If Not ShowPatiInfo(mlng病人ID, 0) Then Unload Me: Exit Sub
    If Not ShowPatiInfo(mlng被合并病人ID, 1) Then Unload Me: Exit Sub
End Sub

Private Sub ClearPatiInfo(X As Integer)
'功能：清除一个病人信息
'参数：x=控件索引,0=源病人,1=目标病人
    txt姓名(X).Text = ""
    txt性别(X).Text = ""
    txt出生日期(X).Text = ""
    txt国籍(X).Text = ""
    txt民族(X).Text = ""
    txt学历(X).Text = ""
    txt身份(X).Text = ""
    txt职业(X).Text = ""
    txt身份证号(X).Text = ""
    txt出生地点(X).Text = ""
    txt家庭地址(X).Text = ""
    txt婚姻状况(X).Text = ""
    txt状态(X).Text = ""
    lbl住院号(X).Caption = "住院号:"
    txt住院号(X).Text = ""
    txt科室(X).Text = ""
    txt床位(X).Text = ""
    txt住院次数(X).Text = ""
    fra(X).Tag = ""
End Sub

Private Function ShowPatiInfo(lngID As Long, X As Integer) As Boolean
'功能：显示一个病人信息
'参数：lngID=病人ID,x=控件索引,0=源病人,1=目标病人
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim str住院号 As String, str门诊号 As String
    
    On Error GoTo errH
    'by lesfeng 2010-03-08 性能优化 select *
    strSQL = "Select 病人ID,门诊号,住院号,就诊卡号,姓名,性别,年龄,出生日期,出生地点,身份证号,身份,职业,民族,国籍,区域,学历,婚姻状况,家庭地址,家庭电话" & _
             "  From 病人信息 Where 病人ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    If rsTmp.EOF Then
        MsgBox "未发现该病人的身份信息！", vbInformation, gstrSysName
        Exit Function
    End If

    txt姓名(X).Text = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
    txt性别(X).Text = IIf(IsNull(rsTmp!性别), "", rsTmp!性别)
    txt出生日期(X).Text = Format(IIf(IsNull(rsTmp!出生日期), "", rsTmp!出生日期), "yyyy年MM月dd日")
    txt国籍(X).Text = IIf(IsNull(rsTmp!国籍), "", rsTmp!国籍)
    txt民族(X).Text = IIf(IsNull(rsTmp!民族), "", rsTmp!民族)
    txt学历(X).Text = IIf(IsNull(rsTmp!学历), "", rsTmp!学历)
    txt身份(X).Text = IIf(IsNull(rsTmp!身份), "", rsTmp!身份)
    txt职业(X).Text = IIf(IsNull(rsTmp!职业), "", rsTmp!职业)
    txt身份证号(X).Text = IIf(IsNull(rsTmp!身份证号), "", rsTmp!身份证号)
    txt出生地点(X).Text = IIf(IsNull(rsTmp!出生地点), "", rsTmp!出生地点)
    txt家庭地址(X).Text = IIf(IsNull(rsTmp!家庭地址), "", rsTmp!家庭地址)
    txt婚姻状况(X).Text = IIf(IsNull(rsTmp!婚姻状况), "", rsTmp!婚姻状况)
    
    str门诊号 = IIf(IsNull(rsTmp!门诊号), 0, rsTmp!门诊号)
    str住院号 = IIf(IsNull(rsTmp!住院号), 0, rsTmp!住院号)
    'by lesfeng 2010-03-08 性能优化 select A.*
    strSQL = "Select A.出院日期,A.住院号,A.出院病床,A.病人ID,A.主页ID,B.名称 as 科室 From 病案主页 A,部门表 B" & _
        " Where A.主页ID=(Select Max(主页ID) From 病案主页 Where 病人ID=[1])" & _
        " And A.出院科室ID=B.ID And A.病人ID=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If rsTmp.EOF Then
        If glngSys Like "8??" Then
            txt状态(X).Text = "正常"
        Else
            txt状态(X).Text = "门诊"
        End If
        lbl住院号(X).Caption = "门诊号:"
        txt住院号(X).Text = IIf(str门诊号 = "0", "", str门诊号)
        txt科室(X).Text = ""
        txt床位(X).Text = ""
        txt住院次数(X).Text = ""
    Else
        txt状态(X).Text = IIf(IsNull(rsTmp!出院日期), "在院", "出院")
        lbl住院号(X).Caption = "住院号:"
        txt住院号(X).Text = IIf(str住院号 = 0, "", str住院号)
        txt科室(X).Text = rsTmp!科室
        txt床位(X).Text = IIf(IsNull(rsTmp!出院病床), "家庭", rsTmp!出院病床)
        txt住院次数(X).Text = rsTmp!主页ID
    End If
            
    fra(X).Tag = lngID
    
    ShowPatiInfo = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPatiSel
End Sub

Private Sub cmdMerge_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, Curdate As Date
    Dim i As Integer, j As Integer
    Dim str合并原因 As String
    
    If Val(fra(1).Tag) = 0 Then
        If glngSys Like "8??" Then
            MsgBox "没有设置要保留的客户,请先选择一个客户！", vbInformation, gstrSysName
        Else
            MsgBox "没有设置要保留的病人,请先选择一个病人！", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
    If Val(fra(1).Tag) = Val(fra(0).Tag) Then
        If glngSys Like "8??" Then
            MsgBox "你选择了同一个客户,请选择其它客户！", vbInformation, gstrSysName
        Else
            MsgBox "你选择了同一个病人,请选择其它病人！", vbInformation, gstrSysName
        End If
        cmdPati.SetFocus: Exit Sub
    End If
        
    Set rsPatiS = GetPatiInfo(CLng(fra(0).Tag))
    Set rsPatiO = GetPatiInfo(CLng(fra(1).Tag))
    
    'A或B有一个办理了预约入院
    If Not IsNull(rsPatiS!主页ID) And Nvl(rsPatiS!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
    End If
    If Not IsNull(rsPatiO!主页ID) And Nvl(rsPatiO!主页ID, 0) = 0 Then
        MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]办理了预约入院登记，请先取消该登记。", vbInformation, gstrSysName
    End If
    
    'AB都住过院
    If Not IsNull(rsPatiS!主页ID) And Not IsNull(rsPatiO!主页ID) Then
        '1.先住院的在院,不允许(先后住院可以为：出院-出院,出院-在院；不允许：在院-出院,在院-在院)
        '因为除病人合并外,程序不额外处理自动出院或撤消出院
        rsPatiS.MoveLast
        rsPatiO.MoveLast
        If rsPatiS!入院日期 <= rsPatiO!入院日期 Then
            If IsNull(rsPatiS!出院日期) Then
                MsgBox "病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            If IsNull(rsPatiO!出院日期) Then
                MsgBox "病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]最后一次住院先入院,但当前未出院,不能执行合并操作！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '2.时间交叉提示是否继续
        Curdate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
                If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期) Or _
                    IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                    MsgBox "发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), Curdate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), Curdate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                        vbCrLf & "互相交叉，不能进行合并！", _
                        vbInformation, gstrSysName
                        Exit Sub
                End If
                rsPatiO.MoveNext
            Next
            rsPatiS.MoveNext
        Next
    End If
    
    '合并原因
    str合并原因 = InputBox("合并操作后不能撤消,请慎重!" & vbCrLf & vbCrLf & "请输入合并原因:" & vbCrLf & vbCrLf, gstrSysName, "")
    If zlCommFun.ActualLen(str合并原因) > 250 Then
        MsgBox "合并原因不能多于250个字符,请按Ctrl+C复制下面的内容,重新执行时再输入:" & _
            vbCrLf & vbCrLf & str合并原因, vbInformation, gstrSysName
        Exit Sub
    ElseIf Trim(str合并原因) = "" Then
        MsgBox "必须输入合并原因才能进行合并!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    DoEvents
    On Error GoTo errH
    strSQL = "zl_病人信息_MERGE(" & Val(fra(0).Tag) & "," & Val(fra(1).Tag) & ",'" & str合并原因 & "','" & UserInfo.姓名 & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    On Error GoTo 0
    Screen.MousePointer = 0
        
    '合并后应只剩一个病人
    'by lesfeng 2010-03-08 性能优化
    strSQL = "Select 病人ID From 病人信息 Where 病人ID IN([1],[2])"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(fra(0).Tag), Val(fra(1).Tag))
    
    mlng病人ID = rsTmp!病人ID

    If glngSys Like "8??" Then
        MsgBox "客户合并成功,合并后的客户ID为 " & mlng病人ID & "。", vbInformation, gstrSysName
    Else
        MsgBox "病人合并成功,合并后的病人ID为 " & mlng病人ID & "。", vbInformation, gstrSysName
    End If
    Call ClearPatiInfo(1)
    Call ShowPatiInfo(mlng病人ID, 0)
    Unload frmPatiSel
    mblnOK = True
    cmdPati.SetFocus
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatiInfo(lng病人ID As Long) As ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    '主页ID=0时(不是NULL)，表示预约入院
    strSQL = _
        " Select A.病人ID,Decode(B.病人ID,NULL,NULL,Nvl(B.主页ID,0)) as 主页ID," & _
        "            decode(B.姓名,NULL,A.姓名,B.姓名) as 姓名,A.住院号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=[1]" & _
        " Order by Nvl(B.主页ID,0)"
    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
    
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
