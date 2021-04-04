VERSION 5.00
Begin VB.Form frmMergePatient 
   Caption         =   "病人合并"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   Icon            =   "frmMergePatient.frx":0000
   LinkTopic       =   "Form1"
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
      TabIndex        =   0
      Top             =   630
      Width           =   1100
   End
   Begin VB.Frame fra 
      Caption         =   "要保留的病人信息"
      Height          =   4725
      Index           =   1
      Left            =   3330
      TabIndex        =   5
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt状态 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   525
         Width           =   2000
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   780
         Width           =   2000
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2000
      End
      Begin VB.TextBox txt出生日期 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2000
      End
      Begin VB.TextBox txt国籍 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2000
      End
      Begin VB.TextBox txt民族 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txt学历 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt婚姻状况 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2000
      End
      Begin VB.TextBox txt职业 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2000
      End
      Begin VB.TextBox txt身份 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txt身份证号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2000
      End
      Begin VB.TextBox txt出生地点 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2000
      End
      Begin VB.TextBox txt家庭地址 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2000
      End
      Begin VB.TextBox txt病案号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2000
      End
      Begin VB.TextBox txt住院次数 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病案号:"
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   69
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lbl状态 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   67
         Top             =   270
         Width           =   450
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   1
         Left            =   675
         TabIndex        =   65
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
         Left            =   675
         TabIndex        =   64
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
         Left            =   315
         TabIndex        =   63
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
         Left            =   675
         TabIndex        =   62
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
         Left            =   675
         TabIndex        =   61
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
         Left            =   675
         TabIndex        =   60
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
         Left            =   315
         TabIndex        =   59
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
         Left            =   675
         TabIndex        =   58
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
         Left            =   675
         TabIndex        =   57
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
         Left            =   315
         TabIndex        =   56
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
         Left            =   315
         TabIndex        =   55
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
         Left            =   315
         TabIndex        =   54
         Top             =   3600
         Width           =   810
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大主页ID:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   53
         Top             =   4140
         Width           =   990
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   495
         TabIndex        =   52
         Top             =   525
         Width           =   630
      End
   End
   Begin VB.Frame fra 
      Caption         =   "被合并的病人信息"
      Height          =   4725
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   3135
      Begin VB.TextBox txt住院次数 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2000
      End
      Begin VB.TextBox txt病案号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3870
         Width           =   2000
      End
      Begin VB.TextBox txt家庭地址 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2000
      End
      Begin VB.TextBox txt出生地点 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3345
         Width           =   2000
      End
      Begin VB.TextBox txt身份证号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3090
         Width           =   2000
      End
      Begin VB.TextBox txt身份 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txt职业 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2580
         Width           =   2000
      End
      Begin VB.TextBox txt婚姻状况 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2325
         Width           =   2000
      End
      Begin VB.TextBox txt学历 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt民族 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1800
         Width           =   2000
      End
      Begin VB.TextBox txt国籍 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1545
         Width           =   2000
      End
      Begin VB.TextBox txt出生日期 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2000
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2000
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   780
         Width           =   2000
      End
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   525
         Width           =   2000
      End
      Begin VB.TextBox txt状态 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   1100
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   270
         Width           =   2000
      End
      Begin VB.Label lbl状态 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "状态:"
         ForeColor       =   &H000000C0&
         Height          =   180
         Index           =   0
         Left            =   660
         TabIndex        =   68
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病案号:"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   66
         Top             =   3840
         Width           =   630
      End
      Begin VB.Label lbl住院号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   19
         Top             =   525
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大主页ID:"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   4140
         Width           =   990
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址:"
         ForeColor       =   &H00333333&
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   17
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
         Left            =   300
         TabIndex        =   16
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
         Left            =   300
         TabIndex        =   15
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
         Left            =   660
         TabIndex        =   14
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
         Left            =   660
         TabIndex        =   13
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
         Left            =   300
         TabIndex        =   12
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
         Left            =   660
         TabIndex        =   11
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
         Left            =   660
         TabIndex        =   10
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
         Left            =   660
         TabIndex        =   9
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
         Left            =   300
         TabIndex        =   8
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
         Left            =   660
         TabIndex        =   7
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
         Left            =   660
         TabIndex        =   6
         Top             =   780
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdMerge 
      Caption         =   "合并(&M)"
      Height          =   350
      Left            =   6675
      TabIndex        =   1
      Top             =   1155
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   6675
      TabIndex        =   2
      Top             =   3570
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6675
      TabIndex        =   3
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
Private mlng病人ID As Long '入：初始要合并的病人ID
Private mlng保留病人 As Long '入：保留病人的病人ID
Private mstrPrivs As String
Private mstr住院号 As String '出:保留病人住院号

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdTurn_Click()
    Dim lngTmp As Long
    
    lngTmp = fra(0).Tag
    Call ShowPatiInfo(CLng(fra(1).Tag), "", 0)
    Call ShowPatiInfo(lngTmp, "", 1)
End Sub



Private Sub ClearPatiInfo(x As Integer)
'功能：清除一个病人信息
'参数：x=控件索引,0=源病人,1=目标病人
    txt姓名(x).Text = ""
    txt性别(x).Text = ""
    txt出生日期(x).Text = ""
    txt国籍(x).Text = ""
    txt民族(x).Text = ""
    txt学历(x).Text = ""
    txt身份(x).Text = ""
    txt职业(x).Text = ""
    txt身份证号(x).Text = ""
    txt出生地点(x).Text = ""
    txt家庭地址(x).Text = ""
    txt婚姻状况(x).Text = ""
    txt状态(x).Text = ""
    lbl住院号(x).Caption = "住院号:"
    txt住院号(x).Text = ""
    txt病案号(x).Text = ""
    txt住院次数(x).Text = ""
    fra(x).Tag = ""
End Sub

Private Sub cmdMerge_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim rsPatiS As ADODB.Recordset
    Dim rsPatiO As ADODB.Recordset
    Dim strSQL As String, curDate As Date
    Dim i As Integer, j As Integer
    Dim str合并原因 As String
        
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
        curDate = zlDatabase.Currentdate
        rsPatiS.MoveFirst
        For i = 1 To rsPatiS.RecordCount
            rsPatiO.MoveFirst
            For j = 1 To rsPatiO.RecordCount
'                If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期) Or _
'                    IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
'                    If MsgBox("发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
'                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期), "yyyy-MM-dd") & _
'                        vbCrLf & "互相交叉，应该不是同一个病人，确实要合并吗？", _
'                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
'                End If
                 If Not (rsPatiO!入院日期 >= IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期) Or _
                    IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期) <= rsPatiS!入院日期) Then
                    MsgBox "发现病人:" & rsPatiS!姓名 & "[" & rsPatiS!住院号 & "]第 " & rsPatiS!主页ID & " 次住院的期间" & Format(rsPatiS!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiS!出院日期), curDate, rsPatiS!出院日期), "yyyy-MM-dd") & vbCrLf & _
                        "与病人:" & rsPatiO!姓名 & "[" & rsPatiO!住院号 & "]的第 " & rsPatiO!主页ID & " 次住院的期间" & Format(rsPatiO!入院日期, "yyyy-MM-dd") & "至" & Format(IIf(IsNull(rsPatiO!出院日期), curDate, rsPatiO!出院日期), "yyyy-MM-dd") & _
                        vbCrLf & "互相交叉，不能合并？", vbInformation, gstrSysName
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
    strSQL = "Select 病人ID From 病人信息 Where 病人ID IN(" & Val(fra(0).Tag) & "," & Val(fra(1).Tag) & ")"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    mlng病人ID = rsTmp!病人ID

    If gclsPros.SysNo Like "8??" Then
        MsgBox "客户合并成功,合并后的客户ID为 " & mlng病人ID & "。", vbInformation, gstrSysName
    Else
        MsgBox "病人合并成功,合并后的病人ID为 " & mlng病人ID & "。", vbInformation, gstrSysName
    End If
    
    
'    Call ClearPatiInfo(1)
    '56792:刘鹏飞,2012-12-12,合并之后病案号应该传""而不是0
    Call ShowPatiInfo(mlng病人ID, "", 0)
    mstr住院号 = txt住院号(0).Text
    Unload Me
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
        " A.姓名,A.住院号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B" & _
        " Where A.病人ID=B.病人ID(+) And A.病人ID=" & lng病人ID & _
        " Order by Nvl(B.主页ID,0)"
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetPatiInfo = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ShowPatiInfo(lngID As Long, str病案号 As String, x As Integer) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo errHandle
    
    '56792:刘鹏飞,2012-12-12
    If str病案号 = "" Then
        '主要检查要合并的病人是否已经编目
        strSQL = "Select A.病人id, A.住院号, A.姓名, A.性别, A.出生日期, A.国籍, A.民族, A.学历, A.婚姻状况, A.职业, A.身份, A.身份证号, " & _
                 "       A.出生地点 , A.家庭地址, A.出院时间, B.病案号, C.最大主页ID 主页ID " & _
                 "From 病人信息 A, 住院病案记录 B, (Select Max(B.主页id) 主页id,A.病人ID,Max(A.主页ID) 最大主页ID From 病案主页 A,住院病案记录 B Where A.病人ID=B.病人ID(+) ANd A.病人id =[1] Group by a.病人ID) C " & _
                 "Where A.病人ID = C.病人ID And C.病人ID=B.病人ID(+) And C.主页id=B.主页ID(+) And A.病人ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
        
    Else
        '提取已经编目的保留病人
        strSQL = "Select A.病人id, A.住院号, A.姓名, A.性别, A.出生日期, A.国籍, A.民族, A.学历, A.婚姻状况, A.职业, A.身份, A.身份证号, " & _
                 "       A.出生地点 , A.家庭地址, A.出院时间,B.病案号, C.最大主页ID 主页ID " & _
                 "From 病人信息 A, 住院病案记录 B, (Select A.病人id,Max(B.主页id) 主页id,Max(A.主页ID) 最大主页ID From 病案主页 A,住院病案记录 B Where A.病人ID=B.病人id  and B.病案号= [1] group by A.病人ID) C " & _
                 "Where A.病人ID = B.病人ID and a.病人id=C.病人id and b.主页ID=c.主页ID  "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str病案号)
    End If
    If Not rsTemp.EOF Then
        
        txt姓名(x).Text = IIf(IsNull(rsTemp!姓名), "", rsTemp!姓名)
        txt性别(x).Text = IIf(IsNull(rsTemp!性别), "", rsTemp!性别)
        txt出生日期(x).Text = Format(IIf(IsNull(rsTemp!出生日期), "", rsTemp!出生日期), "yyyy年MM月dd日")
        txt国籍(x).Text = IIf(IsNull(rsTemp!国籍), "", rsTemp!国籍)
        txt民族(x).Text = IIf(IsNull(rsTemp!民族), "", rsTemp!民族)
        txt学历(x).Text = IIf(IsNull(rsTemp!学历), "", rsTemp!学历)
        txt身份(x).Text = IIf(IsNull(rsTemp!身份), "", rsTemp!身份)
        txt职业(x).Text = IIf(IsNull(rsTemp!职业), "", rsTemp!职业)
        txt身份证号(x).Text = IIf(IsNull(rsTemp!身份证号), "", rsTemp!身份证号)
        txt出生地点(x).Text = IIf(IsNull(rsTemp!出生地点), "", rsTemp!出生地点)
        txt家庭地址(x).Text = IIf(IsNull(rsTemp!家庭地址), "", rsTemp!家庭地址)
        txt婚姻状况(x).Text = IIf(IsNull(rsTemp!婚姻状况), "", rsTemp!婚姻状况)
        txt住院号(x).Text = IIf(IsNull(rsTemp!住院号), "", rsTemp!住院号)
        txt状态(x).Text = IIf(IsNull(rsTemp!出院时间), "在院", "出院")
        txt病案号(x).Text = IIf(IsNull(rsTemp!病案号), "", rsTemp!病案号)
        txt住院次数(x).Text = IIf(IsNull(rsTemp!主页ID), "", rsTemp!主页ID)
        fra(x).Tag = rsTemp!病人ID
        ShowPatiInfo = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    ShowPatiInfo = False
End Function

Public Function MergePatient(lngID As Long, str病案号 As String, frmMain As Object) As String
    '返回住院号
    mstr住院号 = ""
    If ShowPatiInfo(lngID, "", 0) = False Then MergePatient = False: Exit Function
    If ShowPatiInfo(0, str病案号, 1) = False Then MergePatient = False: Exit Function
    
    '56792:刘鹏飞,2012-12-12
    '如果两个病人都已经编目则提示操作员不能合并
    cmdTurn.Enabled = True
    If Trim(txt病案号(0).Text) <> "" And Trim(txt病案号(1).Text) <> "" Then
        ShowMsgbox "合并和保留病人的病案均已经编目，不允许进行病人合并操作！"
        MergePatient = txt住院号(0).Text
        Exit Function
    ElseIf Trim(txt病案号(0).Text) = "" And Trim(txt病案号(1).Text) <> "" Then
        cmdTurn.Enabled = False
    Else
        ShowMsgbox "没有提取到保留病人的病案号，请检查！"
        MergePatient = txt住院号(0).Text
        Exit Function
    End If
    
    frmMergePatient.Show 1, frmMain
    MergePatient = mstr住院号
End Function

