VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProperty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "表格属性"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   5760
   Icon            =   "frmProperty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   900
      ScaleHeight     =   420
      ScaleWidth      =   510
      TabIndex        =   87
      Top             =   5805
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox picTMP 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1485
      ScaleHeight     =   420
      ScaleWidth      =   510
      TabIndex        =   86
      Top             =   5760
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Height          =   345
      Left            =   2055
      TabIndex        =   44
      Top             =   5580
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5385
      Left            =   90
      TabIndex        =   52
      Top             =   90
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9499
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "表格(&T)"
      TabPicture(0)   =   "frmProperty.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "picBackPic"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkSingleClickEdit"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkWordEllipsis"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkEditable"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkEnabled"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkHotTrack"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkHighlightSelectedIcons"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkDrawFocusRect"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "chkAutoHeight"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkShowToolTips"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkSingleLine"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkTabTrip"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtBorderWidth"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtGridLineWidth"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCellMargin"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmbFontQuality"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmbHighLightMode"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "picBorderColor"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "picBackColor"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "picHighlightBackColor"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "picGridLineColor"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "picHighlightForeColor"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdHighlightForeColor"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdHighlightBackColor"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdBorderColor"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdGridLineColor"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdBackColor"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "cmdSelPic"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmdDelPic"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Frame4"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtUserTag"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).ControlCount=   47
      TabCaption(1)   =   " 行/列(&U) "
      TabPicture(1)   =   "frmProperty.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkFixColWidth"
      Tab(1).Control(1)=   "cmdNextCol"
      Tab(1).Control(2)=   "cmdPrevCol"
      Tab(1).Control(3)=   "cmdNextRow"
      Tab(1).Control(4)=   "cmdPrevRow"
      Tab(1).Control(5)=   "txtWidth"
      Tab(1).Control(6)=   "Frame6"
      Tab(1).Control(7)=   "txtHeight"
      Tab(1).Control(8)=   "Frame3"
      Tab(1).Control(9)=   "lblCol"
      Tab(1).Control(10)=   "lblRow"
      Tab(1).Control(11)=   "Label22"
      Tab(1).Control(12)=   "Label21"
      Tab(1).Control(13)=   "Label20"
      Tab(1).Control(14)=   "Label19"
      Tab(1).Control(15)=   "Label18"
      Tab(1).Control(16)=   "Label15"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "单元格(&E)"
      TabPicture(2)   =   "frmProperty.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdClearBkClr"
      Tab(2).Control(1)=   "cmbFontSize"
      Tab(2).Control(2)=   "cmbFontName"
      Tab(2).Control(3)=   "cmdFillColor"
      Tab(2).Control(4)=   "cmdForeColor"
      Tab(2).Control(5)=   "chkProtected"
      Tab(2).Control(6)=   "picBG"
      Tab(2).Control(7)=   "txtFormatString"
      Tab(2).Control(8)=   "chkFormatString"
      Tab(2).Control(9)=   "cmdStrikethrough"
      Tab(2).Control(10)=   "cmdUnderLine"
      Tab(2).Control(11)=   "cmdItalic"
      Tab(2).Control(12)=   "Frame8"
      Tab(2).Control(13)=   "Frame7"
      Tab(2).Control(14)=   "cmdBold"
      Tab(2).Control(15)=   "tbrThis"
      Tab(2).Control(16)=   "Label24"
      Tab(2).Control(17)=   "Label23"
      Tab(2).ControlCount=   18
      Begin VB.TextBox txtUserTag 
         Height          =   300
         Left            =   225
         TabIndex        =   23
         Top             =   4860
         Width           =   2490
      End
      Begin VB.Frame Frame4 
         Height          =   30
         Left            =   600
         TabIndex        =   88
         Top             =   4650
         Width           =   2595
      End
      Begin VB.CheckBox chkFixColWidth 
         Caption         =   "固定列宽(&F)"
         Height          =   225
         Left            =   -70980
         TabIndex        =   28
         Top             =   2408
         Width           =   1365
      End
      Begin VB.CommandButton cmdClearBkClr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -74055
         Picture         =   "frmProperty.frx":0060
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "清除背景色"
         Top             =   2250
         Width           =   330
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   300
         Left            =   -72840
         TabIndex        =   32
         Text            =   "cmbFontSize"
         Top             =   675
         Width           =   960
      End
      Begin VB.ComboBox cmbFontName 
         Height          =   300
         Left            =   -74775
         Sorted          =   -1  'True
         TabIndex        =   31
         Text            =   "cmbFontName"
         Top             =   682
         Width           =   1815
      End
      Begin VB.CommandButton cmdNextCol 
         Caption         =   "下一列(&B)"
         Height          =   330
         Left            =   -72015
         TabIndex        =   30
         Top             =   2835
         Width           =   1770
      End
      Begin VB.CommandButton cmdPrevCol 
         Caption         =   "上一列(&A)"
         Height          =   330
         Left            =   -73995
         TabIndex        =   29
         Top             =   2835
         Width           =   1770
      End
      Begin VB.CommandButton cmdNextRow 
         Caption         =   "下一行(&N)"
         Height          =   330
         Left            =   -72015
         TabIndex        =   26
         Top             =   1350
         Width           =   1770
      End
      Begin VB.CommandButton cmdPrevRow 
         Caption         =   "上一行(&P)"
         Height          =   330
         Left            =   -73995
         TabIndex        =   25
         Top             =   1350
         Width           =   1770
      End
      Begin VB.CommandButton cmdFillColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70050
         Picture         =   "frmProperty.frx":68B2
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   690
         Width           =   330
      End
      Begin VB.CommandButton cmdForeColor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70395
         Picture         =   "frmProperty.frx":6A16
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   690
         Width           =   330
      End
      Begin VB.CheckBox chkProtected 
         Caption         =   "保护(&P)"
         Height          =   285
         Left            =   -74685
         TabIndex        =   43
         Top             =   3150
         Width           =   2760
      End
      Begin VB.PictureBox picBG 
         BackColor       =   &H00E0E0E0&
         Height          =   1455
         Left            =   -73650
         ScaleHeight     =   1395
         ScaleWidth      =   3915
         TabIndex        =   82
         Top             =   1080
         Width           =   3975
         Begin VB.Label lblExam 
            BackStyle       =   0  'Transparent
            Caption         =   "示例文本 Example"
            Height          =   225
            Left            =   90
            TabIndex        =   40
            Top             =   90
            Width           =   3465
         End
      End
      Begin VB.TextBox txtFormatString 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   -73335
         TabIndex        =   42
         Top             =   2835
         Width           =   3615
      End
      Begin VB.CheckBox chkFormatString 
         Caption         =   "格式串(&F):"
         Height          =   285
         Left            =   -74685
         TabIndex        =   41
         Top             =   2835
         Width           =   1590
      End
      Begin VB.CommandButton cmdStrikethrough 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -70755
         Picture         =   "frmProperty.frx":6B59
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   690
         Width           =   330
      End
      Begin VB.CommandButton cmdUnderLine 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71100
         Picture         =   "frmProperty.frx":6BAE
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   690
         Width           =   330
      End
      Begin VB.CommandButton cmdItalic 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71460
         Picture         =   "frmProperty.frx":6C23
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   690
         Width           =   330
      End
      Begin VB.Frame Frame8 
         Height          =   120
         Left            =   -74370
         TabIndex        =   80
         Top             =   2565
         Width           =   4830
      End
      Begin VB.Frame Frame7 
         Height          =   120
         Left            =   -74370
         TabIndex        =   78
         Top             =   450
         Width           =   4830
      End
      Begin VB.CommandButton cmdBold 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -71805
         Picture         =   "frmProperty.frx":6C89
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   690
         Width           =   330
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   -73515
         TabIndex        =   27
         Text            =   "1"
         Top             =   2385
         Width           =   960
      End
      Begin VB.Frame Frame6 
         Height          =   120
         Left            =   -74550
         TabIndex        =   74
         Top             =   1845
         Width           =   5010
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   -73515
         TabIndex        =   24
         Text            =   "1"
         Top             =   945
         Width           =   960
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   -74370
         TabIndex        =   70
         Top             =   450
         Width           =   4830
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   -74595
         TabIndex        =   68
         Top             =   450
         Width           =   5055
      End
      Begin VB.CommandButton cmdDelPic 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4965
         Picture         =   "frmProperty.frx":6D16
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3885
         Width           =   330
      End
      Begin VB.CommandButton cmdSelPic 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4650
         Picture         =   "frmProperty.frx":D568
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3885
         Width           =   330
      End
      Begin VB.CommandButton cmdBackColor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   20
         Top             =   3600
         Width           =   285
      End
      Begin VB.CommandButton cmdGridLineColor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   19
         Top             =   3255
         Width           =   285
      End
      Begin VB.CommandButton cmdBorderColor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   18
         Top             =   2925
         Width           =   285
      End
      Begin VB.CommandButton cmdHighlightBackColor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   17
         Top             =   2580
         Width           =   285
      End
      Begin VB.CommandButton cmdHighlightForeColor 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5010
         TabIndex        =   16
         Top             =   2250
         Width           =   285
      End
      Begin VB.PictureBox picHighlightForeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4650
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2250
         Width           =   330
      End
      Begin VB.PictureBox picGridLineColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4650
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3255
         Width           =   330
      End
      Begin VB.PictureBox picHighlightBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4650
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2580
         Width           =   330
      End
      Begin VB.PictureBox picBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4650
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3600
         Width           =   330
      End
      Begin VB.PictureBox picBorderColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4650
         ScaleHeight     =   210
         ScaleWidth      =   300
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2925
         Width           =   330
      End
      Begin VB.ComboBox cmbHighLightMode 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2505
         Width           =   2490
      End
      Begin VB.ComboBox cmbFontQuality 
         Height          =   300
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   4170
         Width           =   2490
      End
      Begin VB.TextBox txtCellMargin 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   2025
         TabIndex        =   14
         Text            =   "1"
         Top             =   3585
         Width           =   690
      End
      Begin VB.TextBox txtGridLineWidth 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   2025
         TabIndex        =   13
         Text            =   "1"
         Top             =   3240
         Width           =   690
      End
      Begin VB.TextBox txtBorderWidth 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   2025
         TabIndex        =   12
         Text            =   "1"
         Top             =   2902
         Width           =   690
      End
      Begin VB.CheckBox chkTabTrip 
         Appearance      =   0  'Flat
         Caption         =   "捕获Tab键(&K)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   3
         Top             =   720
         Value           =   1  'Checked
         Width           =   1710
      End
      Begin VB.CheckBox chkSingleLine 
         Appearance      =   0  'Flat
         Caption         =   "单行文本(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   4
         Top             =   990
         Width           =   1710
      End
      Begin VB.CheckBox chkShowToolTips 
         Appearance      =   0  'Flat
         Caption         =   "显示提示文本(&I)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   8
         Top             =   1905
         Width           =   1665
      End
      Begin VB.CheckBox chkAutoHeight 
         Appearance      =   0  'Flat
         Caption         =   "自动行高(&W)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   1
         Top             =   990
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.CheckBox chkDrawFocusRect 
         Appearance      =   0  'Flat
         Caption         =   "焦点虚框(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3930
         TabIndex        =   10
         Top             =   1905
         Width           =   1410
      End
      Begin VB.CheckBox chkHighlightSelectedIcons 
         Appearance      =   0  'Flat
         Caption         =   "图标高亮(&G)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   9
         Top             =   1905
         Width           =   1710
      End
      Begin VB.CheckBox chkHotTrack 
         Appearance      =   0  'Flat
         Caption         =   "热跟踪(&H)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   2
         Top             =   1260
         Width           =   1620
      End
      Begin VB.CheckBox chkEnabled 
         Appearance      =   0  'Flat
         Caption         =   "可用(&N)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   225
         TabIndex        =   0
         Top             =   720
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.CheckBox chkEditable 
         Appearance      =   0  'Flat
         Caption         =   "可编辑(&D)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3930
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   600
         TabIndex        =   57
         Top             =   1710
         Width           =   4830
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   600
         TabIndex        =   55
         Top             =   525
         Width           =   4830
      End
      Begin VB.CheckBox chkWordEllipsis 
         Appearance      =   0  'Flat
         Caption         =   "显示未完省略(&R)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2025
         TabIndex        =   5
         Top             =   1260
         Width           =   1710
      End
      Begin VB.CheckBox chkSingleClickEdit 
         Appearance      =   0  'Flat
         Caption         =   "单击编辑(&L)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3930
         TabIndex        =   7
         Top             =   990
         Width           =   1410
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   1020
         Left            =   -74775
         TabIndex        =   39
         Top             =   1080
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1799
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         ImageList       =   "imlAlign"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   8
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   9
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "标记"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   135
         TabIndex        =   89
         Top             =   4575
         Width           =   915
      End
      Begin VB.Image picBackPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   990
         Left            =   3675
         Stretch         =   -1  'True
         Top             =   4185
         Width           =   1605
      End
      Begin VB.Label lblCol 
         Caption         =   "第 1 列:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   84
         Top             =   2115
         Width           =   1185
      End
      Begin VB.Label lblRow 
         Caption         =   "第 1 行:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   83
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "其他"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   -74865
         TabIndex        =   81
         Top             =   2565
         Width           =   915
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "字体"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   -74865
         TabIndex        =   79
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label22 
         Caption         =   "厘米"
         Height          =   195
         Left            =   -72480
         TabIndex        =   77
         Top             =   2430
         Width           =   555
      End
      Begin VB.Label Label21 
         Caption         =   "指定宽度(&W):"
         Height          =   195
         Left            =   -74775
         TabIndex        =   76
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "列"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   -74865
         TabIndex        =   75
         Top             =   1845
         Width           =   915
      End
      Begin VB.Label Label19 
         Caption         =   "厘米"
         Height          =   195
         Left            =   -72480
         TabIndex        =   73
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label18 
         Caption         =   "指定高度(&H):"
         Height          =   195
         Left            =   -74775
         TabIndex        =   72
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "外观"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   -74865
         TabIndex        =   71
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "行"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   -74865
         TabIndex        =   69
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "高亮前景色(&2):"
         Height          =   195
         Left            =   3255
         TabIndex        =   67
         Top             =   2280
         Width           =   1680
      End
      Begin VB.Label Label8 
         Caption         =   "高亮显示模式(&1):"
         Height          =   195
         Left            =   225
         TabIndex        =   66
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label14 
         Caption         =   "字体质量(&G):"
         Height          =   195
         Left            =   225
         TabIndex        =   65
         Top             =   3930
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "单元格边距(&8):"
         Height          =   195
         Left            =   225
         TabIndex        =   64
         Top             =   3615
         Width           =   1410
      End
      Begin VB.Label Label11 
         Caption         =   "背景图片(&P):"
         Height          =   195
         Left            =   3255
         TabIndex        =   63
         Top             =   3930
         Width           =   1185
      End
      Begin VB.Label Label10 
         Caption         =   "高亮背景色(&3):"
         Height          =   195
         Left            =   3255
         TabIndex        =   62
         Top             =   2610
         Width           =   1680
      End
      Begin VB.Label Label7 
         Caption         =   "网格宽度(&7):"
         Height          =   195
         Left            =   225
         TabIndex        =   61
         Top             =   3285
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "网格颜色(&6):"
         Height          =   195
         Left            =   3255
         TabIndex        =   60
         Top             =   3285
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "边框宽度(&5):"
         Height          =   195
         Left            =   225
         TabIndex        =   59
         Top             =   2940
         Width           =   1185
      End
      Begin VB.Label Label4 
         Caption         =   "边框颜色(&4):"
         Height          =   195
         Left            =   3255
         TabIndex        =   58
         Top             =   2940
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "外观"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   135
         TabIndex        =   56
         Top             =   1635
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "行为"
         ForeColor       =   &H000052D9&
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   450
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "背景色(&B):"
         Height          =   195
         Left            =   3255
         TabIndex        =   53
         Top             =   3615
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   345
      Left            =   4575
      TabIndex        =   46
      Top             =   5580
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   345
      Left            =   3315
      TabIndex        =   45
      Top             =   5580
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   225
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.JPG|*.JPG|*.BMP|*.BMP|*.GIF|*.GIF|*.*|*.*"
   End
   Begin MSComctlLib.ImageList imlAlign 
      Left            =   45
      Top             =   5355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":13DBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":13EC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":13F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":1405B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":1415D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":1425F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":14360
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":14426
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProperty.frx":14524
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTable As Object, mfrmParent As Object, mRow As Long, mCol As Long, mRowHeights() As Long, mColWidths() As Long, mCell As cCell
Private lAlignment As Long

Private Sub chkAutoHeight_Click()
    txtHeight.Enabled = (chkAutoHeight.Value = vbUnchecked)
End Sub

Private Sub chkEditable_Click()
    chkSingleClickEdit.Enabled = (chkEditable.Value = vbChecked)
End Sub

Private Sub chkFormatString_Click()
    On Error Resume Next
    txtFormatString.Enabled = (chkFormatString.Value = vbChecked)
    If txtFormatString.Enabled Then txtFormatString.SelStart = 0: txtFormatString.SelLength = Len(txtFormatString): txtFormatString.SetFocus
End Sub

Private Sub chkSingleLine_Click()
    chkWordEllipsis.Value = IIf((chkSingleLine.Value <> vbChecked), chkWordEllipsis.Value, vbUnchecked)
    chkWordEllipsis.Enabled = (chkSingleLine.Value = vbChecked)
End Sub

Private Sub cmbFontName_Change()
    On Error Resume Next
    lblExam.FontName = cmbFontName.Text
    Set picBG.Font = lblExam.Font
    lblExam.Height = picBG.TextHeight(lblExam.Caption)
    Select Case lAlignment
    Case 1
        lblExam.Top = 100
    Case 2
        lblExam.Top = -(lblExam.Height - picBG.ScaleHeight) / 2
    Case 3
        lblExam.Top = picBG.ScaleHeight - lblExam.Height
    End Select
End Sub

Private Sub cmbFontName_Click()
    cmbFontName_Change
End Sub

Private Sub cmbFontSize_Change()
    On Error Resume Next
    lblExam.FontSize = Val(cmbFontSize.Text)
    Set picBG.Font = lblExam.Font
    lblExam.Height = picBG.TextHeight(lblExam.Caption)
    Select Case lAlignment
    Case 1
        lblExam.Top = 100
    Case 2
        lblExam.Top = -(lblExam.Height - picBG.ScaleHeight) / 2
    Case 3
        lblExam.Top = picBG.ScaleHeight - lblExam.Height
    End Select
End Sub

Private Sub cmbFontSize_Click()
    cmbFontSize_Change
End Sub

Private Sub cmdApply_Click()
    With mTable
'        Select Case SSTab1.Tab
'        Case 0
        .Enabled = (chkEnabled.Value = vbChecked)
        .AutoHeight = (chkAutoHeight.Value = vbChecked)
        .HotTrack = (chkHotTrack.Value = vbChecked)
        .TabKeyMoveNextCell = (chkTabTrip.Value = vbChecked)
        .SingleLine = (chkSingleLine.Value = vbChecked)
        .WordEllipsis = IIf(.SingleLine, (chkWordEllipsis.Value = vbChecked), False)
        .Editable = (chkEditable.Value = vbChecked)
        .SingleClickEdit = (chkSingleClickEdit.Value = vbChecked)
        .ShowToolTipText = (chkShowToolTips.Value = vbChecked)
        .DrawFocusRect = (chkDrawFocusRect.Value = vbChecked)
        .HighlightSelectedIcons = (chkHighlightSelectedIcons.Value = vbChecked)
        .HighlightMode = cmbHighLightMode.ListIndex
        .BorderWidth = Val(txtBorderWidth)
        .GridLineWidth = Val(txtGridLineWidth)
        .CellMargin = Val(.CellMargin)
        .FontQuality = cmbFontQuality.ListIndex
        .HighlightForeColor = picHighlightForeColor.BackColor
        .HighlightBackColor = picHighlightBackColor.BackColor
        .BorderColor = picBorderColor.BackColor
        .GridLineColor = picGridLineColor.BackColor
        .BackColor = picBackColor.BackColor
        .BackgroundPicture = IIf(picBackPic.Picture = 0, Nothing, picBackPic.Picture)
        .UserTag = Trim(txtUserTag.Text)
            
        .Font.Bold = dlgThis.FontBold
        .Font.Italic = dlgThis.FontItalic
        .Font.Name = dlgThis.FontName
        .Font.Size = dlgThis.FontSize
        .Font.Strikethrough = dlgThis.FontStrikethru
        .Font.Underline = dlgThis.FontUnderline
        .ForeColor = dlgThis.Color
'        .Refresh
'        Case 1
        .RowHeight(mRow) = Me.ScaleY(Val(txtHeight), vbCentimeters, vbTwips)
        .ColWidth(mCol) = IIf(chkFixColWidth.Value = vbChecked, -1, 1) * Me.ScaleX(Val(txtWidth), vbCentimeters, vbTwips)
'        .Refresh
'        Case 2
        '所有选中单元格一并进行设置！
        Dim i As Long
        For i = 1 To mTable.Cells.Count
            If mTable.Cells(i).Selected Then
                mTable.Cells(i).FontName = lblExam.FontName
                mTable.Cells(i).FontSize = lblExam.FontSize
                mTable.Cells(i).FontBold = lblExam.FontBold
                mTable.Cells(i).FontItalic = lblExam.FontItalic
                mTable.Cells(i).FontStrikeout = lblExam.FontStrikethru
                mTable.Cells(i).FontUnderline = lblExam.FontUnderline
                mTable.Cells(i).ForeColor = lblExam.ForeColor
                mTable.Cells(i).BackColor = IIf(picBG.BackColor = vbWindowBackground, -1, picBG.BackColor)
                Select Case lblExam.Alignment
                Case vbLeftJustify
                    mTable.Cells(i).HAlignment = HALignLeft
                Case vbCenter
                    mTable.Cells(i).HAlignment = HALignCentre
                Case vbRightJustify
                    mTable.Cells(i).HAlignment = HALignRight
                End Select
                Select Case lAlignment
                Case 1
                    mTable.Cells(i).VAlignment = VALignTop
                Case 2
                    mTable.Cells(i).VAlignment = VALignVCentre
                Case 3
                    mTable.Cells(i).VAlignment = VALignBottom
                End Select
                mTable.Cells(i).FormatString = IIf(chkFormatString.Value = vbChecked, txtFormatString, "")
                mTable.Cells(i).Protected = (chkProtected.Value = vbChecked)
            End If
        Next
        .Refresh
'        End Select
        .Modified = True
        mfrmParent.RaiseResizeEvent
    End With
End Sub

Private Sub cmdBackColor_Click()
    On Error GoTo LL
    dlgThis.Color = picBackColor.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picBackColor.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdBold_Click()
    lblExam.Font.Bold = Not lblExam.Font.Bold
End Sub

Private Sub cmdBorderColor_Click()
    On Error GoTo LL
    dlgThis.Color = picBorderColor.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picBorderColor.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Function ShowMe(ByRef frmParent As Object, ByRef oTable As Object, Optional ByVal lStartTab As Long) As Boolean
    Set mfrmParent = frmParent
    Set mTable = oTable
    With cmbHighLightMode
        .AddItem "0-无"
        .AddItem "1-纯色矩形边框"
        .AddItem "2-半透明矩形边框"
        .AddItem "3-纯色矩形填充"
        .AddItem "4-半透明矩形填充"
    End With
    With cmbFontQuality
        .AddItem "0-默认字体质量"
        .AddItem "1-较低质量"
        .AddItem "2-质量优先"
        .AddItem "3-非抗锯齿模式"
        .AddItem "4-抗锯齿模式"
        .AddItem "5-清晰模式"
    End With
    With mTable
        '1
        chkEnabled.Value = IIf(.Enabled, vbChecked, vbUnchecked)
        chkAutoHeight.Value = IIf(.AutoHeight, vbChecked, vbUnchecked)
        chkHotTrack.Value = IIf(.HotTrack, vbChecked, vbUnchecked)
        chkTabTrip.Value = IIf(.TabKeyMoveNextCell, vbChecked, vbUnchecked)
        chkSingleLine.Value = IIf(.SingleLine, vbChecked, vbUnchecked)
        chkWordEllipsis.Value = IIf(.WordEllipsis, vbChecked, vbUnchecked)
        chkWordEllipsis.Enabled = .SingleLine
        chkEditable.Value = IIf(.Editable, vbChecked, vbUnchecked)
        chkSingleClickEdit.Value = IIf(.SingleClickEdit, vbChecked, vbUnchecked)
        chkSingleClickEdit.Enabled = .Editable
        chkShowToolTips.Value = IIf(.ShowToolTipText, vbChecked, vbUnchecked)
        chkDrawFocusRect.Value = IIf(.DrawFocusRect, vbChecked, vbUnchecked)
        chkHighlightSelectedIcons.Value = IIf(.HighlightSelectedIcons, vbChecked, vbUnchecked)
        cmbHighLightMode.ListIndex = .HighlightMode
        txtBorderWidth.Text = .BorderWidth
        txtGridLineWidth.Text = .GridLineWidth
        txtCellMargin.Text = .CellMargin
        cmbFontQuality.ListIndex = .FontQuality
        
'        txtFont.Text = .Font.Name & "," & .Font.Size & "磅" & IIf(.Font.Bold, ",粗体", "") & IIf(.Font.Italic, ",斜体", "")
'        dlgThis.FontBold = .Font.Bold
'        dlgThis.FontItalic = .Font.Italic
'        dlgThis.FontName = .Font.Name
'        dlgThis.FontSize = .Font.Size
'        dlgThis.FontStrikethru = .Font.Strikethrough
'        dlgThis.FontUnderline = .Font.Underline
       
        picHighlightForeColor.BackColor = .HighlightForeColor
        picHighlightBackColor.BackColor = .HighlightBackColor
        picBorderColor.BackColor = .BorderColor
        picGridLineColor.BackColor = .GridLineColor
        picBackColor.BackColor = .BackColor
        Set picBackPic.Picture = .BackgroundPicture
        txtUserTag.Text = .UserTag
        
        '2
        Dim i As Long
        ReDim mRowHeights(1 To .RowCount) As Long
        ReDim mColWidths(1 To .ColCount) As Long
        For i = 1 To .RowCount
            mRowHeights(i) = .RowHeight(i)
        Next
        For i = 1 To .ColCount
            mColWidths(i) = .ColWidth(i)
        Next
        If .SelectedCellKey > 0 Then
            mRow = .Cells(.SelectedCellKey).Row
            mCol = .Cells(.SelectedCellKey).Col
        Else
            mRow = 1
            mCol = 1
            .Cell(1, 1).Selected = True
            cmdPrevRow.Enabled = False
            cmdPrevCol.Enabled = False
        End If
        If .RowCount <= 1 Then cmdNextRow.Enabled = False
        If .ColCount <= 1 Then cmdNextCol.Enabled = False
        lblRow.Caption = "第 " & mRow & " 行"
        lblCol.Caption = "第 " & mCol & " 列"
        chkFixColWidth.Value = IIf(.ColWidth(mCol) < 0, vbChecked, vbUnchecked)
        txtWidth.Text = Abs(Me.ScaleX(.ColWidth(mCol), vbTwips, vbCentimeters))
        txtHeight.Text = Me.ScaleY(.RowHeight(mRow), vbTwips, vbCentimeters)
        txtHeight.Enabled = (chkAutoHeight.Value = vbUnchecked)
        txtWidth.Text = Format(txtWidth.Text, "0.00")
        txtHeight.Text = Format(txtHeight.Text, "0.00")
        lblCol.Width = picBG.ScaleWidth
        Set mCell = .Cell(mRow, mCol).Clone(True)
        
        '3
        Dim hDC As Long
        cmbFontName.Clear
        EnumFontFamilies Me.hDC, vbNullString, AddressOf EnumFontFamProc, ByVal 0&
        cmbFontName.Text = mCell.FontName
        
        cmbFontSize.Clear
        cmbFontSize.AddItem 5
        cmbFontSize.AddItem 5.5
        cmbFontSize.AddItem 6.5
        cmbFontSize.AddItem 7.5
        cmbFontSize.AddItem 8
        cmbFontSize.AddItem 9
        cmbFontSize.AddItem 10
        cmbFontSize.AddItem 10.5
        cmbFontSize.AddItem 11
        cmbFontSize.AddItem 12
        cmbFontSize.AddItem 14
        cmbFontSize.AddItem 16
        cmbFontSize.AddItem 18
        cmbFontSize.AddItem 20
        cmbFontSize.AddItem 22
        cmbFontSize.AddItem 24
        cmbFontSize.AddItem 26
        cmbFontSize.AddItem 28
        cmbFontSize.AddItem 36
        cmbFontSize.AddItem 48
        cmbFontSize.AddItem 72
        cmbFontSize.Text = mCell.FontSize
        
        lblExam.FontName = mCell.FontName
        lblExam.FontSize = mCell.FontSize
        lblExam.FontBold = mCell.FontBold
        lblExam.FontItalic = mCell.FontItalic
        lblExam.FontStrikethru = mCell.FontStrikeout
        lblExam.FontUnderline = mCell.FontUnderline
        lblExam.ForeColor = IIf(mCell.ForeColor = -1, vbBlack, mCell.ForeColor)
        picBG.BackColor = IIf(mCell.BackColor = -1, vbWindowBackground, mCell.BackColor)
        Select Case mCell.HAlignment
        Case HALignLeft
            lblExam.Alignment = vbLeftJustify
        Case HALignCentre
            lblExam.Alignment = vbCenter
        Case HALignRight
            lblExam.Alignment = vbRightJustify
        End Select
        Select Case mCell.VAlignment
        Case VALignTop
            lblExam.Top = 100
            lAlignment = 1
        Case VALignVCentre
            lblExam.Top = -(lblExam.Height - picBG.ScaleHeight) / 2
            lAlignment = 2
        Case VALignBottom
            lblExam.Top = picBG.ScaleHeight - lblExam.Height
            lAlignment = 3
        End Select
        Set picBG.Font = lblExam.Font
        lblExam.Height = picBG.TextHeight(lblExam.Caption)
        chkFormatString.Value = IIf(mCell.FormatString <> "", vbChecked, vbUnchecked)
        txtFormatString = mCell.FormatString
        txtFormatString.Enabled = (chkFormatString.Value = vbChecked)
        chkProtected.Value = IIf(mCell.Protected, vbChecked, vbUnchecked)
        
    End With
    SSTab1.Tab = IIf(lStartTab > 3, 3, IIf(lStartTab < 1, 1, lStartTab)) - 1
    Me.Show vbModal, frmParent
End Function

Private Sub cmdDelPic_Click()
    Set picBackPic.Picture = LoadPicture("")
End Sub

Private Sub cmdFillColor_Click()
    On Error GoTo LL
    dlgThis.Color = picBG.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picBG.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdFont_Click()
    dlgThis.CancelError = False
    dlgThis.flags = cdlCFBoth Or cdlCFEffects
    dlgThis.ShowFont
End Sub

Private Sub cmdForeColor_Click()
    On Error GoTo LL
    dlgThis.Color = lblExam.ForeColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        lblExam.ForeColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdGridLineColor_Click()
    On Error GoTo LL
    dlgThis.Color = picGridLineColor.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picGridLineColor.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdHighlightBackColor_Click()
    On Error GoTo LL
    dlgThis.Color = picHighlightForeColor.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picHighlightBackColor.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdHighlightForeColor_Click()
    On Error GoTo LL
    dlgThis.Color = picHighlightForeColor.BackColor
    dlgThis.CancelError = True
    dlgThis.ShowColor
    If dlgThis.Color <> -1 Then
        picHighlightForeColor.BackColor = dlgThis.Color
    End If
LL:
End Sub

Private Sub cmdItalic_Click()
    lblExam.Font.Italic = Not lblExam.Font.Italic
End Sub

Private Sub cmdNextCol_Click()
    mCol = mCol + 1
    If mCol > mTable.ColCount Then mCol = mTable.ColCount
    cmdPrevCol.Enabled = (mCol > 1)
    cmdNextCol.Enabled = (mCol < mTable.ColCount)
    lblRow.Caption = "第 " & mRow & " 行"
    lblCol.Caption = "第 " & mCol & " 列"
    chkFixColWidth.Value = IIf(mTable.ColWidth(mCol) < 0, vbChecked, vbUnchecked)
    txtWidth.Text = Abs(Me.ScaleX(mTable.ColWidth(mCol), vbTwips, vbCentimeters))
    txtWidth.Text = Format(txtWidth.Text, "0.00")
End Sub

Private Sub cmdNextRow_Click()
    mRow = mRow + 1
    If mRow > mTable.RowCount Then mRow = mTable.RowCount
    cmdPrevRow.Enabled = (mRow > 1)
    cmdNextRow.Enabled = (mRow < mTable.RowCount)
    lblRow.Caption = "第 " & mRow & " 行"
    lblCol.Caption = "第 " & mCol & " 列"
    txtHeight.Text = Me.ScaleY(mTable.RowHeight(mRow), vbTwips, vbCentimeters)
    txtHeight.Text = Format(txtHeight.Text, "0.00")
End Sub

Private Sub cmdOK_Click()
    Call cmdApply_Click
    Unload Me
End Sub

Private Sub cmdPrevCol_Click()
    mCol = mCol - 1
    If mCol < 1 Then mCol = 1
    cmdPrevCol.Enabled = (mCol > 1)
    cmdNextCol.Enabled = (mCol < mTable.ColCount)
    lblRow.Caption = "第 " & mRow & " 行"
    lblCol.Caption = "第 " & mCol & " 列"
    chkFixColWidth.Value = IIf(mTable.ColWidth(mCol) < 0, vbChecked, vbUnchecked)
    txtWidth.Text = Abs(Me.ScaleX(mTable.ColWidth(mCol), vbTwips, vbCentimeters))
    txtWidth.Text = Format(txtWidth.Text, "0.00")
End Sub

Private Sub cmdPrevRow_Click()
    mRow = mRow - 1
    If mRow < 1 Then mRow = 1
    cmdPrevRow.Enabled = (mRow > 1)
    cmdNextRow.Enabled = (mRow < mTable.RowCount)
    lblRow.Caption = "第 " & mRow & " 行"
    lblCol.Caption = "第 " & mCol & " 列"
    txtHeight.Text = Me.ScaleY(mTable.RowHeight(mRow), vbTwips, vbCentimeters)
    txtHeight.Text = Format(txtHeight.Text, "0.00")
End Sub

Private Sub cmdSelPic_Click()
    On Error GoTo LL
    If dlgThis.FileName = "" Then dlgThis.FileName = App.Path
    dlgThis.CancelError = True
    dlgThis.ShowOpen
    If dlgThis.FileName <> "" And dlgThis.FileName <> App.Path Then
        Set picBackPic.Picture = LoadPicture(dlgThis.FileName)
    End If
LL:
End Sub

Private Sub cmdStrikethrough_Click()
    lblExam.Font.Strikethrough = Not lblExam.Font.Strikethrough
End Sub

Private Sub cmdUnderLine_Click()
    lblExam.Font.Underline = Not lblExam.Font.Underline
End Sub

Private Sub cmdClearBkClr_Click()
    picBG.BackColor = vbWindowBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTable = Nothing
    Set mCell = Nothing
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1, 2, 3
        lblExam.Top = 100
        lAlignment = 1
    Case 4, 5, 6
        lblExam.Top = -(lblExam.Height - picBG.ScaleHeight) / 2
        lAlignment = 2
    Case 7, 8, 9
        lblExam.Top = picBG.ScaleHeight - lblExam.Height
        lAlignment = 3
    End Select
    Select Case Button.Index
    Case 1, 4, 7
        lblExam.Alignment = vbLeftJustify
    Case 2, 5, 8
        lblExam.Alignment = vbCenter
    Case 3, 6, 9
        lblExam.Alignment = vbRightJustify
    End Select
End Sub

Private Sub txtBorderWidth_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtCellMargin_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtGridLineWidth_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
