VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#4.3#0"; "ZlPatiAddress.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCertifyRegist 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "实名信息登记"
   ClientHeight    =   12750
   ClientLeft      =   225
   ClientTop       =   -3510
   ClientWidth     =   14700
   Icon            =   "frmCertifyRegist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12750
   ScaleWidth      =   14700
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar vsbMain 
      Height          =   7335
      LargeChange     =   100
      Left            =   0
      Max             =   1000
      SmallChange     =   10
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1800
      Width           =   255
   End
   Begin VB.HScrollBar hsbMain 
      Height          =   255
      LargeChange     =   25
      Left            =   6840
      Max             =   100
      TabIndex        =   78
      Top             =   240
      Width           =   7935
   End
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   76
      Top             =   12390
      Width           =   14700
      _ExtentX        =   25929
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCertifyRegist.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21405
            Key             =   "Info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人状态"
            TextSave        =   "病人状态"
            Key             =   "病人状态"
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
   Begin MSComDlg.CommonDialog cmdialog 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picBig 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      ForeColor       =   &H80000008&
      Height          =   13935
      Left            =   360
      ScaleHeight     =   13905
      ScaleWidth      =   15945
      TabIndex        =   33
      Top             =   360
      Width           =   15975
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   13935
         Left            =   0
         ScaleHeight     =   13935
         ScaleWidth      =   14535
         TabIndex        =   35
         Top             =   600
         Width           =   14535
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6255
            Left            =   240
            ScaleHeight     =   6255
            ScaleWidth      =   14295
            TabIndex        =   42
            Top             =   0
            Width           =   14295
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   0
               Left            =   1485
               MaxLength       =   100
               TabIndex        =   3
               Top             =   1792
               Width           =   2175
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   10485
               TabIndex        =   54
               Top             =   1807
               Width           =   1640
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   1
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   6
                  Top             =   -30
                  Width           =   1620
               End
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   4845
               TabIndex        =   53
               Top             =   1807
               Width           =   1480
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   0
                  Left            =   -30
                  TabIndex        =   4
                  Text            =   "cboInfo"
                  Top             =   -30
                  Width           =   1455
               End
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   3
               Left            =   7320
               MaxLength       =   18
               TabIndex        =   20
               Top             =   5062
               Width           =   2055
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   2
               Left            =   1485
               MaxLength       =   100
               TabIndex        =   14
               Top             =   4575
               Width           =   2175
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   4
               Left            =   4845
               TabIndex        =   52
               Top             =   4590
               Width           =   1480
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   4
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   15
                  Top             =   -30
                  Width           =   1455
               End
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   6
               Left            =   4845
               TabIndex        =   51
               Top             =   5070
               Width           =   1480
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   6
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   19
                  Top             =   -30
                  Width           =   1455
               End
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   7
               Left            =   1485
               TabIndex        =   50
               Top             =   5070
               Width           =   2175
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   7
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   -30
                  Width           =   2165
               End
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   8
               Left            =   10485
               TabIndex        =   49
               Top             =   5070
               Width           =   1640
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   8
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   21
                  Top             =   -30
                  Width           =   1620
               End
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   5
               Left            =   10485
               TabIndex        =   48
               Top             =   4575
               Width           =   1640
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   5
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   17
                  Top             =   -30
                  Width           =   1620
               End
            End
            Begin VB.CommandButton cmdAdress 
               Appearance      =   0  'Flat
               Caption         =   "…"
               Height          =   255
               Index           =   2
               Left            =   5980
               TabIndex        =   0
               Top             =   5520
               Width           =   255
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   1
               Left            =   7365
               MaxLength       =   18
               TabIndex        =   9
               Top             =   2287
               Width           =   2055
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   3
               Left            =   1485
               TabIndex        =   47
               Top             =   2302
               Width           =   2175
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   3
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   7
                  Top             =   -30
                  Width           =   2165
               End
            End
            Begin VB.CommandButton cmdAdress 
               Appearance      =   0  'Flat
               Caption         =   "…"
               Height          =   255
               Index           =   1
               Left            =   11820
               TabIndex        =   1
               Top             =   2745
               Width           =   255
            End
            Begin VB.CommandButton cmdAdress 
               Appearance      =   0  'Flat
               Caption         =   "…"
               Height          =   255
               Index           =   0
               Left            =   5980
               TabIndex        =   2
               Top             =   2760
               Width           =   255
            End
            Begin VB.Frame frmInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   4845
               TabIndex        =   46
               Top             =   2302
               Width           =   1480
               Begin VB.ComboBox cboInfo 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Index           =   2
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   -30
                  Width           =   1455
               End
            End
            Begin VB.CommandButton cmdInfoDate 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   9135
               Picture         =   "frmCertifyRegist.frx":70E4
               Style           =   1  'Graphical
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1807
               Width           =   270
            End
            Begin VB.CommandButton cmdInfoDate 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   9135
               Picture         =   "frmCertifyRegist.frx":71DA
               Style           =   1  'Graphical
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   4597
               Width           =   270
            End
            Begin VB.PictureBox picPicture 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1575
               Left            =   12240
               ScaleHeight     =   1575
               ScaleWidth      =   1575
               TabIndex        =   43
               Top             =   1800
               Width           =   1575
               Begin VB.Image imgPatient 
                  Appearance      =   0  'Flat
                  Height          =   1515
                  Left            =   0
                  Picture         =   "frmCertifyRegist.frx":72D0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   1575
               End
            End
            Begin MSMask.MaskEdBox txtDateInfo 
               Height          =   255
               Index           =   0
               Left            =   7320
               TabIndex        =   5
               Tag             =   "####-##-## ##:##"
               Top             =   1800
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtInfoDate 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   235
               Index           =   0
               Left            =   7320
               TabIndex        =   31
               Top             =   1810
               Width           =   1740
            End
            Begin ZlPatiAddress.PatiAddress patiAdressInfo 
               Height          =   270
               Index           =   0
               Left            =   1440
               TabIndex        =   10
               Tag             =   "出生地址"
               Top             =   2760
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   15659001
               Items           =   3
               Style           =   1
               MaxLength       =   100
            End
            Begin VB.TextBox txtAdressInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   0
               Left            =   1425
               MaxLength       =   100
               TabIndex        =   11
               Top             =   2760
               Width           =   4785
            End
            Begin ZlPatiAddress.PatiAddress patiAdressInfo 
               Height          =   270
               Index           =   1
               Left            =   7320
               TabIndex        =   12
               Tag             =   "现住址"
               Top             =   2760
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               Style           =   1
               MaxLength       =   100
            End
            Begin VB.TextBox txtAdressInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   1
               Left            =   7305
               MaxLength       =   100
               TabIndex        =   13
               Top             =   2760
               Width           =   4785
            End
            Begin ZlPatiAddress.PatiAddress patiAdressInfo 
               Height          =   270
               Index           =   2
               Left            =   1485
               TabIndex        =   22
               Tag             =   "联系人地址"
               Top             =   5535
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483643
               Style           =   1
               MaxLength       =   100
            End
            Begin VB.TextBox txtAdressInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   2
               Left            =   1425
               MaxLength       =   100
               TabIndex        =   23
               Top             =   5535
               Width           =   4785
            End
            Begin MSMask.MaskEdBox txtDateInfo 
               Height          =   255
               Index           =   1
               Left            =   7320
               TabIndex        =   16
               Tag             =   "####-##-## ##:##"
               Top             =   4590
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   0
               Appearance      =   0
               AutoTab         =   -1  'True
               MaxLength       =   16
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   10.5
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "####-##-## ##:##"
               PromptChar      =   "_"
            End
            Begin VB.TextBox txtInfoDate 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   235
               Index           =   1
               Left            =   7305
               TabIndex        =   34
               Top             =   4600
               Width           =   1740
            End
            Begin VB.Label lblDelete 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "清除"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   13440
               TabIndex        =   75
               Top             =   3480
               Width           =   495
            End
            Begin VB.Label lblAdd 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "采集"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   12840
               TabIndex        =   74
               Top             =   3480
               Width           =   495
            End
            Begin VB.Label lblFile 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "文件"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   12240
               TabIndex        =   73
               Top             =   3480
               Width           =   495
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "国    籍"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   9600
               TabIndex        =   72
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "姓    名"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   600
               TabIndex        =   71
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "性    别"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   3960
               TabIndex        =   70
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "出生日期"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   6480
               TabIndex        =   69
               Top             =   1800
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "身份证号"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   13
               Left            =   6480
               TabIndex        =   68
               Top             =   5070
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "姓    名"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   9
               Left            =   600
               TabIndex        =   67
               Top             =   4590
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "性    别"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   10
               Left            =   3960
               TabIndex        =   66
               Top             =   4590
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "出生日期"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   11
               Left            =   6480
               TabIndex        =   65
               Top             =   4590
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "国    籍"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   12
               Left            =   9600
               TabIndex        =   64
               Top             =   4590
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "民    族"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   14
               Left            =   3960
               TabIndex        =   63
               Top             =   5070
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "身份证类型"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   15
               Left            =   435
               TabIndex        =   62
               Top             =   5070
               Width           =   900
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "出生地点"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   17
               Left            =   600
               TabIndex        =   61
               Top             =   5550
               Width           =   720
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "关    系"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   16
               Left            =   9600
               TabIndex        =   60
               Top             =   5070
               Width           =   735
            End
            Begin VB.Label lblTitles 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "陪诊人信息"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   5880
               TabIndex        =   59
               Top             =   3720
               Width           =   1455
            End
            Begin VB.Label lblTitles 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "病人信息"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   6000
               TabIndex        =   58
               Top             =   840
               Width           =   1215
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "住    址"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   8
               Left            =   6480
               TabIndex        =   30
               Top             =   2775
               Width           =   720
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "身份证号"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   6480
               TabIndex        =   57
               Top             =   2295
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "身份证类型"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   6
               Left            =   435
               TabIndex        =   56
               Top             =   2295
               Width           =   900
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "出生地点"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   7
               Left            =   600
               TabIndex        =   32
               Top             =   2760
               Width           =   720
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "民    族"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   5
               Left            =   3960
               TabIndex        =   55
               Top             =   2295
               Width           =   735
            End
         End
         Begin VB.PictureBox picCert 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3255
            Left            =   240
            ScaleHeight     =   3255
            ScaleWidth      =   14295
            TabIndex        =   38
            Top             =   6240
            Width           =   14295
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   5
               Left            =   4845
               MaxLength       =   100
               TabIndex        =   25
               Top             =   120
               Width           =   7260
            End
            Begin VB.TextBox txtInfo 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   270
               Index           =   4
               Left            =   1485
               MaxLength       =   20
               TabIndex        =   24
               Top             =   120
               Width           =   1815
            End
            Begin VB.OptionButton optType 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "病人本身"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   6600
               TabIndex        =   26
               Top             =   600
               Value           =   -1  'True
               Width           =   1095
            End
            Begin VB.OptionButton optType 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "陪诊人"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   8040
               TabIndex        =   27
               Top             =   600
               Width           =   855
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCert 
               Height          =   2295
               Left            =   600
               TabIndex        =   28
               Top             =   960
               Width           =   8295
               _cx             =   14631
               _cy             =   4048
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   9
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   325
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfImg 
               Height          =   2295
               Left            =   9000
               TabIndex        =   77
               Top             =   960
               Width           =   3615
               _cx             =   6376
               _cy             =   4048
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   9
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   325
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "备    注"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   19
               Left            =   3960
               TabIndex        =   41
               Top             =   135
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "手 机 号"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   18
               Left            =   600
               TabIndex        =   40
               Top             =   135
               Width           =   735
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "其它证件"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   20
               Left            =   600
               TabIndex        =   39
               Top             =   600
               Width           =   1095
            End
         End
         Begin VB.PictureBox picInterface 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3135
            Left            =   240
            ScaleHeight     =   3135
            ScaleWidth      =   14295
            TabIndex        =   36
            Top             =   9720
            Width           =   14295
            Begin VSFlex8Ctl.VSFlexGrid vsfInterface 
               Height          =   2295
               Left            =   600
               TabIndex        =   29
               Top             =   600
               Width           =   12015
               _cx             =   21193
               _cy             =   4048
               Appearance      =   0
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MousePointer    =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   4210752
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   3
               HighLight       =   2
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   9
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   325
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   ""
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   -1  'True
               AutoSizeMode    =   0
               AutoSearch      =   0
               AutoSearchDelay =   2
               MultiTotals     =   -1  'True
               SubtotalPosition=   1
               OutlineBar      =   0
               OutlineCol      =   0
               Ellipsis        =   0
               ExplorerBar     =   0
               PicturesOver    =   0   'False
               FillStyle       =   0
               RightToLeft     =   0   'False
               PictureType     =   0
               TabBehavior     =   0
               OwnerDraw       =   0
               Editable        =   2
               ShowComboButton =   1
               WordWrap        =   0   'False
               TextStyle       =   0
               TextStyleFixed  =   0
               OleDragMode     =   0
               OleDropMode     =   0
               DataMode        =   0
               VirtualData     =   -1  'True
               DataMember      =   ""
               ComboSearch     =   3
               AutoSizeMouse   =   -1  'True
               FrozenRows      =   0
               FrozenCols      =   0
               AllowUserFreezing=   0
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin VB.Label lblFeild 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "三方认证"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   21
               Left            =   600
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
         End
         Begin MSComCtl2.MonthView monInfo 
            Height          =   2160
            Left            =   0
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3810
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            BorderStyle     =   1
            Appearance      =   0
            StartOfWeek     =   212205569
            TitleBackColor  =   8421504
            TitleForeColor  =   16777215
            CurrentDate     =   38003
         End
      End
   End
   Begin VB.Image imgAdd 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   11295
      Picture         =   "frmCertifyRegist.frx":819A
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDelete 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   10800
      Picture         =   "frmCertifyRegist.frx":8724
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPic 
      Height          =   375
      Left            =   9360
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image ImgCert 
      Height          =   240
      Left            =   10320
      Picture         =   "frmCertifyRegist.frx":9126
      Top             =   240
      Width           =   240
   End
   Begin VB.Image imgDefual 
      Appearance      =   0  'Flat
      Height          =   1515
      Left            =   6480
      Picture         =   "frmCertifyRegist.frx":F978
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgIdentify 
      Height          =   240
      Left            =   3240
      Picture         =   "frmCertifyRegist.frx":10842
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgLoad 
      Height          =   255
      Left            =   2400
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img图片 
      Height          =   240
      Left            =   1440
      Picture         =   "frmCertifyRegist.frx":17094
      Top             =   120
      Width           =   240
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCertifyRegist.frx":1D8E6
   End
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   120
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCertifyRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private mlng病人ID As Long, mlng实名id As Long, mlng证件id As Long, mintModel As Integer
Private mfrmParent As Object '主窗体
Private mrsPati As New ADODB.Recordset '病人实名信息
Private mrsCert As New ADODB.Recordset '实名证件
Private mrsIneterface As New ADODB.Recordset '三方认证接口
Private mbln扫描身份证登记 As Boolean
Private mblnChange As Boolean
Private mblnInfoChange As Boolean  '数据是否发生变化
Private mblnSave As Boolean  '是否已经保存
Private mblnIdentifySure As Boolean '是否已经确定认证
Private mrsMainInfo As ADODB.Recordset  '病人信息主信息记录集
Private mrsSecdInfo  As ADODB.Recordset '列表记录集
Private mblnLoadFilish As Boolean  '是否加载完毕
Private mobjIdentify As Object  '三方认证接口部件
Private mblnInterface As Boolean '三方接口是否认证通过
Private mintDate As Integer  '时间空间索引
Private mstrReason As String '变更原因
Private mlng图像操作 As Long  '病人图片操作类型
Private mstr采集图片 As String '采集的图片路径
Private mlngImage As Long '证件图片存放序号
Private mlngPati As Long
Private mlngTopVsc As Long
Private mblnChange基本 As Boolean
Private mstrAge As String
Private mstrMsg As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1

Private Enum TXT_Info
    TXT_姓名 = 0
    TXT_身份证号 = 1
    TXT_陪诊人姓名 = 2
    TXT_陪诊人身份证号 = 3
    txt_手机号 = 4
    TXT_备注 = 5
End Enum

Private Enum DATE_Info
    DATE_出生日期 = 0
    DATE_陪诊人出生日期 = 1
End Enum

Private Enum PatiAress_Info
    ADRS_出生地点 = 0
    ADRS_住址 = 1
    ADRS_陪诊人住址 = 2
End Enum

Private Enum VSFCert_COL
    COL_证件ID = 0
    COL_证件号码
    COL_证件类型
    COL_备注
    COL_所有者
    COL_图片
    COL_增加
    COL_Del
End Enum

Private Enum VSFIMG_COL
    IMG_证件ID = 0
    IMG_序号
    IMG_图片
    IMG_备注
    IMG_Del
End Enum

Private Enum VSFInterface_COL
    COLS_接口ID = 0
    COLS_名称
    COLS_部件名
    COLS_说明
    COLS_认证结果
    COLS_认证
End Enum

Private Enum Change_State
    CS_删除行 = -1
    CS_未改变 = 0
    CS_更新行 = 1
    CS_替换行 = 2
    CS_新增行 = 3
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
 GdiplusVersion As Long
 DebugEventCallback As Long
 SuppressBackgroundThread As Long
 SuppressExternalCodecs As Long
End Type

Private Type EncoderParameter
 GUID As GUID
 NumberOfValues As Long
 type As Long
 Value As Long
End Type

Private Type EncoderParameters
 Count As Long
 Parameter As EncoderParameter
End Type

Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, ID As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long

Private Function DrawLin()
'功能：给控件画线
    Dim objText As Object
    
    For Each objText In Me.Controls
        If TypeName(objText) = "TextBox" Or TypeName(objText) = "Frame" Then
            If objText.Name <> "txtAdressInfo" Then
                DrawLineCTL objText
            ElseIf objText.Name = "txtAdressInfo" Then
                If Not gbln启用结构化地址 Then
                    DrawLineCTL objText
                End If
            End If
        End If
    Next
End Function

Private Sub InitVsfGridHeader()
'功能：初始化列表
    Dim strHeader As String
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
     
    '证件信息列表
    strHeader = "证件ID;证件号码,2000,1;证件类型,2000,1;备注,2050,1;所有者,1000,4;,270,4;,270,4;,270,4"
    Call grid.Init(vsfCert, strHeader)
    With vsfCert
        If Not .ColHidden(COL_证件类型) Then
            strSQL = "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '证件类型' 表名 From 证件类型"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "编码", "名称")
            Else
                strTmp = " |"
            End If
            .ColData(COL_证件类型) = strTmp
        End If
    End With
    
    '三方认证接口
    strHeader = "接口ID;名称,3000,1;部件名;说明,6000,1;认证结果,2000,4;,270,4"
    Call grid.Init(vsfInterface, strHeader)
    
    '图片列表
    strHeader = "证件ID;序号;图片,600,4;备注,2650,1;,270,4"
    Call grid.Init(vsfImg, strHeader)
End Sub

Private Sub DrawLineCTL(ByRef objCtl As Object, Optional ByVal bytModel As Byte = 0)
'功能:给指定对象画一条线或清除此原有线条
'objCtl-传入控件对象，根据该控件对象获取对应坐标值
'bytModel=0-画线;1-清除线
    Dim objPic As Object  '容器
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    Select Case TypeName(objCtl)
    Case "TextBox"
        '在每个TextBox 下面画一条线
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width
        y2 = y1
    Case "Frame"
        x1 = objCtl.Left
        y1 = objCtl.Top + objCtl.Height + 3
        x2 = objCtl.Left + objCtl.Width - 60
        y2 = y1
    End Select
    Set objPic = objCtl.Container
    objPic.DrawWidth = 1
    If bytModel = 0 Then
        objPic.Line (x1, y1)-(x2, y2)
    Else
        objPic.Line (x1, y1)-(x2, y2), objPic.BackColor '清除线条
    End If
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Call CheckValueChange(cboInfo(Index))
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strReason As String

    Select Case Control.ID
        Case conMenu_Certify_Save
            If mblnInfoChange Then
                If Not CheckCertifyData Then Exit Sub
                Call CachPatiData
                Call CachCertInterface
                If mintModel = 1 And mblnChange基本 Then
                    frmGetReason.ShowMe Me, strReason
                    mstrReason = strReason
                    If mstrReason = "" Then
                        Exit Sub
                    End If
                End If
                
                If SaveCertifyData(0, mintModel) Then
                    mblnSave = True
                    If mstrMsg <> "" Then
                        MsgBox mstrMsg, vbInformation, gstrSysName
                    End If
                    mstrMsg = ""
                    mblnInfoChange = False
                    mblnChange基本 = False
                    mintModel = 1
                    Call CachAllData
                Else
                    mblnSave = False
                    If mstrMsg <> "" Then
                        MsgBox mstrMsg, vbInformation, gstrSysName
                    End If
                    mstrMsg = ""
                End If
            End If
        Case conMenu_Certify_IdentifySure
            If mblnInterface Then
                If SaveCertifyData(1, mintModel) Then
                    mblnIdentifySure = True
                Else
                    mblnIdentifySure = False
                End If
            Else
                If MsgBox("三方认证接口未认证,是否继续人工认证？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If SaveCertifyData(1, mintModel) Then
                        mblnIdentifySure = True
                    Else
                        mblnIdentifySure = False
                    End If
                Else
                    Exit Sub
                End If
            End If
        Case conMenu_Certify_Cancel
            If mblnIdentifySure Then
                If MsgBox("该病人的实名信息已经认证,确定要取消认证吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If SaveCertifyData(2, mintModel) Then
                        mblnIdentifySure = False
                    Else
                        mblnIdentifySure = True
                    End If
                End If
            End If
        Case conMenu_CertifyHelp_Help
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case conMenu_Certify_Quit
            Unload Me
    End Select
End Sub

Private Function CachCertData() As String
'功能：将证件信息缓存
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim strDels As String
    Dim strAll As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    Dim strAllInfo As String
    
    On Error GoTo errH
    With vsfCert
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_证件号码) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_证件ID)) & "|" & zlCommFun.GetNeedName(.TextMatrix(i, COL_证件类型), "-") & "|" & .TextMatrix(i, COL_证件号码)
                strInfo = strMainInfo & "|" & .TextMatrix(i, COL_备注) & "|" & .TextMatrix(i, COL_所有者)
 
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '相同过每记录
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, COL_证件号码) = &HC0C0FF
                    Call .ShowCell(i, COL_证件号码)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '收集所有诊断用于判断是否有重复行
                End If
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "控件名='vsfCert' and 序号=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!序号 = lngTmp
                   mrsSecdInfo!控件名 = "vsfCert"
               End If
               mrsSecdInfo!现ID = Val(.RowData(i))
               mrsSecdInfo!信息现值 = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!主信息现值 = IIf(strMainInfo = "", Null, strMainInfo)
               mrsSecdInfo!IndexEx = i
               mrsSecdInfo.Update
               lngTmp = lngTmp + 1

               mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "控件名='vsfCert'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
        
        '主信息改变行需要调用删除方法
        mrsSecdInfo.Filter = "(改变状态=" & CS_删除行 & " And 控件名='vsfCert')" ' OR (改变状态=" & CS_替换行 & " And 控件名='vsfCert')"
        Do While Not mrsSecdInfo.EOF
            strDels = "" & mrsSecdInfo!原ID
            If strDels <> "" Then
                strAllInfo = strAllInfo & "," & mlng实名id & "-" & Val(strDels) & "-----"
            End If
            mrsSecdInfo.MoveNext
        Loop
        mrsSecdInfo.Filter = "控件名='vsfCert' And 改变状态>" & CS_未改变
        If Not mrsSecdInfo.EOF Then
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                If mrsSecdInfo!改变状态 = CS_新增行 Then
                    strAllInfo = strAllInfo & "," & mlng实名id & "-" & mrsSecdInfo!原ID & "-" & zlCommFun.GetNeedName(.TextMatrix(lngRow, COL_证件类型), "-") & "-" & .TextMatrix(lngRow, COL_证件号码) & "-" & .TextMatrix(lngRow, COL_备注) & "-" & IIf(.TextMatrix(lngRow, COL_所有者) = "病人本身", 1, 2)
                Else
                    strAllInfo = strAllInfo & "," & mlng实名id & "-" & mrsSecdInfo!原ID & "-" & zlCommFun.GetNeedName(.TextMatrix(lngRow, COL_证件类型), "-") & "-" & .TextMatrix(lngRow, COL_证件号码) & "-" & .TextMatrix(lngRow, COL_备注) & "-" & IIf(.TextMatrix(lngRow, COL_所有者) = "病人本身", 1, 2)
                End If
                mrsSecdInfo.MoveNext
            Loop
        Else
            mrsSecdInfo.Filter = "控件名='vsfCert' And 改变状态=" & CS_未改变
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                strAllInfo = strAllInfo & "," & mlng实名id & "-" & mrsSecdInfo!原ID & "-" & .TextMatrix(lngRow, COL_证件类型) & "-" & .TextMatrix(lngRow, COL_证件号码) & "-" & .TextMatrix(lngRow, COL_备注) & "-" & IIf(.TextMatrix(lngRow, COL_所有者) = "病人本身", 1, 2)
                mrsSecdInfo.MoveNext
            Loop
        End If
    End With
    CachCertData = Mid(strAllInfo, 2)
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function CachCertImgData(ByRef arrSQL As Variant) As Boolean
'功能：将图片信息缓存
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strsInfo As String
    Dim strsMainInfo As String
    Dim strDels As String
    Dim strAll As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long

    On Error GoTo errH
    With vsfImg
       .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, IMG_证件ID) <> "" Then
                strsInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_图片, i, IMG_图片) & "|" & .TextMatrix(i, IMG_序号) & "|" & .TextMatrix(i, IMG_备注)
                strsMainInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_图片, i, IMG_图片) & "|" & .TextMatrix(i, IMG_序号) & "|" & .TextMatrix(i, IMG_备注)
            Else
                strsInfo = ""
                strsMainInfo = ""
            End If
            mrsSecdInfo.Filter = "控件名='vsfImg' and 序号=" & lngTmp
            If mrsSecdInfo.EOF Then
                mrsSecdInfo.AddNew
                mrsSecdInfo!序号 = lngTmp
                mrsSecdInfo!控件名 = "vsfImg"
            End If
            mrsSecdInfo!现ID = Val(.RowData(i))
            mrsSecdInfo!信息现值 = IIf(strsInfo = "", Null, strsInfo)
            mrsSecdInfo!主信息现值 = IIf(strsMainInfo = "", Null, strsMainInfo)
            mrsSecdInfo!IndexEx = i
            mrsSecdInfo.Update
            lngTmp = lngTmp + 1

            mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "控件名='vsfImg'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
        

        '删除行以及主信息改变行需要调用删除方法
        mrsSecdInfo.Filter = "(改变状态=" & CS_删除行 & " And 控件名='vsfImg')" ' OR (改变状态=" & CS_替换行 & " And 控件名='vsfImg')"
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            strDels = "" & .TextMatrix(lngRow, IMG_序号)
            If Val(strDels) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人实名证件图片_Delete(" & .RowData(lngRow) & "," & Val(strDels) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop
    
        '主信息改变以及新增行需要调用插入过程        '次级信息改变，调用更新过程
        mrsSecdInfo.Filter = "控件名='vsfImg' And 改变状态>" & CS_未改变
    
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            If mrsSecdInfo!改变状态 <> CS_新增行 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_病人实名证件图片_Update(" & Val(.RowData(lngRow)) & "," & Val(.TextMatrix(lngRow, IMG_序号)) & ",'" & .TextMatrix(lngRow, IMG_备注) & "')"
            Else
                 Call SaveCertPicture(Val(.RowData(lngRow)), Val(.TextMatrix(lngRow, IMG_序号)), .TextMatrix(lngRow, IMG_备注), .Cell(flexcpData, lngRow, IMG_图片, lngRow, IMG_图片))
            End If
            mrsSecdInfo.MoveNext
        Loop
    End With
    CachCertImgData = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub zlRefresh()
'新增病人实名信息后刷新数据
    Dim strSQL As String
    Dim rsPati As New ADODB.Recordset
    Dim i As Long
    Dim str号码 As String, str证件类型 As String
    
    On Error GoTo errH
        strSQL = "Select 病人ID,实名ID  From 病人实名信息 where 姓名=[1] And 身份证号=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "刷新病人信息", txtInfo(TXT_姓名).Text, txtInfo(TXT_身份证号).Text)
        If rsPati.EOF Then
            strSQL = "Select 病人ID,实名ID From 病人实名信息 where 姓名=[1] And 陪诊人姓名=[2] And 陪诊人身份证号=[3]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "刷新病人信息", txtInfo(TXT_姓名).Text, txtInfo(TXT_陪诊人姓名).Text, txtInfo(TXT_陪诊人身份证号).Text)
            If rsPati.EOF Then
                With vsfCert
                    For i = .FixedRows To .Rows - 1
                        str号码 = .TextMatrix(i, COL_证件号码)
                        str证件类型 = .TextMatrix(i, COL_证件类型)
                        strSQL = "Select A.病人ID,A.实名ID From 病人实名信息 A,病人实名证件 B where A.实名ID=B.实名ID And 姓名=[1] And B.证件类型=[2] And B.证件号码=[3]" & IIf(optType(0).Value, " And B.所有者=[4]", " And A.陪诊人姓名=[4] And B.所有者=[5]")
                        If optType(0).Value Then
                            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "刷新病人信息", txtInfo(TXT_姓名).Text, str证件类型, str号码, 1)
                        Else
                            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "刷新病人信息", txtInfo(TXT_姓名).Text, str证件类型, str号码, txtInfo(TXT_陪诊人姓名).Text, 2)
                        End If
                        If Not rsPati.EOF Then
                            Exit For
                        End If
                    Next
                End With
            End If
        End If
    If Not rsPati.EOF Then
        mlng实名id = rsPati!实名ID & ""
        mlng病人ID = rsPati!病人ID & ""
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub zlRefreshCert()
'新增实名证件信息就刷新数据
    Dim strSQL As String
    Dim rsCert As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim str号码 As String, str证件类型 As String
    
    With vsfCert
        For i = .FixedRows To .Rows - 1
            str号码 = .TextMatrix(i, COL_证件号码)
            str证件类型 = .TextMatrix(i, COL_证件类型)
            strSQL = "Select A.病人ID,A.实名ID,B.ID as 证件ID From 病人实名信息 A,病人实名证件 B where A.实名ID=B.实名ID And A.实名ID=[1] And 证件类型=[2] And 证件号码=[3]"
            Set rsCert = zlDatabase.OpenSQLRecord(strSQL, "刷新病人信息", mlng实名id, str证件类型, str号码)
            If Not rsCert.EOF Then
                .Cell(flexcpData, i, COL_证件ID, i, COL_证件ID) = i & ""
                .TextMatrix(i, COL_证件ID) = rsCert!证件ID & ""
            Else
                .Cell(flexcpData, i, COL_证件ID, i, COL_证件ID) = i & ""
                .TextMatrix(i, COL_证件ID) = zlDatabase.GetNextId("病人实名证件") & ""
            End If
            With vsfImg
                For j = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, j, IMG_证件ID, j, IMG_证件ID) = "" & i & "-" & j Then
                        .RowData(j) = Val(vsfCert.TextMatrix(i, COL_证件ID))
                    End If
                Next
            End With
        Next
    End With
End Sub

Private Function CachCertInterface() As Boolean
'功能：将证件信息缓存
'功能：获取诊断保存的SQL
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    
    On Error GoTo errH
    With vsfInterface
        .Tag = ""
        strVsName = .Name
        strTmp = ""
        lngTmp = 1
        arrMain = Array(COLS_名称, COLS_部件名, COLS_说明, COLS_认证结果, COLS_认证)
        arrWhole = Array(COLS_接口ID, COLS_部件名, COLS_名称, COLS_说明, COLS_认证结果, COLS_认证)
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COLS_认证结果) <> "" Then
                If strTmp <> .TextMatrix(i, COLS_认证结果) Then
                    j = 1: strTmp = .TextMatrix(i, COLS_认证结果)
                Else
                    j = j + 1
                End If
                strInfo = j: strMainInfo = ""
                For k = LBound(arrWhole) To UBound(arrWhole)
                    strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                Next
                For k = LBound(arrMain) To UBound(arrMain)
                    strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                Next
                mrsSecdInfo.Filter = "控件名='" & strVsName & "' and 序号=" & lngTmp
 
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!序号 = lngTmp
                    mrsSecdInfo!控件名 = strVsName
                End If
                mrsSecdInfo!信息现值 = strInfo
                mrsSecdInfo!主信息现值 = strMainInfo
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
                mrsSecdInfo.Filter = 0
            End If
        Next
        mrsSecdInfo.Filter = "控件名='" & strVsName & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng状态 = CS_未改变
            If mrsSecdInfo!信息原值 & "" <> mrsSecdInfo!信息现值 & "" Then
                lng状态 = CS_更新行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息原值) Then
                lng状态 = CS_新增行
            End If
            If lng状态 = CS_更新行 And IsNull(mrsSecdInfo!信息现值) Then
                lng状态 = CS_删除行
            End If
            If lng状态 = CS_更新行 And mrsSecdInfo!主信息原值 & "" <> mrsSecdInfo!主信息现值 & "" Then
                lng状态 = CS_替换行
            End If
            mrsSecdInfo.Update "改变状态", lng状态
            mrsSecdInfo.MoveNext
        Next
    End With
    CachCertInterface = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function CachPatiData() As Boolean
'功能：缓存病人实名信息
    Dim strCtlName As String, strFilter As String, strValue As String
    Dim objCtl As Object
    Dim strBirthdate As String
    
    On Error GoTo errH
    For Each objCtl In Me.Controls
        strCtlName = objCtl.Name
        Select Case strCtlName
            Case "txtInfo"
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = objCtl.Text
                    mrsMainInfo!信息现值 = strValue
                    mrsMainInfo.Update
                End If
            Case "cboInfo"
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = zlCommFun.GetNeedName(objCtl.Text, "-")
                    mrsMainInfo!信息现值 = strValue
                    mrsMainInfo.Update
                End If
            Case "txtDateInfo"
                If objCtl.Mask = "####-##-## ##:##" Then
                    If Format(Mid(objCtl.Text, 12), "HH:MM") = "__:__" Then
                        strBirthdate = Format(Mid(objCtl.Text, 1, 10), "YYYY-MM-DD")
                    Else
                        strBirthdate = Format(objCtl.Text, "YYYY-MM-DD HH:MM")
                    End If
                ElseIf objCtl.Mask = "####-##-##" Then
                    strBirthdate = Format(objCtl.Text, "YYYY-MM-DD")
                End If
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    If IsDate(strBirthdate) Then
                        strValue = strBirthdate
                    Else
                        strValue = ""
                    End If
                    mrsMainInfo!信息现值 = strValue
                    mrsMainInfo.Update
                End If
            Case "txtAdressInfo"
                If gbln启用结构化地址 Then
                    strFilter = "控件名='patiAdressInfo ' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = patiAdressInfo(objCtl.Index).Value
                        mrsMainInfo!信息现值 = strValue
                        mrsMainInfo.Update
                    End If
                Else
                    strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = objCtl.Text
                        mrsMainInfo!信息现值 = strValue
                        mrsMainInfo.Update
                    End If
                End If
        End Select
    Next
    CachPatiData = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function CachAllData() As Boolean
    Dim strCtlName As String, strFilter As String, strValue As String
    Dim objCtl As Object
    Dim i As Long, j As Long, k As Long
    Dim lng状态 As Long
    Dim strTmp As String
    Dim strInfo As String
    Dim strMainInfo As String
    Dim strsInfo As String
    Dim strsMainInfo As String
    Dim strDels As String
    Dim strAll As String
    Dim lngRow As Long
    Dim strVsName As String
    Dim arrWhole As Variant
    Dim arrOther As Variant
    Dim arrMain As Variant
    Dim DatCur As Date
    Dim lngID As Long
    Dim lngTmp As Long
    
    On Error GoTo errH
    For Each objCtl In Me.Controls
        strCtlName = objCtl.Name
        Select Case strCtlName
            Case "txtInfo"
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = objCtl.Text
                    mrsMainInfo!信息原值 = strValue
                    mrsMainInfo.Update
                End If
            Case "cboInfo"
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = zlCommFun.GetNeedName(objCtl.Text, "-")
                    mrsMainInfo!信息原值 = strValue
                    mrsMainInfo.Update
                End If
            Case "txtDateInfo"
                strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    If IsDate(objCtl.Text) Then
                        strValue = objCtl.Text
                    Else
                        strValue = ""
                    End If
                    mrsMainInfo!信息原值 = strValue
                    mrsMainInfo.Update
                End If
            Case "txtAdressInfo"
                If gbln启用结构化地址 Then
                    strFilter = "控件名='patiAdressInfo ' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = patiAdressInfo(objCtl.Index).Value
                        mrsMainInfo!信息原值 = strValue
                        mrsMainInfo.Update
                    End If
                Else
                    strFilter = "控件名='" & strCtlName & "' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = objCtl.Text
                        mrsMainInfo!信息原值 = strValue
                        mrsMainInfo.Update
                    End If
                End If
        End Select
    Next
    
    With vsfCert
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_证件号码) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_证件ID)) & "|" & zlCommFun.GetNeedName(.TextMatrix(i, COL_证件类型), "-") & "|" & .TextMatrix(i, COL_证件号码)
                strInfo = strMainInfo & "|" & .TextMatrix(i, COL_备注) & "|" & .TextMatrix(i, COL_所有者)
                .RowData(i) = .TextMatrix(i, COL_证件ID)
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "控件名='vsfCert' and 序号=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!序号 = lngTmp
                   mrsSecdInfo!控件名 = "vsfCert"
               End If
               mrsSecdInfo!原ID = Val(.RowData(i))
               mrsSecdInfo!信息原值 = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!主信息原值 = IIf(strMainInfo = "", Null, strMainInfo)
               mrsSecdInfo!IndexEx = i
               mrsSecdInfo.Update
               lngTmp = lngTmp + 1
               mrsSecdInfo.Filter = 0
        Next
    End With
    
    With vsfImg
       .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, IMG_证件ID) <> "" Then
                strsInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_图片, i, IMG_图片) & "|" & .TextMatrix(i, IMG_序号) & "|" & .TextMatrix(i, IMG_备注)
                strsMainInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_图片, i, IMG_图片) & "|" & .TextMatrix(i, IMG_序号) & "|" & .TextMatrix(i, IMG_备注)
            Else
                strsInfo = ""
                strsMainInfo = ""
            End If
            mrsSecdInfo.Filter = "控件名='vsfImg' and 序号=" & lngTmp
            If mrsSecdInfo.EOF Then
                mrsSecdInfo.AddNew
                mrsSecdInfo!序号 = lngTmp
                mrsSecdInfo!控件名 = "vsfImg"
            End If
            mrsSecdInfo!原ID = Val(.RowData(i))
            mrsSecdInfo!信息原值 = IIf(strsInfo = "", Null, strsInfo)
            mrsSecdInfo!主信息原值 = IIf(strsMainInfo = "", Null, strsMainInfo)
            mrsSecdInfo!IndexEx = i
            mrsSecdInfo.Update
            lngTmp = lngTmp + 1

            mrsSecdInfo.Filter = 0
        Next
    End With
    CachAllData = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function SaveCertifyData(ByVal intTYPE As Integer, ByVal intModel As Integer) As Boolean
'功能：保存数据
'    intTYPE: 0-保存 1-认证确定 2-取消
'    intModel: 0-新增 1-修改
    Dim arrSQL() As Variant
    Dim strCertifySql As String
    Dim strUpdateSql As String
    Dim blnTrans As Boolean
    Dim i As Long, j As Long, k As Long
    Dim strFile As String, strTmp As String, strArrTmp As Variant
    Dim lng证件ID As Long, lng序号 As Long
    Dim strIDs As String, strArry As Variant
    Dim str身份证 As String, strName As String, strSex As String, strAge As String
    Dim strJsonAsk As String, strJsonOut As String, strCertInfo As String
    Dim lng病人ID As Long, lng实名ID As Long, lng场合 As Long, lng就诊ID As Long, lngID As Long
    Dim blnNew As Boolean, blnCheck As Boolean, blnNotChangeAge As Boolean
    Dim arrInfo  As Variant
    Dim strSQL As String, strBirthdate As String
    Dim rsTmp As ADODB.Recordset
    Dim str姓名 As String, str年龄 As String, str性别 As String, str出生日期 As String, strInfo As String, strMsg As String, strExpalin As String
    Dim blnIn As Boolean
    Dim str原姓名 As String, str原性别 As String, str原年龄 As String, str原出生日期 As String
    Dim str调整说明 As String
    Dim str变更时间 As String
    
    On Error GoTo errH
    arrSQL = Array()
    If Not mrsMainInfo Is Nothing Then
        mrsMainInfo.Filter = "信息名='身份证号'"
        If Not mrsMainInfo.EOF Then str身份证 = mrsMainInfo!信息现值 & ""
        
        mrsMainInfo.Filter = "信息名='姓名'"
        If Not mrsMainInfo.EOF Then str姓名 = mrsMainInfo!信息现值 & ""

        mrsMainInfo.Filter = "信息名='性别'"
        If Not mrsMainInfo.EOF Then str性别 = mrsMainInfo!信息现值 & ""

        If IsDate(txtDateInfo(DATE_出生日期).Text) Then
            str年龄 = GetAge(txtDateInfo(DATE_出生日期).Text, mlng病人ID, zlDatabase.Currentdate)
        End If
        
        mrsMainInfo.Filter = "信息名='出生日期'"
        If Not mrsMainInfo.EOF Then str出生日期 = Format(mrsMainInfo!信息现值 & "", "yyyy-mm-dd hh:mm:ss")
    End If
    If intTYPE = 0 Then
        '保存前的检查
        blnCheck = True
        strCertInfo = GetCertifyData(3, 0, 0, 0)
        '姓名|性别|年龄|出生日期|陪诊人姓名|陪诊人性别|陪诊人出生日期|身份证号|陪诊人身份证号|陪诊人关系|所有者|证件信息
        arrInfo = Split(strCertInfo, "|")
        strJsonAsk = "{""input"":{""opr_fun"":" & intModel & "," & IIf(intModel = 1, """real_id"":" & mlng实名id & ",", "") & """pati_name"":""" & arrInfo(0) & """,""pati_sex"":""" & arrInfo(1) & """,""pati_age"":""" & arrInfo(2) & """,""pati_birthdate"":""" & arrInfo(3) & """,""pati_idcard"":""" & arrInfo(7) & """,""owner"":" & arrInfo(10) & ",""grdn_name"":""" & arrInfo(4) & """,""grdn_sex"":""" & arrInfo(5) & """,""grdn_birthdate"":""" & arrInfo(6) & """,""grdn_idcard"":""" & arrInfo(8) & """,""grdn_relation"":""" & arrInfo(9) & """,""papers_info"":""" & arrInfo(11) & """}}"
        If Not CallService("Zl_Patisvr_Patirealnamecheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
            blnCheck = False
        Else
            If intModel = 0 Then
                lng病人ID = gobjService.GetJsonNodeValue("output.pati_id")
                lng实名ID = gobjService.GetJsonNodeValue("output.real_id")
                blnNew = gobjService.GetJsonNodeValue("output.new_pati") = 1
            End If
        End If
        strJsonOut = ""
        If Not blnCheck Then Exit Function
        If Not blnNew Then
            '获取挂号信息
            strJsonAsk = "{""input"":{""query_type"":0,""occasion"":0,""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & "}}"
            strJsonOut = ""
            If CallService("Zl_Cissvr_Getpativisitid", strJsonAsk, strJsonOut) Then
                lngID = gobjService.GetJsonNodeValue("output.visit_id")
                lng场合 = gobjService.GetJsonNodeValue("output.occasion")
            End If
            '更新病人基本信息的检查
            blnCheck = False
            strSQL = "Select 姓名, 性别, 年龄, 出生日期 From 病人信息 Where 病人id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "更新病人基本信息的检查", IIf(intModel = 0, lng病人ID, mlng病人ID))
            If Not rsTmp.EOF Then
                strName = rsTmp!姓名 & ""
                strSex = rsTmp!性别 & ""
                strAge = rsTmp!年龄 & ""
                strBirthdate = Format(rsTmp!出生日期 & "", "yyyy-mm-dd hh:mm:ss")
                If rsTmp!年龄 & "" <> str年龄 Then
                    '不更新年龄的判断逻辑
                    '新生儿年龄不用更新
                    If rsTmp!年龄 & "" Like "*小时%分钟" Or rsTmp!年龄 & "" Like "*分钟" Or rsTmp!年龄 & "" Like "*天*小时" Or rsTmp!年龄 & "" Like "*小时" Then
                        blnNotChangeAge = True
                    Else
                        blnNotChangeAge = False
                    End If
                End If
            End If
            If blnNotChangeAge Then
              If strName <> str姓名 Or strBirthdate <> str出生日期 Or strSex <> str性别 Then
                 blnCheck = True
              End If
              strAge = strAge
            Else
              If strName <> str姓名 Or strBirthdate <> str出生日期 Or strSex <> str性别 Or strAge <> str年龄 Then
                blnCheck = True
              End If
              strAge = str年龄
            End If
            If blnCheck Then
                blnCheck = True
                If lngID = 0 Then
                    strSQL = "Select 姓名, 性别, 年龄, 出生日期 From 病人信息 Where 病人id = [1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人信息", mlng病人ID)
                    If rsTmp.EOF Then
                        MsgBox "病人ID[" & lng病人ID & "]在病人信息中不存在,不能继续进行病人信息变更操作!", vbInformation, gstrSysName
                        Exit Function
                    Else
                        str原姓名 = rsTmp!姓名 & ""
                        str原年龄 = rsTmp!年龄 & ""
                        str原性别 = rsTmp!性别 & ""
                        str原出生日期 = Format(rsTmp!出生日期 & "", "YYYY-MM-DD HH:MM:SS")
                    End If
                Else
                    strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng场合 & ",""pati_name"":""" & strName & """,""pati_age"":""" & strAge & """,""pati_sex"":""" & strSex & """,""pati_birthdate"":""" & strBirthdate & """}}"
                    If Not CallService("Zl_Cissvr_Checkpatexist", strJsonAsk, strJsonOut, , , False, , , , True) Then
                        blnCheck = False
                    Else
                        str原姓名 = gobjService.GetJsonNodeValue("output.pati_name")
                        str原性别 = gobjService.GetJsonNodeValue("output.pati_sex")
                        str原年龄 = gobjService.GetJsonNodeValue("output.pati_age")
                        str原出生日期 = gobjService.GetJsonNodeValue("output.pati_birthdate")
                    End If
                End If
                strJsonOut = ""
                If Not blnCheck Then Exit Function
                strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & "}}"
                If Not CallService("Zl_Patisvr_Lockcheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
                    blnCheck = False
                End If
                If Not blnCheck Then Exit Function
                strJsonOut = ""
                strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng场合 & "}}"
                If CallService("Zl_Exsesvr_Updpatbaseinfocheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
                    strExpalin = gobjService.GetJsonNodeValue("output.explain")
                Else
                    blnCheck = False
                End If
                strJsonOut = ""
                If Not CallService("Zl_Cissvr_Updpatbaseinfocheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
                    blnCheck = False
                End If
            Else
                blnCheck = True
            End If
            If Not blnCheck Then Exit Function
        End If
    End If
    strJsonOut = ""
    If lngID <> 0 Then
        strJsonAsk = "{""input"":{""pati_id"":" & IIf(intTYPE = 0, IIf(intModel = 0, lng病人ID, mlng病人ID), mlng病人ID) & ",""pati_pageid"":" & lngID & "}}"
        If CallService("zl_cissvr_checkpativisitorin", strJsonAsk, strJsonOut, , , False) Then
            blnIn = Val(gobjService.GetJsonNodeValue("output.isexist")) = 1
        End If
    End If
    strJsonOut = ""
    strCertifySql = GetCertifyData(intTYPE, intModel, IIf(intTYPE = 0, IIf(intModel = 0, lng病人ID, mlng病人ID), mlng病人ID), IIf(intTYPE = 0, IIf(intModel = 0, lng实名ID, mlng实名id), mlng实名id), blnNew, IIf(blnNew = True, 0, lngID))
    If strCertifySql <> "" Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strCertifySql
    End If
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If Not blnNew Then
        str变更时间 = zlDatabase.CallProcedure("Zl_病人信息_基本信息调整_s", "基本信息调整", IIf(intModel = 0, lng病人ID, mlng病人ID), 1109, str姓名, str性别, str年龄, Format(str出生日期, "YYYY-MM-DD HH:MM"), str原姓名, str原性别, str原年龄, Format(str原出生日期, "YYYY-MM-DD HH:MM"), Empty)
    End If
    strJsonOut = ""
    If intTYPE = 0 Then
        If Not blnNew Then
    '        strInfo = zlDatabase.CallProcedure("Zl_病人信息_基本信息调整", "基本信息调整", IIf(intModel = 0, lng病人ID, mlng病人ID), lngID, 1109, str姓名, str性别, strAge, CDate(str出生日期), lng场合, "病人实名信息认证", str原姓名, str性别, str原年龄, IIf(IsDate(str原出生日期), CDate(str原出生日期), "Null"), Empty)
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng场合 & ",""update_info"":{""pati_name"":""" & str姓名 & """,""pati_age"":""" & str年龄 & """,""pati_sex"":""" & str性别 & """,""pati_birthdate"":""" & str出生日期 & """}}}"
            If CallService("Zl_Cissvr_Updatepatibaseinfo", strJsonAsk, strJsonOut, , , False, , , , False) Then
                strMsg = gobjService.GetJsonNodeValue("output.adjust_explain")
                If strMsg <> "" Then
                    strInfo = strInfo & strMsg
                End If
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    If str变更时间 = "" Then str变更时间 = zlDatabase.Currentdate
    strJsonOut = ""
    If intTYPE = 0 Then
        If Not blnNew Then
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng场合 & ",""update_info"":{""pati_name"":""" & str姓名 & """,""pati_age"":""" & strAge & """,""pati_sex"":""" & str性别 & """,""pati_birthdate"":""" & str出生日期 & """,""explain"":""" & strExpalin & """}}}"
            If CallService("Zl_Exsesvr_Updatepatibaseinfo", strJsonAsk, strJsonOut) Then
                strMsg = gobjService.GetJsonNodeValue("output.adjust_explain")
                If strMsg <> "" Then
                    strInfo = strInfo & strMsg
                End If
            End If
            If strInfo <> "" Then
                strInfo = Mid(strInfo, 3)
                strInfo = "修改原因:病人实名信息认证" & Chr(13) & "病人基本信息调整导致以下内容发生变化:" & Chr(13) & strInfo
            End If
            str调整说明 = strInfo
            strInfo = Replace(strInfo, Chr(13), " ")
            strInfo = Replace(strInfo, Chr(10), " ")
            strJsonOut = ""
'            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人id, mlng病人ID) & ",""visit_id"":" & lngID & ",""model"":""" & "实名信息管理" & """,""pati_name_n"":""" & str姓名 & """,""pati_sex_n"":""" & str性别 & """,""pati_age_n"":""" & str年龄 & """,""pati_birthdate_n"":""" & Format(str出生日期, "YYYY-MM-DD HH:MM:SS") & """,""occasion"":" & lng场合 & ",""pati_name_o"":""" & str原姓名 & """,""pati_sex_o"":""" & str原性别 & """,""pati_age_o"":""" & str原年龄 & """,""pati_birthdate_o"":""" & Format(str原出生日期, "YYYY-MM-DD HH:MM:SS") & """,""explain"":""" & strInfo & """}}"
'            Call CallService("Zl_Patisvr_Updatepatibaseinfo", strJsonAsk, strJsonOut)
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng病人ID, mlng病人ID) & ",""visit_id"":" & lngID & ",""pati_name"":""" & str姓名 & """,""pati_sex"":""" & str性别 & """,""pati_age"":""" & strAge & """}}"
            Call CallService("Zl_Pivassvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            strJsonOut = ""
            Call CallService("Zl_Drugsvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            strJsonOut = ""
            Call CallService("Zl_Stuffsvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            Call UpdateChangeInfo(IIf(intModel = 0, lng病人ID, mlng病人ID), strInfo, CDate(str变更时间))
        End If
    End If
    If str调整说明 <> "" Then
        mstrMsg = str调整说明
    End If
    If intTYPE = 0 Then
        If intModel = 0 Then
            mlng病人ID = lng病人ID
            mlng实名id = lng实名ID
            Call zlRefresh
        End If
        SavePatPicture mlng病人ID
        Call zlRefreshCert
        arrSQL = Array()
        Call CachCertImgData(arrSQL)
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        gcnOracle.CommitTrans: blnTrans = False
    End If
    Screen.MousePointer = 0
    SaveCertifyData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function UpdateChangeInfo(ByVal lng病人ID As Long, ByVal strInfo As String, ByVal d变动时间 As Date)
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Zl_病人信息变动_Update(" & lng病人ID & ",'" & strInfo & "'," & zlStr.To_Date(d变动时间, "ymdhms") & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SavePatPicture(lng病人ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存病人照片
    '入参:lng病人ID - 病人ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSQL As String
    
    On Error GoTo Errhand
    Select Case mlng图像操作
        Case 1 '文件
            strFile = cmdialog.filename
        Case 2 '采集
            strFile = mstr采集图片
            mstr采集图片 = ""
        Case 4 '二代身份证
            strFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, strFile
    End Select
    If InStr(1, ",1,2,4,", "," & mlng图像操作 & ",") <> 0 Then
        If strFile = "" Then Exit Sub
        Call PictureBoxSaveJPG(imgPatient.Picture, strFile) '保存压缩后的图片
        If Sys.SaveLob(glngSys, 27, mlng病人ID, strFile) = False Then
            MsgBox "保存照片失败,文件可能被删除!", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf mlng图像操作 = 3 Then
        strSQL = strSQL & "Zl_病人照片_Delete("
        strSQL = strSQL & lng病人ID & ")"
        
        zlDatabase.ExecuteProcedure strSQL, "Zl_病人照片_Delete"
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SaveCertPicture(ByVal lng证件ID As Long, ByVal lng序号 As Long, ByVal strNote As String, ByVal strFile As String)
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    If strFile = "" Then
        Exit Sub
    End If
    If lng序号 = 0 Then
        strSQL = "Select max(序号) as 序号 from 病人实名证件图片 where 证件ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "获取图片序号", lng证件ID)
        If rsTmp.EOF Then
            lng序号 = 1
        Else
            lng序号 = Val("" & rsTmp!序号) + 1
        End If
    Else
        lng序号 = lng序号
    End If
    If Sys.SaveLob(glngSys, 33, lng证件ID & "|" & lng序号 & "|" & strNote, strFile) = False Then
        MsgBox "保存照片失败,文件可能被删除!", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Function GetCertifyData(ByVal intTYPE As Integer, ByVal intModel As Integer, ByVal lng病人ID As Long, ByVal lng实名ID As Long, Optional ByVal bln新病人 As Boolean, Optional ByVal lng主页ID As Long, Optional ByVal blnIn As Boolean) As String
'功能：获取病人实名信息的sql
    Dim strValue As String
    Dim arrFilds As Variant
    Dim strSQL As String, strTmp As String
    Dim i As Long
    Dim CurrDate As Date
    
    On Error GoTo errH
    If intTYPE = 0 Then
        If intModel = 0 Then
            arrFilds = Array("实名id", "病人id", "新病人", "姓名", "性别", "年龄", "出生日期", "国籍", "民族", "身份证类型", "陪诊人姓名", "陪诊人性别", "陪诊人出生日期", "陪诊人国籍", "陪诊人民族", _
                                "陪诊人身份证类型", "出生地点", "住址", "陪诊人住址", "身份证号", "陪诊人身份证号", "陪诊人关系", "手机号", "备注", _
                                "认证状态", "所有者", "证件信息", "是否结构化", "地址信息", "主页id", "是否就诊")
            strSQL = "ZL_病人实名信息_Insert_S("
        Else
            arrFilds = Array("实名id", "病人id", "姓名", "性别", "年龄", "出生日期", "国籍", "民族", "身份证类型", "陪诊人姓名", "陪诊人性别", "陪诊人出生日期", "陪诊人国籍", "陪诊人民族", _
                                "陪诊人身份证类型", "出生地点", "住址", "陪诊人住址", "身份证号", "陪诊人身份证号", "陪诊人关系", "手机号", "备注", _
                                "认证状态", "所有者", "变更原因", "证件信息", "是否结构化", "地址信息", "主页id", "是否就诊")
            strSQL = "ZL_病人实名信息_Update_S("
        End If
        For i = LBound(arrFilds) To UBound(arrFilds)
            strValue = ""
            Select Case Trim(arrFilds(i))
                Case ""
                    strValue = ",Null"
                Case "出生日期", "陪诊人出生日期"
                    mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!信息现值 & ""
                    strValue = "," & zlStr.To_Date(strValue, "ymdhm")
                Case "证件信息"
                     strValue = ",'" & CachCertData & "'"
                Case "认证状态"
                    strValue = IIf(mblnIdentifySure, ",1", ",0")
                Case "所有者"
                    strValue = IIf(optType(0).Value, ",1", ",2")
                Case "是否结构化"
                    strValue = IIf(gbln启用结构化地址, ",1", ",0")
                Case "地址信息"
                    strValue = ",'" & GetPatiAdresInfo & "'"
                Case "身份证类型", "陪诊人身份证类型"
                    mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!信息现值 & ""
                    If strValue <> "" Then
                        strValue = "," & cbo.FindIndex(cboInfo(mrsMainInfo!Index), strValue)
                    Else
                        strValue = ",Null"
                    End If
                Case "实名id"
                    strValue = "," & lng实名ID
                Case "病人id"
                    strValue = "," & lng病人ID
                Case "主页id"
                    If lng主页ID <> 0 Then
                        strValue = "," & lng主页ID
                    Else '
                        strValue = ",Null"
                    End If
                Case "新病人"
                    strValue = "," & IIf(bln新病人 = True, 1, 0)
                Case "是否就诊"
                    strValue = "," & IIf(blnIn = True, 1, 0)
                Case "年龄"
                    If IsDate(txtDateInfo(DATE_出生日期).Text) Then
                        mstrAge = GetAge(txtDateInfo(DATE_出生日期).Text, mlng病人ID, CurrDate)
                    End If
                    strValue = ",'" & mstrAge & "'"
                Case "变更原因"
                    strValue = ",'" & mstrReason & "'"
                Case Else
                    mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!信息现值 & ""
                    strValue = IIf(strValue = "", ",Null", ",'" & strValue & "'")
            End Select
            If i = UBound(arrFilds) Then
                strValue = IIf(strValue = "", "Null", strValue) & ")"
            End If
            strTmp = strTmp & strValue
        Next
    ElseIf intTYPE = 1 Then
        strValue = "," & lng实名ID & "," & lng病人ID & ",1" & ")"
        strSQL = "Zl_病人实名信息_状态_Update(0,"
        strTmp = strValue
    ElseIf intTYPE = 2 Then
        strValue = "," & lng实名ID & "," & lng病人ID & ",0" & ")"
        strSQL = "Zl_病人实名信息_状态_Update(0,"
        strTmp = strValue
    ElseIf intTYPE = 3 Then
        arrFilds = Array("姓名", "性别", "年龄", "出生日期", "陪诊人姓名", "陪诊人性别", "陪诊人出生日期", "身份证号", "陪诊人身份证号", "陪诊人关系", "所有者", "证件信息")
        For i = LBound(arrFilds) To UBound(arrFilds)
            strValue = ""
            Select Case Trim(arrFilds(i))
                Case "出生日期", "陪诊人出生日期"
                    mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!信息现值 & ""
                    strValue = "|" & Format(strValue, "yyyy-mm-dd hh:mm")
                Case "证件信息"
                     strValue = "|" & CachCertData
                Case "所有者"
                    strValue = IIf(optType(0).Value, "|1", "|2")
                Case "年龄"
                    If IsDate(txtDateInfo(DATE_出生日期).Text) Then
                        mstrAge = GetAge(txtDateInfo(DATE_出生日期).Text, mlng病人ID, CurrDate)
                    End If
                    strValue = "|" & mstrAge
                Case Else
                    mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!信息现值 & ""
                    strValue = IIf(strValue = "", "|", "|" & strValue)
            End Select
            strTmp = strTmp & strValue
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End If
    If intTYPE <> 3 Then
        strSQL = strSQL & Mid(strTmp, InStr(strTmp, ",") + 1)
        GetCertifyData = strSQL
    Else
        GetCertifyData = strTmp
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetAge(ByVal DateBir As Date, Optional ByVal lng病人ID As Long, Optional ByVal datCalc As Date) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    lng病人ID = 0
    strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID, DateBir, datCalc)
    If Not rsTmp.EOF Then
        GetAge = "" & rsTmp!old
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUpdateData() As String
'功能：获取病人实名信息变动记录的SQL
    Dim strValue As String
    Dim arrFilds As Variant
    Dim strSQL As String, strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    arrFilds = Array("姓名", "性别", "出生日期", "身份证号", "国籍", "民族", "出生地点", "住址", "身份证类型", "陪诊人姓名", "陪诊人性别", "陪诊人出生日期", "陪诊人身份证号", _
                        "陪诊人身份证类型", "陪诊人关系", "陪诊人住址", "陪诊人国籍", "陪诊人民族", "手机号", "备注", _
                        "变更原因")
                        
    strSQL = "Zl_病人实名信息_基本信息变动(" & mlng实名id & "," & mlng病人ID & ","
    
    For i = LBound(arrFilds) To UBound(arrFilds)
        strValue = ""
        Select Case Trim(arrFilds(i))
            Case ""
                strValue = ",Null"
            Case "出生日期", "陪诊人出生日期"
                mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!信息现值 & ""
                strValue = "," & zlStr.To_Date(strValue, "ymdhm")
            Case "身份证类型", "陪诊人身份证类型"
                mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!信息现值 & ""
                strValue = "," & cbo.FindIndex(cboInfo(mrsMainInfo!Index), strValue)
            Case "变更原因"
                strValue = ",'" & mstrReason & "'"
            Case Else
                mrsMainInfo.Filter = "信息名='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!信息现值 & ""
                strValue = IIf(strValue = "", ",Null", ",'" & strValue & "'")
        End Select
        If i = UBound(arrFilds) Then
            strValue = IIf(strValue = "", "Null", strValue) & ")"
        End If
        strTmp = strTmp & strValue
    Next
    strSQL = strSQL & Mid(strTmp, InStr(strTmp, ",") + 1)
    GetUpdateData = strSQL
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetPatiAdresInfo() As String
'功能：获取所有结构化地址信息的字符串
    Dim strTmp As String
    Dim i As Long
    Dim intTYPE As Integer
    
    On Error GoTo errH
    For i = patiAdressInfo.LBound To patiAdressInfo.UBound
        If patiAdressInfo(i).Value <> "" Then
            '新增\修改
            intTYPE = decode(i, 0, 1, 1, 3, 2, 5)
            strTmp = strTmp & "|" & intTYPE & ";" & patiAdressInfo(i).value省 & ";" & patiAdressInfo(i).value市 & ";" & patiAdressInfo(i).value区县 & ";" & patiAdressInfo(i).value乡镇 & ";" & patiAdressInfo(i).value详细地址 & ";" & patiAdressInfo(i).Code
        End If
    Next
    strTmp = Mid(strTmp, 2)
    GetPatiAdresInfo = strTmp
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function GetCertifyValue() As String
'功能：获取证件信息
    Dim i As Long
    Dim strCert As String

    On Error GoTo errH
    With vsfCert
        For i = .FixedRows To .Rows - 1
            If (.TextMatrix(i, COL_证件号码)) <> "" Then
                strCert = strCert & "," & zlCommFun.GetNeedName(.TextMatrix(i, COL_证件类型), "-") & "-" & .TextMatrix(i, COL_证件号码) & "-" & .TextMatrix(i, COL_备注) & "-" & .Cell(flexcpData, i, COL_所有者, i, COL_所有者) & "-" & .TextMatrix(i, COL_备注)
            End If
        Next
    End With
    strCert = Mid(strCert, 2)
    GetCertifyValue = strCert
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function CheckCertifyData() As Boolean
'功能：保存前的检查
    Dim objCtl As Object, objTmp As Object
    Dim strBirthDay As String, strAge As String, strSex As String, strErrInfo As String, strBaseInfo As String
    Dim strTmp As String, str证件类型 As String, str国籍 As String, strInfo As String
    Dim strMask As String
    Dim blnJudge As Boolean, blnShow As Boolean
    Dim i As Long, j As Long
    Dim strBirthdate As String
    Dim CurrDate As Date
    
    On Error GoTo errH
    CurrDate = zlDatabase.Currentdate
    For Each objCtl In Me.Controls
        Select Case objCtl.Name
            Case "cboInfo"
                strTmp = decode(objCtl.Index, CBO_性别, "病人性别", CBO_陪诊人性别, "陪诊人性别", CBO_国籍, "病人国籍", CBO_陪诊人国籍, "陪诊人国籍", CBO_民族, "病人民族", CBO_陪诊人民族, "陪诊人民族", CBO_身份证类型, "病人身份证类型", CBO_陪诊人身份证类型, "陪诊人身份证类型", CBO_关系, "陪诊人关系")
                If objCtl.Index = CBO_性别 Or objCtl.Index = CBO_国籍 Or objCtl.Index = CBO_民族 Then
                    blnShow = True
                ElseIf objCtl.Index = CBO_陪诊人性别 Or objCtl.Index = CBO_陪诊人国籍 Or objCtl.Index = CBO_陪诊人民族 Or objCtl.Index = CBO_关系 Then
                    If txtInfo(TXT_陪诊人姓名).Text <> "" Then
                        blnShow = True
                    End If
                ElseIf objCtl.Index = CBO_身份证类型 Then
                    If txtInfo(TXT_身份证号).Text <> "" Then
                        blnShow = True
                    End If
                    If zlCommFun.GetNeedName(cboInfo(CBO_身份证类型).Text, "-") = "外国人居留证" Then
                        If zlCommFun.GetNeedName(cboInfo(CBO_国籍), "-") = "中国" Then
                            ShowMessage objCtl, "病人的证件类型为【外国人居留证】对应的国籍不能为【中国】，请检查！", False
                            Exit Function
                        End If
                    End If
                ElseIf objCtl.Index = CBO_陪诊人身份证类型 Then
                    If txtInfo(TXT_陪诊人身份证号).Text <> "" Then
                        blnShow = True
                    End If
                    If zlCommFun.GetNeedName(cboInfo(CBO_陪诊人身份证类型).Text, "-") = "外国人居留证" Then
                        If zlCommFun.GetNeedName(cboInfo(CBO_陪诊人国籍), "-") = "中国" Then
                            ShowMessage objCtl, "病人的证件类型为【外国人居留证】对应的国籍不能为【中国】，请检查！", False
                            Exit Function
                        End If
                    End If
                End If
                If blnShow Then
                    If Trim(objCtl.Text) = "" Then
                        ShowMessage objCtl, strTmp & "必须填写！", False
                        Exit Function
                    End If
                End If
                blnShow = False
            Case "txtInfo"
                Select Case objCtl.Index
                    Case TXT_陪诊人身份证号, TXT_身份证号
                        If Trim(objCtl.Text) <> "" Then
                            str证件类型 = IIf(objCtl.Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_身份证类型).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人身份证类型).Text, "-"))
                            str国籍 = IIf(objCtl.Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_国籍).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人国籍).Text, "-"))
                            If (str证件类型 = "居民身份证" Or str证件类型 = "港澳台居住证") And str国籍 = "中国" Then
                                If CreatePublicPatient() Then
                                    If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(objCtl.Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                                        If strBirthDay <> Format(IIf(objCtl.Index = TXT_陪诊人身份证号, txtDateInfo(DATE_陪诊人出生日期).Text, txtDateInfo(DATE_出生日期).Text), "YYYY-MM-DD") Then
                                            strBaseInfo = strBaseInfo & "," & "出生日期"
                                        End If
                                        If strSex <> zlCommFun.GetNeedName(IIf(objCtl.Index = TXT_陪诊人身份证号, cboInfo(CBO_陪诊人性别).Text, cboInfo(CBO_性别).Text), "-") Then
                                            strBaseInfo = strBaseInfo & "," & "性别"
                                        End If
                                        If Format(strBirthDay, "HH:MM") = "00:00" Then
                                            strMask = "####-##-##"
                                        Else
                                            strMask = "####-##-## ##:##"
                                        End If
                                        strBaseInfo = Mid(strBaseInfo, 2)
                                        If strBaseInfo <> "" Then
                                            If objCtl.Index = TXT_身份证号 Then
                                                If MsgBox("病人身份证号返回的" & strBaseInfo & "与实际填写的" & strBaseInfo & "不符合,是否继续？继续则会将界面上录入的" & strBaseInfo & "替换成身份证号返回的" & strBaseInfo & "！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    txtDateInfo(DATE_出生日期).Mask = strMask
                                                    txtDateInfo(DATE_出生日期).Tag = strMask
                                                    txtDateInfo(DATE_出生日期).Text = Format(strBirthDay, decode(strMask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                                                    cboInfo(CBO_性别).ListIndex = cbo.FindIndex(cboInfo(CBO_性别), strSex)
                                                End If
                                            Else
                                                If MsgBox("陪诊人身份证号返回的" & strBaseInfo & "与实际填写的" & strBaseInfo & "不符合,是否替换？继续则会将界面上录入的" & strBaseInfo & "替换成身份证号返回的" & strBaseInfo & "！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    txtDateInfo(DATE_陪诊人出生日期).Mask = strMask
                                                    txtDateInfo(DATE_陪诊人出生日期).Tag = strMask
                                                    txtDateInfo(DATE_陪诊人出生日期).Text = Format(strBirthDay, decode(strMask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                                                    cboInfo(CBO_陪诊人性别).ListIndex = cbo.FindIndex(cboInfo(CBO_陪诊人性别), strSex)
                                                End If
                                            End If
                                        End If
                                        If objCtl.Index = TXT_身份证号 Then
                                            If IsDate(txtDateInfo(DATE_出生日期).Text) Then
                                                mstrAge = GetAge(txtDateInfo(DATE_出生日期).Text, mlng病人ID, CurrDate)
                                            End If
                                        End If
                                        strBirthDay = "": strAge = "": strSex = "": strErrInfo = "": strBaseInfo = ""
                                    Else
                                        Call ShowMessage(txtInfo(objCtl.Index), strErrInfo)
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Case txt_手机号
                        If Not CheckPhoneNumber(Trim(objCtl.Text)) Then Exit Function
                    Case TXT_姓名
                        If Trim(objCtl.Text) = "" Then
                            ShowMessage objCtl, "必须录入病人姓名！", False
                            Exit Function
                        End If
                    Case TXT_陪诊人姓名
                        If Trim(objCtl.Text) = "" Then
                            If CheckPPatiInfo Then
                                ShowMessage objCtl, "陪诊人姓名还没有录入，请检查！", False
                                Exit Function
                            End If
                        End If
                End Select
            Case "txtAdressInfo"
                strTmp = decode(objCtl.Index, ADRS_出生地点, "病人出生地点", ADRS_住址, "病人住址", ADRS_陪诊人住址, "陪诊人住址")
                If gbln启用结构化地址 Then    '需要检查地址控件的内容
                    If patiAdressInfo(objCtl.Index).CheckNullValue() <> "" Then
                        Call ShowMessage(patiAdressInfo(objCtl.Index), strTmp & "的" & patiAdressInfo(objCtl.Index).CheckNullValue() & "尚未输入，请检查。", False)
                        Exit Function
'                    ElseIf patiAdressInfo(objCtl.Index).Value = "" Then
'                        Call ShowMessage(patiAdressInfo(objCtl.Index), "必须录入" & strTmp & "。", False)
'                        Exit Function
                    End If
                    If patiAdressInfo(objCtl.Index).MaxLength > 0 Then
                        If zlCommFun.ActualLen(patiAdressInfo(objCtl.Index).Value) > patiAdressInfo(objCtl.Index).MaxLength Then
                            Call ShowMessage(patiAdressInfo(objCtl.Index), strTmp & "的内容太长，请检查。(该项目最多允许 " & patiAdressInfo(objCtl.Index).MaxLength & " 个字符或 " & patiAdressInfo(objCtl.Index).MaxLength \ 2 & " 个汉字)", False)
                            Exit Function
                        End If
                    End If
                Else '需要检查TextBox的内容
                    If objCtl.MaxLength <> 0 And objCtl.Text <> "" Then
                        If zlCommFun.ActualLen(objCtl.Text) > objCtl.MaxLength Then
                            Call ShowMessage(objCtl, strTmp & "的内容过长，请检查。(该项目最多允许 " & objCtl.MaxLength & " 个字符或 " & objCtl.MaxLength \ 2 & " 个汉字)", False)
                            Exit Function
                        End If
                    End If
                End If
            Case "txtDateInfo"
                strTmp = decode(objCtl.Index, DATE_出生日期, "出生日期", DATE_陪诊人出生日期, "陪诊人出生日期")
                If objCtl.Mask = "####-##-## ##:##" Then
                    If Format(Mid(objCtl.Text, 12), "HH:MM") = "__:__" Then
                        strBirthdate = Format(Mid(objCtl.Text, 1, 10), "YYYY-MM-DD")
                    Else
                        strBirthdate = Format(objCtl.Text, "YYYY-MM-DD HH:MM")
                    End If
                ElseIf objCtl.Mask = "####-##-##" Then
                    strBirthdate = Format(objCtl.Text, "YYYY-MM-DD")
                End If
                If Not IsDate(strBirthdate) Then
                    If objCtl.Index = DATE_出生日期 Then
                        If objCtl.Text = "____-__-__" Or objCtl.Text = "____-__-__ __:__" And txtInfo(TXT_姓名).Text <> "" Then
                            Call ShowMessage(txtDateInfo(objCtl.Index), "请输入病人的" & strTmp & "!", False)
                            Exit Function
                        Else
                            If txtInfo(TXT_姓名).Text <> "" Then
                                Call ShowMessage(txtDateInfo(objCtl.Index), strTmp & "不是有效的日期格式。", False)
                                Exit Function
                            End If
                        End If
                    ElseIf objCtl.Index = DATE_陪诊人出生日期 Then
                        If (objCtl.Text = "____-__-__" Or objCtl.Text = "____-__-__ __:__") And Trim(txtInfo(TXT_陪诊人姓名).Text) <> "" Then
                            Call ShowMessage(txtDateInfo(objCtl.Index), "请输入" & strTmp & "!", False)
                            Exit Function
                        Else
                            If txtInfo(TXT_陪诊人姓名).Text <> "" Then
                                Call ShowMessage(txtDateInfo(objCtl.Index), strTmp & "不是有效的日期格式。", False)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Case "vsfCert"
                With vsfCert
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, COL_证件号码) <> "" Then
                            If .TextMatrix(i, COL_证件类型) = "" Then
                                Call ShowMessage(vsfCert, "请选择证件类型！", False)
                                Exit Function
                            End If
                        End If
                        strInfo = strInfo & "," & .TextMatrix(i, COL_证件号码)
                        For j = i + 1 To .Rows - 1
                            If .TextMatrix(j, COL_证件号码) <> "" Then
                                If .TextMatrix(j, COL_证件号码) = .TextMatrix(i, COL_证件号码) And zlCommFun.GetNeedName(.TextMatrix(j, COL_证件类型), "-") = zlCommFun.GetNeedName(.TextMatrix(i, COL_证件类型), "-") Then
                                    .Row = i: .Col = COL_证件号码
                                    Call ShowMessage(vsfCert, "有重复的证件信息，请检查！", False)
                                    Exit Function
                                End If
                            End If
                        Next
                    Next
                End With
                If Mid(strInfo, 2) = "" Then
                    If txtInfo(TXT_身份证号).Text = "" And txtInfo(TXT_陪诊人身份证号) = "" Then
                        Call ShowMessage(txtInfo(TXT_身份证号), "病人身份证、病人其他证件、陪诊人身份证、陪诊人其他证件必须录入一个，请检查！", False)
                        Exit Function
                    End If
                End If
        End Select
    Next
    CheckCertifyData = True
    Exit Function
errH:    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub cmbMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Certify_Save
            Control.Visible = True
            Control.Enabled = IIf(mblnInfoChange, True, False)
        Case conMenu_Certify_IdentifySure
            Control.Visible = True
            If mblnSave Then
                If mblnIdentifySure Then
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Else
                Control.Enabled = False
            End If
        Case conMenu_Certify_Cancel
            Control.Visible = True
            Control.Enabled = IIf(mblnIdentifySure, True, False)
    End Select
End Sub

Private Sub cmdAdress_Click(Index As Integer)
'功能：cmdAdressInfo_Click
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte, strCaption As String, strMsg As String, blnRoot As Boolean, blnNonWin As Boolean

    On Error GoTo errH
    Select Case Index
        Case ADRS_出生地点, ADRS_住址, ADRS_陪诊人住址
            '选择地区数据
            strSQL = "Select Rownum as ID,编码,名称,简码 From 地区 Order by 编码"
            strCaption = "区域": strMsg = "字典管理工具": bytStyle = 0: blnRoot = False: blnNonWin = True
    End Select

    '数据处理
    On Error GoTo errH
    '数据处理
    Set objTXTBox = txtAdressInfo(Index)
    vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, bytStyle, strCaption, , , , , blnRoot, blnNonWin, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "没有设置""" & IIf(strCaption = "区域", "地区", strCaption) & """数据，请先到" & strMsg & "中设置。", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        objTXTBox.Text = rsTmp!名称
        objTXTBox.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAdress_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmdInfoDate_Click(Index As Integer)
    Dim objmonInfo As MonthView  '方便调用控件属性
    Dim objCmd As CommandButton
    Dim objMSK As MaskEdBox
    Dim datStart As Date
    Dim DateEnd As Date
    Dim datTmp As Date
    
    On Error GoTo errH
    mintDate = Index
    Set objmonInfo = monInfo
    Set objCmd = cmdInfoDate(Index)
    Set objMSK = txtDateInfo(Index)
    datStart = zlDatabase.Currentdate
    objmonInfo.MinDate = 0
    objmonInfo.MaxDate = zlDatabase.Currentdate
    Select Case Index
        Case DATE_出生日期, DATE_陪诊人出生日期
            objmonInfo.MaxDate = datStart
    End Select
    If IsDate(objMSK.Text) Then
        datTmp = CDate(objMSK.Text)
        If datTmp > objmonInfo.MaxDate Then
            datTmp = objmonInfo.MaxDate
        ElseIf datTmp < objmonInfo.MinDate Then
            datTmp = objmonInfo.MinDate
        End If
        objmonInfo.Value = datTmp
    End If
    objmonInfo.Left = objCmd.Left + objCmd.Width - objmonInfo.Width + objMSK.Container.Left + picPati.Left
    If objCmd.Index = DATE_出生日期 Then
        objmonInfo.Top = objCmd.Top + objCmd.Height + 20 + objMSK.Container.Top
    Else
        objmonInfo.Top = objCmd.Top - objmonInfo.Height - 20 + objMSK.Container.Top
    End If
    objmonInfo.ZOrder
    objmonInfo.Visible = True
    objmonInfo.SetFocus
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
'鼠标滚轮
    Call Form_Resize
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
'鼠标滚轮
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hwnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsbMain.Value
    lngMin = vsbMain.Min
    lngMax = vsbMain.Max
    
    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsbMain.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsbMain.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsbMain.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsbMain.Value = lngMin
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim strFile As String
    Dim objFile As New FileSystemObject
    
    'CommandBar
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cmbMain.VisualTheme = xtpThemeOffice2003
    With Me.cmbMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
'        .UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Set cmbMain.Icons = imgManager.Icons
    cmbMain.EnableCustomization False
    
    Call InitBaseInfo
    
    '初始化地址控件
    patiAdressInfo(ADRS_出生地点).Visible = gbln启用结构化地址
    patiAdressInfo(ADRS_住址).Visible = gbln启用结构化地址
    patiAdressInfo(ADRS_陪诊人住址).Visible = gbln启用结构化地址
    txtAdressInfo(ADRS_出生地点).Visible = Not gbln启用结构化地址: cmdAdress(ADRS_出生地点).Visible = Not gbln启用结构化地址
    txtAdressInfo(ADRS_住址).Visible = Not gbln启用结构化地址: cmdAdress(ADRS_住址).Visible = Not gbln启用结构化地址
    txtAdressInfo(ADRS_陪诊人住址).Visible = Not gbln启用结构化地址: cmdAdress(ADRS_陪诊人住址).Visible = Not gbln启用结构化地址
    If gbln启用结构化地址 Then
        patiAdressInfo(ADRS_出生地点).ShowTown = gbln显示乡镇
        patiAdressInfo(ADRS_住址).ShowTown = gbln显示乡镇
        patiAdressInfo(ADRS_陪诊人住址).ShowTown = gbln显示乡镇
    End If
    
    '画线
    Call DrawLin
    
    '添加工具栏
    Call MainDefCommandBar
    
    '初始化列表
    Call InitVsfGridHeader
    
    '初始化数据
    Call InitCboData
    
    mbln扫描身份证登记 = Val(zlDatabase.GetPara("扫描身份证登记", glngSys, glngModul)) = "1"
    If mintModel = 1 Then
        mblnSave = True
        Screen.MousePointer = 11
        Call LoadPatiInfo(mlng实名id)
        Call LoadPatiPricture(mlng病人ID, imgLoad, strFile)
        If strFile <> "" Then
            mstr采集图片 = App.Path & "/pati"
        End If
        If imgLoad.Picture <> 0 Then
            imgPatient.Picture = imgLoad.Picture
        End If
        Call LoadInterface(mlng实名id)
        Screen.MousePointer = 0
    Else
        mblnSave = False
        mblnIdentifySure = False
        Screen.MousePointer = 11
        Call LoadPatiInfo(mlng实名id)
        Call LoadInterface
        Screen.MousePointer = 0
    End If
    mblnLoadFilish = True
    If mlng实名id <> 0 Then
        stbBar.Panels(2).Text = "实名id:" & mlng实名id
    End If
    
'    vsbMain.Max = 600
'    vsbMain.Min = 0
'    vsbMain.LargeChange = 100
    
    If Not objFile.FolderExists(App.Path & "\CertImg") Then
        objFile.CreateFolder App.Path & "\CertImg"
    End If
End Sub

Private Function LoadPatiInfo(ByVal lng实名ID As Long) As Boolean
    Dim i As Long
    On Error GoTo errH
    Set mrsPati = LoadPatiInfoByID(lng实名ID)
    If Not mrsPati.EOF Then
        For i = 0 To mrsPati.Fields.Count - 1
            LoadCache mrsPati.Fields(i).Name, mrsPati.Fields(i).Value & ""
        Next
        mblnIdentifySure = IIf(Val(mrsPati!认证状态 & "") = 0, False, True)
    End If
    Set mrsCert = LoadPatiCert(0, lng实名ID)
    If Not mrsCert.EOF Then
        LoadCachCert mrsCert
    Else
        vsfCert.TextMatrix(vsfCert.FixedRows, COL_所有者) = IIf(optType(0).Value, "病人本身", "陪诊人")
        vsfCert.Cell(flexcpData, vsfCert.FixedRows, COL_所有者, vsfCert.FixedRows, COL_所有者) = IIf(optType(0).Value, 1, 2)
    End If
    vsfCert.Select vsfCert.FixedRows, COL_证件号码
    Call vsfCert_Click
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadCachCert(ByVal rsTmp As ADODB.Recordset)
'功能：将证件信息加载并缓存

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim strsInfo As String, strsMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim lngTmp As Long
    Dim lngsTmp As Long
    Dim strType As String
    Dim rsImg As New ADODB.Recordset
    Dim strFile As String
    Dim objFile As New FileSystemObject
    
    On Error GoTo errH
    
     '删除之前的缓存
    mrsSecdInfo.Filter = "控件名='vsfCert'"
    If Not mrsSecdInfo.EOF Then
        For i = 1 To mrsSecdInfo.RecordCount
            mrsSecdInfo.Delete
            mrsSecdInfo.Update
            mrsSecdInfo.MoveNext
        Next
    End If

    lngTmp = 1
    lngsTmp = 1
    With vsfCert
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .Rows = .Rows + 1: lngRow = .Rows - 1
            .TextMatrix(lngRow, COL_证件ID) = "" & rsTmp!ID
            .TextMatrix(lngRow, COL_证件类型) = "" & rsTmp!证件类型
            .TextMatrix(lngRow, COL_证件号码) = "" & rsTmp!证件号码
            .TextMatrix(lngRow, COL_备注) = "" & rsTmp!备注
            .TextMatrix(lngRow, COL_所有者) = IIf(Val("" & rsTmp!所有者) = 1, "病人本身", "陪诊人")
            .Cell(flexcpData, lngRow, COL_所有者, lngRow, COL_所有者) = Val("" & rsTmp!所有者)
            
            .Cell(flexcpPicture, lngRow, COL_图片, lngRow, COL_图片) = img图片
            .Cell(flexcpPictureAlignment, lngRow, COL_图片, lngRow, COL_图片) = 4
            
            .Cell(flexcpPicture, lngRow, COL_增加, lngRow, COL_增加) = imgAdd
            .Cell(flexcpPictureAlignment, lngRow, COL_增加, lngRow, COL_增加) = 4
            
            .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
            .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
            
            .RowData(lngRow) = Val(rsTmp!ID & "")
            If Trim(.TextMatrix(lngRow, COL_证件ID)) <> "" Then
                Set rsImg = GetCertPicture(Val(.TextMatrix(lngRow, COL_证件ID)), 0, 1)
                If Not rsImg.EOF Then
                    With vsfImg
                        j = .Rows - 1
                        For k = 0 To rsImg.RecordCount - 1
                            .AddItem "" & lngRow & "-" & j, j
                            .TextMatrix(j, IMG_证件ID) = "" & i
                            .Cell(flexcpData, j, IMG_证件ID, j, IMG_证件ID) = "" & lngRow & "-" & j
                            
                            .TextMatrix(j, IMG_序号) = "" & rsImg!序号
                            .RowData(j) = Val(vsfCert.TextMatrix(lngRow, COL_证件ID))
                            
                            .TextMatrix(j, IMG_备注) = "" & rsImg!备注
                        
                            .Cell(flexcpPicture, j, IMG_图片, j, IMG_图片) = ImgCert
                            .Cell(flexcpPictureAlignment, j, IMG_图片, j, IMG_图片) = 4
                            
                            .Cell(flexcpPicture, j, IMG_Del, j, IMG_Del) = imgDelete
                            .Cell(flexcpAlignment, j, IMG_Del, j, IMG_Del) = 4
                            
                            strsMainInfo = .RowData(j) & "|" & .Cell(flexcpData, j, IMG_图片, j, IMG_图片) & "|" & .TextMatrix(j, IMG_序号) & "|" & .TextMatrix(j, IMG_备注)
                            strsInfo = .RowData(j) & "|" & .Cell(flexcpData, j, IMG_图片, j, IMG_图片) & "|" & .TextMatrix(j, IMG_序号) & "|" & .TextMatrix(j, IMG_备注)
                            mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值"), Array(lngsTmp, Val(.RowData(j)), "vsfImg", strsInfo, strsMainInfo)
                            j = j + 1
                            lngsTmp = lngsTmp + 1
                            strFile = ""
                            rsImg.MoveNext
                        Next
                    End With
                End If
            End If
                
            strMainInfo = rsTmp!ID & "|" & rsTmp!证件类型 & "|" & rsTmp!证件号码
            strInfo = strMainInfo & "|" & rsTmp!备注 & "|" & IIf(Nvl("" & rsTmp!所有者, "") = "1", "病人本身", "陪诊人")
            mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值"), Array(lngTmp, Val(rsTmp!ID & ""), "vsfCert", strInfo, strMainInfo)
            lngTmp = lngTmp + 1
            rsTmp.MoveNext
        Next
        .Row = 1: .Col = COL_证件号码
        If .TextMatrix(1, COL_所有者) <> "" Then
            If .TextMatrix(1, COL_所有者) = "病人本身" Then
                optType(0).Value = True
                optType(1).Value = False
            Else
                optType(1).Value = True
                optType(0).Value = False
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCachInterface(ByVal rsTmp As ADODB.Recordset)
'功能：三方接口信息加载并缓存

    Dim strTmp As String
    Dim i As Long, j As Long, k As Long, lngRow As Long, lngCount As Long
    Dim lngPos As Long
    Dim strInfo As String, strMainInfo As String
    Dim arrWhole As Variant, arrMain As Variant
    Dim lngTmp As Long
    
    On Error GoTo errH
    i = vsfInterface.FixedRows
    If Not rsTmp.EOF Then
        With vsfInterface
            Do While Not rsTmp.EOF
                For j = .FixedRows To .Rows - 1
                    If Val("" & rsTmp!ID) = Val(.TextMatrix(j, COLS_接口ID)) Then
                        lngCount = lngCount + 1
                    End If
                Next
                If lngCount = 0 Then
                    .AddItem "", i
                    .TextMatrix(i, COLS_接口ID) = "" & rsTmp!ID
                    .TextMatrix(i, COLS_名称) = "" & rsTmp!接口名
                    .TextMatrix(i, COLS_部件名) = "" & rsTmp!部件名
                    .TextMatrix(i, COLS_说明) = "" & rsTmp!说明
                    .TextMatrix(i, COLS_认证结果) = decode("" & rsTmp!认证结果, "0", "认证失败", "1", "认证成功", "2", "未认证", "" & rsTmp!认证结果)
                    
                    .Cell(flexcpPicture, i, COLS_认证) = imgIdentify
                    .Cell(flexcpPictureAlignment, lngRow, COL_图片, lngRow, COL_图片) = 4
                    
                    If Val("" & rsTmp!认证结果) = 0 Then
                        mblnInterface = False
                    Else
                        mblnInterface = True
                    End If
                    i = i + 1
                End If
                lngCount = 0
                rsTmp.MoveNext
            Loop
            '数据缓存
            lngTmp = 1
            strTmp = ""
            arrMain = Array(COLS_接口ID, COLS_名称, COLS_说明, COLS_认证结果, COLS_认证)
            arrWhole = Array(COLS_名称, COLS_说明, COLS_认证结果, COLS_认证)
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, COLS_认证结果) <> "" Then
                    If strTmp <> .TextMatrix(i, COLS_认证结果) Then
                        j = 1: strTmp = .TextMatrix(i, COLS_认证结果)
                    Else
                        j = j + 1
                    End If
                    strInfo = j: strMainInfo = ""
                    For k = LBound(arrWhole) To UBound(arrWhole)
                        strInfo = strInfo & "|" & .TextMatrix(i, arrWhole(k))
                    Next
                    For k = LBound(arrMain) To UBound(arrMain)
                        strMainInfo = strMainInfo & "|" & .TextMatrix(i, arrMain(k))
                    Next
                    mrsSecdInfo.AddNew Array("序号", "原ID", "控件名", "信息原值", "主信息原值", "Tag", "信息现值", "主信息现值"), Array(lngTmp, Val(.RowData(i)), vsfInterface.Name, strInfo, strMainInfo, decode(Val(.TextMatrix(i, COLS_认证结果)), "认证失败", 0, "认证成功", 1, "未认证"), Null, Null)
                    lngTmp = lngTmp + 1
                End If
            Next
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadInterface(Optional ByVal lng实名ID As Long) As Boolean
'加载三方接口信息
    On Error GoTo errH
    Set mrsIneterface = LoadCertInterface(0, lng实名ID)
    If Not mrsIneterface.EOF Then
        Call LoadCachInterface(mrsIneterface)
    End If
    Set mrsIneterface = LoadCertInterface(0)
    If Not mrsIneterface.EOF Then
        Call LoadCachInterface(mrsIneterface)
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadCache(ByVal strName As String, ByVal strValue As String)
'功能：初始化缓存信息
    Dim objCtl As Object
    Dim intIndex As Integer
    Dim str控件名 As String
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strFMT As String
    
    On Error GoTo errH
    mrsMainInfo.Filter = "信息名='" & strName & "'"
    If Not mrsMainInfo.EOF Then
        str控件名 = mrsMainInfo!控件名 & ""
        With Me.Controls
            Select Case str控件名
                Case "txtInfo"
                    txtInfo(mrsMainInfo!Index).Text = strValue
                Case "cboInfo"
                    If Not IsNumeric(strValue) Then
                        If strValue <> "" Then
                            intIndex = cbo.FindIndex(cboInfo(mrsMainInfo!Index), "" & strValue)
                        Else
                            If cboInfo(mrsMainInfo!Index).Tag <> "" Then
                                intIndex = cboInfo(mrsMainInfo!Index).Tag
                            Else
                                intIndex = 0
                            End If
                        End If
                    Else
                        intIndex = Val(strValue)
                    End If
                    cboInfo(mrsMainInfo!Index).ListIndex = intIndex
                Case "patiAdressInfo"
                    If gbln启用结构化地址 Then
                        If mlng病人ID <> 0 Then
                            Call SetStructAddress(mlng病人ID, 0, patiAdressInfo(mrsMainInfo!Index), decode(mrsMainInfo!Index, ADRS_出生地点, 1, ADRS_住址, 3, ADRS_陪诊人住址, 5))
                        End If
                    End If
                Case "txtAdressInfo"
                    If Not gbln启用结构化地址 Then
                        txtAdressInfo(mrsMainInfo!Index).Text = strValue
                    End If
                Case "txtDateInfo"
                    strFMT = txtDateInfo(mrsMainInfo!Index).Mask
                    If Format(strValue, "HH:MM") = "00:00" Then
                        txtDateInfo(mrsMainInfo!Index).Mask = "####-##-##"
                        txtDateInfo(mrsMainInfo!Index).Tag = "####-##-##"
                        strFMT = txtDateInfo(mrsMainInfo!Index).Mask
                     Else
                        txtDateInfo(mrsMainInfo!Index).Mask = "####-##-## ##:##"
                        txtDateInfo(mrsMainInfo!Index).Tag = "####-##-## ##:##"
                        strFMT = txtDateInfo(mrsMainInfo!Index).Mask
                    End If
                    If IsDate(strValue) Then
                        strValue = Format(strValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                    Else
                        strValue = Replace(strFMT, "#", "_")
                    End If
                    txtDateInfo(mrsMainInfo!Index).Text = strValue
            End Select
            mrsMainInfo.Update "信息原值", strValue
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function CheckValueChange(Optional ByRef objTmp As Object) As Boolean
'功能：检查首页控件的值是否发生变化
    Dim strOlsInfo As String
    Dim strCurInfo As String
    Dim strCboName As String
    Dim cboTmp As ComboBox
    Dim lngIndex As Long
    Dim blnFind As Boolean
    
    If mblnInfoChange Then Exit Function
    If Not mblnLoadFilish Then Exit Function
    On Error GoTo errH
    If objTmp Is Nothing Then
        mblnInfoChange = True
        mblnSave = False
        Exit Function
    End If
    If TypeName(objTmp) = "ComboBox" Then
        Set cboTmp = objTmp
        strCurInfo = cboTmp.Text
        strCboName = cboTmp.Name
        lngIndex = cboTmp.Index
        mblnChange基本 = True
    Else
        mblnInfoChange = True
        mblnSave = False
        If TypeName(objTmp) <> "VSFlexGrid" Then
            mblnChange基本 = True
        End If
        Exit Function
    End If
    
    If strCboName = "cboInfo" Then
        mrsMainInfo.Filter = "控件名='" & strCboName & "'" & "And Index=" & lngIndex
        If Not mrsMainInfo.EOF Then
            strOlsInfo = Nvl(mrsMainInfo!信息原值)
            blnFind = True
        Else
            mrsSecdInfo.Filter = "控件名='" & strCboName & "'" & "And IndexEx=" & lngIndex
            If Not mrsSecdInfo.EOF Then
                strOlsInfo = Nvl(mrsSecdInfo!信息原值)
                blnFind = True
            End If
        End If
        If blnFind Then
            If strCurInfo <> strOlsInfo And blnFind Then
                mblnInfoChange = True
                mblnSave = False
            End If
        Else
            mblnInfoChange = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CreateIdentifyObj(ByVal strDllName As String) As Boolean
    On Error Resume Next
    Set mobjIdentify = CreateObject(strDllName & ".clsIdentityCert")
    If mobjIdentify Is Nothing Then
        MsgBox "创建三方认证接口部件(" & strDllName & ".clsIdentityCert)失败!", vbInformation, gstrSysName
    Else
        Call mobjIdentify.Initialization(gcnOracle)
    End If
    Err.Clear: On Error GoTo 0
    If Not mobjIdentify Is Nothing Then CreateIdentifyObj = True
End Function


Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbBar.Visible Then Bottom = Me.stbBar.Height
End Sub

Private Sub cmbMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim X As Long, Y As Long

    Call Me.cmbMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    With Me.picBig
        .Left = lngLeft
        .Top = lngTop
        .Width = Me.ScaleWidth
        .Height = lngBottom
    End With
    vsbMain.Max = (picBig.ScaleHeight - picMain.Height) / Screen.TwipsPerPixelY + 200
    vsbMain.Min = 0
    vsbMain.SmallChange = 5
    vsbMain.LargeChange = 50
End Sub

Private Sub Form_Resize()
    If Me.ScaleWidth < picMain.Width Then
        hsbMain.Visible = True
        picMain.Left = Me.ScaleLeft + ((Me.ScaleWidth - picMain.Width) * ((hsbMain.Value) / 100))
    Else
        hsbMain.Visible = False
        picMain.Left = Me.ScaleLeft + (Me.ScaleWidth - picMain.Width) / 2
    End If
    Call picMain_Resize
    vsbMain.Move Me.ScaleWidth - vsbMain.Width, Me.ScaleTop + 530, vsbMain.Width, Me.ScaleHeight + Me.ScaleTop - 870
    vsbMain.LargeChange = 100
    vsbMain.SmallChange = vsbMain.LargeChange / 2
    
    hsbMain.Top = Me.ScaleTop + Me.ScaleHeight - hsbMain.Height - 330
    hsbMain.Left = Me.ScaleLeft
    hsbMain.Width = Me.ScaleLeft + Me.ScaleWidth - 255
    hsbMain.LargeChange = 100 / ((picMain.Width) / Me.ScaleWidth)
    hsbMain.SmallChange = 10
    
    Call vsbMain_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objFile As New FileSystemObject
    
    If mblnInfoChange Then
        If MsgBox("该病人的实名信息还没有保存,是否继续退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        Unload Me
    End If
    ClearValue
    If objFile.FileExists(mstr采集图片) Then
        Kill mstr采集图片
    End If
End Sub

Private Sub hsbMain_Change()
    picMain.Left = Me.ScaleLeft + ((Me.ScaleWidth - picMain.Width) * ((hsbMain.Value) / 100))
End Sub

Private Sub lblAdd_Click()
    Dim strPictureFile As String
    Dim objControl As CommandBarControl
    Dim objFile As New FileSystemObject
    
    If gobjPublicPatient Is Nothing Then
        On Error Resume Next
        Call CreatePublicPatient
        Err.Clear: On Error GoTo 0
    End If
    If gobjPublicPatient Is Nothing Then
        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
        Exit Sub
    End If
    Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
    
    If gobjPublicPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
    If strPictureFile <> "" Then
        objFile.CopyFile strPictureFile, App.Path & "\Person.bmp"
        strPictureFile = App.Path & "\Person.bmp"
        Set imgPatient.Picture = LoadPicture(strPictureFile)
        picPicture.Tag = strPictureFile
        mstr采集图片 = strPictureFile
        mlng图像操作 = 2
    End If
    CheckValueChange imgPatient
End Sub

Private Sub lblDelete_Click()
    mlng图像操作 = 3
    imgPatient.Picture = imgDefual.Picture
    CheckValueChange imgPatient
End Sub

Private Sub lblFile_Click()
'问题号:74421
    Dim strFileDir As String
    On Error GoTo Errhand:
    With cmdialog
        .CancelError = False
        .flags = cdlOFNHideReadOnly
        .Filter = "(*.bmp)|*.bmp"
        .FilterIndex = 2
        .ShowOpen
        strFileDir = .filename
        If strFileDir = "" Then Exit Sub
        imgPatient.Picture = LoadPicture(strFileDir)
    End With
    mlng图像操作 = 1
    CheckValueChange imgPatient
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngIndex As Long, lngPatientID As Long
    Dim objCard As zlOneCardComLib.Card
    Dim strErrMsg As String
    Dim str年龄 As String
    Dim strFMT As String
    
    If Me.ActiveControl Is txtInfo(TXT_身份证号) Then
        If txtInfo(TXT_身份证号).Text = "" Then
            txtInfo(TXT_身份证号).Text = strID
            txtInfo(TXT_姓名).Text = "": txtInfo(TXT_姓名).PasswordChar = ""
            txtInfo(TXT_姓名).IMEMode = 0
            txtInfo(TXT_姓名).Text = strName
            Call cbo.Locate(cboInfo(CBO_性别), strSex)
            Call cbo.Locate(cboInfo(CBO_民族), strNation)
            If Format(datBirthDay, "HH:MM") = "00:00" Then
               strFMT = "####-##-##"
            Else
                strFMT = "####-##-## ##:##"
            End If
            txtDateInfo(DATE_出生日期).Mask = strFMT
            txtDateInfo(DATE_出生日期) = Format(datBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
            txtInfo(TXT_身份证号).Text = strID
            Call LoadIDImage
            txtAdressInfo(ADRS_出生地点).Text = strAddress
            If gbln启用结构化地址 Then
                patiAdressInfo(ADRS_出生地点).Value = strAddress
            End If
            Call cbo.Locate(cboInfo(CBO_身份证类型), "居民身份证")
        End If
    ElseIf Me.ActiveControl Is txtInfo(TXT_陪诊人身份证号) Then
        If txtInfo(TXT_陪诊人身份证号).Text = "" Then
            txtInfo(TXT_陪诊人身份证号).Text = strID
            txtInfo(TXT_陪诊人姓名).Text = "": txtInfo(TXT_陪诊人姓名).PasswordChar = ""
            txtInfo(TXT_陪诊人姓名).IMEMode = 0
            txtInfo(TXT_陪诊人姓名).Text = strName
            Call cbo.Locate(cboInfo(CBO_陪诊人性别), strSex)
            Call cbo.Locate(cboInfo(CBO_陪诊人民族), strNation)
            If Format(datBirthDay, "HH:MM") = "00:00" Then
               strFMT = "####-##-##"
            Else
                strFMT = "####-##-## ##:##"
            End If
            txtDateInfo(DATE_陪诊人出生日期).Mask = strFMT
            txtDateInfo(DATE_陪诊人出生日期) = Format(datBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
            txtInfo(TXT_陪诊人身份证号).Text = strID
            txtAdressInfo(ADRS_陪诊人住址).Text = strAddress
            If gbln启用结构化地址 Then
                patiAdressInfo(ADRS_陪诊人住址).Value = strAddress
            End If
            Call cbo.Locate(cboInfo(CBO_陪诊人身份证类型), "居民身份证")
        End If
    End If
End Sub

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载身份证图像
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Screen.MousePointer = 11
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    Screen.MousePointer = 0
    mlng图像操作 = 4
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub monInfo_DateClick(ByVal DateClicked As Date)
'功能：monInfo_DateClick
    Dim strDate As String, strFMT As String
    Dim objMSK As MaskEdBox

    Set objMSK = txtDateInfo(mintDate)
    '获取时分秒数据
    If objMSK.MaxLength >= Len("####-##-## ##:##") Then
        'yyyy-MM-dd HH:mm:ss 格式时间
        If objMSK.MaxLength > Len("####-##-## ##:##") Then
            strFMT = "HH:mm:ss"
        Else
            'yyyy-MM-dd HH:mm 格式时间
            strFMT = "HH:mm"
        End If
        '原时间是时间类型，这取该时间的时分秒数据，否则取当前时间的时分秒
        If IsDate(objMSK.Text) Then
            strDate = " " & Format(objMSK.Text, strFMT)
        Else
            strDate = " " & Format(zlDatabase.Currentdate, strFMT)
        End If
    End If
    '获取时间
    strDate = Format(DateClicked, "yyyy-MM-dd") & strDate
    objMSK.Text = strDate
    txtDateInfo(objMSK.Index).Text = objMSK.Text
    monInfo.Visible = False
    zlControl.ControlSetFocus objMSK
End Sub


Private Sub monInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        monInfo.Visible = False
    End If
End Sub

Private Sub monInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call monInfo_DateClick(monInfo.Value)
    ElseIf KeyAscii = vbKeyEscape Then
        monInfo.Visible = False
    End If
End Sub

Private Sub monInfo_Validate(Cancel As Boolean)
    monInfo.Visible = False
End Sub

Private Sub mskDate_GotFocus(Index As Integer)
'功能：MskDateInfo_GotFocus
    zlCommFun.OpenIme False
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Long
    With vsfCert
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, COL_所有者) = IIf(Index = 0, "病人本身", "陪诊人")
            .Cell(flexcpData, i, COL_所有者, i, COL_所有者) = IIf(Index = 0, 1, 2)
        Next
    End With
    CheckValueChange optType
End Sub

Private Sub optType_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub patiAdressInfo_Change(Index As Integer)
    Call CheckValueChange(patiAdressInfo(Index))
End Sub

Private Sub picBig_Click()
    monInfo.Visible = False
End Sub

Private Sub picBig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsbMain.SetFocus
End Sub

Private Sub picBig_Resize()
    Dim Y As Long, X As Long
    X = picBig.ScaleWidth / 2 + picBig.ScaleLeft
    Y = picMain.ScaleWidth / 2
    picMain.Top = picBig.Top
    picMain.Left = X - Y
End Sub

Private Sub picCert_Click()
    vsbMain.SetFocus
    monInfo.Visible = False
End Sub

Private Sub picInterface_Click()
    vsbMain.SetFocus
    monInfo.Visible = False
End Sub

Private Sub picMain_Resize()
    picPati.Top = picMain.ScaleTop
    picPati.Left = picMain.ScaleLeft
    
    picCert.Top = picPati.Top + picPati.Height
    picCert.Left = picMain.ScaleLeft
    
    picInterface.Top = picCert.Top + picCert.Height
    picInterface.Left = picMain.ScaleLeft
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '菜单定义
    '-----------------------------------------------------
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Save, "保存", -1, False)
    objControl.IconId = 1
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_IdentifySure, "认证确定", -1, False)
    objControl.IconId = 2
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Cancel, "取消认证", -1, False)
    objControl.IconId = 3
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Quit, "退出", -1, False)
    objControl.IconId = 4
    objControl.BeginGroup = True
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_CertifyHelp_Help, "帮助", -1, False)
    objControl.IconId = 5
    
    For Each objControl In cmbMain.ActiveMenuBar.Controls
      If objControl.type = xtpControlButton Then
          objControl.Style = xtpButtonIconAndCaption
      End If
    Next
        
End Sub

Private Sub InitBaseInfo()
    Dim arrMainFileds() As Variant

    '初始化记录集
    '1、主记录结构定义
    Set mrsMainInfo = New ADODB.Recordset
    With mrsMainInfo
        .Fields.Append "序号", adInteger, , adFldKeyColumn              '主键，标识信息
        .Fields.Append "信息名", adVarChar, 100, adFldKeyColumn   '信息名称
        '该记录集仅记录一个信息对应一个控件的情况或多个信息对应一个控件，其他情况不填写
        .Fields.Append "控件名", adVarChar, 100, adFldIsNullable      '展示信息的控件名称
        .Fields.Append "Index", adInteger, , adFldIsNullable                '为空时表示不是控件数组
        .Fields.Append "ExpState", adInteger                                        '信息扩展状态，0-不扩展，1-初始扩展，2-加载扩展
        .Fields.Append "页码", adInteger                                                '信息所在的页码
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable  '信息在首页加载时的值
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable  '信息在首页检查时的值
        .Fields.Append "ErrInfo", adVarChar, 4000, adFldIsNullable  '控件录入信息不合法提示信息，
        .Fields.Append "Edit", adInteger                                                 '0-可编辑,1-不可编辑，只用于展示,2-不可编辑不保存
        .Fields.Append "是否改变", adInteger                                          '信息是否有改变0-未改变，1-改变了
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    '2、次级信息记录集结构定义
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '本记录集的主键
        .Fields.Append "序号", adInteger                                              '标识信息，引用主记录集
        .Fields.Append "控件名", adVarChar, 100                                       '展示信息的控件名称
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable                        '行号或控件数组Index
        .Fields.Append "页码", adInteger                                              '信息所在的页码
        .Fields.Append "原ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息原值", adVarChar, 2000, adFldIsNullable      '信息在加载时的值
        .Fields.Append "主信息原值", adVarChar, 2000, adFldIsNullable    '信息的主要部分，标识一个信息是否被彻底改变，信息在加载时的值
        .Fields.Append "现ID", adBigInt, , adFldIsNullable
        .Fields.Append "信息现值", adVarChar, 2000, adFldIsNullable      '信息在检查时的值
        .Fields.Append "主信息现值", adVarChar, 2000, adFldIsNullable    '信息在检查时的值
        .Fields.Append "改变状态", adInteger                             '信息改变程度0-未改变，1-次级信息改变，2-主信息改变,3-新增,-1，删除
        .Fields.Append "ID", adBigInt, , adFldIsNullable                 '信息行在数据库中的ID,一般表格类控件使用
        .Fields.Append "Tag", adVarChar, 2000                            '存储额外数据
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With mrsMainInfo
        arrMainFileds = Array("信息名", "控件名", "Index")
        '基本信息页
        .AddNew arrMainFileds, Array("国籍", "cboInfo", CBO_国籍)
        .AddNew arrMainFileds, Array("民族", "cboInfo", CBO_民族)
        .AddNew arrMainFileds, Array("性别", "cboInfo", CBO_性别)
        .AddNew arrMainFileds, Array("身份证类型", "cboInfo", CBO_身份证类型)
        .AddNew arrMainFileds, Array("陪诊人国籍", "cboInfo", CBO_陪诊人国籍)
        .AddNew arrMainFileds, Array("陪诊人民族", "cboInfo", CBO_陪诊人民族)
        .AddNew arrMainFileds, Array("陪诊人性别", "cboInfo", CBO_陪诊人性别)
        .AddNew arrMainFileds, Array("陪诊人身份证类型", "cboInfo", CBO_陪诊人身份证类型)
        .AddNew arrMainFileds, Array("陪诊人关系", "cboInfo", CBO_关系)
        
        .AddNew arrMainFileds, Array("姓名", "txtInfo", TXT_姓名)
        .AddNew arrMainFileds, Array("陪诊人姓名", "txtInfo", TXT_陪诊人姓名)
        .AddNew arrMainFileds, Array("身份证号", "txtInfo", TXT_身份证号)
        .AddNew arrMainFileds, Array("陪诊人身份证号", "txtInfo", TXT_陪诊人身份证号)
        .AddNew arrMainFileds, Array("手机号", "txtInfo", txt_手机号)
        .AddNew arrMainFileds, Array("备注", "txtInfo", TXT_备注)
    
        .AddNew arrMainFileds, Array("出生日期", "txtDateInfo", DATE_出生日期)
        .AddNew arrMainFileds, Array("陪诊人出生日期", "txtDateInfo", DATE_陪诊人出生日期)
        
        If gbln启用结构化地址 Then
            .AddNew arrMainFileds, Array("出生地点", "patiAdressInfo", ADRS_出生地点)
            .AddNew arrMainFileds, Array("住址", "patiAdressInfo", ADRS_住址)
            .AddNew arrMainFileds, Array("陪诊人住址", "patiAdressInfo", ADRS_陪诊人住址)
        Else
            .AddNew arrMainFileds, Array("出生地点", "txtAdressInfo", ADRS_出生地点)
            .AddNew arrMainFileds, Array("住址", "txtAdressInfo", ADRS_住址)
            .AddNew arrMainFileds, Array("陪诊人住址", "txtAdressInfo", ADRS_陪诊人住址)
        End If
        
    End With
End Sub

Private Sub InitCboData()
'给下拉列表添加数据
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    strSQL = _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '民族' 表名 From 民族 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '国籍' 表名 From 国籍 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '性别' 表名 From 性别 Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '社会关系' 表名 From 社会关系  Union ALL" & vbNewLine & _
        "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '证件类型' 表名 From 证件类型"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    Call SetCboFromRec(Array("国籍", "民族", "性别", "社会关系"), Array(CBO_国籍, CBO_民族, CBO_性别, CBO_关系))
    Call SetCboFromRec(Array("社会关系"), Array(CBO_关系), " ")
    Call SetCboFromRec(Array("国籍", "民族", "性别"), Array(CBO_陪诊人国籍, CBO_陪诊人民族, CBO_陪诊人性别))

    Call SetCboFromList(Array("", "0-居民身份证", "1-港澳台居住证", "2-外国人居留证"), Array(CBO_身份证类型))
    Call SetCboFromList(Array("", "0-居民身份证", "1-港澳台居住证", "2-外国人居留证"), Array(CBO_陪诊人身份证类型))
    
    If cboInfo(CBO_身份证类型).ListCount > 0 Then
        cboInfo(CBO_身份证类型).ListIndex = 0
    End If
    If cboInfo(CBO_陪诊人身份证类型).ListCount > 0 Then
        cboInfo(CBO_陪诊人身份证类型).ListIndex = 0
    End If
    
End Sub

Private Sub SetCboFromRec(ByVal arrTab As Variant, ByVal arrCboIdx As Variant, Optional ByVal strAddBeginItems As String = "NULL")
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
    Dim arrItem As Variant
    Dim rsTmp As ADODB.Recordset

    For i = 0 To UBound(arrTab)
        Set rsTmp = GetCboData(arrTab(i))
        If Not rsTmp.EOF Then
            Set objCboTmp = cboInfo(arrCboIdx(i))
                objCboTmp.Clear
            If strAddBeginItems <> "NULL" Then
                arrItem = Split(strAddBeginItems, ",")
                For j = LBound(arrItem) To UBound(arrItem)
                    objCboTmp.AddItem arrItem(j)
                Next
            End If
            For j = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!编码) Then
                    objCboTmp.AddItem rsTmp!名称
                Else
                    objCboTmp.AddItem rsTmp!编码 & "-" & rsTmp!名称
                End If
                objCboTmp.ItemData(objCboTmp.NewIndex) = Nvl(rsTmp!ID, 0)
                If Val(rsTmp!缺省 & "") = 1 Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                    objCboTmp.Tag = objCboTmp.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
    Next
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long

    For i = 0 To UBound(arrCboIdx)
        cboInfo(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboInfo(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboInfo(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub

Public Function ShowMe(frmParent As Object, ByVal intModel As Integer, Optional ByRef lng病人ID As Long, Optional ByRef lng实名ID As Long) As Boolean
    mlng病人ID = lng病人ID
    mlng实名id = lng实名ID
    mintModel = intModel
    Set mfrmParent = frmParent
    Me.Show 1, mfrmParent
    lng病人ID = mlng病人ID
    lng实名ID = mlng实名id
    mlng实名id = 0
    mlng病人ID = 0
End Function

Private Sub picPati_Click()
    vsbMain.SetFocus
    monInfo.Visible = False
End Sub

Private Sub txtAdressInfo_Change(Index As Integer)
    CheckValueChange txtAdressInfo(Index)
End Sub

Private Sub txtAdressInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtDateInfo_Change(Index As Integer)
    Dim CurrDate As Date
    
    CurrDate = zlDatabase.Currentdate
    If IsDate(txtDateInfo(Index)) Then
        mstrAge = GetAge(txtDateInfo(Index), mlng病人ID, CurrDate)
    End If
    CheckValueChange txtDateInfo(Index)
End Sub

Private Sub txtDateInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = Chr(13) Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    Dim strBirthDay As String
    Dim strAge As String
    Dim strSex As String
    Dim strErrInfo As String, strFMT As String
    Dim intIndex As Integer
    Dim str证件类型 As String
    Dim str国籍 As String
    Dim CurrDate As Date
    
    CurrDate = zlDatabase.Currentdate
    If (Index = TXT_身份证号 Or Index = TXT_陪诊人身份证号) And Trim(txtInfo(Index).Text) <> "" Then
        str证件类型 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_身份证类型).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人身份证类型).Text, "-"))
        str国籍 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_国籍).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人国籍).Text, "-"))
        If (str证件类型 = "居民身份证" Or str证件类型 = "港澳台居住证") And str国籍 = "中国" Then
            If mblnLoadFilish Then
                If CreatePublicPatient() Then
                    If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                        If IsDate(strBirthDay) Then
                            intIndex = IIf(Index = TXT_身份证号, DATE_出生日期, DATE_陪诊人出生日期)
                            If Format(strBirthDay, "HH:MM") = "00:00" Then
                               txtDateInfo(intIndex).Mask = "####-##-##"
                               txtDateInfo(intIndex).Tag = "####-##-##"
                               strFMT = txtDateInfo(intIndex).Mask
                            Else
                                txtDateInfo(intIndex).Mask = "####-##-## ##:##"
                                txtDateInfo(intIndex).Tag = "####-##-## ##:##"
                                strFMT = txtDateInfo(intIndex).Mask
                            End If
                            strBirthDay = Format(strBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            txtDateInfo(intIndex).Text = strBirthDay
                            If Index = TXT_身份证号 Then
                                Call cbo.Locate(cboInfo(CBO_性别), strSex, False)
                                mstrAge = GetAge(strBirthDay, mlng病人ID, CurrDate)
                            Else
                                Call cbo.Locate(cboInfo(CBO_陪诊人性别), strSex, False)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    CheckValueChange txtInfo(Index)
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    On Error GoTo errH
 
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    Me.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtInfo_GotFocus(Index As Integer)
    If Index = TXT_身份证号 Or Index = TXT_陪诊人身份证号 Then
        zlControl.TxtSelAll txtInfo(Index)
        If mbln扫描身份证登记 = True Then
            Call OpenIDCard(txtInfo(Index).Text = "")
        End If
    End If
End Sub

Private Function ClearValue()
    Dim objFile As New FileSystemObject
    Dim i As Long
    
    With vsfImg
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, IMG_图片) <> "" Then
                If objFile.FileExists(.Cell(flexcpData, i, IMG_图片)) Then
                    Kill .Cell(flexcpData, i, IMG_图片)
                End If
            End If
        Next
    End With
    mlng证件id = 0
    mintModel = 0
    Set mfrmParent = Nothing
    Set mrsPati = Nothing
    Set mrsCert = Nothing
    Set mrsIneterface = Nothing
    mblnChange = False
    mbln扫描身份证登记 = False
    mblnInfoChange = False
    mblnSave = False
    mblnIdentifySure = False
    Set mrsMainInfo = Nothing
    Set mrsSecdInfo = Nothing
    mblnLoadFilish = False
    Set mobjIdentify = Nothing
    mblnInterface = False
    mintDate = 0
    mstrReason = ""
    mlng图像操作 = 0
    mstr采集图片 = ""
    mlngImage = 0
    mlngPati = 0
    mlngTopVsc = 0
    mblnChange基本 = False
    mstrAge = ""
    mstrMsg = ""
    Set mobjIDCard = Nothing
End Function

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strBirthDay As String
    Dim strAge As String
    Dim strSex As String
    Dim strErrInfo As String, strFMT As String
    Dim intIndex As Long
    Dim str证件类型 As String
    Dim str国籍 As String
    
    If Index = TXT_身份证号 Or Index = TXT_陪诊人身份证号 Then
        str证件类型 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_身份证类型).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人身份证类型).Text, "-"))
        str国籍 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_国籍).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人国籍).Text, "-"))
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            If Trim(txtInfo(Index).Text) <> "" And (str证件类型 = "居民身份证" Or str证件类型 = "港澳台居住证") And str国籍 = "中国" Then
                If Not CreatePublicPatient Then Exit Sub
                If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                    If IsDate(strBirthDay) Then
                        intIndex = IIf(Index = TXT_身份证号, DATE_出生日期, DATE_陪诊人出生日期)
                        If Format(strBirthDay, "HH:MM") = "00:00" Then
                           txtDateInfo(intIndex).Mask = "####-##-##"
                           txtDateInfo(intIndex).Tag = "####-##-##"
                           strFMT = txtDateInfo(intIndex).Mask
                        Else
                            txtDateInfo(intIndex).Mask = "####-##-## ##:##"
                            txtDateInfo(intIndex).Tag = "####-##-## ##:##"
                            strFMT = txtDateInfo(intIndex).Mask
                        End If
                        strBirthDay = Format(strBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                        txtDateInfo(intIndex).Text = strBirthDay
                        If Index = TXT_身份证号 Then
                            Call cbo.Locate(cboInfo(CBO_性别), strSex, False)
                        Else
                            Call cbo.Locate(cboInfo(CBO_陪诊人性别), strSex, False)
                        End If
                    End If
                Else
                    Call ShowMessage(txtInfo(Index), strErrInfo)
                    Exit Sub
                End If
            End If
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    ElseIf Index = TXT_姓名 Or Index = TXT_陪诊人姓名 Then
        If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    ElseIf Index = TXT_备注 Then
        If zlCommFun.ActualLen(txtInfo(TXT_备注)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
            KeyAscii = 0
        End If
    ElseIf Index = txt_手机号 Then
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        If KeyAscii = 13 And Trim(txtInfo(txt_手机号).Text) <> "" Then
            If Not CheckPhoneNumber(Trim(txtInfo(txt_手机号).Text)) Then Exit Sub
        End If
    End If
    If Chr(KeyAscii) = Chr(13) Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_LostFocus(Index As Integer)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled False
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim strBirthDay As String
    Dim strAge As String
    Dim strSex As String
    Dim strErrInfo As String, strFMT As String
    Dim intIndex As Long
    Dim str证件类型 As String
    Dim str国籍 As String
    
    If Index = TXT_身份证号 Or Index = TXT_陪诊人身份证号 And Trim(txtInfo(Index).Text) <> "" Then
        str证件类型 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_身份证类型).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人身份证类型).Text, "-"))
        str国籍 = IIf(Index = TXT_身份证号, zlCommFun.GetNeedName(cboInfo(CBO_国籍).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_陪诊人国籍).Text, "-"))
        If Trim(txtInfo(Index).Text) <> "" And (str证件类型 = "居民身份证" Or str证件类型 = "港澳台居住证") And str国籍 = "中国" Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                    If IsDate(strBirthDay) Then
                        intIndex = IIf(Index = TXT_身份证号, DATE_出生日期, DATE_陪诊人出生日期)
                            If Format(strBirthDay, "HH:MM") = "00:00" Then
                               txtDateInfo(intIndex).Mask = "####-##-##"
                               txtDateInfo(intIndex).Tag = "####-##-##"
                               strFMT = txtDateInfo(intIndex).Mask
                            Else
                                txtDateInfo(intIndex).Mask = "####-##-## ##:##"
                                txtDateInfo(intIndex).Tag = "####-##-## ##:##"
                                strFMT = txtDateInfo(intIndex).Mask
                            End If
                            strBirthDay = Format(strBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                            txtDateInfo(intIndex).Text = strBirthDay
                            If Index = TXT_身份证号 Then
                                Call cbo.Locate(cboInfo(CBO_性别), strSex, False)
                            Else
                                Call cbo.Locate(cboInfo(CBO_陪诊人性别), strSex, False)
                            End If
                    End If
                Else
                    Call ShowMessage(txtInfo(Index), strErrInfo)
                End If
            End If
        End If
    ElseIf Index = txt_手机号 Then
        If Trim(txtInfo(txt_手机号).Text) <> "" Then
            If Not CheckPhoneNumber(Trim(txtInfo(txt_手机号).Text)) Then Exit Sub
        End If
    End If
End Sub

Private Sub vsbMain_Change()
    Call vsbMain_Scroll
End Sub

Private Sub vsbMain_Scroll()
    mlngTopVsc = -1 * vsbMain.Value * Screen.TwipsPerPixelY
    picMain.Top = mlngTopVsc
End Sub

Private Sub vsfCert_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfCert
        If Col = COL_证件类型 Then
            .TextMatrix(Row, Col) = zlStr.NeedName(.TextMatrix(Row, Col))
        End If
    End With
End Sub

Private Sub vsfCert_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewCol As Long, lngNewRow As Long
    Dim arrTmp As Variant
    Dim j As Long
    
    lngNewCol = NewCol
    lngNewRow = NewRow
    If lngNewCol = -1 Then Exit Sub
    With vsfCert
        If lngNewCol = COL_证件类型 Then
            .ComboList = .ColData(lngNewCol)
            If Trim(.TextMatrix(lngNewRow, lngNewCol)) <> "" Then
                arrTmp = Split(.ColData(lngNewCol) & "", "|")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    If zlStr.NeedName(arrTmp(j) & "") = .TextMatrix(lngNewRow, lngNewCol) Then
                        .TextMatrix(lngNewRow, lngNewCol) = arrTmp(j)
                        Exit For
                    End If
                Next
            End If
        ElseIf lngNewCol = COL_Del Or lngNewCol = COL_增加 Then
            .ComboList = "..."
            .FocusRect = flexFocusNone
            Set .CellButtonPicture = IIf(lngNewCol = COL_增加, imgAdd, imgDelete)
        ElseIf lngNewCol = COL_图片 Then
             .ComboList = "..."
             .FocusRect = flexFocusNone
             Set .CellButtonPicture = img图片
        Else
            .ComboList = ""
        End If
        If OldCol = COL_证件类型 Then
            .TextMatrix(OldRow, OldCol) = zlStr.NeedName(.TextMatrix(OldRow, OldCol))
        End If
        If lngNewRow >= .FixedRows Then
            '显示图片
            If lngNewCol <> COL_增加 And .TextMatrix(lngNewRow, COL_证件号码) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '下一行诊断为空则不能新增行
                    If .TextMatrix(lngNewRow + 1, COL_证件号码) = "" Then
                         Set .Cell(flexcpPicture, lngNewRow, COL_增加) = imgAdd
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, COL_增加) = imgAdd
                End If
            End If
            '显示图片
            If lngNewCol <> COL_Del Then Set .Cell(flexcpPicture, lngNewRow, COL_Del) = imgDelete
            If lngNewCol <> COL_图片 And .TextMatrix(lngNewRow, COL_证件号码) <> "" Then Set .Cell(flexcpPicture, lngNewRow, COL_图片) = img图片
        End If
    End With
End Sub

Private Sub vsfCert_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = COL_增加 Or lngCol = COL_Del Or lngCol = COL_图片 Then Cancel = True
End Sub

Private Sub vsfCert_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long, int序号 As Long, lngCount As Long, lng证件ID As Long
    Dim strPictureFile As String
    Dim objFile As New FileSystemObject
    Dim rsTmp As New ADODB.Recordset
    Dim strsInfo As String, strsMainInfo As String
    Dim lngRows As Long, lngCounts As Long
    Dim blnAdd As Boolean
    
    lngCol = Col
    lngRow = Row
    With vsfCert
        For i = .FixedRows To .Rows - 1
            If .RowHidden(i) = False Then
                lngRows = lngRows + 1
            End If
        Next
        Select Case lngCol
            Case COL_增加
                For i = .Rows - 1 To .FixedRows Step -1
                    If Trim(.TextMatrix(i, COL_证件号码)) <> "" And .RowHidden(i) = False Then
                        blnAdd = True
                        Exit For
                    ElseIf Trim(.TextMatrix(i, COL_证件号码)) = "" And .RowHidden(i) = False Then
                        Exit For
                    End If
                Next
                If blnAdd = True Then
                     lngRow = .Rows: .AddItem "", lngRow
                     .Row = lngRow: .Col = COL_证件号码
                     .TextMatrix(lngRow, COL_所有者) = IIf(optType(0).Value, "病人本身", "陪诊人")
                     .Cell(flexcpData, lngRow, COL_所有者, lngRow, COL_所有者) = IIf(optType(0).Value, "1", "2")
                     .ShowCell .Row, COL_证件号码
                End If
                lngCounts = 0
            Case COL_Del
                If Trim(.TextMatrix(lngRow, COL_证件号码)) <> "" Then
                    If MsgBox("确定要删除证件号码为【" & .TextMatrix(lngRow, COL_证件号码) & "】的证件信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If lngRows = .FixedRows Then
                            For i = COL_证件ID To COL_备注
                                .TextMatrix(lngRow, i) = ""
                                .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                            Next
                            With vsfImg
                                For j = .FixedRows To .Rows - 1
                                    If Val(.TextMatrix(j, IMG_证件ID)) = lngRow Then
                                        .RemoveItem j
                                        .AddItem "", j
                                        .RowHidden(j) = True
                                    End If
                                Next
                            End With
                            CheckValueChange vsfCert
                        ElseIf lngRows > .FixedRows Then
                            .RemoveItem lngRow
                            .AddItem "", lngRow
                            .RowHidden(lngRow) = True
                            With vsfImg
                                For j = .FixedRows To .Rows - 1
                                    If Val(.TextMatrix(j, IMG_证件ID)) = lngRow Then
                                        .RemoveItem j
                                        .AddItem "", j
                                        .RowHidden(j) = True
                                    End If
                                Next
                            End With
                            For j = .FixedRows To .Rows - 1
                                If Trim(.TextMatrix(j, COL_证件号码)) <> "" Then
                                    .Row = j: .Col = COL_证件号码
                                    Call vsfCert_Click
                                End If
                            Next
                            CheckValueChange vsfCert
                        End If
                    Else
                        .Row = lngRow: .Col = COL_证件号码
                        .ShowCell .Row, COL_证件号码
                    End If
                Else
                    If .Rows - 1 = .FixedRows Or lngRow = .FixedRows Then
                        Exit Sub
                    Else
                        For i = .FixedRows To .Rows - 1
                            If .TextMatrix(i, COL_证件号码) <> "" Then
                                lngCounts = lngCounts + 1
                            End If
                        Next
                        If lngCounts <> 0 Then
                            .RemoveItem lngRow
                            For j = .FixedRows To .Rows - 1
                                If j <= .Rows - 1 Then
                                    If Trim(.TextMatrix(j, COL_证件号码)) <> "" Then
                                        .Row = j: .Col = COL_证件号码
                                        Call vsfCert_Click
                                    End If
                                End If
                            Next
                            CheckValueChange vsfCert
                        End If
                    End If
                End If
            Case COL_图片
                If .TextMatrix(lngRow, COL_证件号码) <> "" Then
                    If gobjPublicPatient Is Nothing Then
                        On Error Resume Next
                        Call CreatePublicPatient
                        Err.Clear: On Error GoTo 0
                    End If
                    If gobjPublicPatient Is Nothing Then
                        MsgBox "创建病人信息公共部件（zlPublicPatient.clsPublicPatient）失败！", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.用户名)
                    If gobjPublicPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
                    If strPictureFile <> "" Then
                        mlngImage = mlngImage + 1
                        objFile.CopyFile strPictureFile, App.Path & "\CertImg\image" & mlngImage & ".bmp"
                        strPictureFile = App.Path & "\CertImg\image" & mlngImage & ".bmp"
                        If ImgCert.Picture <> 0 Then
                            j = vsfImg.Rows - 1
                            For k = vsfImg.FixedRows To vsfImg.Rows - 1
                                If vsfImg.Cell(flexcpData, k, IMG_证件ID, k, IMG_证件ID) = lngRow & "-" & k Then
                                    int序号 = int序号 + 1
                                End If
                            Next
                            With vsfImg
                                .AddItem "" & lngRow & "-" & j, j
                                .Cell(flexcpPicture, j, IMG_图片, j, IMG_图片) = ImgCert
                                .Cell(flexcpPictureAlignment, j, IMG_图片, j, IMG_图片) = 4
                                
                                .Cell(flexcpPicture, j, IMG_Del, j, IMG_Del) = imgDelete
                                .Cell(flexcpPictureAlignment, j, IMG_Del, j, IMG_Del) = 4
                                
                                .TextMatrix(j, IMG_证件ID) = "" & lngRow
                                .Cell(flexcpData, j, IMG_证件ID, j, IMG_证件ID) = "" & lngRow & "-" & j
                                
                                .TextMatrix(j, IMG_序号) = "" & int序号 + 1
                                imgPic.Picture = LoadPicture(strPictureFile) '打开要压缩的图片
                                Call PictureBoxSaveJPG(imgPic.Picture, strPictureFile) '保存压缩后的图片
                                .Cell(flexcpData, j, IMG_图片, j, IMG_图片) = strPictureFile
                            End With
                            int序号 = 0
                            CheckValueChange vsfImg
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfCert_Click()
    Dim i As Long, j As Long, int序号 As Long, lngCount As Long, lng证件ID As Long
    Dim lngRow As Long, lngCol As Long
    
    With vsfCert
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow <> -1 And lngCol <> -1 Then
            If (lngCol = COL_增加 Or lngCol = COL_Del Or lngCol = COL_图片) And lngRow >= .FixedRows Then
                If lngCol = COL_增加 Then
                    If .TextMatrix(lngRow, COL_证件号码) = "" Then Exit Sub
                End If
                .Select lngRow, lngCol
                Call vsfCert_CellButtonClick(lngRow, lngCol)
            Else
                With vsfImg
                    For i = .FixedRows To .Rows - 1
                        If .Cell(flexcpData, i, IMG_证件ID, i, IMG_证件ID) = "" & lngRow & "-" & i And Val(.TextMatrix(i, IMG_证件ID)) > 0 Then
                            .RowHidden(i) = False
                        Else
                            .RowHidden(i) = True
                        End If
                    Next
                End With
            End If
        End If
    End With
End Sub

Private Sub vsfCert_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    Dim i As Long
    
    With vsfCert
        If Col = COL_证件类型 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If .ComboItem(i) = .TextMatrix(Row, Col) Then
                    .Cell(flexcpData, Row, Col, Row, Col) = i
                    Exit For
                End If
            Next
        End If
    End With
    CheckValueChange vsfCert
End Sub

Private Sub vsfCert_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long

    With vsfCert
        If Col = COL_证件类型 Then
            '定位到匹配项
            For i = 0 To .ComboCount - 1
                If zlStr.NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i
                    .Cell(flexcpData, Row, Col, Row, Col) = i
                    Exit For
                End If
            Next
        End If
    End With
    CheckValueChange
End Sub

Private Sub vsfCert_DblClick()
    Call vsfCert_KeyPress(vbKeySpace)
End Sub

Private Sub vsfCert_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim i As Long

    With vsfCert
        lngCol = .Col
        If KeyCode = vbKeyInsert Then
            lngRow = .Row
            If Trim(.TextMatrix(lngRow, COL_证件号码)) <> "" Then
                 lngRow = .Row + 1: .AddItem "", lngRow
                .Row = lngRow: .Col = COL_证件号码
                .Cell(flexcpPicture, lngRow, COL_增加, lngRow, COL_增加) = imgAdd
                .Cell(flexcpPictureAlignment, lngRow, COL_增加, lngRow, COL_增加) = 4
                .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
                .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
                .ShowCell .Row, .Col
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngRow = .Row
            If Trim(.TextMatrix(lngRow, COL_证件号码)) <> "" Then
                If MsgBox("确定要删除证件类型为【" & .TextMatrix(lngRow, COL_证件类型) & "】的证件信息吗？", vbQuestion + vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If .Rows - 1 = .FixedRows Then
                        For i = COL_证件ID To COL_图片
                            .TextMatrix(lngRow, i) = ""
                            .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                        Next
                    ElseIf .Rows - 1 > .FixedRows Then
                        .RemoveItem lngRow
                    End If
                Else
                    .Row = lngRow: .Col = COL_证件号码
                    .ShowCell .Row, .Col
                End If
            Else
                .RemoveItem lngRow
            End If
        End If
    End With
    CheckValueChange vsfCert
End Sub

Private Sub vsfCert_KeyPress(KeyAscii As Integer)
    Dim intRow As Integer, intCol As Integer
    Dim i As Long, j As Long
    
    With vsfCert
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row < .FixedRows Then .Row = .FixedRows
            For i = .Row To .Rows - 1
                For j = IIf(i = .Row, .Col + 1, COL_证件号码) To COL_Del
                    If CertCellEditable(vsfCert, i, j) Then Exit For
                Next
                If j <= COL_Del Then Exit For
            Next
            If i <= .Rows - 1 Then
                Call .Select(i, j)
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            Select Case .Col
                Case COL_证件号码
                    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Or KeyAscii = Asc("*") Then
                        KeyAscii = 0
                    Else
                        intRow = .Row
                        intCol = .Col
                        .ComboList = "" '使按钮状态进入输入状态
                    End If
                Case COL_图片, COL_增加, COL_Del, COL_所有者
                    .ComboList = "..."
                Case COL_备注
                    If zlCommFun.ActualLen(.TextMatrix(.Row, COL_备注)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                        KeyAscii = 0
                    End If
                Case COL_证件号码
                    If zlCommFun.ActualLen(.TextMatrix(.Row, COL_证件号码)) >= 20 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                        KeyAscii = 0
                    End If
            End Select
        End If
    End With
End Sub

Public Function CertCellEditable(ByVal objVsf As Object, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean

    If objVsf Is Nothing Then Exit Function
    If lngCol <> -1 Then
        If objVsf.Name = "vsfCert" Then
            With vsfCert
                If .ColHidden(lngCol) Then Exit Function
                If Trim(.TextMatrix(lngRow, COL_证件号码)) = "" Then
                    If lngCol > COL_证件号码 Then Exit Function
                End If
            End With
        ElseIf objVsf.Name = "vsfInterface" Then
            With vsfInterface
                .Editable = flexEDNone
                If .ColHidden(lngCol) Then Exit Function
                If lngCol <> COLS_认证 Then Exit Function
                If Trim(.TextMatrix(lngRow, COLS_名称)) = "" Then Exit Function
                .Editable = flexEDKbdMouse
            End With
        End If
    End If
    CertCellEditable = True
End Function

Private Sub vsfCert_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long
    
    lngRow = Row
    lngCol = Col
    With vsfCert
        If lngCol = COL_证件号码 Then
            If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With
    CheckValueChange vsfCert
End Sub

Private Sub vsfCert_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsfCert
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsfCert_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long
    
    lngRow = Row
    lngCol = Col
    With vsfCert
        If lngCol = COL_证件号码 Then
            .TextMatrix(lngRow, COL_证件号码) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_证件号码)) >= 20 Then
                MsgBox "证件号码的字符个数不能大于20个字符！", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_备注 Then
            .TextMatrix(lngRow, COL_备注) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_备注)) >= 100 Then
                MsgBox "备注的字符个数不能大于100个字符或者50个汉字！", vbInformation, gstrSysName
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsfImg_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = IMG_Del Then Cancel = True
End Sub

Private Sub vsfImg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long, lngCertRow As Long, lng证件ID As Long, lng序号 As Long
    
    With vsfImg
        lngRow = Row
        lngCol = Col
        lng证件ID = Val(.RowData(lngRow))
        lng序号 = Val(.TextMatrix(lngRow, IMG_序号))
        lngCertRow = Val(.TextMatrix(lngRow, IMG_证件ID))
        Select Case lngCol
            Case IMG_Del
                If .Cell(flexcpData, lngRow, IMG_证件ID, lngRow, IMG_证件ID) <> "" Then
                    If MsgBox("确定要删除第【" & lngRow & "】行的图片信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem lngRow
                        .AddItem "", lngRow
                        .RowHidden(lngRow) = True
                        .Cell(flexcpData, lngRow, IMG_证件ID, lngRow, IMG_证件ID) = "" & lngCertRow & "-" & lngRow
                        .TextMatrix(lngRow, IMG_序号) = "" & lng序号
                        CheckValueChange vsfImg
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfImg_Click()
    Dim lngRow As Long, lngCol As Long
    Dim lng证件ID As Long, lng序号 As Long, lngCertRow As Long
    Dim strFile As String
    Dim vPoint As POINTAPI
    
    vPoint = GetCoordPos(vsfImg.hwnd, vsfImg.CellLeft, vsfImg.CellTop)
    With vsfImg
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow <> -1 And lngCol <> -1 Then
            lng证件ID = Val(.RowData(lngRow))
            lng序号 = Val(.TextMatrix(lngRow, IMG_序号))
            lngCertRow = Val(.TextMatrix(lngRow, IMG_证件ID))
            If lngCol = IMG_图片 And .RowHidden(lngRow) = False Then
                If Trim(.Cell(flexcpData, lngRow, IMG_图片, lngRow, IMG_图片)) = "" And lng证件ID <> 0 And lng序号 <> 0 Then
                    frmCertPicture.ShowMe Me, lng证件ID, 1, vPoint.X, vPoint.Y, vsfImg.Height, lng序号
                Else
                    frmCertPicture.ShowMe Me, 0, 2, vPoint.X, vPoint.Y, vsfImg.Height, 0, .Cell(flexcpData, lngRow, IMG_图片, lngRow, IMG_图片)
                End If
            ElseIf lngCol = IMG_Del Then
                Call vsfImg_CellButtonClick(lngRow, lngCol)
            End If
        End If
    End With
End Sub

Private Sub vsfImg_DblClick()
    Call vsfImg_KeyPress(vbKeySpace)
End Sub


Private Sub vsfImg_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long
    
    With vsfImg
        lngRow = .Row
        lngCol = .Col
        
        If lngCol = IMG_备注 Then
            If zlCommFun.ActualLen(.TextMatrix(lngRow, lngCol)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsfImg_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long
    
    With vsfImg
        lngRow = Row
        lngCol = Col
        If lngCol = IMG_备注 Then
            .TextMatrix(lngRow, IMG_备注) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, lngCol)) >= 100 Then
                MsgBox "备注的字符个数不能大于100个字符或者50个汉字！", vbInformation, gstrSysName
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vsfInterface_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewCol As Long, lngNewRow As Long
    
    lngNewCol = NewCol
    lngNewRow = NewRow
    If lngNewCol = -1 Then Exit Sub
    With vsfInterface
        If Not CertCellEditable(vsfInterface, lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            If lngNewCol = COLS_认证 Then
                 .ComboList = "..."
                 .FocusRect = flexFocusNone
                 Set .CellButtonPicture = imgIdentify
            End If
        End If
    End With
End Sub

Private Sub vsfInterface_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = COLS_认证 Then Cancel = True
End Sub

Private Sub vsfInterface_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    Dim blnCreate As Boolean, blnTrans As Boolean
    Dim strParIn As String, strParOut As String
    Dim strSQL As String, strReason As String
    Dim i As Long
    Dim CurrDate As Date
    
    On Error GoTo errH
    CurrDate = zlDatabase.Currentdate
    With vsfInterface
         lngRow = .Row
         lngCol = .Col
         If lngCol = COLS_认证 And Trim(.TextMatrix(lngRow, COLS_名称)) <> "" Then
            If mblnInfoChange Then
                If Not CheckCertifyData Then Exit Sub
                Call CachPatiData
                Call CachCertInterface
                If mintModel = 1 And mblnChange基本 Then
                    frmGetReason.ShowMe Me, strReason
                    mstrReason = strReason
                    If mstrReason = "" Then
                        Exit Sub
                    End If
                End If
                If SaveCertifyData(0, mintModel) Then
                    If mstrMsg <> "" Then
                        MsgBox mstrMsg, vbInformation, gstrSysName
                    End If
                    mstrMsg = ""
                    mblnSave = True
                    mblnInfoChange = False
                    mblnChange基本 = False
                    mintModel = 1
                    CachAllData
                Else
                    mblnSave = False
                End If
            End If
            blnCreate = CreateIdentifyObj(.TextMatrix(lngRow, COLS_部件名))
            If blnCreate Then
                If mblnSave Then
                    Call SetParIn(strParIn)
                    If mobjIdentify.IdentityCert(strParIn, strParOut) Then
                        .TextMatrix(lngRow, COLS_认证结果) = "已认证"
                        mblnInterface = True
                        Call SetReturnValue(strParOut)
                    Else
                        .TextMatrix(lngRow, COLS_认证结果) = "认证失败"
                    End If
                    CheckValueChange vsfInterface
                    Screen.MousePointer = 11
                    gcnOracle.BeginTrans: blnTrans = True
                    strSQL = "Zl_实名认证接口日志_Insert(" & mlng实名id & "," & Val(.TextMatrix(lngRow, COLS_接口ID)) & "," & IIf(optType(0).Value, 1, 2) & "," & IIf(.TextMatrix(lngRow, COLS_认证结果) = "已认证", 1, 0) & ",'" & gstrDBUser & "',To_Date('" & CurrDate & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    gcnOracle.CommitTrans: blnTrans = False
                    Call SaveInterfaceRecord(lngRow, strParIn, strParOut, CurrDate)
                    Screen.MousePointer = 0
                End If
            Else
                .TextMatrix(lngRow, COLS_认证结果) = "三方部件创建失败"
            End If
         End If
    End With
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfInterface_Click()
    Dim lngRow As Long, lngCol As Long
    On Error GoTo errH
    With vsfInterface
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngCol > 0 And lngRow > 0 Then
            If Not CertCellEditable(vsfInterface, lngRow, lngCol) Then Exit Sub
            If lngCol = COLS_认证 Then
                Call vsfInterface_CellButtonClick(lngRow, lngCol)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetReturnValue(ByVal strParOut As String)
    Dim xTxt As DOMDocument
    Dim xRoot As IXMLDOMElement
    Dim xNode As IXMLDOMNode
    Dim strName As String, strSex As String, strCountry As String, strNation As String, strPlace As String, strAdress As String, strIdNumer As String, strPhone As String, strDate As String
    Dim strMask As String
    
    On Error GoTo errH
    If strParOut <> "" Then
        Set xTxt = New DOMDocument
        Screen.MousePointer = 11
        xTxt.loadXML strParOut
        If xTxt.documentElement Is Nothing Then
            Set xTxt = Nothing
            Screen.MousePointer = 0
            Exit Sub
        End If
        Set xRoot = xTxt.selectSingleNode("OUT")
        Set xNode = xRoot.selectSingleNode("PATI_NAME")
        If Not xNode Is Nothing Then
            strName = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("SEX")
        If Not xNode Is Nothing Then
            strSex = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("COUNTRY")
        If Not xNode Is Nothing Then
            strCountry = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("NATION")
        If Not xNode Is Nothing Then
            strNation = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("BIRTH_PLACE")
        If Not xNode Is Nothing Then
            strPlace = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("ID_NUMBER")
        If Not xNode Is Nothing Then
            strIdNumer = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("PHONE_NUMBER")
        If Not xNode Is Nothing Then
            strPhone = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("ADDRESS")
        If Not xNode Is Nothing Then
            strAdress = xNode.nodeTypedValue
        End If
        Set xNode = xRoot.selectSingleNode("BIRTH_DATE")
        If Not xNode Is Nothing Then
            strDate = xNode.nodeTypedValue
        End If
        If strDate <> "" Then
            If Format(strDate, "HH:MM") = "00:00" Then
                strMask = "####-##-##"
            Else
                strMask = "####-##-## ##:##"
            End If
            strDate = Format(strDate, decode(strMask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
        End If
        If mlngPati = 0 Then
            txtInfo(TXT_姓名) = strName
            cboInfo(CBO_性别).ListIndex = cbo.FindIndex(cboInfo(CBO_性别), strSex)
            cboInfo(CBO_国籍).ListIndex = cbo.FindIndex(cboInfo(CBO_国籍), strCountry)
            cboInfo(CBO_民族).ListIndex = cbo.FindIndex(cboInfo(CBO_民族), strNation)
            If gbln启用结构化地址 Then
                patiAdressInfo(ADRS_出生地点).Value = strPlace
                patiAdressInfo(ADRS_住址).Value = strAdress
                
            Else
                txtAdressInfo(ADRS_出生地点).Text = strPlace
                txtAdressInfo(ADRS_住址).Text = strAdress
            End If
            txtInfo(TXT_身份证号) = strIdNumer
            txtInfo(txt_手机号) = strPhone
            If IsDate(strDate) Then
                txtDateInfo(DATE_出生日期).Mask = strMask
                txtDateInfo(DATE_出生日期).Tag = strMask
                txtDateInfo(DATE_出生日期) = strDate
            End If
        Else
            txtInfo(TXT_陪诊人姓名) = strName
            cboInfo(CBO_陪诊人性别).ListIndex = cbo.FindIndex(cboInfo(CBO_性别), strSex)
            cboInfo(CBO_陪诊人国籍).ListIndex = cbo.FindIndex(cboInfo(CBO_国籍), strCountry)
            cboInfo(CBO_陪诊人民族).ListIndex = cbo.FindIndex(cboInfo(CBO_民族), strNation)
            If gbln启用结构化地址 Then
                patiAdressInfo(ADRS_陪诊人住址).Value = strAdress
            Else
                txtAdressInfo(ADRS_陪诊人住址).Text = strAdress
            End If
            txtInfo(TXT_陪诊人身份证号) = strIdNumer
            txtInfo(txt_手机号) = strPhone
            If IsDate(strDate) Then
                txtDateInfo(DATE_陪诊人出生日期).Mask = strMask
                txtDateInfo(DATE_陪诊人出生日期).Tag = strMask
                txtDateInfo(DATE_陪诊人出生日期) = strDate
            End If
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveInterfaceRecord(ByVal lngRow As Long, ByVal strParIn As String, ByVal strParOut As String, ByVal CurrDate As Date) As String
    On Error GoTo Errhand

    If Sys.SaveLob(glngSys, 34, mlng实名id & "|" & Val(vsfInterface.TextMatrix(lngRow, COLS_接口ID)) & "|" & CurrDate & "|0", strParIn, 1) = False Then
        MsgBox "实名认证接口日志保存失败！", vbInformation, gstrSysName
    End If
    If Sys.SaveLob(glngSys, 34, mlng实名id & "|" & Val(vsfInterface.TextMatrix(lngRow, COLS_接口ID)) & "|" & CurrDate & "|1", strParOut, 1) = False Then
        MsgBox "实名认证接口日志保存失败！", vbInformation, gstrSysName
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetParIn(ByRef strParIn As String)
    Dim objCtl As Object
    Dim strTmp As String, strCertType As String, strCertNumber As String, strCertPati As String, strPicture As String, strCert As String
    Dim strCertImage As Variant
    Dim i As Long, intTYPE As Integer, j As Long
    Dim strPatiImg As String, strCerts As String
    Dim strFile As String
    
    On Error GoTo errH
    If Trim(txtInfo(TXT_身份证号).Text) <> "" Then
        intTYPE = 0
    Else
        If optType(0).Value = True Then
            If vsfCert.Rows - 1 > vsfCert.FixedRows Then
                intTYPE = 0
            End If
        Else
            If Trim(txtInfo(TXT_陪诊人身份证号).Text) <> "" Then
                intTYPE = 1
            Else
                If optType(1).Value = True Then
                    If vsfCert.Rows - 1 > vsfCert.FixedRows Then
                        intTYPE = 1
                    End If
                End If
            End If
        End If
    End If
    mlngPati = intTYPE
    Screen.MousePointer = 11
    If mstr采集图片 <> "" Then
        Call PictureBoxSaveJPG(imgPatient.Picture, mstr采集图片) '保存压缩后的图片
        strPatiImg = zlStr.EncodeBase64_File(mstr采集图片)
    End If
    If intTYPE = 0 Then
        strTmp = "<cert_id>" & mlng实名id & "</cert_id><pati_Id>" & mlng病人ID & "</pati_Id><pati_name>" & txtInfo(TXT_姓名).Text & "</pati_name>" & _
        "<sex>" & zlCommFun.GetNeedName(cboInfo(CBO_性别).Text, "-") & "</sex><birth_date>" & txtDateInfo(DATE_出生日期) & "</birth_date><country>" & cboInfo(CBO_国籍) & "</country><nation>" & _
        cboInfo(CBO_民族) & "</nation><birth_place>" & IIf(gbln启用结构化地址, patiAdressInfo(ADRS_出生地点).Value, txtAdressInfo(ADRS_出生地点)) & "</birth_place>" & _
        "<address>" & IIf(gbln启用结构化地址, patiAdressInfo(ADRS_住址).Value, txtAdressInfo(ADRS_住址)) & "</address><id_number>" & txtInfo(TXT_身份证号).Text & "</id_number>" & _
        "<phone_number>" & txtInfo(txt_手机号).Text & "</phone_number><pati_Image>" & strPatiImg & "</pati_Image>"
    Else
        strTmp = "<cert_id>" & mlng实名id & "</cert_id><pati_Id>" & mlng病人ID & "</pati_Id><pati_name>" & txtInfo(TXT_陪诊人姓名).Text & "</pati_name>" & _
        "<sex>" & zlCommFun.GetNeedName(cboInfo(CBO_性别).Text, "-") & "</sex><birth_date>" & txtDateInfo(DATE_陪诊人出生日期) & "</birth_date><country>" & cboInfo(CBO_陪诊人国籍) & "</country><nation>" & _
        cboInfo(CBO_陪诊人民族) & "</nation><birth_place></birth_place>" & _
        "<address>" & IIf(gbln启用结构化地址, patiAdressInfo(ADRS_陪诊人住址).Value, txtAdressInfo(ADRS_陪诊人住址)) & "</address><id_number>" & txtInfo(TXT_陪诊人身份证号).Text & "</id_number>" & _
        "<phone_number>" & txtInfo(txt_手机号).Text & "</phone_number><pati_Image>" & strPatiImg & "</pati_Image>"
    End If
    With vsfCert
        If .Rows > .FixedRows Then
            For i = .FixedRows To .Rows - 1
                If IIf(Trim(.TextMatrix(i, COL_所有者)) = "病人本身", 0, 1) = intTYPE Then
                    strCertType = .TextMatrix(i, COL_证件类型)
                    strCertNumber = .TextMatrix(i, COL_证件号码)
                    strCertPati = IIf(Trim(.TextMatrix(i, COL_所有者)) = "病人本身", "1", "2")
                    With vsfImg
                        For j = .FixedRows To .Rows - 1
                            If .RowData(j) = vsfCert.TextMatrix(i, COL_证件ID) Then
                                If .Cell(flexcpData, i, IMG_图片, i, IMG_图片) = "" Then
                                    Call ReadPatPricture(.RowData(j) & "," & Val(.TextMatrix(j, IMG_序号)), imgLoad, strFile)
                                    If strFile <> "" Then
                                        Call PictureBoxSaveJPG(imgLoad.Picture, strFile) '保存压缩后的图片
                                        strPatiImg = zlStr.EncodeBase64_File(strFile)
                                        Kill strFile
                                    End If
                                Else
                                    strPatiImg = zlStr.EncodeBase64_File(.Cell(flexcpData, i, IMG_图片, i, IMG_图片))
                                End If
                                strPicture = strPicture & "<IMAGE><IMAGE_CODE>" & strPatiImg & "</IMAGE_CODE><NOTE>" & .TextMatrix(i, IMG_备注) & "</NOTE></IMAGE>"
                            End If
                        Next
                    End With
                    strCert = strCert & "<CERTS><CERT_TYPE>" & strCertType & "</CERT_TYPE><CERT_NUMBER>" & strCertNumber & "</CERT_NUMBER><CERT_OWNER>" & strCertPati & "</CERT_OWNER><CERT_IMAGES>" & strPicture & "</CERT_IMAGES></CERTS>"
                    strCerts = strCerts & "<OTHER_CERTS>" & strCert & "</OTHER_CERTS>"
                Else
                    strCerts = "<OTHER_CERTS><CERTS><CERT_TYPE></CERT_TYPE><CERT_NUMBER></CERT_NUMBER><CERT_OWNER></CERT_OWNER><CERT_IMAGES></CERT_IMAGES></CERTS></OTHER_CERTS>"
                End If
            Next
        End If
    End With
    Screen.MousePointer = 0
    If strTmp <> "" Then
        strTmp = "<XML_IN>" & strTmp & strCerts & "</OTHER_CERTS></XML_IN>"
    End If
    strParIn = strTmp
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsfInterface_DblClick()
    Call vsfInterface_KeyPress(vbKeySpace)
End Sub

Private Sub vsfInterface_KeyPress(KeyAscii As Integer)
    Dim longRow As Integer, longCol As Integer
    Dim i As Long, j As Long
    
    With vsfInterface
        longRow = .Row
        longCol = .Col
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row < .FixedRows Then .Row = .FixedRows
            For i = .Row To .Rows - 1
                For j = IIf(i = .Row, .Col + 1, COLS_名称) To COLS_认证
                    If CertCellEditable(vsfInterface, i, j) Then Exit For
                Next
                If j <= COLS_认证 Then Exit For
            Next
            If i <= .Rows - 1 Then
                Call .Select(i, j)
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            If Not CertCellEditable(vsfInterface, longRow, longCol) Then
                KeyAscii = 0
                Exit Sub
            End If
            Select Case .Col
                Case COLS_认证
                    .ComboList = "..."
            End Select
        End If
    End With
End Sub

Private Function CheckPPatiInfo() As Boolean
'功能：检查陪诊人信息
    Dim objCtrl As Object
    Dim blnLocked As Boolean
    Dim i As Long
    Dim bln陪诊 As Boolean
    
    With Me.Controls
        For Each objCtrl In Me.Controls
            Select Case objCtrl.Name
                Case "txtInfo"
                    If objCtrl.Index = TXT_陪诊人身份证号 Then
                        bln陪诊 = IIf(objCtrl.Text <> "", True, bln陪诊)
                    End If
                Case "txtAdressInfo"
                    If objCtrl.Index = ADRS_陪诊人住址 Then
                        bln陪诊 = IIf(objCtrl.Text <> "", True, bln陪诊)
                    End If
                Case "patiAdressInfo"
                    If objCtrl.Index = ADRS_陪诊人住址 Then
                        bln陪诊 = IIf(objCtrl.Value <> "", True, bln陪诊)
                    End If
                Case "txtDateInfo"
                    If objCtrl.Index = DATE_陪诊人出生日期 Then
                        If IsDate(objCtrl.Text) Or (objCtrl.Text <> "____-__-__ __:__" And objCtrl.Text <> "____-__-__") Then
                            bln陪诊 = True
                        Else
                            bln陪诊 = bln陪诊
                        End If
                    End If
                Case "txtInfoDate"
'                    If objCtrl.Index = DATE_陪诊人出生日期 Then
'                        bln陪诊 = IIf(objCtrl.Text <> "", True, bln陪诊)
'                    End If
                Case "optType"
                    If objCtrl.Index = 1 Then
                        If objCtrl.Value = True Then
                            For i = vsfCert.FixedRows To vsfCert.Rows - 1
                                If vsfCert.TextMatrix(i, COL_证件号码) <> "" Then
                                    bln陪诊 = True
                                End If
                            Next
                        End If
                    End If
            End Select
        Next
    End With
    CheckPPatiInfo = bln陪诊
End Function

Private Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80) As Boolean
     Dim tSI As GdiplusStartupInput
     Dim lRes As Long
     Dim lGDIP As Long
     Dim lBitmap As Long
    
     '初始化 GDI+
     tSI.GdiplusVersion = 1
     lRes = GdiplusStartup(lGDIP, tSI, 0)
    
     If lRes = 0 Then
        '从句柄创建 GDI+ 图像
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
    
         If lRes = 0 Then
             Dim tJpgEncoder As GUID
             Dim tParams As EncoderParameters
            
             '初始化解码器的GUID标识
             CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            
             '设置解码器参数
             tParams.Count = 1
            With tParams.Parameter ' Quality
            '得到Quality参数的GUID标识
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
            .NumberOfValues = 1
            .type = 4
            .Value = VarPtr(quality)
            End With
        
            '保存图像
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
            
            '销毁GDI+图像
            GdipDisposeImage lBitmap
         End If
    
        '销毁 GDI+
        GdiplusShutdown lGDIP
     End If
    
     If lRes Then
        PictureBoxSaveJPG = False
     Else
        PictureBoxSaveJPG = True
     End If
End Function





