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
   Caption         =   "ʵ����Ϣ�Ǽ�"
   ClientHeight    =   12750
   ClientLeft      =   225
   ClientTop       =   -3510
   ClientWidth     =   14700
   Icon            =   "frmCertifyRegist.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12750
   ScaleWidth      =   14700
   StartUpPosition =   2  '��Ļ����
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "����״̬"
            TextSave        =   "����״̬"
            Key             =   "����״̬"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
               Caption         =   "��"
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
               Caption         =   "��"
               Height          =   255
               Index           =   1
               Left            =   11820
               TabIndex        =   1
               Top             =   2745
               Width           =   255
            End
            Begin VB.CommandButton cmdAdress 
               Appearance      =   0  'Flat
               Caption         =   "��"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Tag             =   "������ַ"
               Top             =   2760
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Tag             =   "��סַ"
               Top             =   2760
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
               Tag             =   "��ϵ�˵�ַ"
               Top             =   5535
               Width           =   4785
               _ExtentX        =   8440
               _ExtentY        =   476
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "���"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�ɼ�"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�ļ�"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "��    ��"
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
               Caption         =   "��    ��"
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
               Caption         =   "��    ��"
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
               Caption         =   "��������"
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
               Caption         =   "���֤��"
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
               Caption         =   "��    ��"
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
               Caption         =   "��    ��"
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
               Caption         =   "��������"
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
               Caption         =   "��    ��"
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
               Caption         =   "��    ��"
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
               Caption         =   "���֤����"
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
               Caption         =   "�����ص�"
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
               Caption         =   "��    ϵ"
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
               Caption         =   "��������Ϣ"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "������Ϣ"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "ס    ַ"
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
               Caption         =   "���֤��"
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
               Caption         =   "���֤����"
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
               Caption         =   "�����ص�"
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
               Caption         =   "��    ��"
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
               Caption         =   "���˱���"
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
               Caption         =   "������"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "��    ע"
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
               Caption         =   "�� �� ��"
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
               Caption         =   "����֤��"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "������֤"
               BeginProperty Font 
                  Name            =   "����"
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
   Begin VB.Image imgͼƬ 
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

Private mlng����ID As Long, mlngʵ��id As Long, mlng֤��id As Long, mintModel As Integer
Private mfrmParent As Object '������
Private mrsPati As New ADODB.Recordset '����ʵ����Ϣ
Private mrsCert As New ADODB.Recordset 'ʵ��֤��
Private mrsIneterface As New ADODB.Recordset '������֤�ӿ�
Private mblnɨ�����֤�Ǽ� As Boolean
Private mblnChange As Boolean
Private mblnInfoChange As Boolean  '�����Ƿ����仯
Private mblnSave As Boolean  '�Ƿ��Ѿ�����
Private mblnIdentifySure As Boolean '�Ƿ��Ѿ�ȷ����֤
Private mrsMainInfo As ADODB.Recordset  '������Ϣ����Ϣ��¼��
Private mrsSecdInfo  As ADODB.Recordset '�б��¼��
Private mblnLoadFilish As Boolean  '�Ƿ�������
Private mobjIdentify As Object  '������֤�ӿڲ���
Private mblnInterface As Boolean '�����ӿ��Ƿ���֤ͨ��
Private mintDate As Integer  'ʱ��ռ�����
Private mstrReason As String '���ԭ��
Private mlngͼ����� As Long  '����ͼƬ��������
Private mstr�ɼ�ͼƬ As String '�ɼ���ͼƬ·��
Private mlngImage As Long '֤��ͼƬ������
Private mlngPati As Long
Private mlngTopVsc As Long
Private mblnChange���� As Boolean
Private mstrAge As String
Private mstrMsg As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1

Private Enum TXT_Info
    TXT_���� = 0
    TXT_���֤�� = 1
    TXT_���������� = 2
    TXT_���������֤�� = 3
    txt_�ֻ��� = 4
    TXT_��ע = 5
End Enum

Private Enum DATE_Info
    DATE_�������� = 0
    DATE_�����˳������� = 1
End Enum

Private Enum PatiAress_Info
    ADRS_�����ص� = 0
    ADRS_סַ = 1
    ADRS_������סַ = 2
End Enum

Private Enum VSFCert_COL
    COL_֤��ID = 0
    COL_֤������
    COL_֤������
    COL_��ע
    COL_������
    COL_ͼƬ
    COL_����
    COL_Del
End Enum

Private Enum VSFIMG_COL
    IMG_֤��ID = 0
    IMG_���
    IMG_ͼƬ
    IMG_��ע
    IMG_Del
End Enum

Private Enum VSFInterface_COL
    COLS_�ӿ�ID = 0
    COLS_����
    COLS_������
    COLS_˵��
    COLS_��֤���
    COLS_��֤
End Enum

Private Enum Change_State
    CS_ɾ���� = -1
    CS_δ�ı� = 0
    CS_������ = 1
    CS_�滻�� = 2
    CS_������ = 3
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
'���ܣ����ؼ�����
    Dim objText As Object
    
    For Each objText In Me.Controls
        If TypeName(objText) = "TextBox" Or TypeName(objText) = "Frame" Then
            If objText.Name <> "txtAdressInfo" Then
                DrawLineCTL objText
            ElseIf objText.Name = "txtAdressInfo" Then
                If Not gbln���ýṹ����ַ Then
                    DrawLineCTL objText
                End If
            End If
        End If
    Next
End Function

Private Sub InitVsfGridHeader()
'���ܣ���ʼ���б�
    Dim strHeader As String
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
     
    '֤����Ϣ�б�
    strHeader = "֤��ID;֤������,2000,1;֤������,2000,1;��ע,2050,1;������,1000,4;,270,4;,270,4;,270,4"
    Call grid.Init(vsfCert, strHeader)
    With vsfCert
        If Not .ColHidden(COL_֤������) Then
            strSQL = "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '֤������' ���� From ֤������"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If Not rsTmp.EOF Then
                strTmp = " |" & Rec.ToComboList(rsTmp, "[0]-[1]|", "����", "����")
            Else
                strTmp = " |"
            End If
            .ColData(COL_֤������) = strTmp
        End If
    End With
    
    '������֤�ӿ�
    strHeader = "�ӿ�ID;����,3000,1;������;˵��,6000,1;��֤���,2000,4;,270,4"
    Call grid.Init(vsfInterface, strHeader)
    
    'ͼƬ�б�
    strHeader = "֤��ID;���;ͼƬ,600,4;��ע,2650,1;,270,4"
    Call grid.Init(vsfImg, strHeader)
End Sub

Private Sub DrawLineCTL(ByRef objCtl As Object, Optional ByVal bytModel As Byte = 0)
'����:��ָ������һ���߻������ԭ������
'objCtl-����ؼ����󣬸��ݸÿؼ������ȡ��Ӧ����ֵ
'bytModel=0-����;1-�����
    Dim objPic As Object  '����
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    Select Case TypeName(objCtl)
    Case "TextBox"
        '��ÿ��TextBox ���滭һ����
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
        objPic.Line (x1, y1)-(x2, y2), objPic.BackColor '�������
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
                If mintModel = 1 And mblnChange���� Then
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
                    mblnChange���� = False
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
                If MsgBox("������֤�ӿ�δ��֤,�Ƿ�����˹���֤��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
                If MsgBox("�ò��˵�ʵ����Ϣ�Ѿ���֤,ȷ��Ҫȡ����֤��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
'���ܣ���֤����Ϣ����
    Dim i As Long, j As Long, k As Long
    Dim lng״̬ As Long
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
            If .TextMatrix(i, COL_֤������) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_֤��ID)) & "|" & zlCommFun.GetNeedName(.TextMatrix(i, COL_֤������), "-") & "|" & .TextMatrix(i, COL_֤������)
                strInfo = strMainInfo & "|" & .TextMatrix(i, COL_��ע) & "|" & .TextMatrix(i, COL_������)
 
                If InStr("," & strAll & ",", "," & strMainInfo & ",") > 0 Then
                    '��ͬ��ÿ��¼
                    .Tag = i
                    .Cell(flexcpBackColor, i, .FixedCols, i, COL_֤������) = &HC0C0FF
                    Call .ShowCell(i, COL_֤������)
                    Exit Function
                Else
                    strAll = strAll & "," & strMainInfo '�ռ�������������ж��Ƿ����ظ���
                End If
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "�ؼ���='vsfCert' and ���=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!��� = lngTmp
                   mrsSecdInfo!�ؼ��� = "vsfCert"
               End If
               mrsSecdInfo!��ID = Val(.RowData(i))
               mrsSecdInfo!��Ϣ��ֵ = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!����Ϣ��ֵ = IIf(strMainInfo = "", Null, strMainInfo)
               mrsSecdInfo!IndexEx = i
               mrsSecdInfo.Update
               lngTmp = lngTmp + 1

               mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "�ؼ���='vsfCert'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
            mrsSecdInfo.MoveNext
        Next
        
        '����Ϣ�ı�����Ҫ����ɾ������
        mrsSecdInfo.Filter = "(�ı�״̬=" & CS_ɾ���� & " And �ؼ���='vsfCert')" ' OR (�ı�״̬=" & CS_�滻�� & " And �ؼ���='vsfCert')"
        Do While Not mrsSecdInfo.EOF
            strDels = "" & mrsSecdInfo!ԭID
            If strDels <> "" Then
                strAllInfo = strAllInfo & "," & mlngʵ��id & "-" & Val(strDels) & "-----"
            End If
            mrsSecdInfo.MoveNext
        Loop
        mrsSecdInfo.Filter = "�ؼ���='vsfCert' And �ı�״̬>" & CS_δ�ı�
        If Not mrsSecdInfo.EOF Then
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                If mrsSecdInfo!�ı�״̬ = CS_������ Then
                    strAllInfo = strAllInfo & "," & mlngʵ��id & "-" & mrsSecdInfo!ԭID & "-" & zlCommFun.GetNeedName(.TextMatrix(lngRow, COL_֤������), "-") & "-" & .TextMatrix(lngRow, COL_֤������) & "-" & .TextMatrix(lngRow, COL_��ע) & "-" & IIf(.TextMatrix(lngRow, COL_������) = "���˱���", 1, 2)
                Else
                    strAllInfo = strAllInfo & "," & mlngʵ��id & "-" & mrsSecdInfo!ԭID & "-" & zlCommFun.GetNeedName(.TextMatrix(lngRow, COL_֤������), "-") & "-" & .TextMatrix(lngRow, COL_֤������) & "-" & .TextMatrix(lngRow, COL_��ע) & "-" & IIf(.TextMatrix(lngRow, COL_������) = "���˱���", 1, 2)
                End If
                mrsSecdInfo.MoveNext
            Loop
        Else
            mrsSecdInfo.Filter = "�ؼ���='vsfCert' And �ı�״̬=" & CS_δ�ı�
            Do While Not mrsSecdInfo.EOF
                lngRow = mrsSecdInfo!IndexEx
                strAllInfo = strAllInfo & "," & mlngʵ��id & "-" & mrsSecdInfo!ԭID & "-" & .TextMatrix(lngRow, COL_֤������) & "-" & .TextMatrix(lngRow, COL_֤������) & "-" & .TextMatrix(lngRow, COL_��ע) & "-" & IIf(.TextMatrix(lngRow, COL_������) = "���˱���", 1, 2)
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
'���ܣ���ͼƬ��Ϣ����
    Dim i As Long, j As Long, k As Long
    Dim lng״̬ As Long
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
            If .TextMatrix(i, IMG_֤��ID) <> "" Then
                strsInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ) & "|" & .TextMatrix(i, IMG_���) & "|" & .TextMatrix(i, IMG_��ע)
                strsMainInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ) & "|" & .TextMatrix(i, IMG_���) & "|" & .TextMatrix(i, IMG_��ע)
            Else
                strsInfo = ""
                strsMainInfo = ""
            End If
            mrsSecdInfo.Filter = "�ؼ���='vsfImg' and ���=" & lngTmp
            If mrsSecdInfo.EOF Then
                mrsSecdInfo.AddNew
                mrsSecdInfo!��� = lngTmp
                mrsSecdInfo!�ؼ��� = "vsfImg"
            End If
            mrsSecdInfo!��ID = Val(.RowData(i))
            mrsSecdInfo!��Ϣ��ֵ = IIf(strsInfo = "", Null, strsInfo)
            mrsSecdInfo!����Ϣ��ֵ = IIf(strsMainInfo = "", Null, strsMainInfo)
            mrsSecdInfo!IndexEx = i
            mrsSecdInfo.Update
            lngTmp = lngTmp + 1

            mrsSecdInfo.Filter = 0
        Next
        mrsSecdInfo.Filter = "�ؼ���='vsfImg'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
            mrsSecdInfo.MoveNext
        Next
        

        'ɾ�����Լ�����Ϣ�ı�����Ҫ����ɾ������
        mrsSecdInfo.Filter = "(�ı�״̬=" & CS_ɾ���� & " And �ؼ���='vsfImg')" ' OR (�ı�״̬=" & CS_�滻�� & " And �ؼ���='vsfImg')"
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            strDels = "" & .TextMatrix(lngRow, IMG_���)
            If Val(strDels) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ʵ��֤��ͼƬ_Delete(" & .RowData(lngRow) & "," & Val(strDels) & ")"
            End If
            mrsSecdInfo.MoveNext
        Loop
    
        '����Ϣ�ı��Լ���������Ҫ���ò������        '�μ���Ϣ�ı䣬���ø��¹���
        mrsSecdInfo.Filter = "�ؼ���='vsfImg' And �ı�״̬>" & CS_δ�ı�
    
        Do While Not mrsSecdInfo.EOF
            lngRow = mrsSecdInfo!IndexEx
            If mrsSecdInfo!�ı�״̬ <> CS_������ Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ʵ��֤��ͼƬ_Update(" & Val(.RowData(lngRow)) & "," & Val(.TextMatrix(lngRow, IMG_���)) & ",'" & .TextMatrix(lngRow, IMG_��ע) & "')"
            Else
                 Call SaveCertPicture(Val(.RowData(lngRow)), Val(.TextMatrix(lngRow, IMG_���)), .TextMatrix(lngRow, IMG_��ע), .Cell(flexcpData, lngRow, IMG_ͼƬ, lngRow, IMG_ͼƬ))
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
'��������ʵ����Ϣ��ˢ������
    Dim strSQL As String
    Dim rsPati As New ADODB.Recordset
    Dim i As Long
    Dim str���� As String, str֤������ As String
    
    On Error GoTo errH
        strSQL = "Select ����ID,ʵ��ID  From ����ʵ����Ϣ where ����=[1] And ���֤��=[1]"
        Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ˢ�²�����Ϣ", txtInfo(TXT_����).Text, txtInfo(TXT_���֤��).Text)
        If rsPati.EOF Then
            strSQL = "Select ����ID,ʵ��ID From ����ʵ����Ϣ where ����=[1] And ����������=[2] And ���������֤��=[3]"
            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ˢ�²�����Ϣ", txtInfo(TXT_����).Text, txtInfo(TXT_����������).Text, txtInfo(TXT_���������֤��).Text)
            If rsPati.EOF Then
                With vsfCert
                    For i = .FixedRows To .Rows - 1
                        str���� = .TextMatrix(i, COL_֤������)
                        str֤������ = .TextMatrix(i, COL_֤������)
                        strSQL = "Select A.����ID,A.ʵ��ID From ����ʵ����Ϣ A,����ʵ��֤�� B where A.ʵ��ID=B.ʵ��ID And ����=[1] And B.֤������=[2] And B.֤������=[3]" & IIf(optType(0).Value, " And B.������=[4]", " And A.����������=[4] And B.������=[5]")
                        If optType(0).Value Then
                            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ˢ�²�����Ϣ", txtInfo(TXT_����).Text, str֤������, str����, 1)
                        Else
                            Set rsPati = zlDatabase.OpenSQLRecord(strSQL, "ˢ�²�����Ϣ", txtInfo(TXT_����).Text, str֤������, str����, txtInfo(TXT_����������).Text, 2)
                        End If
                        If Not rsPati.EOF Then
                            Exit For
                        End If
                    Next
                End With
            End If
        End If
    If Not rsPati.EOF Then
        mlngʵ��id = rsPati!ʵ��ID & ""
        mlng����ID = rsPati!����ID & ""
    End If
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub zlRefreshCert()
'����ʵ��֤����Ϣ��ˢ������
    Dim strSQL As String
    Dim rsCert As New ADODB.Recordset
    Dim i As Long, j As Long
    Dim str���� As String, str֤������ As String
    
    With vsfCert
        For i = .FixedRows To .Rows - 1
            str���� = .TextMatrix(i, COL_֤������)
            str֤������ = .TextMatrix(i, COL_֤������)
            strSQL = "Select A.����ID,A.ʵ��ID,B.ID as ֤��ID From ����ʵ����Ϣ A,����ʵ��֤�� B where A.ʵ��ID=B.ʵ��ID And A.ʵ��ID=[1] And ֤������=[2] And ֤������=[3]"
            Set rsCert = zlDatabase.OpenSQLRecord(strSQL, "ˢ�²�����Ϣ", mlngʵ��id, str֤������, str����)
            If Not rsCert.EOF Then
                .Cell(flexcpData, i, COL_֤��ID, i, COL_֤��ID) = i & ""
                .TextMatrix(i, COL_֤��ID) = rsCert!֤��ID & ""
            Else
                .Cell(flexcpData, i, COL_֤��ID, i, COL_֤��ID) = i & ""
                .TextMatrix(i, COL_֤��ID) = zlDatabase.GetNextId("����ʵ��֤��") & ""
            End If
            With vsfImg
                For j = .FixedRows To .Rows - 1
                    If .Cell(flexcpData, j, IMG_֤��ID, j, IMG_֤��ID) = "" & i & "-" & j Then
                        .RowData(j) = Val(vsfCert.TextMatrix(i, COL_֤��ID))
                    End If
                Next
            End With
        Next
    End With
End Sub

Private Function CachCertInterface() As Boolean
'���ܣ���֤����Ϣ����
'���ܣ���ȡ��ϱ����SQL
    Dim i As Long, j As Long, k As Long
    Dim lng״̬ As Long
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
        arrMain = Array(COLS_����, COLS_������, COLS_˵��, COLS_��֤���, COLS_��֤)
        arrWhole = Array(COLS_�ӿ�ID, COLS_������, COLS_����, COLS_˵��, COLS_��֤���, COLS_��֤)
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COLS_��֤���) <> "" Then
                If strTmp <> .TextMatrix(i, COLS_��֤���) Then
                    j = 1: strTmp = .TextMatrix(i, COLS_��֤���)
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
                mrsSecdInfo.Filter = "�ؼ���='" & strVsName & "' and ���=" & lngTmp
 
                If mrsSecdInfo.EOF Then
                    mrsSecdInfo.AddNew
                    mrsSecdInfo!��� = lngTmp
                    mrsSecdInfo!�ؼ��� = strVsName
                End If
                mrsSecdInfo!��Ϣ��ֵ = strInfo
                mrsSecdInfo!����Ϣ��ֵ = strMainInfo
                mrsSecdInfo!IndexEx = i
                mrsSecdInfo.Update
                lngTmp = lngTmp + 1
                mrsSecdInfo.Filter = 0
            End If
        Next
        mrsSecdInfo.Filter = "�ؼ���='" & strVsName & "'"
        For i = 1 To mrsSecdInfo.RecordCount
            lng״̬ = CS_δ�ı�
            If mrsSecdInfo!��Ϣԭֵ & "" <> mrsSecdInfo!��Ϣ��ֵ & "" Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣԭֵ) Then
                lng״̬ = CS_������
            End If
            If lng״̬ = CS_������ And IsNull(mrsSecdInfo!��Ϣ��ֵ) Then
                lng״̬ = CS_ɾ����
            End If
            If lng״̬ = CS_������ And mrsSecdInfo!����Ϣԭֵ & "" <> mrsSecdInfo!����Ϣ��ֵ & "" Then
                lng״̬ = CS_�滻��
            End If
            mrsSecdInfo.Update "�ı�״̬", lng״̬
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
'���ܣ����没��ʵ����Ϣ
    Dim strCtlName As String, strFilter As String, strValue As String
    Dim objCtl As Object
    Dim strBirthdate As String
    
    On Error GoTo errH
    For Each objCtl In Me.Controls
        strCtlName = objCtl.Name
        Select Case strCtlName
            Case "txtInfo"
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = objCtl.Text
                    mrsMainInfo!��Ϣ��ֵ = strValue
                    mrsMainInfo.Update
                End If
            Case "cboInfo"
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = zlCommFun.GetNeedName(objCtl.Text, "-")
                    mrsMainInfo!��Ϣ��ֵ = strValue
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
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    If IsDate(strBirthdate) Then
                        strValue = strBirthdate
                    Else
                        strValue = ""
                    End If
                    mrsMainInfo!��Ϣ��ֵ = strValue
                    mrsMainInfo.Update
                End If
            Case "txtAdressInfo"
                If gbln���ýṹ����ַ Then
                    strFilter = "�ؼ���='patiAdressInfo ' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = patiAdressInfo(objCtl.Index).Value
                        mrsMainInfo!��Ϣ��ֵ = strValue
                        mrsMainInfo.Update
                    End If
                Else
                    strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = objCtl.Text
                        mrsMainInfo!��Ϣ��ֵ = strValue
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
    Dim lng״̬ As Long
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
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = objCtl.Text
                    mrsMainInfo!��Ϣԭֵ = strValue
                    mrsMainInfo.Update
                End If
            Case "cboInfo"
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    strValue = zlCommFun.GetNeedName(objCtl.Text, "-")
                    mrsMainInfo!��Ϣԭֵ = strValue
                    mrsMainInfo.Update
                End If
            Case "txtDateInfo"
                strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                mrsMainInfo.Filter = strFilter
                If Not mrsMainInfo.EOF Then
                    If IsDate(objCtl.Text) Then
                        strValue = objCtl.Text
                    Else
                        strValue = ""
                    End If
                    mrsMainInfo!��Ϣԭֵ = strValue
                    mrsMainInfo.Update
                End If
            Case "txtAdressInfo"
                If gbln���ýṹ����ַ Then
                    strFilter = "�ؼ���='patiAdressInfo ' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = patiAdressInfo(objCtl.Index).Value
                        mrsMainInfo!��Ϣԭֵ = strValue
                        mrsMainInfo.Update
                    End If
                Else
                    strFilter = "�ؼ���='" & strCtlName & "' and Index= " & objCtl.Index
                    mrsMainInfo.Filter = strFilter
                    If Not mrsMainInfo.EOF Then
                        strValue = objCtl.Text
                        mrsMainInfo!��Ϣԭֵ = strValue
                        mrsMainInfo.Update
                    End If
                End If
        End Select
    Next
    
    With vsfCert
        .Tag = ""
        lngTmp = 1
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_֤������) <> "" Then
                strMainInfo = Val(.TextMatrix(i, COL_֤��ID)) & "|" & zlCommFun.GetNeedName(.TextMatrix(i, COL_֤������), "-") & "|" & .TextMatrix(i, COL_֤������)
                strInfo = strMainInfo & "|" & .TextMatrix(i, COL_��ע) & "|" & .TextMatrix(i, COL_������)
                .RowData(i) = .TextMatrix(i, COL_֤��ID)
            Else
                strMainInfo = ""
                strInfo = ""
            End If
               mrsSecdInfo.Filter = "�ؼ���='vsfCert' and ���=" & lngTmp
               If mrsSecdInfo.EOF Then
                   mrsSecdInfo.AddNew
                   mrsSecdInfo!��� = lngTmp
                   mrsSecdInfo!�ؼ��� = "vsfCert"
               End If
               mrsSecdInfo!ԭID = Val(.RowData(i))
               mrsSecdInfo!��Ϣԭֵ = IIf(strInfo = "", Null, strInfo)
               mrsSecdInfo!����Ϣԭֵ = IIf(strMainInfo = "", Null, strMainInfo)
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
            If .TextMatrix(i, IMG_֤��ID) <> "" Then
                strsInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ) & "|" & .TextMatrix(i, IMG_���) & "|" & .TextMatrix(i, IMG_��ע)
                strsMainInfo = .RowData(i) & "|" & .Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ) & "|" & .TextMatrix(i, IMG_���) & "|" & .TextMatrix(i, IMG_��ע)
            Else
                strsInfo = ""
                strsMainInfo = ""
            End If
            mrsSecdInfo.Filter = "�ؼ���='vsfImg' and ���=" & lngTmp
            If mrsSecdInfo.EOF Then
                mrsSecdInfo.AddNew
                mrsSecdInfo!��� = lngTmp
                mrsSecdInfo!�ؼ��� = "vsfImg"
            End If
            mrsSecdInfo!ԭID = Val(.RowData(i))
            mrsSecdInfo!��Ϣԭֵ = IIf(strsInfo = "", Null, strsInfo)
            mrsSecdInfo!����Ϣԭֵ = IIf(strsMainInfo = "", Null, strsMainInfo)
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
'���ܣ���������
'    intTYPE: 0-���� 1-��֤ȷ�� 2-ȡ��
'    intModel: 0-���� 1-�޸�
    Dim arrSQL() As Variant
    Dim strCertifySql As String
    Dim strUpdateSql As String
    Dim blnTrans As Boolean
    Dim i As Long, j As Long, k As Long
    Dim strFile As String, strTmp As String, strArrTmp As Variant
    Dim lng֤��ID As Long, lng��� As Long
    Dim strIDs As String, strArry As Variant
    Dim str���֤ As String, strName As String, strSex As String, strAge As String
    Dim strJsonAsk As String, strJsonOut As String, strCertInfo As String
    Dim lng����ID As Long, lngʵ��ID As Long, lng���� As Long, lng����ID As Long, lngID As Long
    Dim blnNew As Boolean, blnCheck As Boolean, blnNotChangeAge As Boolean
    Dim arrInfo  As Variant
    Dim strSQL As String, strBirthdate As String
    Dim rsTmp As ADODB.Recordset
    Dim str���� As String, str���� As String, str�Ա� As String, str�������� As String, strInfo As String, strMsg As String, strExpalin As String
    Dim blnIn As Boolean
    Dim strԭ���� As String, strԭ�Ա� As String, strԭ���� As String, strԭ�������� As String
    Dim str����˵�� As String
    Dim str���ʱ�� As String
    
    On Error GoTo errH
    arrSQL = Array()
    If Not mrsMainInfo Is Nothing Then
        mrsMainInfo.Filter = "��Ϣ��='���֤��'"
        If Not mrsMainInfo.EOF Then str���֤ = mrsMainInfo!��Ϣ��ֵ & ""
        
        mrsMainInfo.Filter = "��Ϣ��='����'"
        If Not mrsMainInfo.EOF Then str���� = mrsMainInfo!��Ϣ��ֵ & ""

        mrsMainInfo.Filter = "��Ϣ��='�Ա�'"
        If Not mrsMainInfo.EOF Then str�Ա� = mrsMainInfo!��Ϣ��ֵ & ""

        If IsDate(txtDateInfo(DATE_��������).Text) Then
            str���� = GetAge(txtDateInfo(DATE_��������).Text, mlng����ID, zlDatabase.Currentdate)
        End If
        
        mrsMainInfo.Filter = "��Ϣ��='��������'"
        If Not mrsMainInfo.EOF Then str�������� = Format(mrsMainInfo!��Ϣ��ֵ & "", "yyyy-mm-dd hh:mm:ss")
    End If
    If intTYPE = 0 Then
        '����ǰ�ļ��
        blnCheck = True
        strCertInfo = GetCertifyData(3, 0, 0, 0)
        '����|�Ա�|����|��������|����������|�������Ա�|�����˳�������|���֤��|���������֤��|�����˹�ϵ|������|֤����Ϣ
        arrInfo = Split(strCertInfo, "|")
        strJsonAsk = "{""input"":{""opr_fun"":" & intModel & "," & IIf(intModel = 1, """real_id"":" & mlngʵ��id & ",", "") & """pati_name"":""" & arrInfo(0) & """,""pati_sex"":""" & arrInfo(1) & """,""pati_age"":""" & arrInfo(2) & """,""pati_birthdate"":""" & arrInfo(3) & """,""pati_idcard"":""" & arrInfo(7) & """,""owner"":" & arrInfo(10) & ",""grdn_name"":""" & arrInfo(4) & """,""grdn_sex"":""" & arrInfo(5) & """,""grdn_birthdate"":""" & arrInfo(6) & """,""grdn_idcard"":""" & arrInfo(8) & """,""grdn_relation"":""" & arrInfo(9) & """,""papers_info"":""" & arrInfo(11) & """}}"
        If Not CallService("Zl_Patisvr_Patirealnamecheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
            blnCheck = False
        Else
            If intModel = 0 Then
                lng����ID = gobjService.GetJsonNodeValue("output.pati_id")
                lngʵ��ID = gobjService.GetJsonNodeValue("output.real_id")
                blnNew = gobjService.GetJsonNodeValue("output.new_pati") = 1
            End If
        End If
        strJsonOut = ""
        If Not blnCheck Then Exit Function
        If Not blnNew Then
            '��ȡ�Һ���Ϣ
            strJsonAsk = "{""input"":{""query_type"":0,""occasion"":0,""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & "}}"
            strJsonOut = ""
            If CallService("Zl_Cissvr_Getpativisitid", strJsonAsk, strJsonOut) Then
                lngID = gobjService.GetJsonNodeValue("output.visit_id")
                lng���� = gobjService.GetJsonNodeValue("output.occasion")
            End If
            '���²��˻�����Ϣ�ļ��
            blnCheck = False
            strSQL = "Select ����, �Ա�, ����, �������� From ������Ϣ Where ����id = [1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���²��˻�����Ϣ�ļ��", IIf(intModel = 0, lng����ID, mlng����ID))
            If Not rsTmp.EOF Then
                strName = rsTmp!���� & ""
                strSex = rsTmp!�Ա� & ""
                strAge = rsTmp!���� & ""
                strBirthdate = Format(rsTmp!�������� & "", "yyyy-mm-dd hh:mm:ss")
                If rsTmp!���� & "" <> str���� Then
                    '������������ж��߼�
                    '���������䲻�ø���
                    If rsTmp!���� & "" Like "*Сʱ%����" Or rsTmp!���� & "" Like "*����" Or rsTmp!���� & "" Like "*��*Сʱ" Or rsTmp!���� & "" Like "*Сʱ" Then
                        blnNotChangeAge = True
                    Else
                        blnNotChangeAge = False
                    End If
                End If
            End If
            If blnNotChangeAge Then
              If strName <> str���� Or strBirthdate <> str�������� Or strSex <> str�Ա� Then
                 blnCheck = True
              End If
              strAge = strAge
            Else
              If strName <> str���� Or strBirthdate <> str�������� Or strSex <> str�Ա� Or strAge <> str���� Then
                blnCheck = True
              End If
              strAge = str����
            End If
            If blnCheck Then
                blnCheck = True
                If lngID = 0 Then
                    strSQL = "Select ����, �Ա�, ����, �������� From ������Ϣ Where ����id = [1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������Ϣ", mlng����ID)
                    If rsTmp.EOF Then
                        MsgBox "����ID[" & lng����ID & "]�ڲ�����Ϣ�в�����,���ܼ������в�����Ϣ�������!", vbInformation, gstrSysName
                        Exit Function
                    Else
                        strԭ���� = rsTmp!���� & ""
                        strԭ���� = rsTmp!���� & ""
                        strԭ�Ա� = rsTmp!�Ա� & ""
                        strԭ�������� = Format(rsTmp!�������� & "", "YYYY-MM-DD HH:MM:SS")
                    End If
                Else
                    strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng���� & ",""pati_name"":""" & strName & """,""pati_age"":""" & strAge & """,""pati_sex"":""" & strSex & """,""pati_birthdate"":""" & strBirthdate & """}}"
                    If Not CallService("Zl_Cissvr_Checkpatexist", strJsonAsk, strJsonOut, , , False, , , , True) Then
                        blnCheck = False
                    Else
                        strԭ���� = gobjService.GetJsonNodeValue("output.pati_name")
                        strԭ�Ա� = gobjService.GetJsonNodeValue("output.pati_sex")
                        strԭ���� = gobjService.GetJsonNodeValue("output.pati_age")
                        strԭ�������� = gobjService.GetJsonNodeValue("output.pati_birthdate")
                    End If
                End If
                strJsonOut = ""
                If Not blnCheck Then Exit Function
                strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & "}}"
                If Not CallService("Zl_Patisvr_Lockcheck", strJsonAsk, strJsonOut, , , False, , , , True) Then
                    blnCheck = False
                End If
                If Not blnCheck Then Exit Function
                strJsonOut = ""
                strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng���� & "}}"
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
        strJsonAsk = "{""input"":{""pati_id"":" & IIf(intTYPE = 0, IIf(intModel = 0, lng����ID, mlng����ID), mlng����ID) & ",""pati_pageid"":" & lngID & "}}"
        If CallService("zl_cissvr_checkpativisitorin", strJsonAsk, strJsonOut, , , False) Then
            blnIn = Val(gobjService.GetJsonNodeValue("output.isexist")) = 1
        End If
    End If
    strJsonOut = ""
    strCertifySql = GetCertifyData(intTYPE, intModel, IIf(intTYPE = 0, IIf(intModel = 0, lng����ID, mlng����ID), mlng����ID), IIf(intTYPE = 0, IIf(intModel = 0, lngʵ��ID, mlngʵ��id), mlngʵ��id), blnNew, IIf(blnNew = True, 0, lngID))
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
        str���ʱ�� = zlDatabase.CallProcedure("Zl_������Ϣ_������Ϣ����_s", "������Ϣ����", IIf(intModel = 0, lng����ID, mlng����ID), 1109, str����, str�Ա�, str����, Format(str��������, "YYYY-MM-DD HH:MM"), strԭ����, strԭ�Ա�, strԭ����, Format(strԭ��������, "YYYY-MM-DD HH:MM"), Empty)
    End If
    strJsonOut = ""
    If intTYPE = 0 Then
        If Not blnNew Then
    '        strInfo = zlDatabase.CallProcedure("Zl_������Ϣ_������Ϣ����", "������Ϣ����", IIf(intModel = 0, lng����ID, mlng����ID), lngID, 1109, str����, str�Ա�, strAge, CDate(str��������), lng����, "����ʵ����Ϣ��֤", strԭ����, str�Ա�, strԭ����, IIf(IsDate(strԭ��������), CDate(strԭ��������), "Null"), Empty)
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng���� & ",""update_info"":{""pati_name"":""" & str���� & """,""pati_age"":""" & str���� & """,""pati_sex"":""" & str�Ա� & """,""pati_birthdate"":""" & str�������� & """}}}"
            If CallService("Zl_Cissvr_Updatepatibaseinfo", strJsonAsk, strJsonOut, , , False, , , , False) Then
                strMsg = gobjService.GetJsonNodeValue("output.adjust_explain")
                If strMsg <> "" Then
                    strInfo = strInfo & strMsg
                End If
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    If str���ʱ�� = "" Then str���ʱ�� = zlDatabase.Currentdate
    strJsonOut = ""
    If intTYPE = 0 Then
        If Not blnNew Then
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & ",""visit_id"":" & lngID & ",""occasion"":" & lng���� & ",""update_info"":{""pati_name"":""" & str���� & """,""pati_age"":""" & strAge & """,""pati_sex"":""" & str�Ա� & """,""pati_birthdate"":""" & str�������� & """,""explain"":""" & strExpalin & """}}}"
            If CallService("Zl_Exsesvr_Updatepatibaseinfo", strJsonAsk, strJsonOut) Then
                strMsg = gobjService.GetJsonNodeValue("output.adjust_explain")
                If strMsg <> "" Then
                    strInfo = strInfo & strMsg
                End If
            End If
            If strInfo <> "" Then
                strInfo = Mid(strInfo, 3)
                strInfo = "�޸�ԭ��:����ʵ����Ϣ��֤" & Chr(13) & "���˻�����Ϣ���������������ݷ����仯:" & Chr(13) & strInfo
            End If
            str����˵�� = strInfo
            strInfo = Replace(strInfo, Chr(13), " ")
            strInfo = Replace(strInfo, Chr(10), " ")
            strJsonOut = ""
'            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����id, mlng����ID) & ",""visit_id"":" & lngID & ",""model"":""" & "ʵ����Ϣ����" & """,""pati_name_n"":""" & str���� & """,""pati_sex_n"":""" & str�Ա� & """,""pati_age_n"":""" & str���� & """,""pati_birthdate_n"":""" & Format(str��������, "YYYY-MM-DD HH:MM:SS") & """,""occasion"":" & lng���� & ",""pati_name_o"":""" & strԭ���� & """,""pati_sex_o"":""" & strԭ�Ա� & """,""pati_age_o"":""" & strԭ���� & """,""pati_birthdate_o"":""" & Format(strԭ��������, "YYYY-MM-DD HH:MM:SS") & """,""explain"":""" & strInfo & """}}"
'            Call CallService("Zl_Patisvr_Updatepatibaseinfo", strJsonAsk, strJsonOut)
            strJsonAsk = "{""input"":{""pati_id"":" & IIf(intModel = 0, lng����ID, mlng����ID) & ",""visit_id"":" & lngID & ",""pati_name"":""" & str���� & """,""pati_sex"":""" & str�Ա� & """,""pati_age"":""" & strAge & """}}"
            Call CallService("Zl_Pivassvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            strJsonOut = ""
            Call CallService("Zl_Drugsvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            strJsonOut = ""
            Call CallService("Zl_Stuffsvr_Patiinfoupdate", strJsonAsk, strJsonOut)
            Call UpdateChangeInfo(IIf(intModel = 0, lng����ID, mlng����ID), strInfo, CDate(str���ʱ��))
        End If
    End If
    If str����˵�� <> "" Then
        mstrMsg = str����˵��
    End If
    If intTYPE = 0 Then
        If intModel = 0 Then
            mlng����ID = lng����ID
            mlngʵ��id = lngʵ��ID
            Call zlRefresh
        End If
        SavePatPicture mlng����ID
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

Private Function UpdateChangeInfo(ByVal lng����ID As Long, ByVal strInfo As String, ByVal d�䶯ʱ�� As Date)
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Zl_������Ϣ�䶯_Update(" & lng����ID & ",'" & strInfo & "'," & zlStr.To_Date(d�䶯ʱ��, "ymdhms") & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SavePatPicture(lng����ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���没����Ƭ
    '���:lng����ID - ����ID
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rs As New Recordset
    Dim strFile As String, strSQL As String
    
    On Error GoTo Errhand
    Select Case mlngͼ�����
        Case 1 '�ļ�
            strFile = cmdialog.filename
        Case 2 '�ɼ�
            strFile = mstr�ɼ�ͼƬ
            mstr�ɼ�ͼƬ = ""
        Case 4 '�������֤
            strFile = App.Path & "\SFZIMG.bmp"
            SavePicture imgPatient.Picture, strFile
    End Select
    If InStr(1, ",1,2,4,", "," & mlngͼ����� & ",") <> 0 Then
        If strFile = "" Then Exit Sub
        Call PictureBoxSaveJPG(imgPatient.Picture, strFile) '����ѹ�����ͼƬ
        If Sys.SaveLob(glngSys, 27, mlng����ID, strFile) = False Then
            MsgBox "������Ƭʧ��,�ļ����ܱ�ɾ��!", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf mlngͼ����� = 3 Then
        strSQL = strSQL & "Zl_������Ƭ_Delete("
        strSQL = strSQL & lng����ID & ")"
        
        zlDatabase.ExecuteProcedure strSQL, "Zl_������Ƭ_Delete"
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SaveCertPicture(ByVal lng֤��ID As Long, ByVal lng��� As Long, ByVal strNote As String, ByVal strFile As String)
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim rsTmp As New ADODB.Recordset
    
    If strFile = "" Then
        Exit Sub
    End If
    If lng��� = 0 Then
        strSQL = "Select max(���) as ��� from ����ʵ��֤��ͼƬ where ֤��ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡͼƬ���", lng֤��ID)
        If rsTmp.EOF Then
            lng��� = 1
        Else
            lng��� = Val("" & rsTmp!���) + 1
        End If
    Else
        lng��� = lng���
    End If
    If Sys.SaveLob(glngSys, 33, lng֤��ID & "|" & lng��� & "|" & strNote, strFile) = False Then
        MsgBox "������Ƭʧ��,�ļ����ܱ�ɾ��!", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Function GetCertifyData(ByVal intTYPE As Integer, ByVal intModel As Integer, ByVal lng����ID As Long, ByVal lngʵ��ID As Long, Optional ByVal bln�²��� As Boolean, Optional ByVal lng��ҳID As Long, Optional ByVal blnIn As Boolean) As String
'���ܣ���ȡ����ʵ����Ϣ��sql
    Dim strValue As String
    Dim arrFilds As Variant
    Dim strSQL As String, strTmp As String
    Dim i As Long
    Dim CurrDate As Date
    
    On Error GoTo errH
    If intTYPE = 0 Then
        If intModel = 0 Then
            arrFilds = Array("ʵ��id", "����id", "�²���", "����", "�Ա�", "����", "��������", "����", "����", "���֤����", "����������", "�������Ա�", "�����˳�������", "�����˹���", "����������", _
                                "���������֤����", "�����ص�", "סַ", "������סַ", "���֤��", "���������֤��", "�����˹�ϵ", "�ֻ���", "��ע", _
                                "��֤״̬", "������", "֤����Ϣ", "�Ƿ�ṹ��", "��ַ��Ϣ", "��ҳid", "�Ƿ����")
            strSQL = "ZL_����ʵ����Ϣ_Insert_S("
        Else
            arrFilds = Array("ʵ��id", "����id", "����", "�Ա�", "����", "��������", "����", "����", "���֤����", "����������", "�������Ա�", "�����˳�������", "�����˹���", "����������", _
                                "���������֤����", "�����ص�", "סַ", "������סַ", "���֤��", "���������֤��", "�����˹�ϵ", "�ֻ���", "��ע", _
                                "��֤״̬", "������", "���ԭ��", "֤����Ϣ", "�Ƿ�ṹ��", "��ַ��Ϣ", "��ҳid", "�Ƿ����")
            strSQL = "ZL_����ʵ����Ϣ_Update_S("
        End If
        For i = LBound(arrFilds) To UBound(arrFilds)
            strValue = ""
            Select Case Trim(arrFilds(i))
                Case ""
                    strValue = ",Null"
                Case "��������", "�����˳�������"
                    mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!��Ϣ��ֵ & ""
                    strValue = "," & zlStr.To_Date(strValue, "ymdhm")
                Case "֤����Ϣ"
                     strValue = ",'" & CachCertData & "'"
                Case "��֤״̬"
                    strValue = IIf(mblnIdentifySure, ",1", ",0")
                Case "������"
                    strValue = IIf(optType(0).Value, ",1", ",2")
                Case "�Ƿ�ṹ��"
                    strValue = IIf(gbln���ýṹ����ַ, ",1", ",0")
                Case "��ַ��Ϣ"
                    strValue = ",'" & GetPatiAdresInfo & "'"
                Case "���֤����", "���������֤����"
                    mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!��Ϣ��ֵ & ""
                    If strValue <> "" Then
                        strValue = "," & cbo.FindIndex(cboInfo(mrsMainInfo!Index), strValue)
                    Else
                        strValue = ",Null"
                    End If
                Case "ʵ��id"
                    strValue = "," & lngʵ��ID
                Case "����id"
                    strValue = "," & lng����ID
                Case "��ҳid"
                    If lng��ҳID <> 0 Then
                        strValue = "," & lng��ҳID
                    Else '
                        strValue = ",Null"
                    End If
                Case "�²���"
                    strValue = "," & IIf(bln�²��� = True, 1, 0)
                Case "�Ƿ����"
                    strValue = "," & IIf(blnIn = True, 1, 0)
                Case "����"
                    If IsDate(txtDateInfo(DATE_��������).Text) Then
                        mstrAge = GetAge(txtDateInfo(DATE_��������).Text, mlng����ID, CurrDate)
                    End If
                    strValue = ",'" & mstrAge & "'"
                Case "���ԭ��"
                    strValue = ",'" & mstrReason & "'"
                Case Else
                    mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!��Ϣ��ֵ & ""
                    strValue = IIf(strValue = "", ",Null", ",'" & strValue & "'")
            End Select
            If i = UBound(arrFilds) Then
                strValue = IIf(strValue = "", "Null", strValue) & ")"
            End If
            strTmp = strTmp & strValue
        Next
    ElseIf intTYPE = 1 Then
        strValue = "," & lngʵ��ID & "," & lng����ID & ",1" & ")"
        strSQL = "Zl_����ʵ����Ϣ_״̬_Update(0,"
        strTmp = strValue
    ElseIf intTYPE = 2 Then
        strValue = "," & lngʵ��ID & "," & lng����ID & ",0" & ")"
        strSQL = "Zl_����ʵ����Ϣ_״̬_Update(0,"
        strTmp = strValue
    ElseIf intTYPE = 3 Then
        arrFilds = Array("����", "�Ա�", "����", "��������", "����������", "�������Ա�", "�����˳�������", "���֤��", "���������֤��", "�����˹�ϵ", "������", "֤����Ϣ")
        For i = LBound(arrFilds) To UBound(arrFilds)
            strValue = ""
            Select Case Trim(arrFilds(i))
                Case "��������", "�����˳�������"
                    mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!��Ϣ��ֵ & ""
                    strValue = "|" & Format(strValue, "yyyy-mm-dd hh:mm")
                Case "֤����Ϣ"
                     strValue = "|" & CachCertData
                Case "������"
                    strValue = IIf(optType(0).Value, "|1", "|2")
                Case "����"
                    If IsDate(txtDateInfo(DATE_��������).Text) Then
                        mstrAge = GetAge(txtDateInfo(DATE_��������).Text, mlng����ID, CurrDate)
                    End If
                    strValue = "|" & mstrAge
                Case Else
                    mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                    strValue = mrsMainInfo!��Ϣ��ֵ & ""
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

Private Function GetAge(ByVal DateBir As Date, Optional ByVal lng����ID As Long, Optional ByVal datCalc As Date) As String
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strTmp As String
    lng����ID = 0
    strSQL = "Select Zl_Age_Calc([1],[2],Null) old From Dual"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng����ID, DateBir, datCalc)
    If Not rsTmp.EOF Then
        GetAge = "" & rsTmp!old
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUpdateData() As String
'���ܣ���ȡ����ʵ����Ϣ�䶯��¼��SQL
    Dim strValue As String
    Dim arrFilds As Variant
    Dim strSQL As String, strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    arrFilds = Array("����", "�Ա�", "��������", "���֤��", "����", "����", "�����ص�", "סַ", "���֤����", "����������", "�������Ա�", "�����˳�������", "���������֤��", _
                        "���������֤����", "�����˹�ϵ", "������סַ", "�����˹���", "����������", "�ֻ���", "��ע", _
                        "���ԭ��")
                        
    strSQL = "Zl_����ʵ����Ϣ_������Ϣ�䶯(" & mlngʵ��id & "," & mlng����ID & ","
    
    For i = LBound(arrFilds) To UBound(arrFilds)
        strValue = ""
        Select Case Trim(arrFilds(i))
            Case ""
                strValue = ",Null"
            Case "��������", "�����˳�������"
                mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!��Ϣ��ֵ & ""
                strValue = "," & zlStr.To_Date(strValue, "ymdhm")
            Case "���֤����", "���������֤����"
                mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!��Ϣ��ֵ & ""
                strValue = "," & cbo.FindIndex(cboInfo(mrsMainInfo!Index), strValue)
            Case "���ԭ��"
                strValue = ",'" & mstrReason & "'"
            Case Else
                mrsMainInfo.Filter = "��Ϣ��='" & Trim(arrFilds(i)) & "'"
                strValue = mrsMainInfo!��Ϣ��ֵ & ""
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
'���ܣ���ȡ���нṹ����ַ��Ϣ���ַ���
    Dim strTmp As String
    Dim i As Long
    Dim intTYPE As Integer
    
    On Error GoTo errH
    For i = patiAdressInfo.LBound To patiAdressInfo.UBound
        If patiAdressInfo(i).Value <> "" Then
            '����\�޸�
            intTYPE = decode(i, 0, 1, 1, 3, 2, 5)
            strTmp = strTmp & "|" & intTYPE & ";" & patiAdressInfo(i).valueʡ & ";" & patiAdressInfo(i).value�� & ";" & patiAdressInfo(i).value���� & ";" & patiAdressInfo(i).value���� & ";" & patiAdressInfo(i).value��ϸ��ַ & ";" & patiAdressInfo(i).Code
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
'���ܣ���ȡ֤����Ϣ
    Dim i As Long
    Dim strCert As String

    On Error GoTo errH
    With vsfCert
        For i = .FixedRows To .Rows - 1
            If (.TextMatrix(i, COL_֤������)) <> "" Then
                strCert = strCert & "," & zlCommFun.GetNeedName(.TextMatrix(i, COL_֤������), "-") & "-" & .TextMatrix(i, COL_֤������) & "-" & .TextMatrix(i, COL_��ע) & "-" & .Cell(flexcpData, i, COL_������, i, COL_������) & "-" & .TextMatrix(i, COL_��ע)
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
'���ܣ�����ǰ�ļ��
    Dim objCtl As Object, objTmp As Object
    Dim strBirthDay As String, strAge As String, strSex As String, strErrInfo As String, strBaseInfo As String
    Dim strTmp As String, str֤������ As String, str���� As String, strInfo As String
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
                strTmp = decode(objCtl.Index, CBO_�Ա�, "�����Ա�", CBO_�������Ա�, "�������Ա�", CBO_����, "���˹���", CBO_�����˹���, "�����˹���", CBO_����, "��������", CBO_����������, "����������", CBO_���֤����, "�������֤����", CBO_���������֤����, "���������֤����", CBO_��ϵ, "�����˹�ϵ")
                If objCtl.Index = CBO_�Ա� Or objCtl.Index = CBO_���� Or objCtl.Index = CBO_���� Then
                    blnShow = True
                ElseIf objCtl.Index = CBO_�������Ա� Or objCtl.Index = CBO_�����˹��� Or objCtl.Index = CBO_���������� Or objCtl.Index = CBO_��ϵ Then
                    If txtInfo(TXT_����������).Text <> "" Then
                        blnShow = True
                    End If
                ElseIf objCtl.Index = CBO_���֤���� Then
                    If txtInfo(TXT_���֤��).Text <> "" Then
                        blnShow = True
                    End If
                    If zlCommFun.GetNeedName(cboInfo(CBO_���֤����).Text, "-") = "����˾���֤" Then
                        If zlCommFun.GetNeedName(cboInfo(CBO_����), "-") = "�й�" Then
                            ShowMessage objCtl, "���˵�֤������Ϊ������˾���֤����Ӧ�Ĺ�������Ϊ���й��������飡", False
                            Exit Function
                        End If
                    End If
                ElseIf objCtl.Index = CBO_���������֤���� Then
                    If txtInfo(TXT_���������֤��).Text <> "" Then
                        blnShow = True
                    End If
                    If zlCommFun.GetNeedName(cboInfo(CBO_���������֤����).Text, "-") = "����˾���֤" Then
                        If zlCommFun.GetNeedName(cboInfo(CBO_�����˹���), "-") = "�й�" Then
                            ShowMessage objCtl, "���˵�֤������Ϊ������˾���֤����Ӧ�Ĺ�������Ϊ���й��������飡", False
                            Exit Function
                        End If
                    End If
                End If
                If blnShow Then
                    If Trim(objCtl.Text) = "" Then
                        ShowMessage objCtl, strTmp & "������д��", False
                        Exit Function
                    End If
                End If
                blnShow = False
            Case "txtInfo"
                Select Case objCtl.Index
                    Case TXT_���������֤��, TXT_���֤��
                        If Trim(objCtl.Text) <> "" Then
                            str֤������ = IIf(objCtl.Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_���֤����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_���������֤����).Text, "-"))
                            str���� = IIf(objCtl.Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_�����˹���).Text, "-"))
                            If (str֤������ = "�������֤" Or str֤������ = "�۰�̨��ס֤") And str���� = "�й�" Then
                                If CreatePublicPatient() Then
                                    If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(objCtl.Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                                        If strBirthDay <> Format(IIf(objCtl.Index = TXT_���������֤��, txtDateInfo(DATE_�����˳�������).Text, txtDateInfo(DATE_��������).Text), "YYYY-MM-DD") Then
                                            strBaseInfo = strBaseInfo & "," & "��������"
                                        End If
                                        If strSex <> zlCommFun.GetNeedName(IIf(objCtl.Index = TXT_���������֤��, cboInfo(CBO_�������Ա�).Text, cboInfo(CBO_�Ա�).Text), "-") Then
                                            strBaseInfo = strBaseInfo & "," & "�Ա�"
                                        End If
                                        If Format(strBirthDay, "HH:MM") = "00:00" Then
                                            strMask = "####-##-##"
                                        Else
                                            strMask = "####-##-## ##:##"
                                        End If
                                        strBaseInfo = Mid(strBaseInfo, 2)
                                        If strBaseInfo <> "" Then
                                            If objCtl.Index = TXT_���֤�� Then
                                                If MsgBox("�������֤�ŷ��ص�" & strBaseInfo & "��ʵ����д��" & strBaseInfo & "������,�Ƿ������������Ὣ������¼���" & strBaseInfo & "�滻�����֤�ŷ��ص�" & strBaseInfo & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    txtDateInfo(DATE_��������).Mask = strMask
                                                    txtDateInfo(DATE_��������).Tag = strMask
                                                    txtDateInfo(DATE_��������).Text = Format(strBirthDay, decode(strMask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                                                    cboInfo(CBO_�Ա�).ListIndex = cbo.FindIndex(cboInfo(CBO_�Ա�), strSex)
                                                End If
                                            Else
                                                If MsgBox("���������֤�ŷ��ص�" & strBaseInfo & "��ʵ����д��" & strBaseInfo & "������,�Ƿ��滻��������Ὣ������¼���" & strBaseInfo & "�滻�����֤�ŷ��ص�" & strBaseInfo & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                                    txtDateInfo(DATE_�����˳�������).Mask = strMask
                                                    txtDateInfo(DATE_�����˳�������).Tag = strMask
                                                    txtDateInfo(DATE_�����˳�������).Text = Format(strBirthDay, decode(strMask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                                                    cboInfo(CBO_�������Ա�).ListIndex = cbo.FindIndex(cboInfo(CBO_�������Ա�), strSex)
                                                End If
                                            End If
                                        End If
                                        If objCtl.Index = TXT_���֤�� Then
                                            If IsDate(txtDateInfo(DATE_��������).Text) Then
                                                mstrAge = GetAge(txtDateInfo(DATE_��������).Text, mlng����ID, CurrDate)
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
                    Case txt_�ֻ���
                        If Not CheckPhoneNumber(Trim(objCtl.Text)) Then Exit Function
                    Case TXT_����
                        If Trim(objCtl.Text) = "" Then
                            ShowMessage objCtl, "����¼�벡��������", False
                            Exit Function
                        End If
                    Case TXT_����������
                        If Trim(objCtl.Text) = "" Then
                            If CheckPPatiInfo Then
                                ShowMessage objCtl, "������������û��¼�룬���飡", False
                                Exit Function
                            End If
                        End If
                End Select
            Case "txtAdressInfo"
                strTmp = decode(objCtl.Index, ADRS_�����ص�, "���˳����ص�", ADRS_סַ, "����סַ", ADRS_������סַ, "������סַ")
                If gbln���ýṹ����ַ Then    '��Ҫ����ַ�ؼ�������
                    If patiAdressInfo(objCtl.Index).CheckNullValue() <> "" Then
                        Call ShowMessage(patiAdressInfo(objCtl.Index), strTmp & "��" & patiAdressInfo(objCtl.Index).CheckNullValue() & "��δ���룬���顣", False)
                        Exit Function
'                    ElseIf patiAdressInfo(objCtl.Index).Value = "" Then
'                        Call ShowMessage(patiAdressInfo(objCtl.Index), "����¼��" & strTmp & "��", False)
'                        Exit Function
                    End If
                    If patiAdressInfo(objCtl.Index).MaxLength > 0 Then
                        If zlCommFun.ActualLen(patiAdressInfo(objCtl.Index).Value) > patiAdressInfo(objCtl.Index).MaxLength Then
                            Call ShowMessage(patiAdressInfo(objCtl.Index), strTmp & "������̫�������顣(����Ŀ������� " & patiAdressInfo(objCtl.Index).MaxLength & " ���ַ��� " & patiAdressInfo(objCtl.Index).MaxLength \ 2 & " ������)", False)
                            Exit Function
                        End If
                    End If
                Else '��Ҫ���TextBox������
                    If objCtl.MaxLength <> 0 And objCtl.Text <> "" Then
                        If zlCommFun.ActualLen(objCtl.Text) > objCtl.MaxLength Then
                            Call ShowMessage(objCtl, strTmp & "�����ݹ��������顣(����Ŀ������� " & objCtl.MaxLength & " ���ַ��� " & objCtl.MaxLength \ 2 & " ������)", False)
                            Exit Function
                        End If
                    End If
                End If
            Case "txtDateInfo"
                strTmp = decode(objCtl.Index, DATE_��������, "��������", DATE_�����˳�������, "�����˳�������")
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
                    If objCtl.Index = DATE_�������� Then
                        If objCtl.Text = "____-__-__" Or objCtl.Text = "____-__-__ __:__" And txtInfo(TXT_����).Text <> "" Then
                            Call ShowMessage(txtDateInfo(objCtl.Index), "�����벡�˵�" & strTmp & "!", False)
                            Exit Function
                        Else
                            If txtInfo(TXT_����).Text <> "" Then
                                Call ShowMessage(txtDateInfo(objCtl.Index), strTmp & "������Ч�����ڸ�ʽ��", False)
                                Exit Function
                            End If
                        End If
                    ElseIf objCtl.Index = DATE_�����˳������� Then
                        If (objCtl.Text = "____-__-__" Or objCtl.Text = "____-__-__ __:__") And Trim(txtInfo(TXT_����������).Text) <> "" Then
                            Call ShowMessage(txtDateInfo(objCtl.Index), "������" & strTmp & "!", False)
                            Exit Function
                        Else
                            If txtInfo(TXT_����������).Text <> "" Then
                                Call ShowMessage(txtDateInfo(objCtl.Index), strTmp & "������Ч�����ڸ�ʽ��", False)
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Case "vsfCert"
                With vsfCert
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, COL_֤������) <> "" Then
                            If .TextMatrix(i, COL_֤������) = "" Then
                                Call ShowMessage(vsfCert, "��ѡ��֤�����ͣ�", False)
                                Exit Function
                            End If
                        End If
                        strInfo = strInfo & "," & .TextMatrix(i, COL_֤������)
                        For j = i + 1 To .Rows - 1
                            If .TextMatrix(j, COL_֤������) <> "" Then
                                If .TextMatrix(j, COL_֤������) = .TextMatrix(i, COL_֤������) And zlCommFun.GetNeedName(.TextMatrix(j, COL_֤������), "-") = zlCommFun.GetNeedName(.TextMatrix(i, COL_֤������), "-") Then
                                    .Row = i: .Col = COL_֤������
                                    Call ShowMessage(vsfCert, "���ظ���֤����Ϣ�����飡", False)
                                    Exit Function
                                End If
                            End If
                        Next
                    Next
                End With
                If Mid(strInfo, 2) = "" Then
                    If txtInfo(TXT_���֤��).Text = "" And txtInfo(TXT_���������֤��) = "" Then
                        Call ShowMessage(txtInfo(TXT_���֤��), "�������֤����������֤�������������֤������������֤������¼��һ�������飡", False)
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
'���ܣ�cmdAdressInfo_Click
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim objTXTBox As TextBox
    Dim bytStyle As Byte, strCaption As String, strMsg As String, blnRoot As Boolean, blnNonWin As Boolean

    On Error GoTo errH
    Select Case Index
        Case ADRS_�����ص�, ADRS_סַ, ADRS_������סַ
            'ѡ���������
            strSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            strCaption = "����": strMsg = "�ֵ������": bytStyle = 0: blnRoot = False: blnNonWin = True
    End Select

    '���ݴ���
    On Error GoTo errH
    '���ݴ���
    Set objTXTBox = txtAdressInfo(Index)
    vPoint = GetCoordPos(objTXTBox.Container.hwnd, objTXTBox.Left, objTXTBox.Top)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, bytStyle, strCaption, , , , , blnRoot, blnNonWin, vPoint.X, vPoint.Y, objTXTBox.Height, blnCancel)

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û������""" & IIf(strCaption = "����", "����", strCaption) & """���ݣ����ȵ�" & strMsg & "�����á�", vbInformation, gstrSysName
        End If
        objTXTBox.Tag = ""
        zlControl.ControlSetFocus objTXTBox
    Else
        objTXTBox.Text = rsTmp!����
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
    Dim objmonInfo As MonthView  '������ÿؼ�����
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
        Case DATE_��������, DATE_�����˳�������
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
    If objCmd.Index = DATE_�������� Then
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
'������
    Call Form_Resize
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
'������
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����֤������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��ʼ���Կ�����
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hwnd)
    End If
    '�򿪶�����
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsbMain.Value
    lngMin = vsbMain.Min
    lngMax = vsbMain.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsbMain.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsbMain.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
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
'        .UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Set cmbMain.Icons = imgManager.Icons
    cmbMain.EnableCustomization False
    
    Call InitBaseInfo
    
    '��ʼ����ַ�ؼ�
    patiAdressInfo(ADRS_�����ص�).Visible = gbln���ýṹ����ַ
    patiAdressInfo(ADRS_סַ).Visible = gbln���ýṹ����ַ
    patiAdressInfo(ADRS_������סַ).Visible = gbln���ýṹ����ַ
    txtAdressInfo(ADRS_�����ص�).Visible = Not gbln���ýṹ����ַ: cmdAdress(ADRS_�����ص�).Visible = Not gbln���ýṹ����ַ
    txtAdressInfo(ADRS_סַ).Visible = Not gbln���ýṹ����ַ: cmdAdress(ADRS_סַ).Visible = Not gbln���ýṹ����ַ
    txtAdressInfo(ADRS_������סַ).Visible = Not gbln���ýṹ����ַ: cmdAdress(ADRS_������סַ).Visible = Not gbln���ýṹ����ַ
    If gbln���ýṹ����ַ Then
        patiAdressInfo(ADRS_�����ص�).ShowTown = gbln��ʾ����
        patiAdressInfo(ADRS_סַ).ShowTown = gbln��ʾ����
        patiAdressInfo(ADRS_������סַ).ShowTown = gbln��ʾ����
    End If
    
    '����
    Call DrawLin
    
    '��ӹ�����
    Call MainDefCommandBar
    
    '��ʼ���б�
    Call InitVsfGridHeader
    
    '��ʼ������
    Call InitCboData
    
    mblnɨ�����֤�Ǽ� = Val(zlDatabase.GetPara("ɨ�����֤�Ǽ�", glngSys, glngModul)) = "1"
    If mintModel = 1 Then
        mblnSave = True
        Screen.MousePointer = 11
        Call LoadPatiInfo(mlngʵ��id)
        Call LoadPatiPricture(mlng����ID, imgLoad, strFile)
        If strFile <> "" Then
            mstr�ɼ�ͼƬ = App.Path & "/pati"
        End If
        If imgLoad.Picture <> 0 Then
            imgPatient.Picture = imgLoad.Picture
        End If
        Call LoadInterface(mlngʵ��id)
        Screen.MousePointer = 0
    Else
        mblnSave = False
        mblnIdentifySure = False
        Screen.MousePointer = 11
        Call LoadPatiInfo(mlngʵ��id)
        Call LoadInterface
        Screen.MousePointer = 0
    End If
    mblnLoadFilish = True
    If mlngʵ��id <> 0 Then
        stbBar.Panels(2).Text = "ʵ��id:" & mlngʵ��id
    End If
    
'    vsbMain.Max = 600
'    vsbMain.Min = 0
'    vsbMain.LargeChange = 100
    
    If Not objFile.FolderExists(App.Path & "\CertImg") Then
        objFile.CreateFolder App.Path & "\CertImg"
    End If
End Sub

Private Function LoadPatiInfo(ByVal lngʵ��ID As Long) As Boolean
    Dim i As Long
    On Error GoTo errH
    Set mrsPati = LoadPatiInfoByID(lngʵ��ID)
    If Not mrsPati.EOF Then
        For i = 0 To mrsPati.Fields.Count - 1
            LoadCache mrsPati.Fields(i).Name, mrsPati.Fields(i).Value & ""
        Next
        mblnIdentifySure = IIf(Val(mrsPati!��֤״̬ & "") = 0, False, True)
    End If
    Set mrsCert = LoadPatiCert(0, lngʵ��ID)
    If Not mrsCert.EOF Then
        LoadCachCert mrsCert
    Else
        vsfCert.TextMatrix(vsfCert.FixedRows, COL_������) = IIf(optType(0).Value, "���˱���", "������")
        vsfCert.Cell(flexcpData, vsfCert.FixedRows, COL_������, vsfCert.FixedRows, COL_������) = IIf(optType(0).Value, 1, 2)
    End If
    vsfCert.Select vsfCert.FixedRows, COL_֤������
    Call vsfCert_Click
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub LoadCachCert(ByVal rsTmp As ADODB.Recordset)
'���ܣ���֤����Ϣ���ز�����

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
    
     'ɾ��֮ǰ�Ļ���
    mrsSecdInfo.Filter = "�ؼ���='vsfCert'"
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
            .TextMatrix(lngRow, COL_֤��ID) = "" & rsTmp!ID
            .TextMatrix(lngRow, COL_֤������) = "" & rsTmp!֤������
            .TextMatrix(lngRow, COL_֤������) = "" & rsTmp!֤������
            .TextMatrix(lngRow, COL_��ע) = "" & rsTmp!��ע
            .TextMatrix(lngRow, COL_������) = IIf(Val("" & rsTmp!������) = 1, "���˱���", "������")
            .Cell(flexcpData, lngRow, COL_������, lngRow, COL_������) = Val("" & rsTmp!������)
            
            .Cell(flexcpPicture, lngRow, COL_ͼƬ, lngRow, COL_ͼƬ) = imgͼƬ
            .Cell(flexcpPictureAlignment, lngRow, COL_ͼƬ, lngRow, COL_ͼƬ) = 4
            
            .Cell(flexcpPicture, lngRow, COL_����, lngRow, COL_����) = imgAdd
            .Cell(flexcpPictureAlignment, lngRow, COL_����, lngRow, COL_����) = 4
            
            .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
            .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
            
            .RowData(lngRow) = Val(rsTmp!ID & "")
            If Trim(.TextMatrix(lngRow, COL_֤��ID)) <> "" Then
                Set rsImg = GetCertPicture(Val(.TextMatrix(lngRow, COL_֤��ID)), 0, 1)
                If Not rsImg.EOF Then
                    With vsfImg
                        j = .Rows - 1
                        For k = 0 To rsImg.RecordCount - 1
                            .AddItem "" & lngRow & "-" & j, j
                            .TextMatrix(j, IMG_֤��ID) = "" & i
                            .Cell(flexcpData, j, IMG_֤��ID, j, IMG_֤��ID) = "" & lngRow & "-" & j
                            
                            .TextMatrix(j, IMG_���) = "" & rsImg!���
                            .RowData(j) = Val(vsfCert.TextMatrix(lngRow, COL_֤��ID))
                            
                            .TextMatrix(j, IMG_��ע) = "" & rsImg!��ע
                        
                            .Cell(flexcpPicture, j, IMG_ͼƬ, j, IMG_ͼƬ) = ImgCert
                            .Cell(flexcpPictureAlignment, j, IMG_ͼƬ, j, IMG_ͼƬ) = 4
                            
                            .Cell(flexcpPicture, j, IMG_Del, j, IMG_Del) = imgDelete
                            .Cell(flexcpAlignment, j, IMG_Del, j, IMG_Del) = 4
                            
                            strsMainInfo = .RowData(j) & "|" & .Cell(flexcpData, j, IMG_ͼƬ, j, IMG_ͼƬ) & "|" & .TextMatrix(j, IMG_���) & "|" & .TextMatrix(j, IMG_��ע)
                            strsInfo = .RowData(j) & "|" & .Cell(flexcpData, j, IMG_ͼƬ, j, IMG_ͼƬ) & "|" & .TextMatrix(j, IMG_���) & "|" & .TextMatrix(j, IMG_��ע)
                            mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ"), Array(lngsTmp, Val(.RowData(j)), "vsfImg", strsInfo, strsMainInfo)
                            j = j + 1
                            lngsTmp = lngsTmp + 1
                            strFile = ""
                            rsImg.MoveNext
                        Next
                    End With
                End If
            End If
                
            strMainInfo = rsTmp!ID & "|" & rsTmp!֤������ & "|" & rsTmp!֤������
            strInfo = strMainInfo & "|" & rsTmp!��ע & "|" & IIf(Nvl("" & rsTmp!������, "") = "1", "���˱���", "������")
            mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ"), Array(lngTmp, Val(rsTmp!ID & ""), "vsfCert", strInfo, strMainInfo)
            lngTmp = lngTmp + 1
            rsTmp.MoveNext
        Next
        .Row = 1: .Col = COL_֤������
        If .TextMatrix(1, COL_������) <> "" Then
            If .TextMatrix(1, COL_������) = "���˱���" Then
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
'���ܣ������ӿ���Ϣ���ز�����

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
                    If Val("" & rsTmp!ID) = Val(.TextMatrix(j, COLS_�ӿ�ID)) Then
                        lngCount = lngCount + 1
                    End If
                Next
                If lngCount = 0 Then
                    .AddItem "", i
                    .TextMatrix(i, COLS_�ӿ�ID) = "" & rsTmp!ID
                    .TextMatrix(i, COLS_����) = "" & rsTmp!�ӿ���
                    .TextMatrix(i, COLS_������) = "" & rsTmp!������
                    .TextMatrix(i, COLS_˵��) = "" & rsTmp!˵��
                    .TextMatrix(i, COLS_��֤���) = decode("" & rsTmp!��֤���, "0", "��֤ʧ��", "1", "��֤�ɹ�", "2", "δ��֤", "" & rsTmp!��֤���)
                    
                    .Cell(flexcpPicture, i, COLS_��֤) = imgIdentify
                    .Cell(flexcpPictureAlignment, lngRow, COL_ͼƬ, lngRow, COL_ͼƬ) = 4
                    
                    If Val("" & rsTmp!��֤���) = 0 Then
                        mblnInterface = False
                    Else
                        mblnInterface = True
                    End If
                    i = i + 1
                End If
                lngCount = 0
                rsTmp.MoveNext
            Loop
            '���ݻ���
            lngTmp = 1
            strTmp = ""
            arrMain = Array(COLS_�ӿ�ID, COLS_����, COLS_˵��, COLS_��֤���, COLS_��֤)
            arrWhole = Array(COLS_����, COLS_˵��, COLS_��֤���, COLS_��֤)
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, COLS_��֤���) <> "" Then
                    If strTmp <> .TextMatrix(i, COLS_��֤���) Then
                        j = 1: strTmp = .TextMatrix(i, COLS_��֤���)
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
                    mrsSecdInfo.AddNew Array("���", "ԭID", "�ؼ���", "��Ϣԭֵ", "����Ϣԭֵ", "Tag", "��Ϣ��ֵ", "����Ϣ��ֵ"), Array(lngTmp, Val(.RowData(i)), vsfInterface.Name, strInfo, strMainInfo, decode(Val(.TextMatrix(i, COLS_��֤���)), "��֤ʧ��", 0, "��֤�ɹ�", 1, "δ��֤"), Null, Null)
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

Private Function LoadInterface(Optional ByVal lngʵ��ID As Long) As Boolean
'���������ӿ���Ϣ
    On Error GoTo errH
    Set mrsIneterface = LoadCertInterface(0, lngʵ��ID)
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
'���ܣ���ʼ��������Ϣ
    Dim objCtl As Object
    Dim intIndex As Integer
    Dim str�ؼ��� As String
    Dim i As Long, j As Long
    Dim strTmp As String
    Dim strFMT As String
    
    On Error GoTo errH
    mrsMainInfo.Filter = "��Ϣ��='" & strName & "'"
    If Not mrsMainInfo.EOF Then
        str�ؼ��� = mrsMainInfo!�ؼ��� & ""
        With Me.Controls
            Select Case str�ؼ���
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
                    If gbln���ýṹ����ַ Then
                        If mlng����ID <> 0 Then
                            Call SetStructAddress(mlng����ID, 0, patiAdressInfo(mrsMainInfo!Index), decode(mrsMainInfo!Index, ADRS_�����ص�, 1, ADRS_סַ, 3, ADRS_������סַ, 5))
                        End If
                    End If
                Case "txtAdressInfo"
                    If Not gbln���ýṹ����ַ Then
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
            mrsMainInfo.Update "��Ϣԭֵ", strValue
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
'���ܣ������ҳ�ؼ���ֵ�Ƿ����仯
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
        mblnChange���� = True
    Else
        mblnInfoChange = True
        mblnSave = False
        If TypeName(objTmp) <> "VSFlexGrid" Then
            mblnChange���� = True
        End If
        Exit Function
    End If
    
    If strCboName = "cboInfo" Then
        mrsMainInfo.Filter = "�ؼ���='" & strCboName & "'" & "And Index=" & lngIndex
        If Not mrsMainInfo.EOF Then
            strOlsInfo = Nvl(mrsMainInfo!��Ϣԭֵ)
            blnFind = True
        Else
            mrsSecdInfo.Filter = "�ؼ���='" & strCboName & "'" & "And IndexEx=" & lngIndex
            If Not mrsSecdInfo.EOF Then
                strOlsInfo = Nvl(mrsSecdInfo!��Ϣԭֵ)
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
        MsgBox "����������֤�ӿڲ���(" & strDllName & ".clsIdentityCert)ʧ��!", vbInformation, gstrSysName
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
        If MsgBox("�ò��˵�ʵ����Ϣ��û�б���,�Ƿ�����˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True
            Exit Sub
        End If
    Else
        Unload Me
    End If
    ClearValue
    If objFile.FileExists(mstr�ɼ�ͼƬ) Then
        Kill mstr�ɼ�ͼƬ
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
        MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
        Exit Sub
    End If
    Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
    
    If gobjPublicPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
    If strPictureFile <> "" Then
        objFile.CopyFile strPictureFile, App.Path & "\Person.bmp"
        strPictureFile = App.Path & "\Person.bmp"
        Set imgPatient.Picture = LoadPicture(strPictureFile)
        picPicture.Tag = strPictureFile
        mstr�ɼ�ͼƬ = strPictureFile
        mlngͼ����� = 2
    End If
    CheckValueChange imgPatient
End Sub

Private Sub lblDelete_Click()
    mlngͼ����� = 3
    imgPatient.Picture = imgDefual.Picture
    CheckValueChange imgPatient
End Sub

Private Sub lblFile_Click()
'�����:74421
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
    mlngͼ����� = 1
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
    Dim str���� As String
    Dim strFMT As String
    
    If Me.ActiveControl Is txtInfo(TXT_���֤��) Then
        If txtInfo(TXT_���֤��).Text = "" Then
            txtInfo(TXT_���֤��).Text = strID
            txtInfo(TXT_����).Text = "": txtInfo(TXT_����).PasswordChar = ""
            txtInfo(TXT_����).IMEMode = 0
            txtInfo(TXT_����).Text = strName
            Call cbo.Locate(cboInfo(CBO_�Ա�), strSex)
            Call cbo.Locate(cboInfo(CBO_����), strNation)
            If Format(datBirthDay, "HH:MM") = "00:00" Then
               strFMT = "####-##-##"
            Else
                strFMT = "####-##-## ##:##"
            End If
            txtDateInfo(DATE_��������).Mask = strFMT
            txtDateInfo(DATE_��������) = Format(datBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
            txtInfo(TXT_���֤��).Text = strID
            Call LoadIDImage
            txtAdressInfo(ADRS_�����ص�).Text = strAddress
            If gbln���ýṹ����ַ Then
                patiAdressInfo(ADRS_�����ص�).Value = strAddress
            End If
            Call cbo.Locate(cboInfo(CBO_���֤����), "�������֤")
        End If
    ElseIf Me.ActiveControl Is txtInfo(TXT_���������֤��) Then
        If txtInfo(TXT_���������֤��).Text = "" Then
            txtInfo(TXT_���������֤��).Text = strID
            txtInfo(TXT_����������).Text = "": txtInfo(TXT_����������).PasswordChar = ""
            txtInfo(TXT_����������).IMEMode = 0
            txtInfo(TXT_����������).Text = strName
            Call cbo.Locate(cboInfo(CBO_�������Ա�), strSex)
            Call cbo.Locate(cboInfo(CBO_����������), strNation)
            If Format(datBirthDay, "HH:MM") = "00:00" Then
               strFMT = "####-##-##"
            Else
                strFMT = "####-##-## ##:##"
            End If
            txtDateInfo(DATE_�����˳�������).Mask = strFMT
            txtDateInfo(DATE_�����˳�������) = Format(datBirthDay, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
            txtInfo(TXT_���������֤��).Text = strID
            txtAdressInfo(ADRS_������סַ).Text = strAddress
            If gbln���ýṹ����ַ Then
                patiAdressInfo(ADRS_������סַ).Value = strAddress
            End If
            Call cbo.Locate(cboInfo(CBO_���������֤����), "�������֤")
        End If
    End If
End Sub

Private Sub LoadIDImage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������֤ͼ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Dim objStdPic As StdPicture
    
    If mobjIDCard Is Nothing Then Exit Sub
    Screen.MousePointer = 11
    Call mobjIDCard.GetPhotoAsStdPicture(objStdPic)
    imgPatient.Picture = objStdPic
    Screen.MousePointer = 0
    mlngͼ����� = 4
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub monInfo_DateClick(ByVal DateClicked As Date)
'���ܣ�monInfo_DateClick
    Dim strDate As String, strFMT As String
    Dim objMSK As MaskEdBox

    Set objMSK = txtDateInfo(mintDate)
    '��ȡʱ��������
    If objMSK.MaxLength >= Len("####-##-## ##:##") Then
        'yyyy-MM-dd HH:mm:ss ��ʽʱ��
        If objMSK.MaxLength > Len("####-##-## ##:##") Then
            strFMT = "HH:mm:ss"
        Else
            'yyyy-MM-dd HH:mm ��ʽʱ��
            strFMT = "HH:mm"
        End If
        'ԭʱ����ʱ�����ͣ���ȡ��ʱ���ʱ�������ݣ�����ȡ��ǰʱ���ʱ����
        If IsDate(objMSK.Text) Then
            strDate = " " & Format(objMSK.Text, strFMT)
        Else
            strDate = " " & Format(zlDatabase.Currentdate, strFMT)
        End If
    End If
    '��ȡʱ��
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
'���ܣ�MskDateInfo_GotFocus
    zlCommFun.OpenIme False
End Sub

Private Sub optType_Click(Index As Integer)
    Dim i As Long
    With vsfCert
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, COL_������) = IIf(Index = 0, "���˱���", "������")
            .Cell(flexcpData, i, COL_������, i, COL_������) = IIf(Index = 0, 1, 2)
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
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '�˵�����
    '-----------------------------------------------------
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Save, "����", -1, False)
    objControl.IconId = 1
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_IdentifySure, "��֤ȷ��", -1, False)
    objControl.IconId = 2
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Cancel, "ȡ����֤", -1, False)
    objControl.IconId = 3
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_Certify_Quit, "�˳�", -1, False)
    objControl.IconId = 4
    objControl.BeginGroup = True
    Set objControl = cmbMain.ActiveMenuBar.Controls.Add(xtpControlButton, conMenu_CertifyHelp_Help, "����", -1, False)
    objControl.IconId = 5
    
    For Each objControl In cmbMain.ActiveMenuBar.Controls
      If objControl.type = xtpControlButton Then
          objControl.Style = xtpButtonIconAndCaption
      End If
    Next
        
End Sub

Private Sub InitBaseInfo()
    Dim arrMainFileds() As Variant

    '��ʼ����¼��
    '1������¼�ṹ����
    Set mrsMainInfo = New ADODB.Recordset
    With mrsMainInfo
        .Fields.Append "���", adInteger, , adFldKeyColumn              '��������ʶ��Ϣ
        .Fields.Append "��Ϣ��", adVarChar, 100, adFldKeyColumn   '��Ϣ����
        '�ü�¼������¼һ����Ϣ��Ӧһ���ؼ������������Ϣ��Ӧһ���ؼ��������������д
        .Fields.Append "�ؼ���", adVarChar, 100, adFldIsNullable      'չʾ��Ϣ�Ŀؼ�����
        .Fields.Append "Index", adInteger, , adFldIsNullable                'Ϊ��ʱ��ʾ���ǿؼ�����
        .Fields.Append "ExpState", adInteger                                        '��Ϣ��չ״̬��0-����չ��1-��ʼ��չ��2-������չ
        .Fields.Append "ҳ��", adInteger                                                '��Ϣ���ڵ�ҳ��
        .Fields.Append "��Ϣԭֵ", adVarChar, 2000, adFldIsNullable  '��Ϣ����ҳ����ʱ��ֵ
        .Fields.Append "��Ϣ��ֵ", adVarChar, 2000, adFldIsNullable  '��Ϣ����ҳ���ʱ��ֵ
        .Fields.Append "ErrInfo", adVarChar, 4000, adFldIsNullable  '�ؼ�¼����Ϣ���Ϸ���ʾ��Ϣ��
        .Fields.Append "Edit", adInteger                                                 '0-�ɱ༭,1-���ɱ༭��ֻ����չʾ,2-���ɱ༭������
        .Fields.Append "�Ƿ�ı�", adInteger                                          '��Ϣ�Ƿ��иı�0-δ�ı䣬1-�ı���
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    '2���μ���Ϣ��¼���ṹ����
    Set mrsSecdInfo = New ADODB.Recordset
    With mrsSecdInfo
        .Fields.Append "Sort", adInteger                                              '����¼��������
        .Fields.Append "���", adInteger                                              '��ʶ��Ϣ����������¼��
        .Fields.Append "�ؼ���", adVarChar, 100                                       'չʾ��Ϣ�Ŀؼ�����
        .Fields.Append "IndexEx", adInteger, , adFldIsNullable                        '�кŻ�ؼ�����Index
        .Fields.Append "ҳ��", adInteger                                              '��Ϣ���ڵ�ҳ��
        .Fields.Append "ԭID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣԭֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ�ڼ���ʱ��ֵ
        .Fields.Append "����Ϣԭֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ����Ҫ���֣���ʶһ����Ϣ�Ƿ񱻳��׸ı䣬��Ϣ�ڼ���ʱ��ֵ
        .Fields.Append "��ID", adBigInt, , adFldIsNullable
        .Fields.Append "��Ϣ��ֵ", adVarChar, 2000, adFldIsNullable      '��Ϣ�ڼ��ʱ��ֵ
        .Fields.Append "����Ϣ��ֵ", adVarChar, 2000, adFldIsNullable    '��Ϣ�ڼ��ʱ��ֵ
        .Fields.Append "�ı�״̬", adInteger                             '��Ϣ�ı�̶�0-δ�ı䣬1-�μ���Ϣ�ı䣬2-����Ϣ�ı�,3-����,-1��ɾ��
        .Fields.Append "ID", adBigInt, , adFldIsNullable                 '��Ϣ�������ݿ��е�ID,һ������ؼ�ʹ��
        .Fields.Append "Tag", adVarChar, 2000                            '�洢��������
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    With mrsMainInfo
        arrMainFileds = Array("��Ϣ��", "�ؼ���", "Index")
        '������Ϣҳ
        .AddNew arrMainFileds, Array("����", "cboInfo", CBO_����)
        .AddNew arrMainFileds, Array("����", "cboInfo", CBO_����)
        .AddNew arrMainFileds, Array("�Ա�", "cboInfo", CBO_�Ա�)
        .AddNew arrMainFileds, Array("���֤����", "cboInfo", CBO_���֤����)
        .AddNew arrMainFileds, Array("�����˹���", "cboInfo", CBO_�����˹���)
        .AddNew arrMainFileds, Array("����������", "cboInfo", CBO_����������)
        .AddNew arrMainFileds, Array("�������Ա�", "cboInfo", CBO_�������Ա�)
        .AddNew arrMainFileds, Array("���������֤����", "cboInfo", CBO_���������֤����)
        .AddNew arrMainFileds, Array("�����˹�ϵ", "cboInfo", CBO_��ϵ)
        
        .AddNew arrMainFileds, Array("����", "txtInfo", TXT_����)
        .AddNew arrMainFileds, Array("����������", "txtInfo", TXT_����������)
        .AddNew arrMainFileds, Array("���֤��", "txtInfo", TXT_���֤��)
        .AddNew arrMainFileds, Array("���������֤��", "txtInfo", TXT_���������֤��)
        .AddNew arrMainFileds, Array("�ֻ���", "txtInfo", txt_�ֻ���)
        .AddNew arrMainFileds, Array("��ע", "txtInfo", TXT_��ע)
    
        .AddNew arrMainFileds, Array("��������", "txtDateInfo", DATE_��������)
        .AddNew arrMainFileds, Array("�����˳�������", "txtDateInfo", DATE_�����˳�������)
        
        If gbln���ýṹ����ַ Then
            .AddNew arrMainFileds, Array("�����ص�", "patiAdressInfo", ADRS_�����ص�)
            .AddNew arrMainFileds, Array("סַ", "patiAdressInfo", ADRS_סַ)
            .AddNew arrMainFileds, Array("������סַ", "patiAdressInfo", ADRS_������סַ)
        Else
            .AddNew arrMainFileds, Array("�����ص�", "txtAdressInfo", ADRS_�����ص�)
            .AddNew arrMainFileds, Array("סַ", "txtAdressInfo", ADRS_סַ)
            .AddNew arrMainFileds, Array("������סַ", "txtAdressInfo", ADRS_������סַ)
        End If
        
    End With
End Sub

Private Sub InitCboData()
'�������б��������
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    strSQL = _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����' ���� From ���� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����' ���� From ���� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '�Ա�' ���� From �Ա� Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '����ϵ' ���� From ����ϵ  Union ALL" & vbNewLine & _
        "Select RowNum As ID, ����, ����, ����, ȱʡ��־ ȱʡ, '֤������' ���� From ֤������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    Call SetCboFromRec(Array("����", "����", "�Ա�", "����ϵ"), Array(CBO_����, CBO_����, CBO_�Ա�, CBO_��ϵ))
    Call SetCboFromRec(Array("����ϵ"), Array(CBO_��ϵ), " ")
    Call SetCboFromRec(Array("����", "����", "�Ա�"), Array(CBO_�����˹���, CBO_����������, CBO_�������Ա�))

    Call SetCboFromList(Array("", "0-�������֤", "1-�۰�̨��ס֤", "2-����˾���֤"), Array(CBO_���֤����))
    Call SetCboFromList(Array("", "0-�������֤", "1-�۰�̨��ס֤", "2-����˾���֤"), Array(CBO_���������֤����))
    
    If cboInfo(CBO_���֤����).ListCount > 0 Then
        cboInfo(CBO_���֤����).ListIndex = 0
    End If
    If cboInfo(CBO_���������֤����).ListCount > 0 Then
        cboInfo(CBO_���������֤����).ListIndex = 0
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
                If IsNull(rsTmp!����) Then
                    objCboTmp.AddItem rsTmp!����
                Else
                    objCboTmp.AddItem rsTmp!���� & "-" & rsTmp!����
                End If
                objCboTmp.ItemData(objCboTmp.NewIndex) = Nvl(rsTmp!ID, 0)
                If Val(rsTmp!ȱʡ & "") = 1 Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                    objCboTmp.Tag = objCboTmp.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
    Next
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String����
'      arrCboIdx=ComboBox��������,���ComboBoxʱ,װ��������ͬ
'      intDefaut=ȱʡ����
    Dim i As Long, j As Long

    For i = 0 To UBound(arrCboIdx)
        cboInfo(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboInfo(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboInfo(arrCboIdx(i)).ListIndex = intDefault 'ȱʡΪδѡ��
    Next
End Sub

Public Function ShowMe(frmParent As Object, ByVal intModel As Integer, Optional ByRef lng����ID As Long, Optional ByRef lngʵ��ID As Long) As Boolean
    mlng����ID = lng����ID
    mlngʵ��id = lngʵ��ID
    mintModel = intModel
    Set mfrmParent = frmParent
    Me.Show 1, mfrmParent
    lng����ID = mlng����ID
    lngʵ��ID = mlngʵ��id
    mlngʵ��id = 0
    mlng����ID = 0
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
        mstrAge = GetAge(txtDateInfo(Index), mlng����ID, CurrDate)
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
    Dim str֤������ As String
    Dim str���� As String
    Dim CurrDate As Date
    
    CurrDate = zlDatabase.Currentdate
    If (Index = TXT_���֤�� Or Index = TXT_���������֤��) And Trim(txtInfo(Index).Text) <> "" Then
        str֤������ = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_���֤����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_���������֤����).Text, "-"))
        str���� = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_�����˹���).Text, "-"))
        If (str֤������ = "�������֤" Or str֤������ = "�۰�̨��ס֤") And str���� = "�й�" Then
            If mblnLoadFilish Then
                If CreatePublicPatient() Then
                    If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                        If IsDate(strBirthDay) Then
                            intIndex = IIf(Index = TXT_���֤��, DATE_��������, DATE_�����˳�������)
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
                            If Index = TXT_���֤�� Then
                                Call cbo.Locate(cboInfo(CBO_�Ա�), strSex, False)
                                mstrAge = GetAge(strBirthDay, mlng����ID, CurrDate)
                            Else
                                Call cbo.Locate(cboInfo(CBO_�������Ա�), strSex, False)
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
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
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
    If Index = TXT_���֤�� Or Index = TXT_���������֤�� Then
        zlControl.TxtSelAll txtInfo(Index)
        If mblnɨ�����֤�Ǽ� = True Then
            Call OpenIDCard(txtInfo(Index).Text = "")
        End If
    End If
End Sub

Private Function ClearValue()
    Dim objFile As New FileSystemObject
    Dim i As Long
    
    With vsfImg
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpData, i, IMG_ͼƬ) <> "" Then
                If objFile.FileExists(.Cell(flexcpData, i, IMG_ͼƬ)) Then
                    Kill .Cell(flexcpData, i, IMG_ͼƬ)
                End If
            End If
        Next
    End With
    mlng֤��id = 0
    mintModel = 0
    Set mfrmParent = Nothing
    Set mrsPati = Nothing
    Set mrsCert = Nothing
    Set mrsIneterface = Nothing
    mblnChange = False
    mblnɨ�����֤�Ǽ� = False
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
    mlngͼ����� = 0
    mstr�ɼ�ͼƬ = ""
    mlngImage = 0
    mlngPati = 0
    mlngTopVsc = 0
    mblnChange���� = False
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
    Dim str֤������ As String
    Dim str���� As String
    
    If Index = TXT_���֤�� Or Index = TXT_���������֤�� Then
        str֤������ = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_���֤����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_���������֤����).Text, "-"))
        str���� = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_�����˹���).Text, "-"))
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Then
            KeyAscii = 0
        End If
        If KeyAscii = 13 Then
            If Trim(txtInfo(Index).Text) <> "" And (str֤������ = "�������֤" Or str֤������ = "�۰�̨��ס֤") And str���� = "�й�" Then
                If Not CreatePublicPatient Then Exit Sub
                If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                    If IsDate(strBirthDay) Then
                        intIndex = IIf(Index = TXT_���֤��, DATE_��������, DATE_�����˳�������)
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
                        If Index = TXT_���֤�� Then
                            Call cbo.Locate(cboInfo(CBO_�Ա�), strSex, False)
                        Else
                            Call cbo.Locate(cboInfo(CBO_�������Ա�), strSex, False)
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
    ElseIf Index = TXT_���� Or Index = TXT_���������� Then
        If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    ElseIf Index = TXT_��ע Then
        If zlCommFun.ActualLen(txtInfo(TXT_��ע)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
            KeyAscii = 0
        End If
    ElseIf Index = txt_�ֻ��� Then
        If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        If KeyAscii = 13 And Trim(txtInfo(txt_�ֻ���).Text) <> "" Then
            If Not CheckPhoneNumber(Trim(txtInfo(txt_�ֻ���).Text)) Then Exit Sub
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
    Dim str֤������ As String
    Dim str���� As String
    
    If Index = TXT_���֤�� Or Index = TXT_���������֤�� And Trim(txtInfo(Index).Text) <> "" Then
        str֤������ = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_���֤����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_���������֤����).Text, "-"))
        str���� = IIf(Index = TXT_���֤��, zlCommFun.GetNeedName(cboInfo(CBO_����).Text, "-"), zlCommFun.GetNeedName(cboInfo(CBO_�����˹���).Text, "-"))
        If Trim(txtInfo(Index).Text) <> "" And (str֤������ = "�������֤" Or str֤������ = "�۰�̨��ס֤") And str���� = "�й�" Then
            If CreatePublicPatient() Then
                If gobjPublicPatient.CheckPatiIdcard(Trim(txtInfo(Index).Text), strBirthDay, strAge, strSex, strErrInfo) Then
                    If IsDate(strBirthDay) Then
                        intIndex = IIf(Index = TXT_���֤��, DATE_��������, DATE_�����˳�������)
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
                            If Index = TXT_���֤�� Then
                                Call cbo.Locate(cboInfo(CBO_�Ա�), strSex, False)
                            Else
                                Call cbo.Locate(cboInfo(CBO_�������Ա�), strSex, False)
                            End If
                    End If
                Else
                    Call ShowMessage(txtInfo(Index), strErrInfo)
                End If
            End If
        End If
    ElseIf Index = txt_�ֻ��� Then
        If Trim(txtInfo(txt_�ֻ���).Text) <> "" Then
            If Not CheckPhoneNumber(Trim(txtInfo(txt_�ֻ���).Text)) Then Exit Sub
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
        If Col = COL_֤������ Then
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
        If lngNewCol = COL_֤������ Then
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
        ElseIf lngNewCol = COL_Del Or lngNewCol = COL_���� Then
            .ComboList = "..."
            .FocusRect = flexFocusNone
            Set .CellButtonPicture = IIf(lngNewCol = COL_����, imgAdd, imgDelete)
        ElseIf lngNewCol = COL_ͼƬ Then
             .ComboList = "..."
             .FocusRect = flexFocusNone
             Set .CellButtonPicture = imgͼƬ
        Else
            .ComboList = ""
        End If
        If OldCol = COL_֤������ Then
            .TextMatrix(OldRow, OldCol) = zlStr.NeedName(.TextMatrix(OldRow, OldCol))
        End If
        If lngNewRow >= .FixedRows Then
            '��ʾͼƬ
            If lngNewCol <> COL_���� And .TextMatrix(lngNewRow, COL_֤������) <> "" Then
                If .Rows - 1 <> lngNewRow Then
                    '��һ�����Ϊ������������
                    If .TextMatrix(lngNewRow + 1, COL_֤������) = "" Then
                         Set .Cell(flexcpPicture, lngNewRow, COL_����) = imgAdd
                    End If
                Else
                    Set .Cell(flexcpPicture, lngNewRow, COL_����) = imgAdd
                End If
            End If
            '��ʾͼƬ
            If lngNewCol <> COL_Del Then Set .Cell(flexcpPicture, lngNewRow, COL_Del) = imgDelete
            If lngNewCol <> COL_ͼƬ And .TextMatrix(lngNewRow, COL_֤������) <> "" Then Set .Cell(flexcpPicture, lngNewRow, COL_ͼƬ) = imgͼƬ
        End If
    End With
End Sub

Private Sub vsfCert_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngCol As Long
    
    lngCol = Col
    If lngCol = COL_���� Or lngCol = COL_Del Or lngCol = COL_ͼƬ Then Cancel = True
End Sub

Private Sub vsfCert_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, lngCol As Long
    Dim i As Long, j As Long, k As Long, int��� As Long, lngCount As Long, lng֤��ID As Long
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
            Case COL_����
                For i = .Rows - 1 To .FixedRows Step -1
                    If Trim(.TextMatrix(i, COL_֤������)) <> "" And .RowHidden(i) = False Then
                        blnAdd = True
                        Exit For
                    ElseIf Trim(.TextMatrix(i, COL_֤������)) = "" And .RowHidden(i) = False Then
                        Exit For
                    End If
                Next
                If blnAdd = True Then
                     lngRow = .Rows: .AddItem "", lngRow
                     .Row = lngRow: .Col = COL_֤������
                     .TextMatrix(lngRow, COL_������) = IIf(optType(0).Value, "���˱���", "������")
                     .Cell(flexcpData, lngRow, COL_������, lngRow, COL_������) = IIf(optType(0).Value, "1", "2")
                     .ShowCell .Row, COL_֤������
                End If
                lngCounts = 0
            Case COL_Del
                If Trim(.TextMatrix(lngRow, COL_֤������)) <> "" Then
                    If MsgBox("ȷ��Ҫɾ��֤������Ϊ��" & .TextMatrix(lngRow, COL_֤������) & "����֤����Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        If lngRows = .FixedRows Then
                            For i = COL_֤��ID To COL_��ע
                                .TextMatrix(lngRow, i) = ""
                                .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                            Next
                            With vsfImg
                                For j = .FixedRows To .Rows - 1
                                    If Val(.TextMatrix(j, IMG_֤��ID)) = lngRow Then
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
                                    If Val(.TextMatrix(j, IMG_֤��ID)) = lngRow Then
                                        .RemoveItem j
                                        .AddItem "", j
                                        .RowHidden(j) = True
                                    End If
                                Next
                            End With
                            For j = .FixedRows To .Rows - 1
                                If Trim(.TextMatrix(j, COL_֤������)) <> "" Then
                                    .Row = j: .Col = COL_֤������
                                    Call vsfCert_Click
                                End If
                            Next
                            CheckValueChange vsfCert
                        End If
                    Else
                        .Row = lngRow: .Col = COL_֤������
                        .ShowCell .Row, COL_֤������
                    End If
                Else
                    If .Rows - 1 = .FixedRows Or lngRow = .FixedRows Then
                        Exit Sub
                    Else
                        For i = .FixedRows To .Rows - 1
                            If .TextMatrix(i, COL_֤������) <> "" Then
                                lngCounts = lngCounts + 1
                            End If
                        Next
                        If lngCounts <> 0 Then
                            .RemoveItem lngRow
                            For j = .FixedRows To .Rows - 1
                                If j <= .Rows - 1 Then
                                    If Trim(.TextMatrix(j, COL_֤������)) <> "" Then
                                        .Row = j: .Col = COL_֤������
                                        Call vsfCert_Click
                                    End If
                                End If
                            Next
                            CheckValueChange vsfCert
                        End If
                    End If
                End If
            Case COL_ͼƬ
                If .TextMatrix(lngRow, COL_֤������) <> "" Then
                    If gobjPublicPatient Is Nothing Then
                        On Error Resume Next
                        Call CreatePublicPatient
                        Err.Clear: On Error GoTo 0
                    End If
                    If gobjPublicPatient Is Nothing Then
                        MsgBox "����������Ϣ����������zlPublicPatient.clsPublicPatient��ʧ�ܣ�", vbInformation, Me.Caption
                        Exit Sub
                    End If
                    Call gobjPublicPatient.zlInitCommon(gcnOracle, glngSys, UserInfo.�û���)
                    If gobjPublicPatient.PatiImageGatherer(Me, strPictureFile) = False Then Exit Sub
                    If strPictureFile <> "" Then
                        mlngImage = mlngImage + 1
                        objFile.CopyFile strPictureFile, App.Path & "\CertImg\image" & mlngImage & ".bmp"
                        strPictureFile = App.Path & "\CertImg\image" & mlngImage & ".bmp"
                        If ImgCert.Picture <> 0 Then
                            j = vsfImg.Rows - 1
                            For k = vsfImg.FixedRows To vsfImg.Rows - 1
                                If vsfImg.Cell(flexcpData, k, IMG_֤��ID, k, IMG_֤��ID) = lngRow & "-" & k Then
                                    int��� = int��� + 1
                                End If
                            Next
                            With vsfImg
                                .AddItem "" & lngRow & "-" & j, j
                                .Cell(flexcpPicture, j, IMG_ͼƬ, j, IMG_ͼƬ) = ImgCert
                                .Cell(flexcpPictureAlignment, j, IMG_ͼƬ, j, IMG_ͼƬ) = 4
                                
                                .Cell(flexcpPicture, j, IMG_Del, j, IMG_Del) = imgDelete
                                .Cell(flexcpPictureAlignment, j, IMG_Del, j, IMG_Del) = 4
                                
                                .TextMatrix(j, IMG_֤��ID) = "" & lngRow
                                .Cell(flexcpData, j, IMG_֤��ID, j, IMG_֤��ID) = "" & lngRow & "-" & j
                                
                                .TextMatrix(j, IMG_���) = "" & int��� + 1
                                imgPic.Picture = LoadPicture(strPictureFile) '��Ҫѹ����ͼƬ
                                Call PictureBoxSaveJPG(imgPic.Picture, strPictureFile) '����ѹ�����ͼƬ
                                .Cell(flexcpData, j, IMG_ͼƬ, j, IMG_ͼƬ) = strPictureFile
                            End With
                            int��� = 0
                            CheckValueChange vsfImg
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfCert_Click()
    Dim i As Long, j As Long, int��� As Long, lngCount As Long, lng֤��ID As Long
    Dim lngRow As Long, lngCol As Long
    
    With vsfCert
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow <> -1 And lngCol <> -1 Then
            If (lngCol = COL_���� Or lngCol = COL_Del Or lngCol = COL_ͼƬ) And lngRow >= .FixedRows Then
                If lngCol = COL_���� Then
                    If .TextMatrix(lngRow, COL_֤������) = "" Then Exit Sub
                End If
                .Select lngRow, lngCol
                Call vsfCert_CellButtonClick(lngRow, lngCol)
            Else
                With vsfImg
                    For i = .FixedRows To .Rows - 1
                        If .Cell(flexcpData, i, IMG_֤��ID, i, IMG_֤��ID) = "" & lngRow & "-" & i And Val(.TextMatrix(i, IMG_֤��ID)) > 0 Then
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
        If Col = COL_֤������ Then
            '��λ��ƥ����
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
        If Col = COL_֤������ Then
            '��λ��ƥ����
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
            If Trim(.TextMatrix(lngRow, COL_֤������)) <> "" Then
                 lngRow = .Row + 1: .AddItem "", lngRow
                .Row = lngRow: .Col = COL_֤������
                .Cell(flexcpPicture, lngRow, COL_����, lngRow, COL_����) = imgAdd
                .Cell(flexcpPictureAlignment, lngRow, COL_����, lngRow, COL_����) = 4
                .Cell(flexcpPicture, lngRow, COL_Del, lngRow, COL_Del) = imgDelete
                .Cell(flexcpPictureAlignment, lngRow, COL_Del, lngRow, COL_Del) = 4
                .ShowCell .Row, .Col
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngRow = .Row
            If Trim(.TextMatrix(lngRow, COL_֤������)) <> "" Then
                If MsgBox("ȷ��Ҫɾ��֤������Ϊ��" & .TextMatrix(lngRow, COL_֤������) & "����֤����Ϣ��", vbQuestion + vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If .Rows - 1 = .FixedRows Then
                        For i = COL_֤��ID To COL_ͼƬ
                            .TextMatrix(lngRow, i) = ""
                            .Cell(flexcpData, lngRow, i, lngRow, i) = ""
                        Next
                    ElseIf .Rows - 1 > .FixedRows Then
                        .RemoveItem lngRow
                    End If
                Else
                    .Row = lngRow: .Col = COL_֤������
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
                For j = IIf(i = .Row, .Col + 1, COL_֤������) To COL_Del
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
                Case COL_֤������
                    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Or KeyAscii = Asc("*") Then
                        KeyAscii = 0
                    Else
                        intRow = .Row
                        intCol = .Col
                        .ComboList = "" 'ʹ��ť״̬��������״̬
                    End If
                Case COL_ͼƬ, COL_����, COL_Del, COL_������
                    .ComboList = "..."
                Case COL_��ע
                    If zlCommFun.ActualLen(.TextMatrix(.Row, COL_��ע)) >= 100 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
                        KeyAscii = 0
                    End If
                Case COL_֤������
                    If zlCommFun.ActualLen(.TextMatrix(.Row, COL_֤������)) >= 20 And UCase(Chr(KeyAscii)) <> Chr(8) And UCase(Chr(KeyAscii)) <> Chr(13) Then
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
                If Trim(.TextMatrix(lngRow, COL_֤������)) = "" Then
                    If lngCol > COL_֤������ Then Exit Function
                End If
            End With
        ElseIf objVsf.Name = "vsfInterface" Then
            With vsfInterface
                .Editable = flexEDNone
                If .ColHidden(lngCol) Then Exit Function
                If lngCol <> COLS_��֤ Then Exit Function
                If Trim(.TextMatrix(lngRow, COLS_����)) = "" Then Exit Function
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
        If lngCol = COL_֤������ Then
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
        If lngCol = COL_֤������ Then
            .TextMatrix(lngRow, COL_֤������) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_֤������)) >= 20 Then
                MsgBox "֤��������ַ��������ܴ���20���ַ���", vbInformation, gstrSysName
                Cancel = True
            End If
        ElseIf lngCol = COL_��ע Then
            .TextMatrix(lngRow, COL_��ע) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, COL_��ע)) >= 100 Then
                MsgBox "��ע���ַ��������ܴ���100���ַ�����50�����֣�", vbInformation, gstrSysName
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
    Dim lngRow As Long, lngCol As Long, lngCertRow As Long, lng֤��ID As Long, lng��� As Long
    
    With vsfImg
        lngRow = Row
        lngCol = Col
        lng֤��ID = Val(.RowData(lngRow))
        lng��� = Val(.TextMatrix(lngRow, IMG_���))
        lngCertRow = Val(.TextMatrix(lngRow, IMG_֤��ID))
        Select Case lngCol
            Case IMG_Del
                If .Cell(flexcpData, lngRow, IMG_֤��ID, lngRow, IMG_֤��ID) <> "" Then
                    If MsgBox("ȷ��Ҫɾ���ڡ�" & lngRow & "���е�ͼƬ��Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        .RemoveItem lngRow
                        .AddItem "", lngRow
                        .RowHidden(lngRow) = True
                        .Cell(flexcpData, lngRow, IMG_֤��ID, lngRow, IMG_֤��ID) = "" & lngCertRow & "-" & lngRow
                        .TextMatrix(lngRow, IMG_���) = "" & lng���
                        CheckValueChange vsfImg
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfImg_Click()
    Dim lngRow As Long, lngCol As Long
    Dim lng֤��ID As Long, lng��� As Long, lngCertRow As Long
    Dim strFile As String
    Dim vPoint As POINTAPI
    
    vPoint = GetCoordPos(vsfImg.hwnd, vsfImg.CellLeft, vsfImg.CellTop)
    With vsfImg
        lngRow = .MouseRow
        lngCol = .MouseCol
        If lngRow <> -1 And lngCol <> -1 Then
            lng֤��ID = Val(.RowData(lngRow))
            lng��� = Val(.TextMatrix(lngRow, IMG_���))
            lngCertRow = Val(.TextMatrix(lngRow, IMG_֤��ID))
            If lngCol = IMG_ͼƬ And .RowHidden(lngRow) = False Then
                If Trim(.Cell(flexcpData, lngRow, IMG_ͼƬ, lngRow, IMG_ͼƬ)) = "" And lng֤��ID <> 0 And lng��� <> 0 Then
                    frmCertPicture.ShowMe Me, lng֤��ID, 1, vPoint.X, vPoint.Y, vsfImg.Height, lng���
                Else
                    frmCertPicture.ShowMe Me, 0, 2, vPoint.X, vPoint.Y, vsfImg.Height, 0, .Cell(flexcpData, lngRow, IMG_ͼƬ, lngRow, IMG_ͼƬ)
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
        
        If lngCol = IMG_��ע Then
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
        If lngCol = IMG_��ע Then
            .TextMatrix(lngRow, IMG_��ע) = .EditText
            If zlCommFun.ActualLen(.TextMatrix(lngRow, lngCol)) >= 100 Then
                MsgBox "��ע���ַ��������ܴ���100���ַ�����50�����֣�", vbInformation, gstrSysName
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
            If lngNewCol = COLS_��֤ Then
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
    If lngCol = COLS_��֤ Then Cancel = True
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
         If lngCol = COLS_��֤ And Trim(.TextMatrix(lngRow, COLS_����)) <> "" Then
            If mblnInfoChange Then
                If Not CheckCertifyData Then Exit Sub
                Call CachPatiData
                Call CachCertInterface
                If mintModel = 1 And mblnChange���� Then
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
                    mblnChange���� = False
                    mintModel = 1
                    CachAllData
                Else
                    mblnSave = False
                End If
            End If
            blnCreate = CreateIdentifyObj(.TextMatrix(lngRow, COLS_������))
            If blnCreate Then
                If mblnSave Then
                    Call SetParIn(strParIn)
                    If mobjIdentify.IdentityCert(strParIn, strParOut) Then
                        .TextMatrix(lngRow, COLS_��֤���) = "����֤"
                        mblnInterface = True
                        Call SetReturnValue(strParOut)
                    Else
                        .TextMatrix(lngRow, COLS_��֤���) = "��֤ʧ��"
                    End If
                    CheckValueChange vsfInterface
                    Screen.MousePointer = 11
                    gcnOracle.BeginTrans: blnTrans = True
                    strSQL = "Zl_ʵ����֤�ӿ���־_Insert(" & mlngʵ��id & "," & Val(.TextMatrix(lngRow, COLS_�ӿ�ID)) & "," & IIf(optType(0).Value, 1, 2) & "," & IIf(.TextMatrix(lngRow, COLS_��֤���) = "����֤", 1, 0) & ",'" & gstrDBUser & "',To_Date('" & CurrDate & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                    gcnOracle.CommitTrans: blnTrans = False
                    Call SaveInterfaceRecord(lngRow, strParIn, strParOut, CurrDate)
                    Screen.MousePointer = 0
                End If
            Else
                .TextMatrix(lngRow, COLS_��֤���) = "������������ʧ��"
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
            If lngCol = COLS_��֤ Then
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
            txtInfo(TXT_����) = strName
            cboInfo(CBO_�Ա�).ListIndex = cbo.FindIndex(cboInfo(CBO_�Ա�), strSex)
            cboInfo(CBO_����).ListIndex = cbo.FindIndex(cboInfo(CBO_����), strCountry)
            cboInfo(CBO_����).ListIndex = cbo.FindIndex(cboInfo(CBO_����), strNation)
            If gbln���ýṹ����ַ Then
                patiAdressInfo(ADRS_�����ص�).Value = strPlace
                patiAdressInfo(ADRS_סַ).Value = strAdress
                
            Else
                txtAdressInfo(ADRS_�����ص�).Text = strPlace
                txtAdressInfo(ADRS_סַ).Text = strAdress
            End If
            txtInfo(TXT_���֤��) = strIdNumer
            txtInfo(txt_�ֻ���) = strPhone
            If IsDate(strDate) Then
                txtDateInfo(DATE_��������).Mask = strMask
                txtDateInfo(DATE_��������).Tag = strMask
                txtDateInfo(DATE_��������) = strDate
            End If
        Else
            txtInfo(TXT_����������) = strName
            cboInfo(CBO_�������Ա�).ListIndex = cbo.FindIndex(cboInfo(CBO_�Ա�), strSex)
            cboInfo(CBO_�����˹���).ListIndex = cbo.FindIndex(cboInfo(CBO_����), strCountry)
            cboInfo(CBO_����������).ListIndex = cbo.FindIndex(cboInfo(CBO_����), strNation)
            If gbln���ýṹ����ַ Then
                patiAdressInfo(ADRS_������סַ).Value = strAdress
            Else
                txtAdressInfo(ADRS_������סַ).Text = strAdress
            End If
            txtInfo(TXT_���������֤��) = strIdNumer
            txtInfo(txt_�ֻ���) = strPhone
            If IsDate(strDate) Then
                txtDateInfo(DATE_�����˳�������).Mask = strMask
                txtDateInfo(DATE_�����˳�������).Tag = strMask
                txtDateInfo(DATE_�����˳�������) = strDate
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

    If Sys.SaveLob(glngSys, 34, mlngʵ��id & "|" & Val(vsfInterface.TextMatrix(lngRow, COLS_�ӿ�ID)) & "|" & CurrDate & "|0", strParIn, 1) = False Then
        MsgBox "ʵ����֤�ӿ���־����ʧ�ܣ�", vbInformation, gstrSysName
    End If
    If Sys.SaveLob(glngSys, 34, mlngʵ��id & "|" & Val(vsfInterface.TextMatrix(lngRow, COLS_�ӿ�ID)) & "|" & CurrDate & "|1", strParOut, 1) = False Then
        MsgBox "ʵ����֤�ӿ���־����ʧ�ܣ�", vbInformation, gstrSysName
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
    If Trim(txtInfo(TXT_���֤��).Text) <> "" Then
        intTYPE = 0
    Else
        If optType(0).Value = True Then
            If vsfCert.Rows - 1 > vsfCert.FixedRows Then
                intTYPE = 0
            End If
        Else
            If Trim(txtInfo(TXT_���������֤��).Text) <> "" Then
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
    If mstr�ɼ�ͼƬ <> "" Then
        Call PictureBoxSaveJPG(imgPatient.Picture, mstr�ɼ�ͼƬ) '����ѹ�����ͼƬ
        strPatiImg = zlStr.EncodeBase64_File(mstr�ɼ�ͼƬ)
    End If
    If intTYPE = 0 Then
        strTmp = "<cert_id>" & mlngʵ��id & "</cert_id><pati_Id>" & mlng����ID & "</pati_Id><pati_name>" & txtInfo(TXT_����).Text & "</pati_name>" & _
        "<sex>" & zlCommFun.GetNeedName(cboInfo(CBO_�Ա�).Text, "-") & "</sex><birth_date>" & txtDateInfo(DATE_��������) & "</birth_date><country>" & cboInfo(CBO_����) & "</country><nation>" & _
        cboInfo(CBO_����) & "</nation><birth_place>" & IIf(gbln���ýṹ����ַ, patiAdressInfo(ADRS_�����ص�).Value, txtAdressInfo(ADRS_�����ص�)) & "</birth_place>" & _
        "<address>" & IIf(gbln���ýṹ����ַ, patiAdressInfo(ADRS_סַ).Value, txtAdressInfo(ADRS_סַ)) & "</address><id_number>" & txtInfo(TXT_���֤��).Text & "</id_number>" & _
        "<phone_number>" & txtInfo(txt_�ֻ���).Text & "</phone_number><pati_Image>" & strPatiImg & "</pati_Image>"
    Else
        strTmp = "<cert_id>" & mlngʵ��id & "</cert_id><pati_Id>" & mlng����ID & "</pati_Id><pati_name>" & txtInfo(TXT_����������).Text & "</pati_name>" & _
        "<sex>" & zlCommFun.GetNeedName(cboInfo(CBO_�Ա�).Text, "-") & "</sex><birth_date>" & txtDateInfo(DATE_�����˳�������) & "</birth_date><country>" & cboInfo(CBO_�����˹���) & "</country><nation>" & _
        cboInfo(CBO_����������) & "</nation><birth_place></birth_place>" & _
        "<address>" & IIf(gbln���ýṹ����ַ, patiAdressInfo(ADRS_������סַ).Value, txtAdressInfo(ADRS_������סַ)) & "</address><id_number>" & txtInfo(TXT_���������֤��).Text & "</id_number>" & _
        "<phone_number>" & txtInfo(txt_�ֻ���).Text & "</phone_number><pati_Image>" & strPatiImg & "</pati_Image>"
    End If
    With vsfCert
        If .Rows > .FixedRows Then
            For i = .FixedRows To .Rows - 1
                If IIf(Trim(.TextMatrix(i, COL_������)) = "���˱���", 0, 1) = intTYPE Then
                    strCertType = .TextMatrix(i, COL_֤������)
                    strCertNumber = .TextMatrix(i, COL_֤������)
                    strCertPati = IIf(Trim(.TextMatrix(i, COL_������)) = "���˱���", "1", "2")
                    With vsfImg
                        For j = .FixedRows To .Rows - 1
                            If .RowData(j) = vsfCert.TextMatrix(i, COL_֤��ID) Then
                                If .Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ) = "" Then
                                    Call ReadPatPricture(.RowData(j) & "," & Val(.TextMatrix(j, IMG_���)), imgLoad, strFile)
                                    If strFile <> "" Then
                                        Call PictureBoxSaveJPG(imgLoad.Picture, strFile) '����ѹ�����ͼƬ
                                        strPatiImg = zlStr.EncodeBase64_File(strFile)
                                        Kill strFile
                                    End If
                                Else
                                    strPatiImg = zlStr.EncodeBase64_File(.Cell(flexcpData, i, IMG_ͼƬ, i, IMG_ͼƬ))
                                End If
                                strPicture = strPicture & "<IMAGE><IMAGE_CODE>" & strPatiImg & "</IMAGE_CODE><NOTE>" & .TextMatrix(i, IMG_��ע) & "</NOTE></IMAGE>"
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
                For j = IIf(i = .Row, .Col + 1, COLS_����) To COLS_��֤
                    If CertCellEditable(vsfInterface, i, j) Then Exit For
                Next
                If j <= COLS_��֤ Then Exit For
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
                Case COLS_��֤
                    .ComboList = "..."
            End Select
        End If
    End With
End Sub

Private Function CheckPPatiInfo() As Boolean
'���ܣ������������Ϣ
    Dim objCtrl As Object
    Dim blnLocked As Boolean
    Dim i As Long
    Dim bln���� As Boolean
    
    With Me.Controls
        For Each objCtrl In Me.Controls
            Select Case objCtrl.Name
                Case "txtInfo"
                    If objCtrl.Index = TXT_���������֤�� Then
                        bln���� = IIf(objCtrl.Text <> "", True, bln����)
                    End If
                Case "txtAdressInfo"
                    If objCtrl.Index = ADRS_������סַ Then
                        bln���� = IIf(objCtrl.Text <> "", True, bln����)
                    End If
                Case "patiAdressInfo"
                    If objCtrl.Index = ADRS_������סַ Then
                        bln���� = IIf(objCtrl.Value <> "", True, bln����)
                    End If
                Case "txtDateInfo"
                    If objCtrl.Index = DATE_�����˳������� Then
                        If IsDate(objCtrl.Text) Or (objCtrl.Text <> "____-__-__ __:__" And objCtrl.Text <> "____-__-__") Then
                            bln���� = True
                        Else
                            bln���� = bln����
                        End If
                    End If
                Case "txtInfoDate"
'                    If objCtrl.Index = DATE_�����˳������� Then
'                        bln���� = IIf(objCtrl.Text <> "", True, bln����)
'                    End If
                Case "optType"
                    If objCtrl.Index = 1 Then
                        If objCtrl.Value = True Then
                            For i = vsfCert.FixedRows To vsfCert.Rows - 1
                                If vsfCert.TextMatrix(i, COL_֤������) <> "" Then
                                    bln���� = True
                                End If
                            Next
                        End If
                    End If
            End Select
        Next
    End With
    CheckPPatiInfo = bln����
End Function

Private Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80) As Boolean
     Dim tSI As GdiplusStartupInput
     Dim lRes As Long
     Dim lGDIP As Long
     Dim lBitmap As Long
    
     '��ʼ�� GDI+
     tSI.GdiplusVersion = 1
     lRes = GdiplusStartup(lGDIP, tSI, 0)
    
     If lRes = 0 Then
        '�Ӿ������ GDI+ ͼ��
        lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
    
         If lRes = 0 Then
             Dim tJpgEncoder As GUID
             Dim tParams As EncoderParameters
            
             '��ʼ����������GUID��ʶ
             CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
            
             '���ý���������
             tParams.Count = 1
            With tParams.Parameter ' Quality
            '�õ�Quality������GUID��ʶ
            CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
            .NumberOfValues = 1
            .type = 4
            .Value = VarPtr(quality)
            End With
        
            '����ͼ��
            lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
            
            '����GDI+ͼ��
            GdipDisposeImage lBitmap
         End If
    
        '���� GDI+
        GdiplusShutdown lGDIP
     End If
    
     If lRes Then
        PictureBoxSaveJPG = False
     Else
        PictureBoxSaveJPG = True
     End If
End Function





