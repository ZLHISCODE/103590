VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISManageEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������Ȩ"
   ClientHeight    =   9555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7290
   Icon            =   "frmCISManageEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   7290
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   64
      Top             =   1440
      Width           =   7300
      Begin VB.PictureBox picDept 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5985
         Left            =   240
         ScaleHeight     =   5985
         ScaleWidth      =   6900
         TabIndex        =   81
         Top             =   1320
         Width           =   6900
         Begin XtremeReportControl.ReportControl rptDept 
            Height          =   5445
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   6855
            _Version        =   589884
            _ExtentX        =   12091
            _ExtentY        =   9604
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.CheckBox chkDept 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ֻ��ʾ���ʿ���"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3720
            TabIndex        =   11
            Top             =   0
            Width           =   1575
         End
         Begin VB.TextBox txtDeptFind 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1245
            TabIndex        =   10
            Top             =   0
            Width           =   1935
         End
         Begin VB.Label lblDept3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Ҷ�λ(&F)"
            Height          =   180
            Left            =   240
            TabIndex        =   82
            Top             =   45
            Width           =   990
         End
      End
      Begin VB.Frame fraPatiType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����������ʵ�ָ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6375
         Index           =   1
         Left            =   120
         TabIndex        =   66
         Top             =   1080
         Width           =   7095
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   4
            Left            =   1800
            Picture         =   "frmCISManageEdit.frx":6852
            ScaleHeight     =   1455
            ScaleWidth      =   3495
            TabIndex        =   72
            Top             =   2280
            Width           =   3495
         End
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Index           =   3
            Left            =   1800
            Picture         =   "frmCISManageEdit.frx":872A
            ScaleHeight     =   1455
            ScaleWidth      =   3495
            TabIndex        =   71
            Top             =   2400
            Width           =   3495
         End
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6015
            Left            =   120
            ScaleHeight     =   6015
            ScaleWidth      =   6900
            TabIndex        =   67
            Top             =   240
            Width           =   6900
            Begin VB.CommandButton cmdDel 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISManageEdit.frx":A696
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   960
               Width           =   330
            End
            Begin VB.CommandButton cmdAdd 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISManageEdit.frx":10EE8
               Style           =   1  'Graphical
               TabIndex        =   17
               Top             =   480
               Width           =   330
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1025
               ScaleHeight     =   240
               ScaleWidth      =   1905
               TabIndex        =   69
               Top             =   130
               Width           =   1935
               Begin VB.ComboBox cboDept 
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   13
                  Top             =   -30
                  Width           =   1960
               End
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   3495
               ScaleHeight     =   240
               ScaleWidth      =   1140
               TabIndex        =   68
               Top             =   120
               Width           =   1170
               Begin VB.ComboBox cboFind 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   14
                  Top             =   -30
                  Width           =   1215
               End
            End
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   4680
               TabIndex        =   15
               Top             =   120
               Width           =   1815
            End
            Begin VSFlex8Ctl.VSFlexGrid vsPati 
               Height          =   5475
               Left            =   0
               TabIndex        =   16
               Top             =   480
               Width           =   6525
               _cx             =   1967205621
               _cy             =   1967203769
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
               MouseIcon       =   "frmCISManageEdit.frx":1773A
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16444122
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   16777215
               GridColorFixed  =   16777215
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   0
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   2000
               ColWidthMin     =   0
               ColWidthMax     =   10000
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmCISManageEdit.frx":18014
               ScrollTrack     =   -1  'True
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   0
               MergeCompare    =   0
               AutoResize      =   0   'False
               AutoSizeMode    =   1
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
               OwnerDraw       =   1
               Editable        =   0
               ShowComboButton =   1
               WordWrap        =   -1  'True
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
               AllowUserFreezing=   1
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
               Begin VB.PictureBox picTmp 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   1920
                  ScaleHeight     =   240
                  ScaleWidth      =   480
                  TabIndex        =   70
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   480
               End
            End
            Begin VB.Label lblDept 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "�����˿���"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   80
               TabIndex        =   83
               Top             =   160
               Width           =   900
            End
            Begin VB.Image imgSentence 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   2915
               Picture         =   "frmCISManageEdit.frx":180AF
               ToolTipText     =   "��ʾ��ǰѡ���������Ĳ���"
               Top             =   90
               Width           =   360
            End
         End
      End
      Begin VB.Frame fraPatiType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���ʷ�Χ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   7095
         Begin VB.OptionButton opt��Χ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ָ������"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   3
            Left            =   5280
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt��Χ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���Ʋ���"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   2
            Left            =   3840
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton opt��Χ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ָ�����Ҳ���"
            ForeColor       =   &H80000008&
            Height          =   495
            Index           =   1
            Left            =   1920
            TabIndex        =   7
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton opt��Χ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ȫԺ����"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   600
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.ImageList imgPati 
      Left            =   7320
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":18799
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":1EFFB
            Key             =   "boy"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAppInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   52
      Top             =   1440
      Width           =   7300
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��ϸ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6375
         Left            =   240
         TabIndex        =   53
         Top             =   960
         Width           =   6855
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��챨��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   1440
            TabIndex        =   24
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1000
            Index           =   3
            Left            =   360
            TabIndex        =   59
            Top             =   3120
            Width           =   6135
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "���еĻ����¼"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "���µ�"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   32
               Top             =   350
               Width           =   855
            End
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ���Ļ����¼"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   33
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtHlInfo 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   720
               Width           =   4095
            End
            Begin VB.Image imgHlInfo 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISManageEdit.frx":2585D
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   670
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   700
            Index           =   2
            Left            =   360
            TabIndex        =   58
            Top             =   5640
            Width           =   6255
            Begin VB.OptionButton optJybg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "���м��鱨��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton optJybg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ���ļ��鱨��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   41
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtJybgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   360
               Width           =   4095
            End
            Begin VB.Image imgJybgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISManageEdit.frx":25F47
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   310
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   1
            Left            =   240
            TabIndex        =   57
            Top             =   4440
            Width           =   6255
            Begin VB.OptionButton optJcbg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "���м�鱨��"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   36
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optJcbg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ���ļ�鱨��"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   37
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   525
               Width           =   4095
            End
            Begin VB.Image imgJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5880
               Picture         =   "frmCISManageEdit.frx":26631
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   480
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   0
            Left            =   240
            TabIndex        =   56
            Top             =   1560
            Width           =   6375
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "���е��Ӳ���"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ�����͵Ĳ���"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ָ���Ĳ����ļ�"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   28
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtDzblTpye 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Index           =   0
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   480
               Width           =   4095
            End
            Begin VB.TextBox txtDzblTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   885
               Width           =   4095
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   0
               Left            =   5880
               Picture         =   "frmCISManageEdit.frx":26D1B
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   435
               Width           =   360
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   1
               Left            =   5880
               Picture         =   "frmCISManageEdit.frx":27405
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   840
               Width           =   360
            End
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ٴ�·��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   1440
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���鱨��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   39
            Top             =   5280
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��鱨��"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������ҳ"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ҽ���嵥"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "���Ӳ���"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�����¼"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkAllInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��������"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame fraTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6375
         Left            =   240
         TabIndex        =   54
         Top             =   960
         Width           =   6735
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1410
            Index           =   5
            Left            =   1680
            Picture         =   "frmCISManageEdit.frx":27AEF
            ScaleHeight     =   1410
            ScaleWidth      =   3435
            TabIndex        =   55
            Top             =   2280
            Width           =   3435
         End
      End
      Begin VB.Line lineTmp 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   6840
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   7320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":29907
            Key             =   "Male"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":30169
            Key             =   "feMale"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":369CB
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":36B25
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":3D387
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManageEdit.frx":3D4E1
            Key             =   "dept"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDoctor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   76
      Top             =   1440
      Width           =   7300
      Begin VB.Frame fraDoctor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���÷�����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   7095
         Begin XtremeReportControl.ReportControl rptDoc 
            Height          =   6255
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   6855
            _Version        =   589884
            _ExtentX        =   12091
            _ExtentY        =   11033
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1245
            ScaleHeight     =   240
            ScaleWidth      =   1905
            TabIndex        =   79
            Top             =   240
            Width           =   1935
            Begin VB.ComboBox cboDocDept 
               Height          =   300
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   -30
               Width           =   1960
            End
         End
         Begin VB.CheckBox chkDoctor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ֻ��ʾ��Ȩ����Ա"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3720
            TabIndex        =   4
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtDocFind 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1245
            TabIndex        =   3
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label lblDept1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            Height          =   180
            Left            =   600
            TabIndex        =   80
            Top             =   285
            Width           =   630
         End
         Begin VB.Label lblDept2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���Ҷ�λ(&F)"
            Height          =   180
            Left            =   240
            TabIndex        =   78
            Top             =   645
            Width           =   990
         End
      End
   End
   Begin VB.Frame fraReault 
      BorderStyle     =   0  'None
      Caption         =   "������Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   7335
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   960
         MaxLength       =   50
         TabIndex        =   0
         Top             =   120
         Width           =   6255
      End
      Begin VB.TextBox txtReault 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   960
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label lblTmp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   75
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lblTmp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��  ע"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   74
         Top             =   600
         Width           =   540
      End
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   60
      Top             =   1440
      Width           =   7300
      Begin VB.Frame fraTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "���÷���ʱ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7455
         Left            =   120
         TabIndex        =   61
         Top             =   120
         Width           =   7095
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�ѹ鵵�Ĳ���(�����������ʷסԺ����)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   47
            Top             =   2520
            Width           =   4000
         End
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "δ�鵵�Ĳ���(��������ǰ��Ժ)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   46
            Top             =   2160
            Width           =   4000
         End
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   45
            Top             =   1800
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   1635
            TabIndex        =   43
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   216203267
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   3960
            TabIndex        =   44
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   216203267
            CurrentDate     =   40976
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   3600
            X2              =   3880
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblTmp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "�������ݵ�ʱ�����ƣ�"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   63
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   840
            TabIndex        =   62
            Top             =   503
            Width           =   735
         End
         Begin VB.Line lineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            Index           =   1
            X1              =   240
            X2              =   6720
            Y1              =   960
            Y2              =   960
         End
      End
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&Q)"
      Height          =   375
      Left            =   5760
      TabIndex        =   50
      Top             =   9120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   4200
      TabIndex        =   49
      Top             =   9120
      Width           =   1215
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7980
      Left            =   0
      TabIndex        =   48
      Top             =   1080
      Width           =   7335
      _Version        =   589884
      _ExtentX        =   12938
      _ExtentY        =   14076
      _StockProps     =   64
   End
   Begin VB.Image imtmp 
      Height          =   360
      Left            =   120
      Picture         =   "frmCISManageEdit.frx":43D43
      Stretch         =   -1  'True
      Top             =   9120
      Width           =   360
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblTmp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����������Ȩ"
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   51
      Top             =   9240
      Width           =   1080
   End
End
Attribute VB_Name = "frmCISManageEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '=0ʱΪ������Ȩ,=1ʱΪ�޸���Ȩ
Private mlngApplyID As Long
Private mblnOk As Boolean
Private mlngFindNum As Long '����ҽ��
Private mlngFindDept As Long '���ҿ���

Private mstrNewEMR As String

Private Enum colDoc
    COLD_��ԱID = 0
    COLD_ѡ�� = 1
    COLD_��� = 2
    COLD_���� = 3
    COLD_�Ա� = 4
    COLD_ƴ������ = 5
    COLD_��ʼ��� = 6
    COLD_�������� = 7
    COLD_��������ID = 8
End Enum

Private Enum colDept
    COLB_����ID = 0
    COLB_ѡ�� = 1
    COLB_���� = 2
    COLB_���� = 3
    COLB_���� = 4
End Enum


Private Enum colList
    col_����Id = 0
    col_���� = 1
    col_�Ա� = 2
    col_���� = 3
    COL_��ʶ�� = 4
    col_���� = 5
    COL_��ǰ״̬ = 6
End Enum

Private Enum FileIndex
    File_��ҳ = 0
    File_ҽ�� = 1
    File_���� = 2
    File_���� = 3
    File_��� = 4
    File_���� = 5
    File_·�� = 6
    File_��� = 7
End Enum


Private Enum CmdIndex
    Cmd_���п��� = 1
    Cmd_������� = 2
    Cmd_סԺ���� = 3
End Enum

Public Function ShowEdit(frmParent As Object, ByVal intType As String, ByRef lngApplyID As Long) As Boolean
'���ܣ�������Ȩ���ݱ༭��
    On Error Resume Next
    mintType = intType
    mlngApplyID = lngApplyID
    mblnOk = False
    
    If mlngApplyID = 0 And mintType = 1 Then Exit Function
    Me.Show 1, frmParent
    lngApplyID = mlngApplyID
    ShowEdit = mblnOk
    On Error GoTo 0
End Function

Private Sub cboDocDept_Click()
    Call LoadDoc
End Sub

Private Sub chkDept_Click()
    Call SetDeptShow
End Sub

Private Sub chkDoctor_Click()
    Call SetDocShow
End Sub

Private Sub SetDocShow()
    Dim i As Long
    
    cboDocDept.Enabled = Not (chkDoctor.Value = 1)
    txtDocFind.Enabled = Not (chkDoctor.Value = 1)
    
    For i = 0 To rptDoc.Records.Count - 1
        If chkDoctor.Value = 1 Then
            rptDoc.Records(i).Visible = rptDoc.Records(i).Tag = "1"
        Else
            rptDoc.Records(i).Visible = True
        End If
    Next
    rptDoc.Populate
End Sub


Private Sub SetDeptShow()
    Dim i As Long
    
    For i = 0 To rptDept.Records.Count - 1
        If chkDept.Value = 1 Then
            rptDept.Records(i).Visible = rptDept.Records(i).Tag = "1"
        Else
            rptDept.Records(i).Visible = True
        End If
    Next
    rptDept.Populate
End Sub

Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_Activate()
    If txtName.Enabled And txtName.Visible Then txtName.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlngFindNum = 0
    mlngFindDept = 0
End Sub

Private Sub opt��Χ_Click(Index As Integer)
    Call SetPatiCtl
End Sub


Private Sub chkAllInfo_Click()
    Dim i As Long
    If chkAllInfo.Tag = "1" Then Exit Sub
    
    fraInfo.Visible = Not (chkAllInfo.Value = 1)
    fraTmp.Visible = chkAllInfo.Value = 1
    
    If chkAllInfo.Value = 0 Then
        For i = 0 To 7
            chkInfo(i).Value = 1
        Next
        optDzbl(0).Value = True
        chkHlInfo(0).Value = 1
        optJcbg(0).Value = True
        optJybg(0).Value = True
        Call SetFileCtl
    End If
End Sub

Private Sub chkHlInfo_Click(Index As Integer)
    If chkHlInfo(0).Tag = "1" Then Exit Sub
    chkHlInfo(0).Tag = "1"
    If Index = 0 Then
        If chkHlInfo(0).Value = 1 Then chkHlInfo(1).Value = 0: chkHlInfo(2).Value = 0
    Else
        If chkHlInfo(1).Value = 1 Or chkHlInfo(2).Value Then chkHlInfo(0).Value = 0
    End If
    chkHlInfo(0).Tag = ""
    Call SetFileCtl
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub SetPatiCtl()
    fraPatiType(1).Caption = IIf(opt��Χ(1).Value, "������Ȩ���ʵ�ָ������", IIf(opt��Χ(3).Value, "������Ȩ���ʵ�ָ������", IIf(opt��Χ(0).Value, "ȫԺ����", "���Ʋ���")))
    fraPatiType(1).Enabled = (opt��Χ(1).Value Or opt��Χ(3).Value)
    picDept.Visible = opt��Χ(1).Value
    picTmp(4).Visible = opt��Χ(2).Value
    picPati.Visible = opt��Χ(3).Value
    picTmp(3).Visible = opt��Χ(0).Value
End Sub


Private Sub SetFileCtl()
    fraInfo.Visible = Not (chkAllInfo.Value = 1)
    fraTmp.Visible = chkAllInfo.Value = 1

    'File_����
    fraFile(0).Enabled = chkInfo(File_����).Value = 1
    optDzbl(0).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    optDzbl(1).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    optDzbl(2).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    txtDzblTpye(0).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    txtDzblTpye(1).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    
    txtDzblTpye(0).BackColor = IIf(optDzbl(1).Value = True And chkInfo(File_����).Value = 1, &HFFFFFF, &H80000004)
    txtDzblTpye(1).BackColor = IIf(optDzbl(2).Value = True And chkInfo(File_����).Value = 1, &HFFFFFF, &H80000004)
    
     'File_����
    fraFile(3).Enabled = chkInfo(File_����).Value = 1
    chkHlInfo(0).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    chkHlInfo(1).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    chkHlInfo(2).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    txtHlInfo.ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    
    txtHlInfo.BackColor = IIf(chkHlInfo(2).Value = 1 And chkInfo(File_����).Value = 1, &HFFFFFF, &H80000004)
    
    'File_���
    fraFile(1).Enabled = chkInfo(File_���).Value = 1
    optJcbg(0).ForeColor = IIf(chkInfo(File_���).Value = 1, &H0, &H80000000)
    optJcbg(1).ForeColor = IIf(chkInfo(File_���).Value = 1, &H0, &H80000000)
    txtJcbgTpye.ForeColor = IIf(chkInfo(File_���).Value = 1, &H0, &H80000000)
    
    txtJcbgTpye.BackColor = IIf(optJcbg(1).Value = True And chkInfo(File_���).Value = 1, &HFFFFFF, &H80000004)
    
    'File_����
    fraFile(2).Enabled = chkInfo(File_����).Value = 1
    optJybg(0).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    optJybg(1).ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    txtJybgTpye.ForeColor = IIf(chkInfo(File_����).Value = 1, &H0, &H80000000)
    
    txtJybgTpye.BackColor = IIf(optJybg(1).Value = True And chkInfo(File_����).Value = 1, &HFFFFFF, &H80000004)
    
    '��ʼ��
    If optDzbl(0).Value = False And optDzbl(1).Value = False And optDzbl(2).Value = False Then optDzbl(0).Value = True
    If chkHlInfo(0).Value = 0 And chkHlInfo(1).Value = 0 And chkHlInfo(2).Value = 0 Then chkHlInfo(0).Value = 1
    If optJcbg(0).Value = False And optJcbg(1).Value = False Then optJcbg(0).Value = True
    If optJybg(0).Value = False And optJybg(1).Value = False Then optJybg(0).Value = True
End Sub



Private Sub cmdAdd_Click()
    If Val(vsPati.TextMatrix(vsPati.Rows - 1, col_����Id)) <> 0 Or vsPati.Rows < 2 Then
        vsPati.Rows = vsPati.Rows + 1
    End If
    vsPati.Row = vsPati.Rows - 1
    vsPati.SetFocus
End Sub


Private Sub cmdDel_Click()
    If vsPati.Row < 1 Then Exit Sub
    If Val(vsPati.TextMatrix(vsPati.Row, col_����Id)) <> 0 Then
        vsPati.Tag = Replace(vsPati.Tag, Val(vsPati.TextMatrix(vsPati.Row, col_����Id)), "")
    End If
    vsPati.RemoveItem vsPati.Row
    If vsPati.Rows < 2 Then
        Call cmdAdd_Click
    End If
    vsPati.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim str����ids As String
    Dim str�������� As String
    Dim str������ids As String
    Dim str����ids As String
    Dim strXML As String
    Dim lngID As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim i As Long
    Dim curDate As Date
    Dim blnTran As Boolean
    Dim int���ʲ��� As Integer  '0-ȫԺ���ˣ�1-���Ʋ��ˣ�2-ָ�����Ҳ��ˣ�3-ָ�����ˣ�4-���Ϊָ�������Ĳ��ˣ�5-ָ�������Ĳ��ˡ�2-4�Ķ�������ͨ���ӱ�洢';
    Dim rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    
    On Error GoTo errH
   '��鷽����
    If txtName.Text = "" Then
        MsgBox "��ǰ��δ¼�뷽����,������¼�롣", vbInformation, gstrSysName
        txtName.SetFocus
        Exit Sub
    End If
    
    If ZLCommFun.ActualLen(txtName.Text) > txtName.MaxLength Then
        MsgBox "���������ݹ��࣬������� " & txtName.MaxLength \ 2 & " �����ֻ� " & txtName.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtName.SetFocus: Exit Sub
    End If
    
    strSQL = "select Count(1) as ��� from ���Ӳ���������Ȩ where ������=[1]" & IIf(mlngApplyID = 0, "", " and ID<>[2]")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtName.Text, mlngApplyID)
    
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            If Val(rsTmp!��� & "") > 0 Then
                MsgBox "��ǰ¼��ķ������ѱ�ʹ�ã�������¼�롣", vbInformation, gstrSysName
                txtName.SetFocus: Exit Sub
            End If
        End If
    End If
    
    '��鷽����ע
    If ZLCommFun.ActualLen(txtReault.Text) > txtReault.MaxLength Then
        MsgBox "������ע���ݹ��࣬������� " & txtReault.MaxLength \ 2 & " �����ֻ� " & txtReault.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtReault.SetFocus: Exit Sub
    End If
    
    '��ȡ������
    For i = 0 To rptDoc.Records.Count - 1
        If rptDoc.Records(i).Tag = "1" And Val(rptDoc.Records(i)(COLD_��ԱID).Value) <> 0 Then
            str������ids = str������ids & "," & rptDoc.Records(i)(COLD_��ԱID).Value
        End If
    Next
    str������ids = Mid(str������ids, 2)
    If str������ids = "" Then
        Me.tbcSub.Item(0).Selected = True
        MsgBox "��ǰ��δ¼���������Ϣ,������¼�롣", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡ���ʲ���
    If opt��Χ(3).Value = True Then
        For i = 1 To vsPati.Rows - 1
            If Val(vsPati.TextMatrix(i, col_����Id)) <> 0 Then
                str����ids = str����ids & "," & Val(vsPati.TextMatrix(i, col_����Id))
                str�������� = str�������� & "," & Val(vsPati.TextMatrix(i, col_����))
            End If
        Next
        str����ids = Mid(str����ids, 2)
        str�������� = Mid(str��������, 2)
        
        If str����ids = "" Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "��ǰ��δ¼����Ҫ��Ȩ���ʲ����Ĳ�����Ϣ,������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
    ElseIf opt��Χ(1).Value = True Then
        For i = 0 To rptDept.Records.Count - 1
            If rptDept.Records(i).Tag = "1" And Val(rptDept.Records(i)(COLB_����ID).Value) <> 0 Then
                str����ids = str����ids & "," & rptDept.Records(i)(COLB_����ID).Value
            End If
        Next
        str����ids = Mid(str����ids, 2)
        
        If str����ids = "" Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "��ǰ��δ¼����Ҫ��Ȩ���ʲ����Ĳ��˿���,������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    int���ʲ��� = IIf(opt��Χ(1).Value, 2, IIf(opt��Χ(3).Value, 3, IIf(opt��Χ(0).Value, 0, 1)))
    
    '����������
    If chkAllInfo.Value = 0 Then
        For i = 0 To 7
            If chkInfo(i).Value = 1 Then
                lngTmp = lngTmp + 1
            End If
        Next
        If lngTmp = 0 Then
            Me.tbcSub.Item(2).Selected = True
            MsgBox "��ǰ��δ¼����Ҫ������ʲ�����Ȩ������,������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    For i = 0 To 1
        If txtDzblTpye(i).BackColor = &HFFFFFF And txtDzblTpye(i).Text = "" And chkAllInfo.Value = 0 Then
            Me.tbcSub.Item(2).Selected = True
            MsgBox "��ǰ��δ¼�벡���ļ�" & IIf(i = 0, "����", "") & ",������¼��!!!", vbInformation, gstrSysName
            txtDzblTpye(i).SetFocus
            Exit Sub
        End If
    Next
    If txtHlInfo.BackColor = &HFFFFFF And txtHlInfo.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "��ǰ��δ¼�뻤���¼�ļ�,������¼�롣", vbInformation, gstrSysName
        txtHlInfo.SetFocus
        Exit Sub
    End If
    If txtJcbgTpye.BackColor = &HFFFFFF And txtJcbgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "��ǰ��δ¼���鱨������,������¼�롣", vbInformation, gstrSysName
        txtJcbgTpye.SetFocus
        Exit Sub
    End If
    If txtJybgTpye.BackColor = &HFFFFFF And txtJybgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "��ǰ��δ¼����鱨������,������¼�롣", vbInformation, gstrSysName
        txtJybgTpye.SetFocus
        Exit Sub
    End If
   
    '������ʱ��
    If dtpTime(0).Value >= dtpTime(1).Value Then
        Me.tbcSub.Item(3).Selected = True
        MsgBox "��ǰ������ʼʱ�����С����ֹʱ��,������¼�롣", vbInformation, gstrSysName
        txtReault.SetFocus
        Exit Sub
    End If
    
    strXML = GetInfoXml
    
    '��������
    lngID = mlngApplyID
    If lngID = 0 Then lngID = zlDatabase.GetNextId("���Ӳ���������Ȩ")
    curDate = zlDatabase.Currentdate
    strSQL = "Zl_���Ӳ���������Ȩ_Update(" & mintType & "," & lngID & ",1,NULL,'" & txtName.Text & "'," & int���ʲ��� & ",'" & strXML & "',To_Date('" & Format(dtpTime(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            "To_Date('" & Format(dtpTime(1).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            IIf(optTimeTpye(0).Value, 0, IIf(optTimeTpye(1).Value, 1, 2)) & ",'" & UserInfo.���� & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & txtReault.Text & "')"
    
    
    arrSQL = Array()
    
    '��ȡ���ʲ���
    If opt��Χ(3).Value = True Then
        For i = 1 To vsPati.Rows - 1
            If Val(vsPati.TextMatrix(i, col_����Id)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�����Ȩ���ʲ���_Insert(" & lngID & ",2," & Val(vsPati.TextMatrix(i, col_����Id)) & ")"
            End If
        Next
    ElseIf opt��Χ(1).Value = True Then
        For i = 0 To rptDept.Records.Count - 1
            If rptDept.Records(i).Tag = "1" And Val(rptDept.Records(i)(COLB_����ID).Value) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�����Ȩ���ʲ���_Insert(" & lngID & ",3," & Val(rptDept.Records(i)(COLB_����ID).Value) & ")"
            End If
        Next
    End If
    
    For i = 0 To rptDoc.Records.Count - 1
        If rptDoc.Records(i).Tag = "1" And Val(rptDoc.Records(i)(COLB_����ID).Value) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ�����Ȩ������Ա_Insert(" & lngID & "," & Val(rptDoc.Records(i)(COLD_��ԱID).Value) & ")"
        End If
    Next

    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    For i = 0 To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    mlngApplyID = lngID
    mblnOk = True
    On Error GoTo 0
    Screen.MousePointer = 0
    Unload Me
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub




Private Function ReadXmlSet() As Boolean
    '��ȡ��Ȩ���ݵ�Xml������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String

    
    On Error GoTo errH
    If mlngApplyID = 0 Then Exit Function
    
    strXML = Sys.ReadXML("���Ӳ���������Ȩ", "��������", "ID=[1]", strErr, mlngApplyID)
    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function
    
    '��������
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    chkAllInfo.Tag = "1"
    chkAllInfo.Value = Val(strValue)
    chkAllInfo.Tag = ""
    If Val(strValue) = 0 Then
        '������ҳ��ҽ�����ٴ�·��
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_��ҳ).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_ҽ��).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_·��).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_���).Value = 1
        
        '�����¼
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_����).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber): If Val(strValue) = 1 Then chkHlInfo(0).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then chkHlInfo(1).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber):
            If Val(strValue) = 1 Then
                chkHlInfo(2).Value = 1
                If GetXmlString(objXML, "nursing_info/file_name", strValue) Then
                    txtHlInfo.Text = strValue
                End If
            End If
        End If
        
        '��鱨��
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_���).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0���м�鱨�� =1ָ�����͵ļ�鱨��
            optJcbg(Val(strValue)).Value = True
            If Val(strValue) = 1 Then

                If GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue) Then
                    txtJcbgTpye.Text = strValue
                End If
            End If
        End If
        
        '���鱨��
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_����).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 ���м��鱨�� =1ָ�����͵ļ��鱨��
            optJybg(Val(strValue)).Value = True
            
            If Val(strValue) = 1 Then
                If GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue) Then
                    txtJybgTpye.Text = strValue
                End If
            End If
        End If
        
        '���Ӳ���
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_����).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 ���е��Ӳ���  =1ָ�����͵ĵ��Ӳ���  =1ָ������ĵ��Ӳ���
            optDzbl(Val(strValue)) = True
            
            If Val(strValue) = 1 Then
                If GetXmlString(objXML, "emr_info/standard_class/class_name", strValue) Then
                    txtDzblTpye(0).Text = strValue
                End If
            ElseIf Val(strValue) = 2 Then
                If GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue) Then
                    txtDzblTpye(1).Text = strValue
                End If
            End If
        End If
    End If
    ReadXmlSet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetInfoXml() As String
    Dim objXML As New zl9ComLib.clsXML
    Dim i As Long
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "app_info"                          '���ڵ�[������Ϣ]
        .appendData "all_files", chkAllInfo.Value       '<��������>���ͣ�N
        If chkAllInfo.Value <> 1 Then
            .appendData "medical_record", chkInfo(File_��ҳ).Value  '<������ҳ>���ͣ�N
            .appendData "advice", chkInfo(File_ҽ��).Value          '<ҽ���嵥>���ͣ�N
            .appendData "emr", chkInfo(File_����).Value             '<���Ӳ���>���ͣ�N
                If chkInfo(File_����).Value = 1 Then
                    .AppendNode "emr_info"  '���ڵ�[���Ӳ�����ϸ]
                        .appendData "emr_type", IIf(optDzbl(0).Value = True, 0, IIf(optDzbl(1).Value = True, 1, 2)) '<���Ӳ�������>���ͣ�N
                        If optDzbl(1).Value And txtDzblTpye(0).Text <> "" Then
                            .AppendNode "standard_class"  '���ڵ�[����׼����]
                            For i = 0 To UBound(Split(txtDzblTpye(0).Text, ","))
                                .appendData "class_name", Split(txtDzblTpye(0).Text, ",")(i)
                            Next
                            .AppendNode "standard_class", True
                        ElseIf optDzbl(2).Value And txtDzblTpye(1).Text <> "" Then
                            .AppendNode "antetype_class"  '���ڵ�[������ԭ��]
                            For i = 0 To UBound(Split(txtDzblTpye(1).Text, ","))
                                .appendData "class_name", Split(txtDzblTpye(1).Text, ",")(i)
                            Next
                            .AppendNode "antetype_class", True
                        End If
                    .AppendNode "emr_info", True
                End If
            .appendData "nursing_record", chkInfo(File_����).Value      '<�����¼>���ͣ�N
                If chkInfo(File_����).Value = 1 Then
                    .AppendNode "nursing_info"  '���ڵ�[�����¼��ϸ]
                        .appendData "nursing_all", chkHlInfo(0).Value  '<���л����¼>���ͣ�N
                        .appendData "thermometer", chkHlInfo(1).Value  '<�Ƿ�����������µ�>���ͣ�N
                        .appendData "record_file", chkHlInfo(2).Value   '<�Ƿ�ָ�������¼>���ͣ�N
                        If chkHlInfo(2).Value = 1 And txtHlInfo.Text <> "" Then
                            For i = 0 To UBound(Split(txtHlInfo.Text, ","))
                                .appendData "file_name", Split(txtHlInfo.Text, ",")(i)
                            Next
                        End If
                    .AppendNode "nursing_info", True
                End If
            .appendData "pacs_report", chkInfo(File_���).Value         '<��鱨��>���ͣ�N
                If chkInfo(File_���).Value = 1 Then
                    .AppendNode "pacs_info"  '���ڵ�[��鱨����ϸ]
                        .appendData "pacs_type", IIf(optJcbg(0).Value = True, 0, 1) '<��鱨������>���ͣ�N
                        If optJcbg(1).Value And txtJcbgTpye.Text <> "" Then
                            .AppendNode "pacs_report_type"  '���ڵ�[����׼����]
                            For i = 0 To UBound(Split(txtJcbgTpye.Text, ","))
                                .appendData "type_name", Split(txtJcbgTpye.Text, ",")(i)
                            Next
                            .AppendNode "pacs_report_type", True
                        End If
                    .AppendNode "pacs_info", True
                End If
            .appendData "lis_report", chkInfo(File_����).Value          '<���鱨��>���ͣ�N
                If chkInfo(File_����).Value = 1 Then
                    .AppendNode "lis_info"  '���ڵ�[���鱨����ϸ]
                        .appendData "lis_type", IIf(optJybg(0).Value = True, 0, 1) '<���鱨������>���ͣ�N
                        If optJybg(1).Value And txtJybgTpye.Text <> "" Then
                            .AppendNode "lis_report_type"  '���ڵ�[����׼����]
                            For i = 0 To UBound(Split(txtJybgTpye.Text, ","))
                                .appendData "type_name", Split(txtJybgTpye.Text, ",")(i)
                            Next
                            .AppendNode "lis_report_type", True
                        End If
                    .AppendNode "lis_info", True
                End If
            .appendData "cispath", chkInfo(File_·��).Value             '<�ٴ�·��>���ͣ�N
            .appendData "patipeis", chkInfo(File_���).Value             '<��챨��>���ͣ�N
         End If
        .AppendNode "app_info", True
        GetInfoXml = .XmlText
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim curDate As Date

    On Error GoTo errH
    Me.Caption = IIf(mintType = 0, "����������Ȩ", "�޸ķ�����Ȩ")
    lblTmp(0).Caption = Me.Caption
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        
        .InsertItem(0, "������", picDoctor.hwnd, 0).Tag = "������"
        .InsertItem(1, "���ò���", picParent.hwnd, 0).Tag = "���ò���"
        .InsertItem(2, "��������", picAppInfo.hwnd, 0).Tag = "��������"
        .InsertItem(3, "����ʱ��", picTime.hwnd, 0).Tag = "����ʱ��"
        
        .Item(3).Selected = True
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call LoadDept
    Call LoadPatiDept

    
    '��ʼ�����˱��
    Call InitPatiTable
    
    Call InitReportColumn
    
    Call LoadDoc
    
     
    'ִ�н�������˵���ʼ��
    cboFind.Clear
    cboFind.AddItem "����"
    cboFind.AddItem "���֤��"
    cboFind.AddItem "�����"
    cboFind.AddItem "סԺ��"
    cboFind.AddItem "����ID"
    cboFind.ListIndex = 0
    
    
    If mintType = 1 Then
        '���ػ�����Ϣ
        Call LoadOther
        
        '���ط�����
        Call SetDoc
        chkDoctor.Value = 1
        Call SetDocShow
        
        '���ط��ʷ�Χ
        If opt��Χ(3).Value = True Then
            Call LoadPati
        ElseIf opt��Χ(1).Value = True Then
            Call SetDept
            chkDept.Value = 1
            Call SetDeptShow
        End If
        '���ط�������
        Call ReadXmlSet
    Else
        chkAllInfo.Value = 1
        curDate = zlDatabase.Currentdate
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd hh:mm")
        dtpTime(1).Value = Format(curDate + 7, "yyyy-MM-dd hh:mm")
        optTimeTpye(0).Value = True
    End If
    Call SetPatiCtl
    Call SetFileCtl
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadOther()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select a.Id, a.��Ȩ����, a.����id, a.������, a.���ʲ���, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.��ע From ���Ӳ���������Ȩ A Where a.Id =[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            dtpTime(0).Value = Format(rsTmp!���ʿ�ʼʱ�� & "", "yyyy-MM-dd hh:mm")
            dtpTime(1).Value = Format(rsTmp!���ʽ���ʱ�� & "", "yyyy-MM-dd hh:mm")
            optTimeTpye(Val(rsTmp!����ʱ�� & "")).Value = True
            txtReault.Text = rsTmp!��ע & ""
            txtName.Text = rsTmp!������ & ""
            opt��Χ(decode(Val(rsTmp!���ʲ��� & ""), 2, 1, 3, 3, 0, 0, 2)) = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPati()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    
    On Error GoTo errH
    With vsPati
        strSQL = "Select d.Id, d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��, d.��ǰ״̬" & vbNewLine & _
                "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��," & vbNewLine & _
                "                     Decode(a.ִ��״̬, 1, '��' || To_Char(a.ִ��ʱ��, 'yyyy-mm-dd') || '���������Ժ', '�������ھ���') As ��ǰ״̬" & vbNewLine & _
                "              From ���˹Һż�¼ A, ���Ӳ�����Ȩ���ʲ��� G" & vbNewLine & _
                "              Where g.��Ȩ���� = a.����id And g.��Ȩid = [1] And ��¼״̬ = 1" & vbNewLine & _
                "              Union All" & vbNewLine & _
                "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��," & vbNewLine & _
                "                     Decode(b.��Ժ����, Null, '��Ժ', '��' || b.��ҳid || '��סԺ��Ժ') As ��ǰ״̬" & vbNewLine & _
                "              From ������ҳ B, ���Ӳ�����Ȩ���ʲ��� H" & vbNewLine & _
                "              Where h.��Ȩ���� = b.����id And h.��Ȩid = [1]) C) D, ���ű� G" & vbNewLine & _
                "Where g.Id = d.���� And d.Top = 1" & vbNewLine & _
                "Order By d.����ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                    If Val(.TextMatrix(.Rows - 1, col_����Id)) <> 0 Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    
                    .TextMatrix(lngRow, col_����Id) = rsTmp!ID & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    Set .Cell(flexcpPicture, lngRow, col_����) = imgPati.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
                    .TextMatrix(lngRow, col_�Ա�) = rsTmp!�Ա� & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    .TextMatrix(lngRow, COL_��ʶ��) = rsTmp!��ʶ�� & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    .TextMatrix(lngRow, COL_��ǰ״̬) = rsTmp!��ǰ״̬ & ""
                    .Tag = .Tag & "," & rsTmp!ID & ","
                End If
                rsTmp.MoveNext
            Loop
            .WordWrap = True
            '�Զ������и�
            .AutoSize col_����, COL_��ǰ״̬
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetDept()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select a.��Ȩ���� From ���Ӳ�����Ȩ���ʲ��� A Where a.��Ȩid = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
    If Not rsTmp Is Nothing Then
        For i = 0 To rptDept.Records.Count - 1
            If Val(rptDept.Records(i)(COLB_����ID).Value) <> 0 Then
                rsTmp.Filter = "��Ȩ���� =" & Val(rptDept.Records(i)(COLB_����ID).Value)
                If rsTmp.RecordCount > 0 Then
                    rptDept.Records(i)(COLB_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                    rptDept.Records(i).Tag = "1"
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDoc()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    strSQL = "Select a.��ԱID From ���Ӳ�����Ȩ������Ա A Where a.��Ȩid = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
    If Not rsTmp Is Nothing Then
        For i = 0 To rptDoc.Records.Count - 1
            If Val(rptDoc.Records(i)(COLD_��ԱID).Value) <> 0 Then
                rsTmp.Filter = "��ԱID =" & Val(rptDoc.Records(i)(COLD_��ԱID).Value)
                If rsTmp.RecordCount > 0 Then
                    rptDoc.Records(i)(COLD_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                    rptDoc.Records(i).Tag = "1"
                End If
            End If
        Next
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub imgDzblTpye_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String

    If txtDzblTpye(Index).BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgDzblTpye(Index).Container.hwnd, imgDzblTpye(Index).Left, imgDzblTpye(Index).Top)
    blnCancel = True
    On Error GoTo errH
    If Index = 0 Then
        strSQL = "select 1 as ID, '���ﲡ��' as ��������," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",���ﲡ��,"), 1, 0) & " as �ѹ�ѡcheck from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 2 as ID, 'סԺ����' as ��������," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",סԺ����,"), 1, 0) & " as �ѹ�ѡcheck from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 4 as ID, '������' as ��������," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",������,"), 1, 0) & " as �ѹ�ѡcheck from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 5 as ID, '����֤������' as ��������," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",����֤������,"), 1, 0) & " as �ѹ�ѡcheck from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 6 as ID, '֪���ļ�' as ��������," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",֪���ļ�,"), 1, 0) & " as �ѹ�ѡcheck from dual"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ�����ļ�����", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(0).Height, blnCancel, True, True)
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!��������
                    rsTmp.MoveNext
                Loop
                txtDzblTpye(0).Text = Mid(strTmp, 2)
            Else
                MsgBox "δ���ҵ�����ѡ��Ĳ����ļ�����!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    Else

        '�²���
        
        If mstrNewEMR = "" Then
            Set rsTmp = Nothing
            On Error Resume Next
            If Not gobjEmr Is Nothing Then
                strSQL = "Select Title as ���� From Antetype_List Where nvl(disable,0)=0 Order By Code"
                Call gobjEmr.OpenSQLRecordset(strSQL, "", rsTmp)
            End If
            Err.Clear: On Error GoTo 0
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    mstrNewEMR = mstrNewEMR & "," & rsTmp!����
                    rsTmp.MoveNext
                Loop
            End If
            mstrNewEMR = Mid(mstrNewEMR, 2)
        End If
            
        strSQL = ""
        If mstrNewEMR <> "" Then
            strSQL = "Select Rownum + 100000 As ID, '�°没��' As ��������, b.C2 As ����, Decode(d.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
                     "From Table(Cast(f_Str2list2([2]) As Zltools.t_Strlist2)) B," & vbNewLine & _
                     "     (Select Replace(C2, '���°没����', '') As C2" & vbNewLine & _
                     "       From Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) C" & vbNewLine & _
                     "       Where Instr(C2, '���°没����') > 0) D" & vbNewLine & _
                     "Where b.C2 = d.C2(+) union all "
        End If

        strSQL = strSQL & " Select * from (Select a.ID,Decode(a.����, 1, '���ﲡ��', 2, 'סԺ����', 4, '������', 5, '����֤��', 6, '֪���ļ�') As ��������, a.����," & vbNewLine & _
                "       Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
                "From �����ļ��б� A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
                "Where a.���� In (1, 2, 4, 5, 6) And a.���� = b.C2(+)" & vbNewLine & _
                "Order By ��������, ���)"

        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ�����ļ�", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(1).Height, blnCancel, True, True, txtDzblTpye(1).Text, mstrNewEMR)
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & IIf(rsTmp!�������� & "" = "�°没��", "���°没����", "") & rsTmp!����
                    rsTmp.MoveNext
                Loop
                txtDzblTpye(1).Text = Mid(strTmp, 2)
            Else
                MsgBox "δ���ҵ�����ѡ��Ĳ����ļ�!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgHlInfo_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtHlInfo.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgHlInfo.Container.hwnd, imgHlInfo.Left, imgHlInfo.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.ID,'�����¼' As ��������, a.����," & vbNewLine & _
            "       Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
            "From �����ļ��б� A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.���� =3 AND A.����<>-1 And a.���� = b.C2(+)" & vbNewLine & _
            "Order By ����, ���"

    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ�����¼�ļ�", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(1).Height, blnCancel, True, True, txtHlInfo.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            txtHlInfo.Text = Mid(strTmp, 2)
        Else
            MsgBox "δ���ҵ�����ѡ��Ļ����¼�ļ�!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgJcbgTpye_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtJcbgTpye.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgJcbgTpye.Container.hwnd, imgJcbgTpye.Left, imgJcbgTpye.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.���� As ID, a.����, Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
            "From ���Ƽ������ A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.���� = b.C2(+)" & vbNewLine & _
            "Order By ����"
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ���鱨������", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgJcbgTpye.Height, blnCancel, True, True, txtJcbgTpye.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            txtJcbgTpye.Text = Mid(strTmp, 2)
        Else
            MsgBox "δ���ҵ�����ѡ��ļ�鱨������!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgJybgTpye_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtJybgTpye.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgJybgTpye.Container.hwnd, imgJybgTpye.Left, imgJybgTpye.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.���� As ID, a.����, Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
            "From ���Ƽ������� A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.���� = b.C2(+)" & vbNewLine & _
            "Order By ����"
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ����鱨������", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgJybgTpye.Height, blnCancel, True, True, txtJybgTpye.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            txtJybgTpye.Text = Mid(strTmp, 2)
        Else
            MsgBox "δ���ҵ�����ѡ��ļ��鱨������!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgSentence_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim lngRow As Long
    
    vPoint = zlcontrol.GetCoordPos(cboDept.Container.hwnd, cboDept.Left, cboDept.Top)
    blnCancel = True
    On Error GoTo errH
    
    If lblDept.Tag = "" Then
        strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                    "              From ���˹Һż�¼ A" & vbNewLine & _
                    "              Where ��¼״̬=1 And a.ִ��ʱ�� Between Sysdate - 7 And Sysdate" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                    "              From ������ҳ B" & vbNewLine & _
                    "              Where b.��Ժ���� Between Sysdate - 7 And Sysdate) C) D, ���ű� G" & vbNewLine & _
                    "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                    "Order By d.����ʱ�� Desc"
    ElseIf lblDept.Tag = "����" Then
        strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                    "              From ���˹Һż�¼ A" & vbNewLine & _
                    "              Where ��¼״̬=1 And a.ִ��ʱ�� Between Sysdate - 7 And Sysdate) C) D, ���ű� G" & vbNewLine & _
                    "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                    "Order By d.����ʱ�� Desc"
    ElseIf lblDept.Tag = "סԺ" Then
        strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                    "              From ������ҳ B" & vbNewLine & _
                    "              Where b.��Ժ���� Between Sysdate - 7 And Sysdate) C) D, ���ű� G" & vbNewLine & _
                    "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                    "Order By d.����ʱ�� Desc"
    End If
    
    
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ�����7��Ĳ���", True, "", "", True, True, True, vPoint.X, vPoint.Y, cboDept.Height, blnCancel, True, True, cboDept.ItemData(cboDept.ListIndex))
    With vsPati
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                        If Val(.TextMatrix(.Rows - 1, col_����Id)) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        
                        .TextMatrix(lngRow, col_����Id) = rsTmp!ID & ""
                        .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                        Set .Cell(flexcpPicture, lngRow, col_����) = imgPati.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
                        .TextMatrix(lngRow, col_�Ա�) = rsTmp!�Ա� & ""
                        .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                        .TextMatrix(lngRow, COL_��ʶ��) = rsTmp!��ʶ�� & ""
                        .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                        .TextMatrix(lngRow, COL_��ǰ״̬) = rsTmp!��ǰ״̬ & ""
                        .Tag = .Tag & "," & rsTmp!ID & ","
                    End If
                    rsTmp.MoveNext
                Loop
                .WordWrap = True
                '�Զ������и�
                .AutoSize col_����, COL_��ǰ״̬
            Else
                 MsgBox "δ���ҵ������ҽ��ڵ�" & lblDept.Tag & "����!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End With
    Exit Sub
errH:
    MsgBox "δ���ҵ������ҽ��ڵ�" & lblDept.Tag & "����!", vbInformation, Me.Caption
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, True)
End Sub


Private Sub picPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, False)
End Sub

Private Sub optDzbl_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub optJcbg_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub optJybg_Click(Index As Integer)
    Call SetFileCtl
End Sub


Private Sub txtDzblTpye_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtDzblTpye(Index).hwnd, Replace(txtDzblTpye(Index).Text, ",", "��" & vbCrLf), True, True)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim lngRow As Long
    
    Dim colPati As Collection, str����ids As String, i As Long
    
    If KeyAscii = vbKeyReturn Then
        If Len(txtFind.Text) < 1 Then Exit Sub
        vPoint = zlcontrol.GetCoordPos(cboDept.Container.hwnd, cboDept.Left, cboDept.Top)
        blnCancel = True
        On Error GoTo errH
        
        If cboFind.Text = "�����" Then
            strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                        "              From ���˹Һż�¼ A" & vbNewLine & _
                        "              Where A.��¼״̬=1 And A.�����=[2]) C) D, ���ű� G" & vbNewLine & _
                        "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                        "Order By d.����ʱ�� Desc"
        ElseIf cboFind.Text = "סԺ��" Then
            strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                        "              From ������ҳ B" & vbNewLine & _
                        "              Where B.סԺ��=[2]) C) D, ���ű� G" & vbNewLine & _
                        "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                        "Order By d.����ʱ�� Desc"
        Else
            If cboFind.Text = "���֤��" Then
                Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, txtFind.Text)
            End If
        
            If Not colPati Is Nothing Then
                If colPati.Count > 0 Then
                    For i = 1 To colPati.Count
                        If InStr("," & str����ids & ",", "," & Val(GetColVal(colPati(i), "_pati_id")) & ",") = 0 Then
                           str����ids = str����ids & "," & Val(GetColVal(colPati(i), "_pati_id"))
                        End If
                    Next
                End If
            End If
            If str����ids <> "" Then str����ids = Mid(str����ids, 2)
        
        
        
        
        
            If lblDept.Tag = "" Then
                strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                            "              From ���˹Һż�¼ A" & vbNewLine & _
                            "              Where A.��¼״̬=1 And " & decode(cboFind.Text, "���֤��", " A.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "A.����ID =[2]", "����", "A.���� like [2]") & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                            "              From ������ҳ B" & vbNewLine & _
                            "              Where " & decode(cboFind.Text, "���֤��", " B.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[2]", "����", "B.���� like [2]") & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            ElseIf lblDept.Tag = "סԺ" Then
                strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                            "              From ������ҳ B" & vbNewLine & _
                            "              Where " & decode(cboFind.Text, "���֤��", " B.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[2]", "����", "B.���� like [2]") & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            ElseIf lblDept.Tag = "����" Then
                strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��,decode(A.ִ��״̬,1,'��'||to_char(A.ִ��ʱ��,'yyyy-mm-dd') || '���������Ժ','�������ھ���') as ��ǰ״̬" & vbNewLine & _
                            "              From ���˹Һż�¼ A" & vbNewLine & _
                            "              Where A.��¼״̬=1 And " & decode(cboFind.Text, "���֤��", " A.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "A.����ID =[2]", "����", "A.���� like [2]") & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            End If
        End If
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "���Ҳ���", True, "", "", True, True, True, vPoint.X, vPoint.Y, cboDept.Height, blnCancel, True, True, cboDept.ItemData(cboDept.ListIndex), IIf(InStr(",�����,סԺ��,����ID,", cboFind.Text) > 0, Val(txtFind.Text), IIf(cboFind.Text = "����", txtFind.Text & "%", txtFind.Text)), str����ids)
        With vsPati
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                            If Val(.TextMatrix(.Rows - 1, col_����Id)) <> 0 Then
                                .Rows = .Rows + 1
                            End If
                            lngRow = .Rows - 1
                            
                            .TextMatrix(lngRow, col_����Id) = rsTmp!ID & ""
                            .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                            Set .Cell(flexcpPicture, lngRow, col_����) = imgPati.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
                            .TextMatrix(lngRow, col_�Ա�) = rsTmp!�Ա� & ""
                            .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                            .TextMatrix(lngRow, COL_��ʶ��) = rsTmp!��ʶ�� & ""
                            .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                            .TextMatrix(lngRow, COL_��ǰ״̬) = rsTmp!��ǰ״̬ & ""
                            .Tag = .Tag & "," & rsTmp!ID & ","
                        End If
                        rsTmp.MoveNext
                    Loop
                    .WordWrap = True
                    '�Զ������и�
                    .AutoSize col_����, COL_��ǰ״̬
                Else
                    MsgBox "�ڵ�ǰ����δ���ҵ�����!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    Else
        Select Case cboFind.Text
            Case "סԺ��", "�����", "����ID"
                If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "���֤��"
                If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "����"
        End Select
    End If
    Exit Sub
errH:
    MsgBox "�ڵ�ǰ����δ���ҵ�����!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptDoc
        Set objCol = .Columns.Add(COLD_��ԱID, "��ԱID", 0, False)
        Set objCol = .Columns.Add(COLD_ѡ��, "", 18, False)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(COLD_���, "���", 100, True)
        Set objCol = .Columns.Add(COLD_����, "����", 100, True)
        Set objCol = .Columns.Add(COLD_�Ա�, "�Ա�", 60, True)
        Set objCol = .Columns.Add(COLD_ƴ������, "ƴ������", 0, False)
        Set objCol = .Columns.Add(COLD_��ʼ���, "��ʼ���", 0, False)
        Set objCol = .Columns.Add(COLD_��������, "��������", 100, True)
        Set objCol = .Columns.Add(COLD_��������ID, "��������ID", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�ķ�����..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
    
    With rptDept
        Set objCol = .Columns.Add(COLB_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COLB_ѡ��, "", 20, True)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentRight
            objCol.Icon = img16.ListImages("unCheck").Index - 1
        Set objCol = .Columns.Add(COLB_����, "����", 100, True)
        Set objCol = .Columns.Add(COLB_����, "����", 150, True)
        Set objCol = .Columns.Add(COLB_����, "����", 100, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
            .HighlightBackColor = &HFFEDCA
            .HighlightForeColor = vbBlack
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub


Private Sub LoadPatiDept()
'���ز�ѯPati����
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    Dim strFiter As String
    
    strSQL = "Select B.ID,B.����,B.���� From " & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            "  And C.�������� = '�ٴ�' " & decode(lblDept.Tag, "", " And C.������� <> 0 ", "����", " And C.������� in (1,3) ", "סԺ", " And C.������� in (2,3) ") & "  And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"

    On Error GoTo errH
    cboDept.Clear
    '���в���
    cboDept.AddItem "���в���"
    cboDept.ItemData(cboDept.NewIndex) = -1
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID & ""
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDept()
'���ز���Ա��������
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    
    
    
    strSQL = "Select B.ID,B.����,B.����,B.���� From " & _
            " ���ű� B, ��������˵�� C" & vbNewLine & _
            " Where B.Id = C.����id " & _
            "  And C.�������� = '�ٴ�' And C.������� <> 0  And (B.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.����ʱ�� Is Null) Order By B.����"


    On Error GoTo errH
    
    cboDocDept.Clear
    '���в���
    cboDocDept.AddItem "���в���"
    cboDocDept.ItemData(cboDocDept.NewIndex) = -1
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    rptDept.Records.DeleteAll
    With rptDept
        For i = 1 To rsTmp.RecordCount
            cboDocDept.AddItem rsTmp!���� & "-" & rsTmp!����
            cboDocDept.ItemData(cboDocDept.NewIndex) = rsTmp!ID & ""
            
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem("")
                objItem.Icon = img16.ListImages("unCheck").Index - 1
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                objItem.Icon = img16.ListImages.Item("dept").Index - 1
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                
            rsTmp.MoveNext
        Next
        .Populate
    End With

    
    If cboDocDept.ListIndex = -1 And cboDocDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDocDept.hwnd, 0)
    End If
    mlngFindDept = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadDoc()
    '����ҽ��
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem

    If cboDocDept.ListIndex = -1 Then Exit Sub
    
    strSQL = "Select DISTINCT A.���,a.Id,a.����, A.�Ա� ,b.����ID,e.���� as ��������, Upper(zlSpellCode(a.����)) As ƴ������, Upper(Zlwbcode(a.����)) As ��ʼ���" & vbNewLine & _
            "From ��Ա�� A, ������Ա B, ��Ա����˵�� D,���ű� E" & vbNewLine & _
            "Where a.Id = b.��Աid And e.ID=b.����ID And d.��Աid = a.Id  And (d.��Ա���� = 'ҽ��' Or d.��Ա���� = '��ʿ') And " & vbNewLine & _
            "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) and " & IIf(Val(cboDocDept.ItemData(cboDocDept.ListIndex)) = -1, "b.ȱʡ=1 ", "b.����id=[1]")

    On Error GoTo errH

    rptDoc.Records.DeleteAll

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(cboDocDept.ItemData(cboDocDept.ListIndex)))

    With rptDoc
        Do While Not rsTmp.EOF
            Set objRecord = .Records.Add()
            Set objItem = objRecord.AddItem(rsTmp!ID & "")
            Set objItem = objRecord.AddItem("")
                objItem.Icon = img16.ListImages("unCheck").Index - 1
            Set objItem = objRecord.AddItem(rsTmp!��� & "")
            Set objItem = objRecord.AddItem(rsTmp!���� & "")
                objItem.Icon = img16.ListImages.Item(IIf(rsTmp!�Ա� & "" = "Ů", "feMale", "Male")).Index - 1
            Set objItem = objRecord.AddItem(rsTmp!�Ա� & "")
            Set objItem = objRecord.AddItem(rsTmp!ƴ������ & "")
            Set objItem = objRecord.AddItem(rsTmp!��ʼ��� & "")
            Set objItem = objRecord.AddItem(rsTmp!�������� & "")
            Set objItem = objRecord.AddItem(rsTmp!����ID & "")
            rsTmp.MoveNext
        Loop
        .Populate
    End With
    mlngFindNum = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlcontrol.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtHlInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtHlInfo.hwnd, Replace(txtHlInfo.Text, ",", "��" & vbCrLf), True, True)
End Sub

Private Sub txtJcbgTpye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtJcbgTpye.hwnd, Replace(txtJcbgTpye.Text, ",", "��" & vbCrLf), True, True)
End Sub

Private Sub txtJybgTpye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtJybgTpye.hwnd, Replace(txtJybgTpye.Text, ",", "��" & vbCrLf), True, True)
End Sub


Private Sub InitPatiTable()
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "����ID;����,1300,1;�Ա�,700,4;����,700,4;��ʶ��,950,1;����,1000,1;��ǰ״̬,1700,1"
    arrHead = Split(strHead, ";")
    With vsPati
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusLight
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &HFAEADA

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
    End With
End Sub


Private Sub txtName_GotFocus()
    Call zlcontrol.TxtSelAll(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub txtReault_GotFocus()
    Call zlcontrol.TxtSelAll(txtReault)
End Sub

Private Sub txtReault_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub rptDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptDoc.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptDocCheck(rptDoc.SelectedRows(0), rptDoc.SelectedRows(0).Record.Item(COLD_ѡ��))
        End If
    End If
End Sub

Private Sub rptDoc_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim objHitTest As ReportHitTestInfo
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptDoc.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptDoc.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COLD_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptDoc.Columns(COLD_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(COLD_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                            rptDoc.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptDoc.Columns(COLD_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptDoc.Records.Count - 1
                            rptDoc.Records(i)(COLD_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                            rptDoc.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        ElseIf rptDoc.HitTest(X, Y).ht = xtpHitTestReportArea Then
            Set objHitTest = rptDoc.HitTest(X, Y)
            If Not objHitTest.Column Is Nothing And Not objHitTest.Row Is Nothing Then
                If objHitTest.Column.Index = COLD_ѡ�� Then
                    If rptDoc.SelectedRows.Count > 0 Then
                        Call rptDocCheck(objHitTest.Row, objHitTest.Row.Record.Item(COLD_ѡ��))
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptDocCheck(Row As XtremeReportControl.IReportRow, Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COLD_ѡ��).Icon = img16.ListImages.Item("unCheck").Index - 1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COLD_ѡ��).Icon = img16.ListImages.Item("AllCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptDoc.Populate
End Sub

Private Sub rptDoc_SelectionChanged()
    Dim i As Long, j As Long, blnDo As Boolean
    If mlngFindNum <> 0 Then mlngFindNum = rptDoc.SelectedRows(0).Index + 1
    
    
    If rptDoc.Rows.Count <= 0 Then Exit Sub
    For i = 0 To rptDoc.Rows.Count - 1
        For j = 0 To rptDoc.Columns.Count - 1
            If rptDoc.Rows(i).Record.Item(j).Bold Then
                rptDoc.Rows(i).Record.Item(j).Bold = False
                rptDoc.Rows(i).Record.Item(j).BackColor = rptDoc.PaintManager.BackColor
                blnDo = True
            End If
        Next
    Next
    If blnDo Then
        blnDo = False
        rptDoc.Redraw
    End If
    
    For i = 0 To rptDoc.Columns.Count - 1
       rptDoc.SelectedRows(0).Record.Item(i).Bold = True
       rptDoc.SelectedRows(0).Record.Item(i).BackColor = RGB(153, 204, 255)
    Next

End Sub

Private Sub rptDoc_SortOrderChanged()
    mlngFindNum = 0
End Sub


Private Sub txtDocFind_Change()
    mlngFindNum = 0
End Sub

Private Sub txtDocFind_GotFocus()
    If txtDocFind.Text <> "" Then
        Call zlcontrol.TxtSelAll(txtDocFind)
    End If
End Sub

Private Sub txtDocFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDoc
        strMsg = UCase(Trim(txtDocFind.Text))
        If ZLCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindNum To rptDoc.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(COLD_����).Value Like "*" & strMsg & "*" Or _
                            .Rows(i).Record(IIf(0 = 0, COLD_ƴ������, COLD_��ʼ���)).Value Like "*" & strMsg & "*" Then
                        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        rptDoc.SelectedRows(0).Selected = False
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(COLD_����).Value Like "*" & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindNum = i + 1
                        blnIsFind = True
                        rptDoc.SelectedRows(0).Selected = False
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindNum = 0 Then
            MsgBox "��ǰ����û���ҵ������ҵ���Ա��", vbInformation, Me.Caption
        ElseIf mlngFindNum <> 0 And blnIsFind = False Then
            MsgBox "�Ѿ������һ����Ա�ˡ�", vbInformation, Me.Caption
            mlngFindNum = 0
        End If
    End With
End Sub



Private Sub rptDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If rptDept.SelectedRows.Count > 0 Then
        If KeyCode = vbKeySpace Then
            Call rptDeptCheck(rptDept.SelectedRows(0), rptDept.SelectedRows(0).Record.Item(COLB_ѡ��))
        End If
    End If
End Sub

Private Sub rptDept_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim objHitTest As ReportHitTestInfo
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptDept.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptDept.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COLB_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        rptDept.Columns(COLB_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                        For i = 0 To rptDept.Records.Count - 1
                            rptDept.Records(i)(COLB_ѡ��).Icon = img16.ListImages("AllCheck").Index - 1
                            rptDept.Records(i).Tag = "1"
                        Next
                    Else
                        objColumn.Caption = ""
                        rptDept.Columns(COLB_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                        For i = 0 To rptDept.Records.Count - 1
                            rptDept.Records(i)(COLB_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
                            rptDept.Records(i).Tag = "0"
                        Next
                    End If
                End If
            End If
        ElseIf rptDept.HitTest(X, Y).ht = xtpHitTestReportArea Then
            Set objHitTest = rptDept.HitTest(X, Y)
            If Not objHitTest.Column Is Nothing And Not objHitTest.Row Is Nothing Then
                If objHitTest.Column.Index = COLB_ѡ�� Then
                    If rptDept.SelectedRows.Count > 0 Then
                        Call rptDeptCheck(objHitTest.Row, objHitTest.Row.Record.Item(COLB_ѡ��))
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptDeptCheck(Row As XtremeReportControl.IReportRow, Item As XtremeReportControl.IReportRecordItem)
    If Row.Record.Tag = "1" Then
        Row.Record.Item(COLB_ѡ��).Icon = img16.ListImages("unCheck").Index - 1
        Row.Record.Tag = "0"
    Else
        Row.Record.Item(COLB_ѡ��).Icon = img16.ListImages.Item("AllCheck").Index - 1
        Row.Record.Tag = "1"
    End If
    rptDept.Populate
End Sub

Private Sub rptDept_SelectionChanged()
    Dim i As Long, j As Long, blnDo As Boolean
    If mlngFindDept <> 0 Then mlngFindDept = rptDept.SelectedRows(0).Index + 1
    
    
    If rptDept.Rows.Count <= 0 Or rptDept.SelectedRows.Count <= 0 Then Exit Sub
    For i = 0 To rptDept.Rows.Count - 1
        For j = 0 To rptDept.Columns.Count - 1
            If rptDept.Rows(i).Record.Item(j).Bold Then
                rptDept.Rows(i).Record.Item(j).Bold = False
                rptDept.Rows(i).Record.Item(j).BackColor = rptDept.PaintManager.BackColor
                blnDo = True
            End If
        Next
    Next
    If blnDo Then
        blnDo = False
        rptDept.Redraw
    End If
    
    For i = 0 To rptDept.Columns.Count - 1
       rptDept.SelectedRows(0).Record.Item(i).Bold = True
       rptDept.SelectedRows(0).Record.Item(i).BackColor = RGB(153, 204, 255)
    Next
End Sub

Private Sub rptDept_SortOrderChanged()
    mlngFindDept = 0
End Sub


Private Sub txtDeptFind_Change()
    mlngFindDept = 0
End Sub

Private Sub txtDeptFind_GotFocus()
    If txtDeptFind.Text <> "" Then
        Call zlcontrol.TxtSelAll(txtDeptFind)
    End If
End Sub

Private Sub txtDeptFind_KeyPress(KeyAscii As Integer)
    Dim strMsg As String
    Dim i As Long
    Dim blnIsAllChar As Boolean
    Dim blnIsFind As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    With rptDept
        strMsg = UCase(Trim(txtDeptFind.Text))
        If ZLCommFun.IsCharAlpha(strMsg) Then blnIsAllChar = True
        
        For i = mlngFindDept To rptDept.Rows.Count - 1
            If Not .Rows(i).GroupRow Then
                If blnIsAllChar Then
                    If .Rows(i).Record(COLB_����).Value Like "*" & strMsg & "*" Or _
                            .Rows(i).Record(COLB_����).Value Like "*" & strMsg & "*" Then
                        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
                        Set .FocusedRow = .Rows(i)
                        mlngFindDept = i + 1
                        blnIsFind = True
                        rptDept.SelectedRows(0).Selected = False
                        Exit Sub
                    End If
                Else
                    If .Rows(i).Record(COLB_����).Value Like "*" & strMsg & "*" Then
                        Set .FocusedRow = .Rows(i)
                        mlngFindDept = i + 1
                        blnIsFind = True
                        rptDept.SelectedRows(0).Selected = False
                        Exit Sub
                    End If
                End If
            End If
        Next
        If mlngFindDept = 0 Then
            MsgBox "��ǰ����û���ҵ������ҵĲ��š�", vbInformation, Me.Caption
        ElseIf mlngFindDept <> 0 And blnIsFind = False Then
            MsgBox "�Ѿ������һ�������ˡ�", vbInformation, Me.Caption
            mlngFindDept = 0
        End If
    End With
End Sub



Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    lblDept.Tag = Control.Parameter
    lblDept.Caption = decode(lblDept.Tag, "", "�����˿���", "����", "���������", "סԺ", "��סԺ����")
    Call LoadPatiDept
    
    'ִ�н�������˵���ʼ��
    cboFind.Clear
    cboFind.AddItem "����"
    cboFind.AddItem "���֤��"
    cboFind.AddItem "����ID"
    If lblDept.Tag = "" Or lblDept.Tag = "����" Then
        cboFind.AddItem "�����"
    End If
    
    If lblDept.Tag = "" Or lblDept.Tag = "סԺ" Then
        cboFind.AddItem "סԺ��"
    End If
    
    cboFind.ListIndex = 0
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Control.Checked = Control.Parameter = lblDept.Tag
End Sub


Private Sub lblDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSQL As String
    Dim str��λ As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, Cmd_���п���, "���п���")
        objControl.Parameter = ""
        Set objControl = .Add(xtpControlButton, Cmd_סԺ����, "סԺ����")
        objControl.Parameter = "סԺ"
        Set objControl = .Add(xtpControlButton, Cmd_�������, "�������")
        objControl.Parameter = "����"
    End With
    GetWindowRect picPati.hwnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lblDept.Left + lblDept.Width, vRect.Top * Screen.TwipsPerPixelY + lblDept.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLBLFace(ByRef objCtl As Object, ByVal blnOver As Boolean)
    If blnOver Then
        If objCtl.BorderStyle = 0 Then
            objCtl.BorderStyle = 1
            objCtl.BackStyle = 1
        End If
    Else
        If objCtl.BorderStyle = 1 Then
            objCtl.BorderStyle = 0
            objCtl.BackStyle = 0
        End If
    End If
End Sub

