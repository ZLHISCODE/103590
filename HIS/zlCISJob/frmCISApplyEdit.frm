VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISApplyEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   Icon            =   "frmCISApplyEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7320
   StartUpPosition =   1  '����������
   Begin MSComctlLib.ImageList img16 
      Left            =   2040
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":6852
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":D0B4
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":13916
            Key             =   "����ʱ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":13EB0
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":1444A
            Key             =   "����ҽ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":149E4
            Key             =   "���ʲ���"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":14F7E
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":150D8
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   52
      Top             =   360
      Width           =   7300
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
         Height          =   7335
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   7095
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6975
            Left            =   120
            ScaleHeight     =   6975
            ScaleWidth      =   6900
            TabIndex        =   54
            Top             =   240
            Width           =   6900
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   4680
               TabIndex        =   2
               Top             =   120
               Width           =   1815
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
               TabIndex        =   56
               Top             =   120
               Width           =   1170
               Begin VB.ComboBox cboFind 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   -30
                  Width           =   1215
               End
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1065
               ScaleHeight     =   240
               ScaleWidth      =   1905
               TabIndex        =   55
               Top             =   130
               Width           =   1935
               Begin VB.ComboBox cboDept 
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   -30
                  Width           =   1960
               End
            End
            Begin VB.CommandButton cmdAdd 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISApplyEdit.frx":15232
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   480
               Width           =   330
            End
            Begin VB.CommandButton cmdDel 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISApplyEdit.frx":1BA84
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   960
               Width           =   330
            End
            Begin VSFlex8Ctl.VSFlexGrid vsPati 
               Height          =   6435
               Left            =   0
               TabIndex        =   3
               Top             =   480
               Width           =   6525
               _cx             =   11509
               _cy             =   11351
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
               MouseIcon       =   "frmCISApplyEdit.frx":222D6
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
               FormatString    =   $"frmCISApplyEdit.frx":22BB0
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
                  TabIndex        =   57
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
               TabIndex        =   58
               Top             =   160
               Width           =   900
            End
            Begin VB.Image imgSentence 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   2970
               Picture         =   "frmCISApplyEdit.frx":22C4B
               ToolTipText     =   "��ʾ��ǰѡ���������Ĳ���"
               Top             =   90
               Width           =   360
            End
         End
      End
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
      TabIndex        =   40
      Top             =   360
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
         TabIndex        =   41
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
            TabIndex        =   11
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
            TabIndex        =   47
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
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   20
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtHlInfo 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   720
               Width           =   4095
            End
            Begin VB.Image imgHlInfo 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISApplyEdit.frx":23335
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
            TabIndex        =   46
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
               TabIndex        =   27
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
               TabIndex        =   28
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtJybgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   4095
            End
            Begin VB.Image imgJybgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISApplyEdit.frx":23A1F
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
            TabIndex        =   45
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
               TabIndex        =   23
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
               TabIndex        =   24
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   525
               Width           =   4095
            End
            Begin VB.Image imgJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":24109
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
            TabIndex        =   44
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
               TabIndex        =   12
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
               TabIndex        =   13
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
               TabIndex        =   15
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
               TabIndex        =   14
               Top             =   480
               Width           =   4095
            End
            Begin VB.TextBox txtDzblTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   885
               Width           =   4095
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   0
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":247F3
               ToolTipText     =   "ѡ�񱾿�������Ĳ���"
               Top             =   435
               Width           =   360
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   1
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":24EDD
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
            TabIndex        =   9
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
            TabIndex        =   26
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
            TabIndex        =   22
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
            TabIndex        =   7
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
            TabIndex        =   8
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
            TabIndex        =   10
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
            TabIndex        =   17
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
         TabIndex        =   6
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
         TabIndex        =   42
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
            Picture         =   "frmCISApplyEdit.frx":255C7
            ScaleHeight     =   1410
            ScaleWidth      =   3435
            TabIndex        =   43
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
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   48
      Top             =   360
      Width           =   7300
      Begin VB.Frame fraReault 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "����ԭ��"
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
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   6360
         Width           =   7095
         Begin VB.TextBox txtReault 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   6855
         End
      End
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
         Height          =   6135
         Left            =   120
         TabIndex        =   49
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
            Top             =   1800
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   1635
            TabIndex        =   30
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   214892547
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   3960
            TabIndex        =   31
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   214892547
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
            TabIndex        =   51
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   50
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
      Left            =   5880
      TabIndex        =   38
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   4320
      TabIndex        =   37
      Top             =   8040
      Width           =   1215
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7980
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   7330
      _Version        =   589884
      _ExtentX        =   12929
      _ExtentY        =   14076
      _StockProps     =   64
   End
   Begin VB.Image imtmp 
      Height          =   360
      Left            =   120
      Picture         =   "frmCISApplyEdit.frx":273DF
      Stretch         =   -1  'True
      Top             =   8040
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
      Caption         =   "������������"
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   39
      Top             =   8160
      Width           =   1080
   End
End
Attribute VB_Name = "frmCISApplyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '=0ʱΪ��������,=1ʱΪ�޸�����
Private mrsPati As ADODB.Recordset
Private mblnOk As Boolean
Private mlngApplyID As Long

Private mstrNewEMR As String


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

Public Function ShowEdit(frmParent As Object, ByVal intType As String, ByRef lngApplyID As Long, Optional ByRef rsPati As ADODB.Recordset) As Boolean
'���ܣ������������ݱ༭��
'rsPati ��������¼��
    On Error Resume Next
    mintType = intType
    Set mrsPati = rsPati
    mblnOk = False
    mlngApplyID = lngApplyID
    
    If mlngApplyID = 0 And mintType = 1 Then Exit Function
    Me.Show 1, frmParent
    lngApplyID = mlngApplyID
    ShowEdit = mblnOk
    On Error GoTo 0
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    lblDept.Tag = Control.Parameter
    lblDept.Caption = decode(lblDept.Tag, "", "�����˿���", "����", "���������", "סԺ", "��סԺ����")
    Call LoadDept
    
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
    Dim strXML As String
    Dim lngID As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim i As Long
    Dim curDate As Date
    Dim blnTran As Boolean
    Dim lngTmp As Long
    
    
    On Error GoTo errH
    '��ȡ���ʲ���
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, col_����Id)) <> 0 Then
            str����ids = str����ids & "," & Val(vsPati.TextMatrix(i, col_����Id))
            str�������� = str�������� & "," & Val(vsPati.TextMatrix(i, col_����))
        End If
    Next
    str����ids = Mid(str����ids, 2)
    str�������� = Mid(str��������, 2)
    
    If str����ids = "" Then
        Me.tbcSub.Item(0).Selected = True
        MsgBox "��ǰ��δ¼����Ҫ������ʲ����Ĳ�����Ϣ,������¼�롣", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If chkAllInfo.Value = 0 Then
        For i = 0 To 7
            If chkInfo(i).Value = 1 Then
                lngTmp = lngTmp + 1
            End If
        Next
        If lngTmp = 0 Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "��ǰ��δ¼����Ҫ������ʲ�����Ȩ������,������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '����������
    For i = 0 To 1
        If txtDzblTpye(i).BackColor = &HFFFFFF And txtDzblTpye(i).Text = "" And chkAllInfo.Value = 0 Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "��ǰ��δ¼�벡���ļ�" & IIf(i = 0, "����", "") & ",������¼��!!!", vbInformation, gstrSysName
            txtDzblTpye(i).SetFocus
            Exit Sub
        End If
    Next
    If txtHlInfo.BackColor = &HFFFFFF And txtHlInfo.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "��ǰ��δ¼�뻤���¼�ļ�,������¼�롣", vbInformation, gstrSysName
        txtHlInfo.SetFocus
        Exit Sub
    End If
    If txtJcbgTpye.BackColor = &HFFFFFF And txtJcbgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "��ǰ��δ¼���鱨������,������¼�롣", vbInformation, gstrSysName
        txtJcbgTpye.SetFocus
        Exit Sub
    End If
    If txtJybgTpye.BackColor = &HFFFFFF And txtJybgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "��ǰ��δ¼����鱨������,������¼�롣", vbInformation, gstrSysName
        txtJybgTpye.SetFocus
        Exit Sub
    End If
    
    '������ԭ��
    If txtReault.Text = "" Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "��ǰ��δ¼�����ԭ��,������¼�롣", vbInformation, gstrSysName
        txtReault.SetFocus
        Exit Sub
    End If
    
    If ZLCommFun.ActualLen(txtReault.Text) > txtReault.MaxLength Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "����ԭ�����ݹ��࣬������� " & txtReault.MaxLength \ 2 & " �����ֻ� " & txtReault.MaxLength & " ���ַ���", vbInformation, gstrSysName
        txtReault.SetFocus: Exit Sub
    End If
    
    
    '������ʱ��
    If dtpTime(0).Value >= dtpTime(1).Value Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "��ǰ������ʼʱ�����С����ֹʱ��,������¼�롣", vbInformation, gstrSysName
        txtReault.SetFocus
        Exit Sub
    End If
    
    strXML = GetInfoXml
    
    '��������
    lngID = mlngApplyID
    If lngID = 0 Then lngID = zlDatabase.GetNextId("���Ӳ�����������")
    curDate = zlDatabase.Currentdate
    strSQL = "Zl_���Ӳ�����������_Update(" & mintType & "," & lngID & ",'" & strXML & "',To_Date('" & Format(dtpTime(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            "To_Date('" & Format(dtpTime(1).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            IIf(optTimeTpye(0).Value, 0, IIf(optTimeTpye(1).Value, 1, 2)) & ",'" & Replace(txtReault.Text, "'", "") & _
                            "','" & UserInfo.���� & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    arrSQL = Array()
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, col_����Id)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_���Ӳ���������ʲ���_Insert(" & lngID & "," & Val(vsPati.TextMatrix(i, col_����Id)) & ")"
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
    
    '����xml����
'    Call Sys.SaveLobV2("���Ӳ�����������", "��������", "ID =[1]", "", mlngApplyID)
    
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
    '��ȡ�������ݵ�Xml������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String

    
    On Error GoTo errH
    If mlngApplyID = 0 Then Exit Function
    
    strXML = Sys.ReadXML("���Ӳ�����������", "��������", "ID=[1]", strErr, mlngApplyID)
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
        '������ҳ��ҽ�����ٴ�·�����������
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
            .appendData "patipeis", chkInfo(File_���).Value            '<�������>���ͣ�N
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
    Me.Caption = IIf(mintType = 0, "������������", "�޸ķ�������")
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
        .InsertItem(0, "���ò���", picParent.hwnd, 0).Tag = "���ò���"
        .InsertItem(1, "��������", picAppInfo.hwnd, 0).Tag = "��������"
        .InsertItem(2, "����ʱ�޺�ԭ��", picTime.hwnd, 0).Tag = "����ʱ��"
        
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call LoadDept
    
    '��ʼ�����˱��
    Call InitPatiTable
    
     
    'ִ�н�������˵���ʼ��
    cboFind.Clear
    cboFind.AddItem "����"
    cboFind.AddItem "���֤��"
    cboFind.AddItem "�����"
    cboFind.AddItem "סԺ��"
    cboFind.AddItem "����ID"
    cboFind.ListIndex = 0

    
    If mintType = 1 Then
        Call LoadPati
        Call ReadXmlSet
        Call SetFileCtl
        Call LoadOther
    Else
        '�����湴ѡ����
        If Not mrsPati Is Nothing Then
            If mrsPati.State <> 0 Then
                If Not mrsPati.EOF Then
                    Call LoadPati
                End If
            End If
        End If
        chkAllInfo.Value = 1
        curDate = zlDatabase.Currentdate
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd hh:mm")
        dtpTime(1).Value = Format(curDate + 7, "yyyy-MM-dd hh:mm")
        optTimeTpye(0).Value = True
        SetFileCtl
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadOther()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.����ԭ�� From ���Ӳ����������� A Where a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            dtpTime(0).Value = Format(rsTmp!���ʿ�ʼʱ�� & "", "yyyy-MM-dd hh:mm")
            dtpTime(1).Value = Format(rsTmp!���ʽ���ʱ�� & "", "yyyy-MM-dd hh:mm")
            optTimeTpye(Val(rsTmp!����ʱ�� & "")).Value = True
            txtReault.Text = rsTmp!����ԭ�� & ""
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
        If mrsPati Is Nothing Then
            strSQL = "Select d.Id, d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��, d.��ǰ״̬" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '����' As ����, a.����id As ID, a.����, a.�Ա�, a.����, a.ִ�в���id As ����, a.ִ��ʱ�� As ����ʱ��, a.����� As ��ʶ��," & vbNewLine & _
                    "                     Decode(a.ִ��״̬, 1, '��' || To_Char(a.ִ��ʱ��, 'yyyy-mm-dd') || '���������Ժ', '�������ھ���') As ��ǰ״̬" & vbNewLine & _
                    "              From ���˹Һż�¼ A, ���Ӳ���������ʲ��� G" & vbNewLine & _
                    "              Where g.����id = a.����id And g.����id = [1] And ��¼״̬ = 1" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��," & vbNewLine & _
                    "                     Decode(b.��Ժ����, Null, '��Ժ', '��' || b.��ҳid || '��סԺ��Ժ') As ��ǰ״̬" & vbNewLine & _
                    "              From ������ҳ B, ���Ӳ���������ʲ��� H" & vbNewLine & _
                    "              Where h.����id = b.����id And h.����id = [1]) C) D, ���ű� G" & vbNewLine & _
                    "Where g.Id = d.���� And d.Top = 1" & vbNewLine & _
                    "Order By d.����ʱ�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
        Else
            Set rsTmp = mrsPati
            If Not rsTmp Is Nothing Then rsTmp.MoveFirst
        End If
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                    If Val(.TextMatrix(.Rows - 1, col_����Id)) <> 0 Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    
                    .TextMatrix(lngRow, col_����Id) = rsTmp!ID & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    Set .Cell(flexcpPicture, lngRow, col_����) = img16.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
                    .TextMatrix(lngRow, col_�Ա�) = rsTmp!�Ա� & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    .TextMatrix(lngRow, COL_��ʶ��) = rsTmp!��ʶ�� & ""
                    .TextMatrix(lngRow, col_����) = rsTmp!���� & ""
                    .TextMatrix(lngRow, COL_��ǰ״̬) = rsTmp!��ǰ״̬ & ""
                    .Tag = .Tag & "," & rsTmp!ID & ","
                End If
                rsTmp.MoveNext
            Loop
        End If
        .WordWrap = True
        '�Զ������и�
        .AutoSize col_����, COL_��ǰ״̬
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set mrsPati = Nothing
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

        strSQL = strSQL & " Select a.ID,Decode(a.����, 1, '���ﲡ��', 2, 'סԺ����', 4, '������', 5, '����֤��', 6, '֪���ļ�') As ��������, a.����," & vbNewLine & _
                "       Decode(b.C2, Null, 0, 1) As �ѹ�ѡcheck" & vbNewLine & _
                "From �����ļ��б� A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
                "Where a.���� In (1, 2, 4, 5, 6) And a.���� = b.C2(+)" & vbNewLine & _
                "Order By ����, ���"

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
                    "              Where  b.��Ժ���� Between Sysdate - 7 And Sysdate) C) D, ���ű� G" & vbNewLine & _
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
                    "              Where  b.��Ժ���� Between Sysdate - 7 And Sysdate) C) D, ���ű� G" & vbNewLine & _
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
                        Set .Cell(flexcpPicture, lngRow, col_����) = img16.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
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
    MsgBox "δ���ҵ������ҽ��ڵ�" & lblDept.Tag & "����!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, True)
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

Private Sub picPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, False)
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
                            "              Where " & decode(cboFind.Text, "���֤��", " b.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[2]", "����", "B.���� like [2]") & ") C) D, ���ű� G" & vbNewLine & _
                            "Where g.Id = d.���� And d.Top = 1" & IIf(cboDept.Text = "���в���", "", " And D.����=[1]") & vbNewLine & _
                            "Order By d.����ʱ�� Desc"
            ElseIf lblDept.Tag = "סԺ" Then
                strSQL = "Select d.Id,d.����, d.����, d.�Ա�, d.����, g.���� As ����, d.��ʶ��,d.��ǰ״̬" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By ����ʱ�� Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select 'סԺ' As ����, b.����id As ID, b.����, b.�Ա�, b.����, b.��Ժ����id As ����, b.��Ժ���� As ����ʱ��, b.סԺ�� As ��ʶ��,decode(B.��Ժ����,null,'��Ժ','��'||b.��ҳid||'��סԺ��Ժ') as ��ǰ״̬" & vbNewLine & _
                            "              From ������ҳ B" & vbNewLine & _
                            "              Where " & decode(cboFind.Text, "���֤��", " b.����ID in (Select Column_Value As ����id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "����ID", "B.����ID =[2]", "����", "B.���� like [2]") & ") C) D, ���ű� G" & vbNewLine & _
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
                            Set .Cell(flexcpPicture, lngRow, col_����) = img16.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
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
                    MsgBox "�ڵ�ǰ��Χδ���ҵ�" & lblDept.Tag & "����!", vbInformation, Me.Caption
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
    MsgBox "�ڵ�ǰ��Χδ���ҵ�" & lblDept.Tag & "����!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDept()
'���ز�ѯ����
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


Private Sub txtReault_GotFocus()
    Call zlcontrol.TxtSelAll(txtReault)
End Sub

Private Sub txtReault_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
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
