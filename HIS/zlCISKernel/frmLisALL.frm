VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLisALL 
   Caption         =   "���鱨�����"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18135
   Icon            =   "frmLisALL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   18135
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   6480
      ScaleHeight     =   1695
      ScaleWidth      =   2295
      TabIndex        =   50
      Top             =   240
      Width           =   2295
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   735
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   615
         _Version        =   589884
         _ExtentX        =   1085
         _ExtentY        =   1296
         _StockProps     =   0
         AutoColumnSizing=   0   'False
      End
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   4545
      Top             =   915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picWB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   12810
      ScaleHeight     =   1455
      ScaleWidth      =   2760
      TabIndex        =   48
      Top             =   195
      Width           =   2760
      Begin SHDocVwCtl.WebBrowser webSub 
         Height          =   1320
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   2010
         ExtentX         =   3545
         ExtentY         =   2328
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   915
      Left            =   525
      TabIndex        =   0
      Top             =   720
      Width           =   1110
      _Version        =   589884
      _ExtentX        =   1958
      _ExtentY        =   1614
      _StockProps     =   64
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7590
      Left            =   645
      ScaleHeight     =   7590
      ScaleWidth      =   15240
      TabIndex        =   1
      Top             =   1785
      Width           =   15240
      Begin VB.Frame fraLR 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2745
         Left            =   5760
         MousePointer    =   9  'Size W E
         TabIndex        =   6
         Top             =   1320
         Width           =   45
      End
      Begin VB.Frame fraRpt 
         Height          =   5115
         Left            =   4050
         TabIndex        =   5
         Top             =   1785
         Width           =   7650
         Begin XtremeSuiteControls.TabControl tbcArchive 
            Height          =   1290
            Left            =   915
            TabIndex        =   47
            Top             =   570
            Width           =   2055
            _Version        =   589884
            _ExtentX        =   3625
            _ExtentY        =   2275
            _StockProps     =   64
         End
      End
      Begin VB.Frame fraPList 
         Height          =   5250
         Left            =   510
         TabIndex        =   4
         Top             =   1305
         Width           =   5235
         Begin XtremeSuiteControls.TabControl tbcHistory 
            Height          =   1395
            Left            =   345
            TabIndex        =   46
            Top             =   1365
            Width           =   855
            _Version        =   589884
            _ExtentX        =   1508
            _ExtentY        =   2461
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   480
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   3
         Top             =   105
         Width           =   1425
         Begin XtremeCommandBars.CommandBars cbsMain 
            Left            =   0
            Top             =   0
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
      Begin MSComctlLib.StatusBar stbThis 
         Bindings        =   "frmLisALL.frx":6852
         Height          =   360
         Left            =   1290
         TabIndex        =   2
         Top             =   7080
         Width           =   12450
         _ExtentX        =   21960
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   2
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Bevel           =   2
               Object.Width           =   2355
               MinWidth        =   882
               Picture         =   "frmLisALL.frx":6866
               Text            =   "�������"
               TextSave        =   "�������"
               Key             =   "ZLFLAG"
               Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   1
               Object.Width           =   19526
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
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   3990
         ScaleHeight     =   975
         ScaleWidth      =   7695
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   495
         Width           =   7695
         Begin VB.Frame fraInfo 
            Caption         =   " ����������Ϣ "
            Height          =   840
            Left            =   90
            TabIndex        =   8
            Top             =   60
            Width           =   7500
            Begin VB.Frame fraIn 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   195
               TabIndex        =   27
               Top             =   255
               Visible         =   0   'False
               Width           =   7170
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   4770
                  TabIndex        =   45
                  Top             =   0
                  Width           =   1080
               End
               Begin VB.Label lbl����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   4305
                  TabIndex        =   44
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lblסԺ��zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "סԺ��:"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   43
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   42
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   41
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   40
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lblҽ����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ҽ����:"
                  Height          =   180
                  Index           =   0
                  Left            =   5940
                  TabIndex        =   39
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lbl��Ժzy 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��Ժ:"
                  Height          =   180
                  Index           =   0
                  Left            =   4305
                  TabIndex        =   38
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl����zy 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   3150
                  TabIndex        =   37
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl����zy 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��  ��:"
                  Height          =   180
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   36
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   35
                  Top             =   255
                  Width           =   900
               End
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H000000FF&
                  Height          =   180
                  Index           =   1
                  Left            =   3585
                  TabIndex        =   34
                  Top             =   255
                  Width           =   675
               End
               Begin VB.Label lbl��Ժzy 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   4770
                  TabIndex        =   33
                  Top             =   255
                  Width           =   90
               End
               Begin VB.Label lblҽ����zy 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   6600
                  TabIndex        =   32
                  Top             =   0
                  Width           =   90
               End
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3585
                  TabIndex        =   31
                  Top             =   0
                  Width           =   675
               End
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   435
                  TabIndex        =   30
                  Top             =   255
                  Width           =   1080
               End
               Begin VB.Label lbl����zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   435
                  TabIndex        =   29
                  Top             =   0
                  Width           =   1080
               End
               Begin VB.Label lblסԺ��zy 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2190
                  TabIndex        =   28
                  Top             =   0
                  Width           =   900
               End
            End
            Begin VB.Frame fraOut 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   450
               Left            =   195
               TabIndex        =   9
               Top             =   240
               Visible         =   0   'False
               Width           =   7170
               Begin VB.Label lbl�� 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "��"
                  BeginProperty Font 
                     Name            =   "����"
                     Size            =   21.75
                     Charset         =   134
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   435
                  Left            =   6705
                  TabIndex        =   26
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Label lbl�Һŵ�mz 
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3870
                  TabIndex        =   25
                  Top             =   0
                  Width           =   1065
               End
               Begin VB.Label lbl�Һŵ�mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�Һŵ�:"
                  Height          =   180
                  Index           =   0
                  Left            =   3255
                  TabIndex        =   24
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lblҽ��mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2385
                  TabIndex        =   23
                  Top             =   0
                  Width           =   780
               End
               Begin VB.Label lblҽ��mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ҽ��:"
                  Height          =   180
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   22
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl������mz 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   5655
                  TabIndex        =   21
                  Top             =   255
                  Width           =   90
               End
               Begin VB.Label lbl������mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "������:"
                  Height          =   180
                  Index           =   0
                  Left            =   5025
                  TabIndex        =   20
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl�����mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�����:"
                  Height          =   180
                  Index           =   0
                  Left            =   3240
                  TabIndex        =   19
                  Top             =   255
                  Width           =   630
               End
               Begin VB.Label lbl����mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   18
                  Top             =   0
                  Width           =   450
               End
               Begin VB.Label lbl�ѱ�mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "�ѱ�:"
                  Height          =   180
                  Index           =   0
                  Left            =   1935
                  TabIndex        =   17
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lblҽ����mz 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ҽ����:"
                  Height          =   180
                  Index           =   0
                  Left            =   5025
                  TabIndex        =   16
                  Top             =   0
                  Width           =   630
               End
               Begin VB.Label lblҽ����mz 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00008000&
                  Height          =   180
                  Index           =   1
                  Left            =   5655
                  TabIndex        =   15
                  Top             =   0
                  Width           =   90
               End
               Begin VB.Label lbl�ѱ�mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   2385
                  TabIndex        =   14
                  Top             =   255
                  Width           =   765
               End
               Begin VB.Label lbl����mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   450
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1425
               End
               Begin VB.Label lbl�����mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   3870
                  TabIndex        =   12
                  Top             =   255
                  Width           =   1095
               End
               Begin VB.Label lbl����mz 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "����:"
                  Height          =   180
                  Index           =   0
                  Left            =   0
                  TabIndex        =   11
                  Top             =   255
                  Width           =   450
               End
               Begin VB.Label lbl����mz 
                  BackColor       =   &H00C0FFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "#"
                  ForeColor       =   &H00800000&
                  Height          =   180
                  Index           =   1
                  Left            =   450
                  TabIndex        =   10
                  Top             =   255
                  Width           =   1455
               End
            End
         End
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5040
      Top             =   240
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
            Picture         =   "frmLisALL.frx":70FA
            Key             =   "δִ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":7694
            Key             =   "��ִ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":7C2E
            Key             =   "�ܾ�ִ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":81C8
            Key             =   "����ִ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":8762
            Key             =   "�ѱ���"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":8CFC
            Key             =   "CheckCol"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9296
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9830
            Key             =   "�Ѻ˶�"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":9DCA
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2880
      Top             =   480
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
            Picture         =   "frmLisALL.frx":1062C
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":16E8E
            Key             =   "object_report"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":17428
            Key             =   "object_case"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":179C2
            Key             =   "object_tend"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":17F5C
            Key             =   "object_first"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":184F6
            Key             =   "object_advice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":18A90
            Key             =   "object_file"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":1902A
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLisALL.frx":1F88C
            Key             =   "Path"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLisALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum PATIREPORT_COLUMN
    COL_ѡ��
    COL_ͼ��
    COL_��Ŀ '��������
    COL_���ʱ��
    COL_�걾
    COL_����ʱ��
    COL_������
    COL_�ɼ�ʱ��
    
    '������
    COL_ҽ��ID
    COL_����ID
    COL_����
    COL_��ӡ����
    COL_�ĵ�����
    COL_��ֹ��ӡ
End Enum

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fmask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private mblnMoved As Boolean
Private mlng����ID As Long
Private mlng����ID As Long
Private mlngPreIndex As Long
Private mrsData As ADODB.Recordset
Private mlngPreDept As Long
Private mstr�Һŵ� As String
Private mlngModel As Long '1252-p����ҽ���´1253-pסԺҽ���´�
Private mstrMainPrivs As String
Private mlng����ID  As Long
Private mlng����ID  As Long
Private mfrmLisCom As Object 'LIS����ӿڷ��صĴ��壬�����°���ϰ�
Private mstrFilesTemp As String '�������ɵ���ʱ�ļ�Ŀ¼����<STAB>�ָ�
Private mstrCurFile As String '��ǰ��ҳ�ؼ�ռ�õ��ļ�Ŀ¼
Private mobjPopup As CommandBarPopup
Private mlngPages As Long
Private Const M_S As String = "<SPL>" '�ַ����ָ��
Private mcolCpt As Collection '���ϣ�����
Private mstrCpage As String '��ǰѡ������ξ���ı���
Private mlngҽ��ID As Long
Private mlng����ID As Long
Private mlng���� As Long '0-PDF�ļ�
Private mstrPrePati As String
Private mfrmPDF As Object '������PDF�ؼ�����Ĵ���
Private mblnPDF As Boolean
Private mclsPDF As Object
Private mbln��ֹ��ӡ As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng����id As Long, ByVal lng����ID As Long, ByVal lngModel As Long, ByVal strPrivs As String) As Boolean
'���ܣ������ӿ�
'������
'      frmParent     �������
'      lng����ID     ����id
'      lng����ID     ���ﲡ�˶�Ӧ�Һ�id,סԺ����Ϊ��ҳid
'      lng����ID     �������ڿ��ң��Һſ���/���˵�ǰ����
'      lng����ID     �������ڲ��������ﲡ�˴�0
'      lngModel      ģ��ţ�1252-p����ҽ���´1253-pסԺҽ���´�
'      strPrivs      Ȩ��

    mblnMoved = False
    mlng����ID = lng����ID
    mlng����ID = lng����ID
    mlngModel = lngModel
    mlng����ID = lng����id
    mlng����ID = lng����ID
    mstrMainPrivs = strPrivs
    mlngPreDept = -1
    Me.Show , frmParent
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrpara As Variant
    
    arrpara = Split(mobjPopup.Parameter, ",")
    mobjPopup.Visible = mlngPages > 0
    
    Select Case Control.ID
    Case conMenu_View_Forward, conMenu_View_Backward   '��һҳ,��һҳ
        Control.Visible = mlngPages > 1
        If Control.Visible Then
            If Val(arrpara(0)) = 1 And Control.ID = conMenu_View_Forward Then
                Control.Enabled = False
            ElseIf Val(arrpara(0)) = mlngPages And Control.ID = conMenu_View_Backward Then
                Control.Enabled = False
            Else
                Control.Enabled = True
            End If
        End If
    Case conMenu_File_Print
        Control.Enabled = Not mbln��ֹ��ӡ
    End Select
End Sub

Private Sub Form_Load()
    Dim intIdx As Integer
    Dim strErr As String
    Dim objTab As Object
    Dim strFile As String
    
    '��ʼ���˵�
    Call InitBar
    strFile = mstrCurFile
    Call CreatePDFobj
    Call InitObjLis(mlngModel)
    If Not gobjLIS Is Nothing Then
        If mlngModel = p����ҽ���´� Then
            Call gobjLIS.PatientSampleBrowse(Me, mlng����ID, mstrMainPrivs, mlng����ID, 0, 1, 0, strErr, False, mfrmLisCom)
        Else
            Call gobjLIS.PatientSampleBrowse(Me, mlng����ID, mstrMainPrivs, mlng����ID, mlng����ID, 2, mlng����ID, strErr, False, mfrmLisCom)
        End If
        If strErr <> "" Then
            MsgBox strErr, vbInformation, Me.Caption
        End If
    Else
        Call frmLisView.ShowMe(mlng����ID, mlngModel, Me, False, mfrmLisCom)
    End If
         
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem(intIdx, "��ͨ������", mfrmLisCom.hwnd, 0).Tag = "��ͨ": intIdx = intIdx + 1
        .InsertItem(intIdx, "΢����/ϸ��ѧ����", picMain.hwnd, 0).Tag = "΢����": intIdx = intIdx + 1
        .Item(0).Selected = True '�½�ʱ���Զ�ѡ�������,�����ټ����¼�
    End With
    
    '��Ϣչʾ����
    With tbcArchive
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .Color = xtpTabColorOffice2003
            .Layout = xtpTabLayoutAutoSize
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        Set objTab = .InsertItem(intIdx, "����", picWB.hwnd, 0): objTab.Tag = objTab.Caption
        If mblnPDF Then
            intIdx = intIdx + 1
            Set objTab = .InsertItem(intIdx, "����", mfrmPDF.hwnd, 0): objTab.Tag = objTab.Caption
                objTab.Visible = False
        Else
            .Item(0).Selected = True
        End If
    End With
    
    '������ʷ�б�
    '-----------------------------------------------------
    With tbcHistory
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .Color = xtpTabColorOffice2003
            .DisableLunaColors = False
            .BoldSelected = True
            .HotTracking = True
            .ShowIcons = True
        End With
        .SetImageList ils16
    End With
    Call InitReportColumn
    '���ξ�����Ϣ
    Call InitBasicData
    mstrCurFile = strFile
    Call DeleteLISTempFile(0)
    mstrCurFile = ""
    Call RestoreWinState(Me, App.ProductName)
    stbThis.Visible = True
End Sub

Private Sub InitBar()
'���ܣ��˵���ʼ��
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl
    
    '������----------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
     
    '����������:������������
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False

    With objBar.Controls
    
        Set mobjPopup = .Add(xtpControlPopup, conMenu_Edit_NewItem, "ҳ��")
        mobjPopup.IconId = conMenu_Edit_Modify
        mobjPopup.Style = xtpButtonIconAndCaption
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Forward, "��һҳ", -1, False)
        objControl.IconId = conMenu_View_Forward
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Backward, "��һҳ", -1, False)
        objControl.IconId = conMenu_View_Backward
         
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ������")
            objControl.IconId = conMenu_File_Parameter
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
'        Set objControl = .Add(xtpControlButton, conMenu_Tool_Reference, "�ο�")
'            objControl.BeginGroup = True
    End With
    objBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    objBar.ContextMenuPresent = False
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlCustom And objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngPage As Long
    Dim strCaption As String
    Dim arrpara As Variant
    Dim i As Long
    
    Select Case Control.ID
        Case conMenu_View_Jump
            arrpara = Split(Control.Parameter, M_S)
            lngPage = Val(arrpara(0))
            strCaption = arrpara(1)
        Case conMenu_View_Forward, conMenu_View_Backward   '��һҳ,��һҳ
            arrpara = Split(mobjPopup.Parameter, M_S)
            strCaption = arrpara(1)
            If Control.ID = conMenu_View_Forward Then
                lngPage = Val(arrpara(0)) - 1
            ElseIf Control.ID = conMenu_View_Backward Then
                lngPage = Val(arrpara(0)) + 1
            End If
        Case conMenu_File_Print '��ӡ
            Call FuncPrint
        Case conMenu_Tool_Reference
            Call FunReference
        Case conMenu_File_PrintSet
            cdgPrint.ShowPrinter
    End Select
    
    Select Case Control.ID
    Case conMenu_View_Jump, conMenu_View_Forward, conMenu_View_Backward
        mstrCpage = mcolCpt("" & lngPage)
        mobjPopup.Caption = mstrCpage
        mobjPopup.Parameter = lngPage & M_S & strCaption
        mobjPopup.SetFocus
        For i = 1 To mobjPopup.CommandBar.Controls.Count
            If lngPage = i Then
                mobjPopup.CommandBar.Controls(i).Checked = True
            Else
                mobjPopup.CommandBar.Controls(i).Checked = False
            End If
        Next
        Call LoadFaceData(mlngPages - lngPage + 1)
        cbsMain.RecalcLayout
    End Select
End Sub

Private Sub InitBasicData()
'���ܣ���ʼ��һЩ�������ݣ��������б���ص�
    Dim strSQL As String
    Dim objTab As TabControlItem
    Dim strTmp As String
    Dim str����IDs As String
    Dim rsTmp As ADODB.Recordset
    Dim str���֤�� As String
    Dim strTemp As String
    Dim n As Long, p As Long
    Dim strThis As String
    Dim strSQLPati As String
    Dim varPar(0 To 10) As String
    Dim objControl As CommandBarControl
    Dim objCloc As CommandBarControl '���ڶ�λ����һ�ξ���
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If Not mobjPopup Is Nothing Then mobjPopup.CommandBar.Controls.DeleteAll
    
    mlngPreIndex = -1
    strSQL = "select a.���֤�� from ������Ϣ a where a.����id=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    strTmp = rsTmp!���֤�� & ""
    If strTmp <> "" Then
        '��֤���֤�ŵĺϷ���
        If InitObjPublicPatient Then
            If gobjPublicPatient.CheckPatiIdcard(strTmp) Then
                str���֤�� = strTmp
            End If
        End If
    End If
         
    'ͨ�����֤�Ų����
    If str���֤�� <> "" Then
        strSQL = "select a.����id from ������Ϣ a where a.����id<>[1] and a.���֤��=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, str���֤��)
        Do While Not rsTmp.EOF
            str����IDs = str����IDs & "," & rsTmp!����ID
            rsTmp.MoveNext
        Loop
        str����IDs = Mid(str����IDs, 2)
    End If
    
    'ͨ������������
'    strTmp = GetPatiRelate(mlng����ID, str���֤��)
'    If strTmp <> "" Then
'        If str����IDs <> "" Then
'            str����IDs = str����IDs & "," & strTmp
'        Else
'            str����IDs = strTmp
'        End If
'    End If
    
    If str����IDs = "" Then
        strSQL = " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,0 as ����ת��,-1 as ��������,null as ����� From ���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
            " Union ALL" & _
            " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,1 as ����ת��,-1 as ��������,null as ����� From H���˹Һż�¼ Where ����ID=[1] And ��¼����=1 And ��¼״̬=1" & _
            " Union ALL" & _
            " Select ����id,��ҳID as ����ID,Null,��Ժ���� as ��ʼʱ��,��Ժ����,��Ժ����ID,����ת��,NVL(��������,0) as ��������,null as ����� From ������ҳ Where ����ID=[1] And Nvl(��ҳID,0)<>0"
        strSQL = "Select Rownum As ���,a.����ID,A.����ID,A.NO,A.��ʼʱ��,A.����ʱ��,B.���� as ����,A.����ת�� ,A.��������,a.����� From (" & strSQL & ") A,���ű� B Where A.����ID=B.ID Order by ��ʼʱ�� Desc"
        Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    Else
        str����IDs = mlng����ID & "," & str����IDs
        
        '����4000���ȵĲ��
        strTemp = "Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) X"
        n = 0
        Do While True
            If Len(str����IDs) < 4000 Then
                p = Len(str����IDs) + 1
            Else
                p = InStrRev(Mid(str����IDs, 1, 4000), ",")
            End If
            strThis = Mid(str����IDs, 1, p - 1)
            
            If n > 10 Then
                strSQLPati = strSQLPati & vbNewLine & " Union All " & Replace(strTemp, "[1]", "'" & strThis & "'")
            Else
                varPar(n) = strThis
                strSQLPati = IIF(strSQLPati = "", "", strSQLPati & vbNewLine & " Union All ") & Replace(strTemp, "[1]", "[" & (n + 1) & "]")
            End If
            
            n = n + 1
            str����IDs = Mid(str����IDs, p + 1)
            If str����IDs = "" Then Exit Do
        Loop
        strTmp = " ����ID In (" & strSQLPati & ")"
        
        
        strSQL = " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,0 as ����ת��,-1 as ��������,null as ����� From ���˹Һż�¼ Where " & strTmp & " And ��¼����=1 And ��¼״̬=1 and NO is not null" & _
            " Union ALL" & _
            " Select ����id,ID as ����ID,NO,����ʱ�� as ��ʼʱ��,Null as ����ʱ��,ִ�в���ID as ����ID,1 as ����ת��,-1 as ��������,null as ����� From H���˹Һż�¼ Where " & strTmp & " And ��¼����=1 And ��¼״̬=1 and NO is not null" & _
            " Union ALL" & _
            " Select ����id,��ҳID as ����ID,Null,��Ժ���� as ��ʼʱ��,��Ժ����,��Ժ����ID,����ת��,NVL(��������,0) as ��������,סԺ�� as ����� From ������ҳ Where " & strTmp & " And Nvl(��ҳID,0)<>0"
        strSQL = "Select Rownum As ���,a.����ID,A.����ID,A.NO,A.��ʼʱ��,A.����ʱ��,B.���� as ����,A.����ת�� ,A.��������,a.����� From (" & strSQL & ") A,���ű� B Where A.����ID=B.ID  Order by ��ʼʱ�� Desc"
        Set mrsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, varPar(0), varPar(1), varPar(2), varPar(3), varPar(4), varPar(5), varPar(6), varPar(7), varPar(8), varPar(9), varPar(10))
    End If
    
    Set mcolCpt = New Collection
    n = 0
    mlngPages = mrsData.RecordCount
    Do While Not mrsData.EOF
        
        strTmp = IIF(IsNull(mrsData!NO), "��" & mrsData!����ID & "��" & IIF(mrsData!�������� = 1, "��������", IIF(mrsData!�������� = 2, "סԺ����", "סԺ")), "�������") & ":" & mrsData!���� & "," & Format(mrsData!��ʼʱ��, "yyyy-MM-dd HH:mm") & _
            IIF(Not IsNull(mrsData!����ʱ��), "��" & Format(mrsData!����ʱ��, "yyyy-MM-dd HH:mm"), "")
            
        If mrsData.AbsolutePosition = 1 Then
            Set objTab = tbcHistory.InsertItem(tbcHistory.ItemCount, strTmp, picList.hwnd, IIF(IsNull(mrsData!NO), 0, 1))
        End If
        n = n + 1
        Set objControl = mobjPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strTmp, -1, False)
        objControl.Parameter = n & M_S & Val(mrsData!���) & M_S & strTmp
        mcolCpt.Add strTmp, n & ""
        If mlng����ID = Val(mrsData!����ID & "") And mlng����ID = Val(mrsData!����ID & "") Then
            Set objCloc = objControl
        End If
        mrsData.MoveNext
    Loop
    
    If objCloc Is Nothing Then
        objControl.Execute
    Else
        objCloc.Execute
    End If
        
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    tbcSub.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    mlngҽ��ID = 0
    Set mrsData = Nothing
    Unload mfrmLisCom
    Set mfrmLisCom = Nothing
    Call DeleteLISTempFile(1)
    If Not mfrmPDF Is Nothing Then
        Unload mfrmPDF
        Set mfrmPDF = Nothing
    End If
    Set mclsPDF = Nothing
    mblnPDF = False
    mbln��ֹ��ӡ = False
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    
    
    '--���沼�֣��˵�����࣬�ָ��ߣ����ϣ��м�
    
    picBar.Move 0, 0, picMain.ScaleWidth, 500
    stbThis.Width = picMain.ScaleWidth
    stbThis.Left = 0
    stbThis.Top = picMain.ScaleHeight - stbThis.Height
    
    fraPList.Top = picBar.Height
    fraPList.Left = 0
    fraPList.Height = picMain.Height - picBar.Height - stbThis.Height
 
    fraLR.Top = fraPList.Top
    fraLR.Height = fraPList.Height
    fraLR.Left = fraPList.Width
    
    picInfo.Top = fraPList.Top
    picInfo.Left = fraLR.Left + 45
    picInfo.Width = picMain.ScaleWidth - picInfo.Left
    picInfo.Height = 840
    
    fraRpt.Top = picInfo.Top + picInfo.Height
    fraRpt.Left = picInfo.Left
    fraRpt.Width = picInfo.Width
    fraRpt.Height = fraLR.Height - picInfo.Height
    
    
    '--��ಡ���б�����б������ڲ�
    tbcHistory.Move 0, 140, fraPList.Width - 50, fraPList.Height - 150
    
    '�м�����
    tbcArchive.Move 0, 0, fraRpt.Width, fraRpt.Height
    
    '--���˻�����Ϣ�ı���ɫ
    picInfo.BackColor = fraLR.BackColor
    fraInfo.BackColor = picInfo.BackColor
    fraIn.BackColor = picInfo.BackColor
    fraOut.BackColor = picInfo.BackColor
    
    '--picInfo�����ڲ��ؼ�λ������
    fraInfo.Left = 0
    fraInfo.Top = 0
    fraInfo.Width = picInfo.Width - fraInfo.Left * 3
    fraIn.Width = fraInfo.Width - fraIn.Left * 2
    fraOut.Width = fraIn.Width
    lbl��.Left = fraOut.Width - lbl��.Width - 60
   
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    If Button = 1 Then
        If fraLR.Left + X < 1000 Or fraLR.Left + X > picMain.Width - 3000 Then Exit Sub
        fraPList.Width = fraPList.Width + X
        Call picMain_Resize
    End If
End Sub
 
Private Function ShowOutPatiInfo() As Boolean
'���ܣ�ѡ�����ﲡ��ĳ����ʷ�����¼ʱ����ȡ��صĲ�����Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng����ID <> 0 Then
        strSQL = "Select B.Id,B.NO,B.�����,B.����,B.�Ա�,B.����,A.ҽ�Ƹ��ʽ," & _
            " A.�ѱ�,A.����,A.ҽ����,B.����,B.����ʱ��,B.ִ����,B.ִ��״̬,B.ִ��ʱ��," & _
            " B.ִ�в���ID as ����ID,B.����,B.����,D.������,C.���� as ����" & _
            " From ������Ϣ A,���˹Һż�¼ B,���ű� C,����������Ϣ D" & _
            " Where A.����ID=B.����ID And B.ID=[1] And B.ִ�в���ID=C.ID" & _
            " And B.����ID=D.����ID(+) And B.����=D.����(+) And B.��¼����=1 And B.��¼״̬=1"
        If mblnMoved Then
            strSQL = Replace(strSQL, "���˹Һż�¼", "H���˹Һż�¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        With rsTmp
            '���ղ���������ɫ��ʾ
            lbl����mz(1).Caption = NVL(!����)
            If Not IsNull(!����) Then
                lbl����mz(1).ForeColor = vbRed
            Else
                lbl����mz(1).ForeColor = lbl�����mz(1).ForeColor
            End If
            lblҽ��mz(1).Caption = NVL(!ִ����)
            lbl�Һŵ�mz(1).Caption = !NO
            lbl�����mz(1).Caption = NVL(!�����)
            lbl����mz(1).Caption = NVL(!ҽ�Ƹ��ʽ)
            lbl�ѱ�mz(1).Caption = NVL(!�ѱ�)
            lblҽ����mz(1).Caption = NVL(!ҽ����)
            lbl������mz(1).Caption = NVL(!������)
            lbl��.Visible = NVL(!����, 0) <> 0
        End With
    Else
        fraOut.Visible = True
        lbl����mz(1).Caption = ""
        lblҽ��mz(1).Caption = ""
        lbl�Һŵ�mz(1).Caption = ""
        lbl�����mz(1).Caption = ""
        lbl����mz(1).Caption = ""
        lbl�ѱ�mz(1).Caption = ""
        lblҽ����mz(1).Caption = ""
        lbl������mz(1).Caption = ""
    End If
    ShowOutPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowInPatiInfo() As Boolean
'���ܣ�ѡ��ĳ��סԺ��¼ʱ����ȡ��صĲ�����Ϣ
'���أ�blnMoved=����סԺ�����Ƿ�ת����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    If mlng����ID <> 0 Then
        strSQL = "Select NVL(B.����,A.����) ����, NVL(B.�Ա�,A.�Ա�) �Ա�, NVL(B.����,A.����) ����,B.סԺ��,B.��Ժ����,B.ҽ�Ƹ��ʽ," & _
            " D.��Ϣֵ as ҽ����,B.����,B.��ǰ����,C.���� as ����ȼ�,B.��Ժ����," & _
            " B.��Ժ����,B.��������,B.״̬,B.��Ժ����ID,B.��ǰ����ID,A.סԺ����" & _
            " From ������Ϣ A,������ҳ B,�շ���ĿĿ¼ C,������ҳ�ӱ� D" & _
            " Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2] And B.����ȼ�ID=C.ID(+)" & _
            " And B.����ID=D.����ID(+) And B.��ҳID=D.��ҳID(+) And D.��Ϣ��(+)='ҽ����'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID)
        
        With rsTmp
            '���ղ�����ɫ������ʾ
            lbl����zy(1).Caption = NVL(!����)
            lbl����zy(1).ForeColor = zlDatabase.GetPatiColor(NVL(!��������))   'GetPatiColor(NVL(!��������))
            
            lblסԺ��zy(1).Caption = NVL(!סԺ��)
            lbl����zy(1).Caption = NVL(!��Ժ����)
            lblҽ����zy(1).Caption = NVL(!ҽ����)
            lbl����zy(1).Caption = NVL(!����ȼ�)
            lbl����zy(1).Caption = NVL(!ҽ�Ƹ��ʽ)
            
            'Σ�ز��˲�����ɫ��ʾ
            lbl����zy(1).Caption = NVL(!��ǰ����)
            If NVL(!��ǰ����) = "Σ" Or NVL(!��ǰ����) = "��" Or NVL(!��ǰ����) = "��" Then
                lbl����zy(1).ForeColor = vbRed
            Else
                lbl����zy(1).ForeColor = lblסԺ��zy(1).ForeColor
            End If
            
            lbl��Ժzy(1).Caption = Format(!��Ժ����, "yyyy-MM-dd HH:mm")
            If Not IsNull(!��Ժ����) Then
                lbl��Ժzy(1).Caption = lbl��Ժzy(1).Caption & "��" & Format(!��Ժ����, "yyyy-MM-dd HH:mm")
            End If
            lbl����zy(1).Caption = NVL(!��������)
        End With
    Else
        '���ղ�����ɫ������ʾ
        fraIn.Visible = True
        lbl����zy(1).Caption = ""
        lblסԺ��zy(1).Caption = ""
        lbl����zy(1).Caption = ""
        lblҽ����zy(1).Caption = ""
        lbl����zy(1).Caption = ""
        lbl����zy(1).Caption = ""
        lbl����zy(1).Caption = ""
        lbl��Ժzy(1).Caption = ""
        lbl����zy(1).Caption = ""
    End If
    ShowInPatiInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub picWB_Resize()
    On Error Resume Next
    webSub.Move 0, 0, picWB.Width, picWB.Height
End Sub
 
Private Sub WebShow(ByVal strKey As String, ByVal strCpt As String)
'���ܣ��ؼ�չʾ�ļ�Ҫ�����Ƿ������PDF�ؼ������
'����: strKey �̶���ʽ,����ID;ҽ��ID;����
'       strCpt ����
    Dim strUrl As String '�ļ�·��
    
    If strKey = "" Then
        Call webSub.Navigate("about:blank")
        webSub.Visible = False
        mstrCurFile = ""
        
        If mblnPDF Then
            tbcArchive(1).Visible = False
            tbcArchive(1).Selected = False
        End If
        
        tbcArchive(0).Visible = True
        tbcArchive(0).Selected = True
        tbcArchive(0).Caption = strCpt
    Else
        strUrl = GetLisRptFile(strKey)
        If mblnPDF And mlng���� = 0 Then
            Call mclsPDF.LoadFile(strUrl)
            tbcArchive(0).Visible = False
            tbcArchive(0).Selected = False
            tbcArchive(1).Visible = True
            tbcArchive(1).Selected = True
            tbcArchive(1).Caption = strCpt
	    mfrmPDF.Enabled = Not mbln��ֹ��ӡ
        Else
            If strUrl <> "" Then
                webSub.Navigate strUrl
                mstrCurFile = strUrl
            End If
            webSub.Visible = True
            
            If mblnPDF Then
                tbcArchive(1).Visible = False
                tbcArchive(1).Selected = False
            End If
            
            tbcArchive(0).Visible = True
            tbcArchive(0).Selected = True
            tbcArchive(0).Caption = strCpt
	    picWB.Enabled = Not mbln��ֹ��ӡ
        End If
    End If
End Sub
 
Private Function GetLisRptFile(ByVal strTag As String) As String
'���ܣ���LIS�����ļ��鿴����ȡ��ʱ�ļ�·��
'����: strTag �����,�̶���ʽ,����ID;ҽ��ID;����
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim lng����ID As Long
    Dim lngҽ��ID As Long
    Dim lng���� As String
    Dim varTmp As Variant
    Dim strSuffix As String '�ļ���׺��
    
    Screen.MousePointer = 11
    varTmp = Split(strTag, ";")
    lng����ID = Val(varTmp(0))
    lngҽ��ID = Val(varTmp(1))
    lng���� = varTmp(2)
    If lng���� = 0 Then
        strSuffix = "pdf"
    ElseIf lng���� = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng����ID & "." & strSuffix
    If InStr(mstrFilesTemp, strFile) = 0 Then
        mstrFilesTemp = mstrFilesTemp & "<STAB>" & strFile
    End If
    If Not objFile.FileExists(strFile) Then
        strFile = Sys.ReadLob(glngSys, 22, lng����ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, gstrSysName:
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function

Private Sub FuncPrint()
'���ܣ���ӡ��ǰPDF�ļ������߶���ļ�,�����Ƕ���ļ�������xps��pdf�ļ�
'      ����ļ�ʱ�ϲ����ٴ�ӡ
    Dim i As Long
    Dim lng����ID As Long
    Dim lngҽ��ID As Long
    Dim rsPati As ADODB.Recordset
    Dim lngPage As Long
    
    lngPage = mlngPreDept
    Set rsPati = New ADODB.Recordset
    rsPati.Fields.Append "ҽ��ID", adBigInt
    rsPati.Fields.Append "����ID", adBigInt
    rsPati.Fields.Append "����", adBigInt
    rsPati.Fields.Append "��Ŀ", adVarChar, 4000
    rsPati.CursorLocation = adUseClient
    rsPati.LockType = adLockOptimistic
    rsPati.CursorType = adOpenStatic
    rsPati.Open
    
    For i = 0 To rptPati.Rows.Count - 1
        If Not rptPati.Rows(i).GroupRow Then
            If rptPati.Rows(i).Record(COL_ѡ��).Checked Then
                With rptPati.Rows(i).Record
                    If Val(.Item(COL_��ֹ��ӡ).value) = 0 Then
                        rsPati.AddNew
                        rsPati!ҽ��ID = Val(.Item(COL_ҽ��ID).value)
                        rsPati!����ID = Val(.Item(COL_����ID).value)
                        rsPati!���� = Val(.Item(COL_����).value)
                        rsPati!��Ŀ = .Item(COL_��Ŀ).value
                        rsPati.Update
                    End If
                End With
            End If
        End If
    Next
    
    '��û�й�ѡ��ʱ��Ŵ�ӡ��ǰѡ�е���һ��
    If rptPati.SelectedRows.Count > 0 And rsPati.RecordCount = 0 Then
        With rptPati.SelectedRows(0)
            lngҽ��ID = Val(.Record(COL_ҽ��ID).value)
            lng����ID = Val(.Record(COL_����ID).value)
            rsPati.Filter = "ҽ��ID=" & lngҽ��ID & " and ����ID=" & lng����ID
            If rsPati.EOF Then
                If Val(.Item(COL_��ֹ��ӡ).value) = 0 Then
                    rsPati.AddNew
                    rsPati!ҽ��ID = lngҽ��ID
                    rsPati!����ID = lng����ID
                    rsPati!���� = Val(.Record(COL_����).value)
                    rsPati!��Ŀ = .Record(COL_��Ŀ).value
                    rsPati.Update
                End If
            End If
        End With
    End If
    
    rsPati.Filter = 0
    If rsPati.RecordCount = 0 Then
        Exit Sub
    End If
    
    '��PDF�ؼ�ʱ
    If Not mblnPDF Then
        Call PrintNoPdf(rsPati)
    Else
        '����PDF�ؼ�ʱ,����xps��PDF
        rsPati.Filter = "����=2"
        If Not rsPati.EOF Then
            Call PrintNoPdf(rsPati)
        End If
        rsPati.Filter = "����=0"
        If Not rsPati.EOF Then
            Call PrintPdf(rsPati)
        End If
    End If
    Call LoadFaceData(lngPage)
End Sub

Private Sub PrintPdf(ByRef rsPati As ADODB.Recordset)
'����:��ӡ��Ӧ����PDF�ؼ����ļ�
    Dim i As Long
    Dim strTag As String
    Dim lng����ID As Long
    Dim strFileSource As String
    Dim strSQL As String
    Dim lngCnt As Long
    
   On Error GoTo errH

    lngCnt = rsPati.RecordCount
    
    For i = 1 To rsPati.RecordCount
        lng����ID = rsPati!����ID
        strTag = rsPati!����ID & ";" & rsPati!ҽ��ID & ";" & rsPati!����
        strFileSource = GetLisRptFile(strTag)

        Call mclsPDF.LoadFile(strFileSource)
        Call mclsPDF.PrintFile(0)
        
        If lngCnt <> 1 Then
            '�������ʱ�ŵ����ȴ�
            Call mclsPDF.WaitTime(0, strFileSource, rsPati!��Ŀ & "")
        End If
        
        strSQL = "Zl_ҽ����������_Print(" & lng����ID & ",0)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        
        rsPati.MoveNext
    Next
    mlngPreDept = -1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PrintNoPdf(ByRef rsPati As ADODB.Recordset)
'����:��ӡ����Ӧ����PDF�ؼ����ļ�
    Dim i As Long
    Dim strTag As String
    Dim lng����ID As Long
    Dim strFileSource As String
    
    For i = 1 To rsPati.RecordCount
        strTag = rsPati!����ID & ";" & rsPati!ҽ��ID & ";" & rsPati!����
        strFileSource = GetLisRptFile(strTag)
        lng����ID = Split(strTag, ";")(0)
        Call FunFastPrint(strFileSource, lng����ID)
        rsPati.MoveNext
    Next
    mlngPreDept = -1
End Sub

Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'���ܣ�API���ÿ��ٴ�ӡPDF�ļ�
'������strFile �ļ�·��
    Dim RetVal As Long
    Dim strSQL As String
    Dim ShExInfo As SHELLEXECUTEINFO
    
    On Error GoTo errH
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fmask = &H40
        .hwnd = 0
        .lpVerb = "print"
        .lpFile = strFile
        .lpParameters = ""
        .lpDirectory = vbNullChar
        .nShow = 0
    End With
    RetVal = ShellExecuteEx(ShExInfo)
    If RetVal = 0 Then
        Exit Sub
    End If
    strSQL = "Zl_ҽ����������_Print(" & lngRptID & ",0)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function DeleteLISTempFile(ByVal intType As Integer) As Boolean
'���ܣ�ɾ����ʱ�ļ�
'������intType 0-�򿪴���ʱ���ã�1-�رմ���ʱ����
    Dim objFile As New FileSystemObject
    Dim i As Long
    Dim varFiles As Variant
    Dim strTmp As String
    
    On Error GoTo errH
    If intType = 1 Then
        If mstrFilesTemp = "" Then Exit Function
        Call webSub.Navigate("about:blank")
        varFiles = Split(mstrFilesTemp, "<STAB>")
        For i = 0 To UBound(varFiles)
            strTmp = varFiles(i)
            If strTmp <> "" And strTmp <> mstrCurFile Then
                If objFile.FileExists(strTmp) Then
                    objFile.DeleteFile strTmp, True
                End If
            End If
        Next
        mstrFilesTemp = ""
        strTmp = mstrCurFile
        If strTmp <> "" Then
            On Error Resume Next
            If objFile.FileExists(strTmp) Then
                objFile.DeleteFile strTmp, True
            End If
        End If
    Else
        strTmp = mstrCurFile
        If strTmp <> "" Then
            On Error Resume Next
            If objFile.FileExists(strTmp) Then
                objFile.DeleteFile strTmp, True
            End If
        End If
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiRelate(ByVal lng����ID As Long, ByVal strIDNo As String) As String
'���ܣ���ȡָ������֮�ص����Ĳ���ID�������в�������ǰ����Ĳ���ID
'������lng����ID������id;strIDNo ���֤��
'˵��������������֤�Ų�Ϊ���ַ���ʱ�򷵻����в��ˣ�����Ϊ�����ų����֤ΪstrIDNo�Ĳ��ˣ�Ҫ�󴫵����֤�ű����ǺϷ������֤�ż�ͨ��ZLHIS���֤����֤
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If strIDNo = "" Then
        strSQL = "select b.����id from ������ݹ��� a,������ݹ��� b where a.����id=b.����id and a.����id=[1] and b.����id+0<>[1]"
    Else
        strSQL = "select b.����id from ������ݹ��� a,������ݹ��� b,������Ϣ c where a.����id=b.����id and b.����id=c.����id and a.����id=[1] and b.����id+0<>[1] and nvl(c.���֤��,'-')<>[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetPatiRelate", lng����ID, strIDNo)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        strSQL = strSQL & "," & rsTmp!����ID
        rsTmp.MoveNext
    Loop
    
    GetPatiRelate = Mid(strSQL, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadFaceData(ByVal lngPage As Long)

    If mlngPreDept = lngPage Then Exit Sub
    
    mlngPreDept = lngPage
    
    mrsData.Filter = "���=" & mlngPreDept
    
    mlng����ID = mrsData!����ID
    mlng����ID = mrsData!����ID

    If Not mrsData.EOF Then
        mstr�Һŵ� = NVL(mrsData!NO, "")
        mblnMoved = Val(NVL(mrsData!����ת��, "")) = 1
    End If
    '��ʾ������Ϣ
    If mstr�Һŵ� <> "" Then
        Call ShowOutPatiInfo
    Else
        Call ShowInPatiInfo
    End If
    
    fraOut.Visible = mstr�Һŵ� <> ""
    fraIn.Visible = mstr�Һŵ� = ""

    '��ʾ����Ŀ¼
    Me.tbcHistory(0).Caption = mstrCpage
    Call LoadPatients
    Call Form_Resize
End Sub

Private Sub picList_Resize()
    On Error Resume Next
    rptPati.Move 0, 0, picList.Width, picList.Height
End Sub

Private Sub InitReportColumn()
'����:��ʼ�����
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(COL_ѡ��, "", 18, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = 5
         
        Set objCol = .Columns.Add(COL_ͼ��, "", 18, False)
            objCol.Sortable = False
            objCol.Alignment = xtpAlignmentCenter
            objCol.AllowDrag = False
        Set objCol = .Columns.Add(COL_��Ŀ, "��Ŀ", 120, True)
        Set objCol = .Columns.Add(COL_���ʱ��, "���ʱ��", 106, True)
        Set objCol = .Columns.Add(COL_�걾, "�걾", 60, True)
        Set objCol = .Columns.Add(COL_����ʱ��, "����ʱ��", 106, True)
        Set objCol = .Columns.Add(COL_������, "������", 55, True)
        Set objCol = .Columns.Add(COL_�ɼ�ʱ��, "�ɼ�ʱ��", 106, True)
        
        '������
        Set objCol = .Columns.Add(COL_ҽ��ID, "ҽ��ID", 0, False)
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(COL_����, "����", 0, False)
        Set objCol = .Columns.Add(COL_��ӡ����, "��ӡ����", 0, False)
        Set objCol = .Columns.Add(COL_�ĵ�����, "�ĵ�����", 0, False)
        Set objCol = .Columns.Add(COL_��ֹ��ӡ, "��ֹ��ӡ", 0, False)
        
        For Each objCol In .Columns
            If objCol.Index <> COL_ѡ�� Then objCol.Editable = False
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .TreeIndent = 0 '�з�����ʱ�������߱��ϻ�����һ������
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList Me.img16
    End With
End Sub

Private Sub LoadPatients()
'���ܣ����ز����б�
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
     
    '����LIS����
    If mstr�Һŵ� = "" Then
        strSQL = "select b.id as ����ID,b.������ as ��Ŀ, To_Char(b.����ʱ��,'YYYY-MM-DD HH24:MI') as ���ʱ��,b.������||','||To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as �ĵ�����,c.ҽ��ID,b.����,b.��ӡ����,b.�Ƿ��ֹ��ӡ," & _
            " To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,a.�걾��λ as �걾,a.����ҽ�� as ������,To_Char(d.����ʱ��,'YYYY-MM-DD HH24:MI') as �ɼ�ʱ��" & _
            " from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c,����ҽ������ d" & _
            " where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and a.id=d.ҽ��id(+) and b.���� in (0,2) and a.����id=[1] and a.��ҳid=[2]" & _
            " order by b.����ʱ�� desc"
    Else
        strSQL = "select b.id as ����ID,b.������ as ��Ŀ,To_Char(b.����ʱ��,'YYYY-MM-DD HH24:MI') as ���ʱ��,b.������||','||To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as �ĵ�����,c.ҽ��ID,b.����,b.��ӡ����,b.�Ƿ��ֹ��ӡ," & _
            " To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,a.�걾��λ as �걾,a.����ҽ�� as ������,To_Char(d.����ʱ��,'YYYY-MM-DD HH24:MI') as �ɼ�ʱ��" & _
            " from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c,����ҽ������ d" & _
            " where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and a.id=d.ҽ��id(+) and b.���� in (0,2) and a.�Һŵ�=[3]" & _
            " order by b.����ʱ�� desc"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "ҽ����������", "Hҽ����������")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
 
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng����ID, mstr�Һŵ�)
    mstrPrePati = ""
    Call WebShow("", "����")
    rptPati.Records.DeleteAll
    Do Until rsTmp.EOF
        Set objRecord = Me.rptPati.Records.Add()
            objRecord.Tag = CStr("_" & rsTmp!ҽ��ID & "_" & rsTmp!����ID) '���ڶ�λΨһ��
        'ѡ����
        Set objItem = objRecord.AddItem("")
        objItem.HasCheckbox = True
        
        'ͼ��
        Set objItem = objRecord.AddItem("")
        If Val(rsTmp!��ӡ���� & "") > 0 Then
            objItem.Icon = 8
        End If
        
        '��Ŀ
        objRecord.AddItem rsTmp!��Ŀ & ""
        
        '���ʱ��
        objRecord.AddItem rsTmp!���ʱ�� & ""
        
        '�걾
        objRecord.AddItem rsTmp!�걾 & ""
        
        '����ʱ��
        objRecord.AddItem rsTmp!����ʱ�� & ""
        
        '������
        objRecord.AddItem rsTmp!������ & ""
        
        '�ɼ�ʱ��
        objRecord.AddItem rsTmp!�ɼ�ʱ�� & ""
        
        'ҽ��ID
        objRecord.AddItem rsTmp!ҽ��ID & ""
        
        '����ID
        objRecord.AddItem rsTmp!����ID & ""
        
        '����
        objRecord.AddItem rsTmp!���� & ""
        
        '��ӡ����
        objRecord.AddItem rsTmp!��ӡ���� & ""
        
        '�ĵ�����
        objRecord.AddItem rsTmp!�ĵ����� & ""
        
        '��ֹ��ӡ
        objRecord.AddItem rsTmp!�Ƿ��ֹ��ӡ & ""
        
        rsTmp.MoveNext
    Loop
    rptPati.Populate
   
    Screen.MousePointer = 0
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
  
    Dim objColumn As ReportColumn
        
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = COL_ѡ�� Then
                    If objColumn.Caption = "" Then
                        objColumn.Caption = "1"
                        Call SelectALLPati(True)
                    Else
                        objColumn.Caption = ""
                        Call SelectALLPati(False)
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub SelectALLPati(ByVal blnSelect As Boolean)
'����:����ȫѡ����ȫ��
    Dim i As Long
    
    If rptPati.Columns(COL_ѡ��).Visible And rptPati.SelectedRows.Count > 0 Then
        '��������м�¼��ѡ��״̬
        For i = 0 To rptPati.Records.Count - 1
            rptPati.Records(i)(COL_ѡ��).Checked = False
        Next
        
        For i = 0 To rptPati.Rows.Count - 1
            rptPati.Rows(i).Record(COL_ѡ��).Checked = blnSelect
        Next
        rptPati.Redraw
    End If
End Sub

Private Sub rptPati_SelectionChanged()
'����:�л���
    Dim strCurPati As String
    Dim strKey As String
    If rptPati.SelectedRows.Count <= 0 Then Exit Sub
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then strCurPati = .Record.Tag
        If strCurPati = mstrPrePati Then Exit Sub
        mstrPrePati = strCurPati
        If Not .GroupRow Then
            mlngҽ��ID = .Record(COL_ҽ��ID).value
            mlng����ID = .Record(COL_����ID).value
            mlng���� = Val(.Record(COL_����).value)
            mbln��ֹ��ӡ = Val(.Record(COL_��ֹ��ӡ).value) = 1
        End If
        strKey = mlng����ID & ";" & mlngҽ��ID & ";" & mlng����
        Call zlControl.FormLock(Me.hwnd)
        Call WebShow(strKey, .Record(COL_�ĵ�����).value)
        Call zlControl.FormLock(0)
    End With
End Sub

Private Sub CreatePDFobj()
'����:��������
    On Error Resume Next
    mblnPDF = False
    Set mclsPDF = CreateObject("zlPDFViewer.clsPDFViewer")
    If Not mclsPDF Is Nothing Then
        Set mfrmPDF = mclsPDF.GetFrm
        If Not mfrmPDF Is Nothing Then
            mblnPDF = True
        End If
    End If
End Sub

Private Sub FunReference()
'���ܣ��������Ʋο�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strIDs As String
    Dim i As Long
    
    On Error GoTo errH
    
    If 0 <> mlngҽ��ID Then
        strSQL = "select a.������ĿID from ����ҽ����¼ a where a.id=[1] or a.���id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngҽ��ID)
        
        For i = 1 To rsTmp.RecordCount
            strIDs = strIDs & "," & rsTmp!������ĿID
            rsTmp.MoveNext
        Next
    End If
    
'    Call frmClinicHelp.ShowMe(0, Me, 0, True, Mid(strIDs, 2))
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
