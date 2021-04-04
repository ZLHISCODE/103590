VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPageSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "页面设置"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   Icon            =   "frmPageSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picButton 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   5385
      TabIndex        =   21
      Top             =   5025
      Width           =   5385
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4170
         TabIndex        =   23
         Top             =   45
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   2850
         TabIndex        =   22
         Top             =   45
         Width           =   1100
      End
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   2250
      ScaleHeight     =   4680
      ScaleWidth      =   5205
      TabIndex        =   24
      Top             =   825
      Width           =   5205
      Begin VB.CheckBox chkVAlignCenter 
         Caption         =   "垂直"
         Height          =   270
         Index           =   1
         Left            =   1095
         TabIndex        =   48
         Top             =   4230
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CheckBox chkHAlignCenter 
         Caption         =   "水平"
         Height          =   270
         Index           =   0
         Left            =   1095
         TabIndex        =   47
         Top             =   3810
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   3
         Left            =   1080
         TabIndex        =   45
         Top             =   3645
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   0
         Left            =   1080
         TabIndex        =   29
         Top             =   135
         Width           =   3810
      End
      Begin VB.ComboBox cboKind 
         Height          =   300
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   345
         Width           =   3750
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   3435
         MaxLength       =   6
         TabIndex        =   6
         Text            =   "297.08"
         Top             =   810
         Width           =   735
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   1095
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "210.05"
         Top             =   810
         Width           =   735
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   1
         Left            =   1080
         TabIndex        =   28
         Top             =   1335
         Width           =   3810
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   3
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   14
         Text            =   "31.7"
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   2
         Left            =   1095
         MaxLength       =   6
         TabIndex        =   12
         Text            =   "31.7"
         Top             =   1935
         Width           =   735
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   1
         Left            =   3450
         MaxLength       =   6
         TabIndex        =   10
         Text            =   "25.4"
         Top             =   1530
         Width           =   735
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   0
         Left            =   1095
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "25.4"
         Top             =   1530
         Width           =   735
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   2
         Left            =   1080
         TabIndex        =   27
         Top             =   2460
         Width           =   1590
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "纵向(&P)"
         Height          =   270
         Index           =   0
         Left            =   1095
         TabIndex        =   15
         Top             =   2715
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "横向(&S)"
         Height          =   270
         Index           =   1
         Left            =   1095
         TabIndex        =   16
         Top             =   3165
         Width           =   1065
      End
      Begin VB.PictureBox picViewer 
         BackColor       =   &H00808080&
         Height          =   2160
         Left            =   2730
         ScaleHeight     =   2100
         ScaleWidth      =   2100
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2475
         Width           =   2160
         Begin VB.PictureBox picPaper 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1830
            Left            =   405
            ScaleHeight     =   1800
            ScaleWidth      =   1335
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   60
            Width           =   1365
            Begin VB.Line linMarjin 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   2
               X1              =   105
               X2              =   105
               Y1              =   0
               Y2              =   1530
            End
            Begin VB.Line linMarjin 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   0
               X1              =   0
               X2              =   1410
               Y1              =   105
               Y2              =   105
            End
            Begin VB.Line linMarjin 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   3
               X1              =   930
               X2              =   930
               Y1              =   0
               Y2              =   1530
            End
            Begin VB.Line linMarjin 
               BorderColor     =   &H00C0C0C0&
               BorderStyle     =   3  'Dot
               Index           =   1
               X1              =   0
               X2              =   1410
               Y1              =   1215
               Y2              =   1215
            End
         End
      End
      Begin MSComCtl2.UpDown udHeight 
         Height          =   300
         Left            =   4185
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   810
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196617
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4410
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWidth 
         Height          =   300
         Left            =   1830
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   810
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196618
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   0
         Left            =   1830
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1530
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(0)"
         BuddyDispid     =   196619
         BuddyIndex      =   0
         OrigLeft        =   1680
         OrigTop         =   1530
         OrigRight       =   1920
         OrigBottom      =   1830
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   1
         Left            =   4185
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1530
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(1)"
         BuddyDispid     =   196619
         BuddyIndex      =   1
         OrigLeft        =   4035
         OrigTop         =   1530
         OrigRight       =   4275
         OrigBottom      =   1830
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   2
         Left            =   1830
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1935
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(2)"
         BuddyDispid     =   196619
         BuddyIndex      =   2
         OrigLeft        =   1680
         OrigTop         =   1935
         OrigRight       =   1920
         OrigBottom      =   2235
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   3
         Left            =   4185
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1935
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(3)"
         BuddyDispid     =   196619
         BuddyIndex      =   3
         OrigLeft        =   4035
         OrigTop         =   1935
         OrigRight       =   4275
         OrigBottom      =   2235
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblAlignCenter 
         AutoSize        =   -1  'True
         Caption         =   "居中方式"
         Height          =   180
         Left            =   330
         TabIndex        =   46
         Top             =   3570
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         Caption         =   "纸张种类"
         Height          =   180
         Left            =   330
         TabIndex        =   44
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   0
         Left            =   2085
         TabIndex        =   43
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高度(&H)"
         Height          =   180
         Left            =   2700
         TabIndex        =   5
         Top             =   870
         Width           =   630
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽度(&W)"
         Height          =   180
         Left            =   390
         TabIndex        =   3
         Top             =   870
         Width           =   630
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   1
         Left            =   4470
         TabIndex        =   42
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         Caption         =   "页边距"
         Height          =   180
         Left            =   330
         TabIndex        =   41
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label lblMarjin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右(&R)"
         Height          =   180
         Index           =   3
         Left            =   2880
         TabIndex        =   13
         Top             =   1995
         Width           =   450
      End
      Begin VB.Label lblMarjin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左(&L)"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   11
         Top             =   1995
         Width           =   450
      End
      Begin VB.Label lblMarjin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下(&B)"
         Height          =   180
         Index           =   1
         Left            =   2880
         TabIndex        =   9
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label lblMarjin 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上(&T)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   7
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   2
         Left            =   2085
         TabIndex        =   40
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   3
         Left            =   4470
         TabIndex        =   39
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   4
         Left            =   2085
         TabIndex        =   38
         Top             =   1995
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   5
         Left            =   4470
         TabIndex        =   37
         Top             =   1995
         Width           =   360
      End
      Begin VB.Label lblOrient 
         AutoSize        =   -1  'True
         Caption         =   "纸张方向"
         Height          =   180
         Left            =   330
         TabIndex        =   36
         Top             =   2385
         Width           =   720
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   1
         Left            =   570
         Picture         =   "frmPageSetup.frx":000C
         Top             =   3045
         Width           =   480
      End
      Begin VB.Image imgOrient 
         Height          =   480
         Index           =   0
         Left            =   570
         Picture         =   "frmPageSetup.frx":08D6
         Top             =   2625
         Width           =   480
      End
      Begin VB.Label lblKind 
         AutoSize        =   -1  'True
         Caption         =   "尺寸(&K)"
         Height          =   180
         Left            =   390
         TabIndex        =   1
         Top             =   405
         Width           =   630
      End
   End
   Begin VB.PictureBox picHeadFoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4680
      Left            =   30
      ScaleHeight     =   4680
      ScaleWidth      =   5205
      TabIndex        =   49
      Top             =   210
      Width           =   5205
      Begin VB.CommandButton cmdHead 
         Caption         =   "加入页眉(&U)"
         Height          =   350
         Left            =   3735
         TabIndex        =   65
         Top             =   1740
         Width           =   1200
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "页眉字体(&M)"
         Height          =   350
         Index           =   0
         Left            =   2550
         TabIndex        =   64
         Top             =   1740
         Width           =   1200
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "清除图片(&E)"
         Height          =   350
         Index           =   1
         Left            =   1365
         TabIndex        =   62
         Top             =   1740
         Width           =   1200
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "选择图片(&P)"
         Height          =   350
         Index           =   0
         Left            =   180
         TabIndex        =   63
         Top             =   1740
         Width           =   1200
      End
      Begin VB.ComboBox cboMode 
         Height          =   300
         ItemData        =   "frmPageSetup.frx":11A0
         Left            =   1110
         List            =   "frmPageSetup.frx":11A2
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2130
         Width           =   3825
      End
      Begin VB.CommandButton cmdFoot 
         Caption         =   "加入页脚(&D)"
         Height          =   350
         Left            =   3735
         TabIndex        =   58
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton cmdFont 
         Caption         =   "页脚字体(J)"
         Height          =   350
         Index           =   1
         Left            =   2550
         TabIndex        =   57
         Top             =   2460
         Width           =   1200
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   5
         Left            =   720
         TabIndex        =   51
         Top             =   195
         Width           =   2025
      End
      Begin VB.TextBox txtHead 
         Height          =   300
         Left            =   3510
         MaxLength       =   6
         TabIndex        =   18
         Text            =   "15"
         Top             =   60
         Width           =   735
      End
      Begin VB.TextBox txtFoot 
         Height          =   300
         Left            =   3510
         MaxLength       =   6
         TabIndex        =   20
         Text            =   "15"
         Top             =   2925
         Width           =   735
      End
      Begin VB.Frame fraLine 
         Height          =   30
         Index           =   6
         Left            =   720
         TabIndex        =   50
         Top             =   3060
         Width           =   2025
      End
      Begin MSComCtl2.UpDown UPHead 
         Height          =   300
         Left            =   4275
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   60
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHead"
         BuddyDispid     =   196643
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4410
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UPFoot 
         Height          =   300
         Left            =   4275
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   2925
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtFoot"
         BuddyDispid     =   196644
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeadContent 
         Height          =   1215
         Left            =   1380
         MultiLine       =   -1  'True
         TabIndex        =   67
         Top             =   420
         Width           =   3540
      End
      Begin VB.PictureBox picHead 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   180
         ScaleHeight     =   1185
         ScaleWidth      =   1110
         TabIndex        =   66
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox txtFootContent 
         Height          =   1215
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   68
         Top             =   3360
         Width           =   4740
      End
      Begin VB.Label lblHead 
         AutoSize        =   -1  'True
         Caption         =   "页眉"
         Height          =   180
         Left            =   330
         TabIndex        =   61
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblMode 
         AutoSize        =   -1  'True
         Caption         =   "替换要素"
         Height          =   180
         Left            =   330
         TabIndex        =   60
         Top             =   2190
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   6
         Left            =   4545
         TabIndex        =   56
         Top             =   2985
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "边距(&B)"
         Height          =   180
         Left            =   2775
         TabIndex        =   17
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblFoot 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "边距(&A)"
         Height          =   180
         Index           =   1
         Left            =   2775
         TabIndex        =   19
         Top             =   2985
         Width           =   630
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   7
         Left            =   4545
         TabIndex        =   55
         Top             =   120
         Width           =   360
      End
      Begin VB.Label lblFoot 
         AutoSize        =   -1  'True
         Caption         =   "页脚"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   54
         Top             =   2985
         Width           =   360
      End
   End
   Begin XtremeSuiteControls.TabControl tabPageSet 
      Height          =   1665
      Left            =   1515
      TabIndex        =   0
      Top             =   135
      Width           =   2175
      _Version        =   589884
      _ExtentX        =   3836
      _ExtentY        =   2937
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmPageSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean, Doc As cTableEPR, mblnKind As Boolean
Private maryItems() As String
Public Function ShowMe(ByVal frmObj As Object, TmpDoc As cTableEPR) As Boolean
Dim i As Integer
    On Error GoTo errHand
    Set Doc = TmpDoc
    '添加纸张
    With cboKind
        .Clear
        For i = 1 To 42
            .AddItem GetPaperName(i)
            .ItemData(.NewIndex) = Split(GetPaperName(i), ",")(9)
        Next
    End With
    
    '显示页眉页脚
    txtHead.Text = Round(Me.ScaleY(Doc.EPRFileInfo.HeadMargin, vbTwips, vbMillimeters), 2): If txtHead.Text < 0 Then txtHead.Text = 0
    txtFoot.Text = Round(Me.ScaleY(Doc.EPRFileInfo.FootMargin, vbTwips, vbMillimeters), 2): If txtFoot.Text < 0 Then txtFoot.Text = 0
    With Doc.EPRFileInfo
        txtHeadContent.FontName = .HeadFontName: txtHeadContent.FontSize = .HeadFontSize: txtHeadContent.FontBold = .HeadFontBold: txtHeadContent.FontItalic = .HeadFontItalic
        txtHeadContent.FontUnderline = .HeadFontUnderline: txtHeadContent.Font.Strikethrough = .HeadFontStrikethrough: txtHeadContent.ForeColor = .HeadFontColor
        txtHeadContent.Text = .HeadConText
    
        txtFootContent.FontName = .FootFontName: txtFootContent.FontSize = .FootFontSize: txtFootContent.FontBold = .FootFontBold: txtFootContent.FontItalic = .FootFontItalic
        txtFootContent.FontUnderline = .FootFontUnderline: txtFootContent.Font.Strikethrough = .FootFontStrikethrough: txtFootContent.ForeColor = .FootFontColor
        txtFootContent.Text = .FootConText
    End With
    
    If Doc.EPRFileInfo.HeadPic.Handle <> 0 Then
        picHead.AutoRedraw = True: picHead.ZOrder 0
        Call picHead.PaintPicture(Doc.EPRFileInfo.HeadPic, 0, 0, picHead.Width, picHead.Height)
    End If
    
    '显示纸张
    cboKind.ListIndex = SeekCboIndex(cboKind, Doc.EPRFileInfo.PaperKind)
    If Doc.EPRFileInfo.PaperOrient = vbPRORPortrait Then '纵向
        optOrient(0).Value = True
        txtMarjin(0).Text = Round(Me.ScaleY(Doc.EPRFileInfo.MarginTop, vbTwips, vbMillimeters), 2)
        txtMarjin(1).Text = Round(Me.ScaleY(Doc.EPRFileInfo.MarginBottom, vbTwips, vbMillimeters), 2)
        txtMarjin(2).Text = Round(Me.ScaleX(Doc.EPRFileInfo.MarginLeft, vbTwips, vbMillimeters), 2)
        txtMarjin(3).Text = Round(Me.ScaleX(Doc.EPRFileInfo.MarginRight, vbTwips, vbMillimeters), 2)
        If Val(txtHead.Text) = 0 Then txtHead.Text = txtMarjin(0).Text
        If Val(txtFoot.Text) = 0 Then txtFoot.Text = txtMarjin(1).Text
    Else
        optOrient(1).Value = True
        txtMarjin(2).Text = Round(Me.ScaleY(Doc.EPRFileInfo.MarginTop, vbTwips, vbMillimeters), 2)
        txtMarjin(3).Text = Round(Me.ScaleY(Doc.EPRFileInfo.MarginBottom, vbTwips, vbMillimeters), 2)
        txtMarjin(0).Text = Round(Me.ScaleX(Doc.EPRFileInfo.MarginLeft, vbTwips, vbMillimeters), 2)
        txtMarjin(1).Text = Round(Me.ScaleX(Doc.EPRFileInfo.MarginRight, vbTwips, vbMillimeters), 2)
        If Val(txtHead.Text) = 0 Then txtHead.Text = txtMarjin(2).Text
        If Val(txtFoot.Text) = 0 Then txtFoot.Text = txtMarjin(3).Text
    End If
    If cboKind.ListIndex = cboKind.ListCount - 1 Then
        txtHeight.Text = CInt(Me.ScaleY(Doc.EPRFileInfo.PaperHeight, vbTwips, vbMillimeters))
        txtWidth.Text = CInt(Me.ScaleX(Doc.EPRFileInfo.PaperWidth, vbTwips, vbMillimeters))
    End If

    '添加预定义内容
    With Me.cboMode
        .AddItem "第[页码]页":                                      .AddItem "第[页码]页，共[总页数]页"
        .AddItem "文件：[文件名]":                                  .AddItem "打印日期：[打印日期]"
        .AddItem "打印时间：[打印时间]":                            .AddItem "[单位名称][病历名称]"
        .AddItem "姓名：[姓名]    性别：[性别]    年龄：[年龄]    标识号：[标识号]"
        .AddItem "门诊号：[门诊号]":                                .AddItem "住院号：[住院号]"
        .AddItem "科室：[入院科室]":                                .AddItem "病区：[入院病区]"
        .AddItem "科室：[当前科室]":                                .AddItem "床号：[当前床号]"
        .AddItem "住院日期：[入院日期]～[出院日期]":                .AddItem "第[住院次数]住院"
        .AddItem "经治医师：[住院医师]":                            .AddItem "责任护士：[责任护士]"
        .ListIndex = 0
    End With
    
    mblnOK = False
    Me.Show 1, frmObj
    ShowMe = mblnOK
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboKind_Click()
    mblnKind = True
    maryItems = Split(cboKind.Text, ",")

    Me.txtHeight.Text = Round(Me.ScaleY(maryItems(1), vbTwips, vbMillimeters), 2)
    Me.txtWidth.Text = Round(Me.ScaleX(maryItems(2), vbTwips, vbMillimeters), 2)
    Me.txtMarjin(0).Text = Round(Me.ScaleY(maryItems(3), vbTwips, vbMillimeters), 2)
    Me.txtMarjin(1).Text = Round(Me.ScaleY(maryItems(4), vbTwips, vbMillimeters), 2)
    Me.txtMarjin(2).Text = Round(Me.ScaleX(maryItems(5), vbTwips, vbMillimeters), 2)
    Me.txtMarjin(3).Text = Round(Me.ScaleX(maryItems(6), vbTwips, vbMillimeters), 2)
    Me.optOrient(0).Value = True
    If Round(Me.txtHead.Text, 2) < Round(Me.ScaleY(maryItems(7), vbTwips, vbMillimeters), 2) Then
        Me.txtHead.Text = Round(Me.ScaleY(maryItems(7), vbTwips, vbMillimeters), 2)
    End If
    If Round(Me.txtFoot.Text, 2) < Round(Me.ScaleY(maryItems(8), vbTwips, vbMillimeters), 2) Then
        Me.txtFoot.Text = Round(Me.ScaleY(maryItems(8), vbTwips, vbMillimeters), 2)
    End If
    
    Call RedrawSample
    mblnKind = False
End Sub

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Unload Me
End Sub

Private Sub cmdFont_Click(Index As Integer)
Dim lngCorlor As Long, tmpFont As StdFont
    If Index = 0 Then
        lngCorlor = txtHeadContent.ForeColor: Set tmpFont = txtHeadContent.Font
        If SetFont(Me.hWnd, Me.hdc, tmpFont, lngCorlor) Then Set txtHeadContent.Font = tmpFont: txtHeadContent.ForeColor = lngCorlor
    Else
        lngCorlor = txtFootContent.ForeColor: Set tmpFont = txtFootContent.Font
        If SetFont(Me.hWnd, Me.hdc, tmpFont, lngCorlor) Then Set txtFootContent.Font = tmpFont: txtFootContent.ForeColor = lngCorlor
    End If
End Sub

Private Sub cmdFoot_Click()
    txtFootContent.Text = txtFootContent.Text & Space(1) & cboMode.Text
End Sub

Private Sub cmdHead_Click()
    txtHeadContent.Text = txtHeadContent.Text & Space(1) & cboMode.Text
End Sub

Private Sub cmdOK_Click()
    If Not ValidSet() Then Exit Sub
    maryItems = Split(Me.cboKind.Text, ",")
    
    With Doc.EPRFileInfo
        '纸张
        .PaperKind = cboKind.ItemData(cboKind.ListIndex): .PaperOrient = IIf(optOrient(0).Value, vbPRORPortrait, vbPRORLandscape)
        If .PaperOrient = vbPRORPortrait Then '纵向
            If .PaperKind <> cboKind.ItemData(cboKind.ListCount - 1) Then
                .PaperHeight = maryItems(1): .PaperWidth = maryItems(2)
            Else
                .PaperHeight = CInt(Me.ScaleY(txtHeight.Text, vbMillimeters, vbTwips)): .PaperWidth = CInt(ScaleX(txtWidth.Text, vbMillimeters, vbTwips))
            End If
            .MarginTop = Int(ScaleY(txtMarjin(0).Text, vbMillimeters, vbTwips)): .MarginBottom = Int(ScaleY(txtMarjin(1).Text, vbMillimeters, vbTwips))
            .MarginLeft = Int(ScaleX(txtMarjin(2).Text, vbMillimeters, vbTwips)): .MarginRight = Int(ScaleY(txtMarjin(3).Text, vbMillimeters, vbTwips))
        Else
            If .PaperKind <> cboKind.ListCount - 1 Then
                .PaperHeight = maryItems(2): .PaperWidth = maryItems(1)
            Else
                .PaperHeight = Int(Me.ScaleY(txtWidth.Text, vbMillimeters, vbTwips)): .PaperWidth = Int(ScaleX(txtHeight.Text, vbMillimeters, vbTwips))
            End If
            .MarginTop = Int(ScaleY(txtMarjin(2).Text, vbMillimeters, vbTwips)): .MarginBottom = Int(ScaleY(txtMarjin(3).Text, vbMillimeters, vbTwips))
            .MarginLeft = Int(ScaleX(txtMarjin(0).Text, vbMillimeters, vbTwips)): .MarginRight = Int(ScaleY(txtMarjin(1).Text, vbMillimeters, vbTwips))
        End If
        
        '页眉页脚
        .HeadConText = Replace(txtHeadContent.Text, "'", "’"):      .HeadFontName = txtHeadContent.FontName:       .HeadFontSize = txtHeadContent.FontSize
        .HeadFontBold = txtHeadContent.FontBold:   .HeadFontItalic = txtHeadContent.FontItalic:   .HeadFontUnderline = txtHeadContent.FontUnderline
        .HeadFontStrikethrough = txtHeadContent.FontStrikethru: .HeadFontColor = txtHeadContent.ForeColor
        If picHead.Picture.Handle <> 0 Then
            Set .HeadPic = picHead.Picture
        End If
        
        .FootConText = Replace(txtFootContent.Text, "'", "’"):        .FootFontName = txtFootContent.FontName:       .FootFontSize = txtFootContent.FontSize
        .FootFontBold = txtFootContent.FontBold:   .FootFontItalic = txtFootContent.FontItalic:   .FootFontUnderline = txtFootContent.FontUnderline
        .FootFontStrikethrough = txtFootContent.FontStrikethru: .FootFontColor = txtFootContent.ForeColor
        
        .HeadMargin = CInt(Me.ScaleY(txtHead.Text, vbMillimeters, vbTwips)): If .HeadMargin = 0 Then .HeadMargin = .MarginTop
        .FootMargin = CInt(Me.ScaleY(txtFoot.Text, vbMillimeters, vbTwips)): If .FootMargin = 0 Then .FootMargin = .MarginBottom

    End With
    
    mblnOK = True: Unload Me
End Sub

Private Sub cmdPicture_Click(Index As Integer)
Dim strFile As String
    If Index = 1 Then
        Set picHead.Picture = Nothing
    Else
        strFile = GetOpenFile(Me.hWnd, "", "图像文件(*.jpg;*.gif;*.ico;*.bmp;jpeg)" & Chr(0) & "*.jpg;*.gif;*.ico;*.bmp;jpeg" & Chr(0) & Chr(0), "导入页眉图片")
        If strFile <> "" Then
            Set picHead.Picture = LoadPicture(strFile): picHead.ZOrder 0
        End If
        If picHead.Picture.Handle <> 0 Then
            picHead.AutoRedraw = True
            Call picHead.PaintPicture(picHead.Picture, 0, 0, picHead.Width, picHead.Height)
        End If
    End If
End Sub

Private Sub Form_Load()
    With tabPageSet
        .Top = 0: .Left = 0: .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - picButton.Height
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        .InsertItem 1, "页面设置", picPage.hWnd, 0
        .InsertItem 2, "页眉页脚", picHeadFoot.hWnd, 0
        
        .Item(0).Selected = True
        picPage.BackColor = &H8000000F: picHeadFoot.BackColor = &H8000000F
    End With
End Sub

Private Sub optOrient_Click(Index As Integer)
 Dim strCaption As String
    
    strCaption = Me.lblWidth.Caption
    Me.lblWidth.Caption = Me.lblHeight.Caption
    Me.lblHeight.Caption = strCaption
    
    strCaption = Me.lblMarjin(0).Caption
    Me.lblMarjin(0).Caption = Me.lblMarjin(2).Caption
    Me.lblMarjin(2).Caption = strCaption
    
    strCaption = Me.lblMarjin(1).Caption
    Me.lblMarjin(1).Caption = Me.lblMarjin(3).Caption
    Me.lblMarjin(3).Caption = strCaption
    
    Call RedrawSample

End Sub

Private Sub txtFoot_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtHead_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtHeight_Change()
    If mblnKind Then Exit Sub
    Me.cboKind.ListIndex = Me.cboKind.ListCount - 1
    Call RedrawSample
End Sub
Private Sub txtHeight_GotFocus()
    With txtHeight
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    Call RedrawMarjin(Index)
End Sub

Private Sub txtMarjin_GotFocus(Index As Integer)
    With txtMarjin(Index)
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtMarjin_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    If mblnKind Then Exit Sub
    Me.cboKind.ListIndex = Me.cboKind.ListCount - 1
    Call RedrawSample
End Sub

Private Sub txtWidth_GotFocus()
    With txtWidth
        .SelStart = 0: .SelLength = Len(.Text)
    End With
    Call OpenIme(False)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call PressKey(vbKeyTab): Exit Sub
    If InStr("1234567890." & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function ValidSet() As Boolean
'功能：检查设置的合理性，并提示进行自动调整
'cboKind.txt 名称,高度,宽度,边距上,下,左,右,页眉边距,页脚边距,打印纸张序号
    Dim dblMarjin As Double
    maryItems = Split(Me.cboKind.Text, ",")
    
    '自定义纸张，需要检测宽度高度是否超过边界
    If Me.cboKind.ListIndex = Me.cboKind.ListCount - 1 Then
        If Val(txtHeight.Text) = 0 Or Val(txtWidth.Text) = 0 Then
            MsgBox "请指定纸张宽度和高度！", vbInformation, gstrSysName
            Exit Function
        End If
        If Val(txtHeight.Text) > Round(Me.ScaleY(maryItems(1), vbTwips, vbMillimeters), 2) Then
            ValidSet = False
            If MsgBox(IIf(optOrient(0).Value = True, "高度", "宽度") & "超过自定义纸张限制。是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtHeight.Text = Round(Me.ScaleY(maryItems(1), vbTwips, vbMillimeters), 2)
            Else
                Exit Function
            End If
            tabPageSet.Item(0).Selected = True: zlControl.TxtSelAll txtHeight
        End If
        If Val(txtWidth.Text) > Round(Me.ScaleX(maryItems(2), vbTwips, vbMillimeters), 2) Then
            ValidSet = False
            If MsgBox(IIf(optOrient(0).Value = True, "宽度", "高度") & "超过自定义纸张限制。是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtWidth.Text = Round(Me.ScaleX(maryItems(2), vbTwips, vbMillimeters), 2)
            Else
                Exit Function
            End If
            tabPageSet.Item(0).Selected = True: zlControl.TxtSelAll txtWidth
        End If
    End If
    
    '页眉边距判断'页脚边距判断
    If optOrient(0).Value Then
        If Val(txtHead.Text) = 0 Or Val(txtFoot.Text) = 0 Then
            MsgBox "请指定页眉边距和页脚边距！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Int(Me.ScaleY(Val(txtHead.Text), vbMillimeters, vbTwips)) > Int(Me.ScaleY(Val(txtMarjin(0)), vbMillimeters, vbTwips)) Then
            If MsgBox("页眉边距不能大于上边距，是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtHead.Text = txtMarjin(0).Text
            Else
                Exit Function
            End If
            tabPageSet.Item(1).Selected = True: zlControl.TxtSelAll txtHead
        End If
        
        If Int(Me.ScaleY(Val(txtFoot.Text), vbMillimeters, vbTwips)) > Int(Me.ScaleY(Val(txtMarjin(1)), vbMillimeters, vbTwips)) Then
            If MsgBox("页脚边距不能大于下边距，是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtFoot.Text = txtMarjin(1).Text
            Else
                Exit Function
            End If
            tabPageSet.Item(1).Selected = True: zlControl.TxtSelAll txtFoot
        End If
    Else
        If Int(Me.ScaleY(Val(txtHead.Text), vbMillimeters, vbTwips)) > Int(Me.ScaleY(Val(txtMarjin(2)), vbMillimeters, vbTwips)) Then
            If MsgBox("页眉边距不能大于上边距，是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtHead.Text = txtMarjin(3).Text
            Else
                Exit Function
            End If
            tabPageSet.Item(1).Selected = True: Call zlControl.TxtSelAll(txtHead)
        End If
        
        If Int(Me.ScaleY(Val(txtFoot.Text), vbMillimeters, vbTwips)) > Int(Me.ScaleY(Val(txtMarjin(3)), vbMillimeters, vbTwips)) Then
            If MsgBox("页脚边距不能大于下边距，是否自动调整？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                txtFoot.Text = txtMarjin(3).Text
            Else
                Exit Function
            End If
            tabPageSet.Item(1).Selected = True: zlControl.TxtSelAll txtFoot
        End If
    End If

    ValidSet = True
End Function

Private Sub RedrawSample()
    '功能：重新绘制页面示范
    Dim dblWidth As Double, dblHeight As Double
    
    If Val(Trim(txtWidth.Text)) = 0 Then Exit Sub
    If Val(Trim(txtHeight.Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(0).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(1).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(2).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(3).Text)) = 0 Then Exit Sub
    
    If optOrient(0).Value Then
        dblWidth = Val(txtWidth.Text): dblHeight = Val(txtHeight.Text)
    Else
        dblWidth = Val(txtHeight.Text): dblHeight = Val(txtWidth.Text)
    End If
    
    With picPaper
        If dblWidth < dblHeight Then
            .Top = 45: .Height = picViewer.ScaleHeight - 90
            .Width = .Height / dblHeight * dblWidth
            .Left = (picViewer.ScaleWidth - .Width) / 2
        Else
            .Left = 45: .Width = picViewer.ScaleWidth - 90
            .Height = .Width / dblWidth * dblHeight
            .Top = (picViewer.ScaleHeight - .Height) / 2
        End If
    End With
    
    Call RedrawMarjin(0)
    Call RedrawMarjin(1)
    Call RedrawMarjin(2)
    Call RedrawMarjin(3)

End Sub
Private Sub RedrawMarjin(Index As Integer)
    '功能：重新绘制指定的边距示范线
    '参数：index，0、1、2、3分别为上下左右边距设置，在方向变化时和边距线对应关系变化
    If Val(Trim(txtWidth.Text)) = 0 Then Exit Sub
    If Val(Trim(txtHeight.Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(0).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(1).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(2).Text)) = 0 Then Exit Sub
    If Val(Trim(txtMarjin(3).Text)) = 0 Then Exit Sub
    
    Select Case Index
    Case 0
        If optOrient(0).Value Then
            With linMarjin(0)
                .X1 = 0: .X2 = picPaper.ScaleWidth - 15
                .Y1 = Val(txtMarjin(0).Text) / Val(txtHeight.Text) * (picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        Else
            With linMarjin(2)
                .X1 = Val(txtMarjin(0).Text) / Val(txtHeight.Text) * (picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = picPaper.ScaleHeight - 15
            End With
        End If
    Case 1
        If optOrient(0).Value Then
            With linMarjin(1)
                .X1 = 0: .X2 = picPaper.ScaleWidth - 15
                .Y1 = (1 - Val(txtMarjin(1).Text) / Val(txtHeight.Text)) * (picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        Else
            With linMarjin(3)
                .X1 = (1 - Val(txtMarjin(1).Text) / Val(txtHeight.Text)) * (picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = picPaper.ScaleHeight - 15
            End With
        End If
    Case 2
        If optOrient(0).Value Then
            With linMarjin(2)
                .X1 = Val(txtMarjin(2).Text) / Val(txtWidth.Text) * (picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = picPaper.ScaleHeight - 15
            End With
        Else
            With linMarjin(0)
                .X1 = 0: .X2 = picPaper.ScaleWidth - 15
                .Y1 = Val(txtMarjin(2).Text) / Val(txtWidth.Text) * (picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        End If
    Case 3
        If optOrient(0).Value Then
            With linMarjin(3)
                .X1 = (1 - Val(txtMarjin(3).Text) / Val(txtWidth.Text)) * (picPaper.ScaleWidth - 15): .X2 = .X1
                .Y1 = 0: .Y2 = picPaper.ScaleHeight - 15
            End With
        Else
            With linMarjin(1)
                .X1 = 0: .X2 = picPaper.ScaleWidth - 15
                .Y1 = (1 - Val(txtMarjin(3).Text) / Val(txtWidth.Text)) * (picPaper.ScaleHeight - 15): .Y2 = .Y1
            End With
        End If
    End Select
End Sub

