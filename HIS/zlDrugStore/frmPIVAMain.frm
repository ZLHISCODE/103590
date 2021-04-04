VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmPIVAMain 
   Caption         =   "静脉输液配置中心管理"
   ClientHeight    =   11700
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   17910
   Icon            =   "frmPIVAMain.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   11700
   ScaleWidth      =   17910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraTip 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9240
      TabIndex        =   107
      Top             =   1200
      Width           =   840
      Begin VB.PictureBox pic自备药 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         Picture         =   "frmPIVAMain.frx":058A
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   108
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lbl自备药 
         AutoSize        =   -1  'True
         Caption         =   "自备药"
         Height          =   180
         Index           =   3
         Left            =   255
         TabIndex        =   109
         Top             =   30
         Width           =   540
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   3480
      ScaleHeight     =   780
      ScaleWidth      =   9375
      TabIndex        =   50
      Top             =   7200
      Width           =   9375
      Begin VB.TextBox txtFindItem 
         Appearance      =   0  'Flat
         Height          =   350
         Left            =   7320
         MaxLength       =   13
         TabIndex        =   51
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label lblMsg 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6480
         TabIndex        =   53
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblFindItem 
         AutoSize        =   -1  'True
         Caption         =   "瓶签号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         TabIndex        =   52
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   2
      Left            =   8760
      Picture         =   "frmPIVAMain.frx":6DDC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   45
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   1
      Left            =   8280
      ScaleHeight     =   1695
      ScaleWidth      =   3015
      TabIndex        =   37
      Top             =   8160
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   1200
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Width           =   2760
         _cx             =   4868
         _cy             =   2117
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":D62E
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
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   0
      Left            =   3720
      ScaleHeight     =   1695
      ScaleWidth      =   3015
      TabIndex        =   26
      Top             =   7800
      Width           =   3015
      Begin VSFlex8Ctl.VSFlexGrid vsfDept 
         Height          =   1200
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   2760
         _cx             =   4868
         _cy             =   2117
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":D6DD
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
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "frmPIVAMain.frx":D78C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPacker 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "frmPIVAMain.frx":DD16
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPrint 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   8520
      Picture         =   "frmPIVAMain.frx":E2A0
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picPrint 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   0
      Left            =   8280
      Picture         =   "frmPIVAMain.frx":E82A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picDetailList 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   3600
      ScaleHeight     =   6255
      ScaleWidth      =   13695
      TabIndex        =   13
      Top             =   1800
      Width           =   13695
      Begin VB.TextBox txtDia 
         Enabled         =   0   'False
         Height          =   855
         Left            =   10800
         TabIndex        =   92
         Top             =   3720
         Width           =   2655
      End
      Begin VB.Frame fraMedis 
         Height          =   615
         Left            =   360
         TabIndex        =   79
         Top             =   240
         Width           =   9855
         Begin VB.CheckBox chkCheck 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "全选"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   240
            Width           =   2055
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   84
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5370
            TabIndex        =   83
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   6420
            TabIndex        =   82
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   7470
            TabIndex        =   81
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkResult 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   8520
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblType 
            BackColor       =   &H80000004&
            Caption         =   "配药类型"
            Height          =   180
            Left            =   1200
            TabIndex        =   87
            Top             =   277
            Width           =   780
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":EDB4
            Height          =   240
            Index           =   0
            Left            =   4560
            Picture         =   "frmPIVAMain.frx":15606
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":1BE58
            Height          =   240
            Index           =   1
            Left            =   5610
            Picture         =   "frmPIVAMain.frx":226AA
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":28EFC
            Height          =   240
            Index           =   2
            Left            =   6660
            Picture         =   "frmPIVAMain.frx":2F74E
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":35FA0
            Height          =   240
            Index           =   3
            Left            =   7710
            Picture         =   "frmPIVAMain.frx":3C7F2
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image ImgResult 
            DragIcon        =   "frmPIVAMain.frx":43044
            Height          =   240
            Index           =   4
            Left            =   8760
            Picture         =   "frmPIVAMain.frx":49896
            Top             =   240
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin VB.Frame fraDetailCtr 
         BackColor       =   &H00FFEDDD&
         Height          =   840
         Left            =   -120
         TabIndex        =   54
         Top             =   840
         Width           =   15015
         Begin VB.ComboBox cboSort 
            Height          =   300
            Left            =   12720
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   480
            Width           =   2500
         End
         Begin VB.ComboBox cboFrequency 
            Height          =   300
            Left            =   5760
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox chkSure 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "已确认"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   73
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSure 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "未确认"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   72
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "打包"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   71
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox cboMedi 
            Height          =   300
            Left            =   8040
            Style           =   2  'Dropdown List
            TabIndex        =   70
            Top             =   480
            Width           =   3700
         End
         Begin VB.ComboBox cboLevel 
            Height          =   300
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   480
            Width           =   615
         End
         Begin VB.ComboBox cboBatch 
            Height          =   300
            Left            =   2640
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   480
            Width           =   975
         End
         Begin VB.CheckBox chkSendType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "已发送"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5400
            TabIndex        =   67
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSendType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "未发送"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   4800
            TabIndex        =   66
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "配药"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   65
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkPack 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "按配药(打包)汇总"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1440
            TabIndex        =   64
            Top             =   150
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton optShowType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "简要"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   7080
            TabIndex        =   63
            Top             =   150
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton optShowType 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "详细"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   8040
            TabIndex        =   62
            Top             =   150
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chkAll 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "全选"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   150
            Width           =   735
         End
         Begin VB.CheckBox chkDept 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "按病区汇总"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   150
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "已打印"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   59
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkPrint 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "未打印"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   58
            Top             =   150
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkChange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "已变"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   57
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkChange 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFEDDD&
            Caption         =   "未变"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   56
            Top             =   150
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ComboBox cboDosType 
            Height          =   300
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblSort 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "排序方式"
            Height          =   180
            Left            =   11880
            TabIndex        =   90
            Top             =   540
            Width           =   720
         End
         Begin VB.Label lblFrequency 
            BackColor       =   &H00FFEDDD&
            Caption         =   "频次"
            Height          =   180
            Left            =   5280
            TabIndex        =   89
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblMedi 
            BackColor       =   &H00FFEDDD&
            Caption         =   "药品"
            Height          =   180
            Left            =   7560
            TabIndex        =   78
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblBatch 
            BackColor       =   &H00FFEDDD&
            Caption         =   "批次"
            Height          =   180
            Left            =   2280
            TabIndex        =   77
            Top             =   540
            Width           =   420
         End
         Begin VB.Label lblLevel 
            BackColor       =   &H00FFEDDD&
            Caption         =   "优先级"
            Height          =   180
            Left            =   3840
            TabIndex        =   76
            Top             =   540
            Width           =   540
         End
         Begin VB.Label lblVolu 
            BackColor       =   &H00FFEDDD&
            Caption         =   "容量：0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   9960
            TabIndex        =   75
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblDosType 
            BackColor       =   &H00FFEDDD&
            Caption         =   "类型"
            Height          =   180
            Left            =   120
            TabIndex        =   74
            Top             =   540
            Width           =   420
         End
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   9000
         TabIndex        =   47
         ToolTipText     =   "热键：F2"
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtLog 
         Height          =   855
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   3960
         Width           =   3855
      End
      Begin VB.Frame fraH 
         Height          =   30
         Left            =   240
         MousePointer    =   7  'Size N S
         TabIndex        =   31
         Top             =   2760
         Width           =   9375
      End
      Begin VB.PictureBox picHelp 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   240
         ScaleHeight     =   225
         ScaleWidth      =   9975
         TabIndex        =   22
         Top             =   0
         Width           =   9975
         Begin VB.PictureBox picHelpIcon 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   0
            Picture         =   "frmPIVAMain.frx":500E8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   23
            Top             =   20
            Width           =   240
         End
         Begin VB.Label lblCount 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   5760
            TabIndex        =   28
            Top             =   45
            Width           =   4170
         End
         Begin VB.Label lblHelp 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "提示："
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   260
            TabIndex        =   24
            Top             =   50
            Width           =   540
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfTrans 
         Height          =   840
         Left            =   360
         TabIndex        =   14
         Top             =   1680
         Width           =   4560
         _cx             =   8043
         _cy             =   1482
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   61
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5693A
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
         ExplorerBar     =   2
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
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmPIVAMain.frx":5710D
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfSumDrug 
         Height          =   1200
         Left            =   8280
         TabIndex        =   16
         Top             =   1680
         Width           =   1920
         _cx             =   3387
         _cy             =   2117
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5765B
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
         ExplorerBar     =   3
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfMedis 
         Height          =   840
         Left            =   5400
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   2400
         _cx             =   4233
         _cy             =   1482
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   35
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":5781B
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
         Editable        =   1
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
      Begin VSFlex8Ctl.VSFlexGrid VSFLook 
         Height          =   1440
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   8640
         _cx             =   15240
         _cy             =   2540
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
         BackColorSel    =   16771280
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777215
         GridColor       =   10329501
         GridColorFixed  =   10329501
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   20
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPIVAMain.frx":57C6F
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
         ExplorerBar     =   2
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
         Height          =   855
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   1508
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPIVAMain.frx":57F0D
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
      Begin VB.Label lblLog 
         Caption         =   "审核理由"
         Height          =   255
         Left            =   8280
         TabIndex        =   94
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label lblDia 
         Caption         =   "病人诊断"
         Height          =   255
         Left            =   10800
         TabIndex        =   93
         Top             =   3480
         Width           =   1335
      End
   End
   Begin VB.PictureBox picDetail 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   3480
      ScaleHeight     =   1455
      ScaleWidth      =   2295
      TabIndex        =   10
      Top             =   120
      Width           =   2295
      Begin VB.Frame fraLineV1 
         BackColor       =   &H80000012&
         Height          =   2085
         Left            =   120
         TabIndex        =   11
         Top             =   -120
         Width           =   50
      End
      Begin XtremeSuiteControls.TabControl tbcDetail 
         Height          =   975
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         _Version        =   589884
         _ExtentX        =   2566
         _ExtentY        =   1720
         _StockProps     =   64
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox picCondition 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8055
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.PictureBox picMsg 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   120
         MouseIcon       =   "frmPIVAMain.frx":57F5B
         ScaleHeight     =   2055
         ScaleWidth      =   2895
         TabIndex        =   39
         Tag             =   "0"
         Top             =   6000
         Width           =   2895
         Begin VB.Frame fraMsg 
            Height          =   50
            Left            =   -20
            TabIndex        =   42
            Top             =   0
            Width           =   3405
         End
         Begin VB.PictureBox picUpOrDown 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   2400
            Picture         =   "frmPIVAMain.frx":58265
            ScaleHeight     =   270
            ScaleWidth      =   270
            TabIndex        =   40
            Top             =   60
            Width           =   270
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfMsg 
            Height          =   1560
            Left            =   0
            TabIndex        =   44
            Top             =   405
            Width           =   2880
            _cx             =   5080
            _cy             =   2752
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
            BackColorSel    =   16771280
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   16777215
            GridColor       =   10329501
            GridColorFixed  =   10329501
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPIVAMain.frx":585A7
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
            Editable        =   0
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
         Begin VB.Label lblMsgComment 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "消息提醒(0)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   180
            Left            =   45
            TabIndex        =   41
            Top             =   90
            Width           =   1095
         End
      End
      Begin VB.PictureBox picLook 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   120
         ScaleHeight     =   1305
         ScaleWidth      =   2025
         TabIndex        =   34
         Top             =   4320
         Width           =   2055
         Begin XtremeSuiteControls.TabControl tbcLook 
            Height          =   1215
            Left            =   0
            TabIndex        =   35
            Top             =   240
            Width           =   2055
            _Version        =   589884
            _ExtentX        =   3625
            _ExtentY        =   2143
            _StockProps     =   64
            Enabled         =   -1  'True
         End
      End
      Begin VB.PictureBox picWork 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   1200
         ScaleHeight     =   1305
         ScaleWidth      =   1905
         TabIndex        =   32
         Top             =   3840
         Width           =   1935
         Begin XtremeSuiteControls.TabControl tabWork 
            Height          =   1095
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1695
            _Version        =   589884
            _ExtentX        =   2990
            _ExtentY        =   1931
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picDeptList 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   2895
         TabIndex        =   8
         Top             =   3840
         Width           =   2895
         Begin VB.CommandButton cmdRefreshTrans 
            Caption         =   "刷新明细"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1920
            TabIndex        =   48
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox chkAllDept 
            Appearance      =   0  'Flat
            Caption         =   "全选"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   43
            Top             =   40
            Width           =   735
         End
         Begin VB.Frame fraLineH1 
            Height          =   50
            Left            =   -20
            TabIndex        =   9
            Top             =   0
            Width           =   3405
         End
         Begin XtremeSuiteControls.TabControl tabDeptList 
            Height          =   1455
            Left            =   120
            TabIndex        =   36
            Top             =   0
            Width           =   2535
            _Version        =   589884
            _ExtentX        =   4471
            _ExtentY        =   2566
            _StockProps     =   64
         End
      End
      Begin VB.PictureBox picTime 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   3135
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtdept 
            Height          =   315
            Left            =   840
            TabIndex        =   105
            Top             =   1920
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdDrug 
            Caption         =   "..."
            Height          =   255
            Left            =   2640
            TabIndex        =   104
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txtTag 
            Height          =   315
            Left            =   840
            TabIndex        =   103
            Top             =   3120
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtDrug 
            Height          =   315
            Left            =   840
            TabIndex        =   101
            Top             =   2700
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   840
            TabIndex        =   99
            Top             =   2280
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.PictureBox picShowSendType 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            MouseIcon       =   "frmPIVAMain.frx":586C4
            ScaleHeight     =   270
            ScaleWidth      =   3015
            TabIndex        =   95
            Tag             =   "0"
            Top             =   1560
            Width           =   3015
            Begin VB.PictureBox picUpOrDown1 
               BackColor       =   &H00FFEDDD&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   2640
               Picture         =   "frmPIVAMain.frx":589CE
               ScaleHeight     =   270
               ScaleWidth      =   270
               TabIndex        =   96
               Top             =   0
               Width           =   270
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "其他过滤条件"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   120
               TabIndex        =   97
               Top             =   45
               Width           =   1080
            End
         End
         Begin VB.ComboBox cbo时间范围 
            Height          =   300
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   420
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker Dtp结束时间 
            Height          =   315
            Left            =   885
            TabIndex        =   3
            Top             =   1140
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   116654083
            CurrentDate     =   39998
         End
         Begin MSComCtl2.DTPicker Dtp开始时间 
            Height          =   300
            Left            =   885
            TabIndex        =   4
            Top             =   780
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   116654083
            CurrentDate     =   39998
         End
         Begin VB.Label lbldept 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "科 室"
            Height          =   180
            Left            =   240
            TabIndex        =   106
            Top             =   1980
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblTag 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "瓶签号"
            Height          =   180
            Left            =   225
            TabIndex        =   102
            Top             =   3180
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblDrug 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "药 品"
            Height          =   180
            Left            =   240
            TabIndex        =   100
            Top             =   2760
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "姓名↓"
            Height          =   180
            Left            =   225
            TabIndex        =   98
            Top             =   2340
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "输液单执行时间范围"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   45
            TabIndex        =   15
            Top             =   120
            Width           =   1755
         End
         Begin VB.Label lblTimeBegin 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "开始时间"
            Height          =   180
            Left            =   45
            TabIndex        =   7
            Top             =   840
            Width           =   720
         End
         Begin VB.Label lblTimeEnd 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "结束时间"
            Height          =   180
            Left            =   45
            TabIndex        =   6
            Top             =   1200
            Width           =   720
         End
         Begin VB.Label lbl时间范围 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            Caption         =   "时间范围"
            Height          =   180
            Left            =   45
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   11340
      Width           =   17916
      _ExtentX        =   31591
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPIVAMain.frx":58D10
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24712
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   840
      Left            =   6000
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   2280
      _cx             =   4022
      _cy             =   1482
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
      BackColorSel    =   16771280
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPIVAMain.frx":595A4
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
   Begin MSComctlLib.ImageList ImgList 
      Left            =   10080
      Top             =   360
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
            Picture         =   "frmPIVAMain.frx":59632
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":5FE94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":666F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":6CF58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":737BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":7A01C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":8087E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":870E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgPro 
      Left            =   10800
      Top             =   360
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
            Picture         =   "frmPIVAMain.frx":8D942
            Key             =   "不取药"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPIVAMain.frx":941A4
            Key             =   "自备药"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   7320
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmPIVAMain.frx":9AA06
      Left            =   6480
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmPIVAMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngMode As Long
Private mstrPrivs As String
Private mblnLoad As Boolean
Private mblnActive As Boolean
Private mobjCISJOB As Object  '电子病案查阅对象
Private mobjPlugIn As Object    '外挂接口对象
    
Private mrsDeptAdvice As ADODB.Recordset        '病区对应的医嘱数
Private mrsTrans As ADODB.Recordset             '输液单记录，包含输液单内容（药品）
Private mrsDeptTrans As ADODB.Recordset         '病区对应的输液单数

Private mrsWorkBatch As ADODB.Recordset         '输液配置中心的工作批次

Private mstr上次病区ID As String                '上次选择的病区
Private mstr上次IDS As String                   '上次过滤数据
Public mblnParamsRefresh As Boolean
Private mstrFilter As String
Private mstrUnVisble  As String
Private mstrUnallowSetColHide As String
Private mblnFilter As Boolean

Private mstrCenterName As String
Private mlngPassPati As Long

'消息相关对象变量
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule
Attribute mobjMipModule.VB_VarHelpID = -1
Private mdateToday As Date                      '今日日期
Private mrsMsg As Recordset
Private mrsSendMsg As Recordset

Private mstrLastLabel As String                 '上次选择的瓶签号
Private mintCountPack As Integer                '摆药之后切换打包状态的单据数量
Private mintBeginRow As Integer
Private mintEndRow As Integer
Private mlng已扫描  As Long
Private mlng未扫描 As Long
Private mlngNum As Long


Private mfrmPIVCard As frmPIVCard
Private mfrmPrintPlan As frmPrintPlan
Private mfrmPlan As frmPlan

Private mstr批次 As String
Private mstr打包 As String

Private mint标志 As Integer

Private mstr配药id As String
Private mblnLock As Boolean

Private mrsPRI As Recordset
Private mrsVol As Recordset
Private mrstemp As Recordset
Private mrsMedi As Recordset

Private mblnShowOhters As Boolean             '是否显示自备药与不取药

'输液单列表中，允许修改的列颜色
Private Const CSTCOLOR_MODIFY = &HE1FFE1        '浅绿色
'输液单列表中，不允许修改的列颜色
Private Const CSTCOLOR_UNMODIFY = &H80000005    '白色
'各种状态的输液单有记录时按钮的颜色
Private Const CSTCOLOR_RECORDS = &HE1FFE1       '浅绿色
'各种状态的输液单没有记录时按钮的颜色
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '灰色
'当前状态按钮的颜色
Private Const CSTCOLOR_COMMAND = &HFFEDDD       '浅蓝色

'权限
Private Type Type_Privs
    bln核查确认 As Boolean
    bln取消审核 As Boolean
    bln摆药确认 As Boolean
    bln取消摆药 As Boolean
    bln配药确认 As Boolean
    bln取消配药 As Boolean
    bln发送确认 As Boolean
    bln取消发送 As Boolean
    bln参数设置 As Boolean
    bln销帐审核 As Boolean
    bln确认拒绝 As Boolean
    bln销帐拒绝 As Boolean
    bln排班设置 As Boolean
End Type
Private mPrives As Type_Privs

'使用到的参数（来自系统参数表或其它参数表或本机注册表）
Private Type Type_Params
    '参数表中的系统参数
    lng配置中心 As Long
    bln允许未审核处方发药 As Boolean
    bln允许未收费处方发药 As Boolean
    bln允许取消发药 As Boolean
    bln医嘱作废 As Boolean
    bln审核划价单 As Boolean
    bln报警包含划价费用 As Boolean
    int药品名称显示 As Integer          '0-编码和名称，1-仅编码，2-仅名称

    '参数表中的其它参数
    int摆药后打印 As Integer            '0-提示打印;1-自动打印;2-不打
    int发送后打印 As Integer            '0-提示打印;1-自动打印;2-不打
    bln批次设置 As Boolean
    bln打包设置 As Boolean
    bln审核 As Boolean
    int皮试有效天数 As Integer          '皮试有效天数
    int打印汇总 As Integer            '0-提示打印;1-自动打印;2-不打
    blnLastBatch As Boolean             '保持上次批次
    bln处方审查 As Boolean
    blnByMedi As Boolean                '按药品,批次排序
    blnFilter As Boolean             '是否按设置的常用药品进行过滤
    
'    int瓶签自动打印 As Integer          '0-不自动打印;1-摆药后自动打印;2-配药后自动打印
    int瓶签摆药后打印 As Integer        '0-提示打印;1-自动打印;2-不打
    int瓶签配药后打印 As Integer        '0-提示打印;1-自动打印;2-不打
    bln瓶签手工打印 As Boolean
    strBatchList  As String             '工作批次列表
    intCount As Integer                 '卡片模式下，单行显示列数
    intNum As Integer
    str常用药品 As String
    blnTwoCode As Boolean               '扫描一次进行发送或者配药操作
    intCheck As Integer                 '审核该药房所有医嘱
    blnRePeople As Boolean              '打印瓶签时是否填写实际操作员
    
    '注册表参数
    intFont As Integer                  '表格字体大小
    intAutoSelect As Integer            '自动选择上次选择的输液单
    strSort As String                   '输液单排序
    strVsfTrans As String               '明细表格列宽
    strVsfLook As String                 '已摆药表格列宽
    strVsfSum As String                  '汇总表格列宽
    '库存检查
    IntCheckStock As Integer            '0-不检查;1-不足提醒;2-不足禁止
    
    '是否显示合理用药（PASS）
    intShowPass As Integer
    
    int病区排序 As Integer
    int药品名称显示方式 As Integer      '0-编码和名称，1-仅名称，2-仅编码
    strSourceDep As String              '显示来源病区
End Type
Private mParams As Type_Params

Private Type Type_Condition
    lngCenterID As Long            '输液配置中心的部门ID
    strCenterName As String
    intTransTimeSel As Integer
    strTransStartTime As String
    strTransEndTime As String
    strTransStep As String
End Type
Private mcondition As Type_Condition

'明细分页
Private Enum mDetailType
    输液单列表 = 0
    输液单卡片
    药品汇总列表
End Enum

'业务/查看分页
Private Const CNUMWORK = 0
Private Const CNUMLOOK = 1

'业务分类/步骤
Private Const M_STR_CALSS_AUDIT = "00"              '审核医嘱
Private Const M_STR_CALSS_PREPARE = "01"            '摆药印签
Private Const M_STR_CALSS_DOSAGE = "02"             '配药核查
Private Const M_STR_CALSS_SEND = "03"               '发送核查
Private Const M_STR_CALSS_VERIFY = "04"             '销帐审核
Private Const M_STR_CALSS_PASSEDAUDIT = "10"        '审核已通过医嘱
Private Const M_STR_CALSS_FAILAUDIT = "11"          '审核未通过医嘱
Private Const M_STR_CALSS_SENDED = "12"             '已发送查看
Private Const M_STR_CALSS_SIGNED = "13"             '已签收查看
Private Const M_STR_CALSS_REFUSETOSIGN = "14"       '拒绝签收查看
Private Const M_STR_CALSS_INVALID = "15"            '已作废查看
Private Const M_STR_CALSS_DEVICERETURN = "16"       '医嘱回退查看

Private Enum mTransStatus
    填制 = 1
    摆药 = 2
    校对 = 3
    配药 = 4
    发送 = 5
    签收 = 6
    拒绝签收 = 7
    确认拒收 = 8
    销帐申请 = 9
    销帐审核通过 = 10
    销账审核未通过 = 11
End Enum


'业务汇总表格
Private Const MINTSUMCOLS = 13      '总列数
Private mintcolsum病区 As Integer
Private mintcolsum打包 As Integer
Private mintcolsum药品名称 As Integer
Private mintcolsum商品名 As Integer
Private mintcolsum英文名 As Integer
Private mintcolsum规格 As Integer
Private mintcolsum产地 As Integer
Private mintcolsum批号 As Integer
Private mintcolsum数量 As Integer
Private mintcolsum发药数量 As Integer
Private mintcolsum库存数量 As Integer
Private mintcolsum缺药标志 As Integer
Private mintcolsum是否打包 As Integer


Private Const MINTCOLS = 63      '总列数
'Private mIntCol瓶签号 As Integer
'Private mintcol批次 As Integer
'Private mIntCol摆药人 As Integer
'Private mIntCol摆药时间 As Integer
'Private mIntCol摆药单号 As Integer
'Private mIntCol医嘱发送时间 As Integer
'Private mIntCol执行时间 As Integer
'Private mIntCol药品名称 As Integer
'Private mintcol规格 As Integer
'Private mIntCol单量 As Integer
'Private mintcol数量 As Integer
'Private mIntColNO As Integer
'Private mIntCol单据 As Integer
'Private mIntCol剂量单位 As Integer
'Private mIntCol用法 As Integer
'Private mintcol药品id As Integer
'Private mIntCol配药id As Integer
Private mIntCol当前行 As Integer
Private mIntCol审 As Integer
Private mintcol选择 As Integer
Private mIntCol变 As Integer
Private mIntCol锁 As Integer
Private mIntCol医嘱     As Integer
Private mIntCol调 As Integer
Private mIntCol打印 As Integer
Private mIntCol打包 As Integer
Private mintcol批次 As Integer
Private mIntCol拒收原因 As Integer
Private mIntCol优先级 As Integer
Private mIntCol病区 As Integer
Private mIntCol科室 As Integer
Private mIntCol姓名 As Integer
Private mIntCol性别 As Integer
Private mIntCol年龄 As Integer
Private mIntCol床号 As Integer
Private mIntCol住院号 As Integer
Private mIntCol警 As Integer
Private mIntCol药品名称 As Integer
Private mIntCol皮 As Integer
Private mintcol规格 As Integer
Private mIntCol配药类型 As Integer
Private mIntCol单量 As Integer
Private mintcol数量 As Integer
Private mIntCol执行时间 As Integer
Private mIntCol执行频次 As Integer
Private mIntCol瓶签号 As Integer
Private mIntCol摆药单号 As Integer
Private mIntCol医嘱发送时间 As Integer
Private mIntCol摆药人 As Integer
Private mIntCol摆药时间 As Integer
Private mIntCol配药人 As Integer
Private mIntCol配药时间 As Integer
Private mIntCol发送人 As Integer
Private mIntCol发送时间 As Integer
Private mIntCol销帐申请人 As Integer
Private mIntCol销帐申请时间 As Integer
Private mIntCol销帐审核人 As Integer
Private mIntCol销帐审核时间 As Integer
Private mIntCol销帐原因 As Integer
Private mIntCol操作状态 As Integer

Private mIntCol标志 As Integer
Private mIntCol是否锁定 As Integer
Private mIntCol作废类型 As Integer
Private mIntColNO As Integer
Private mIntCol单据 As Integer
Private mIntCol剂量单位 As Integer
Private mIntCol用法 As Integer
Private mintcol药品id As Integer
Private mIntCol核查人 As Integer
Private mIntCol核查时间 As Integer
Private mIntCol打印标志 As Integer
Private mIntCol配药id As Integer
Private mIntCol是否打包 As Integer
Private mIntCol原批次 As Integer
Private mIntCol抗菌药物 As Integer
Private mIntCol主页id As Integer
Private mIntCol病人ID As Integer
Private mIntCol背景号 As Integer
Private mIntCol警告 As Integer
Private mIntCol溶媒 As Integer
Private mIntCol对应医嘱ID As Integer

'Private Enum mTransStep
'    审核医嘱 = 0
'    摆药印签
'    配药核查
'    发送核查
'    销帐审核
'End Enum
'
'Private Enum mTransLook
'    审核已通过医嘱 = 0
'    审核未通过医嘱
'    已发送查看
'    已签收查看
'    拒绝签收查看
'    已作废查看
'End Enum

'弹出菜单
Private Const conMenu_OperPopup = 300                   '操作

Private Const conMenu_Oper_PrintLabel = 301
Private Const conMenu_Oper_PrintLabel_SelRow = 302    '打印标签（当前选中行）
Private Const conMenu_Oper_PrintLabel_SelBatch = 303    '打印标签（当前选中批次）
Private Const conMenu_Oper_PrintLabel_SelDept = 304    '打印标签（当前选中病区）
Private Const conMenu_Oper_PrintLabel_SelPati = 305    '打印标签（当前选中病人）
Private Const conMenu_Oper_PrintLabel_AllRow = 306    '打印标签（所有选择的行）
Private Const conMenu_Oper_PrintLabel_SelSendNo = 307    '打印标签（当前选中的摆药单号）

Private Const conMenu_Oper_DelBatch = 311
Private Const conMenu_Oper_DelBatch_SelRow = 312       '删除批次（当前选中行）
Private Const conMenu_Oper_DelBatch_SelBatch = 313   '删除批次（当前选中批次）
Private Const conMenu_Oper_DelBatch_SelDept = 314    '删除批次（当前选中病区）
Private Const conMenu_Oper_DelBatch_SelPati = 315    '删除批次（当前选中病人）
Private Const conMenu_Oper_DelBatch_AllRow = 316    '删除批次（所有选择的行）

Private Const conMenu_Oper_Select = 321
Private Const conMenu_Oper_Select_SelRow = 322       '选择（当前选中行）
Private Const conMenu_Oper_Select_SelBatch = 323   '选择（当前选中批次）
Private Const conMenu_Oper_Select_SelDept = 324    '选择（当前选中病区）
Private Const conMenu_Oper_Select_SelPati = 325    '选择（当前选中病人）
Private Const conMenu_Oper_Select_SelAll = 326    '选择（所有行）
Private Const conMenu_Oper_Select_SelSendNo = 327    '选择（当前选中的摆药单号）
Private Const conMenu_Oper_Select_SelMed = 328    '选择（当前所有的抗菌药物）
Private Const conMenu_Oper_Select_CancleSelDept = 329    '取消选择（当前选中病区）
Private Const conMenu_Oper_Select_CancleSelPati = 330    '取消选择（当前选中病人）

Private Const conMenu_Oper_Bag = 331
Private Const conMenu_Oper_Bag_Batch = 332   '打包（打包当前批次）
Private Const conMenu_Oper_Bag_All = 333   '全部打包（打包当前批次）

Private Const conMenu_Oper_Look = 341        '电子病案查阅

Private Const mconMenu_SortPopup = 6000                  '排序方式
Private Const mconMenu_SortPopup_ByCode = 6001           '按编码
Private Const mconMenu_SortPopup_ByName = 6002           '按名称
          
'医保接口
Private gclsInsure As New clsInsure

Private Type TYPE_MedicarePAR
    负数记帐 As Boolean
    记帐上传 As Boolean
    记帐完成后上传 As Boolean
    记帐作废上传 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR
Private Function CheckPriceAdjustByID() As Boolean
    '根据收发ID检查与药品零差价
    Dim rstemp As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    Dim strDrugList As String
    Dim i As Integer
    
    On Error GoTo errHandle
    
    '如果没开启全局的零差价管理，则不进行后续检查，返回true
    If Val(zlDatabase.GetPara(275, 100, , 0)) = 0 Then CheckPriceAdjustByID = True: Exit Function
    
    If mrsTrans Is Nothing Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "执行标志=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Sort = "收发ID"
    
    Set rstemp = mrsTrans
    If mrsTrans.RecordCount = 0 Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rstemp.EOF
        If i >= 5 Then Exit Do
        gstrSQL = "Select a.药品id, Nvl(a.批次, 0) As 批次," & vbNewLine & _
            "       '[' || c.编码 || ']' || c.名称 || Decode(c.产地, Null, Null, '(' || c.产地 || ')') || c.规格 As 通用名" & vbNewLine & _
            " From 药品收发记录 A, 药品规格 B, 收费项目目录 C" & vbNewLine & _
            " Where a.药品id = b.药品id And b.药品id = c.Id And b.是否零差价管理 = 1 And a.Id = [1] "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckPriceAdjustByID", Val(rstemp!收发ID))
        
        If Not rsData.EOF Then
            If InStr(1, "," & strDrugList & ",", "," & rsData!药品ID & ",") = 0 Then
                strDrugList = IIf(strDrugList = "", "", strDrugList & ",") & rsData!药品ID
                If CheckPriceAdjust(rsData!药品ID, mcondition.lngCenterID, rsData!批次) = False Then
                    i = i + 1
                    strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & rsData!通用名
                End If
            End If
        End If
        
        rstemp.MoveNext
    Loop
    
    If strMsg = "" Then
        CheckPriceAdjustByID = True
        Exit Function
    Else
        MsgBox "以下药品启用了零差价管理，但库存中售价和成本价不一致，不能进行业务，请检查！" & vbCrLf & strMsg, vbInformation, gstrSysName
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitVSFLook()
    Dim arr列设置 As Variant
    Dim n As Integer
    Dim i As Integer
    
    mIntCol操作状态 = 0
    mIntCol瓶签号 = 1
    mintcol批次 = 2
    mIntCol打包 = 3
    mIntCol摆药人 = 4
    mIntCol摆药时间 = 5
    mIntCol摆药单号 = 6
    mIntCol医嘱发送时间 = 7
    mIntCol执行时间 = 8
    mIntCol药品名称 = 9
    mintcol规格 = 10
    mIntCol单量 = 11
    mintcol数量 = 12
    mIntColNO = 13
    mIntCol单据 = 14
    mIntCol剂量单位 = 15
    mIntCol用法 = 16
    mintcol药品id = 17
    mIntCol配药id = 18
    
    '恢复用户自定义列顺序
    If mParams.strVsfLook <> "" Then
        arr列设置 = Split(mParams.strVsfLook, "|")
        
        For n = 0 To UBound(arr列设置)
            SetVsfLookValue Split(arr列设置(n), ",")(0), n
        Next
    End If
    
    With VSFLook
        .rows = 1
        .rows = 2
'        .Cols = 17
        
        VsfGridColFormat VSFLook, mIntCol操作状态, "操作状态", 1200, flexAlignLeftCenter, "操作状态"
        VsfGridColFormat VSFLook, mIntCol瓶签号, "瓶签号", 2000, flexAlignRightCenter, "瓶签号"
        VsfGridColFormat VSFLook, mintcol批次, "批次", 1000, flexAlignLeftCenter, "批次"
        VsfGridColFormat VSFLook, mIntCol打包, "打包", 1000, flexAlignLeftCenter, "打包"
        VsfGridColFormat VSFLook, mIntCol摆药人, "摆药人", 1200, flexAlignLeftCenter, "摆药人"
        VsfGridColFormat VSFLook, mIntCol摆药时间, "摆药时间", 2000, flexAlignLeftCenter, "摆药时间"
        VsfGridColFormat VSFLook, mIntCol摆药单号, "摆药单号", 2000, flexAlignLeftCenter, "摆药单号"
        VsfGridColFormat VSFLook, mIntCol医嘱发送时间, "医嘱发送时间", 2000, flexAlignLeftCenter, "医嘱发送时间"
        VsfGridColFormat VSFLook, mIntCol执行时间, "执行时间", 1800, flexAlignLeftCenter, "执行时间"
        VsfGridColFormat VSFLook, mIntCol药品名称, "药品名称", 1800, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat VSFLook, mintcol规格, "规格", 1800, flexAlignLeftCenter, "规格"
        VsfGridColFormat VSFLook, mIntCol单量, "单量", 1800, flexAlignLeftCenter, "单量"
        VsfGridColFormat VSFLook, mintcol数量, "数量", 1800, flexAlignLeftCenter, "数量"
        VsfGridColFormat VSFLook, mIntColNO, "NO", 1800, flexAlignLeftCenter, "NO"
        VsfGridColFormat VSFLook, mIntCol单据, "单据", 1800, flexAlignLeftCenter, "单据"
        VsfGridColFormat VSFLook, mIntCol剂量单位, "剂量单位", 1800, flexAlignLeftCenter, "剂量单位"
        VsfGridColFormat VSFLook, mIntCol用法, "用法", 1800, flexAlignLeftCenter, "用法"
        VsfGridColFormat VSFLook, mintcol药品id, "药品id", 1800, flexAlignLeftCenter, "药品id"
        VsfGridColFormat VSFLook, mIntCol配药id, "配药id", 1800, flexAlignLeftCenter, "配药id"
    End With
    
    '恢复列宽
    If mParams.strVsfLook <> "" Then
        arr列设置 = Split(mParams.strVsfLook, "|")
        For n = 0 To UBound(arr列设置)
            For i = 0 To VSFLook.Cols - 1
                If Split(arr列设置(n), ",")(0) = VSFLook.ColKey(i) Then
                    VSFLook.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                End If
            Next
        Next
    End If
End Sub

Private Sub SetVsfLookValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "瓶签号"
            mIntCol瓶签号 = intValue
        Case "批次"
            mintcol批次 = intValue
        Case "打包"
            mIntCol打包 = intValue
        Case "摆药人"
            mIntCol摆药人 = intValue
        Case "摆药时间"
            mIntCol摆药时间 = intValue
        Case "摆药单号"
            mIntCol摆药单号 = intValue
        Case "医嘱发送时间"
            mIntCol医嘱发送时间 = intValue
        Case "执行时间"
            mIntCol执行时间 = intValue
        Case "药品名称"
            mIntCol药品名称 = intValue
        Case "数量"
            mintcol数量 = intValue
        Case "规格"
            mintcol规格 = intValue
        Case "单量"
            mIntCol单量 = intValue
        Case "NO"
            mIntColNO = intValue
        Case "单据"
            mIntCol单据 = intValue
        Case "剂量单位"
            mIntCol剂量单位 = intValue
        Case "用法"
            mIntCol用法 = intValue
        Case "用法"
            mintcol药品id = intValue
        Case "配药id"
            mIntCol配药id = intValue
    End Select
End Sub
Private Sub InitVsfSum()
    Dim arr列设置 As Variant
    Dim n As Integer
    Dim i As Integer
    
    mintcolsum病区 = 0
    mintcolsum打包 = 1
    mintcolsum药品名称 = 2
    mintcolsum商品名 = 3
    mintcolsum英文名 = 4
    mintcolsum规格 = 5
    mintcolsum产地 = 6
    mintcolsum批号 = 7
    mintcolsum数量 = 8
    mintcolsum发药数量 = 9
    mintcolsum库存数量 = 10
    mintcolsum缺药标志 = 11
    mintcolsum是否打包 = 12
    
    '恢复用户自定义列顺序
    If mParams.strVsfSum <> "" Then
        arr列设置 = Split(mParams.strVsfSum, "|")
        
        For n = 0 To UBound(arr列设置)
            SetColumnValue Split(arr列设置(n), ",")(0), n
        Next
    End If
    
    With vsfSumDrug
        .rows = 1
        .rows = 2
        .Cols = MINTSUMCOLS

        VsfGridColFormat vsfSumDrug, mintcolsum病区, "病区", 450, flexAlignRightCenter, "病区"
        VsfGridColFormat vsfSumDrug, mintcolsum打包, "打包", 2500, flexAlignLeftCenter, "打包"
        VsfGridColFormat vsfSumDrug, mintcolsum药品名称, "药品名称", 400, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfSumDrug, mintcolsum商品名, "商品名", 2000, flexAlignLeftCenter, "商品名"
        VsfGridColFormat vsfSumDrug, mintcolsum英文名, "英文名", 2000, flexAlignLeftCenter, "英文名"
        VsfGridColFormat vsfSumDrug, mintcolsum规格, "规格", 1800, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfSumDrug, mintcolsum产地, "产地", 1800, flexAlignLeftCenter, "产地"
        VsfGridColFormat vsfSumDrug, mintcolsum批号, "批号", 1800, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfSumDrug, mintcolsum数量, "数量", 1800, flexAlignLeftCenter, "数量"
        VsfGridColFormat vsfSumDrug, mintcolsum发药数量, "发药数量", 1800, flexAlignLeftCenter, "发药数量"
        VsfGridColFormat vsfSumDrug, mintcolsum库存数量, "库存数量", 1800, flexAlignLeftCenter, "库存数量"
        VsfGridColFormat vsfSumDrug, mintcolsum缺药标志, "缺药标志", 1800, flexAlignLeftCenter, "缺药标志"
        VsfGridColFormat vsfSumDrug, mintcolsum是否打包, "是否打包", 1800, flexAlignLeftCenter, "是否打包"
    End With
    
    '恢复列宽
    If mParams.strVsfSum <> "" Then
        arr列设置 = Split(mParams.strVsfSum, "|")
        For n = 0 To UBound(arr列设置)
            For i = 0 To vsfSumDrug.Cols - 1
                If Split(arr列设置(n), ",")(0) = vsfSumDrug.ColKey(i) Then
                    vsfSumDrug.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                End If
            Next
        Next
    End If
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "病区"
            mintcolsum病区 = intValue
        Case "打包"
            mintcolsum打包 = intValue
        Case "药品名称"
            mintcolsum药品名称 = intValue
        Case "英文名"
            mintcolsum英文名 = intValue
        Case "商品名"
            mintcolsum商品名 = intValue
        Case "规格"
            mintcolsum规格 = intValue
        Case "产地"
            mintcolsum产地 = intValue
        Case "批号"
            mintcolsum批号 = intValue
        Case "数量"
            mintcolsum数量 = intValue
        Case "发药数量"
            mintcolsum发药数量 = intValue
        Case "库存数量"
            mintcolsum库存数量 = intValue
        Case "缺药标志"
            mintcolsum缺药标志 = intValue
        Case "是否打包"
            mintcolsum是否打包 = intValue
    End Select
                   
End Sub

Private Sub SetTransColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "当前行"
            mIntCol当前行 = intValue
        Case "审"
            mIntCol审 = intValue
        Case "选择"
            mintcol选择 = intValue
        Case "变"
            mIntCol变 = intValue
        Case "锁"
            mIntCol锁 = intValue
        Case "医嘱"
            mIntCol医嘱 = intValue
        Case "调"
            mIntCol调 = intValue
        Case "打印"
            mIntCol打印 = intValue
        Case "打包"
            mIntCol打包 = intValue
        Case "配药批次"
            mintcol批次 = intValue
        Case "拒收原因"
            mIntCol拒收原因 = intValue
        Case "优先级"
            mIntCol优先级 = intValue
        Case "病区"
            mIntCol病区 = intValue
        Case "科室"
            mIntCol科室 = intValue
        Case "姓名"
            mIntCol姓名 = intValue
        Case "性别"
            mIntCol性别 = intValue
        Case "年龄"
            mIntCol年龄 = intValue
        Case "床号"
            mIntCol床号 = intValue
        Case "住院号"
            mIntCol住院号 = intValue
        Case "审查结果"
            mIntCol警 = intValue
        Case "药品名称"
            mIntCol药品名称 = intValue
        Case "皮"
            mIntCol皮 = intValue
        Case "规格"
            mintcol规格 = intValue
        Case "配药类型"
            mIntCol配药类型 = intValue
        Case "单量"
            mIntCol单量 = intValue
        Case "数量"
            mintcol数量 = intValue
        Case "执行时间"
            mIntCol执行时间 = intValue
        Case "瓶签号"
            mIntCol瓶签号 = intValue
        Case "摆药单号"
            mIntCol摆药单号 = intValue
        Case "医嘱发送时间"
            mIntCol医嘱发送时间 = intValue
        Case "摆药人"
            mIntCol摆药人 = intValue
        Case "摆药时间"
            mIntCol摆药时间 = intValue
        Case "配药人"
            mIntCol配药人 = intValue
        Case "配药时间"
            mIntCol配药时间 = intValue
        Case "发送人"
            mIntCol发送人 = intValue
        Case "发送时间"
            mIntCol发送时间 = intValue
        Case "销帐申请人"
            mIntCol销帐申请人 = intValue
        Case "销帐申请时间"
            mIntCol销帐申请时间 = intValue
        Case "销帐审核人"
            mIntCol销帐审核人 = intValue
        Case "销帐审核时间"
            mIntCol销帐审核时间 = intValue
        Case "标志"
            mIntCol标志 = intValue
        Case "作废类型"
            mIntCol作废类型 = intValue
        Case "NO"
            mIntColNO = intValue
        Case "剂量单位"
            mIntCol剂量单位 = intValue
        Case "用法"
            mIntCol用法 = intValue
        Case "药品id"
            mintcol药品id = intValue
        Case "核查人"
            mIntCol核查人 = intValue
        Case "核查时间"
            mIntCol核查时间 = intValue
        Case "打印标志"
            mIntCol打印标志 = intValue
        Case "配药id"
            mIntCol配药id = intValue
        Case "抗菌药物"
            mIntCol抗菌药物 = intValue
        Case "原批次"
            mIntCol原批次 = intValue
        Case "主页id"
            mIntCol主页id = intValue
        Case "病人ID"
            mIntCol病人ID = intValue
        Case "背景号"
            mIntCol背景号 = intValue
        Case "是否锁定"
            mIntCol是否锁定 = intValue
        Case "单据"
            mIntCol单据 = intValue
        Case "是否打包"
            mIntCol是否打包 = intValue
        Case "销帐原因"
            mIntCol销帐原因 = intValue
    End Select
End Sub

Private Sub InitVsfTrans()
    Dim arr列设置 As Variant
    Dim n As Integer
    Dim i As Integer
    Dim strRows As String
    
    '初始化列
    mIntCol当前行 = 0
    mIntCol审 = 1
    mintcol选择 = 2
    mIntCol作废类型 = 3
    mIntCol变 = 4
    mIntCol锁 = 5
    mIntCol医嘱 = 6
    mIntCol调 = 7
    mIntCol打印 = 8
    mIntCol打包 = 9
    mintcol批次 = 10
    mIntCol拒收原因 = 11
    mIntCol优先级 = 12
    mIntCol病区 = 13
    mIntCol科室 = 14
    mIntCol姓名 = 15
    mIntCol性别 = 16
    mIntCol年龄 = 17
    mIntCol床号 = 18
    mIntCol住院号 = 19
    mIntCol警 = 20
    mIntCol药品名称 = 21
    mIntCol皮 = 22
    mintcol规格 = 23
    mIntCol配药类型 = 24
    mIntCol单量 = 25
    mintcol数量 = 26
    mIntCol执行时间 = 27
    mIntCol执行频次 = 28
    mIntCol瓶签号 = 29
    mIntCol摆药单号 = 30
    mIntCol医嘱发送时间 = 31
    mIntCol摆药人 = 32
    mIntCol摆药时间 = 33
    mIntCol配药人 = 34
    mIntCol配药时间 = 35
    mIntCol发送人 = 36
    mIntCol发送时间 = 37
    mIntCol销帐申请人 = 38
    mIntCol销帐申请时间 = 39
    mIntCol销帐审核人 = 40
    mIntCol销帐审核时间 = 41
    mIntCol销帐原因 = 42
    mIntCol标志 = 43
    mIntCol是否锁定 = 44
    mIntColNO = 45
    mIntCol单据 = 46
    mIntCol剂量单位 = 47
    mIntCol用法 = 48
    mintcol药品id = 49
    mIntCol核查人 = 50
    mIntCol核查时间 = 51
    mIntCol打印标志 = 52
    mIntCol配药id = 53
    mIntCol是否打包 = 54
    mIntCol原批次 = 55
    mIntCol抗菌药物 = 56
    mIntCol主页id = 57
    mIntCol病人ID = 58
    mIntCol背景号 = 59
    mIntCol警告 = 60
    mIntCol溶媒 = 61
    mIntCol对应医嘱ID = 62
    
    '若用户以前没有保存"配药类型"列,则进行初始化
    If InStr(mParams.strVsfTrans, "配药类型") = 0 Then mParams.strVsfTrans = ""
    
    '恢复用户自定义列顺序
    If mParams.strVsfTrans <> "" Then
        arr列设置 = Split(mParams.strVsfTrans, "|")
        
        For n = 0 To UBound(arr列设置)
            SetTransColumnValue Split(arr列设置(n), ",")(0), n
        Next
    End If
    
    With vsfTrans
        .Cols = MINTCOLS
        
        VsfGridColFormat vsfTrans, mIntCol当前行, " ", 200, flexAlignRightCenter, "当前行"
        VsfGridColFormat vsfTrans, mIntCol审, "审", 400, flexAlignRightCenter, "审"
        VsfGridColFormat vsfTrans, mintcol选择, "选择", 400, flexAlignLeftCenter, "选择"
        VsfGridColFormat vsfTrans, mIntCol变, "变", 400, flexAlignLeftCenter, "变"
        VsfGridColFormat vsfTrans, mIntCol锁, "锁", 400, flexAlignLeftCenter, "锁"
        VsfGridColFormat vsfTrans, mIntCol医嘱, "医嘱", 400, flexAlignLeftCenter, "医嘱"
        VsfGridColFormat vsfTrans, mIntCol调, "调", 400, flexAlignLeftCenter, "调"
        VsfGridColFormat vsfTrans, mIntCol打印, "打印", 400, flexAlignLeftCenter, "打印"
        VsfGridColFormat vsfTrans, mIntCol打包, "打包", 400, flexAlignLeftCenter, "打包"
        VsfGridColFormat vsfTrans, mintcol批次, "批次", 400, flexAlignLeftCenter, "配药批次"
        VsfGridColFormat vsfTrans, mIntCol拒收原因, "拒收原因", 0, flexAlignLeftCenter, "拒收原因"
        VsfGridColFormat vsfTrans, mIntCol优先级, "优先级", 600, flexAlignLeftCenter, "优先级"
        VsfGridColFormat vsfTrans, mIntCol病区, "病区", 1800, flexAlignLeftCenter, "病区"

        VsfGridColFormat vsfTrans, mIntCol科室, "科室", 1800, flexAlignRightCenter, "科室"
        VsfGridColFormat vsfTrans, mIntCol姓名, "姓名", 1800, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfTrans, mIntCol性别, "性别", 400, flexAlignLeftCenter, "性别"
        VsfGridColFormat vsfTrans, mIntCol年龄, "年龄", 800, flexAlignLeftCenter, "年龄"
        VsfGridColFormat vsfTrans, mIntCol床号, "床号", 800, flexAlignLeftCenter, "床号"
        VsfGridColFormat vsfTrans, mIntCol住院号, "住院号", 800, flexAlignLeftCenter, "住院号"
        VsfGridColFormat vsfTrans, mIntCol警, "警", 400, flexAlignLeftCenter, "审查结果"
        VsfGridColFormat vsfTrans, mIntCol药品名称, "药品名称", 1800, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfTrans, mIntCol皮, "皮", 600, flexAlignLeftCenter, "皮"
        VsfGridColFormat vsfTrans, mintcol规格, "规格", 1800, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfTrans, mIntCol配药类型, "配药类型", 1800, flexAlignLeftCenter, "配药类型"
        VsfGridColFormat vsfTrans, mIntCol单量, "单量", 1800, flexAlignLeftCenter, "单量"
        VsfGridColFormat vsfTrans, mintcol数量, "数量", 1800, flexAlignLeftCenter, "数量"


        VsfGridColFormat vsfTrans, mIntCol执行时间, "执行时间", 2000, flexAlignRightCenter, "执行时间"
        VsfGridColFormat vsfTrans, mIntCol执行频次, "执行频次", 1200, flexAlignRightCenter, "执行频次"
        VsfGridColFormat vsfTrans, mIntCol瓶签号, "瓶签号", 2500, flexAlignLeftCenter, "瓶签号"
        VsfGridColFormat vsfTrans, mIntCol摆药单号, "摆药单号", 2000, flexAlignLeftCenter, "摆药单号"
        VsfGridColFormat vsfTrans, mIntCol医嘱发送时间, "医嘱发送时间", 2000, flexAlignLeftCenter, "医嘱发送时间"
        VsfGridColFormat vsfTrans, mIntCol摆药人, "摆药人", 2000, flexAlignLeftCenter, "摆药人"
        VsfGridColFormat vsfTrans, mIntCol摆药时间, "摆药时间", 1800, flexAlignLeftCenter, "摆药时间"
        VsfGridColFormat vsfTrans, mIntCol配药人, "配药人", 1800, flexAlignLeftCenter, "配药人"
        VsfGridColFormat vsfTrans, mIntCol配药时间, "配药时间", 1800, flexAlignLeftCenter, "配药时间"
        VsfGridColFormat vsfTrans, mIntCol发送人, "发送人", 1800, flexAlignLeftCenter, "发送人"
        VsfGridColFormat vsfTrans, mIntCol发送时间, "发送时间", 1800, flexAlignLeftCenter, "发送时间"
        VsfGridColFormat vsfTrans, mIntCol销帐申请人, "销帐申请人", 1800, flexAlignLeftCenter, "销帐申请人"
        VsfGridColFormat vsfTrans, mIntCol销帐申请时间, "销帐申请时间", 1800, flexAlignLeftCenter, "销帐申请时间"

        VsfGridColFormat vsfTrans, mIntCol销帐审核人, "销帐审核人", 450, flexAlignRightCenter, "销帐审核人"
        VsfGridColFormat vsfTrans, mIntCol销帐审核时间, "销帐审核时间", 2500, flexAlignLeftCenter, "销帐审核时间"
        VsfGridColFormat vsfTrans, mIntCol销帐原因, "销帐原因", 2500, flexAlignLeftCenter, "销帐原因"
        VsfGridColFormat vsfTrans, mIntCol标志, "标志", 400, flexAlignLeftCenter, "标志"
        VsfGridColFormat vsfTrans, mIntCol是否锁定, "是否锁定", 0, flexAlignLeftCenter, "是否锁定"
        VsfGridColFormat vsfTrans, mIntCol作废类型, "作废类型", 2000, flexAlignLeftCenter, "作废类型"
        VsfGridColFormat vsfTrans, mIntColNO, "NO", 1800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfTrans, mIntCol单据, "单据", 1800, flexAlignLeftCenter, "单据"
        VsfGridColFormat vsfTrans, mIntCol剂量单位, "剂量单位", 1800, flexAlignLeftCenter, "剂量单位"
        VsfGridColFormat vsfTrans, mIntCol用法, "用法", 1800, flexAlignLeftCenter, "用法"
        VsfGridColFormat vsfTrans, mintcol药品id, "药品id", 1800, flexAlignLeftCenter, "药品id"
        VsfGridColFormat vsfTrans, mIntCol核查人, "核查人", 1800, flexAlignLeftCenter, "核查人"
        VsfGridColFormat vsfTrans, mIntCol核查时间, "核查时间", 1800, flexAlignLeftCenter, "核查时间"
        VsfGridColFormat vsfTrans, mIntCol打印标志, "打印标志", 0, flexAlignLeftCenter, "打印标志"

        VsfGridColFormat vsfTrans, mIntCol配药id, "配药id", 0, flexAlignLeftCenter, "配药id"
        VsfGridColFormat vsfTrans, mIntCol是否打包, "是否打包", 0, flexAlignLeftCenter, "是否打包"
        VsfGridColFormat vsfTrans, mIntCol原批次, "原批次", 0, flexAlignLeftCenter, "原批次"
        VsfGridColFormat vsfTrans, mIntCol抗菌药物, "抗菌药物", 0, flexAlignLeftCenter, "抗菌药物"
        VsfGridColFormat vsfTrans, mIntCol主页id, "主页id", 0, flexAlignLeftCenter, "主页id"
        VsfGridColFormat vsfTrans, mIntCol病人ID, "病人ID", 0, flexAlignLeftCenter, "病人ID"
        VsfGridColFormat vsfTrans, mIntCol背景号, "背景号", 0, flexAlignLeftCenter, "背景号"
        VsfGridColFormat vsfTrans, mIntCol警告, "警告", 0, flexAlignLeftCenter, "警告"
        VsfGridColFormat vsfTrans, mIntCol溶媒, "溶媒", 0, flexAlignLeftCenter, "溶媒"
        VsfGridColFormat vsfTrans, mIntCol对应医嘱ID, "对应医嘱ID", 0, flexAlignLeftCenter, "对应医嘱ID"
    End With

    '恢复个性设置
    If mParams.strVsfTrans <> "" Then
        arr列设置 = Split(mParams.strVsfTrans, "|")
        For n = 0 To UBound(arr列设置)
            For i = 0 To vsfTrans.Cols - 1
                If Split(arr列设置(n), ",")(0) = vsfTrans.ColKey(i) Then
                    vsfTrans.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                End If
            Next
        Next
    End If
    
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        strRows = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\表格", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = mstrUnVisble & "审;摆药单号;作废类型;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;"
    End If
    
    
    If strRows <> "" Then
        For n = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(n) & ";") > 0 Then
                vsfTrans.ColHidden(n) = True
            Else
                vsfTrans.ColHidden(n) = False
            End If
        Next
    End If
    
    '初始环节为摆药环节
    Call InitColSelList(mstrUnVisble & "审;摆药单号;作废类型;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;")
End Sub




Private Sub SetSortFlag(Optional ByVal blnSpecial As Boolean = False)
    '设置排序标志
    Dim intCol, intSortCount As Integer
    
    With vsfTrans
        .Redraw = flexRDNone
        
        '取消排序标志
        For intCol = 0 To .Cols - 1
            If InStr(1, .TextMatrix(0, intCol), "①") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "①", "")
            If InStr(1, .TextMatrix(0, intCol), "②") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "②", "")
            If InStr(1, .TextMatrix(0, intCol), "③") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "③", "")
            If InStr(1, .TextMatrix(0, intCol), "④") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "④", "")
            If InStr(1, .TextMatrix(0, intCol), "⑤") > 0 Then .TextMatrix(0, intCol) = Replace(.TextMatrix(0, intCol), "⑤", "")
        Next
        
        '设置排序标志
        If blnSpecial = True Then
            '特殊的按药品排序(过滤药品，溶媒，过滤药品单量)
            .TextMatrix(0, .ColIndex("药品名称")) = .TextMatrix(0, .ColIndex("药品名称")) & "①"
        ElseIf mParams.strSort <> "" Then
            For intCol = 0 To .Cols - 1
                For intSortCount = 0 To UBound(Split(mParams.strSort, ","))
                    If .ColKey(intCol) = IIf(Split(mParams.strSort, ",")(intSortCount) = "排序床号", "床号", Split(mParams.strSort, ",")(intSortCount)) Then
                        Select Case intSortCount + 1
                            Case 1
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "①"
                            Case 2
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "②"
                            Case 3
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "③"
                            Case 4
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "④"
                            Case 5
                                .TextMatrix(0, intCol) = .TextMatrix(0, intCol) & "⑤"
                        End Select
                        
                        Exit For
                    End If
                Next
            Next
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetCount()
    '计算统计信息：所选病区数，所选输液单数
    '在状态栏显示
    Dim lngCount As Long
    Dim lngRow As Long
    Dim lng相关ID As Long
    Dim lngVolume As Long
    
    stbThis.Panels(2).Text = ""
    
    With vsfDept(Me.tabDeptList.Selected.index)
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("病区ID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    lngCount = lngCount + 1
                End If
            End If
        Next
    End With
    
    If lngCount = 0 Then Exit Sub
    stbThis.Panels(2).Text = "当前选择病区：" & lngCount
    
    lngCount = 0
    lngVolume = 0
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("溶媒"))) = 1 Then
                        lngVolume = lngVolume + Val(.TextMatrix(lngRow, .ColIndex("单量")))
                    End If
                    
                    If lng相关ID <> Val(.TextMatrix(lngRow, .ColIndex("配药id"))) Then
                        lng相关ID = Val(.TextMatrix(lngRow, .ColIndex("配药id")))
                        lngCount = lngCount + 1
                    End If
                End If
            End If
        Next
    End With
    
    If Not mrsTrans Is Nothing Then
        If mrsTrans.RecordCount > 0 Then
            mrsTrans.Filter = ""
            mrsTrans.Sort = "组号"
            mrsTrans.MoveLast
            lblCount.Caption = "输液单：" & mlng已扫描 + mlng未扫描 & " 已：" & mlng已扫描 & "  未：" & mlng未扫描 & " 当前选择输液单：" & lngCount
            mrsTrans.MoveFirst
        End If
    End If
    
    lblVolu.Visible = False
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        lblVolu.Visible = True
        Me.lblVolu.Caption = "容量：" & lngVolume
    End If
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "  当前选择输液单：" & lngCount & IIf(mcondition.strTransStep = M_STR_CALSS_DOSAGE, " 当前改变了打包状态的输液单：" & mintCountPack, "")
    
End Sub

Private Function Check自备药() As Boolean
    '功能：检查自备药的待发药数量
    Dim strSQL As String
    Dim rs汇总 As ADODB.Recordset
    Dim rs实际 As ADODB.Recordset
    Dim rsData As ADODB.Recordset
    Dim i As Integer
    Dim str病人ids As String     '例如：病人1,病人2...
    Dim lng病人id As Long
    Dim str配药ids As String     '例如：配药id1,配药id2...
    Dim str当前病人 As String
    Dim str当前药品 As String
    Dim lng当前配药id As Long
    
    On Error GoTo errHandle
    
    Check自备药 = False
    
    If Not mblnShowOhters Then Exit Function
    
    If mrsTrans Is Nothing Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "执行标志=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Set rsData = mrsTrans
    
    '对输液单进行按病人分组
    With rsData
        
        .Sort = "病人id"
        
        Do While Not .EOF
            '收集病人ID
            If InStr("," & str病人ids & ",", "," & !病人ID & ",") = 0 Then
                str病人ids = str病人ids & IIf(str病人ids = "", "", ",") & !病人ID
            End If
            
            .MoveNext
        Loop
        
        '集中查询该病人的是否存在自备药记录
        For i = 0 To UBound(Split(str病人ids, ","))
            lng病人id = Split(str病人ids, ",")(i)
            
            .Filter = "执行标志=1 and 病人id =" & lng病人id
            .Sort = "配药id"
            
            '不同病人需要初始化
            str配药ids = ""
            
            '收集配药id
            Do While Not .EOF
                If InStr("," & str配药ids & ",", "," & !配药id & ",") = 0 Then
                    str配药ids = str配药ids & IIf(str配药ids = "", "", ",") & !配药id
                End If

                .MoveNext
            Loop
            
            '【1.先进行药品汇总检查】
            strSQL = "Select c.药品id, Sum((b.单次用量 / c.剂量系数)) As 数量, a.姓名, e.名称" & vbNewLine & _
                    "From 输液配药记录 A, 病人医嘱记录 B, 药品规格 C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D, 收费项目目录 E" & vbNewLine & _
                    "Where a.医嘱id = b.相关id And a.Id = d.Column_Value And b.收费细目id = c.药品id And c.药品id = e.Id And b.执行性质 = 5 And b.执行标记 = 0 And" & vbNewLine & _
                    "      b.收费细目id In (Select d.药品id From 输液自备药清单 D Where d.是否检查库存 = 1)" & vbNewLine & _
                    "Group By c.药品id, a.姓名, e.名称"
                    
            Set rs汇总 = zlDatabase.OpenSQLRecord(strSQL, "检查自备药", str配药ids)
                        
            '检查对应药品数量是否足够
            Do While Not rs汇总.EOF
                str当前病人 = rs汇总!姓名
                str当前药品 = rs汇总!名称
                
                strSQL = "Select Sum(b.实际数量) As 实际数量, c.名称" & vbNewLine & _
                        "From 未发药品记录 A, 药品收发记录 B, 收费项目目录 C" & vbNewLine & _
                        "Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And b.药品id = c.Id And b.审核人 Is Null And b.审核日期 Is Null And" & vbNewLine & _
                        "      Mod(b.记录状态, 3) = 1 And a.病人id = [1] And a.库房id = [2] And b.药品id = [3] And Exists (Select 1 From 门诊费用记录 C Where c.Id = b.费用id)" & vbNewLine & _
                        "Group By 名称"
                
                Set rs实际 = zlDatabase.OpenSQLRecord(strSQL, "核对自备药数量", lng病人id, mParams.lng配置中心, rs汇总!药品ID)
                
                If rs实际.EOF Then
                    MsgBox "病人 " & str当前病人 & " 的自备药药品【" & str当前药品 & "】的待发药数量不足，无法进行摆药！", vbExclamation, "自备药待发数量检查"
                    Exit Function
                Else
                    '若数量不够，弹出提示并终止检查
                    If nvl(rs实际!实际数量, 0) < nvl(rs汇总!数量, 0) Then
                        MsgBox "病人 " & str当前病人 & " 的自备药药品【" & str当前药品 & "】的待发药数量不足，无法进行摆药！", vbExclamation, "自备药待发数量检查"
                        Exit Function
                    End If
                End If
                
                rs汇总.MoveNext
            Loop
            
            '【2.再进行退药待发的药品汇总检查】
            strSQL = "Select c.药品id, Sum((b.单次用量 / c.剂量系数)) As 数量, a.姓名, e.名称, a.Id" & vbNewLine & _
                    "From 输液配药记录 A, 病人医嘱记录 B, 药品规格 C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D, 收费项目目录 E" & vbNewLine & _
                    "Where a.医嘱id = b.相关id And a.Id = d.Column_Value And b.收费细目id = c.药品id And c.药品id = e.Id And b.执行性质 = 5 And b.执行标记 = 0 And" & vbNewLine & _
                    "      b.收费细目id In (Select d.药品id From 输液自备药清单 D Where d.是否检查库存 = 1)" & vbNewLine & _
                    "Group By c.药品id, a.姓名, e.名称, a.Id"
                    
            Set rs汇总 = zlDatabase.OpenSQLRecord(strSQL, "检查自备药", str配药ids)
            
            '检查退药待发对应药品数量是否足够
            Do While Not rs汇总.EOF
                str当前病人 = rs汇总!姓名
                str当前药品 = rs汇总!名称
                lng当前配药id = rs汇总!Id
                
                strSQL = "Select Sum(b.实际数量) As 实际数量, c.名称" & vbNewLine & _
                        "From 未发药品记录 A, 药品收发记录 B, 收费项目目录 C" & vbNewLine & _
                        "Where a.单据 = b.单据 And a.No = b.No And a.库房id = b.库房id And b.药品id = c.Id And b.审核人 Is Null And b.审核日期 Is Null And b.计划id = [4] And" & vbNewLine & _
                        "      Mod(b.记录状态, 3) = 1 And a.病人id = [1] And a.库房id = [2] And b.药品id = [3]" & vbNewLine & _
                        "Group By 名称"
                
                Set rs实际 = zlDatabase.OpenSQLRecord(strSQL, "核对自备药数量", lng病人id, mParams.lng配置中心, rs汇总!药品ID, lng当前配药id)
                
                If Not rs实际.EOF Then
                    '若数量不够，弹出提示并终止检查
                    If nvl(rs实际!实际数量, 0) < nvl(rs汇总!数量, 0) Then
                        MsgBox "病人 " & str当前病人 & " 的自备药药品【" & str当前药品 & "】的待发药数量不足，无法进行摆药！", vbExclamation, "自备药待发数量检查"
                        Exit Function
                    End If
                End If
                
                rs汇总.MoveNext
            Loop
        Next
    End With

    Check自备药 = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InputIsScaner(ByRef txtInput As Object, ByVal KeyAscii As Integer) As Boolean
'功能：判断指定文本框中当前输入是否是由条码设备读入
'参数：KeyAscii=在KeyPress事件中调用的参数
    Static sngInputBegin As Single
    Dim sngNow As Single, blnScaner As Boolean, strText As String
    
    '处理当前键入后显示的内容(还未显示出来)
    strText = txtInput.Text
    If txtInput.SelLength = Len(txtInput.Text) Then strText = ""
    If KeyAscii = 8 Or KeyAscii = 9 Or KeyAscii = 10 Then
        If strText <> "" Then strText = Mid(strText, 1, Len(strText) - 1)
    Else
        strText = strText & Chr(KeyAscii)
    End If
    
    '判断是否由条码设备读入
    sngNow = Timer
    If txtInput.Text = "" Or strText = "" Then
        sngInputBegin = sngNow
    Else
        If Format((sngNow - sngInputBegin) / Len(strText), "0.000") < 0.04 Then blnScaner = True
    End If
    
    InputIsScaner = blnScaner
End Function



Private Function CheckBill(ByVal str配药ID串 As String) As Boolean
    Dim str收发ID串 As String
    Dim lngCount As Long
    
    If str配药ID串 = "" Then Exit Function
    If mrsTrans Is Nothing Then Exit Function
    
    mrsTrans.Filter = ""
    If mrsTrans.RecordCount = 0 Then Exit Function
    
    With mrsTrans
        For lngCount = 0 To UBound(Split(str配药ID串, ","))
            If Val(Split(str配药ID串, ",")(lngCount)) > 0 Then
                .Filter = "配药ID=" & Val(Split(str配药ID串, ",")(lngCount))
                If .RecordCount > 0 Then
                    Do While Not .EOF
                        If InStr(1, "," & str收发ID串 & ",", "," & Val(!收发ID) & ",") = 0 Then
                            str收发ID串 = IIf(str收发ID串 = "", "", str收发ID串 & ",") & Val(!收发ID)
                        End If
                        .MoveNext
                    Loop
                End If
            End If
        Next
    End With
    
    If str收发ID串 = "" Then Exit Function
    
    For lngCount = 0 To UBound(Split(str收发ID串, ","))
        If Val(Split(str收发ID串, ",")(lngCount)) > 0 Then
            If DeptSendWork_CheckBill(1, Val(Split(str收发ID串, ",")(lngCount)), mParams.bln允许未审核处方发药) > 0 Then Exit Function
        End If
    Next
    
    CheckBill = True
End Function
Private Function CheckStock() As Boolean
    '检查库存
    Dim lng收发ID As Long
    Dim str收发ID As String
    Dim rsData As ADODB.Recordset
    Dim blnIsShort As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    If mParams.IntCheckStock = 0 Then
        CheckStock = True
        Exit Function
    End If
    
    If mrsTrans Is Nothing Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Filter = "执行标志=1"
    
    If mrsTrans.RecordCount = 0 Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsTrans.Sort = "收发ID"
    
    Set rstemp = mrsTrans
    If mrsTrans.RecordCount = 0 Then
        MsgBox "读取数据异常，请重新刷新数据！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Do While Not rstemp.EOF
        If lng收发ID <> rstemp!收发ID Then
            lng收发ID = rstemp!收发ID
            str收发ID = IIf(str收发ID = "", "", str收发ID & ",") & rstemp!收发ID
        End If
        rstemp.MoveNext
        
        If Len(str收发ID) >= 3950 Then
            '重新更新库存
            gstrSQL = "Select /*+ Rule*/ " & _
                " A.ID As 收发id, A.实际数量 * Nvl(付数, 1) / D.住院包装 As 发药数量, B.实际数量 / D.住院包装 As 库存数量 " & _
                " From 药品收发记录 A, " & _
                " (Select 库房id, 药品id, Nvl(批次, 0) As 批次, Nvl(实际数量, 0) As 实际数量 " & _
                " From 药品库存 Where 性质 = 1 And 库房id = [1]) B, " & _
                " Table(Cast(f_Num2list([2]) As zlTools.t_Numlist)) C, 药品规格 D " & _
                " Where A.库房id + 0 = B.库房id(+) And A.药品id + 0 = B.药品id(+) And Nvl(A.批次, 0) = B.批次(+) And A.药品id + 0 = D.药品id " & _
                " And A.审核日期 Is Null And A.库房id + 0 = [1] And A.ID = C.Column_Value"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "检查库存", mcondition.lngCenterID, str收发ID)
            
            Do While Not rsData.EOF
                If rsData!发药数量 > rsData!库存数量 Then
                    blnIsShort = True
                End If
                
                mrsTrans.Filter = "收发ID=" & rsData!收发ID
                Do While Not mrsTrans.EOF
                    mrsTrans!库存数量 = rsData!库存数量
                    mrsTrans.Update
                    
                    mrsTrans.MoveNext
                Loop
                rsData.MoveNext
            Loop
            
            str收发ID = ""
        End If
    Loop
    
    If blnIsShort = True Then
        If mParams.IntCheckStock = 1 Then
            CheckStock = (MsgBox("本次选择配药的输液单对应的有些药品库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
            Exit Function
        Else
            MsgBox "本次选择配药的输液单对应的有些药品库存不足，不能继续，请在药品汇总列表中查看！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub GetPrivs()
    With mPrives
        .bln核查确认 = IsInString(mstrPrivs, "核查确认", ";")
        .bln取消审核 = IsInString(mstrPrivs, "取消审核", ";")
        .bln摆药确认 = IsInString(mstrPrivs, "摆药确认", ";")
        .bln取消摆药 = IsInString(mstrPrivs, "取消摆药", ";")
        .bln配药确认 = IsInString(mstrPrivs, "配药确认", ";")
        .bln取消配药 = IsInString(mstrPrivs, "取消配药", ";")
        .bln发送确认 = IsInString(mstrPrivs, "发送确认", ";")
        .bln取消发送 = IsInString(mstrPrivs, "取消发送", ";")
        .bln参数设置 = IsInString(mstrPrivs, "参数设置", ";")
        .bln销帐审核 = IsInString(mstrPrivs, "销帐审核", ";")
        .bln确认拒绝 = IsInString(mstrPrivs, "确认拒绝", ";")
        .bln销帐拒绝 = IsInString(mstrPrivs, "销帐拒绝", ";")
        .bln排班设置 = IsInString(mstrPrivs, "排班设置", ";")
    End With
End Sub

Private Sub GetParams()
    Dim strAutoPrint As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    With mParams
        '系统参数
        .lng配置中心 = Val(zlDatabase.GetPara("配置中心", glngSys, 1345, 0))
        .bln允许未审核处方发药 = (gtype_UserSysParms.P6_未审核记帐处方发药 = 1)
        .bln允许未收费处方发药 = (gtype_UserSysParms.P148_未收费处方发药 = 1)
        .bln允许取消发药 = (gtype_UserSysParms.P15_门诊收费与发药分离 = 1 Or gtype_UserSysParms.P16_住院记帐与发药分离 = 1)
        .bln医嘱作废 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0)
        .bln审核划价单 = True
        '获取处方审查系统的系统参数
        .bln处方审查 = (gtype_UserSysParms.P240_药房处方审查 = 2 Or gtype_UserSysParms.P240_药房处方审查 = 3)
        .bln审核 = (gtype_UserSysParms.P214_首次医嘱执行需要审核 = 1 And Not .bln处方审查)
        .int皮试有效天数 = (gtype_UserSysParms.P70_过敏登记有效天数 = 1)
        
        '参数设置：基础
        .int摆药后打印 = Val(zlDatabase.GetPara("摆药后打印", glngSys, 1345, 0))
        .int发送后打印 = Val(zlDatabase.GetPara("发送后打印", glngSys, 1345, 0))
        .bln批次设置 = (Val(zlDatabase.GetPara("批次设置", glngSys, 1345, 0)) = 1)
        .bln打包设置 = (Val(zlDatabase.GetPara("打包设置", glngSys, 1345, 0)) = 1)
        strAutoPrint = zlDatabase.GetPara("瓶签自动打印", glngSys, 1345, "00|00")
        .bln瓶签手工打印 = (Val(zlDatabase.GetPara("瓶签手工打印", glngSys, 1345, 0)) = 1)
        .intCount = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\输液卡片", "卡片数量", 3))
        .intNum = Val(zlDatabase.GetPara("瓶签打印份数", glngSys, 1345, 0))
        .int打印汇总 = Val(zlDatabase.GetPara("打印标签后是否打印汇总报表", glngSys, 1345, 2))
        .blnLastBatch = (Val(zlDatabase.GetPara("保持上次批次", glngSys, 1345, 0)) = 1)
        .blnTwoCode = (Val(zlDatabase.GetPara("扫两次瓶签号自动发送", glngSys, 1345, 0)) = 1)
        .blnByMedi = (Val(zlDatabase.GetPara("按批次，药品排序", glngSys, 1345, 0)) = 1)
        .intCheck = zlDatabase.GetPara("审核该药房的所有数据", glngSys, 1345, 0)
        .blnFilter = (Val(zlDatabase.GetPara("是否按设置的常用药品进行药品过滤操作", glngSys, 1345, 0)) = 1)
        .blnRePeople = (Val(zlDatabase.GetPara("打印瓶签时填写各个环节的实际操作员", glngSys, 1345, 0)) = 1)
        
        .int药品名称显示方式 = Val(zlDatabase.GetPara("药品名称显示方式", glngSys, 1345, 0))
        
        .strSourceDep = zlDatabase.GetPara("显示来源病区", glngSys, 1345, "")
        
        If InStr(1, strAutoPrint, "|") = 0 Or Len(strAutoPrint) <> 5 Then
            strAutoPrint = "00|00"
        End If
        
        If Mid(strAutoPrint, 1, 1) = 1 Then
            If Val(Mid(strAutoPrint, 2, 1)) = 1 Then
                .int瓶签摆药后打印 = 1
            Else
                .int瓶签摆药后打印 = 0
            End If
        Else
            .int瓶签摆药后打印 = 2
        End If
        If Mid(strAutoPrint, 4, 1) = 1 Then
            If Val(Mid(strAutoPrint, 5, 1)) = 1 Then
                .int瓶签配药后打印 = 1
            Else
                .int瓶签配药后打印 = 0
            End If
        Else
            .int瓶签配药后打印 = 2
        End If
            
        '库存检查规则
        .IntCheckStock = MediWork_GetCheckStockRule(.lng配置中心)

        'PASS
        .intShowPass = gintPass
'        .blnShowPass = True
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub BillPrint_Prepare()
    '打印摆药单
    Dim StrDate As String
    
    With vsfTrans
        If .Row > 0 Then StrDate = .TextMatrix(.Row, .ColIndex("摆药时间"))
    End With
      
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_1", Me, _
        "部门=" & mcondition.lngCenterID, _
        "摆药时间=" & StrDate, "操作人员=" & gstrUserName, "PrintEmpty=0", 1)
End Sub

Private Sub BillPrint_Send()
    '打印发送单
    Dim StrDate As String
    
    With vsfTrans
        If .Row > 0 Then StrDate = .TextMatrix(.Row, .ColIndex("发送时间"))
    End With
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_2", Me, _
            "部门=" & mcondition.lngCenterID, _
            "发送时间=" & StrDate, "操作人员=" & gstrUserName, "PrintEmpty=0", 1)
End Sub

Private Sub BillPrint_Return()
    '打印退药销帐清单
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_2", Me, "包装系数=C.住院包装", 2)
End Sub

Private Sub ClearDetailList()
    '清除明细列表数据
    vsfTrans.rows = 1
    vsfTrans.rows = 2
    
'    vsfDrug.rows = 1
'    vsfDrug.rows = 2
    
    vsfSumDrug.rows = 1
    vsfSumDrug.rows = 2
    
    Me.VSFLook.rows = 1
    Me.VSFLook.rows = 2
    
    Me.vsfMedis.rows = 1
    Me.vsfMedis.rows = 2
End Sub

Private Sub RefreshPrintSign(ByVal str配药id As String, ByVal dateNow As Date, Optional ByVal str工作人员 As String)
    On Error GoTo errHandle

    '更新打印标志
    gstrSQL = "Zl_输液配药记录_打印("
    '配药ID
    gstrSQL = gstrSQL & "'" & str配药id & "'"
    gstrSQL = gstrSQL & ",To_Date('" & dateNow & "','yyyy-MM-dd hh24:mi:ss')"
    gstrSQL = gstrSQL & IIf(str工作人员 <> "", ",'" & str工作人员 & " '", ",Null")
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新打印标志")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeConditionArea()
    Dim intCount As Integer
    
    On Error Resume Next

    '时间条件栏
    With picTime
        .Left = 0
        .Top = 0
        .Width = picCondition.Width
    End With
    
    With cbo时间范围
        .Width = picTime.Width - .Left - 100
    End With
    
    lblTimeBegin.Visible = (cbo时间范围.ListIndex = 3)
    With Dtp开始时间
        .Visible = (cbo时间范围.ListIndex = 3)
        .Width = cbo时间范围.Width
    End With
    
    lblTimeEnd.Visible = (cbo时间范围.ListIndex = 3)
    With Dtp结束时间
        .Visible = (cbo时间范围.ListIndex = 3)
        .Width = cbo时间范围.Width
    End With

    With picShowSendType
        If cbo时间范围.ListIndex = 3 Then
            .Top = Dtp结束时间.Top + Dtp结束时间.Height + 20
        Else
            .Top = cbo时间范围.Top + cbo时间范围.Height + 20
        End If
        .Width = picTime.Width
    End With
    
    picUpOrDown1.Left = picShowSendType.Width - picUpOrDown1.Width - 10
    
    Me.lbldept.Top = picShowSendType.Top + picShowSendType.Height + 20
    Me.txtdept.Top = picShowSendType.Top + picShowSendType.Height + 20
    txtdept.Width = Dtp开始时间.Width
    
    Me.lblName.Top = txtdept.Top + txtdept.Height + 20
    Me.txtName.Top = txtdept.Top + txtdept.Height + 20
    txtName.Width = Dtp开始时间.Width
    
    Me.lblDrug.Top = txtName.Top + txtName.Height + 20
    Me.txtDrug.Top = txtName.Top + txtName.Height + 20
    txtDrug.Width = Dtp开始时间.Width
    cmdDrug.Top = txtDrug.Top
    cmdDrug.Left = txtDrug.Left + txtDrug.Width - cmdDrug.Width
    
    Me.lblTag.Top = txtDrug.Top + txtDrug.Height + 20
    Me.txtTag.Top = txtDrug.Top + txtDrug.Height + 20
    txtTag.Width = Dtp开始时间.Width
    
    
    With picTime
        If txtTag.Visible = True Then
            .Height = txtTag.Top + txtTag.Height
        Else
            .Height = picShowSendType.Top + picShowSendType.Height
        End If
    End With
    
    '消息列表
    With Me.picMsg
        .Left = 0
        .Width = picCondition.Width
        .Height = picUpOrDown.Top + picUpOrDown.Height + 50 + IIf(lblMsgComment.Tag = "1", vsfMsg.Height + 50, 0)
        .Top = picCondition.Height - .Height - 50
    End With
    
    '部门列表栏
    With picDeptList
        .Left = 0
        .Top = picTime.ScaleTop + picTime.ScaleHeight
        .Width = picCondition.Width
        .Height = picCondition.Height - .Top - IIf(picMsg.Visible, Me.picMsg.Height, 0) - 50
    End With
   
    With fraLineH1
        .Top = 50
        .Width = picTime.Width + 100
    End With
   
 End Sub

Private Sub DeleteBatch(ByVal lngType As Long)
    '删除工作批次
    Dim strInputID As String
    Dim lngRow As Long
    Dim strCom As String
    
    On Error GoTo errHandle
    
    With vsfTrans
        If lngType = conMenu_Oper_DelBatch_SelBatch Then
            strCom = .TextMatrix(.Row, .ColIndex("配药批次"))
        ElseIf lngType = conMenu_Oper_DelBatch_SelDept Then
            strCom = .TextMatrix(.Row, .ColIndex("病区"))
        ElseIf lngType = conMenu_Oper_DelBatch_SelPati Then
            strCom = .TextMatrix(.Row, .ColIndex("病区")) & .TextMatrix(.Row, .ColIndex("姓名")) & .TextMatrix(.Row, .ColIndex("床号"))
        End If
        
        If lngType = conMenu_Oper_DelBatch_SelRow Then
            '当前行
            If .TextMatrix(.Row, .ColIndex("配药批次")) <> "" Then
                strInputID = Val(.TextMatrix(.Row, .ColIndex("配药ID")))
            End If
        Else
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                    If lngType = conMenu_Oper_DelBatch_SelBatch Then
                        If .TextMatrix(lngRow, .ColIndex("配药批次")) <> "" And .TextMatrix(lngRow, .ColIndex("配药批次")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_SelDept Then
                        If .TextMatrix(lngRow, .ColIndex("配药批次")) <> "" And .TextMatrix(lngRow, .ColIndex("病区")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_SelPati Then
                        If .TextMatrix(lngRow, .ColIndex("配药批次")) <> "" And .TextMatrix(lngRow, .ColIndex("病区")) & .TextMatrix(lngRow, .ColIndex("姓名")) & .TextMatrix(lngRow, .ColIndex("床号")) = strCom Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                        End If
                    ElseIf lngType = conMenu_Oper_DelBatch_AllRow Then
                        If .TextMatrix(lngRow, .ColIndex("配药批次")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    If strInputID = "" Then Exit Sub
    
    gstrSQL = "Zl_输液配药记录_清除批次("
    '配药ID
    gstrSQL = gstrSQL & "'" & strInputID & "'"
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "清除批次")
    
    DoEvents
    
    '本地数据集更新
    strInputID = "," & strInputID & ","
    With mrsTrans
        .Filter = ""
        Do While Not .EOF
            If InStr(strInputID, "," & !配药id & ",") > 0 Then
                !配药批次 = ""
                .Update
            End If
            .MoveNext
        Loop
    End With
    
    DoEvents
    
    '更新列表显示
    With vsfTrans
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If InStr(strInputID, "," & .TextMatrix(lngRow, .ColIndex("配药ID")) & ",") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("配药批次")) = ""
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    MsgBox "清除批次完成！", vbInformation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub GetCondition()
    Dim strStartTime As String
    Dim strEndTime As String
    
    '时间范围
    Select Case cbo时间范围.ListIndex
        Case 0
            '当日
            mcondition.intTransTimeSel = 0
            mcondition.strTransStartTime = Format(mdateToday, "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(mdateToday, "yyyy-mm-dd") & " 23:59:59"
        Case 1
            '明日
            mcondition.intTransTimeSel = 1
            mcondition.strTransStartTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 23:59:59"
        Case 2
            '今日和明日
            mcondition.intTransTimeSel = 2
            mcondition.strTransStartTime = Format(mdateToday, "yyyy-mm-dd") & " 00:00:00"
            mcondition.strTransEndTime = Format(DateAdd("d", 1, mdateToday), "yyyy-mm-dd") & " 23:59:59"
        Case 3
            '指定日期范围
            mcondition.intTransTimeSel = 3
            mcondition.strTransStartTime = Format(Dtp开始时间.Value, "yyyy-mm-dd hh:mm:ss")
            mcondition.strTransEndTime = Format(Dtp结束时间.Value, "yyyy-mm-dd hh:mm:ss")
    End Select
End Sub

Private Sub GetTransCount(ByVal dateStart As Date, ByVal dateEnd As Date)
    '取病区及病区对应的输液单据数量
    Dim rsTmp As ADODB.Recordset
    Dim lngCount As Long
    Dim intTabIndex As Integer
    Dim strCaption As String
    Dim lng病区id As Long
    Dim intType As Integer
    Dim lng数量 As Long
    Dim str记录id As String
    Dim str类型 As String
    
    '病区对应的输液单数
    Set mrsDeptTrans = New ADODB.Recordset
    With mrsDeptTrans
        If .State = 1 Then .Close
        
        .Fields.Append "选择", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "类型", adDouble, 1, adFldIsNullable
        .Fields.Append "病区ID", adDouble, 18, adFldIsNullable
        .Fields.Append "病区", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 10, adFldIsNullable
        .Fields.Append "记录id", adLongVarChar, 20000, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        If Mid(Me.lblName.Caption, 1, Len(Me.lblName.Caption) - 1) = "姓名" Then
            intType = 1
        ElseIf Mid(Me.lblName.Caption, 1, Len(Me.lblName.Caption) - 1) = "床号" Then
            intType = 2
        Else
            intType = 3
        End If
        
        Set rsTmp = PIVA_GetTransCount(mcondition.lngCenterID, dateStart, dateEnd, mParams.bln审核, mParams.bln处方审查, intType, Me.txtName.Text, Val(Me.txtDrug.Tag), Me.txtTag.Text, Val(Me.txtdept.Tag), mParams.intCheck, mParams.strSourceDep)
        
        '重组记录集
        rsTmp.Sort = "类型,病区id"
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                If rsTmp!类型 = "00" And mParams.intCheck = 1 Then
                    
                    If lng病区id <> rsTmp!病区ID Or str类型 <> rsTmp!类型 Then
                    
                        If lng病区id <> 0 Then
                            !数量 = lng数量
                            !记录id = str记录id
                        End If
                        
                        lng病区id = rsTmp!病区ID
                        .AddNew
                        !选择 = 0
                        !类型 = rsTmp!类型
                        !病区ID = rsTmp!病区ID
                        !病区 = rsTmp!病区
                        !名称 = rsTmp!名称
                        !编码 = rsTmp!编码
                        str类型 = rsTmp!类型
                        If nvl(rsTmp!药师审核标志, 0) = 0 Or nvl(rsTmp!药师审核标志, 0) = 3 Then
                            lng数量 = 1
                        Else
                            lng数量 = 0
                        End If
                        str记录id = rsTmp!Id
                    Else
                        If nvl(rsTmp!药师审核标志, 0) = 0 Or nvl(rsTmp!药师审核标志, 0) = 3 Then
                            lng数量 = lng数量 + 1
                        End If
                        str记录id = str记录id & "," & rsTmp!Id
                    End If
                Else
                    If lng病区id <> rsTmp!病区ID Or str类型 <> rsTmp!类型 Then
                    
                        If lng病区id <> 0 Then
                            !数量 = lng数量
                            !记录id = str记录id
                        End If
                        
                        lng病区id = rsTmp!病区ID
                        .AddNew
                        !选择 = 0
                        !类型 = rsTmp!类型
                        !病区ID = rsTmp!病区ID
                        !病区 = rsTmp!病区
                        !名称 = rsTmp!名称
                        !编码 = rsTmp!编码
                        str类型 = rsTmp!类型
                        lng数量 = 1
                        str记录id = rsTmp!Id
                    Else
                        lng数量 = lng数量 + 1
                        str记录id = str记录id & "," & rsTmp!Id
                    End If
                End If
                .Update
                
                rsTmp.MoveNext
                
                If rsTmp.EOF Then
                    !数量 = lng数量
                    !记录id = str记录id
                End If
                
            Loop
        End If
    End With
    
    '计算各业务环节的医嘱或输液单数量，并显示在分页标签上
    For intTabIndex = 0 To Me.tbcLook.ItemCount - 1
        lngCount = 0

        If Not mrsDeptTrans Is Nothing Then
            mrsDeptTrans.Filter = "类型='" & tbcLook.Item(intTabIndex).Tag & "'"
            Do While Not mrsDeptTrans.EOF
                lngCount = lngCount + mrsDeptTrans!数量

                mrsDeptTrans.MoveNext
            Loop
            
            strCaption = tbcLook.Item(intTabIndex).Caption
            strCaption = Mid(strCaption, 1, InStr(1, strCaption, "(")) & lngCount & ")"
            tbcLook.Item(intTabIndex).Caption = strCaption
        End If
    Next
    
    For intTabIndex = 0 To Me.tabWork.ItemCount - 1
        lngCount = 0

        If Not mrsDeptTrans Is Nothing Then
            mrsDeptTrans.Filter = "类型='" & tabWork.Item(intTabIndex).Tag & "'"
            Do While Not mrsDeptTrans.EOF
                lngCount = lngCount + mrsDeptTrans!数量

                mrsDeptTrans.MoveNext
            Loop
            
            strCaption = Me.tabWork.Item(intTabIndex).Caption
            strCaption = Mid(strCaption, 1, InStr(1, strCaption, "(")) & lngCount & ")"
            tabWork.Item(intTabIndex).Caption = strCaption
        End If
    Next
    Call SetTabColor(tabWork)
    Call SetTabColor(tbcLook)
End Sub
Private Sub GetWorkBatchRec()
    '取输液配置中心的工作批次
    On Error GoTo errHandle
    gstrSQL = "Select 批次,颜色, 配药时间, 给药时间, 打包, 1 正常 From 配药工作批次 Where 启用=1 and 配置中心ID=[1] " & _
        " Union All " & _
        " Select Max(Nvl(批次, 0))+1 批次,0 颜色, '' 配药时间, '' 给药时间, 0 打包, 0 正常 From 配药工作批次 where 配置中心ID=[1] " & _
        " Order By 批次"
    Set mrsWorkBatch = zlDatabase.OpenSQLRecord(gstrSQL, "取输液配置中心工作批次", mParams.lng配置中心)
    
    mstr批次 = ""
    mParams.strBatchList = ""
    mstr打包 = ""
    With mrsWorkBatch
        cboBatch.Clear
        cboBatch.AddItem "<全部>"
        Do While Not .EOF
            If !正常 = 1 Then
                mParams.strBatchList = IIf(mParams.strBatchList = "", "", mParams.strBatchList & "|") & !批次 & "#" & _
                    vbTab & "配药时间" & !配药时间 & _
                    vbTab & "给药时间" & !给药时间
                    
                mstr打包 = mstr打包 & "," & !批次 & "#," & zlStr.nvl(!打包, 0)
                    
                mstr批次 = IIf(mstr批次 = "", "", mstr批次 & "/") & !批次 & "#" & _
                    "|" & "配药时间" & !配药时间 & _
                    "|" & "给药时间" & !给药时间 & "," & IIf(zlStr.nvl(!颜色) = "", 0, !颜色)
            Else
                mParams.strBatchList = IIf(mParams.strBatchList = "", "", mParams.strBatchList & "|") & zlStr.nvl(!批次, 1) & "#" & vbTab & "（备用批次）"
                mstr批次 = IIf(mstr批次 = "", "", mstr批次 & "/") & zlStr.nvl(!批次, 1) & "#|（备用批次）" & "," & IIf(zlStr.nvl(!颜色) = "", 0, !颜色)
                mstr打包 = IIf(mstr打包 = "", "", mstr打包 & ",") & zlStr.nvl(!批次, 1) & "#,0"
            End If
            
            '记载批次信息到下拉框   IIf(mstr打包 = "", "", mstr打包 & ",") & NVL(!批次, 1) & "#,0"
            cboBatch.AddItem !批次 & "#"
            
            .MoveNext
        Loop
        cboBatch.Text = "<全部>"
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniPriRec()
    Set mrstemp = New ADODB.Recordset
    With mrstemp
        If .State = 1 Then .Close
        .Fields.Append "配药id", adDouble, 18, adFldIsNullable
        .Fields.Append "部门id", adDouble, 18, adFldIsNullable
        .Fields.Append "配药类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub IniTransRec()
    '输液单记录集
    Set mrsTrans = New ADODB.Recordset
    With mrsTrans
        If .State = 1 Then .Close
        
        '该记录对应的输液配药记录信息
        .Fields.Append "组号", adDouble, 18, adFldIsNullable
        .Fields.Append "配药id", adDouble, 18, adFldIsNullable
        .Fields.Append "部门id", adDouble, 18, adFldIsNullable
        .Fields.Append "序号", adDouble, 3, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "床号排序", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病区", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "执行时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人id", adDouble, 18, adFldIsNullable
        .Fields.Append "主页id", adDouble, 18, adFldIsNullable
        .Fields.Append "优先级", adDouble, 18, adFldIsNullable
        .Fields.Append "病人科室id", adDouble, 18, adFldIsNullable
        .Fields.Append "打包时间", adLongVarChar, 20, adFldIsNullable
        
        '输液配药记录业务操作信息
        .Fields.Append "配药批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "瓶签号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "打印标志", adDouble, 1, adFldIsNullable
        .Fields.Append "是否打包", adDouble, 1, adFldIsNullable
        .Fields.Append "核查人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "核查时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "摆药单号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "配药时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发送人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "发送时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐申请人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐申请时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐审核人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "销帐审核时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "抗菌药物", adLongVarChar, 1, adFldIsNullable
        .Fields.Append "药师审核时间", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "是否调整批次", adDouble, 1, adFldIsNullable
        .Fields.Append "是否锁定", adDouble, 1, adFldIsNullable
        .Fields.Append "新配药批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "手工调整批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "拒收原因", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "是否确认调整", adDouble, 1, adFldIsNullable
        
        '输液配药记录对应的药品信息
        .Fields.Append "收发id", adDouble, 18, adFldIsNullable
        .Fields.Append "单据", adDouble, 2, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable   '编码+通用名/商品名
        .Fields.Append "药品编码名称", adLongVarChar, 50, adFldIsNullable   '固定显示编码+通用名/商品名,用于汇总列表的排序
        .Fields.Append "通用名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "商品名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "英文名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "产地", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单量", adDouble, 20, adFldIsNullable
        .Fields.Append "剂量单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "用法", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药名id", adDouble, 18, adFldIsNullable
        .Fields.Append "费用序号", adDouble, 3, adFldIsNullable
        .Fields.Append "费用id", adDouble, 18, adFldIsNullable
        .Fields.Append "配药类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "溶媒", adDouble, 1, adFldIsNullable
        .Fields.Append "是否皮试", adDouble, 1, adFldIsNullable
        .Fields.Append "配药类型1", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "批次标记", adDouble, 1, adFldIsNullable
        .Fields.Append "溶媒id", adDouble, 18, adFldIsNullable
        .Fields.Append "排序药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "排序单量", adDouble, 20, adFldIsNullable
        .Fields.Append "是否常用", adDouble, 1, adFldIsNullable
        
        .Fields.Append "发药数量", adDouble, 18, adFldIsNullable
        .Fields.Append "库存数量", adDouble, 18, adFldIsNullable
        .Fields.Append "实际数量", adDouble, 18, adFldIsNullable
        
        .Fields.Append "审查结果", adDouble, 1, adFldIsNullable
        .Fields.Append "医嘱发送时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "开嘱时间", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "皮试结果", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "对应医嘱ID", adDouble, 18, adFldIsNullable
        .Fields.Append "发送号", adDouble, 18, adFldIsNullable
        .Fields.Append "执行频次", adLongVarChar, 50, adFldIsNullable
        
        .Fields.Append "执行标志", adDouble, 1, adFldIsNullable
        .Fields.Append "作废类型", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "险类", adDouble, 5, adFldIsNullable
        .Fields.Append "颜色", adDouble, 18, adFldIsNullable
        .Fields.Append "销帐原因", adLongVarChar, 200, adFldIsNullable
        
        .Fields.Append "实际配药类型", adLongVarChar, 50, adFldIsNullable       '用于显示所有药品的配药类型，包括“溶媒药品”
        
        .Fields.Append "执行性质", adDouble, 1, adFldIsNullable
        .Fields.Append "执行标记", adDouble, 1, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub


Private Sub InitPanes()
    '初始化分栏控件
    'DockingPane
    '-----------------------------------------------------
    Me.dkpMain.SetCommandBars Me.cbsMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
'    Me.dkpMain.Options.DefaultPaneOptions = PaneNoCloseable + PaneNoFloatable + PaneNoHideable + PaneNoCaption
    
    Dim objPaneCon As Pane

    Set objPaneCon = Me.dkpMain.CreatePane(mconPane_PIVA_Condition, 225, 100, DockLeftOf, Nothing)
    objPaneCon.Title = mstrCenterName
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
End Sub

Private Function Find作废类型(ByVal rsData As ADODB.Recordset, ByVal lng配药id As Long) As String
    '功能：显示自备药或不取药的作废类型
    
    Find作废类型 = ""
    
    '若未启用相关显示，则不进行查找
    If Not mblnShowOhters Then Exit Function
    
    rsData.Filter = "配药id = " & lng配药id
    
    Do While Not rsData.EOF
        If nvl(rsData!作废类型) <> "" Then
            Find作废类型 = rsData!作废类型
            Exit Function
        End If
        
        rsData.MoveNext
    Loop
    
    Find作废类型 = ""
    
End Function

Private Sub LoadTrans(ByVal strIDS As String, ByVal strStep As String, ByVal intPack As Integer, ByVal intSend As Integer)
    Dim rsTrans As ADODB.Recordset
    Dim lng配药id As Long
    Dim int配药类型 As Integer
    Dim rstemp As Recordset
    Dim dbl容量 As Double
    Dim strOld执行时间 As String
    Dim strOld批次 As String
    Dim lngOld病人id As Long
    Dim lngOld配药id As Long
    Dim lng单量 As Long
    Dim lngCount As Long
    Dim lng优先级 As Long
    Dim int序号 As Integer
    Dim i As Integer
    Dim lng溶媒id  As Long
    Dim lng药品id As Long
    Dim str配药类型 As String
    Dim dbl单量 As Double
    Dim arrExecute As Variant
    Dim rsSel As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Call IniTransRec
    
    Call IniPriRec
    mintCountPack = 0
    If Not mParams.blnTwoCode Then Me.cboBatch.ListIndex = 0
    
    mblnFilter = False
    Me.cboLevel.ListIndex = 0
    Me.cboMedi.ListIndex = 0
    Me.cboDosType.ListIndex = 0
    Me.cboFrequency.ListIndex = 0
    mblnFilter = True
    
    If Not (mcondition.strTransStep = M_STR_CALSS_SEND And mParams.blnTwoCode = True) Then Me.cboBatch.ListIndex = 0
    
    arrExecute = GetArrayByStr(strIDS, 3950, ",")
    For i = 0 To UBound(arrExecute)
    
        
        Set rsTrans = Piva_GetTrans(CStr(arrExecute(i)), mParams.lng配置中心, strStep, intPack, mblnShowOhters)
        
        Set rsSel = rsTrans.Clone
        
        With rsTrans
            Set rstemp = rsTrans
            If .RecordCount > 0 Then
                rsTrans.Sort = "病人id,配药id,执行时间,配药批次,溶媒,医嘱序号"
                Do While Not .EOF
                    lngCount = lngCount + 1
                                    
                    mrsTrans.AddNew
                    mrsTrans!配药id = !配药id
                    mrsTrans!部门ID = !部门ID
                    mrsTrans!序号 = !序号
                    mrsTrans!姓名 = IIf(IsNull(!姓名), "", !姓名)
                    mrsTrans!性别 = IIf(IsNull(!性别), "", !性别)
                    mrsTrans!年龄 = IIf(IsNull(!年龄), "", !年龄)
                    mrsTrans!住院号 = IIf(IsNull(!住院号), "", !住院号)
                    mrsTrans!床号 = IIf(IsNull(!床号), "", !床号)
                    mrsTrans!床号排序 = IIf(IsNull(!床号排序), "", !床号排序)
                    mrsTrans!编码 = IIf(IsNull(!编码), "", !编码)
                    mrsTrans!病区 = !病人病区
                    mrsTrans!科室 = !病人科室
                    mrsTrans!执行时间 = IIf(IsNull(!执行时间), "", Format(!执行时间, "YYYY-MM-DD HH:MM"))
                    mrsTrans!病人ID = IIf(IsNull(!病人ID), 0, !病人ID)
                    mrsTrans!主页id = IIf(IsNull(!主页id), 0, !主页id)
                    mrsTrans!病人科室id = IIf(IsNull(!病人科室id), 0, !病人科室id)
                    mrsTrans!打包时间 = nvl(!打包时间)
                    
                    mrsTrans!配药批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
                    mrsTrans!新配药批次 = IIf(IsNull(!配药批次), "", !配药批次 & "#")
                    mrsTrans!瓶签号 = IIf(IsNull(!瓶签号), "", !瓶签号)
                    mrsTrans!打印标志 = IIf(IIf(IsNull(!打印标志), 0, !打印标志) = 0, 0, 1)
                    mrsTrans!是否打包 = IIf(IsNull(!是否打包), 0, !是否打包)
                    mrsTrans!核查人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!核查时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!摆药人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!摆药时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!摆药单号 = IIf(IsNull(!摆药单号), "", !摆药单号)
                    mrsTrans!配药人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!配药时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!发送人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!发送时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!销帐申请人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!销帐申请时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!销帐审核人 = IIf(IsNull(!操作人员), "", !操作人员)
                    mrsTrans!销帐审核时间 = IIf(IsNull(!操作时间), "", Format(!操作时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!抗菌药物 = 1
                    mrsTrans!药师审核时间 = IIf(IsNull(!药师审核时间), 0, !药师审核时间)
                    mrsTrans!是否调整批次 = IIf(IsNull(!是否调整批次), 0, !是否调整批次)
                    mrsTrans!是否锁定 = IIf(IsNull(!是否锁定), 0, !是否锁定)
                    mrsTrans!手工调整批次 = IIf(IsNull(!手工调整批次), 0, !手工调整批次)
                    mrsTrans!拒收原因 = nvl(!拒收原因)
                    mrsTrans!是否确认调整 = IIf(IsNull(!是否确认调整), 0, !是否确认调整)
                    
                    mrsTrans!收发ID = !收发ID
                    mrsTrans!单据 = !单据
                    mrsTrans!NO = nvl(!NO)
                    
                    mrsTrans!药品编码名称 = IIf(IsNull(!药品编码), !通用名, "[" & !药品编码 & "]" & !通用名)
                    
                    If mParams.int药品名称显示方式 = 0 Then
                        '编码和名称
                        mrsTrans!药品名称 = IIf(IsNull(!药品编码), !通用名, "[" & !药品编码 & "]" & !通用名)
                    ElseIf mParams.int药品名称显示方式 = 1 Then
                        '名称
                        mrsTrans!药品名称 = !通用名
                    ElseIf mParams.int药品名称显示方式 = 2 Then
                        '编码
                        mrsTrans!药品名称 = IIf(IsNull(!药品编码), "", "[" & !药品编码 & "]")
                    End If

                    mrsTrans!通用名 = !通用名
                    mrsTrans!商品名 = IIf(IsNull(!商品名), "", !商品名)
                    mrsTrans!英文名 = IIf(IsNull(!英文名), "", !英文名)
                    mrsTrans!规格 = IIf(IsNull(!规格), "", !规格)
                    mrsTrans!产地 = IIf(IsNull(!产地), "", !产地)
                    mrsTrans!批号 = IIf(IsNull(!批号), "", !批号)
                    mrsTrans!单量 = FormatEx(nvl(!单量, 0), 2)
                    mrsTrans!剂量单位 = !剂量单位
                    mrsTrans!频次 = IIf(IsNull(!频次), "", !频次)
                    mrsTrans!数量 = nvl(!数量, 0)
                    mrsTrans!单位 = !单位
                    mrsTrans!批次 = !批次
                    mrsTrans!用法 = IIf(IsNull(!用法), "", !用法)
                    mrsTrans!药品ID = nvl(!药品ID, 0)
                    mrsTrans!药名ID = !药名ID
                    mrsTrans!费用序号 = !费用序号
                    mrsTrans!费用ID = !费用ID
                    mrsTrans!优先级 = nvl(!优先级, 0)
                    mrsTrans!批次标记 = nvl(!批次标记, 0)
                    mrsTrans!配药类型 = !配药类型1
                    mrsTrans!实际配药类型 = !配药类型1
                    mrsTrans!发药数量 = nvl(!发药数量, 0)
                    mrsTrans!库存数量 = nvl(!库存数量, 0)
                    mrsTrans!实际数量 = nvl(!实际数量, 0)
                    
                    mrsTrans!医嘱id = !医嘱id
                    mrsTrans!对应医嘱ID = !对应医嘱ID
                    mrsTrans!发送号 = !发送号
                    mrsTrans!开嘱时间 = IIf(IsNull(!开嘱时间), "", Format(!开嘱时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!皮试结果 = nvl(!皮试结果)
                    mrsTrans!医嘱发送时间 = IIf(IsNull(!医嘱发送时间), "", Format(!医嘱发送时间, "YYYY-MM-DD HH:MM:SS"))
                    mrsTrans!审查结果 = nvl(!审查结果, 0)
                    mrsTrans!作废类型 = IIf(IsNull(!作废类型), Find作废类型(rsSel, !配药id), !作废类型)
                    mrsTrans!险类 = !险类
                    mrsTrans!颜色 = !颜色
                    mrsTrans!执行标志 = 0
                    mrsTrans!溶媒 = nvl(!溶媒, 0)
                    mrsTrans!是否皮试 = nvl(!是否皮试, 0)
                    mrsTrans!执行频次 = nvl(!执行频次, 0)
                    mrsTrans!排序药品ID = nvl(!药品ID, 0)
                    mrsTrans!排序单量 = FormatEx(nvl(!单量, 0), 2)
                    mrsTrans!销帐原因 = IIf(IsNull(!销帐原因), "", !销帐原因)
                    
                    mrsTrans!执行性质 = nvl(!执行性质, 0)
                    mrsTrans!执行标记 = nvl(!执行标记, 0)
                    
                    If mParams.blnFilter And mParams.str常用药品 <> "" Then
                        mrsTrans!是否常用 = IIf(InStr(1, "," & mParams.str常用药品 & ",", !药品ID) > 0, 1, 0)
                    End If
                    
                    If !配药id <> lng配药id Then
                        int序号 = int序号 + 1
                    End If
                    mrsTrans!组号 = int序号
                    mrsTrans.Update
                    
                    
                    If !配药id = lng配药id Then
                        
                        If Val(!配药类型) > 0 Then
                            int配药类型 = 1
                        ElseIf int配药类型 = 0 And Val(!配药类型) = 0 Then
                            int配药类型 = 0
                        End If
                        
                        If str配药类型 = "" Then
                            str配药类型 = nvl(!配药类型1)
                        End If
                    Else
                        int配药类型 = Val(!配药类型)
                        If nvl(!溶媒, 0) = 0 Then
                            lng药品id = nvl(!药品ID, 0)
                            dbl单量 = FormatEx(nvl(!单量, 0), 2)
                            str配药类型 = nvl(!配药类型1)
                        End If
                    End If
                    
                    If !溶媒 = 1 Then
                        lng溶媒id = nvl(!药品ID, 0)
                    End If
                    
                    mrsTrans.Filter = ""
                    lng配药id = !配药id
                    
                    .MoveNext
                    
                    If .EOF Then
                        mrsTrans.Filter = "配药id=" & lng配药id
                        mrsTrans.MoveFirst
                        Do While Not mrsTrans.EOF
                            mrsTrans.Update "抗菌药物", int配药类型
                            mrsTrans.Update "溶媒id", lng溶媒id
                            mrsTrans.Update "排序药品id", lng药品id
                            mrsTrans.Update "配药类型", str配药类型
                            mrsTrans.MoveNext
                        Loop
                    Else
                        If lng配药id <> !配药id Then
                            mrsTrans.Filter = "配药id=" & lng配药id
                            mrsTrans.MoveFirst
                            Do While Not mrsTrans.EOF
                                mrsTrans.Update "抗菌药物", int配药类型
                                mrsTrans.Update "溶媒id", lng溶媒id
                                mrsTrans.Update "排序药品id", lng药品id
                                mrsTrans.Update "配药类型", str配药类型
                                mrsTrans.Update "排序单量", dbl单量
                                mrsTrans.MoveNext
                            Loop
                            lng药品id = 0
                            lng溶媒id = 0
                            str配药类型 = ""
                        End If
                    End If
                    
                    
                Loop
            End If
        End With
    Next
    
    mrsTrans.Filter = ""
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboDosType_Click()
  Call SetFilter
End Sub

Private Sub cboFrequency_Click()
    Call SetFilter
End Sub

Private Sub cboMedi_Click()
    Call SetFilter
End Sub



Private Sub cboType_Click()
    Call LoadVsfMedi(mstr上次IDS, False)
    If mcondition.strTransStep = M_STR_CALSS_SEND Then Me.txtFindItem.SetFocus
End Sub

Private Sub chkAllDept_Click()
    Dim lngRow As Long
    
    With vsfDept(tabDeptList.Selected.index)
        If .rows = 1 Then Exit Sub
        If Val(.TextMatrix(1, .ColIndex("病区ID"))) = 0 Then Exit Sub
        
        For lngRow = 1 To .rows - 1
            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(chkAllDept.Value = 0, 0, -1)
        Next
    End With
    
'    DoEvents
'    Call RefreshDetailList(Me.tabDeptList.Selected.index)
End Sub

Private Sub chkChange_Click(index As Integer)
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub chkCheck_Click()
    Dim lngRow As Long
    
    With Me.vsfMedis
        For lngRow = 1 To .rows - 1
            If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
                .TextMatrix(lngRow, .ColIndex("标志")) = IIf(Me.chkCheck.Value = 1, "1", "0")
                .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = IIf(Me.chkCheck.Value = 1, Me.ImgList.ListImages(3).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
            Else
                If nvl(.TextMatrix(lngRow, .ColIndex("相关id")), "0") <> "00" And Val(.TextMatrix(lngRow, .ColIndex("摆药标志"))) <> 1 Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = IIf(chkCheck.Value = 0, 0, -1)
                End If
            End If
        Next
    End With
End Sub

Private Sub chkResult_Click(index As Integer)
    Call LoadVsfMedi(mstr上次IDS, True)
End Sub

Private Sub chkSure_Click(index As Integer)
    '切换确认状态时清除勾选标志
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub chkPrint_Click(index As Integer)
    '切换确认状态时清除勾选标志
    Call UpdateExeSign(0, 0)
    chkAll.Value = 0
    
    Call SetFilter
End Sub

Private Sub SetBeach()
    Dim lngRow As Long
    Dim strInput As String
    Dim lng配药id As Long
    Dim arrExecute As Variant
    
    With Me.vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药id"))) <> 0 Then
                mrsTrans.Filter = "配药id=" & Val(.TextMatrix(lngRow, .ColIndex("配药id")))
                If mrsTrans!手工调整批次 <> 1 And mrsTrans!是否调整批次 = 1 Then
                    .TextMatrix(lngRow, .ColIndex("配药批次")) = mrsTrans!新配药批次
                    If lng配药id <> Val(.TextMatrix(lngRow, .ColIndex("配药id"))) Then
                        lng配药id = Val(.TextMatrix(lngRow, .ColIndex("配药id")))
                        
                        If zlStr.nvl(mrsTrans!新配药批次) = "" Then
                            strInput = IIf(strInput = "", "", strInput & "|") & mrsTrans!配药id & ",:" & zlStr.nvl(mrsTrans!优先级)
                        Else
                            strInput = IIf(strInput = "", "", strInput & "|") & mrsTrans!配药id & "," & Mid(mrsTrans!新配药批次, 1, IIf(Len(mrsTrans!新配药批次) = 0, 0, Len(mrsTrans!新配药批次) - 1)) & ":" & zlStr.nvl(mrsTrans!优先级)
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    On Error GoTo errHandle
    
    arrExecute = GetArrayByStr(strInput, 3950, "|")
    For lngRow = 0 To UBound(arrExecute)
        If mrsPRI.RecordCount > 0 Or mrsVol.RecordCount > 0 Then
            gstrSQL = "Zl_输液配药记录_分批("
            '配药ID,批次
            gstrSQL = gstrSQL & "'" & arrExecute(lngRow) & "'"
            gstrSQL = gstrSQL & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "设置批次")
        End If
    Next

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLock(ByVal intType As Integer, ByVal str配药id As String, Optional ByVal blnRow As Boolean)
'将指定的医嘱锁定或解锁
'intType:1-锁定,0-解锁
'str配药id:配药id的字符串
    Dim arrExecute As Variant
    Dim blnBeginTrans As Boolean
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    If str配药id = "" Then
        If mrsTrans Is Nothing Then Exit Sub
        With mrsTrans
            .Filter = "执行标志=1"
            .Sort = "病区,配药批次,住院号"
            If .RecordCount > 0 Then
                .MoveFirst
            Else
                Exit Sub
            End If
            
            Do While Not .EOF
                If InStr(1, "," & str配药id & ",", "," & !配药id & ",") = 0 Then
                    str配药id = IIf(str配药id = "", "", str配药id & ",") & !配药id
                End If
                .MoveNext
            Loop
        End With
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(str配药id, 3950, ",")
    For lngRow = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_锁定("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrExecute(lngRow) & "'"
        '是否锁定
        gstrSQL = gstrSQL & "," & intType
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "SetLock-锁定")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    If Not blnRow Then
        With Me.vsfTrans
            For lngRow = 1 To .rows - 1
                If InStr(1, "," & str配药id & ",", "," & .TextMatrix(lngRow, .ColIndex("配药id")) & ",") > 0 And Val(.TextMatrix(lngRow, .ColIndex("配药id"))) > 0 Then
                    .TextMatrix(lngRow, .ColIndex("是否锁定")) = intType
                    .Cell(flexcpPicture, lngRow, .ColIndex("锁"), lngRow, .ColIndex("锁")) = IIf(.TextMatrix(lngRow, .ColIndex("是否锁定")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("锁"), lngRow, .ColIndex("锁")) = flexPicAlignCenterCenter
                End If
            Next
        End With
    End If
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub DelTransRec()
    '执行功能后更新记录集：删除界面已选择的记录
    Dim lngRow As Long
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                    Do While Not mrsTrans.EOF
                        mrsTrans.Delete
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
            End If
        Next
    End With
End Sub
Private Sub RefreshDeptList(ByVal index As Integer)
    If mblnActive = True And vsfDept(index).Visible = True Then vsfDept(index).SetFocus
    mlng已扫描 = 0
    mlng未扫描 = 0
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If
    
    '清除输液单明细列表
    Call ClearDetailList
    '清除输液单卡片
    mfrmPIVCard.ClearCard
    '组织查询条件
    Call GetCondition
    '取病区及病区对应的输液单数量
    Call GetTransCount(CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime))
    '显示病区及病区对应的输液单据数量
    Call ShowDeptTrans(index, IIf(index = CNUMWORK, tabWork.Selected.Tag, tbcLook.Selected.Tag))
    
    chkAllDept.Value = 0
End Sub

Public Sub RefreshDetailList(ByVal index As Long)
    '刷新和显示输液单据列表
    Dim str病区id As String
    Dim i As Integer
    Dim bln列表 As Boolean
    Dim bln卡片 As Boolean
    Dim strIDS As String
    
    On Error GoTo errHandle
    
    Call AviShow(Me)
    
    chkAll.Enabled = False
    chkAll.Value = 0
    Me.chkCheck.Value = 0
    Me.lblVolu.Caption = "容量：0"
    Me.lblMsg.Visible = True
    Me.lblMsg.Caption = ""
    
    Call ClearDetailList
    Call mfrmPIVCard.ClearCard
    
    If vsfDept(index).Visible Then vsfDept(index).SetFocus

    If Not tbcDetail.Item(mDetailType.输液单卡片).Selected Then
        tbcDetail.Item(mDetailType.输液单列表).Selected = True
    End If
    
    mstr上次病区ID = ""
    mstr上次IDS = ""
    
    With vsfDept(index)
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("病区id"))) > 0 And Val(.TextMatrix(i, .ColIndex("选择"))) = -1 Then
                strIDS = IIf(strIDS = "", "", strIDS & ",") & .TextMatrix(i, .ColIndex("记录id"))
                str病区id = IIf(str病区id = "", "", str病区id & ",") & .TextMatrix(i, .ColIndex("病区id"))
            End If
        Next
    End With
    
    If strIDS = "" Then Call AviShow(Me, False): Exit Sub
    
    If Not mParams.blnFilter Then
        mblnFilter = False
        Call zlGetMediNum(CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime), str病区id, mcondition.strTransStep)
        mblnFilter = True
    End If
    
    mstr上次病区ID = str病区id
    mstr上次IDS = strIDS
    
    Call GetCondition
    
    If mParams.bln审核 And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "病人医嘱列表"
        Call LoadVsfMedi(strIDS)
        bln卡片 = mfrmPIVCard.ShowDetailCard(mrsMedi, mstr批次, False, mParams.intCount, mParams.bln批次设置, mParams.bln打包设置, mcondition.strTransStep, mParams.bln审核)
    Else
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "输液单列表"
        '重新取所选病区对应的输液单明细
        Call LoadTrans(strIDS, mcondition.strTransStep, Val(vsfTrans.Tag), Val(fraDetailCtr.Tag))
         '在状态栏显示所选的病区和输液单数量
'        Call GetCount
'        mrsTrans.Filter = ""
'
'        '显示输液单卡片
'        bln卡片 = mfrmPIVCard.ShowDetailCard(mrsTrans, mstr批次, mcondition.strTransStep = M_STR_CALSS_PREPARE, mParams.intCount, mParams.bln批次设置, mParams.bln打包设置, mcondition.strTransStep, mParams.bln审核)
'        '显示输液单明细列表
'        bln列表 = ShowTrans(index)
'        '显示输液单药品汇总列表
'        Call ShowSumDrug

        Call SetFilter
       
        If bln列表 And bln卡片 Then
            chkAll.Enabled = True
        End If
        
        If mParams.blnFilter Then Call zlGetMediNumNew
    End If
    
    Call AviShow(Me, False)
    Exit Sub
errHandle:
    Call AviShow(Me, False)
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlGetMediNum(ByVal dateBegin As Date, ByVal dateEnd As Date, ByVal str病区ids As String, ByVal int操作状态 As Integer)
    Dim rstemp As Recordset
    Dim strTmp As String
    
    On Error GoTo errHandle

    strTmp = " Select distinct a.瓶签号, c.No, c.药品id, f.名称" & vbNewLine & _
        "              From 输液配药记录 A, 输液配药内容 B, 药品收发记录 C, 药品特性 D, 药品规格 E, 收费项目目录 F, 病人医嘱记录 M, 住院费用记录 N" & vbNewLine & _
        "              Where a.执行时间 Between [1] And [2] And a.Id = b.记录id And b.收发id = c.Id And" & vbNewLine & _
        "                    c.药品id = e.药品id And d.溶媒 <> 1 And a.部门id = [3] And e.药名id = d.药名id And m.相关id = a.医嘱id And m.Id = n.医嘱序号 And" & vbNewLine & _
        "                    c.费用id = n.Id And e.药品id = f.Id And a.病人病区id In (Select Column_Value From Table(Cast(f_Str2list([4]) As Zltools.t_Strlist))) And 操作状态 = [5] "
    strTmp = strTmp & " Union All " & Replace(strTmp, "住院费用记录", "门诊费用记录")
    
    gstrSQL = "Select 药品id, 名称, 数量" & vbNewLine & _
        "From (Select 药品id, 名称, Count(瓶签号) 数量" & vbNewLine & _
        "       From (" & strTmp & ")" & vbNewLine & _
        "       Group By 药品id, 名称" & vbNewLine & _
        "       Order By 数量 Desc) where rownum<4 "
        
       
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取药品信息", dateBegin, dateEnd, mParams.lng配置中心, str病区ids, int操作状态)
        
    cboMedi.Clear
    Me.cboMedi.AddItem "<全部>"
    Do While Not rstemp.EOF
        Me.cboMedi.AddItem rstemp!名称 & IIf(mParams.blnFilter, "", "(" & rstemp!数量 & ")")
        Me.cboMedi.ItemData(Me.cboMedi.ListCount - 1) = rstemp!药品ID
        
'        mParams.str常用药品 = IIf(mParams.str常用药品 = "", "", mParams.str常用药品 & ",") & rsTemp!药品ID
        rstemp.MoveNext
    Loop
    Me.cboMedi.Text = "<全部>"
        
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub zlGetMediNumNew()
    Dim rsFilter As Recordset
    Dim lng药品id As Long
    Dim str药品名称 As String
    Dim lngCount As Long
    Dim lng配药id As Long
    
    On Error GoTo errHandle
    
    Set rsFilter = New ADODB.Recordset
    With rsFilter
        If .State = 1 Then .Close
        
        '该记录对应的输液配药记录信息
        .Fields.Append "药品id", adDouble, 18, adFldIsNullable
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    cboMedi.Clear
    Me.cboMedi.AddItem "<全部>", 0
    Me.cboMedi.Text = "<全部>"
     
    With mrsTrans
        .Filter = "溶媒<>1 And 是否常用=1"
        If .RecordCount = 0 Then Exit Sub
        
        .Sort = "药品id,配药id"
        
        Do While Not .EOF
            If lng配药id <> mrsTrans!配药id Then
                If lng药品id <> mrsTrans!药品ID Then
                    If lng药品id <> 0 Then
                        rsFilter.AddNew
                        rsFilter!药品ID = lng药品id
                        rsFilter!药品名称 = str药品名称
                        rsFilter!数量 = lngCount
                        
                        rsFilter.Update
                    End If
                
                    lng药品id = mrsTrans!药品ID
                    str药品名称 = mrsTrans!药品名称
                    lngCount = 1
                Else
                    lngCount = lngCount + 1
                End If
            End If
            
            lng配药id = mrsTrans!配药id
                        
            .MoveNext
            
            If .EOF Then
                rsFilter.AddNew
                rsFilter!药品ID = lng药品id
                rsFilter!药品名称 = str药品名称
                rsFilter!数量 = lngCount
                
                rsFilter.Update
            End If
        Loop
    End With
    
    
    
    rsFilter.Filter = ""
    rsFilter.Sort = "数量 Desc"
    Do While Not rsFilter.EOF
        Me.cboMedi.AddItem rsFilter!药品名称 & "(" & rsFilter!数量 & ")"
        Me.cboMedi.ItemData(Me.cboMedi.NewIndex) = rsFilter!药品ID
        
'        mParams.str常用药品 = IIf(mParams.str常用药品 = "", "", mParams.str常用药品 & ",") & rsTemp!药品ID
        rsFilter.MoveNext
    Loop
       
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub SelectBatch(ByVal lngType As Long, ByVal index As Long)
    Dim lngRow As Long
    Dim strCompare As String
    Dim str配药id As String
    Dim i As Integer
    Dim str打包id As String
    Dim intFirst As Integer
    Dim datCur As Date
    Dim lng配药id As Long
    
    With vsfTrans
        If tbcDetail.Item(mDetailType.输液单卡片).Selected Then
            str配药id = mfrmPIVCard.ChooseOne

            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("配药ID")) = str配药id Then
                    .Row = i
                    Exit For
                End If
            Next
        End If
        
        str配药id = ";"
        
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, .ColIndex("配药ID")) = "" Then Exit Sub
        
        .Redraw = flexRDNone
        
        Select Case lngType
            Case conMenu_Oper_Select_SelRow
                '选择当前行
                str配药id = str配药id & .TextMatrix(.Row, .ColIndex("配药ID")) & ";"
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                        If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = 0 Then
                            .TextMatrix(lngRow, .ColIndex("选择")) = -1
                        End If
                        str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Select_SelBatch
                '选择当前批次
                strCompare = .TextMatrix(.Row, .ColIndex("配药批次"))
                
                If strCompare <> "" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                            If .TextMatrix(lngRow, .ColIndex("配药批次")) = strCompare Then
                                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = 0 Then
                                    .TextMatrix(lngRow, .ColIndex("选择")) = -1
                                End If
                                str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                            End If
                        End If
                    Next
                End If
            Case conMenu_Oper_Select_SelDept, conMenu_Oper_Select_CancleSelDept
                '选择当前病区
                strCompare = .TextMatrix(.Row, .ColIndex("病区"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("病区")) = strCompare Then
                            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(lngType = conMenu_Oper_Select_SelDept, -1, 0)
                            str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                        End If
                    End If
                Next
                
            Case conMenu_Oper_Select_SelPati, conMenu_Oper_Select_CancleSelPati
                '选择当前病人
                strCompare = .TextMatrix(.Row, .ColIndex("病区")) & .TextMatrix(.Row, .ColIndex("科室")) & .TextMatrix(.Row, .ColIndex("姓名")) & .TextMatrix(.Row, .ColIndex("床号"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("病区")) & .TextMatrix(lngRow, .ColIndex("科室")) & .TextMatrix(lngRow, .ColIndex("姓名")) & .TextMatrix(lngRow, .ColIndex("床号")) = strCompare Then
                            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(lngType = conMenu_Oper_Select_SelPati, -1, 0)
                            str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                        End If
                    End If
                Next
            Case conMenu_Oper_Select_SelSendNo
                '选择当前摆药单
                If mcondition.strTransStep = M_STR_CALSS_PREPARE Then Exit Sub
                strCompare = .TextMatrix(.Row, .ColIndex("摆药单号"))
                
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                        If .TextMatrix(lngRow, .ColIndex("摆药单号")) = strCompare Then
                            If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = 0 Then
                                .TextMatrix(lngRow, .ColIndex("选择")) = -1
                            End If
                            str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                        End If
                    End If
                Next
            Case conMenu_Oper_Select_SelMed
            '选择所有的抗菌药物
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("抗菌药物")) = 1 Then
                        If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = "0" Then
                            .TextMatrix(lngRow, .ColIndex("选择")) = -1
                        End If
                        str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Select_SelAll
                '选择所有行
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                        If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = "0" Then
                            .TextMatrix(lngRow, .ColIndex("选择")) = -1
                        End If
                        str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                    End If
                Next
            Case conMenu_Oper_Bag_Batch
                strCompare = .TextMatrix(.Row, .ColIndex("配药批次"))
                datCur = Sys.Currentdate
                
                If strCompare <> "" Then
                    For lngRow = 1 To .rows - 1
                        intFirst = 0
                        If .TextMatrix(lngRow, .ColIndex("配药ID")) <> "" Then
                            If .TextMatrix(lngRow, .ColIndex("配药批次")) = strCompare Then
                                If InStr("|" & str打包id, "|" & .TextMatrix(lngRow, .ColIndex("配药ID"))) < 1 Then
                                    str打包id = IIf(str打包id = "", "", str打包id & "|") & .TextMatrix(lngRow, .ColIndex("配药ID")) & ",2"
                                End If
                                
                                '更新配液(打包)图标
                                .Col = .ColIndex("打包")
                                .Cell(flexcpPicture, lngRow, .ColIndex("打包")) = picPacker(2).Picture
                                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("打包")) = flexPicAlignCenterCenter
                                .TextMatrix(lngRow, .ColIndex("是否打包")) = 2
                                
                                mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(.Row, .ColIndex("配药ID")))
                                Do While Not mrsTrans.EOF
                                    intFirst = intFirst + 1
                                    mrsTrans!是否打包 = Val(.TextMatrix(.Row, .ColIndex("是否打包")))
                                    
                                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE And intFirst = 1 And .TextMatrix(.Row, .ColIndex("是否打包")) > 0 Then
                                        mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                                    Else
                                        If IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")) Then
                                            mintCountPack = mintCountPack - 1
                                        End If
                                    End If
                                    
                                    mrsTrans!打包时间 = IIf(.TextMatrix(.Row, .ColIndex("是否打包")) = 0, "", datCur)
                                    mrsTrans.Update
                                    mrsTrans.MoveNext
                                Loop
                                
                                mfrmPIVCard.PackCard Val(.TextMatrix(lngRow, .ColIndex("配药ID"))), 2
                            End If
                        End If
                    Next
                End If
                
                Call GetCount
                
                If str打包id <> "" Then
                    gstrSQL = "Zl_输液配药记录_打包("
                    '配药ID,打包
                    gstrSQL = gstrSQL & "'" & str打包id & "'"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打包设置")
                End If
            Case conMenu_Oper_Bag_All
                datCur = Sys.Currentdate
                
                '更新打包图标
                With Me.vsfTrans
                    .Col = .ColIndex("打包")
                    .Cell(flexcpPicture, 1, .ColIndex("打包"), .rows - 1, .ColIndex("打包")) = picPacker(2).Picture
                    .Cell(flexcpPictureAlignment, 1, .ColIndex("打包"), .rows - 1, .ColIndex("打包")) = flexPicAlignCenterCenter
                    .Cell(flexcpText, 1, .ColIndex("是否打包"), .rows - 1, .ColIndex("是否打包")) = 2
                
                    For lngRow = 1 To .rows - 1
                        If lng配药id <> Val(.TextMatrix(lngRow, .ColIndex("配药id"))) And Val(.TextMatrix(lngRow, .ColIndex("配药id"))) <> 0 Then
                            lng配药id = Val(.TextMatrix(lngRow, .ColIndex("配药id")))
                            If InStr("|" & str打包id, "|" & .TextMatrix(lngRow, .ColIndex("配药ID"))) < 1 Then
                                str打包id = IIf(str打包id = "", "", str打包id & "|") & .TextMatrix(lngRow, .ColIndex("配药ID")) & ",2"
                            End If
                            
                            mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                            
                            '改变内部数据集的值
                            mrsTrans!是否打包 = Val(.TextMatrix(lngRow, .ColIndex("是否打包")))
                            
                            If mcondition.strTransStep = M_STR_CALSS_DOSAGE And .TextMatrix(lngRow, .ColIndex("是否打包")) > 0 Then
                                mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                            Else
                                If IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")) Then
                                    mintCountPack = mintCountPack - 1
                                End If
                            End If
                            
                            mrsTrans!打包时间 = IIf(.TextMatrix(lngRow, .ColIndex("是否打包")) = 0, "", datCur)
                            mrsTrans.Update
                            mrsTrans.MoveNext
                        End If
                    Next
                    
                    Call GetCount
                
                    If str打包id <> "" Then
                        gstrSQL = "Zl_输液配药记录_打包("
                        '配药ID,打包
                        gstrSQL = gstrSQL & "'" & str打包id & "'"
                        gstrSQL = gstrSQL & ")"
                        Call zlDatabase.ExecuteProcedure(gstrSQL, "打包设置")
                    End If
                End With
        End Select
        
        Call mfrmPIVCard.BatchChoose(str配药id)
            
        .Redraw = flexRDDirect
        
        '更新数据集
        If lngType = conMenu_Oper_Select_SelRow Then
            Call UpdateExeSign(Val(.TextMatrix(.Row, .ColIndex("配药ID"))), IIf(Val(.TextMatrix(.Row, .ColIndex("选择"))) = -1, 1, 0))
        ElseIf lngType = conMenu_Oper_Select_SelAll Then
            Call UpdateExeSign(0, 1)
        Else
            Call UpdateExeSign(-1, index)
        End If
    End With
End Sub

Private Sub SetCommand()
    '设置菜单及快捷按钮
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    
    '取消按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Cancel, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Cancel, , True)

    If Not cbrMenu Is Nothing Then
        cbrMenu.Visible = False
        If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then cbrMenu.Visible = (mPrives.bln取消审核 And mParams.bln审核)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then cbrMenu.Visible = mPrives.bln取消摆药
        If mcondition.strTransStep = M_STR_CALSS_SEND Then cbrMenu.Visible = mPrives.bln取消配药
        If mcondition.strTransStep = M_STR_CALSS_SENDED Then cbrMenu.Visible = mPrives.bln取消发送
        
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            cbrMenu.Caption = "取消摆药(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            cbrMenu.Caption = "取消配药(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            cbrMenu.Caption = "取消发送(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            cbrMenu.Caption = "取消审核(&C)"
        End If
    End If
    
    If Not cbrControl Is Nothing Then
        cbrControl.Visible = False
        If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then cbrControl.Visible = (mPrives.bln取消审核 And mParams.bln审核)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then cbrControl.Visible = mPrives.bln取消摆药
        If mcondition.strTransStep = M_STR_CALSS_SEND Then cbrControl.Visible = mPrives.bln取消配药
        If mcondition.strTransStep = M_STR_CALSS_SENDED Then cbrControl.Visible = mPrives.bln取消发送
        
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            cbrControl.Caption = "取消摆药(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            cbrControl.Caption = "取消配药(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            cbrControl.Caption = "取消发送(&C)"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            cbrControl.Caption = "取消审核(&C)"
        End If
    End If

    '打印瓶签
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, , True)
    
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mParams.bln瓶签手工打印 = True And (mcondition.strTransStep <> M_STR_CALSS_AUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT))
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mParams.bln瓶签手工打印 = True And (mcondition.strTransStep <> M_STR_CALSS_AUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT))
    
    
    '打印摆药单
'    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, , True)
'    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, , True)

'    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (Val(cmdNext.Tag) = mType.输液单 And mcondition.strTransStep = M_STR_CALSS_PREPARE)
'    If Not cbrControl Is Nothing Then cbrControl.Visible = (Val(cmdNext.Tag) = mType.输液单 And mcondition.strTransStep = M_STR_CALSS_PREPARE)
    
    '审核,拒绝医嘱按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Approve, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Approve, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_AUDIT And mPrives.bln核查确认 And mParams.bln审核)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_AUDIT And mPrives.bln核查确认 And mParams.bln审核)
    
    '排班设置
    Set cbrMenu = cbsMain.FindControl(, mconMenu_PlanPopup)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = mPrives.bln排班设置
    
    '锁定,解锁按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Lock, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Lock, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_UnLock, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_UnLock, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE)
    
    '调整批次按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Beach, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Beach, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)

    '确认调整按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, , True)
    
    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)
    
    '摆药按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Prepare, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Prepare, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_PREPARE And mPrives.bln摆药确认)

    '配药按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Dosage, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Dosage, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_DOSAGE And mPrives.bln配药确认)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_DOSAGE And mPrives.bln配药确认)
    
    '发送按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Send, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Send, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_SEND And mPrives.bln发送确认)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_SEND And mPrives.bln发送确认)
    
    '确认拒绝按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_REFUSETOSIGN And mPrives.bln确认拒绝)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_REFUSETOSIGN And mPrives.bln确认拒绝)
        
    '销帐按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_VERIFY And mPrives.bln销帐审核)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_VERIFY And mPrives.bln销帐审核)
    
    '删除按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Delete, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_PIVA_Delete, , True)

    If Not cbrMenu Is Nothing Then cbrMenu.Visible = (mcondition.strTransStep = M_STR_CALSS_INVALID)
    If Not cbrControl Is Nothing Then cbrControl.Visible = (mcondition.strTransStep = M_STR_CALSS_INVALID)
    
    '批量打包按钮
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButtonPopup, conMenu_Oper_Bag, , True)
    Set cbrControl = cbsMain(2).Controls.Find(xtpControlButtonPopup, conMenu_Oper_Bag, , True)
    
    If mParams.bln打包设置 Then
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = True
    Else
        If Not cbrMenu Is Nothing Then cbrMenu.Visible = False
    End If
End Sub

Private Sub SetTabColor(ByVal tbcObj As TabControl)
    '设置分页按钮的颜色
    Dim intTabIndex As Integer
    Dim strCount As String

    With tbcObj
        For intTabIndex = 0 To .ItemCount - 1
            If .Item(intTabIndex).Selected = True Then
                .Item(intTabIndex).Color = CSTCOLOR_COMMAND
            Else
                strCount = Mid(.Item(intTabIndex).Caption, InStr(1, .Item(intTabIndex).Caption, "(") + 1)
                strCount = Mid(strCount, 1, InStr(1, strCount, ")") - 1)
                If Val(strCount) > 0 Then
                    .Item(intTabIndex).Color = CSTCOLOR_RECORDS
                Else
                    .Item(intTabIndex).Color = CSTCOLOR_NORECORDS
                End If
            End If
        Next
    End With
End Sub

Private Sub SetListBar()
    '根据明细列表页面选择显示不同的选项条件
    Select Case tbcDetail.Selected.index
        Case mDetailType.输液单列表
            chkDept.Visible = False
            chkPack.Visible = False
            chkAll.Visible = True
            chkType(0).Visible = True
            chkType(1).Visible = True
            
            If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                chkSendType(0).Visible = True
                chkSendType(1).Visible = True
            Else
                chkSendType(0).Visible = False
                chkSendType(1).Visible = False
            End If
            
            Me.lblBatch.Visible = True
            Me.cboBatch.Visible = True
            Me.lblLevel.Visible = True
            Me.cboLevel.Visible = True
            
            Me.cboFrequency.Visible = True
            Me.lblFrequency.Visible = True
            
            lblMedi.Visible = True
            cboMedi.Visible = True
            
            Me.lblVolu.Visible = True
            
            chkPrint(0).Visible = True
            chkPrint(1).Visible = True
            
            chkChange(0).Visible = True
            chkChange(1).Visible = True
            
            chkSure(0).Visible = mcondition.strTransStep = M_STR_CALSS_PREPARE
            chkSure(1).Visible = mcondition.strTransStep = M_STR_CALSS_PREPARE
            
            optShowType(0).Visible = True
            optShowType(1).Visible = True
            
            lblDosType.Visible = True
            cboDosType.Visible = True
        Case mDetailType.药品汇总列表
            chkDept.Visible = True
            chkPack.Visible = True
            chkAll.Visible = False
            chkType(0).Visible = False
            chkType(1).Visible = False
            chkSendType(0).Visible = False
            chkSendType(1).Visible = False
            optShowType(0).Visible = False
            optShowType(1).Visible = False
            
            Me.lblBatch.Visible = False
            Me.cboBatch.Visible = False
            Me.lblLevel.Visible = False
            Me.cboLevel.Visible = False
            
            chkSure(0).Visible = False
            chkSure(1).Visible = False
            
            chkPrint(0).Visible = False
            chkPrint(1).Visible = False
            
            chkChange(0).Visible = False
            chkChange(1).Visible = False
            
            Me.cboFrequency.Visible = False
            Me.lblFrequency.Visible = False
            
            lblMedi.Visible = False
            cboMedi.Visible = False
            
            lblDosType.Visible = False
            cboDosType.Visible = False
            
            Me.lblVolu.Visible = False
            
            vsfTrans.Visible = False
            vsfSumDrug.Visible = True
    End Select
    
    chkSure(0).Left = IIf(chkType(1).Visible, chkType(1).Left + chkType(1).Width, chkAll.Left + chkAll.Width) + 200
    chkSure(1).Left = chkSure(0).Left + chkSure(0).Width + 50
    
    chkPrint(0).Left = IIf(chkSure(1).Visible, chkSure(1).Left + chkSure(1).Width, chkType(1).Left + chkType(1).Width) + 200
    chkPrint(1).Left = chkPrint(0).Left + chkPrint(0).Width + 50
    
    chkChange(0).Left = IIf(chkPrint(1).Visible, chkPrint(1).Left + chkPrint(1).Width, chkSure(1).Left + chkSure(1).Width) + 200
    chkChange(1).Left = chkChange(0).Left + chkChange(0).Width + 50
    
    chkSendType(0).Left = IIf(chkChange(1).Visible, chkChange(1).Left + chkChange(1).Width, chkPrint(1).Left + chkPrint(1).Width) + 200
    chkSendType(1).Left = chkSendType(0).Left + chkSendType(0).Width + 50
    
    optShowType(0).Left = IIf(chkSendType(1).Visible, chkSendType(1).Left + chkSendType(1).Width, chkChange(1).Left + chkChange(1).Width) + 200
    optShowType(1).Left = optShowType(0).Left + optShowType(0).Width + 50
    
End Sub

Private Sub BillPrint_Sum()
    '打印汇总单据
    Dim StrDate As String
    
    Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
        "部门=" & mcondition.lngCenterID, _
        "打印时间=" & StrDate, "PrintEmpty=0", 1)
End Sub

Private Sub SetSumDrugColHide()
    '设置汇总药品列表中的列隐藏属性
    With vsfSumDrug
        .ColHidden(.ColIndex("病区")) = (chkDept.Value = 0)
        .ColHidden(.ColIndex("打包")) = (chkPack.Value = 0)
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            .ColHidden(.ColIndex("库存数量")) = False
        Else
            .ColHidden(.ColIndex("库存数量")) = True
        End If
    End With
End Sub

Private Sub SetTransColHide()
    '设置输液单据表格列隐藏属性

    With vsfTrans
        .ColHidden(.ColIndex("销帐申请人")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        .ColHidden(.ColIndex("销帐申请时间")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        
        .ColHidden(.ColIndex("作废类型")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        .ColHidden(.ColIndex("销帐审核人")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        .ColHidden(.ColIndex("销帐审核时间")) = (mcondition.strTransStep <> M_STR_CALSS_INVALID)
        
        .ColHidden(.ColIndex("摆药单号")) = (mcondition.strTransStep <= M_STR_CALSS_PREPARE)
        .ColHidden(.ColIndex("锁")) = (mcondition.strTransStep <> M_STR_CALSS_PREPARE)
        
        
        .ColHidden(.ColIndex("变")) = (mcondition.strTransStep > M_STR_CALSS_PREPARE)
        .ColHidden(.ColIndex("医嘱")) = (mcondition.strTransStep > M_STR_CALSS_PREPARE)
        
        .ColHidden(.ColIndex("拒收原因")) = (mcondition.strTransStep <> M_STR_CALSS_REFUSETOSIGN)
        .ColHidden(.ColIndex("调")) = (mcondition.strTransStep <> M_STR_CALSS_PREPARE)
        
        .ColHidden(.ColIndex("审")) = (mcondition.strTransStep <> M_STR_CALSS_VERIFY)
        .ColHidden(.ColIndex("选择")) = (mcondition.strTransStep = M_STR_CALSS_VERIFY)
        
        If optShowType(0).Value = True Then
            .ColHidden(.ColIndex("摆药人")) = True
            .ColHidden(.ColIndex("摆药时间")) = True
            .ColHidden(.ColIndex("配药人")) = True
            .ColHidden(.ColIndex("配药时间")) = True
            .ColHidden(.ColIndex("发送人")) = True
            .ColHidden(.ColIndex("发送时间")) = True
            .ColHidden(.ColIndex("医嘱发送时间")) = True
        Else
            .ColHidden(.ColIndex("医嘱发送时间")) = False
            .ColHidden(.ColIndex("摆药人")) = (mcondition.strTransStep <> M_STR_CALSS_DOSAGE)
            .ColHidden(.ColIndex("摆药时间")) = (mcondition.strTransStep <> M_STR_CALSS_DOSAGE)
            .ColHidden(.ColIndex("配药人")) = (mcondition.strTransStep <> M_STR_CALSS_SEND)
            .ColHidden(.ColIndex("配药时间")) = (mcondition.strTransStep <> M_STR_CALSS_SEND)
            .ColHidden(.ColIndex("发送人")) = (mcondition.strTransStep <> M_STR_CALSS_SENDED)
            .ColHidden(.ColIndex("发送时间")) = (mcondition.strTransStep <> M_STR_CALSS_SENDED)
        End If
    End With
End Sub

Private Sub ShowMedicalRecord()
    '【功能】:查阅当前病人的电子病案

    With vsfMedis
        '检查
        If .Row < 1 Then Exit Sub
        
        '调用电子病案查阅接口
        If Not mobjCISJOB Is Nothing Then
            On Error Resume Next
            Call mobjCISJOB.ShowArchive(Me, Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("病人id"))), Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("主页id"))))
            err.Clear: On Error GoTo 0
        End If
        
    End With
End Sub

Private Sub ShowComment(ByVal intTab As Integer, ByVal strStep As String)
    '显示当前流程的提示信息
    
    lblHelp.Caption = ""
    
    If intTab = mDetailType.输液单列表 Then
        If strStep = M_STR_CALSS_PREPARE Then
            If mPrives.bln摆药确认 = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "你没有摆药确认的权限"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认"
                If mParams.int摆药后打印 = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认后提示打印摆药单"
                ElseIf mParams.int摆药后打印 = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认后自动打印摆药单"
                Else
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认后不打印摆药单"
                End If
                If mParams.int瓶签摆药后打印 = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认后提示打印标签"
                ElseIf mParams.int瓶签摆药后打印 = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "摆药确认后自动打印标签"
                End If
            End If
            If mParams.bln批次设置 = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "允许调整批次"
            End If
            If mParams.bln打包设置 = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "允许调整打包状态"
            End If
        ElseIf strStep = M_STR_CALSS_DOSAGE Then
            If mPrives.bln配药确认 = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "你没有配药确认的权限"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "配药确认"
                If mParams.int瓶签配药后打印 = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "配药确认后提示打印标签"
                ElseIf mParams.int瓶签配药后打印 = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "配药确认后自动打印标签"
                End If
            End If
            If mParams.bln打包设置 = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "允许调整打包状态"
            End If
        ElseIf strStep = M_STR_CALSS_SEND Then
            If mPrives.bln发送确认 = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "你没有发送确认的权限"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "发送确认"
                If mParams.int发送后打印 = 0 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "发送确认后提示打印发送单"
                ElseIf mParams.int发送后打印 = 1 Then
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "发送确认后自动打印发送单"
                Else
                    lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "发送确认后不打印发送单"
                End If
            End If
        ElseIf strStep = M_STR_CALSS_SENDED Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "已发送查看"
        ElseIf strStep = M_STR_CALSS_SIGNED Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "已签收查看"
        ElseIf strStep = M_STR_CALSS_REFUSETOSIGN Then
            If mPrives.bln确认拒绝 = False Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "你没有拒绝确认的权限"
            Else
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "拒绝签收申请查看"
            End If
        ElseIf strStep = M_STR_CALSS_VERIFY Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "销帐申请查看"
            If mPrives.bln销帐审核 = True Then
                lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "允许销帐审核"
            End If
        ElseIf strStep = M_STR_CALSS_INVALID Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "已作废查看"
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "显示已销帐审核的输液单"
        ElseIf strStep = M_STR_CALSS_DEVICERETURN Then
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "医嘱回退查看"
            lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "显示因医嘱回退作废的输液单"
        ElseIf strStep = M_STR_CALSS_AUDIT Then
            lblHelp.Caption = "对未审核的医嘱进行审核操作"
        ElseIf strStep = M_STR_CALSS_PASSEDAUDIT Then
            lblHelp.Caption = "对已通过审核的医嘱进行取消操作"
        ElseIf strStep = M_STR_CALSS_FAILAUDIT Then
            lblHelp.Caption = "对未通过审核的医嘱进行取消操作"
        End If
        
        mfrmPIVCard.LoadHelp lblHelp.Caption
    ElseIf intTab = mDetailType.药品汇总列表 Then
        lblHelp.Caption = IIf(lblHelp.Caption = "", "", lblHelp.Caption & ";") & "药品汇总查看;红色表示缺药"
    End If
    
    
End Sub

Private Sub ShowSumDrug()
    '根据已选择的输液单，汇总药品
    Dim lngRow As Long
    Dim strCurr As String
    Dim dblSum As Double
    Dim dblSendSum As Double
    Dim lng收发ID As Long
    
    With vsfSumDrug
        .rows = 1
        .rows = 2
        
        .MergeCells = flexMergeNever
        
        If mrsTrans Is Nothing Then Exit Sub
        
        mrsTrans.Filter = "执行标志=1"
        
        If mrsTrans.RecordCount = 0 Then
            mrsTrans.Filter = ""
            Exit Sub
        End If
        
        .Redraw = flexRDNone
        .rows = 1
        
        If chkDept.Value = 1 And chkPack.Value = 1 Then
            mrsTrans.Sort = "病区,是否打包,药品编码名称,批次,收发ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!病区 & mrsTrans!是否打包 & mrsTrans!药品编码名称 & mrsTrans!批次 Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!病区 & mrsTrans!是否打包 & mrsTrans!药品编码名称 & mrsTrans!批次
                    dblSum = zlStr.nvl(mrsTrans!数量, 0)
                    
                    lng收发ID = mrsTrans!收发ID
                    dblSendSum = mrsTrans!发药数量
                    
                    .TextMatrix(lngRow, .ColIndex("病区")) = mrsTrans!病区
                    .TextMatrix(lngRow, .ColIndex("是否打包")) = mrsTrans!是否打包
                    .TextMatrix(lngRow, .ColIndex("药品名称")) = mrsTrans!药品编码名称
                    .TextMatrix(lngRow, .ColIndex("商品名")) = mrsTrans!商品名
                    .TextMatrix(lngRow, .ColIndex("英文名")) = mrsTrans!英文名
                    .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.nvl(mrsTrans!规格)
                    .TextMatrix(lngRow, .ColIndex("产地")) = mrsTrans!产地
                    .TextMatrix(lngRow, .ColIndex("批号")) = mrsTrans!批号
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("库存数量")) = FormatEx(mrsTrans!库存数量, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!数量, 0)
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    
                    If lng收发ID <> mrsTrans!收发ID Then
                        lng收发ID = mrsTrans!收发ID
                        dblSendSum = dblSendSum + mrsTrans!发药数量
                        .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                        .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
            
            mrsTrans.Filter = ""
        ElseIf chkDept.Value = 1 Then
            mrsTrans.Sort = "病区,药品编码名称,批次,收发ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!病区 & mrsTrans!药品编码名称 & mrsTrans!批次 Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!病区 & mrsTrans!药品编码名称 & mrsTrans!批次
                    dblSum = zlStr.nvl(mrsTrans!数量, 0)
                    
                    lng收发ID = mrsTrans!收发ID
                    dblSendSum = mrsTrans!发药数量
                    
                    .TextMatrix(lngRow, .ColIndex("病区")) = mrsTrans!病区
                    .TextMatrix(lngRow, .ColIndex("药品名称")) = mrsTrans!药品编码名称
                    .TextMatrix(lngRow, .ColIndex("商品名")) = mrsTrans!商品名
                    .TextMatrix(lngRow, .ColIndex("英文名")) = mrsTrans!英文名
                    .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.nvl(mrsTrans!规格)
                    .TextMatrix(lngRow, .ColIndex("产地")) = mrsTrans!产地
                    .TextMatrix(lngRow, .ColIndex("批号")) = mrsTrans!批号
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("库存数量")) = FormatEx(mrsTrans!库存数量, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!数量, 0)
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    
                    If lng收发ID <> mrsTrans!收发ID Then
                        lng收发ID = mrsTrans!收发ID
                        dblSendSum = dblSendSum + mrsTrans!发药数量
                        .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                        .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        ElseIf chkPack.Value = 1 Then
            mrsTrans.Sort = "是否打包,药品编码名称,批次,收发ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!是否打包 & mrsTrans!药品编码名称 & mrsTrans!批次 Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!是否打包 & mrsTrans!药品编码名称 & mrsTrans!批次
                    dblSum = zlStr.nvl(mrsTrans!数量, 0)
                    
                    lng收发ID = mrsTrans!收发ID
                    dblSendSum = mrsTrans!发药数量
                    
                    .TextMatrix(lngRow, .ColIndex("是否打包")) = mrsTrans!是否打包
                    .TextMatrix(lngRow, .ColIndex("药品名称")) = mrsTrans!药品编码名称
                    .TextMatrix(lngRow, .ColIndex("商品名")) = mrsTrans!商品名
                    .TextMatrix(lngRow, .ColIndex("英文名")) = mrsTrans!英文名
                    .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.nvl(mrsTrans!规格)
                    .TextMatrix(lngRow, .ColIndex("产地")) = mrsTrans!产地
                    .TextMatrix(lngRow, .ColIndex("批号")) = mrsTrans!批号
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("库存数量")) = FormatEx(mrsTrans!库存数量, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!数量, 0)
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    
                    If lng收发ID <> mrsTrans!收发ID Then
                        lng收发ID = mrsTrans!收发ID
                        dblSendSum = dblSendSum + mrsTrans!发药数量
                        .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                        .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        Else
            mrsTrans.Sort = "药品编码名称,批次,收发ID"
            
            Do While Not mrsTrans.EOF
                If strCurr <> mrsTrans!药品编码名称 & mrsTrans!批次 Then
                    lngRow = lngRow + 1
                    .rows = .rows + 1
                    
                    strCurr = mrsTrans!药品编码名称 & mrsTrans!批次
                    dblSum = zlStr.nvl(mrsTrans!数量, 0)
                    
                    lng收发ID = mrsTrans!收发ID
                    dblSendSum = mrsTrans!发药数量

                    .TextMatrix(lngRow, .ColIndex("药品名称")) = mrsTrans!药品编码名称
                    .TextMatrix(lngRow, .ColIndex("商品名")) = mrsTrans!商品名
                    .TextMatrix(lngRow, .ColIndex("英文名")) = mrsTrans!英文名
                    .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.nvl(mrsTrans!规格)
                    .TextMatrix(lngRow, .ColIndex("产地")) = mrsTrans!产地
                    .TextMatrix(lngRow, .ColIndex("批号")) = mrsTrans!批号
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("库存数量")) = FormatEx(mrsTrans!库存数量, 2) & mrsTrans!单位
                    .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                Else
                    dblSum = dblSum + zlStr.nvl(mrsTrans!数量, 0)
                    .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(dblSum, 2) & mrsTrans!单位
                    
                    If lng收发ID <> mrsTrans!收发ID Then
                        lng收发ID = mrsTrans!收发ID
                        dblSendSum = dblSendSum + mrsTrans!发药数量
                        .TextMatrix(lngRow, .ColIndex("发药数量")) = FormatEx(dblSendSum, 2) & mrsTrans!单位
                        .TextMatrix(lngRow, .ColIndex("缺药标志")) = IIf(dblSendSum > mrsTrans!库存数量, 1, 0)
                    End If
                End If
                
                mrsTrans.MoveNext
            Loop
        End If
        
        For lngRow = 1 To .rows - 1
            '标识缺药药品
            If .TextMatrix(lngRow, .ColIndex("药品名称")) <> "" Then
                If .TextMatrix(lngRow, .ColIndex("缺药标志")) = 1 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
                End If
            End If
            
            '更新配液(打包)图标
            If chkPack.Value = 1 Then
                If Val(.TextMatrix(lngRow, .ColIndex("是否打包"))) > 0 Then
'                    .Row = lngRow
'                    .Col = .ColIndex("打包")
'                    .CellPicture = picPacker(1).Picture
'                    .CellPictureAlignment = flexPicAlignCenterCenter
                    .Cell(flexcpPicture, lngRow, .ColIndex("打包"), lngRow, .ColIndex("打包")) = picPacker(Val(.TextMatrix(lngRow, .ColIndex("是否打包")))).Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("打包"), lngRow, .ColIndex("打包")) = flexPicAlignCenterCenter
                End If
            End If
        Next
        
        '合并病区
        If chkDept.Value = 1 Then
            .MergeCells = flexMergeRestrictRows
            .MergeCol(.ColIndex("病区")) = True
        End If
        
        Call SetSumDrugColHide
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Function ShowTrans(ByVal index As Long) As Boolean
    Dim lngRow As Long
    Dim lng配药id As Long
    Dim strMaxNo As String
    Dim j As Integer
    Dim dteCur As Date
    Dim intBlackColor As Integer
    Dim lng病人id As Long
    Dim lng药品id As Long
    Dim LngID As Long
    Dim dateCurrent As Date
    Dim strSort As String
    Dim lngNum As Long
'    chkAll.Value = 0
    
    vsfTrans.rows = 1
    vsfTrans.rows = 2
    intBlackColor = 2
'    vsfDrug.rows = 1
'    vsfDrug.rows = 2

    mlng未扫描 = 0
'
    dteCur = Sys.Currentdate
    If mrsTrans Is Nothing Then Exit Function
    If mrsTrans.RecordCount = 0 Then Exit Function
    mrsTrans.MoveFirst
    
    If Val(vsfTrans.Tag) = 1 And mParams.blnByMedi = True Then
        Call MediSort
    Else
        strSort = mParams.strSort
        If InStr(1, strSort, "配药批次") > 0 Then
            strSort = Replace(strSort, "配药批次", "配药批次,优先级")
        End If
        
        If InStr(1, strSort, "床号") > 0 Then
            strSort = Replace(strSort, "床号", "编码,床号排序")
        End If
        
        mrsTrans.Sort = IIf(strSort <> "", strSort & ",配药id", "配药id")
    End If
    
    If mrsTrans.RecordCount = 0 Then Exit Function
    
    dateCurrent = Sys.Currentdate
    
    With Me.vsfDept(0)
        For lngNum = 1 To .rows - 1
            If Val(.TextMatrix(lngNum, .ColIndex("病区id"))) > 0 And Val(.TextMatrix(lngNum, .ColIndex("选择"))) = -1 Then
                .TextMatrix(lngNum, .ColIndex("数量")) = 0
            End If
        Next
    End With
    
    With vsfTrans
        .MergeCells = flexMergeFree
        .Redraw = flexRDNone
        .rows = 1
'        .rows = mrsTrans.RecordCount + 1
        lngRow = 1
        Do While Not mrsTrans.EOF
            .rows = .rows + 1
            If mstrFilter <> "" Then
                If lng配药id <> LngID Or LngID = 0 Then
                    LngID = Split(mstrFilter, ",")(0)
                    mrsTrans.Filter = "配药id=" & Split(mstrFilter, ",")(0)
                    mstrFilter = Mid(mstrFilter, Len(Split(mstrFilter, ",")(0)) + 2)
                End If
            End If
            
            .MergeCol(.ColIndex("姓名")) = True
            .MergeRow(lngRow) = False
            
            If lng病人id <> mrsTrans!病人ID Then
                lng病人id = mrsTrans!病人ID
                If intBlackColor = 2 Then
                    intBlackColor = 1
                ElseIf intBlackColor = 1 Then
                    intBlackColor = 2
                End If
            End If
            
            If lng配药id <> mrsTrans!配药id Then
                With Me.vsfDept(0)
                    For lngNum = 1 To .rows - 1
                        If mrsTrans!病区 = Mid(.TextMatrix(lngNum, .ColIndex("病区")), InStr(1, .TextMatrix(lngNum, .ColIndex("病区")), "]") + 1) And Val(.TextMatrix(lngNum, .ColIndex("选择"))) = -1 Then
                            .TextMatrix(lngNum, .ColIndex("数量")) = Val(.TextMatrix(lngNum, .ColIndex("数量"))) + 1
                        End If
                    Next
                End With
                mlngNum = mlngNum + 1
                mlng未扫描 = mlng未扫描 + 1
                If lng配药id <> 0 Then
                    .rows = .rows + 1
                    .RowHidden(lngRow) = True
                    For j = 0 To .Cols - 1
                        .TextMatrix(lngRow, j) = "00"
                    Next
                    lngRow = lngRow + 1
                End If
                lng配药id = mrsTrans!配药id
            Else
                .MergeCol(.ColIndex("选择")) = True
                .MergeCol(.ColIndex("病区")) = True
                .MergeCol(.ColIndex("科室")) = True
                .MergeCol(.ColIndex("姓名")) = True
                .MergeCol(.ColIndex("性别")) = True
                .MergeCol(.ColIndex("年龄")) = True
                .MergeCol(.ColIndex("床号")) = True
                .MergeCol(.ColIndex("住院号")) = True
                .MergeCol(.ColIndex("配药批次")) = True
                .MergeCol(.ColIndex("作废类型")) = True
                .MergeCol(.ColIndex("执行时间")) = True
                .MergeCol(.ColIndex("瓶签号")) = True
                .MergeCol(.ColIndex("优先级")) = True
                .MergeCol(.ColIndex("核查人")) = True
                .MergeCol(.ColIndex("核查时间")) = True
                .MergeCol(.ColIndex("摆药人")) = True
                .MergeCol(.ColIndex("摆药时间")) = True
                .MergeCol(.ColIndex("摆药单号")) = True
                .MergeCol(.ColIndex("配药人")) = True
                .MergeCol(.ColIndex("配药时间")) = True
                .MergeCol(.ColIndex("发送人")) = True
                .MergeCol(.ColIndex("发送时间")) = True
                .MergeCol(.ColIndex("销帐申请人")) = True
                .MergeCol(.ColIndex("销帐申请时间")) = True
                .MergeCol(.ColIndex("销帐审核人")) = True
                .MergeCol(.ColIndex("销帐审核时间")) = True
                .MergeCol(.ColIndex("医嘱发送时间")) = True
                .MergeCol(.ColIndex("是否打包")) = True
                .MergeCol(.ColIndex("打印标志")) = True
                .MergeCol(.ColIndex("配药ID")) = True
                .MergeCol(.ColIndex("抗菌药物")) = True
                .MergeCol(.ColIndex("打印")) = True
                .MergeCol(.ColIndex("医嘱")) = True
                .MergeCol(.ColIndex("打包")) = True
                .MergeCol(.ColIndex("锁")) = True
                .MergeCol(.ColIndex("变")) = True
                .MergeCol(.ColIndex("背景号")) = True
                .MergeCol(.ColIndex("拒收原因")) = True
                .MergeCol(.ColIndex("调")) = True
                .MergeCol(.ColIndex("审")) = True
                .MergeCol(.ColIndex("执行频次")) = True
                .MergeCol(.ColIndex("销帐原因")) = True
            End If
            
            
            
                lng药品id = mrsTrans!药品ID
                .TextMatrix(lngRow, .ColIndex("背景号")) = intBlackColor
                .TextMatrix(lngRow, 1) = lngRow
                
                .TextMatrix(lngRow, .ColIndex("选择")) = IIf(mrsTrans!执行标志 = 1, -1, "")
                
                If mParams.intAutoSelect = 1 Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = IIf(InStr(1, mstrLastLabel, mrsTrans!瓶签号) > 0, -1, "")
                End If
                
                .TextMatrix(lngRow, .ColIndex("作废类型")) = IIf(IsNull(mrsTrans!作废类型), "", mrsTrans!作废类型)
                
                .TextMatrix(lngRow, .ColIndex("打印")) = " "
                .TextMatrix(lngRow, .ColIndex("医嘱")) = " "
                .TextMatrix(lngRow, .ColIndex("打包")) = " "
                .TextMatrix(lngRow, .ColIndex("锁")) = " "
                .TextMatrix(lngRow, .ColIndex("变")) = " "
                .TextMatrix(lngRow, .ColIndex("调")) = " "
                .TextMatrix(lngRow, .ColIndex("审")) = " "
                .TextMatrix(lngRow, .ColIndex("标志")) = 0
                .TextMatrix(lngRow, .ColIndex("配药批次")) = IIf(zlStr.nvl(mrsTrans!配药批次) = "", " ", zlStr.nvl(mrsTrans!配药批次))
                .TextMatrix(lngRow, .ColIndex("原批次")) = mrsTrans!配药批次
                .TextMatrix(lngRow, .ColIndex("病区")) = mrsTrans!病区
                .TextMatrix(lngRow, .ColIndex("科室")) = mrsTrans!科室
                .TextMatrix(lngRow, .ColIndex("姓名")) = mrsTrans!姓名
                .TextMatrix(lngRow, .ColIndex("床号")) = IIf(zlStr.nvl(mrsTrans!床号) = "", "<空>", mrsTrans!床号)
                .TextMatrix(lngRow, .ColIndex("住院号")) = mrsTrans!住院号
                .TextMatrix(lngRow, .ColIndex("性别")) = mrsTrans!性别
                .TextMatrix(lngRow, .ColIndex("年龄")) = mrsTrans!年龄
                .TextMatrix(lngRow, .ColIndex("执行时间")) = mrsTrans!执行时间
                .TextMatrix(lngRow, .ColIndex("瓶签号")) = mrsTrans!瓶签号
                .TextMatrix(lngRow, .ColIndex("病人id")) = mrsTrans!病人ID
                .TextMatrix(lngRow, .ColIndex("主页id")) = mrsTrans!主页id
                .TextMatrix(lngRow, .ColIndex("优先级")) = Val(zlStr.nvl(mrsTrans!优先级))
                .TextMatrix(lngRow, .ColIndex("是否锁定")) = zlStr.nvl(mrsTrans!是否锁定, 0)
                                                    
                .Cell(flexcpPicture, lngRow, .ColIndex("锁"), lngRow, .ColIndex("锁")) = IIf(.TextMatrix(lngRow, .ColIndex("是否锁定")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("锁"), lngRow, .ColIndex("锁")) = flexPicAlignCenterCenter
                
                .Cell(flexcpPicture, lngRow, .ColIndex("调"), lngRow, .ColIndex("调")) = IIf(mrsTrans!是否确认调整 = 1, Me.ImgList.ListImages(8).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("调"), lngRow, .ColIndex("调")) = flexPicAlignCenterCenter
                                
                .TextMatrix(lngRow, .ColIndex("核查人")) = mrsTrans!核查人
                .TextMatrix(lngRow, .ColIndex("核查时间")) = mrsTrans!核查时间
                .TextMatrix(lngRow, .ColIndex("摆药人")) = mrsTrans!摆药人
                .TextMatrix(lngRow, .ColIndex("摆药时间")) = mrsTrans!摆药时间
                .TextMatrix(lngRow, .ColIndex("摆药单号")) = IIf(zlStr.nvl(mrsTrans!摆药单号) = "", " ", mrsTrans!摆药单号)
                .TextMatrix(lngRow, .ColIndex("配药人")) = mrsTrans!配药人
                .TextMatrix(lngRow, .ColIndex("配药时间")) = mrsTrans!配药时间
                .TextMatrix(lngRow, .ColIndex("发送人")) = mrsTrans!发送人
                .TextMatrix(lngRow, .ColIndex("发送时间")) = mrsTrans!发送时间
                .TextMatrix(lngRow, .ColIndex("销帐申请人")) = mrsTrans!销帐申请人
                .TextMatrix(lngRow, .ColIndex("销帐申请时间")) = mrsTrans!销帐申请时间
                .TextMatrix(lngRow, .ColIndex("销帐审核人")) = mrsTrans!销帐审核人
                .TextMatrix(lngRow, .ColIndex("销帐审核时间")) = mrsTrans!销帐审核时间
                .TextMatrix(lngRow, .ColIndex("医嘱发送时间")) = mrsTrans!医嘱发送时间
                .TextMatrix(lngRow, .ColIndex("拒收原因")) = mrsTrans!拒收原因
                .TextMatrix(lngRow, .ColIndex("销帐原因")) = mrsTrans!销帐原因
                
                .TextMatrix(lngRow, .ColIndex("是否打包")) = mrsTrans!是否打包
                .TextMatrix(lngRow, .ColIndex("打印标志")) = mrsTrans!打印标志
                .TextMatrix(lngRow, .ColIndex("配药ID")) = mrsTrans!配药id
                
                .TextMatrix(lngRow, .ColIndex("抗菌药物")) = mrsTrans!抗菌药物
                
                '加载药品信息
                .TextMatrix(lngRow, .ColIndex("药品名称")) = mrsTrans!药品名称
                .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.nvl(mrsTrans!规格)
                .TextMatrix(lngRow, .ColIndex("配药类型")) = IIf(IsNull(mrsTrans!实际配药类型), "", mrsTrans!实际配药类型)
                .TextMatrix(lngRow, .ColIndex("单量")) = FormatEx(mrsTrans!单量, 2) & mrsTrans!剂量单位
                .TextMatrix(lngRow, .ColIndex("数量")) = FormatEx(mrsTrans!数量, 2) & mrsTrans!单位
                .TextMatrix(lngRow, .ColIndex("NO")) = zlStr.nvl(mrsTrans!NO)
                .TextMatrix(lngRow, .ColIndex("单据")) = nvl(mrsTrans!单据)
                .TextMatrix(lngRow, .ColIndex("药品id")) = mrsTrans!药品ID
                .TextMatrix(lngRow, .ColIndex("执行频次")) = mrsTrans!执行频次
                .TextMatrix(lngRow, .ColIndex("溶媒")) = mrsTrans!溶媒
                .TextMatrix(lngRow, .ColIndex("对应医嘱ID")) = mrsTrans!对应医嘱ID
                
                If mrsTrans!是否皮试 = 1 Then
                    .TextMatrix(lngRow, .ColIndex("皮")) = Get皮试结果(Val(mrsTrans!病人ID), Val(mrsTrans!药名ID), dateCurrent, CDate(mrsTrans!开嘱时间), mrsTrans!主页id)
                End If
                
                .Cell(flexcpPicture, lngRow, .ColIndex("医嘱"), lngRow, .ColIndex("医嘱")) = IIf(Format(zlStr.nvl(mrsTrans!药师审核时间), "YYYY-MM-dd") = Format(dteCur, "YYYY-MM-dd"), Me.ImgList.ListImages(2).Picture, Me.ImgList.ListImages(1).Picture)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("医嘱"), lngRow, .ColIndex("医嘱")) = flexPicAlignCenterCenter
                
                .Cell(flexcpPicture, lngRow, .ColIndex("变"), lngRow, .ColIndex("变")) = IIf(mrsTrans!是否调整批次 = 1, Me.ImgList.ListImages(7).Picture, Nothing)
                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("变"), lngRow, .ColIndex("变")) = flexPicAlignCenterCenter
                
                If Not gobjPass Is Nothing Then
                    .Cell(flexcpPicture, lngRow, .ColIndex("审查结果"), lngRow, .ColIndex("审查结果")) = gobjPass.zlPassSetWarnLight_YF(Val(mrsTrans!审查结果))
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审查结果"), lngRow, .ColIndex("审查结果")) = flexPicAlignCenterCenter
                End If
                
                '显示[自备药]标志
                If mrsTrans!执行性质 = 5 And mrsTrans!执行标记 = 0 Then
                    .Cell(flexcpPicture, lngRow, .ColIndex("药品名称"), lngRow, .ColIndex("药品名称")) = Me.ImgPro.ListImages("自备药").Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("药品名称"), lngRow, .ColIndex("药品名称")) = flexPicAlignLeftCenter
                End If
                
                .Cell(flexcpForeColor, lngRow, .ColIndex("配药批次"), lngRow, .ColIndex("配药批次")) = IIf(mrsTrans!批次标记 = 2, vbRed, IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln批次设置 = True, vbBlue, vbBlack))
                lngRow = lngRow + 1
                
                mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) > IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")), 0, 1)
            
            mrsTrans.MoveNext
            
            If mstrFilter <> "" And mrsTrans.EOF Then
                mrsTrans.Filter = ""
                LngID = 0
            End If
        Loop
        
'        '选择列加粗，蓝色显示
'        .Cell(flexcpFontBold, 0, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = True
'        .Cell(flexcpForeColor, 0, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = vbBlue
        .Cell(flexcpBackColor, 1, .ColIndex("选择"), .rows - 1, .ColIndex("选择")) = CSTCOLOR_MODIFY
        
        .Cell(flexcpFontBold, 0, .ColIndex("审"), 0, .ColIndex("审")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("审"), 0, .ColIndex("审")) = vbBlue
        
        '打包、批次列显示：根据参数不同而定
        .Cell(flexcpFontBold, 0, .ColIndex("打包"), 0, .ColIndex("打包")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln打包设置 = True, True, False)
        .Cell(flexcpForeColor, 0, .ColIndex("打包"), .rows - 1, .ColIndex("打包")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln打包设置 = True, vbBlue, vbBlack)
        .Cell(flexcpBackColor, 1, .ColIndex("打包"), .rows - 1, .ColIndex("打包")) = IIf((mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE) And mParams.bln打包设置 = True, CSTCOLOR_MODIFY, CSTCOLOR_UNMODIFY)
       
        .Cell(flexcpFontBold, 0, .ColIndex("配药批次"), 0, .ColIndex("配药批次")) = IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln批次设置 = True, True, False)
        .Cell(flexcpBackColor, 1, .ColIndex("配药批次"), .rows - 1, .ColIndex("配药批次")) = IIf(mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln批次设置 = True, CSTCOLOR_MODIFY, CSTCOLOR_UNMODIFY)
        
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                '更新打印标志图标
                If Val(.TextMatrix(lngRow, .ColIndex("打印标志"))) = 1 Then
                    .Row = lngRow
                    .Col = .ColIndex("打印")
                    .CellPicture = picPrint(1).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
                
                '更新配液(打包)图标
                If Val(.TextMatrix(lngRow, .ColIndex("是否打包"))) > 0 Then
                    .Row = lngRow
                    .Col = .ColIndex("打包")
                    .CellPicture = picPacker(Val(.TextMatrix(lngRow, .ColIndex("是否打包")))).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
                
                '设置单个病人的背景色
                If Val(.TextMatrix(lngRow, .ColIndex("背景号"))) = 1 Then
                    .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = &H80000005
                Else
                    .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = &HC0FFC0
                End If
            End If
        Next
        
'        Call SetTransColHide
        Call GetCount
        
        .Redraw = flexRDDirect
    End With
    
    Call UpdateExeSign(-1)
    ShowTrans = True
End Function

Private Sub PIVAWork_Sure()
'确认调整
Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                    End If
                    
                    .Cell(flexcpPicture, lngRow, .ColIndex("调"), lngRow, .ColIndex("调")) = Me.ImgList.ListImages(8).Picture
                    .Cell(flexcpPictureAlignment, lngRow, .ColIndex("调"), lngRow, .ColIndex("调")) = flexPicAlignCenterCenter
                    
                    mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                    Do While Not mrsTrans.EOF
                        mrsTrans!是否确认调整 = 1
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
                
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "请选择要确认调整的输液单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_确认调整("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "确认调整")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ShowDeptAdvice(ByVal intStep As Integer, ByVal index As Long)
    Dim lng病区id As Long
    Dim dblCount As Double
    
    With vsfDept(0)
        .rows = 1
        .rows = 2
            
        .Cell(flexcpText, 1, .ColIndex("选择"), 1, .Cols - 1) = "没有医嘱信息......"
        .MergeCells = flexMergeRestrictRows
        .MergeRow(1) = True
     
        If mrsDeptAdvice Is Nothing Then Exit Sub
        
        mrsDeptAdvice.Filter = "核查标志=" & intStep
        mrsDeptAdvice.Sort = "病区,病区ID,核查标志"
        
        If mrsDeptAdvice.RecordCount = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        
        .rows = 1
        
        mrsDeptAdvice.MoveFirst
        Do While Not mrsDeptAdvice.EOF
            If lng病区id <> mrsDeptAdvice!病区ID Then
                lng病区id = mrsDeptAdvice!病区ID
                
                .rows = .rows + 1
                
                If mstr上次病区ID <> "" Then
                    .TextMatrix(.rows - 1, .ColIndex("选择")) = IIf(InStr(1, mstr上次病区ID, mrsDeptAdvice!病区ID) > 0, -1, 0)
                Else
                    .TextMatrix(.rows - 1, .ColIndex("选择")) = IIf(mrsDeptAdvice!选择 = 1, -1, 0)
                End If
                .TextMatrix(.rows - 1, .ColIndex("病区")) = mrsDeptAdvice!病区
                .TextMatrix(.rows - 1, .ColIndex("数量")) = mrsDeptAdvice!数量
                .TextMatrix(.rows - 1, .ColIndex("病区ID")) = mrsDeptAdvice!病区ID
            Else
                .TextMatrix(.rows - 1, .ColIndex("数量")) = Val(.TextMatrix(.rows - 1, .ColIndex("数量"))) + mrsDeptAdvice!数量
            End If
            
            mrsDeptAdvice.MoveNext
        Loop
        
        .Cell(flexcpFontBold, 1, .ColIndex("选择"), .rows - 1, .ColIndex("选择")) = True
        .Cell(flexcpForeColor, 1, .ColIndex("选择"), .rows - 1, .ColIndex("选择")) = vbBlue
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub ShowDeptTrans(ByVal intType As Integer, ByVal strType As String)
    '显示病区及病区对应的输液单据数量
    With vsfDept(intType)
        .rows = 1
        .rows = 2
        
        .Cell(flexcpText, 1, .ColIndex("选择"), 1, .Cols - 1) = "没有输液单信息......"
        .MergeCells = flexMergeRestrictRows
        .MergeRow(1) = True
        
        If mrsDeptTrans Is Nothing Then Exit Sub
        
        mrsDeptTrans.Filter = "类型='" & strType & "'"
        
        If mrsDeptTrans.RecordCount = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        
        .rows = 1
        
        If mParams.int病区排序 = 1 Then
            mrsDeptTrans.Sort = "编码"
        Else
            mrsDeptTrans.Sort = "名称"
        End If
        
        Do While Not mrsDeptTrans.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("选择")) = IIf(mrsDeptTrans!选择 = 1, -1, 0)
            .TextMatrix(.rows - 1, .ColIndex("病区")) = mrsDeptTrans!病区
            .TextMatrix(.rows - 1, .ColIndex("数量")) = mrsDeptTrans!数量
            .TextMatrix(.rows - 1, .ColIndex("病区ID")) = mrsDeptTrans!病区ID
            .TextMatrix(.rows - 1, .ColIndex("记录id")) = mrsDeptTrans!记录id
            
            mrsDeptTrans.MoveNext
        Loop
'        .Cell(flexcpFontBold, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = True
'        .Cell(flexcpForeColor, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = vbBlue
        .Cell(flexcpBackColor, 1, .ColIndex("选择"), .rows - 1, .ColIndex("选择")) = CSTCOLOR_MODIFY
        
        .Redraw = flexRDDirect
    End With
    
    '统计选择的病区和输液单据数量
    Call GetCount
End Sub

Private Sub UpdateExeSign(ByVal LngID As Long, Optional ByVal intSign As Integer)
    '根据表格数据更新数据集执行标志
    'lngID：0-所有记录按intSign值更新;>0-对应的数据集记录按intSign值更新(医嘱时表示相关ID；输液单时表示配药ID);"-1"-根据表格数据更新
    'intSign：当lngID=0,>0时传入
    Dim lngCount As Long
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If LngID = 0 Then
        If mrsTrans.RecordCount = 0 And mrsTrans.Filter <> "" Then mrsTrans.Filter = ""
        Do While Not mrsTrans.EOF
            mrsTrans!执行标志 = intSign
            mrsTrans.Update
            mrsTrans.MoveNext
        Loop
    ElseIf LngID = -1 Then
        With vsfTrans
            For lngCount = 1 To .rows - 1
                If .TextMatrix(lngCount, .ColIndex("配药ID")) <> "" Then
                    mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(lngCount, .ColIndex("配药ID")))
                    Do While Not mrsTrans.EOF
                        If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                            mrsTrans!执行标志 = Val(.TextMatrix(lngCount, .ColIndex("标志")))
                        Else
                            mrsTrans!执行标志 = IIf(Val(.TextMatrix(lngCount, .ColIndex("选择"))) = -1, 1, 0)
                        End If
                        mrsTrans.Update
                        mrsTrans.MoveNext
                    Loop
                End If
            Next
        End With
    Else
        mrsTrans.Filter = "配药ID=" & LngID
        Do While Not mrsTrans.EOF
            mrsTrans!执行标志 = intSign
            mrsTrans.Update
            mrsTrans.MoveNext
        Loop
    End If
    
    DoEvents
    
    Call GetCount
End Sub

Private Sub cboBatch_Click()
    mlng已扫描 = 0
    Call SetFilter
    If mcondition.strTransStep = M_STR_CALSS_SEND Then Me.txtFindItem.SetFocus
End Sub

Private Sub cboLevel_Click()
    Call SetFilter
End Sub

Private Sub SetFilter()
    Dim bln卡片 As Boolean
    Dim bln列表 As Boolean
    Dim lngRow As Long
    Dim lngCount As Long
    
    Me.chkAll.Value = 0
    
    Call ClearDetailList
    If mblnFilter = False Then Exit Sub
    With vsfDept(Me.tabDeptList.Selected.index)
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("病区ID")) <> "" Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    lngCount = lngCount + 1
                End If
            End If
        Next
    End With
    
    If lngCount = 0 Then Exit Sub
    If mrsTrans Is Nothing Then Exit Sub
    
    mrsTrans.Filter = ""
    If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("病人id")) = "" Then
        If mrsTrans Is Nothing Then
            Exit Sub
        Else
            If mrsTrans.RecordCount = 0 Then Exit Sub
        End If
    End If
    
    If cboBatch.Text = "<全部>" And cboLevel.Text = "<全部>" Then
        mrsTrans.Filter = ""
    ElseIf cboBatch.Text <> "<全部>" And cboLevel.Text = "<全部>" Then
        mrsTrans.Filter = "配药批次=" & cboBatch.Text
    ElseIf cboBatch.Text = "<全部>" And cboLevel.Text <> "<全部>" Then
        mrsTrans.Filter = "优先级=" & cboLevel.Text
    Else
        mrsTrans.Filter = "配药批次=" & cboBatch.Text & IIf(cboLevel.Text = "<全部>", "", " And 优先级=" & cboLevel.Text)
    End If
    
    If Me.cboFrequency.Text <> "<全部>" Then
        mrsTrans.Filter = IIf(mrsTrans.Filter = 0, "执行频次='" & Me.cboFrequency.Text & "'", mrsTrans.Filter & " And 执行频次='" & Me.cboFrequency.Text & "'")
    End If
    
    If Me.chkSure(1).Value = 1 And Me.chkSure(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 是否确认调整=1", "是否确认调整=1")
    ElseIf Me.chkSure(0).Value = 1 And Me.chkSure(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 是否确认调整=0", "是否确认调整=0")
    End If
    
    If Me.chkPrint(1).Value = 1 And Me.chkPrint(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 打印标志=1", "打印标志=1")
    ElseIf Me.chkPrint(0).Value = 1 And Me.chkPrint(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 打印标志=0", "打印标志=0")
    End If
    
    If Me.chkChange(1).Value = 1 And Me.chkChange(0).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 是否调整批次=1", "是否调整批次=1")
    ElseIf Me.chkChange(0).Value = 1 And Me.chkChange(1).Value = 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 是否调整批次=0", "是否调整批次=0")
    End If
    
    If Me.cboDosType.ListIndex <> 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 配药类型='" & Me.cboDosType.Text & "'", "配药类型='" & Me.cboDosType.Text & "'")
    End If

    '按过滤的药品进行排序
    If Me.cboMedi.Text <> "<全部>" Or (mParams.blnByMedi = True Or mParams.blnFilter = False) Then
        Call MediSort
        
        If mrsTrans.RecordCount = 0 Then
            Me.vsfTrans.rows = 1
            Me.vsfTrans.rows = 2
            Exit Sub
        End If
        
        Call SetSortFlag(True)
    Else
        Call SetSortFlag
    End If
    
    '显示输液单明细列表
    bln列表 = ShowTrans(Me.tabDeptList.Selected.index)
    '显示输液单药品汇总列表
    Call ShowSumDrug
    '在状态栏显示所选的病区和输液单数量
    Call GetCount
    '显示输液单卡片
    bln卡片 = mfrmPIVCard.ShowDetailCard(mrsTrans, mstr批次, mcondition.strTransStep = M_STR_CALSS_PREPARE, mParams.intCount, mParams.bln批次设置, mParams.bln打包设置, mcondition.strTransStep, mParams.bln审核)
    
    If bln列表 And bln卡片 Then
        chkAll.Enabled = True
    End If
End Sub

Private Sub LoadData()
    Dim rstemp As Recordset
    On Error GoTo errHandle
    
    gstrSQL = "select 科室id,科室名称,配药类型,频次,有效,优先级 from 输液药品优先级 order by 优先级"
    Set mrsPRI = zlDatabase.OpenSQLRecord(gstrSQL, "获取优先级数据")
    
    gstrSQL = "select distinct 优先级 from 输液药品优先级 order by 优先级"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取优先级数据")
    
    cboLevel.Clear
    Me.cboLevel.AddItem "<全部>"
    Do While Not rstemp.EOF
        Me.cboLevel.AddItem rstemp!优先级
        rstemp.MoveNext
    Loop
    cboLevel.Text = "<全部>"

    gstrSQL = "select 科室id,科室名称,容量,配药批次 from 科室容量设置 where 配置中心ID=[1]"
    Set mrsVol = zlDatabase.OpenSQLRecord(gstrSQL, "获取科室容量数据", mParams.lng配置中心)
    
    cboMedi.Clear
    Me.cboMedi.AddItem "<全部>"
    If mParams.blnFilter Then
        gstrSQL = "select distinct 药品id,名称 from 输液优先打印药品"
        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取药品信息")

        Do While Not rstemp.EOF
            Me.cboMedi.AddItem rstemp!名称
            Me.cboMedi.ItemData(Me.cboMedi.ListCount - 1) = rstemp!药品ID

            mParams.str常用药品 = IIf(mParams.str常用药品 = "", "", mParams.str常用药品 & ",") & rstemp!药品ID
            rstemp.MoveNext
        Loop
    End If
    cboMedi.Text = "<全部>"
    
    Set rstemp = DeptSendWork_Get配药类型
    cboType.Clear
    cboDosType.Clear
    Me.cboType.AddItem "<全部>"
    Me.cboDosType.AddItem "<全部>"
    Do While Not rstemp.EOF
        Me.cboType.AddItem rstemp!编码 & "-" & rstemp!名称
        Me.cboDosType.AddItem rstemp!编码 & "-" & rstemp!名称
        rstemp.MoveNext
    Loop
    cboDosType.Text = "<全部>"
    cboType.Text = "<全部>"
    
    gstrSQL = "select 名称 from 诊疗频率项目"
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "执行频次")
    
    cboFrequency.Clear
    Me.cboFrequency.AddItem "<全部>"
    Do While Not rstemp.EOF
        Me.cboFrequency.AddItem rstemp!名称
        rstemp.MoveNext
    Loop
    Me.cboFrequency.Text = "<全部>"
    
    Me.cboBatch.Text = "<全部>"
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub MediSort()
    Dim lng药品id As Long
    Dim strFilter As String
    Dim str配药ids As String
    Dim lng配药id As Long
    Dim i As Long
    Dim j As Long
    Dim strSort As String
    Dim strIDSOld As String
    Dim str配药id As String
    Dim strIDS As String
    Dim strTemp As String
    
    lng药品id = Val(Me.cboMedi.ItemData(Me.cboMedi.ListIndex))
    strFilter = IIf(mrsTrans.Filter = 0, "", mrsTrans.Filter)
    mstrFilter = ""
    
    If lng药品id <> 0 Then
        mrsTrans.Filter = IIf(mrsTrans.Filter <> 0, mrsTrans.Filter & " and 药品id=" & lng药品id, "药品id=" & lng药品id)
        mrsTrans.Sort = "药品id,单量"
        
        Do While Not mrsTrans.EOF
            strIDSOld = IIf(strIDSOld = "", mrsTrans!配药id, strIDSOld & "," & mrsTrans!配药id)
            str配药id = IIf(str配药id = "", "配药id=" & mrsTrans!配药id, str配药id & " or 配药id=" & mrsTrans!配药id)
            mrsTrans.MoveNext
        Loop
        
        If str配药id = "" Then Exit Sub
        mrsTrans.Filter = str配药id
        mrsTrans.Sort = "溶媒,药品id"
        
        lng药品id = 0
        Do While Not mrsTrans.EOF
            If mrsTrans!溶媒 = 1 Then
                If lng药品id <> mrsTrans!药品ID Or lng药品id = 0 Then
                    lng药品id = mrsTrans!药品ID
                    strIDS = IIf(strIDS = "", mrsTrans!配药id, strIDS & "|" & mrsTrans!配药id)
                Else
                    strIDS = IIf(strIDS = "", mrsTrans!配药id, strIDS & "," & mrsTrans!配药id)
                End If
'                mstrFilter = IIf(mstrFilter = "", mrsTrans!配药id, mstrFilter & "," & mrsTrans!配药id)
            End If
            mrsTrans.MoveNext
        Loop
        
        str配药id = ""
        For i = 0 To UBound(Split(strIDS, "|"))
            For j = 0 To UBound(Split(strIDSOld, ","))
                If InStr(1, "," & Split(strIDS, "|")(i) & ",", "," & Split(strIDSOld, ",")(j) & ",") > 0 Then
                    If InStr(1, "," & mstrFilter & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                        mstrFilter = IIf(mstrFilter = "", Split(strIDSOld, ",")(j), mstrFilter & "," & Split(strIDSOld, ",")(j))
                    End If
                End If
            Next
        Next
        
        For j = 0 To UBound(Split(strIDSOld, ","))
            If InStr(1, "," & mstrFilter & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                If InStr(1, "," & strTemp & ",", "," & Split(strIDSOld, ",")(j) & ",") < 1 Then
                    strTemp = strTemp & "," & Split(strIDSOld, ",")(j)
                End If
            End If
        Next
        
        mstrFilter = IIf(mstrFilter = "", Mid(strTemp, 2), mstrFilter & strTemp)
        mrsTrans.Filter = ""
    Else
        mrsTrans.Sort = "配药批次,排序药品id,溶媒id,排序单量,瓶签号,溶媒"
    End If
    
    
    
'    If mParams.blnByMedi = True And Val(Me.vsfTrans.Tag) = 1 Then
'        mrsTrans.Sort = "配药批次,排序药品id,溶媒id,排序单量,瓶签号,溶媒"
'    Else
'        strSort = mParams.strSort
'        If InStr(1, mParams.strSort, "配药批次") > 0 Then
'            strSort = Replace(mParams.strSort, "配药批次", "配药批次,优先级")
'        End If
'        If InStr(1, mParams.strSort, "床号") > 0 Then
'            strSort = Replace(mParams.strSort, "床号", "编码,床号")
'        End If
'        mrsTrans.Sort = IIf(strSort <> "", strSort & ",配药id,溶媒", "配药id,溶媒")
'    End If
End Sub













Private Sub cbo时间范围_Click()
    Dim dteTime As Date
    
    With cbo时间范围
        If .ListIndex <> Val(.Tag) Then
            If (Val(.Tag) = 3 And .ListIndex < 3) Or (Val(.Tag) < 3 And .ListIndex = 3) Then
                Call ResizeConditionArea
            End If
            .Tag = .ListIndex
            
            dteTime = Sys.Currentdate
            
            If .ListIndex = 0 Then
                Dtp开始时间.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
                Dtp结束时间.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
            ElseIf .ListIndex = 1 Then
                Dtp开始时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
                Dtp结束时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
            ElseIf .ListIndex = 2 Then
                Dtp开始时间.Value = CDate(Format(dteTime, "YYYY-MM-DD"))
                Dtp结束时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD"))
            ElseIf .ListIndex = 3 Then
                If mcondition.strTransStartTime <> "" Then
                    Dtp开始时间.Value = CDate(Format(mcondition.strTransStartTime, "YYYY-MM-DD hh:mm:ss"))
                    Dtp结束时间.Value = CDate(Format(mcondition.strTransEndTime, "YYYY-MM-DD hh:mm:ss"))
                Else
                    Dtp开始时间.Value = CDate(Format(dteTime, "YYYY-MM-DD hh:mm:ss"))
                    Dtp结束时间.Value = CDate(Format(dteTime, "YYYY-MM-DD hh:mm:ss"))
                End If
            End If
            mcondition.intTransTimeSel = .ListIndex
        End If
    End With
    
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
End Sub


Private Sub cbsMain_ControlSelected(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Me.cbsMain Is Nothing Then Exit Sub
    If Control Is Nothing Then Exit Sub
    
    '弹出菜单：选择
    If Control.Id = conMenu_Oper_Select Then
        '按批次选择
        Set cbrControl = Me.cbsMain.FindControl(xtpControlButton, conMenu_Oper_Select_SelBatch, False, True)
        If Not cbrControl Is Nothing Then
            cbrControl.Enabled = True
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim lngPatiID As Long
    Dim str挂号单 As String
    Dim lng主页ID As Long
    Dim lngCurrAdviceID As Long

    Select Case Control.Id
        ''''文件
        Case mconMenu_File_PrintSet     '打印设置
            zlPrintSet
        Case mconMenu_File_Preview      '打印预览
            zlSubPrint 2
        Case mconMenu_File_Print        '打印
            zlSubPrint 1
        Case mconMenu_File_Excel        '输出到Excel
            zlSubPrint 3

        Case mconMenu_File_PIVA_BillPrintLable
            '打印标签
            Call BillPrint_Label(Control.Id)
        Case mconMenu_Edit_PIVA_Approve
            '审核
            Call PIVAWork_Approve
        Case mconMenu_Edit_PIVA_Beach
            '重排批次
            Call SetBeach
        Case MCONMENU_EDIT_PIVA_SURE
            '确认调整
            Call PIVAWork_Sure
        Case mconMenu_Edit_PIVA_Prepare
            '执行预调价
            Call setNOtExcetePrice
    
            '摆药确认
            Call PIVAWork_Prepare(1)
        Case mconMenu_Edit_PIVA_Dosage
            '配药确认
            Call PIVAWork_Dosage
        Case mconMenu_Edit_PIVA_Send
            '发送确认
            Call PIVAWork_Send
        Case MCONMENU_EDIT_PIVA_REFUSE
            '确认拒绝
            Call PIVAWork_Refuse
        Case mconMenu_Edit_PIVA_ReVerify
            '执行预调价
            Call setNOtExcetePrice
            
            '销帐审核
            Call PIVAWork_ReturnVerify
        Case mconMenu_Edit_PIVA_Cancel
            '执行预调价
            Call setNOtExcetePrice
            
            '取消上一步操作
            Call PIVAWork_Cancel
        Case mconMenu_Edit_PIVA_Delete
            '删除已回退了医嘱的输液配药记录
            Call PIVAWork_Delete
            
        Case MCONMENU_PLAN_PIVA_DESK
            Call frmDesk.ShowMe(mParams.lng配置中心, Me)
        Case MCONMENU_PLAN_PIVA_DESKDRUG
            Call frmDeskMedi.ShowMe(mParams.lng配置中心, Me)
        Case MCONMENU_PLAN_PIVA_PERWORK
            Call frmPlan.ShowMe(mParams.lng配置中心, Me)
        Case mconMenu_Edit_PIVA_PASS
            '功能：对病人过敏史/病生状态进行管理
            'Pass
            If Not Me.vsfMedis Is ActiveControl Then Exit Sub
            If vsfMedis.Row = 0 Then Exit Sub
            
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("病人id")))
            lng主页ID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("主页id")))
            
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng主页ID, "")
        
        '打印操作
        Case mconMenu_File_PIVA_BillPrintWait
            '打印摆药单
            Call BillPrint_Prepare
        Case mconMenu_File_PIVA_BillPrintTotal
            '打印发送单
            Call BillPrint_Send
        Case mconMenu_File_PIVA_BillPrintReturn
            '打印退药销帐清单
            Call BillPrint_Return
        Case mconMenu_File_PIVA_BillPrintNext
            Call frmPrint.ShowMe(Me)
        Case mconMenu_File_PIVA_BillPrintSum
            '打印汇总报表
            Call BillPrint_Sum
        Case mconMenu_File_Parameter
            '参数设置
            ResetParams
        Case MCONMENU_EDIT_PIVA_SORTSET
            '设置排序规则
            frmPIVASortSet.Show 1, Me
            Call SetSort(True)
        Case mconMenu_View_Refresh
            '刷新
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
        
        Case mconMenu_File_Exit
            '退出
            Unload Me
        
        ''''查看
        Case mconMenu_View_ToolBar_Button               '标准按钮
            Control.Checked = Not Control.Checked
            Me.cbsMain(2).Visible = Control.Checked
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Text                 '文本标签
            Control.Checked = Not Control.Checked
            For Each cbrControl In Me.cbsMain(2).Controls
                cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_ToolBar_Size                 '大图标
            Control.Checked = Not Control.Checked
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_StatusBar                    '状态栏
            Control.Checked = Not Control.Checked
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3                   '字号设置
            mParams.intFont = Val(Control.Parameter)
            Call SetFontSize
        Case mconMenu_View_ShowHistory
            Control.Checked = Not Control.Checked
            mParams.intAutoSelect = IIf(Control.Checked, 1, 0)
        Case mconMenu_Edit_PIVA_Lock
            '当前数据全部锁定
            Call SetLock(1, "")
        Case mconMenu_Edit_PIVA_UnLock
            '当前数据全部解锁
            Call SetLock(0, "")
            
        ''''帮助
        Case mconMenu_Help_Help                         '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, "Frm部门发药管理")
        Case mconMenu_Help_Web                          'WEB上的中联
        Case mconMenu_Help_Web_Home                     '中联主页
            Call zlHomePage(Me.hWnd)
        Case mconMenu_Help_Web_Forum                    '中联论坛
            Call zlWebForum(Me.hWnd)
        Case mconMenu_Help_Web_Mail                     '发送反馈
            Call zlMailTo(Me.hWnd)
        Case mconMenu_Help_About                        '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        
        ''弹出菜单
        Case conMenu_Oper_PrintLabel_SelRow, conMenu_Oper_PrintLabel_SelBatch, conMenu_Oper_PrintLabel_SelDept, conMenu_Oper_PrintLabel_SelPati, conMenu_Oper_PrintLabel_AllRow, conMenu_Oper_PrintLabel_SelSendNo
            '打印标签
            Call BillPrint_Label(Control.Id)
        Case conMenu_Oper_DelBatch_SelRow, conMenu_Oper_DelBatch_SelBatch, conMenu_Oper_DelBatch_SelDept, conMenu_Oper_DelBatch_SelPati, conMenu_Oper_DelBatch_AllRow
            '删除批次
            Call DeleteBatch(Control.Id)
        Case conMenu_Oper_Select_SelRow, conMenu_Oper_Select_SelBatch, conMenu_Oper_Select_SelDept, conMenu_Oper_Select_CancleSelDept, conMenu_Oper_Select_SelPati, conMenu_Oper_Select_CancleSelPati, conMenu_Oper_Select_SelSendNo, conMenu_Oper_Select_SelAll, conMenu_Oper_Select_SelMed, conMenu_Oper_Bag_Batch, conMenu_Oper_Bag_All
            '批量选择,打包
            Call SelectBatch(Control.Id, (Me.tabDeptList.Selected.index))
'        Case conMenu_Oper_DelLabel_SelRow, conMenu_Oper_DelLabel_SelBatch, conMenu_Oper_DelLabel_SelDept, conMenu_Oper_DelLabel_SelPati, conMenu_Oper_DelLabel_AllRow
'            '删除标签
'            Call DeleteLabel(Control.Id)
        Case conMenu_Oper_Look
            On Error Resume Next
            '电子病案查阅
            If Not mobjCISJOB Is Nothing Then
                Call mobjCISJOB.ShowArchive(Me, Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("病人id"))), Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("主页id"))))
            End If
            
            err.Clear
        Case MCONMENU_EDIT_PIVA_MedicalRecord
            '电子病案查阅(工具栏)
            Call ShowMedicalRecord
        Case mconMenu_Edit_PlugIn + 1 To mconMenu_Edit_PlugIn + 99 '外挂发药业务功能调用
            PivaExPlugNormal Control.Parameter
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            lngCurrAdviceID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("医嘱id")))
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("病人id")))
            lng主页ID = Val(vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("主页id")))
            
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng主页ID, "", lngCurrAdviceID)
        '弹出菜单：PASS命令
'        Case mconMenu_PASS_Item
'            '药物临床信息参考
'            Call PassDoCommand(101)
'        Case mconMenu_PASS_Item + 1
'            '药品说明书
'            Call PassDoCommand(102)
'        Case mconMenu_PASS_Item + 2
'            '中国药典
'            Call PassDoCommand(107)
'        Case mconMenu_PASS_Item + 3
'            '病人用药教育
'            PassDoCommand (103)
'        Case mconMenu_PASS_Item + 4
'             '检验值
'            Call PassDoCommand(104)
'        Case mconMenu_PASS_Item + 6
'            '医药信息中心
'            Call PassDoCommand(106)
'        Case mconMenu_PASS_Item + 7
'            '药品配对信息
'             Call PassDoCommand(13)
'        Case mconMenu_PASS_Item + 8
'            '给药途径配对信息
'            Call PassDoCommand(14)
'        Case mconMenu_PASS_Item + 9
'            '医院药品信息
'            Call PassDoCommand(105)
'
'        '功能：执行专项PASS命令
'        Case mconMenu_PASS_Spec
'            '药物-药物相互作用
'            Call PassDoCommand(201)
'        Case mconMenu_PASS_Spec + 1
'            '药物-食物相互使用
'            Call PassDoCommand(202)
'        Case mconMenu_PASS_Spec + 2
'            '国内注射剂配伍
'            Call PassDoCommand(203)
'        Case mconMenu_PASS_Spec + 3
'            '国外注射剂配伍
'            Call PassDoCommand(204)
'        Case mconMenu_PASS_Spec + 4
'            '禁忌症
'            Call PassDoCommand(205)
'        Case mconMenu_PASS_Spec + 5
'            '副作用
'            Call PassDoCommand(206)
'        Case mconMenu_PASS_Spec + 6
'            '老年人用药
'            Call PassDoCommand(207)
'        Case mconMenu_PASS_Spec + 7
'            '儿童用药
'            Call PassDoCommand(208)
'        Case mconMenu_PASS_Spec + 8
'            '妊娠期用药
'            Call PassDoCommand(209)
'        Case mconMenu_PASS_Spec + 9
'            '哺乳期用药
'            Call PassDoCommand(210)
'        Case mconMenu_PASS_Spec + 10
'            Call AdviceCheckWarn(9, "0000000", 2, 1, Me.vsfMedis.TextMatrix(Me.vsfMedis.Row, Me.vsfMedis.ColIndex("医嘱id")))
        Case Else
            '查找菜单
            If Control.Id > mconMenu_Look And Control.Id < mconMenu_Look + 10 Then
                lblFindItem.Caption = Control.Caption
                
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                        If cbrControl.Caption = lblFindItem.Caption Then
                            cbrControl.Checked = True
                        End If
                    Next
                End If
            End If
            
            If Control.Id > mconMenu_Filter And Control.Id < mconMenu_Filter + 10 Then
                lblName.Caption = Control.Caption & "↓"
                
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Filter)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                        If cbrControl.Caption = Mid(lblName.Caption, 1, Len(lblName.Caption) - 1) Then
                            cbrControl.Checked = True
                        End If
                    Next
                End If
            End If
             
            If Control.Id > 401 And Control.Id < 499 Then
                '执行自定义报表
                Call BillPrint_Custom(Control)
            End If
        
            '病区排序弹出菜单
            If Control.Id > mconMenu_SortPopup And Control.Id < mconMenu_SortPopup + 10 Then
                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_SortPopup)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.CommandBar.Controls
                        cbrControl.Checked = False
                    Next
                End If
                
                Control.Checked = True
                If mParams.int病区排序 <> Control.Id - mconMenu_SortPopup Then
                    mParams.int病区排序 = Control.Id - mconMenu_SortPopup
                    Call ShowDeptTrans(Me.tabDeptList.Selected.index, IIf(Me.tabDeptList.Selected.index = CNUMWORK, tabWork.Selected.Tag, tbcLook.Selected.Tag))
                End If
            End If
    End Select
End Sub

Private Sub PivaExPlugNormal(ByVal strFunName As String)
    Dim lng配药id As Long
    
    If Not mobjPlugIn Is Nothing Then
        If vsfTrans.rows > 1 Then
            lng配药id = Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("配药ID")))
        End If
        
        On Error Resume Next
        Call mobjPlugIn.PivaWorkNormal(glngModul, strFunName, mParams.lng配置中心, lng配药id)
        err.Clear: On Error GoTo 0
    End If
    
End Sub

Private Sub SetSort(Optional ByVal BlnRefresh As Boolean = False)
    '设置输液单排序
    Dim strSortString As String
    Dim i As Integer
    Const ALL_SORT_ITEM As String = "病区,科室,床号,配药批次,姓名,瓶签号,执行时间"
    Const DEFAULT_SORT As String = "病区,床号,姓名"
    
    strSortString = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "输液配置中心管理", "输液单排序", "")
    If strSortString = "" Then
        strSortString = DEFAULT_SORT
    ElseIf InStr(1, strSortString, "|") = 0 Then
        strSortString = DEFAULT_SORT
    ElseIf Mid(strSortString, InStr(1, strSortString, "|") + 1) = "" Then
        strSortString = DEFAULT_SORT
    Else
        strSortString = Mid(strSortString, InStr(1, strSortString, "|") + 1)
        For i = 0 To UBound(Split(strSortString, ","))
            If Split(strSortString, ",")(i) <> "" Then
                If InStr(1, "," & ALL_SORT_ITEM & ",", "," & Split(strSortString, ",")(i) & ",") = 0 Then
                    strSortString = DEFAULT_SORT
                    Exit For
                End If
            End If
        Next
    End If
    
    If strSortString <> mParams.strSort Then
        mParams.strSort = strSortString
        If BlnRefresh = True Then Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    Call SetSortFlag
End Sub
Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strName As String
    
    strName = Split(Control.Parameter, ",")(1)

    Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strName, Me, _
        "开始时间=" & mcondition.strTransStartTime, _
        "结束时间=" & mcondition.strTransEndTime, _
        "配药记录=", _
        "瓶签号=")

End Sub
Private Sub SetFontSize()
    Dim intFont As Integer
    Dim stdfnt As StdFont
    
    Select Case mParams.intFont
        Case 0
            intFont = 9
        Case 1
            intFont = 11
        Case 2
            intFont = 15
        Case Else
            intFont = 9
    End Select
    
    With vsfDept(0)
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
    
    With vsfTrans
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
    
'    With vsfDrug
'        .Font.Size = intFont
'        Me.Font.Size = .Font.Size
'        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
'
'        .RowHeightMin = TextHeight("刘") + 120
'        .RowHeightMax = TextHeight("刘") + 120
'        .Refresh
'    End With
    
    With vsfSumDrug
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 120
        .RowHeightMax = TextHeight("刘") + 120
        .Refresh
    End With
    
    If Not tbcDetail.PaintManager.Font Is Nothing Then
        With tbcDetail
            Set stdfnt = .PaintManager.Font
            stdfnt.Size = intFont
             Set .PaintManager.Font = stdfnt
              .PaintManager.Layout = xtpTabLayoutAutoSize
        End With
    End If
    Me.FontSize = intFont
End Sub
Private Sub zlSubPrint(ByVal bytMode As Byte)
    'bytMode：1-打印；2-预览；3-输出到Excel
    Dim ObjThis As Object
    Dim objPrint As New zlPrint1Grd
    Dim ObjAppRow As New zlTabAppRow
    Dim strTitle As String
    
    '取打印列表对象
    Select Case tbcDetail.Selected.index
        Case mDetailType.输液单列表
            If vsfTrans.rows = 1 Then Exit Sub
            If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("配药ID")) = "" Then Exit Sub
            
            Set ObjThis = GetPrintObj(vsfTrans)
            
            Select Case mcondition.strTransStep
                Case M_STR_CALSS_PREPARE
                    strTitle = "待摆药输液单清单"
                Case M_STR_CALSS_DOSAGE
                    strTitle = "待配药输液单清单"
                Case M_STR_CALSS_SEND
                    strTitle = "待发送输液单清单"
                Case M_STR_CALSS_SENDED
                    strTitle = "已发送输液单清单"
            End Select
        Case mDetailType.药品汇总列表
            If vsfSumDrug.rows = 1 Then Exit Sub
            If vsfSumDrug.TextMatrix(1, vsfSumDrug.ColIndex("药品名称")) = "" Then Exit Sub
            
            Set ObjThis = GetPrintObj(vsfSumDrug)
            
            Select Case mcondition.strTransStep
                Case M_STR_CALSS_PREPARE
                    strTitle = "待摆药药品清单"
                Case M_STR_CALSS_DOSAGE
                    strTitle = "待配药药品清单"
                Case M_STR_CALSS_SEND
                    strTitle = "待发送药品清单"
                Case M_STR_CALSS_SENDED
                    strTitle = "已发送药品清单"
            End Select
    End Select
    
    If ObjThis Is Nothing Then Exit Sub
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "打印人:" & gstrUserName
    ObjAppRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy-MM-dd HH:MM:SS")
    objPrint.BelowAppRows.Add ObjAppRow
    
    Set ObjAppRow = New zlTabAppRow
    ObjAppRow.Add "开始时间:" & Format(Dtp开始时间.Value, "yyyy-MM-dd ")
    ObjAppRow.Add "结束时间:" & Format(Dtp结束时间.Value, "yyyy-MM-dd ")
    objPrint.UnderAppRows.Add ObjAppRow
    
    objPrint.Title.Text = strTitle
    Set objPrint.Body = ObjThis
    
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Function GetPrintObj(ByVal vsfObj As VSFlexGrid, Optional ByVal strMerge As String = "") As VSFlexGrid
    '返回打印控件
'    Dim vsfPrint As VSFlexGrid
    Dim lngCol As Long, lngRow As Long
    Dim lngCols As Long
    Dim lngPrintCol As Long
    
    With vsfPrint
        .Cols = 1
        .rows = 1
        
        For lngCol = 0 To vsfObj.Cols - 1
            If vsfObj.ColHidden(lngCol) = False And vsfObj.ColWidth(lngCol) > 0 And vsfObj.ColKey(lngCol) <> "当前行" Then
                lngCols = lngCols + 1
                
                If lngCols > 1 Then
                    .Cols = .Cols + 1
                End If
                
                .ColKey(lngCols - 1) = vsfObj.ColKey(lngCol)
                .ColWidth(lngCols - 1) = vsfObj.ColWidth(lngCol)
                .FixedAlignment(lngCols - 1) = vsfObj.FixedAlignment(lngCol)
                .ColAlignment(lngCols - 1) = vsfObj.ColAlignment(lngCol)
            End If
        Next
                
        For lngRow = 0 To vsfObj.rows - 1
            For lngCol = 0 To vsfObj.Cols - 1
                For lngPrintCol = 0 To .Cols - 1
                    If .ColKey(lngPrintCol) = vsfObj.ColKey(lngCol) Then
                        .TextMatrix(lngRow, lngPrintCol) = vsfObj.TextMatrix(lngRow, lngCol)
                        If .ColKey(lngPrintCol) = "选择" And lngRow > 0 Then
                            If Val(vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("选择"))) = -1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "√"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        
                        If .ColKey(lngPrintCol) = "打印" And lngRow > 0 Then
                            If vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("打印标志")) = 1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "√"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        
                        If .ColKey(lngPrintCol) = "打包" And lngRow > 0 Then
                            If vsfObj.TextMatrix(lngRow, vsfObj.ColIndex("是否打包")) = 1 Then
                                .TextMatrix(lngRow, lngPrintCol) = "√"
                            Else
                                .TextMatrix(lngRow, lngPrintCol) = ""
                            End If
                        End If
                        Exit For
                    End If
                Next
            Next
            .rows = .rows + 1
        Next
    End With
    
    Set GetPrintObj = vsfPrint
End Function
Private Sub ResetParams()
    mblnParamsRefresh = False
    
    With frmPIVAParaSet
        .mstrPrivs = mstrPrivs
        .mlng库房id = mParams.lng配置中心
        .Show 1, Me
    End With
    
    If mblnParamsRefresh = True Then
        Call GetParams
        
        If mcondition.lngCenterID <> mParams.lng配置中心 Then
            mcondition.lngCenterID = mParams.lng配置中心
        End If
        
        Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
        Call SetCommand
        
        DoEvents
        
        Call RefreshDeptList(0)
        
        DoEvents
        
        Call RefreshDetailList(0)
    End If
End Sub

Private Sub BillPrint_Label(ByVal lngType As Long)
    '打印标签
    Dim strInputID As String    '配药ID...
    Dim strPrintID As String    '配药ID,瓶签号|配药号,瓶签号...
    Dim lngRow As Long
    Dim strCom As String
    Dim arrParams
    Dim strMsg As String
    Dim i As Integer
    Dim str配药id As String
    Dim dateNow As Date
    Dim blnPrint As Boolean
    Dim str操作员 As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    If vsfTrans.rows = 1 Then Exit Sub
    If vsfTrans.TextMatrix(1, vsfTrans.ColIndex("配药ID")) = "" Then Exit Sub
    
    
    With vsfTrans
'        If Me.tbcDetail.Item(mDetailType.输液单卡片).Selected Then
'            For i = 1 To .rows - 1
'                If .TextMatrix(i, .ColIndex("配药ID")) = mstr配药id Then
'                    .Row = i
'                    Exit For
'                End If
'            Next
'        End If
        If lngType = conMenu_Oper_PrintLabel_SelBatch Then
            strCom = .TextMatrix(.Row, .ColIndex("配药批次"))
            strMsg = "打印当前批次为【" & .TextMatrix(.Row, .ColIndex("配药批次")) & "】的所有瓶签，是否继续？"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelDept Then
            strCom = .TextMatrix(.Row, .ColIndex("病区"))
            strMsg = "打印当前病区为【" & .TextMatrix(.Row, .ColIndex("病区")) & "】的所有瓶签，是否继续？"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelPati Then
            strCom = .TextMatrix(.Row, .ColIndex("病区")) & .TextMatrix(.Row, .ColIndex("姓名")) & .TextMatrix(.Row, .ColIndex("床号"))
            strMsg = "打印当前病人为【" & .TextMatrix(.Row, .ColIndex("姓名")) & "】的所有瓶签，是否继续？"
        ElseIf lngType = conMenu_Oper_PrintLabel_SelSendNo Then
            strCom = .TextMatrix(.Row, .ColIndex("摆药单号"))
            strMsg = "打印当前摆药单号为【" & .TextMatrix(.Row, .ColIndex("摆药单号")) & "】的所有瓶签，是否继续？"
        ElseIf lngType = conMenu_Oper_PrintLabel_AllRow Then
            strMsg = "打印当前已选择的所有瓶签，是否继续？"
        End If
        
        If lngType = conMenu_Oper_PrintLabel_SelRow Then
            '当前行
            strInputID = Val(.TextMatrix(.Row, .ColIndex("配药ID")))
            strPrintID = Val(.TextMatrix(.Row, .ColIndex("配药ID"))) & "," & .TextMatrix(.Row, .ColIndex("瓶签号"))
        Else
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                    If lngType = conMenu_Oper_PrintLabel_SelBatch Then
                        If .TextMatrix(lngRow, .ColIndex("配药批次")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelDept Then
                        If .TextMatrix(lngRow, .ColIndex("病区")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelPati Then
                        If .TextMatrix(lngRow, .ColIndex("病区")) & .TextMatrix(lngRow, .ColIndex("姓名")) & .TextMatrix(lngRow, .ColIndex("床号")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    ElseIf lngType = conMenu_Oper_PrintLabel_SelSendNo Then
                        If .TextMatrix(lngRow, .ColIndex("摆药单号")) = strCom Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    ElseIf lngType = mconMenu_File_PIVA_BillPrintLable Or lngType = conMenu_Oper_PrintLabel_AllRow Then
                        If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                            If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                                strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                                strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    End If
                End If
            Next
        End If
    End With
    
    If strPrintID = "" Then
        MsgBox "请选择要打印标签的输液单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If strMsg <> "" Then
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    If mParams.blnRePeople And (mcondition.strTransStep = M_STR_CALSS_PREPARE Or mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_SEND) Then
        str操作员 = Mid(frmPeople.ShowMe(mParams.lng配置中心), 2)
    End If
    
    dateNow = Sys.Currentdate
    
    arrParams = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrParams)
        Call RefreshPrintSign(CStr(arrParams(i)), dateNow, str操作员)
    Next
    
    DoEvents
    
    '本地数据集更新
    With mrsTrans
        arrParams = Split(strPrintID, "|")
        For lngRow = 0 To UBound(arrParams)
            If arrParams(lngRow) <> "" Then
                .Filter = "配药ID=" & Val(Split(arrParams(lngRow), ",")(0))
                
                Do While Not .EOF
                    !打印标志 = 1
                    !瓶签号 = Split(arrParams(lngRow), ",")(1)
                    .Update
                    .MoveNext
                Loop
            End If
        Next
    End With
    
    DoEvents
    
    '更新列表显示
    With vsfTrans
        str配药id = ";"
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("打印标志"))) = 0 And InStr(1, "|" & strPrintID, "|" & .TextMatrix(lngRow, .ColIndex("配药ID")) & ",") > 0 Then
                    str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ";"
                    .Row = lngRow
                    .Col = .ColIndex("打印")
                    .CellPicture = picPrint(1).Picture
                    .CellPictureAlignment = flexPicAlignCenterCenter
                End If
            End If
        Next
        .Redraw = flexRDDirect
    End With
    
    mfrmPIVCard.BatchPrint str配药id
    
    DoEvents
    
    '调用报表打印标签
    arrParams = Split(strPrintID, "|")
    For lngRow = 0 To UBound(arrParams)
        If arrParams(lngRow) <> "" Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
                "配药ID=" & Val(Split(arrParams(lngRow), ",")(0)), _
                "瓶签号=" & Split(arrParams(lngRow), ",")(1), _
                "PrintEmpty=0", 2)
        End If
    Next
    
    '打印汇总清单
    If mParams.int打印汇总 = 0 Then
        blnPrint = (MsgBox("是否打印本次打印的药品汇总清单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        
        If blnPrint Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
            "部门=" & mcondition.lngCenterID, _
            "打印时间=" & dateNow, "PrintEmpty=0", 1)
        End If
    ElseIf mParams.int打印汇总 = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_3", Me, _
        "部门=" & mcondition.lngCenterID, _
        "打印时间=" & dateNow, "PrintEmpty=0", 1)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Cancel()
    'PIVA工作：取消（根据当前步骤处理）
    Dim strID As String
    Dim lngRow As Long
    Dim strMsg As String
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim strErr As String
    Dim intRow As Integer
    Dim lng配药id As Long
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing And mcondition.strTransStep <> M_STR_CALSS_FAILAUDIT And mcondition.strTransStep <> M_STR_CALSS_PASSEDAUDIT Then Exit Sub
    
'    If MsgBox("是否" & strMsg & "？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
         With Me.vsfMedis
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) <> 0 Then
                    If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
                        If Val(.TextMatrix(lngRow, .ColIndex("医嘱id"))) <> 0 Then
                            '检查医嘱是否已经摆药
                            If Not CheckIs摆药(Val(.TextMatrix(lngRow, .ColIndex("医嘱id")))) Then
                                strID = strID & .TextMatrix(lngRow, .ColIndex("医嘱id")) & "," & .TextMatrix(lngRow, .ColIndex("标志")) & "|"
                            Else
                                strErr = "所选的医嘱中有已经摆药的医嘱，已经摆药的医嘱不能进行取消审核操作，是否继续取消其他医嘱的审核？"
                            End If
                        End If
                    Else
                        strID = strID & .TextMatrix(lngRow, .ColIndex("医嘱id")) & "," & .TextMatrix(lngRow, .ColIndex("标志")) & "|"
                    End If
                End If
            Next
        End With
        
        If strErr <> "" Then
            If MsgBox(strErr, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
        End If
    Else
        If mcondition.strTransStep <> M_STR_CALSS_PREPARE Then
            With mrsTrans
                .Filter = "执行标志=1"
                .Sort = "病区,配药批次,住院号"
                
                Do While Not .EOF
                    If InStr(1, "," & strID & ",", "," & !配药id & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & !配药id
                    End If
                    
                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_SEND Then
                        If IsOutPatient(mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!病人ID, mrsTrans!主页id, 3) = False Then Exit Sub
                        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!费用序号, 0)), 2, 2, 3) = False Then Exit Sub
                    ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                        If IsOutPatient(mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!病人ID, mrsTrans!主页id, 4) = False Then Exit Sub
                        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!费用序号, 0)), 2, 2, 4) = False Then Exit Sub
                    End If
                    
                    .MoveNext
                Loop
            End With
        
            If strID = "" Then
                MsgBox "请选择要取消的输液单据！", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            With Me.VSFLook
                For intRow = 1 To .rows - 1
                    If lng配药id <> Val(.TextMatrix(intRow, .ColIndex("配药id"))) And .TextMatrix(intRow, .ColIndex("操作状态")) = "已摆药" Then
                        lng配药id = Val(.TextMatrix(intRow, .ColIndex("配药id")))
                        strID = IIf(strID = "", "", strID & ",") & lng配药id
                    End If
                Next
            End With
        End If
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    If mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
        arrExecute = GetArrayByStr(strID, 3950, "|")
    Else
        arrExecute = GetArrayByStr(strID, 3950, ",")
    End If
    For i = 0 To UBound(arrExecute)
        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Or mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            strMsg = "取消摆药"
            
            gstrSQL = "Zl_输液配药记录_取消摆药("
            '配药ID串
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
            strMsg = "取消配药"
            
            gstrSQL = "Zl_输液配药记录_取消配药("
            '配药ID串
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
            strMsg = "取消发送"
            
            gstrSQL = "Zl_输液配药记录_取消发送("
            '配药ID串
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            gstrSQL = gstrSQL & ")"
        ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
            strMsg = "取消审核"
            
            gstrSQL = "Zl_输液配药记录_审核("
            '医嘱ID串
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "|'"
            gstrSQL = gstrSQL & ")"
        End If
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, strMsg)
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
'    If mcondition.strTransStep >= M_STR_CALSS_PREPARE Then
'
'        '本地数据集更新
'        Call DelTransRec
'
'        mrsTrans.Filter = ""
'        Call ShowTrans
'    End If
'
    '刷新
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.VSFLook.rows = 1
    End If
    
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Prepare(ByVal intType As Integer)
    'PIVA工作：摆药确认
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim strPrintID As String    '配药ID,瓶签号|配药号,瓶签号...
    Dim arrParams
    Dim blnBeginTrans As Boolean
    Dim str收发ID串 As String
    Dim arrExecute As Variant
    Dim i As Integer
    Dim arrSql As Variant
    Dim strInput As String
    Dim blnlock As Boolean
    Dim strOn As String
    Dim strOff As String
    Dim dateNow As Date
    Dim curPrepareNo As Currency
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim intCount As Integer
    Dim strExcLable As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    arrSql = Array()

    With mrsTrans
        .Filter = "执行标志=1 and 是否确认调整=0"
        If Not .EOF Then
            If MsgBox("当前还有尚未确认调整的输液单，是否摆药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        .Filter = "执行标志=1 " & IIf(intType = 1, "", " and 是否打包=1")
        .Sort = "病区,配药批次,床号"
        
        Do While Not .EOF
           
            If InStr(1, "," & str收发ID串 & ",", "," & !收发ID & ",") = 0 Then
                str收发ID串 = IIf(str收发ID串 = "", "", str收发ID串 & ",") & !收发ID
            End If
            
            .MoveNext
        Loop
    End With
    
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                    End If
                    
                    If InStr(1, "|" & strPrintID & "|", "|" & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号")) & "|") = 0 Then
                        strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                    End If
                
                    If .TextMatrix(lngRow, .ColIndex("是否锁定")) = 1 Then
                        blnlock = True
                    End If
                End If
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "请选择要摆药的输液单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '检查异常状态
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, 1)
        If Not rsExcStatus Is Nothing Then
            Do While Not rsExcStatus.EOF
                '记录异常状态的输液单id
                strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                
                '从待发送的输液单串中去掉异常状态的输液单id
                strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                
                intCount = intCount + 1
                
                '记录最多5个瓶签号用于提示
                If intCount <= 5 Then
                    strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!瓶签号
                End If
                
                rsExcStatus.MoveNext
            Loop
        End If
        
        '去掉前面的","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '组织提示内容
        If strExcLable <> "" Then
            strExcLable = "注意：以下输液单不能摆药，可能已被其他人摆药或销账！" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "还有其他" & intCount - 5 & "个输液单......"
            End If
        End If
    End If
    
    '根据异常数据和待发数据的情况分别提示
    If strExcLable <> "" Then
        '有异常数据时
        If strInputID = "" Then
            '所选择的都是异常数据时
            MsgBox strExcLable & vbCrLf & "所选择的输液单都已被其他人摆药或销账，请重新选择！", vbInformation, gstrSysName
                       
            '刷新
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
            Call RefreshDetailList(Me.tabDeptList.Selected.index)
            
            Exit Sub
        Else
            '排除异常数据外还有正常数据时
            If MsgBox(strExcLable & vbCrLf & "是否对剩余的输液单摆药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    Else
        If MsgBox("是否摆药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    
    
    '库存检查
    If CheckStock = False Then Exit Sub
    
    '自备药检查
    If Check自备药 = False Then Exit Sub
    
    '零差价管理
    If CheckPriceAdjustByID = False Then Exit Sub
    
    '检查是否结账
    mrsTrans.Filter = "执行标志=1"
    mrsTrans.Sort = "单据, NO, 费用序号"
    Do While Not mrsTrans.EOF
        If IsOutPatient(mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), 2, 2, mrsTrans!病人ID, mrsTrans!主页id, 2) = False Then Exit Sub
        If IsReceiptBalance_Charge(1, mstrPrivs, Val(mrsTrans!单据), CStr(nvl(mrsTrans!NO)), Val(nvl(mrsTrans!费用序号, 0)), 2, 2, 2) = False Then Exit Sub
        
        mrsTrans.MoveNext
    Loop
    
    '取摆药单号(汇总发药号)
    curPrepareNo = Val(zlDatabase.GetNextNo(20))
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_摆药("
        '部门ID
        gstrSQL = gstrSQL & mcondition.lngCenterID
        '配药ID
        gstrSQL = gstrSQL & ",'" & arrExecute(i) & "'"
        '摆药单号
        gstrSQL = gstrSQL & "," & curPrepareNo
        '摆药人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '摆药时间
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        gstrSQL = gstrSQL & ")"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    Next
    
    If mParams.bln审核划价单 = True Then
        arrExecute = GetArrayByStr(str收发ID串, 3950, ",")
        For i = 0 To UBound(arrExecute)
            gstrSQL = "Zl_住院记帐记录_发药审核("
            '收发ID串
            gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
            '操作员编号
            gstrSQL = gstrSQL & ",'" & gstrUserCode & "'"
            '操作员姓名
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '审核时间
            gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        Next
    End If
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    For i = 0 To UBound(arrSql)
        Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "住院记帐审核")
    Next
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '本地数据集更新
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    '刷新明细
'    Call ShowTrans
'    Call ShowSumDrug
'
    '解锁
    If blnlock Then
        Call SetLock(0, strInputID)
    End If
    '刷新
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    lblCount.Caption = "输液单：0 已：0  未：0 当前选择输液单：0"
    
    '打印摆药报表
    If mParams.int摆药后打印 = 0 Then
        blnPrint = (MsgBox("是否打印本次摆药的药品汇总清单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    ElseIf mParams.int摆药后打印 = 1 Then
        blnPrint = True
    End If
    
    If blnPrint = True Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_1", Me, _
            "部门=" & mcondition.lngCenterID, _
            "摆药时间=" & StrCurDate, "操作人员=" & gstrUserName, "PrintEmpty=0", 2)
    End If
    
    '打印瓶签
    blnPrint = False
    If mParams.int瓶签摆药后打印 = 0 Then
        blnPrint = (MsgBox("是否打印输液瓶签？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
    ElseIf mParams.int瓶签摆药后打印 = 1 Then
        blnPrint = True
    End If
    
    If blnPrint = True Then
        '调用独立窗体进行打印
        Call mfrmPrintPlan.ShowMe(Me, strInputID, mParams.intNum)
'        dateNow = Sys.Currentdate
'        arrExecute = GetArrayByStr(strInputID, 3950, ",")
'        For i = 0 To UBound(arrExecute)
'            Call RefreshPrintSign(arrExecute(i), dateNow)
'        Next
'
'        arrParams = Split(strPrintID, "|")
'        For lngRow = 0 To UBound(arrParams)
'            If arrParams(lngRow) <> "" Then
'                For i = 1 To mParams.intNum
'                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
'                        "配药ID=" & Val(Split(arrParams(lngRow), ",")(0)), _
'                        "PrintEmpty=0", 2)
'                Next
'            End If
'        Next
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PIVAWork_Approve()
    'PIVA工作：审核
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim strPrintID As String    '配药ID,瓶签号|配药号,瓶签号...
    Dim str收发ID串 As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strInput As String
    Dim strTemp As String
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
        
    Call InitSendMsgRs
    With Me.vsfMedis
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("医嘱id"))) > 0 And Val(.TextMatrix(lngRow, .ColIndex("标志"))) > 0 Then
                strInputID = strInputID & .TextMatrix(lngRow, .ColIndex("医嘱id")) & "," & .TextMatrix(lngRow, .ColIndex("标志")) & "|"
                
                If .TextMatrix(lngRow, .ColIndex("标志")) = 2 Then
                    mrsSendMsg.AddNew
                    mrsSendMsg!医嘱id = .TextMatrix(lngRow, .ColIndex("医嘱id"))
                    mrsSendMsg!发送号 = "111"
                    mrsSendMsg!病人ID = .TextMatrix(lngRow, .ColIndex("病人ID"))
                    mrsSendMsg!姓名 = .TextMatrix(lngRow, .ColIndex("名字"))
                    mrsSendMsg!住院号 = .TextMatrix(lngRow, .ColIndex("住院号"))
                    mrsSendMsg!主页id = .TextMatrix(lngRow, .ColIndex("主页id"))
                    mrsSendMsg!病区ID = .TextMatrix(lngRow, .ColIndex("病区id"))
                    mrsSendMsg!科室ID = .TextMatrix(lngRow, .ColIndex("科室id"))
                    mrsSendMsg!床号 = .TextMatrix(lngRow, .ColIndex("床号"))
                    mrsSendMsg.Update
                End If
            End If
        Next
    End With
    
    If strInputID = "" Then Exit Sub
    
    arrExecute = GetArrayByStr(strInputID, 3950, "|")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_审核("
        '医嘱ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "|'"
        '审核药师
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        gstrSQL = gstrSQL & ")"
    
        Call zlDatabase.ExecuteProcedure(gstrSQL, "首次执行医嘱审核")
    Next
    
    '发送消息
    Call SendMsgModule
    
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Delete()
    'PIVA工作：删除因医嘱回退而作废的输液配药记录
    Dim strID As String
    Dim lngRow As Long
    
    On Error GoTo errHandle
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .strID("选择"))) = -1 Then
                    If InStr(1, "," & strID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                        strID = IIf(strID = "", "", strID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                    End If
                End If
            End If
        Next
    End With
    
    If strID = "" Then
        MsgBox "请选择要删除的输液单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("是否要删除已作废的输液单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    gstrSQL = "Zl_输液配药记录_删除("
    '配药ID串
    gstrSQL = gstrSQL & "'" & strID & "'"
    gstrSQL = gstrSQL & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "删除")
    
    DoEvents
    
'    '本地数据集更新
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    '刷新明细
'    Call ShowTrans
'    Call ShowSumDrug
    
    
    '刷新
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PIVAWork_Dosage(Optional ByVal lng配药id As Long, Optional ByVal str操作说明 As String)
    'PIVA工作：配药确认
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim strPrintID As String    '配药ID,瓶签号|配药号,瓶签号...
    Dim arrParams
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim dateNow As Date
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim strExcLable As String
    Dim intCount As Integer
    Dim blnlock As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If lng配药id = 0 Then
        With vsfTrans
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                        If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                        End If
                        
                        If InStr(1, "," & strPrintID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                            strPrintID = IIf(strPrintID = "", "", strPrintID & "|") & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & "," & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                        End If
                        
                        If .TextMatrix(lngRow, .ColIndex("是否锁定")) = 1 Then
                            blnlock = True
                        End If
                    End If
                End If
            Next
        End With
    Else
        mlng已扫描 = mlng已扫描 + 1
        mlng未扫描 = mlng未扫描 - 1
        
        strInputID = lng配药id
    End If
    
    If strInputID = "" Then Exit Sub
    
    '检查异常状态
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, mTransStatus.摆药)
        If Not rsExcStatus Is Nothing Then
            If rsExcStatus.EOF And lng配药id <> 0 Then
                lblMsg.Caption = "正常"
                
                With Me.vsfDept(0)
                    For i = 1 To .rows - 1
                        If Mid(.TextMatrix(i, .ColIndex("病区")), InStr(1, .TextMatrix(i, .ColIndex("病区")), "]") + 1) = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("病区")) Then
                            .TextMatrix(i, .ColIndex("数量")) = Val(.TextMatrix(i, .ColIndex("数量"))) - 1
                            Exit For
                        End If
                    Next
                End With
            ElseIf Not rsExcStatus.EOF And lng配药id <> 0 Then
                If rsExcStatus!操作状态 = 4 Then
                    Me.lblMsg.Caption = "该瓶签已扫描"
                ElseIf rsExcStatus!操作状态 = 5 Then
                    Me.lblMsg.Caption = "该瓶签已发送"
                ElseIf rsExcStatus!操作状态 >= 9 Then
                    Me.lblMsg.Caption = "该条医嘱已停止或销账"
                ElseIf nvl(rsExcStatus!是否打包, 0) > 0 Then
                    Me.lblMsg.Caption = "该瓶签已打包"
                End If
                Exit Sub
            Else
                Do While Not rsExcStatus.EOF
                    If rsExcStatus!操作状态 <> 2 Then
                    
                        '记录异常状态的输液单id
                        strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                        
                        '从待发送的输液单串中去掉异常状态的输液单id
                        strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                        
                        intCount = intCount + 1
                        
                        '记录最多5个瓶签号用于提示
                        If intCount <= 5 Then
                            strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!瓶签号
                        End If
                    End If
                    rsExcStatus.MoveNext
                Loop
            End If
        End If
        
        '去掉前面的","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '组织提示内容
        If strExcLable <> "" Then
            strExcLable = "注意：以下输液单不能配药，可能已被其他人配药或销账！" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "还有其他" & intCount - 5 & "个输液单......"
            End If
        End If
    End If
    
    '根据异常数据和待发数据的情况分别提示
    If strExcLable = "" Then
        '无异常数据时
        If lng配药id = 0 Then If MsgBox("是否配药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        '有异常数据时
        If strInputID = "" Then
            '所选择的都是异常数据时
            MsgBox strExcLable & vbCrLf & "所选择的输液单都已被其他人配药或销账，请重新选择！", vbInformation, gstrSysName
                       
            '刷新
            Call RefreshDeptList(0)
            Call RefreshDetailList(0)
            
            Exit Sub
        Else
            '排除异常数据外还有正常数据时
            If MsgBox(strExcLable & vbCrLf & "是否对剩余的输液单配药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_配药("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '配药人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '配药时间
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        '操作说明
        gstrSQL = gstrSQL & ",'" & str操作说明 & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "配药确认")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
'
'    '本地数据集更新
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    '刷新明细
'    Call ShowTrans
'    Call ShowSumDrug
    
    '解锁
    If blnlock Then
        Call SetLock(0, strInputID)
    End If
    
    '刷新
    If lng配药id = 0 Then
        lblCount.Caption = "输液单：0 已：0  未：0 当前选择输液单：0"
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    If lng配药id <> 0 Then
        mrsTrans.Filter = ""
        mrsTrans.Filter = "配药id=" & lng配药id
        Do While Not mrsTrans.EOF
            mrsTrans.Delete (adAffectCurrent)
            mrsTrans.MoveNext
        Loop
        Call SetFilter
        lblCount.Caption = "输液单：" & mlng已扫描 + mlng未扫描 & " 已：" & mlng已扫描 & "  未：" & mlng未扫描 & " 当前选择输液单：0"
        Me.txtFindItem.SetFocus
    End If
    
    '打印瓶签
    If lng配药id = 0 Then
        blnPrint = False
        If mParams.int瓶签配药后打印 = 0 Then
            blnPrint = (MsgBox("是否打印输液瓶签？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        ElseIf mParams.int瓶签配药后打印 = 1 Then
            blnPrint = True
        End If
        
        If blnPrint = True Then
            '调用独立窗体进行打印
            Call mfrmPrintPlan.ShowMe(Me, strInputID, mParams.intNum)
    '        dateNow = Sys.Currentdate
    '        arrExecute = GetArrayByStr(strInputID, 3950, ",")
    '        For i = 0 To UBound(arrExecute)
    '            Call RefreshPrintSign(arrExecute(i), dateNow)
    '        Next
    '
    '        arrParams = Split(strPrintID, "|")
    '        For lngRow = 0 To UBound(arrParams)
    '            If arrParams(lngRow) <> "" Then
    '                For i = 1 To mParams.intNum
    '                    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_1", Me, _
    '                        "配药ID=" & Val(Split(arrParams(lngRow), ",")(0)), _
    '                        "PrintEmpty=0", 2)
    '                Next
    '            End If
    '        Next
        End If
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_ReturnVerify()
    'PIVA工作：销帐审核
    Dim strCurrent As String
    Dim str配药id As String
    Dim strNo As String
    Dim str序号数量 As String
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If MsgBox("是否销帐审核？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strCurrent = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    With mrsTrans
        .Filter = "执行标志>0"
        
        If .EOF Then
            MsgBox "请选择要销帐审核的输液单据！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查是否结账
        .Filter = "执行标志>0"
        .Sort = "单据, NO, 费用序号"
        Do While Not .EOF
            If IsOutPatient(mstrPrivs, Val(!单据), CStr(!NO), 2, 2, mrsTrans!病人ID, mrsTrans!主页id, 1) = False Then Exit Sub
            If IsReceiptBalance_Charge(1, mstrPrivs, Val(!单据), CStr(!NO), Val(!费用序号), 2, 2, 1) = False Then Exit Sub
            
            .MoveNext
        Loop
        
        '配药记录销帐处理（在"Zl_输液配药记录_销帐审核"中统一处理退药和销帐审核）
        .Filter = "执行标志>0"
        .Sort = "配药ID"
        Do While Not .EOF
            If InStr(1, str配药id, !配药id) = 0 Then
                str配药id = IIf(str配药id = "", "", str配药id & ",") & !配药id & "," & !执行标志
            End If
            
            .MoveNext
        Loop
        If str配药id <> "" Then
            gstrSQL = "Zl_输液配药记录_销帐审核("
            'str配药ID
            gstrSQL = gstrSQL & "'" & str配药id & "'"
            '审核人
            gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
            '审核时间
            gstrSQL = gstrSQL & ",To_Date('" & strCurrent & "','yyyy-MM-dd hh24:mi:ss')"
            gstrSQL = gstrSQL & ")"
            
            Call zlDatabase.ExecuteProcedure(gstrSQL, "PIVAWork_ReturnVerify")
        End If
        
        gclsInsure.InitOracle gcnOracle
        
        .Filter = "执行标志=1"
        .Sort = "NO,费用序号"
        Do While Not .EOF
            If strNo <> !NO Or str序号数量 <> !费用序号 & ":" & !实际数量 Then
                strNo = !NO
                str序号数量 = !费用序号 & ":" & !实际数量
                
                '医保处理
                If Not IsNull(!险类) And InStr(1, strMCNO, !NO) = 0 Then
                    MCPAR.记帐作废上传 = gclsInsure.GetCapability(support记帐作废上传, , Val(!险类))
                    MCPAR.记帐完成后上传 = gclsInsure.GetCapability(support记帐完成后上传, , Val(!险类))
                    strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !险类 & _
                            "," & IIf(MCPAR.记帐作废上传, "1", "0") & "," & IIf(MCPAR.记帐完成后上传, "1", "0")
                End If
            End If
            .MoveNext
        Loop
    End With
    
    '医保，记帐作废上传，作废时上传
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:
                    Screen.MousePointer = 0
                    Exit Sub
                End If
            End If
        Next
    End If
    
    '医保，记帐作废上传，完成后上传
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "单据""" & CStr(arrMCPar(0)) & """的销帐数据向医保传送失败，该单据已销帐。", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    
    If MsgBox("你需要打印退药销帐清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1345_2", Me, "退药时间=" & strCurrent, "包装系数=C.住院包装", 2)
    End If
    
'    '本地数据集更新
'    Call DelTransRec
'
'    DoEvents
'
'    mrsTrans.Filter = ""
'    '刷新明细
'    Call ShowTrans
'    Call ShowSumDrug
    
    '刷新
    Call RefreshDeptList(0)
    Call RefreshDetailList(0)
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Refuse()
'确认拒发
Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                    If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                        strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                    End If
                End If
                
            End If
        Next
    End With
    
    If strInputID = "" Then
        MsgBox "请选择要确认拒绝的输液单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("是否确认拒绝？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_确认拒绝("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '拒绝人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "确认拒绝")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '刷新
    Call RefreshDeptList(Me.tabDeptList.Selected.index)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PIVAWork_Send(Optional ByVal lng配药id As Long, Optional ByVal str操作说明 As String)
    'PIVA工作：发送确认
    Dim strInputID As String
    Dim lngRow As Long
    Dim StrCurDate As String
    Dim blnPrint As Boolean
    Dim arrExecute As Variant
    Dim i As Integer
    Dim blnBeginTrans As Boolean
    Dim rsExcStatus As ADODB.Recordset
    Dim strExcID As String
    Dim strExcLable As String
    Dim intCount As Integer
    Dim blnAutoPrint As Boolean
    Dim lng病区id As Long
    Dim str病区名称 As String
    
    On Error GoTo errHandle
    
    If mrsTrans Is Nothing Then Exit Sub
    
    If lng配药id = 0 Then
        With vsfTrans
            For lngRow = 1 To .rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) > 0 Then
                    If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                        If InStr(1, "," & strInputID & ",", "," & Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) & ",") = 0 Then
                            strInputID = IIf(strInputID = "", "", strInputID & ",") & .TextMatrix(lngRow, .ColIndex("配药ID"))
                        End If
                    End If
                    
                End If
            Next
        End With
    Else
        mlng已扫描 = mlng已扫描 + 1
        mlng未扫描 = mlng未扫描 - 1
        With Me.vsfDept(0)
            For i = 1 To .rows - 1
                If Mid(.TextMatrix(i, .ColIndex("病区")), InStr(1, .TextMatrix(i, .ColIndex("病区")), "]") + 1) = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("病区")) Then
                    .TextMatrix(i, .ColIndex("数量")) = Val(.TextMatrix(i, .ColIndex("数量"))) - 1
                    blnAutoPrint = (.TextMatrix(i, .ColIndex("数量")) = 0)
                    lng病区id = Val(.TextMatrix(i, .ColIndex("病区id")))
                    str病区名称 = vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("病区"))
                    Exit For
                End If
            Next
        End With
        strInputID = lng配药id
    End If
    
    '检查异常状态
    If strInputID <> "" Then
        strInputID = "," & strInputID
        
        Set rsExcStatus = PIVA_GetExcStatus(strInputID, mTransStatus.配药)
        If Not rsExcStatus Is Nothing Then
            If rsExcStatus.EOF And lng配药id <> 0 Then
                lblMsg.Caption = "正常"
            ElseIf Not rsExcStatus.EOF And lng配药id <> 0 Then
                If rsExcStatus!操作状态 = 5 Then
                    Me.lblMsg.Caption = "该瓶签已扫描"
                ElseIf rsExcStatus!操作状态 >= 9 Then
                    Me.lblMsg.Caption = "该条医嘱已停止或销账"
                End If
                Exit Sub
            Else
                Do While Not rsExcStatus.EOF
                    '记录异常状态的输液单id
                    strExcID = IIf(strExcID = "", "", strExcID & ",") & rsExcStatus!Id
                    
                    '从待发送的输液单串中去掉异常状态的输液单id
                    strInputID = Replace(strInputID, "," & rsExcStatus!Id, "")
                    
                    intCount = intCount + 1
                    
                    '记录最多5个瓶签号用于提示
                    If intCount <= 5 Then
                        strExcLable = IIf(strExcLable = "", "", strExcLable & vbCrLf) & rsExcStatus!瓶签号
                    End If
                    
                    rsExcStatus.MoveNext
                Loop
            End If
        End If
        
        
        '去掉前面的","
        If strInputID <> "" Then
            strInputID = Mid(strInputID, 2)
        End If
        
        '组织提示内容
        If strExcLable <> "" Then
            strExcLable = "注意：以下输液单不能发送，可能已被其他人发送或销账！" & vbCrLf & strExcLable
            
            If intCount > 5 Then
                strExcLable = strExcLable & vbCrLf & "还有其他" & intCount - 5 & "个输液单......"
            End If
        End If
    End If
    
    '根据异常数据和待发数据的情况分别提示
    If strExcLable = "" Then
        '无异常数据时
        If lng配药id = 0 Then If MsgBox("是否发送？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        '有异常数据时
        If strInputID = "" Then
            '所选择的都是异常数据时
            MsgBox strExcLable & vbCrLf & "所选择的输液单都已被其他人发送或销账，请重新选择！", vbInformation, gstrSysName
                       
            '刷新
            Call RefreshDeptList(Me.tabDeptList.Selected.index)
            Call RefreshDetailList(Me.tabDeptList.Selected.index)
            
            Exit Sub
        Else
            '排除异常数据外还有正常数据时
            If MsgBox(strExcLable & vbCrLf & "是否发送剩余的输液单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
    End If
    
    StrCurDate = Format(Sys.Currentdate, "YYYY-MM-DD HH:MM:SS")
    
    gcnOracle.BeginTrans
    blnBeginTrans = True
    
    arrExecute = GetArrayByStr(strInputID, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = "Zl_输液配药记录_发送("
        '配药ID
        gstrSQL = gstrSQL & "'" & arrExecute(i) & "'"
        '发送人
        gstrSQL = gstrSQL & ",'" & gstrUserName & "'"
        '发送时间
        gstrSQL = gstrSQL & ",To_Date('" & StrCurDate & "','yyyy-MM-dd hh24:mi:ss')"
        '操作说明
        gstrSQL = gstrSQL & ",'" & str操作说明 & "'"
        gstrSQL = gstrSQL & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "发送确认")
    Next
    
    gcnOracle.CommitTrans
    blnBeginTrans = False
    
    '打印发送报表
    If lng配药id = 0 Then
        If mParams.int发送后打印 = 0 Then
            blnPrint = (MsgBox("是否打印本次发送的药品汇总清单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        ElseIf mParams.int发送后打印 = 1 Then
            blnPrint = True
        End If
        
        If blnPrint = True Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_2", Me, _
            "部门=" & mcondition.lngCenterID, _
            "发送时间=" & StrCurDate, "操作人员=" & gstrUserName, "PrintEmpty=0", 2)
        End If
    End If
    
    '刷新
    If lng配药id = 0 Then
        lblCount.Caption = "输液单：0 已：0  未：0 当前选择输液单：0"
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
    
    If lng配药id <> 0 Then
        mrsTrans.Filter = ""
        mrsTrans.Filter = "配药id=" & lng配药id
        Do While Not mrsTrans.EOF
            mrsTrans.Delete (adAffectCurrent)
            mrsTrans.MoveNext
        Loop
        Call SetFilter
        Me.txtFindItem.SetFocus
    End If
    
    DoEvents
    If blnAutoPrint Then
        lblCount.Caption = "输液单：" & mlng已扫描 + mlng未扫描 & " 已：" & mlng已扫描 & "  未：" & mlng未扫描 & " 当前选择输液单：0"
        If Me.cboBatch.Text = "<全部>" Then Exit Sub
        If MsgBox("是否打印" & str病区名称 & "病区" & Me.cboBatch.Text & "批次发送的药品汇总清单？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1345_4", Me, _
                "部门=" & mcondition.lngCenterID, _
                "病人病区=" & lng病区id, _
                "配药批次=" & Mid(Me.cboBatch.Text, 1, 1), "PrintEmpty=0", 2)
        End If
    End If
    
    Exit Sub
errHandle:
    If blnBeginTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case mconMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case mconMenu_View_FontSize_1, mconMenu_View_FontSize_2, mconMenu_View_FontSize_3       '字体
            Control.Checked = Val(Control.Parameter) = mParams.intFont
        Case MCONMENU_EDIT_PIVA_MedicalRecord               '电子病案查阅
            If mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
    End Select
End Sub

Private Sub chkAll_Click()
    mint标志 = 0
    Chk_all
End Sub

Private Sub Chk_all()
    Dim lngRow As Long
    Dim str配药id As String
    Dim strFilter As String
    
    mstrLastLabel = ""
    With vsfTrans
        For lngRow = 1 To .rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("配药ID"))) <> 0 Then
                 str配药id = str配药id & .TextMatrix(lngRow, .ColIndex("配药ID")) & ","
                .TextMatrix(lngRow, .ColIndex("选择")) = IIf(chkAll.Value = 1, -1, 0)
                
                If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
                    If chkAll.Value = 1 Then
                        .TextMatrix(lngRow, .ColIndex("标志")) = "1"
                        .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
                    Else
                        .TextMatrix(lngRow, .ColIndex("标志")) = "0"
                        .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Nothing
                    End If
                End If
                
                mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & .TextMatrix(lngRow, .ColIndex("瓶签号"))
            End If
        Next
    End With
        
    If chkAll.Value = 1 Then
        Call UpdateExeSign(-1, 1)
        
        DoEvents
    Else
        Call UpdateExeSign(-1, 0)
    End If
    
    If mint标志 <> 1 Then
        '同步卡片的数据
        mfrmPIVCard.chkClick Me.chkAll.Value
        mint标志 = 0
    End If
End Sub

Private Sub chkDept_Click()
    Call ShowSumDrug
End Sub


Private Sub chkPack_Click()
    Call ShowSumDrug
End Sub

Private Sub chkSendType_Click(index As Integer)
    Dim n As Integer
    
    If chkSendType(0).Value = 0 And chkSendType(1).Value = 0 Then
        chkSendType(index).Value = 1
    End If
    
    If chkSendType(0).Value = 1 And chkSendType(1).Value = 1 Then
        n = 0
    ElseIf chkSendType(0).Value = 1 Then
        n = 1
    Else
        n = 2
    End If
    
    If n <> Val(fraDetailCtr.Tag) Then
        fraDetailCtr.Tag = n
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
End Sub

Private Sub chkType_Click(index As Integer)
    Dim n As Integer
    
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        chkType(index).Value = 1
    End If
    
    '同步卡片和列表的数据
    mfrmPIVCard.CheckType index, chkType(index).Value
    
    If chkType(0).Value = 1 And chkType(1).Value = 1 Then
        n = 0
    ElseIf chkType(0).Value = 1 Then
        n = 1
    Else
        n = 2
    End If
    
    If n <> Val(vsfTrans.Tag) Then
        vsfTrans.Tag = n
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)
    End If
End Sub

Private Sub cmdDrug_Click()
    Dim RecReturn As Recordset
    '获取药品管理器
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "静脉配置中心", mParams.lng配置中心, mParams.lng配置中心)
    End If
    
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, Me.txtDrug.Text, , , mParams.lng配置中心, , , 0, True, True, True, , , mstrPrivs)
    
    If Not RecReturn.EOF Then
        Me.txtDrug.Text = "(" & RecReturn!药品编码 & ")" & RecReturn!通用名
        Me.txtDrug.Tag = RecReturn!药品ID
    End If
     
End Sub

Private Sub cmdRefreshTrans_Click()
    Me.cboType.ListIndex = 0
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    If Me.cboBatch.ListIndex > 0 Then Call cboBatch_Click
    Me.txtFindItem.SetFocus
End Sub

Private Sub CmdSave_Click()
    '保存原因
    Dim strSQL As String
    
     On Error GoTo errHandle
     
     If txtLog.Text = "" Then
        MsgBox "请填写药师审核原因！", vbInformation, gstrSysName
        Exit Sub
     End If
     
     strSQL = "Zl_病人医嘱记录_SaveReason("
     strSQL = strSQL & Val(Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("医嘱id")))
     strSQL = strSQL & ",'" & txtLog.Text & "'"
     strSQL = strSQL & ")"
     
     Call zlDatabase.ExecuteProcedure(strSQL, "保存原因")
     Me.vsfMedis.TextMatrix(vsfMedis.Row, vsfMedis.ColIndex("药师审核原因")) = txtLog.Text
     
     MsgBox "药师审核原因保存成功！", vbInformation, gstrSysName
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
   SetListBar
   mblnActive = True
End Sub

Private Function ShowOhters() As Boolean
    '功能：根据当前用户环境参数设置，判断是否需要显示[不取药]和[自备药]
    Dim strSQL As String
    Dim rstemp As ADODB.Recordset
    Dim bln自备药 As Boolean
    Dim bln不取药 As Boolean
    
    On Error GoTo errHandle
    
    ShowOhters = False
    
    strSQL = "Select 1 From 输液自备药清单 Where 是否检查库存 = 1 And Rownum < 2"
    
    Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "读取输液自备药清单")
    
    bln自备药 = (Val(zlDatabase.GetPara("自备药允许发往静配中心", glngSys, 1345, 0)) = 1)
    bln不取药 = (Val(zlDatabase.GetPara("不取药允许发往静配中心", glngSys, 1345, 0)) = 1)
    
    If Not rstemp.EOF Or bln自备药 Or bln不取药 Then
        ShowOhters = True
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub Form_Load()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim i As Integer
    
    mblnLoad = True
    mblnActive = False
    
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    mintBeginRow = 0
    mintEndRow = 0
    
    lblMsgComment.Tag = 1
    picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("DownArrow").Picture
    
    mblnShowOhters = ShowOhters()
    fraTip.Visible = mblnShowOhters
    
    '读取权限和参数
    Call GetPrivs
    Call GetParams
    
    '数据检查
    If DependOnCheck = False Then
        Exit Sub
    End If
    stbThis.Panels(3).Text = gstrUserName
    
    '创建电子病案查阅对象
    If mobjCISJOB Is Nothing Then
        On Error Resume Next
        Set mobjCISJOB = CreateObject("zl9CISJob.clsCISJob")
        
        If Not mobjCISJOB Is Nothing Then
            Call mobjCISJOB.InitCISJob(gcnOracle, Me, glngSys, mstrPrivs, gobjBrower.mobjEmr)
        End If
        err.Clear
        
        On Error GoTo 0
    End If
    
    mdateToday = Sys.Currentdate
    
    '初始化窗体对象
    Set mfrmPIVCard = New frmPIVCard
    Set mfrmPlan = New frmPlan
    Set mfrmPrintPlan = New frmPrintPlan
    
    If Not mParams.bln审核 Then
        mcondition.strTransStep = "01"
        fraMedis.Visible = False
    Else
        mcondition.strTransStep = "00"
        fraMedis.Visible = True
    End If
    
    If mParams.bln审核 Then
        mcondition.strTransStep = "00"
        If mParams.intShowPass = 1 Then
            For i = 0 To Me.ImgResult.count - 1
                Me.chkResult(i).Visible = True
                Me.ImgResult(i).Visible = True
                
            Next
        Else
            For i = 0 To Me.ImgResult.count - 1
                Me.chkResult(i).Visible = False
                Me.ImgResult(i).Visible = False
                
            Next
        End If
    Else
        mcondition.strTransStep = "01"
        For i = 0 To Me.ImgResult.count - 1
            Me.chkResult(i).Visible = False
            Me.ImgResult(i).Visible = False
            
        Next
        fraMedis.Visible = False
    End If
    
    '初始化病区排序
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        mParams.int病区排序 = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "病区排序", "1")
    Else
        mParams.int病区排序 = 1
    End If
    
    mcondition.intTransTimeSel = 0
    mcondition.lngCenterID = mParams.lng配置中心
    
    '外挂接口
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    
    '加载界面控件
    InitPanes
    InitTabControl
    InitComandBars
    
    '恢复正常颜色
    picDeptList.BackColor = &HFFFFFF
    picTime.BackColor = &HFFFFFF
    lblNote.BackColor = &HFFFFFF
    lbl时间范围.BackColor = &HFFFFFF
    lblTimeBegin.BackColor = &HFFFFFF
    lblTimeEnd.BackColor = &HFFFFFF
    lblName.BackColor = &HFFFFFF
    lblDrug.BackColor = &HFFFFFF
    lblTag.BackColor = &HFFFFFF
    lbldept.BackColor = &HFFFFFF
    
    '定义不显示的列
    mstrUnVisble = "当前行;标志;是否锁定;NO;单据;剂量单位;用法;药品id;核查人;核查时间;打印标志;配药id;是否打包;原批次;抗菌药物;主页id;病人ID;警告;溶媒;背景号;对应医嘱ID;"
    mstrUnallowSetColHide = "选择;打包;批次;姓名;瓶签号;药品名称;单量;"
    
    '添加自定义报表
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '恢复窗口
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        '恢复个性化参数
        Call LoadCustomSet
        
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
        
        Call RestoreWinState(Me, App.ProductName)
    End If
    
    Call GetWorkBatchRec
    
    Call InitVsfTrans
    Call InitVsfSum
    Call InitVSFLook
    
    Call SetCommand
    
    '界面输液单排序规则
    Call SetSort
    
    Select Case mcondition.strTransStep
        Case M_STR_CALSS_AUDIT, M_STR_CALSS_PASSEDAUDIT, M_STR_CALSS_FAILAUDIT
            lblFindItem.Caption = "床号"
        Case Else
            lblFindItem.Caption = "瓶签号"
    End Select
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            cbrControl.Checked = False
            If cbrControl.Caption = lblFindItem.Caption Then
                cbrControl.Checked = True
            End If
        Next
    End If
    
    err = 0
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    Call mobjMipModule.InitMessage(glngSys, mlngMode, mstrPrivs)
    Call AddMipModule(mobjMipModule)
    
    mblnLoad = False
    
    Call Load时间范围
    
    lblSort.Visible = False
    cboSort.Visible = False
    With cboSort
        .Clear
        .AddItem "0-按已设置排序规则排序"
        .AddItem "1-按过滤药品及数量排序"
        .AddItem "2-按溶媒及数量排序"
    End With
    
    If mobjMipModule Is Nothing Then
        picMsg.Visible = False
    Else
        picMsg.Visible = True
    End If
    
    '加载图标
    For i = 0 To Me.ImgResult.count
        Me.ImgResult(i).Picture = frmPublic.imgPass.ListImages(i + 1).Picture
    Next
    
    '加载列表初始化数据
    LoadData
End Sub

Private Sub LoadCustomSet()
    Dim cbrMenu As CommandBarControl
    
    mParams.intFont = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "字体", 0))
    mParams.intAutoSelect = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "自动选择", 0))
    mParams.strVsfTrans = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "摆药界面明细", "")
    mParams.strVsfSum = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "汇总表格列宽", "")
    mParams.strVsfLook = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "已摆药表格列宽", "")
    
    Call SetFontSize
    
    mcondition.intTransTimeSel = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "时间范围", 0))
    If mcondition.intTransTimeSel < 0 Or mcondition.intTransTimeSel > 3 Then
        mcondition.intTransTimeSel = 0
    End If
    
    Set cbrMenu = cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_ShowHistory, , True)
    If Not cbrMenu Is Nothing Then
        cbrMenu.Checked = (mParams.intAutoSelect = 1)
    End If
End Sub

Private Sub SaveCustomSet()
    Dim i As Integer
    Dim str列设置 As String
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "字体", mParams.intFont
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "时间范围", mcondition.intTransTimeSel
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "自动选择", mParams.intAutoSelect
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "病区排序", mParams.int病区排序
    
    str列设置 = ""
    With Me.vsfTrans
        For i = 0 To .Cols - 1
            str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "摆药界面明细", str列设置)
    
    str列设置 = ""
    With Me.vsfSumDrug
        For i = 0 To .Cols - 1
            str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "汇总表格列宽", str列设置)
    
    str列设置 = ""
    With Me.VSFLook
        For i = 0 To .Cols - 1
            str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & .ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "已摆药表格列宽", str列设置)
End Sub
Private Function DependOnCheck() As Boolean
    '依赖数据检测
    Dim rsTmp As ADODB.Recordset
    
    DependOnCheck = False
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.ID, A.名称" & _
        " From 部门表 A, 部门性质说明 B " & _
        " Where A.ID = B.部门id And B.工作性质 = '配制中心' And " & _
        " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房') " & _
        " And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取配置中心")
    
    '当前部门
    If mParams.lng配置中心 = 0 Then
'        MsgBox "请在参数设置中设置当前的配置中心！", vbInformation, gstrSysName
        frmPIVAParaSet.Show 1, Me
        Call GetParams
        DependOnCheck = True
        Exit Function
    Else
        Do While Not rsTmp.EOF
            If mParams.lng配置中心 = rsTmp!Id Then
                mstrCenterName = rsTmp!名称
                Exit Do
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    '检查部门人员
    gstrSQL = "Select Distinct P.ID, P.名称" & _
        " From 部门表 P " & _
        " Where (P.站点 = '" & gstrNodeNo & "' Or P.站点 is Null) And P.ID In (Select Distinct A.部门id " & _
        " From 部门人员 A, 部门性质说明 B " & _
        " Where A.人员id = [1] And A.部门id = B.部门id And B.工作性质 = '配制中心' And " & _
        " B.部门id In (Select Distinct 部门id From 部门性质说明 Where 工作性质 Like '%药房')) And " & _
        " (P.撤档时间 Is Null Or P.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "取配置中心人员", glngUserId)
    
    If rsTmp.RecordCount = 0 Then
        MsgBox "你不是输液配制中心人员，不能使用本模块！", vbInformation, gstrSysName
        Exit Function
    End If
    
    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1
            Item.Handle = picCondition.hWnd
    End Select
End Sub

Private Sub InitTabControl()
    '初始化分页控件
    Dim lngColor As Long
    
    With Me.tabDeptList
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(CNUMWORK, "业务", picWork.hWnd, 0).Tag = "业务_"
        .InsertItem(CNUMLOOK, "查看", picLook.hWnd, 0).Tag = "查看_"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
'        Call SetTabColor(tabDeptList)
    End With
    
    
    With Me.tabWork
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        Me.fraH.Tag = "1"
        
        If mParams.bln审核 Then
            .InsertItem(0, "审核医嘱(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_AUDIT
            
            lblCount.Visible = False
            vsfMedis.Visible = True
            If mParams.intShowPass = 1 Then fraMedis.Visible = True
            vsfTrans.Visible = False
    
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("审")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("选择")) = True
            Me.fraH.Tag = "2"
        Else
            fraMedis.Visible = False
        End If
        
        .InsertItem(1, "摆药印签(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_PREPARE
        .InsertItem(2, "配药核查(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_DOSAGE
        .InsertItem(3, "发送核查(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_SEND
        .InsertItem(4, "销帐审核(0)", picDept(CNUMWORK).hWnd, 0).Tag = M_STR_CALSS_VERIFY
        
        .Item(1).Selected = True
        .Item(0).Selected = True

        Call SetTabColor(tabWork)
    End With
    
    
    With Me.tbcLook
        With .PaintManager
            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        Me.fraH.Tag = "1"

        If mParams.bln审核 Then
            .InsertItem(0, "审核已通过医嘱(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_PASSEDAUDIT
            .InsertItem(1, "审核未通过医嘱(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_FAILAUDIT
        
            lblCount.Visible = False
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("审")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("选择")) = True
            Me.fraH.Tag = "2"
        End If
        
        .InsertItem(2, "已发送查看(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_SENDED
        .InsertItem(3, "已签收查看(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_SIGNED
        .InsertItem(4, "拒绝签收查看(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_REFUSETOSIGN
        .InsertItem(5, "已销账审核查看(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_INVALID
        .InsertItem(6, "医嘱回退查看(0)", picDept(CNUMLOOK).hWnd, 0).Tag = M_STR_CALSS_DEVICERETURN
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        Call SetTabColor(tbcLook)
    End With
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
            .Position = xtpTabPositionTop
        End With
        
        .InsertItem(mDetailType.输液单列表, "输液单列表(&0)", picDetailList.hWnd, 0).Tag = "输液单列表_"
        .InsertItem(mDetailType.输液单卡片, "输液单卡片(&1)", mfrmPIVCard.hWnd, 0).Tag = "输液单卡片_"
        .InsertItem(mDetailType.药品汇总列表, "药品汇总列表(&2)", picDetailList.hWnd, 0).Tag = "药品汇总列表_"
        
        .Item(1).Selected = True
        .Item(0).Selected = True
        
        If mParams.bln审核 Then .Item(mDetailType.药品汇总列表).Visible = False
        If mParams.bln审核 Then .Item(mDetailType.输液单列表).Caption = "病人医嘱列表"
    End With
End Sub

Private Sub InitComandBars()
    '初始化菜单：加载全部菜单，工具栏，弹出菜单等
    Dim cbrControlMain As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrMenuBar As CommandBarPopup
    Dim cbrToolBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim cbrControlCustom As CommandBarControlCustom
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = frmPublic.imgPIVA.Icons
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsMain.ActiveMenuBar.Title = "菜单"
    Me.cbsMain.ActiveMenuBar.EnableDocking (xtpFlagStretched)
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.Id = mconMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Preview, "预览(&V)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Print, "打印(&P)")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Excel, "输出到&Excel…")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrint, "单据打印(&B)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, "打印药品摆药单(&C)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintTotal, "打印发送清单(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintReturn, "打印退药销帐清单(&W)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintSum, "打印汇总报表(&S)")
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintNext, "续打瓶签(&W)")
             
        Set cbrControlMain = .Add(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, "打印标签(&R)")
        cbrControlMain.Visible = mParams.bln瓶签手工打印
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "打印当前记录(&R)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "打印当前批次(&B)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "打印当前病区(&D)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "打印当前病人(&P)", -1, False)
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "打印当前摆药单号(&S)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "打印所有选择的记录(&A)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Parameter, "参数设置(&T)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出(&X)")
        cbrControlMain.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Cancel, "取消(&C)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Approve, "审核(&A)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Lock, "锁定(&S)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_UnLock, "解锁(&S)")
        cbrControlMain.Visible = False
        
        '调整批次
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Beach, "调整批次(&B)")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, "确认调整(&O)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Prepare, "摆药(&H)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Dosage, "配药(打包)(&R)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Send, "发送(&B)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, "确认拒绝(&R)")
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Delete, "删除(&D)")
'        cbrControlMain.BeginGroup = True
'        cbrControlMain.Visible = False

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, "销帐确认(&V)")
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_PLAN, "排班(&P)")
'        cbrControlMain.BeginGroup = True
        
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PIVA_PASS, mconMenu_Edit_PIVA_PASS)
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetMenu(glngSys, glngModul, mobjPlugIn, cbrMenuBar.CommandBar.Controls, mconMenu_Edit_PlugIn)
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PlanPopup, "排班设置(&E)", -1, False)
    cbrMenuBar.Id = mconMenu_PlanPopup
    
    With cbrMenuBar.CommandBar.Controls
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_DESK, "配液台设置(&D)")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_DESKDRUG, "配液台药品对照(&M)")
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_PLAN_PIVA_PERWORK, "人员安排设置(&P)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.Id = mconMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_ToolBar, "工具栏(&T)")
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Text, "文本标签(&T)", -1, False)
        cbrControl.Checked = True
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_ToolBar_Size, "大图标(&B)", -1, False)
        cbrControl.Checked = True
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_StatusBar, "状态栏(&S)")
        cbrControlMain.Checked = True
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_View_FontSize, "字体(&F)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_1, "小字体(&S)", -1, False)
        cbrControl.Parameter = 0
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_2, "中字体(&M)", -1, False)
        cbrControl.Parameter = 1
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_View_FontSize_3, "大字体(&B)", -1, False)
        cbrControl.Parameter = 2
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_ShowHistory, "自动勾选上次选择的输液单(&A)")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SORTSET, "设置排序规则(&S)")
    End With
    
    Set cbrMenuBar = Me.cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.Id = mconMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_Help, "帮助主题(&H)")
        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB上的中联")
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "中联主页(&H)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "中联论坛(&F)", -1, False
        cbrControlMain.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Help_About, "关于(&A)…")
        cbrControlMain.BeginGroup = True
    End With
    
    '快键绑定
    With Me.cbsMain.KeyBindings
'        .Add FCONTROL, Asc("S"), mconMenu_Edit_Save
'        .Add FCONTROL, Asc("Z"), mconMenu_Edit_Untread
'        .Add FCONTROL, Asc("M"), mconMenu_Edit_Modify
'        .Add FSHIFT, VK_DELETE, mconMenu_Edit_Delete
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
    End With

    '设置不常用菜单
    With Me.cbsMain.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With
    
    '设置病区排列方式菜单
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_SortPopup, "病人排序(&P)", -1, False)
    cbrMenuBar.Id = mconMenu_SortPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByCode, "按编码排序(&0)")
        cbrControlMain.Checked = (mParams.int病区排序 = 1)
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_SortPopup_ByName, "按名称排序(&1)")
        cbrControlMain.Checked = (mParams.int病区排序 = 2)
    End With
    
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Cancel, "取消")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButtonPopup, mconMenu_File_PIVA_BillPrintLable, "打印标签")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = mParams.bln瓶签手工打印
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "打印当前记录(&R)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "打印当前批次(&B)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "打印当前病区(&D)", -1, False)
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "打印当前病人(&P)", -1, False)
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "打印当前摆药单号(&S)", -1, False)
        cbrControl.BeginGroup = True
        
        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "打印所有选择的记录(&A)", -1, False)
        cbrControl.BeginGroup = True
         
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_PIVA_BillPrintWait, "打印摆药单")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Approve, "审核")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
               
        
        '解锁,锁定按钮
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Lock, "锁定")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_UnLock, "解锁")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Beach, "调整批次")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_SURE, "确认调整")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Prepare, "摆药")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Dosage, "配药(打包)")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Send, "发送")
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_REFUSE, "确认拒绝")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_Delete, "删除")
'        cbrControlMain.BeginGroup = True
'        cbrControlMain.Visible = False

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_ReVerify, "销帐确认")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = False
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Edit_PIVA_PASS, "过敏史/病生状态")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Visible = (mParams.intShowPass = 1 And IsInString(gstrprivs, "合理用药监测", ";"))
        
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_View_Refresh, "刷新")
        cbrControlMain.BeginGroup = True
        
        '电子病案查阅
        Set cbrControlMain = .Add(xtpControlButton, MCONMENU_EDIT_PIVA_MedicalRecord, "电子病案查阅")
        cbrControlMain.BeginGroup = True
        
        '外挂部件有扩展功能
        Call zlPlugIn_SetToolbar(glngSys, glngModul, mobjPlugIn, cbrToolBar.Controls, mconMenu_Edit_PlugIn)

        Set cbrControlMain = .Add(xtpControlButton, mconMenu_File_Exit, "退出")
        cbrControlMain.BeginGroup = True
        
        Set cbrControlCustom = .Add(xtpControlCustom, mconMenu_View_Find, "查找")
        cbrControlCustom.Handle = picFind.hWnd
        cbrControlCustom.Flags = xtpFlagRightAlign
    End With
    
    For Each cbrControlMain In cbrToolBar.Controls
        cbrControlMain.Style = xtpButtonIconAndCaption
    Next
    
    '设置弹出菜单
    '打印标签
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_OperPopup, "操作(&O)", -1, False)
    cbrMenuBar.Id = conMenu_OperPopup
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_Select, "选择(&S)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelRow, "选择当前记录(&R)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelBatch, "选择当前批次(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelDept, "选择当前病区(&D)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_CancleSelDept, "取消选择当前病区(&C)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelPati, "选择当前病人(&P)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_CancleSelPati, "取消选择当前病人(&R)")
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelSendNo, "选择当前摆药单号(&S)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelMed, "选择所有抗菌药物(&M)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Select_SelAll, "选择所有记录(&A)")
        cbrControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_PrintLabel, "打印标签(&P)")
        objPopup.IconId = mconMenu_File_PIVA_BillPrintLable
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelRow, "打印当前记录(&R)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelBatch, "打印当前批次(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelDept, "打印当前病区(&D)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelPati, "打印当前病人(&P)")
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_SelSendNo, "打印当前摆药单号(&S)")
        cbrControl.BeginGroup = True
        
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_PrintLabel_AllRow, "打印所有选择的记录(&A)")
        cbrControl.BeginGroup = True
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_Bag, "打包(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Bag_Batch, "打包当前批次(&B)")
        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_Bag_All, "打包所有记录(&B)")
        
        '查看病人医嘱等相关信息
        Set cbrControl = .Add(xtpControlButton, conMenu_Oper_Look, "电子病案查阅(&I)")
                
'        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Oper_DelBatch, "删除批次(&D)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelRow, "删除当前行批次(&R)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelBatch, "删除当前批次(&B)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelDept, "删除当前行病区批次(&D)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_SelPati, "删除当前行病人批次(&P)")
'        Set cbrControl = objPopup.CommandBar.Controls.Add(xtpControlButton, conMenu_Oper_DelBatch_AllRow, "删除所有选择的行批次(&A)")
    End With
    
    '设置弹出菜单，PASS
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS（&P)", 1, False)
    cbrMenuBar.Id = mconMenu_PASS
    cbrMenuBar.Visible = False
'    With cbrMenuBar.CommandBar.Controls
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 0, "药物临床信息参考(&C)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 1, "药品说明书(&D)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 2, "中国药典(&N)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 3, "病人用药教育(&S)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 4, "检验值(&T)")
'
'        Set cbrControlMain = .Add(xtpControlPopup, mconMenu_PASS_Item + 5, "专项信息(&P)")
'        cbrControlMain.BeginGroup = True
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 0, "药物-药物相互作用(&D)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 1, "药物-食物相互作用(&F)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 2, "国内注射剂配伍(&M)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 3, "国外注射剂配伍(&T)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 4, "禁忌症(&C)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 5, "副作用(&S)", -1, False)
'
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 6, "老年人用药(&G)", -1, False)
'        cbrControl.BeginGroup = True
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 7, "儿童用药(&P)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 8, "妊娠期用药(&E)", -1, False)
'        Set cbrControl = cbrControlMain.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 9, "哺乳期用药(&L)", -1, False)
'
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 6, "医药信息中心(&I)")
'        cbrControlMain.BeginGroup = True
'
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 7, "药品配对信息(&M)")
'        cbrControlMain.BeginGroup = True
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 8, "给药途径配对信息(&R)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 9, "医院药品信息(&F)")
'        Set cbrControlMain = .Add(xtpControlButton, mconMenu_PASS_Item + 10, "审查(&S)")
'    End With
    
    '设置弹出菜单，查找
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_Look, "查找项目（&F)", 1, False)
    cbrMenuBar.Id = mconMenu_Look
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 1, "瓶签号")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 2, "住院号")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 3, "床号")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 4, "姓名")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Look + 5, "摆药单号")
    End With
    
     '设置姓名过滤菜单，过滤
    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_Filter, "过滤项目（&F)", 1, False)
    cbrMenuBar.Id = mconMenu_Filter
    cbrMenuBar.Visible = False
    With cbrMenuBar.CommandBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 1, "姓名")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 2, "住院号")
        Set cbrControlMain = .Add(xtpControlButton, mconMenu_Filter + 3, "床号")
    End With
End Sub

Private Sub Load时间范围()
    Dim dteTime As Date
    
    dteTime = Sys.Currentdate
    Dtp开始时间.Value = Format(dteTime, "yyyy-MM-dd") & " 00:00:00"
    Dtp结束时间.Value = Format(dteTime, "yyyy-MM-dd") & " 23:59:59"
    
    With cbo时间范围
        .Clear
        .AddItem "0-今日"
        .AddItem "1-明日"
        .AddItem "2-今日和明日"
        .AddItem "3-指定时间范围"
    End With
    
    With cbo时间范围
        .ListIndex = mcondition.intTransTimeSel
        
        If .ListIndex <> Val(.Tag) Then
            .Tag = .ListIndex
        End If
        
        If .ListIndex = 0 Then
            Dtp开始时间.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp结束时间.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 1 Then
            Dtp开始时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 00:00:00")
            Dtp结束时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 2 Then
            Dtp开始时间.Value = CDate(Format(dteTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp结束时间.Value = CDate(Format(DateAdd("D", 1, dteTime), "YYYY-MM-DD") & " 23:59:59")
        ElseIf .ListIndex = 3 Then
            If mcondition.strTransStartTime = "" Then
                mcondition.strTransStartTime = Format(dteTime, "YYYY-MM-DD") & " 00:00:00"
            End If
            If mcondition.strTransEndTime = "" Then
                mcondition.strTransEndTime = Format(dteTime, "YYYY-MM-DD") & " 23:59:59"
            End If
            
            Dtp开始时间.Value = CDate(Format(mcondition.strTransStartTime, "YYYY-MM-DD") & " 00:00:00")
            Dtp结束时间.Value = CDate(Format(mcondition.strTransEndTime, "YYYY-MM-DD") & " 23:59:59")
        End If
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 9000 Then Me.Height = 9000
    If Me.Width < 12000 Then Me.Width = 12000
    
    ResizeConditionArea
    picDetailList_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsDeptAdvice = Nothing
    Set mrsTrans = Nothing
    Set mrsDeptTrans = Nothing
    
    Set mobjCISJOB = Nothing
        
    '保存窗口及参数
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        
        Call SaveWinState(Me, App.ProductName)
    
        '保存个性化设置
        Call SaveCustomSet
    End If
    
    '卸载消息对象
    If Not mobjMipModule Is Nothing Then
        Call mobjMipModule.CloseMessage
        Call DelMipModule(mobjMipModule)
        Set mobjMipModule = Nothing
    End If
    mcondition.strTransStep = ""
    
    If Not mfrmPIVCard Is Nothing Then Unload mfrmPIVCard
    If Not mfrmPlan Is Nothing Then Unload mfrmPlan
    If Not mfrmPrintPlan Is Nothing Then Unload mfrmPrintPlan
    
    '卸载外挂接口
    Call zlPlugIn_Unload(mobjPlugIn)
    
    Unload Me
End Sub

Private Sub lblTransDrug_Click()
'    If Val(picHscTransDrug.Tag) = "1" Then
'        Call imgDown_Click
'    Else
'        Call imgUp_Click
'    End If
End Sub

Private Sub fraH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfTrans.Height + y <= 1200 Then Exit Sub
        If VSFLook.Height - y < 1200 Then Exit Sub

        fraH.Top = fraH.Top + y
        VSFLook.Top = VSFLook.Top + y
        VSFLook.Height = VSFLook.Height - y
        vsfTrans.Height = vsfTrans.Height + y
        
        txtLog.Top = txtLog.Top + y
        txtLog.Height = txtLog.Height - y
        Me.vsfMedis.Height = vsfMedis.Height + y
        Me.Refresh
    End If
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfTrans.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfTrans.ColHidden(.RowData(i)) Or vsfTrans.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfTrans.Top + vsfTrans.RowHeight(0) + 30
                
                If .Top + .Height > Me.ScaleHeight - vsfTrans.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfTrans.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = vsfTrans.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub


Private Sub lblFindItem_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button = 1 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
        If Not objPopup Is Nothing Then
            For Each cbrControl In objPopup.CommandBar.Controls
                If cbrControl.Caption = "摆药单号" Then
                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_PREPARE _
                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                ElseIf cbrControl.Caption = "瓶签号" Then
                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                Else
                    cbrControl.Visible = True
                End If
            Next
                
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub lblName_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    
    If Button = 1 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Filter)
        If Not objPopup Is Nothing Then
'            For Each cbrControl In objPopup.CommandBar.Controls
'                If cbrControl.Caption = "摆药单号" Then
'                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_PREPARE _
'                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
'                        cbrControl.Visible = False
'                    Else
'                        cbrControl.Visible = True
'                    End If
'                ElseIf cbrControl.Caption = "瓶签号" Then
'                    If mcondition.strTransStep = M_STR_CALSS_AUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
'                        Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
'                        cbrControl.Visible = False
'                    Else
'                        cbrControl.Visible = True
'                    End If
'                Else
'                    cbrControl.Visible = True
'                End If
'            Next
                
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub optShowType_Click(index As Integer)
    Call SetTransColHide
End Sub

Private Sub picCondition_Resize()
    Call ResizeConditionArea
End Sub

Private Sub picDept_Resize(index As Integer)
    On Error Resume Next

    With vsfDept(index)
        .Move picDept(index).ScaleLeft, picDept(index).ScaleTop, picDept(index).ScaleWidth, picDept(index).ScaleHeight
    End With
End Sub

Private Sub picDeptList_Resize()
    On Error Resume Next
    
    With Me.tabDeptList
        .Move picDeptList.ScaleLeft, picDeptList.ScaleTop + 150, picDeptList.ScaleWidth, picDeptList.ScaleHeight - 150
    End With
    
    With Me.cmdRefreshTrans
        .Move picDeptList.Width - .Width - 50, Me.tabDeptList.Top - 50
    End With
    
    With Me.chkAllDept
'        .Move picDeptList.Width - .Width - 50, Me.tabDeptList.Top + 50
        .Move cmdRefreshTrans.Left - .Width - 50, Me.tabDeptList.Top + 50
    End With
End Sub


Private Sub picDetail_Resize()
    On Error Resume Next
    
    With fraLineV1
'        .Top = 0
        .Left = 0
        .Height = picDetail.Height + 100
    End With
    
    With tbcDetail
        .Top = 0
        .Left = fraLineV1.Left + 50
        .Width = picDetail.Width - fraLineV1.Width
        .Height = picDetail.Height - 50
    End With
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    Me.picDetail.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    
    With fraTip
        .ZOrder 0
        .Top = stbThis.Top + 90
        .Left = stbThis.Panels(2).Left + stbThis.Panels(2).Width - .Width - 50
    End With
End Sub

Private Sub picDetailList_Resize()
    On Error Resume Next
    
    With picHelp
        .Top = 0
        .Left = 0
        .Width = picDetailList.Width
    End With
     
    With fraDetailCtr
        .Top = picHelp.Top + picHelp.Height
        .Left = 0
        .Width = picDetailList.Width - 50
    End With
    
    With Me.fraMedis
        .Top = picHelp.Top + picHelp.Height
        .Left = 0
        .Width = picDetailList.Width - 50
    End With
    
    With vsfTrans
        .Top = fraDetailCtr.Top + fraDetailCtr.Height + 50
        .Left = 0
        .Width = picDetailList.Width - 50
'        .Height = picDetailList.Height - .Top
    End With
    
    With Me.vsfMedis
        .Top = IIf(fraMedis.Visible, fraMedis.Top + fraMedis.Height, picHelp.Top + picHelp.Height) + 50
        .Left = fraDetailCtr.Left
        .Width = Me.vsfTrans.Width
        .Height = picDetailList.Height - 100 - IIf(fraMedis.Visible, fraMedis.Height, 0)
    End With
    
    With vsfSumDrug
        .Top = fraDetailCtr.Top + fraDetailCtr.Height + 50
        .Left = 0
        .Width = picDetailList.Width
        .Height = picDetailList.Height - .Top
    End With
    
    If fraH.Tag = "1" Then
        VSFLook.Visible = True
        Me.fraH.Visible = True
        txtLog.Visible = False
        CmdSave.Visible = False
        Me.txtDia.Visible = False
        Me.lblDia.Visible = False
        Me.lblLog.Visible = False
        
        VSFLook.Top = picDetailList.Height - VSFLook.Height + 20
        VSFLook.Left = 0
        VSFLook.Width = picDetailList.Width - 50
        
        fraH.Top = VSFLook.Top - fraH.Height - 50
        fraH.Left = VSFLook.Left
        fraH.Width = VSFLook.Width
        vsfTrans.Height = fraH.Top - vsfTrans.Top - 50
'        VSFLook.Height = picDetailList.Height - fraH.Top - fraH.Height - 50
    ElseIf fraH.Tag = "2" Then
        txtLog.Text = ""
        txtDia.Text = ""
        VSFLook.Visible = False
        Me.fraH.Visible = False
        txtLog.Visible = True
        CmdSave.Visible = True
        Me.txtDia.Visible = True
        Me.lblDia.Visible = True
        Me.lblLog.Visible = True
        
        
        
        Me.txtLog.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50
        Me.lblLog.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height
        
        Me.txtDia.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height - Me.txtDia.Height
        
        Me.lblDia.Top = picDetailList.Height - txtLog.Height - IIf(Me.tabDeptList.Selected.index = 1, 0, CmdSave.Height) - 50 - Me.lblLog.Height - Me.txtDia.Height - lblDia.Height
        
'        txtLog.Height = VSFLook.Height - CmdSave.Height - 50
        txtLog.Left = 0
        txtLog.Width = picDetailList.Width - 50
        txtDia.Left = 0
        txtDia.Width = picDetailList.Width - 50
        lblDia.Left = 0
        lblLog.Left = 0
        
        CmdSave.Left = txtLog.Left + txtLog.Width / 2
        CmdSave.Top = txtLog.Height + txtLog.Top + 50
        
        fraH.Top = lblDia.Top - fraH.Height - 50
        fraH.Left = txtLog.Left
        fraH.Width = txtLog.Width
        vsfMedis.Height = fraH.Top - vsfMedis.Top - 50
        
        If Me.tabDeptList.Selected.index = 1 Then
            txtLog.Enabled = False
        Else
            txtLog.Enabled = True
        End If
    Else
        VSFLook.Visible = False
        txtLog.Visible = False
        Me.txtDia.Visible = False
        Me.lblDia.Visible = False
        Me.lblLog.Visible = False
        Me.fraH.Visible = False
        CmdSave.Visible = False
        vsfTrans.Height = picDetailList.Height - fraDetailCtr.Height - 300
        vsfMedis.Height = picDetailList.Height - fraDetailCtr.Height - 300
    End If
    
End Sub

Private Sub picHelp_Resize()
    Me.lblCount.Left = picHelp.Width - lblCount.Width - 50
End Sub


Private Sub picLook_Resize()
    On Error Resume Next
     
    Me.tbcLook.Move picLook.ScaleLeft, picLook.ScaleTop, picLook.ScaleWidth, picLook.ScaleHeight
    err.Clear
End Sub

Private Sub picMsg_Resize()
    On Error Resume Next

    With vsfMsg
        .Move picMsg.ScaleLeft, picMsg.ScaleTop + lblMsgComment.Top + lblMsgComment.Height + 100, picMsg.ScaleWidth, .Height
    End With
    
    Me.picUpOrDown.Left = picMsg.Width - picUpOrDown.Width - 50
    fraMsg.Width = picMsg.Width
    
    err.Clear
End Sub

Private Sub picUpOrDown_Click()
    If Me.lblMsgComment.Tag = "1" Then
        Me.lblMsgComment.Tag = "0"
        vsfMsg.Visible = False
        picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("UpArrow").Picture
    Else
        Me.lblMsgComment.Tag = "1"
        vsfMsg.Visible = True
        picUpOrDown.Picture = frmPublic.ImgList.ListImages.Item("DownArrow").Picture
    End If
    
    Call ResizeConditionArea
End Sub

Private Sub picUpOrDown1_Click()
    If Me.txtTag.Visible = False Then
        lblName.Visible = True
        txtName.Visible = True
        
        lblDrug.Visible = True
        txtDrug.Visible = True
        
        lblTag.Visible = True
        txtTag.Visible = True
        
        lbldept.Visible = True
        txtdept.Visible = True
    Else
        lblName.Visible = False
        txtName.Visible = False
        
        lblDrug.Visible = False
        txtDrug.Visible = False
        
        lblTag.Visible = False
        txtTag.Visible = False
        
        lbldept.Visible = False
        txtdept.Visible = False
    End If
    
     Call ResizeConditionArea
End Sub

Private Sub picWork_Resize()
    On Error Resume Next
     
    Me.tabWork.Move picWork.ScaleLeft, picWork.ScaleTop, picWork.ScaleWidth, picWork.ScaleHeight
    err.Clear
End Sub

Private Sub tabDeptList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnLoad = True Then Exit Sub
    
    chkAllDept.Value = 0
    
    If Item.index = 0 Then
        mcondition.strTransStep = tabWork.Selected.Tag
        Call tabWork_SelectedChanged(tabWork.Selected)
    Else
        mcondition.strTransStep = tbcLook.Selected.Tag
        Call tbcLook_SelectedChanged(tbcLook.Selected)
    End If
    
    DoEvents
    Call SetCommand
    DoEvents
    
    Call RefreshDeptList(Item.index)
    
    Select Case mcondition.strTransStep
        Case M_STR_CALSS_AUDIT, M_STR_CALSS_PASSEDAUDIT, M_STR_CALSS_FAILAUDIT
            lblFindItem.Caption = "床号"
        Case Else
            lblFindItem.Caption = "瓶签号"
    End Select
    
End Sub

'Private Sub picHscTransDrug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 1 Then
'        If vsfTrans.Height + Y <= 1200 Then Exit Sub
'        If vsfDrug.Height - Y < 1200 Then Exit Sub
'
'        picHscTransDrug.Top = picHscTransDrug.Top + Y
'        vsfDrug.Top = vsfDrug.Top + Y
'        vsfDrug.Height = vsfDrug.Height - Y
'
'        vsfTrans.Height = vsfTrans.Height + Y
'        Me.Refresh
'    End If
'End Sub
    

Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.index
        Case mDetailType.输液单列表
            If (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) And mParams.bln审核 Then
                Me.vsfMedis.Visible = True
                vsfTrans.Visible = False
                fraDetailCtr.Visible = False
            Else
                Me.vsfMedis.Visible = False
                vsfTrans.Visible = True
                fraDetailCtr.Visible = True
            End If
            
            vsfSumDrug.Visible = False
        Case mDetailType.药品汇总列表
            vsfTrans.Visible = False
            vsfSumDrug.Visible = True
            vsfMedis.Visible = False
            fraDetailCtr.Visible = True
            Call ShowSumDrug
    End Select
    
    Call SetListBar
    Call ShowComment(Item.index, mcondition.strTransStep)
End Sub

Private Sub tabWork_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intCount As Integer
    Dim intSum As Integer
    Dim strUnvisble As String
    Dim strRows As String
    Dim i As Integer
    
    If mblnLoad = True Then Exit Sub
    
    lblCount.Visible = True
    mcondition.strTransStep = Item.Tag
    
    Call SetTabColor(tabWork)
    
    DoEvents
    Call SetCommand
    DoEvents
    
    mstrLastLabel = ""
    
    chkAllDept.Value = 0
    chkAll.Value = 0
    
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If

    If mParams.bln审核 Then
        If mcondition.strTransStep >= M_STR_CALSS_PREPARE Then
            lblFindItem.Caption = "瓶签号"
            fraMedis.Visible = False
            vsfMedis.Visible = False
            vsfTrans.Visible = True
            fraDetailCtr.Visible = True
            Me.tbcDetail.Item(mDetailType.药品汇总列表).Visible = True
            lblCount.Visible = True
        Else
            lblFindItem.Caption = "床号"
            If mParams.intShowPass = 1 Then
                For i = 0 To Me.ImgResult.count - 1
                    Me.chkResult(i).Visible = True
                    Me.ImgResult(i).Visible = True
                Next
            End If
            fraMedis.Visible = True
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            Me.tbcDetail.Item(mDetailType.药品汇总列表).Visible = False
            lblCount.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("审")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("选择")) = True
        End If
    Else
        fraMedis.Visible = False
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.fraH.Tag = "1"
    ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        Me.fraH.Tag = "2"
    Else
        Me.fraH.Tag = ""
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        Me.lblVolu.Visible = True
    Else
        Me.lblVolu.Visible = False
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        Me.lblNote.Caption = "输液单销帐申请时间范围"
    Else
        Me.lblNote.Caption = "输液单执行时间范围"
    End If
    
    If mParams.bln审核 And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "病人医嘱列表"
    Else
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "输液单列表"
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
        strUnvisble = mstrUnVisble & "审;摆药人;作废类型;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
        strUnvisble = mstrUnVisble & "审;配药人;作废类型;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
        strUnvisble = mstrUnVisble & "审;摆药人;作废类型;摆药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;配药类型;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        strUnvisble = mstrUnVisble & "选择;作废类型;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐审核人;销帐审核时间;配药类型;"
    End If
    
    Me.VSFLook.rows = 1
    Call picDetailList_Resize
    
    Call SetListBar
    
    vsfColSel.Visible = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        strRows = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\表格", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = strUnvisble
    End If
    
    If strRows <> "" Then
        For i = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(i) & ";") > 0 Then
                vsfTrans.ColHidden(i) = True
            Else
                vsfTrans.ColHidden(i) = False
            End If
        Next
    End If
    
    Call SetLookMenu
'    Call SetTransColHide
    Call InitColSelList(strUnvisble)
    Call SetSumDrugColHide
    Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
    
    Set mrsTrans = Nothing
    Call ShowDeptTrans(Me.tabDeptList.Selected.index, tabWork.Selected.Tag)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
    
End Sub

Private Sub tbcLook_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim intCount As Integer
    Dim intSum As Integer
    Dim strUnvisble As String
    Dim strRows As String
    Dim i As Integer
    
    If mblnLoad = True Then Exit Sub
    
    lblCount.Visible = True
    mcondition.strTransStep = Item.Tag
    
    Call SetTabColor(tbcLook)
    
    DoEvents
    Call SetCommand
    DoEvents
    
    mstrLastLabel = ""
   
    chkAllDept.Value = 0
    chkAll.Value = 0
    
    If Me.cboType.ListCount <> 0 Then
        Me.cboType.ListIndex = 0
    End If
    
    If mParams.bln审核 Then
        If mcondition.strTransStep >= M_STR_CALSS_SENDED Then
            lblFindItem.Caption = "瓶签号"
            fraMedis.Visible = False
            vsfMedis.Visible = False
            vsfTrans.Visible = True
            fraDetailCtr.Visible = True
            Me.tbcDetail.Item(mDetailType.药品汇总列表).Visible = True
            lblCount.Visible = True
        Else
            lblFindItem.Caption = "床号"
            If mParams.intShowPass = 1 Then
                For i = 0 To Me.ImgResult.count - 1
                    Me.chkResult(i).Visible = True
                    Me.ImgResult(i).Visible = True

                Next
            End If
            fraMedis.Visible = True
            vsfMedis.Visible = True
            vsfTrans.Visible = False
            fraDetailCtr.Visible = False
            vsfSumDrug.Visible = False
            Me.tbcDetail.Item(mDetailType.药品汇总列表).Visible = False
            lblCount.Visible = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("选择")) = False
            vsfMedis.ColHidden(vsfMedis.ColIndex("审")) = True
        End If
    Else
        fraMedis.Visible = False
'        For i = 0 To Me.ImgResult.count - 1
'            Me.chkResult(i).Visible = False
'            Me.ImgResult(i).Visible = False
'
'        Next
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
        Me.fraH.Tag = "2"
    Else
        Me.fraH.Tag = ""
    End If
    
    Me.lblVolu.Visible = False
    If mcondition.strTransStep = M_STR_CALSS_VERIFY Then
        Me.lblNote.Caption = "输液单销帐申请时间范围"
    Else
        Me.lblNote.Caption = "输液单执行时间范围"
    End If
                                                                                          
    If mParams.bln审核 And (mcondition.strTransStep = M_STR_CALSS_AUDIT Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "病人医嘱列表"
    Else
        Me.tbcDetail.Item(mDetailType.输液单列表).Caption = "输液单列表"
    End If
    
    If mcondition.strTransStep = M_STR_CALSS_INVALID Then
        strUnvisble = mstrUnVisble & "审;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐原因;配药类型;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_SENDED Then
        strUnvisble = mstrUnVisble & "审;摆药人;作废类型;摆药时间;配药人;配药时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;配药类型;"
    ElseIf mcondition.strTransStep = M_STR_CALSS_DEVICERETURN Then
        strUnvisble = mstrUnVisble & "审;摆药单号;作废类型;医嘱发送时间;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;"
    Else
        strUnvisble = mstrUnVisble & "审;作废类型;摆药人;摆药时间;配药人;配药时间;发送人;发送时间;销帐申请人;销帐申请时间;销帐审核人;销帐审核时间;销帐原因;配药类型;"
    End If
    
    Me.VSFLook.rows = 1
    Call picDetailList_Resize
    
    Call SetListBar
    
    
    vsfColSel.Visible = False
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        strRows = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\表格", mcondition.strTransStep, "")
    End If
    
    If strRows = "" Then
        strRows = strUnvisble
    End If
    
    If strRows <> "" Then
        For i = 1 To Me.vsfTrans.Cols - 1
            If InStr(1, ";" & strRows & ";", ";" & vsfTrans.ColKey(i) & ";") > 0 Then
                vsfTrans.ColHidden(i) = True
            Else
                vsfTrans.ColHidden(i) = False
            End If
        Next
    End If
'    Call SetTransColHide
    
    Call SetLookMenu
    
    Call InitColSelList(strUnvisble)
    Call SetSumDrugColHide
    Call ShowComment(tbcDetail.Selected.index, mcondition.strTransStep)
    
    Set mrsTrans = Nothing
    Call ShowDeptTrans(Me.tabDeptList.Selected.index, tbcLook.Selected.Tag)
    Call RefreshDetailList(Me.tabDeptList.Selected.index)
End Sub


Private Sub SetLookMenu()
    Dim cbrControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_Look)
    
    If Not objPopup Is Nothing Then
        For Each cbrControl In objPopup.CommandBar.Controls
            cbrControl.Checked = False
            If cbrControl.Caption = lblFindItem.Caption Then
                cbrControl.Checked = True
            End If
        Next
    End If
End Sub


Private Sub txtdept_GotFocus()
    Call zlControl.TxtSelAll(txtdept)
End Sub

Private Sub txtdept_KeyPress(KeyAscii As Integer)
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    Dim vRect As RECT
    Dim rstemp As ADODB.Recordset
    
    Me.txtdept.Tag = ""
    On Error GoTo errHandle
    If KeyAscii = 13 Then
        If txtdept.Text <> "" Then
            gstrSQL = " Select A.ID,b.名称 As 站点名称, b.编号 As 站点,A.编码||'-'||A.名称 科室 From 部门表 A, Zlnodelist B " & _
                " Where a.站点 = b.编号(+) And A.ID in (Select 部门ID From 部门性质说明 Where 工作性质='临床' And 服务对象 IN(2,3))" & _
                " And (A.撤档时间 Is Null Or A.撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                " And (A.编码 Like [1] Or A.名称 Like [1] Or A.简码 Like [1])"
        Else
            Exit Sub
        End If
        
        gstrSQL = gstrSQL & " Order By a.编码 || '-' || a.名称 "
        
        '有多条记录则显示，供选择
        vRect = zlControl.GetControlRect(txtdept.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = txtdept.Height
        
        Set rstemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "选择病区或科室", False, "", "选择病区或科室", False, False, True, sngX, sngY, sngH, True, False, False, UCase(txtdept.Text) & "%")
        
        If Not rstemp Is Nothing Then
            If Not rstemp.EOF Then
                txtdept.Tag = rstemp!Id
                txtdept.Text = rstemp!科室
            End If
        End If
        
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDrug_GotFocus()
    Call zlControl.TxtSelAll(txtDrug)
End Sub

Private Sub txtDrug_KeyPress(KeyAscii As Integer)
    Dim rsReturn As Recordset
    
    Me.txtDrug.Tag = ""
    If KeyAscii = 13 Then
    
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(2, "静脉配置中心", mParams.lng配置中心, mParams.lng配置中心)
        End If
    
        Set rsReturn = frmSelector.ShowMe(Me, 1, 1, Me.txtDrug.Text, , , mParams.lng配置中心, mParams.lng配置中心, , 0, True, True, True, , , mstrPrivs)
'        Set RecReturn = frmSelector.showMe(Me, 1, IIf(mint编辑状态 = 8 Or mbln退货, 2, 1), strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , IIf(mint编辑状态 = 8 Or mbln退货, Val(txtProvider.Tag), 0), True, True, True, , , mstrPrivs)
    
        If Not rsReturn.EOF Then
            Me.txtDrug.Text = "(" & rsReturn!药品编码 & ")" & rsReturn!通用名
            Me.txtDrug.Tag = rsReturn!药品ID
        End If
    End If
End Sub

Private Sub txtFinditem_GotFocus()
    Call zlControl.TxtSelAll(txtFindItem)
End Sub

Private Sub txtFinditem_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long
    Dim StrDate As String
    Dim blnScaner As Boolean
    Dim blnDoIt As Boolean
    Dim intCol As Integer
    Dim blnFindItem As Boolean
    Dim strFind As String
    Dim rstemp As Recordset
    Dim i As Integer
    
    On Error GoTo errHandle
    'Me.lblMsg.Caption = ""
    If KeyAscii <> 13 Then
        If lblFindItem.Caption = "瓶签号" Then
            blnScaner = InputIsScaner(txtFindItem, KeyAscii)
        End If
    Else
        txtFindItem.Text = Trim(txtFindItem.Text)
        If txtFindItem.Text = "" Then Exit Sub
        blnScaner = InputIsScaner(txtFindItem, KeyAscii)
        blnDoIt = True
    End If
    
    If blnDoIt = True Then
        If mcondition.strTransStep = M_STR_CALSS_AUDIT _
            Or mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT _
            Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
            With vsfMedis
                If .rows = 1 Then Exit Sub
                If .TextMatrix(1, .ColIndex("医嘱ID")) = "" Then Exit Sub
            
            
                '检查输液单列表中是否有查找项目
                blnFindItem = False
                For intCol = 1 To .Cols - 1
                    If .ColKey(intCol) = lblFindItem.Caption Then
                        blnFindItem = True
                        Exit For
                    End If
                Next
                If blnFindItem = False Then Exit Sub
                
                strFind = txtFindItem.Text
            
                lngRow = .FindRow(strFind, 1, .ColIndex(lblFindItem.Caption))
                 
                If lngRow > 0 Then
                    .Row = lngRow
                    .TopRow = lngRow
                Else
                    MsgBox "没找到" & Me.lblFindItem.Caption & "为[" & strFind & "]的医嘱单 。", vbInformation, gstrSysName
                    If tbcDetail.Item(mDetailType.输液单列表).Selected Then txtFindItem.SetFocus
                End If
            
                txtFindItem.Text = ""
                If tbcDetail.Item(mDetailType.输液单列表).Selected Then txtFindItem.SetFocus
            End With
        Else
            With vsfTrans
                
                
                If .rows = 1 Then Exit Sub
                If .TextMatrix(1, .ColIndex("配药ID")) = "" Then Exit Sub
                If Me.txtFindItem.Text = "" Then Exit Sub
                
                '检查输液单列表中是否有查找项目
                blnFindItem = False
                For intCol = 1 To .Cols - 1
                    If .ColKey(intCol) = lblFindItem.Caption Then
                        blnFindItem = True
                        Exit For
                    End If
                Next
                If blnFindItem = False Then Exit Sub
                
                strFind = txtFindItem.Text
            
                lngRow = .FindRow(strFind, 1, .ColIndex(lblFindItem.Caption))
                 
                
                If lngRow > 0 Then
                    .Row = lngRow
                    .TopRow = lngRow
                    
                    
                    
                    If blnScaner = True And Me.lblFindItem.Caption = "瓶签号" And Me.txtFindItem.Text <> "" Then
                        If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = 0 Then
                            For i = 1 To .rows - 1
                                If .TextMatrix(i, .ColIndex("配药ID")) = .TextMatrix(lngRow, .ColIndex("配药ID")) Then
                                    .TextMatrix(i, .ColIndex("选择")) = -1
                                End If
                            Next
                            
                            Call UpdateExeSign(Val(.TextMatrix(lngRow, .ColIndex("配药ID"))), IIf(Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1, 1, 0))
                            
                            DoEvents
                            If InStr(1, mstrLastLabel, .TextMatrix(lngRow, .ColIndex("瓶签号"))) = 0 Then
                                mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & .TextMatrix(lngRow, .ColIndex("瓶签号"))
                            End If
                        End If
                    End If
                    
                    If mParams.blnTwoCode = True And lblFindItem.Caption = "瓶签号" And blnScaner And Me.txtFindItem.Text <> "" Then
                        
                        
                        If mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
                            txtFindItem.Text = ""
                            Call PIVAWork_Dosage(Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("配药ID"))), "扫描")
                            Exit Sub
                        ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                            txtFindItem.Text = ""
                            
                            
                            Call PIVAWork_Send(Val(vsfTrans.TextMatrix(vsfTrans.Row, vsfTrans.ColIndex("配药ID"))), "扫描")
                            Exit Sub
                        End If
                    End If
                Else

                    If mParams.blnTwoCode = True And lblFindItem.Caption = "瓶签号" And blnScaner And Me.txtFindItem.Text <> "" Then
                        
                        gstrSQL = "select 操作状态 from 输液配药记录 where 瓶签号=[1] and 执行时间 between [2] and [3]"
                        Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询瓶签号", txtFindItem.Text, CDate(mcondition.strTransStartTime), CDate(mcondition.strTransEndTime))
                        
                        Debug.Print "2"
                        
                        DoEvents
                        If rstemp.EOF Then
                            Me.lblMsg.Caption = "该瓶签不存在"
                        Else
                            If mcondition.strTransStep = M_STR_CALSS_DOSAGE Then
                                If rstemp!操作状态 = 4 Then
                                    Me.lblMsg.Caption = "该瓶签已扫描"
                                ElseIf rstemp!操作状态 = 1 Then
                                    Me.lblMsg.Caption = "该瓶签在摆药环节"
                                ElseIf rstemp!操作状态 = 2 Then
                                    Me.lblMsg.Caption = "该瓶签在配药环节"
                                ElseIf rstemp!操作状态 = 5 Then
                                    Me.lblMsg.Caption = "该瓶签在已发送环节"
                                ElseIf rstemp!操作状态 >= 9 Then
                                    Me.lblMsg.Caption = "该条医嘱已停止或销账"
                                End If
                            ElseIf mcondition.strTransStep = M_STR_CALSS_SEND Then
                                If rstemp!操作状态 = 5 Then
                                    Me.lblMsg.Caption = "该瓶签已扫描"
                                ElseIf rstemp!操作状态 = 1 Then
                                    Me.lblMsg.Caption = "该瓶签在摆药环节"
                                ElseIf rstemp!操作状态 = 2 Then
                                    Me.lblMsg.Caption = "该瓶签在配药环节"
                                ElseIf rstemp!操作状态 >= 9 Then
                                    Me.lblMsg.Caption = "该条医嘱已停止或销账"
                                End If
                            End If
                        End If
                    Else
                        MsgBox "没找到" & Me.lblFindItem.Caption & "为[" & strFind & "]的输液单 。", vbInformation, gstrSysName
                    End If
                End If
            
                txtFindItem.Text = ""
                If tbcDetail.Item(mDetailType.输液单列表).Selected Then txtFindItem.SetFocus
                If tbcDetail.Item(mDetailType.输液单卡片).Selected And lngRow > 0 Then
                    mfrmPIVCard.GetForce Val(.TextMatrix(lngRow, .ColIndex("配药ID")))
                End If
            End With
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtLog_GotFocus()
    If Me.txtLog.ForeColor = &H80000000 Then
        Me.txtLog.ForeColor = &H80000001
        txtLog.Text = ""
    End If
End Sub

Private Sub txtName_GotFocus()
    Call zlControl.TxtSelAll(txtdept)
End Sub

Private Sub txtTag_GotFocus()
    Call zlControl.TxtSelAll(txtTag)
End Sub

Private Sub txtTag_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
'            vsfTrans.ColWidth(lngCol) = vsfTrans.ColData(lngCol)
            vsfTrans.ColHidden(lngCol) = False
        Else
'            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = 0
            vsfTrans.ColHidden(lngCol) = True
        End If
    End If
    
    Call SaveListColState
End Sub

Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub

Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub



Private Sub InitColSelList(ByVal strUnvisble As String)
    Dim i As Integer
    
    With vsfColSel
        .rows = .FixedRows
        For i = 1 To vsfTrans.Cols - 1
            '不在不允许显示列表的列才能加入列选择列表
            If IsInString(strUnvisble, vsfTrans.ColKey(i), ";") = False Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = vsfTrans.ColKey(i)
                .RowData(.rows - 1) = i
'
'                '列宽为空或者隐藏的列设置为不勾选
'                If Not (vsfTrans.ColWidth(i) = 0 Or vsfTrans.ColHidden(i)) Then
'                    .TextMatrix(.rows - 1, 0) = 0
'                End If
                
                '指定的列设置为不能设置隐藏
                If IsInString(mstrUnallowSetColHide, vsfTrans.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub

Private Sub SaveListColState()
    Dim strType As String
    Dim str列设置 As String
    Dim i As Integer
    
    With vsfTrans
        For i = 0 To .Cols - 1
            If .ColHidden(i) = True Then
                str列设置 = IIf(str列设置 = "", "", str列设置 & ";") & .ColKey(i)
            End If
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\表格", mcondition.strTransStep, str列设置)
End Sub



Private Sub vsfDept_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    With vsfDept(0)
        If Row = 0 Then Exit Sub
        If Col <> .ColIndex("选择") Then Exit Sub
        If .MouseRow <> Row Or .MouseCol <> Col Then Exit Sub
    End With

    Call GetCount

'    DoEvents
'    Call RefreshDetailList(index)
End Sub


Private Sub vsfDept_EnterCell(index As Integer)
    With vsfDept(0)
        .Editable = flexEDNone
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("选择") Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub

Private Function AdviceCheckWarn(ByVal Int单据 As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long, Optional ByVal lng医嘱id As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号，lngCmd=0时需要
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String, lng药品id As Long, str单量单位 As String
    Dim strSQL As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lng主页ID As Long
    Dim str挂号单 As String
    Dim rs医嘱 As Recordset
    Dim str频率 As String
    Dim blnDo As Boolean
    Dim strTmp As String
    
    AdviceCheckWarn = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    If strNo = "" Then Exit Function

    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
    If lng医嘱id = 0 Then
        strSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
            " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
            " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
            " And A.单据=[2] And A.no=[1] "
        strTmp = Replace(strSQL, "住院费用记录", "门诊费用记录")
        strSQL = strSQL & " Union All " & strTmp
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo, Int单据)
    
        If rsTmp.RecordCount = 0 Then
            rsTmp.Close
            Exit Function
        End If
        
        lngPatiID = rsTmp!病人ID
        str挂号单 = nvl(rsTmp!挂号单)
        lng主页ID = rsTmp!主页id
    Else
        strSQL = "select A.病人id,A.相关id,A.主页id,A.收费细目id,A.开嘱医生,A.单次用量,A.频率次数,A.频率间隔,A.间隔单位,A.医嘱期效,A.开始执行时间 开始时间,A.执行终止时间 结束时间,C.名称 用法,D.名称 药品名称,D.计算单位 from 病人医嘱记录 A,病人医嘱记录 B,诊疗项目目录 C,收费细目 D where A.相关id=B.医嘱id and B.诊疗项目id=C.id and A.收费细目id=d.id and A.id=[1]"
        Set rs医嘱 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱id)
        
        lngPatiID = rs医嘱!病人ID
        lng主页ID = rs医嘱!主页id
    End If
    
    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    If lngPatiID <> mlngPassPati Then
        If str挂号单 <> "" Then               '门诊病人
            strSQL = "Select 病人ID,Count(Distinct Trunc(登记时间)) as 就诊次数 From 病人挂号记录 Where 记录性质=1 And 记录状态=1 And 病人ID=[1] Group by 病人ID"
            strSQL = "Select D.就诊次数,A.姓名,A.性别,A.出生日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,(" & strSQL & ") D,人员表 E" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=D.病人ID" & _
                " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, str挂号单)
            If rsTmp.EOF Then
                Screen.MousePointer = 0
                Exit Function
            End If

            Call PassSetPatientInfo(lngPatiID, rsTmp!就诊次数, rsTmp!姓名, zlStr.nvl(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.nvl(rsTmp!医生码) & "/" & zlStr.nvl(rsTmp!医生名), ""), "")
        Else                                    '住院病人
            strSQL = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And A.主页id=B.主页id And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngPatiID, lng主页ID)
            If rsTmp.EOF Then
                Screen.MousePointer = 0
                Exit Function
            End If

            Call PassSetPatientInfo(lngPatiID, lng主页ID, rsTmp!姓名, zlStr.nvl(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.nvl(rsTmp!医生码) & "/" & zlStr.nvl(rsTmp!医生名), ""), _
                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        End If
        mlngPassPati = lngPatiID
    End If

    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        '取药品名称
         str药品 = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("药品名称"))
         lng药品id = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("药品ID"))
         str单量单位 = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("剂量单位"))
         '取药品给药途径
         str用法 = vsfTrans.TextMatrix(lngRow, vsfTrans.ColIndex("用法"))

        If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
        If InStr(str药品, "]") > 0 Then str药品 = Mid(str药品, InStr(str药品, "]") + 1, Len(str药品) - InStr(str药品, "]"))
        '传入查询药品信息
        Call PassSetQueryDrug(lng药品id, str药品, str单量单位, str用法)

        '设置菜单可用状态
        Call SetPassMenuState

        AdviceCheckWarn = 1 '表示可以弹出菜单

        Screen.MousePointer = 0: Exit Function
    Else
        With rs医嘱
            '用药审核或用药研究
            str药品 = "": str用法 = "": str频率 = ""
            i = 1
            If !开嘱医生 <> "" Then
                strSQL = "select 编号 from 人员表 where 姓名=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "", rs医嘱!开嘱医生)
            End If
            
            blnDo = lng医嘱id <> 0 And !收费细目id <> 0
            If blnDo Then
                '取药品名称
                str药品 = !药品名称
                
                '取药品给药途径
                str用法 = !用法
                
                '取用药频率(次/天),都为整数四舍五入
                If !间隔单位 = "天" Then
                    str频率 = !频率次数 & "/" & !频率间隔
                ElseIf !间隔单位 = "周" Then
                    str频率 = !频率次数 & "/7"
                ElseIf !间隔单位 = "小时" Then
                    If Val(!频率间隔) <= 24 Then
                        str频率 = Format(24 / Val(!频率间隔) * Val(!频率次数), "0") & "/1"
                    Else
                        str频率 = Val(!频率次数) & "/" & Format(Val(!频率间隔) / 24, "0")
                    End If
                ElseIf !间隔单位 = "分钟" Then
                    str频率 = Format((24 * 60) / Val(!频率间隔) * Val(!频率次数), "0") & "/1"
                End If
                
                Call PassSetRecipeInfo(lng医嘱id, !收费细目id, str药品, _
                    !单次用量, !计算单位, str频率, _
                    Format(!开始时间, "yyyy-MM-dd"), Format(!结束时间, "yyyy-MM-dd"), str用法, _
                    !相关id, !医嘱期效, rsTmp!编号 & "\" & !开嘱医生)
            End If
            
            '无可审查的药品
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) Then
                Screen.MousePointer = 0: Exit Function
            End If
        End With
    End If

    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub SetPassMenuState()
    '功能：设置Pass菜单可用状态
    'Pass
    Dim objPopup As CommandBarControl

    ''''一级菜单
    '药物临床信息参考
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPRRes") = 1

    '药品说明书
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Directions") = 1

    '中国药典
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Chp") = 1

    '病人用药教育
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPERes") = 1

    '检验值
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CheckRes") = 1

    '专项信息
'    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 5, , True)
'    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("") = 1

    '医药信息中心
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MEDInfo") = 1

    '药品配对信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-DRUG") = 1

    '给药途径配对信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-ROUTE") = 1

    '医院药品信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("HisDrugInfo") = 1
    
    
    ''''专项信息二级菜单
    '药物-药物相互作用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDIM") = 1
    
    '药物-食物相互使用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DFIM") = 1
    
    '国内注射剂体外配伍
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MatchRes") = 1
    
    '国外注射剂体外配伍
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("TriessRes") = 1
    
    '禁忌症
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDCM") = 1
    
    '副作用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 5, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("SIDE") = 1
    
    '老年人用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("GERI") = 1
    
    '儿童用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PEDI") = 1
    
    '妊娠期用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PREG") = 1
    
    '哺乳期用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("LACT") = 1
End Sub

Private Sub vsfDept_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_SortPopup)
        If Not objPopup Is Nothing Then
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub vsfMedis_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long
    
    With Me.vsfMedis
        If Col = .ColIndex("选择") Then
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("相关id")) = .TextMatrix(Row, .ColIndex("相关id")) And _
                    .TextMatrix(lngRow, .ColIndex("摆药标志")) = .TextMatrix(Row, .ColIndex("摆药标志")) Then
                    .TextMatrix(lngRow, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                End If
            Next
        End If
    
        Call mfrmPIVCard.ChooseOneRec(.TextMatrix(Row, .ColIndex("相关id")), IIf(.TextMatrix(Row, .ColIndex("选择")) = -1, 1, 0))
    
    End With
End Sub

Private Sub vsfMedis_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Me.vsfMedis.ColIndex("审") And Col <> Me.vsfMedis.ColIndex("选择") Then Cancel = True
    
    If (mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT) Then
        If Col = Me.vsfMedis.ColIndex("选择") Then
            If Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("摆药标志"))) = 1 Then
                Cancel = True
            End If
        End If
    End If
    
    If ((Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("已审核"))) <> 0 And Val(vsfMedis.TextMatrix(Row, vsfMedis.ColIndex("已审核"))) <> 3) And mcondition.strTransStep = M_STR_CALSS_AUDIT) Then
        Cancel = True
    ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        Cancel = False
    End If
End Sub

Private Sub vsfMedis_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    With vsfMedis
        If Col = .ColIndex("审") Or _
            Col = .ColIndex("当前行") Or _
            Col = .ColIndex("选择") Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfMedis_DblClick()
    Dim lngRow As Long
    Dim str医嘱ID串 As String
    
    With vsfMedis
        If .Row = 0 Then Exit Sub
        
        '取当前一并给药的医嘱ID串
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("相关ID")) = .TextMatrix(.Row, .ColIndex("相关ID")) Then
                str医嘱ID串 = str医嘱ID串 & IIf(str医嘱ID串 = "", "", ",") & .TextMatrix(lngRow, .ColIndex("医嘱id"))
            End If
        Next
        
        If Val(.TextMatrix(.Row, .ColIndex("已审核"))) <> 0 And Val(.TextMatrix(.Row, .ColIndex("已审核"))) <> 3 And mcondition.strTransStep = M_STR_CALSS_AUDIT Then Exit Sub
        If .Col = .ColIndex("审") Then
            If .TextMatrix(.Row, .ColIndex("标志")) = "0" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("相关id")) = .TextMatrix(.Row, .ColIndex("相关id")) Then
                        .TextMatrix(lngRow, .ColIndex("标志")) = "1"
                        .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
                    End If
                Next
                
            ElseIf .TextMatrix(.Row, .ColIndex("标志")) = "1" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("相关id")) = .TextMatrix(.Row, .ColIndex("相关id")) Then
                        .TextMatrix(lngRow, .ColIndex("标志")) = "2"
                        .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Me.ImgList.ListImages(4).Picture
                        .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
                    End If
                Next
                
            ElseIf .TextMatrix(.Row, .ColIndex("标志")) = "2" Then
                For lngRow = 1 To .rows - 1
                    If .TextMatrix(lngRow, .ColIndex("相关id")) = .TextMatrix(.Row, .ColIndex("相关id")) Then
                        .TextMatrix(lngRow, .ColIndex("标志")) = "0"
                        .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Nothing
                    End If
                Next
            End If
        End If
        
        If .Col = .ColIndex("警") Then
            If IsInString(gstrprivs, "合理用药监测", ";") And Not gobjPass Is Nothing Then
                Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, .TextMatrix(.Row, .ColIndex("住院号")), "2", Val(.TextMatrix(.Row, .ColIndex("病人ID"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), "", str医嘱ID串)
            End If
        End If
        
        If .Col = .ColIndex("药品名称") Then
            If IsInString(gstrprivs, "合理用药监测", ";") And Not gobjPass Is Nothing Then
                Call gobjPass.zlPassAdviceMainPoint_YF("2", .TextMatrix(.Row, .ColIndex("药品id")), .TextMatrix(.Row, .ColIndex("药名")))
            End If
        End If
        
        If .TextMatrix(.Row, .ColIndex("相关ID")) <> "" Then
            Call mfrmPIVCard.ChooseOneRec(.TextMatrix(.Row, .ColIndex("相关ID")), .TextMatrix(.Row, .ColIndex("标志")))
        End If
    End With
    
     
End Sub

Private Sub vsfMedis_EnterCell()
    Dim intRow As Integer
    Dim intBegin As Integer
    Dim intEnd As Integer
    Dim strDiag As String
    Dim i As Integer
    
    With Me.vsfMedis
        txtLog.Text = ""
        txtDia.Text = ""
        Me.txtLog.Text = .TextMatrix(.Row, .ColIndex("药师审核原因"))
        If Val(.TextMatrix(.Row, .ColIndex("相关ID"))) = 0 Then
            txtLog.Enabled = False
            Me.CmdSave.Enabled = False
            Exit Sub
        End If
        If Me.tabDeptList.Selected.index = 0 Then
            If Me.txtLog.Text = "" Then
                txtLog.Text = "请输入审核原因"
                txtLog.ForeColor = &H80000000
            Else
                txtLog.ForeColor = &H80000001
            End If
            txtLog.Enabled = True
            CmdSave.Enabled = True
        Else
            txtLog.ForeColor = &H80000001
        End If
        If mintBeginRow <> 0 And mintBeginRow <= .rows - 1 Then
            If Val(.TextMatrix(mintBeginRow, .ColIndex("背景号"))) = 1 Then
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HC0FFC0
            End If
        End If
        
        For intRow = IIf(.Row > 8, .Row - 8, 1) To IIf(.Row + 8 > .rows - 1, .rows - 1, .Row + 8)
            If Val(.TextMatrix(.Row, .ColIndex("相关ID"))) = Val(.TextMatrix(intRow, .ColIndex("相关ID"))) And .Row > intRow And intBegin = 0 Then
                intBegin = intRow
            ElseIf .Row < intRow And Val(.TextMatrix(.Row, .ColIndex("相关ID"))) = Val(.TextMatrix(intRow, .ColIndex("相关ID"))) Then
                intEnd = intRow
            End If
        Next
    
        intRow = 0
        mintBeginRow = IIf(intBegin = 0, .Row, intBegin)
        mintEndRow = IIf(intEnd = 0, .Row, intEnd)
'        .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HFFE8D0
        
        .Redraw = flexRDNone

        '初始化框框
        For intRow = 0 To .rows - 1
            .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 0, 0, 0, 0, 0, 0
        Next
        
        intBegin = 0
        intEnd = 0
        '查找选中列的框框范围
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("相关ID")) = .TextMatrix(.Row, .ColIndex("相关ID")) Then
                If intBegin = 0 Then intBegin = intRow
                intEnd = intRow
            End If
        Next
        
        '对当前选中的行赋个框框
        If intBegin = intEnd Then
            '①只有一行
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 2, 0, 0
        ElseIf intBegin + 1 = intEnd Then
            '②只有2行
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '上部分
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '下部分
        Else
            '③3行及以上
            For intRow = intBegin + 1 To intEnd - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 2, 0, 2, 0, 0, 0     '中间部分
            Next
            
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '上部分
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '下部分
        End If
        
        '去除选中时的背景色
        .BackColorSel = .Cell(flexcpBackColor, .Row, 1)
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        
        .Redraw = flexRDDirect
        
        intRow = 0
        
        strDiag = RecipeSendWork_GetDiagnosis(2, Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))))
        
        If InStr(1, strDiag, "西医入院诊断") >= 1 Then
            strDiag = Mid(strDiag, InStr(1, strDiag, "西医入院诊断") + 7)
            
            If InStr(1, strDiag, "|") >= 1 Then
                strDiag = Mid(strDiag, 1, InStr(1, strDiag, "|") - 1)
            End If
            
            txtDia.Text = ""
            If strDiag <> "" Then
                strDiag = strDiag & ";"
                For i = 0 To UBound(Split(strDiag, ";"))
                    If Split(strDiag, ";")(i) <> "" Then
                        If InStr(1, txtDia.Text & "※", "※" & Split(strDiag, ";")(i) & "※") < 1 Then
                            txtDia.Text = IIf(txtDia.Text = "", " ※", txtDia.Text & " ※") & Split(strDiag, ";")(i)
                        End If
                    End If
                Next
            End If
        End If
        
        
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, .ColIndex("药品id")), .TextMatrix(.Row, .ColIndex("药名")))
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
    End With
End Sub

Private Sub vsfMedis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim lngPatiID As Long
    Dim lng主页ID As Long
    Dim str审查结果 As String
    Dim lng医嘱id As Long
    
    If Button = 2 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        If IsInString(gstrprivs, "合理用药监测", ";") And vsfMedis.MouseCol = vsfMedis.ColIndex("警") And mParams.intShowPass <> 2 And Not gobjPass Is Nothing Then
            '检查Pass状态
            lng医嘱id = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("医嘱id")))
            str审查结果 = vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("警告"))
            lngPatiID = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("病人id")))
            lng主页ID = Val(vsfMedis.TextMatrix(vsfMedis.MouseRow, vsfMedis.ColIndex("主页id")))

            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            objPopup.Visible = True
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng主页ID, "", str审查结果, lng医嘱id)
            
            objPopup.CommandBar.ShowPopup
        Else
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_OperPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.CommandBar.Controls
                    
                    If cbrControl.Id = conMenu_Oper_Look Then
                        cbrControl.Visible = True
                    Else
                        cbrControl.Visible = False
                    End If
                Next
                    
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsfMedis_RowColChange()
    With vsfMedis
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
    End With
End Sub

Private Sub vsfMsg_DblClick()
    Dim i As Integer
    
    With Me.vsfMsg
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("时间")) = "" Then Exit Sub
        
        If DateDiff("D", mdateToday, Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss")) > 2 Then
            Me.cbo时间范围.ListIndex = 3
            
            If Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss") > Me.Dtp结束时间.Value Then
                Me.Dtp结束时间.Value = Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss")
            ElseIf Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss") < Me.Dtp开始时间.Value Then
                Me.Dtp开始时间.Value = Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss")
            End If
        Else
            Me.cbo时间范围.ListIndex = DateDiff("d", mdateToday, Format(.TextMatrix(.Row, .ColIndex("执行时间")), "yyyy-mm-dd hh:mm:ss"))
        End If
        
        If (Val(Me.cbo时间范围.Tag) = 3 And Me.cbo时间范围.ListIndex < 3) Or (Val(Me.cbo时间范围.Tag) < 3 And Me.cbo时间范围.ListIndex = 3) Then
            Call ResizeConditionArea
        End If
        
        Me.cbo时间范围.Tag = Me.cbo时间范围.ListIndex
        If .TextMatrix(.Row, .ColIndex("类型")) = "医嘱作废" Then
            Me.tabDeptList.Item(1).Selected = True
            Me.tbcLook.Item(5).Selected = True
        ElseIf .TextMatrix(.Row, .ColIndex("类型")) = "销帐申请" Then
            Me.tabDeptList.Item(0).Selected = True
            Me.tabWork.Item(4).Selected = True
        ElseIf .TextMatrix(.Row, .ColIndex("类型")) = "批次调整" Then
            Me.tabDeptList.Item(0).Selected = True
            Me.tabWork.Item(1).Selected = True
        End If
        
        Call RefreshDeptList(Me.tabDeptList.Selected.index)
        
        For i = 1 To Me.vsfDept(0).rows - 1
            If vsfDept(0).TextMatrix(i, vsfDept(0).ColIndex("病区id")) = .TextMatrix(.Row, .ColIndex("病人病区id")) Then
                vsfDept(0).TextMatrix(i, vsfDept(0).ColIndex("选择")) = -1
                Call vsfDept_AfterEdit(0, i, .ColIndex("选择"))
            End If
        Next
        
        Call RefreshDetailList(Me.tabDeptList.Selected.index)

        .RemoveItem (.Row)
        lblMsgComment.Caption = "消息提醒(" & vsfMsg.rows - 1 & ")"
        
    End With
    
End Sub

Private Sub vsfSumDrug_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfSumDrug
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsfSumDrug_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfSumDrug
        If Col = .ColIndex("打包") Then
            .Col = .ColIndex("是否打包")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfSumDrug_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    With vsfSumDrug
        If Col = .ColIndex("打包") Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfTrans_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    Dim blnNext As Boolean
    Dim strLabel As String
    Dim int打包 As Integer
    Dim i As Integer
    
    With vsfTrans
        If Row = 0 Then Exit Sub
        
        If Col = .ColIndex("选择") Then
            Call UpdateExeSign(Val(.TextMatrix(Row, .ColIndex("配药ID"))), IIf(.TextMatrix(Row, .ColIndex("选择")) = -1, 1, 0))
            
            DoEvents
            
            strLabel = .TextMatrix(Row, .ColIndex("瓶签号"))
            
            If Val(.TextMatrix(Row, .ColIndex("选择"))) = -1 Then
                If InStr(1, mstrLastLabel, strLabel) = 0 Then
                    mstrLastLabel = IIf(mstrLastLabel = "", "", mstrLastLabel & ",") & strLabel
                End If
            Else
                mstrLastLabel = Replace(mstrLastLabel, strLabel & ",", "")
                mstrLastLabel = Replace(mstrLastLabel, strLabel, "")
            End If
            
            Call mfrmPIVCard.ChooseOneRec(.TextMatrix(Row, .ColIndex("配药ID")), IIf(.TextMatrix(Row, .ColIndex("选择")) = -1, 1, 0))
        End If
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            If Col = .ColIndex("配药批次") Then
                .ColComboList(.ColIndex("配药批次")) = ""
                If mrsTrans Is Nothing Then Exit Sub
                
                If .TextMatrix(Row, .ColIndex("配药批次")) = .TextMatrix(Row, .ColIndex("原批次")) Then
                    mfrmPIVCard.Changebatch Val(.TextMatrix(Row, .ColIndex("配药ID"))), .TextMatrix(Row, .ColIndex("配药批次"))
                    Exit Sub
                End If
                
                If MsgBox("是否确认把批次由[" & .TextMatrix(Row, .ColIndex("原批次")) & "]调整为[" & .TextMatrix(Row, .ColIndex("配药批次")) & "]？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    .TextMatrix(Row, .ColIndex("配药批次")) = .TextMatrix(Row, .ColIndex("原批次"))
                    
                    For i = Row - 1 To 0 Step -1
                        If .TextMatrix(Row, .ColIndex("配药ID")) = .TextMatrix(i, .ColIndex("配药ID")) Then
                            .TextMatrix(i, .ColIndex("配药批次")) = .TextMatrix(Row, .ColIndex("原批次"))
                        End If
                    Next
                    
                    For i = Row + 1 To .rows - 1
                        If .TextMatrix(Row, .ColIndex("配药ID")) = .TextMatrix(i, .ColIndex("配药ID")) Then
                            .TextMatrix(i, .ColIndex("配药批次")) = .TextMatrix(Row, .ColIndex("原批次"))
                        End If
                    Next
                    Exit Sub
                End If
                
                .TextMatrix(Row, .ColIndex("原批次")) = .TextMatrix(Row, .ColIndex("配药批次"))
                mfrmPIVCard.Changebatch Val(.TextMatrix(Row, .ColIndex("配药ID"))), .TextMatrix(Row, .ColIndex("配药批次"))
                
                
                int打包 = Mid(mstr打包, InStr(mstr打包, "," & .TextMatrix(Row, .ColIndex("配药批次")) & ",") + Len("," & .TextMatrix(Row, .ColIndex("配药批次")) & ","), 1)
                If int打包 = 1 Then int打包 = 2
                .TextMatrix(Row, .ColIndex("是否打包")) = int打包
                .Cell(flexcpPicture, Row, .ColIndex("打包"), Row, .ColIndex("打包")) = IIf(int打包 = 2, picPacker(2).Picture, Nothing)
                .Cell(flexcpPictureAlignment, Row, .ColIndex("打包"), Row, .ColIndex("打包")) = flexPicAlignCenterCenter
                .Cell(flexcpForeColor, Row, .ColIndex("配药批次")) = vbBlue
                
                mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(Row, .ColIndex("配药ID")))
                Do While Not mrsTrans.EOF
                    mrsTrans!配药批次 = .TextMatrix(Row, .ColIndex("配药批次"))
                    mrsTrans!是否打包 = IIf(int打包 <> 0, 2, 0)
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                
                DoEvents
                
                strInput = .TextMatrix(Row, .ColIndex("配药ID")) & "," & Left(.TextMatrix(Row, .ColIndex("配药批次")), 1) & ":"
                
                On Error GoTo errHandle
                
                If strInput <> "" Then
                    gstrSQL = "Zl_输液配药记录_分批("
                    '配药ID,批次
                    gstrSQL = gstrSQL & "'" & strInput & "'"
                    '是否调整批次
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "设置批次")
                    
                    gstrSQL = "Zl_输液配药记录_打包("
                    '配药ID,打包
                    gstrSQL = gstrSQL & "'" & .TextMatrix(Row, .ColIndex("配药ID")) & "," & int打包 & "'"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打包设置")
                End If
            End If
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfTrans_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfTrans
        If Col = .ColIndex("打包") Then
            .Col = .ColIndex("是否打包")
            .Sort = Order
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfTrans
        If Row = 0 Then Exit Sub
        If Val(.TextMatrix(Row, .ColIndex("配药ID"))) = 0 Then Exit Sub
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            If Col = .ColIndex("配药批次") Then
                If mParams.bln批次设置 = False Then Exit Sub
                .ColComboList(.ColIndex("配药批次")) = mParams.strBatchList
            End If
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '设置不能移动的列
    With vsfTrans
        If Col = .ColIndex("审") Then
            Position = .ColIndex("审")
        End If
        
        If Col = .ColIndex("选择") Then
            Position = .ColIndex("选择")
        End If
        
        If Col = .ColIndex("打印") Then
            Position = .ColIndex("打印")
        End If
        
        If Col = .ColIndex("打包") Then
            Position = .ColIndex("打包")
        End If
        
        If Col = .ColIndex("配药批次") Then
            Position = .ColIndex("配药批次")
        End If
        
        If (Col <> .ColIndex("审") And Position = .ColIndex("审")) Or _
            (Col <> .ColIndex("选择") And Position = .ColIndex("选择")) Or _
            (Col <> .ColIndex("打印") And Position = .ColIndex("打印")) Or _
            (Col <> .ColIndex("打包") And Position = .ColIndex("打包")) Or _
            (Col <> .ColIndex("配药批次") And Position = .ColIndex("配药批次")) Then
            Position = Col
        End If
    End With
End Sub

Private Sub vsfTrans_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    With vsfTrans
        If Col = .ColIndex("审") Or _
            Col = .ColIndex("当前行") Or _
            Col = .ColIndex("选择") Or Col = .ColIndex("打包") Or _
            Col = .ColIndex("配药批次") Or Col = .ColIndex("打印") Then
            Cancel = True
        End If
    End With
End Sub
Private Sub vsfTrans_DblClick()
    Dim strInput As String
    Dim intFirst As Integer
    Dim lngRow As Long
    Dim str医嘱ID串 As String
    
    With vsfTrans
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("配药ID")) = "" Then Exit Sub
        If mrsTrans Is Nothing Then Exit Sub
        
        '取当前一并给药的医嘱ID串
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                str医嘱ID串 = str医嘱ID串 & IIf(str医嘱ID串 = "", "", ",") & .TextMatrix(lngRow, .ColIndex("对应医嘱ID"))
            End If
        Next
        
        Select Case .Col
            Case .ColIndex("打包")
                If mcondition.strTransStep <> M_STR_CALSS_PREPARE And mcondition.strTransStep <> M_STR_CALSS_DOSAGE Then Exit Sub
                If mParams.bln打包设置 = False Then Exit Sub
                
                If MsgBox("是否调整为" & IIf(Val(.TextMatrix(.Row, .ColIndex("是否打包"))) = 0, """打包""", """不打包""") & "状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
                If Val(.TextMatrix(.Row, .ColIndex("是否打包"))) = 0 Then
                    .TextMatrix(.Row, .ColIndex("是否打包")) = 2
                Else
                    .TextMatrix(.Row, .ColIndex("是否打包")) = 0
                End If
                
                '更新配液(打包)图标
                .Col = .ColIndex("打包")
                .CellPicture = IIf(.TextMatrix(.Row, .ColIndex("是否打包")) = 2, picPacker(2).Picture, Nothing)
                .CellPictureAlignment = flexPicAlignCenterCenter
                
                mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(.Row, .ColIndex("配药ID")))
                Do While Not mrsTrans.EOF
                    intFirst = intFirst + 1
                    mrsTrans!是否打包 = Val(.TextMatrix(.Row, .ColIndex("是否打包")))
                    
                    If mcondition.strTransStep = M_STR_CALSS_DOSAGE And intFirst = 1 And .TextMatrix(.Row, .ColIndex("是否打包")) > 0 Then
                        mintCountPack = mintCountPack + IIf(IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")), 0, 1)
                    Else
                        If IIf(IsNull(mrsTrans!摆药时间), "", Format(mrsTrans!摆药时间, "YYYY-MM-DD HH:MM:SS")) <= IIf(IsNull(mrsTrans!打包时间), "", Format(mrsTrans!打包时间, "YYYY-MM-DD HH:MM:SS")) Then
                            mintCountPack = mintCountPack - 1
                        End If
                    End If
                    
                    mrsTrans!打包时间 = IIf(.TextMatrix(.Row, .ColIndex("是否打包")) = 0, "", Sys.Currentdate)
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                
                Call GetCount
                
                mfrmPIVCard.PackCard Val(.TextMatrix(.Row, .ColIndex("配药ID"))), .TextMatrix(.Row, .ColIndex("是否打包"))
                
                DoEvents
                
                strInput = .TextMatrix(.Row, .ColIndex("配药ID")) & "," & .TextMatrix(.Row, .ColIndex("是否打包"))
                
                On Error GoTo errHandle
                
                If strInput <> "" Then
                    gstrSQL = "Zl_输液配药记录_打包("
                    '配药ID,打包
                    gstrSQL = gstrSQL & "'" & strInput & "'"
                    '手工调整打包
                    gstrSQL = gstrSQL & ",1"
                    gstrSQL = gstrSQL & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, "打包设置")
                End If
            Case .ColIndex("锁")
                .TextMatrix(.Row, .ColIndex("是否锁定")) = IIf(.TextMatrix(.Row, .ColIndex("是否锁定")) = "1", 0, 1)
                .Cell(flexcpPicture, .Row, .ColIndex("锁"), .Row, .ColIndex("锁")) = IIf(.TextMatrix(.Row, .ColIndex("是否锁定")) = "1", Me.ImgList.ListImages(5).Picture, Me.ImgList.ListImages(6).Picture)
                .Cell(flexcpPictureAlignment, .Row, .ColIndex("锁"), .Row, .ColIndex("锁")) = flexPicAlignCenterCenter
                    
                 Call SetLock(.TextMatrix(.Row, .ColIndex("是否锁定")), .TextMatrix(.Row, .ColIndex("配药id")), True)
            Case .ColIndex("审查结果")
                If IsInString(gstrprivs, "合理用药监测", ";") And Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, .TextMatrix(.Row, .ColIndex("住院号")), "2", Val(.TextMatrix(.Row, .ColIndex("病人ID"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))), "", str医嘱ID串)
                End If
            Case .ColIndex("药品名称")
                If IsInString(gstrprivs, "合理用药监测", ";") And Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassAdviceMainPoint_YF("2", .TextMatrix(.Row, .ColIndex("药品id")), Mid(.TextMatrix(.Row, .ColIndex("药品名称")), InStr(.TextMatrix(.Row, .ColIndex("药品名称")), "]") + 1))
                End If
            Case .ColIndex("审")
                If .TextMatrix(.Row, .ColIndex("标志")) = "0" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                            .TextMatrix(lngRow, .ColIndex("标志")) = "1"
                            .TextMatrix(lngRow, .ColIndex("选择")) = 1
                            .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Me.ImgList.ListImages(3).Picture
                            .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
                        End If
                    Next
                ElseIf .TextMatrix(.Row, .ColIndex("标志")) = "1" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                            If mPrives.bln销帐拒绝 Then
                                .TextMatrix(lngRow, .ColIndex("标志")) = "2"
                                .TextMatrix(lngRow, .ColIndex("选择")) = 1
                                .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Me.ImgList.ListImages(4).Picture
                                .Cell(flexcpPictureAlignment, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = flexPicAlignCenterCenter
                            Else
                                .TextMatrix(lngRow, .ColIndex("标志")) = "0"
                                .TextMatrix(lngRow, .ColIndex("选择")) = 0
                                .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Nothing
                            End If
                            Call UpdateExeSign(lngRow, .TextMatrix(lngRow, .ColIndex("标志")))
                        End If
                    Next
                ElseIf .TextMatrix(.Row, .ColIndex("标志")) = "2" Then
                    For lngRow = 1 To .rows - 1
                        If .TextMatrix(lngRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                            .TextMatrix(lngRow, .ColIndex("标志")) = "0"
                            .TextMatrix(lngRow, .ColIndex("选择")) = 0
                            .Cell(flexcpPicture, lngRow, .ColIndex("审"), lngRow, .ColIndex("审")) = Nothing
                            
                            Call UpdateExeSign(lngRow, .TextMatrix(lngRow, .ColIndex("标志")))
                        End If
                    Next
                End If
                
                '更新数据集执行标志
                Call UpdateExeSign(.TextMatrix(.Row, .ColIndex("配药ID")), .TextMatrix(.Row, .ColIndex("标志")))
                
                If .TextMatrix(.Row, .ColIndex("配药ID")) <> "" Then
                    Call mfrmPIVCard.ChooseOneRec(.TextMatrix(.Row, .ColIndex("配药ID")), .TextMatrix(.Row, .ColIndex("标志")))
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfTrans_EnterCell()
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim rstemp As Recordset
    Dim intRow As Integer
    Dim lng配药id As Long
    Dim intBegin As Integer
    Dim intEnd As Integer
    
    With vsfTrans
        If .Row = 0 Then Exit Sub
        If .Redraw = flexRDNone Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("配药ID"))) = 0 Then Exit Sub
        
        If mintBeginRow <> 0 And mintBeginRow <= .rows - 1 Then
            If mintEndRow > .rows - 1 Then mintEndRow = .rows - 1
            If Val(.TextMatrix(mintBeginRow, .ColIndex("背景号"))) = 1 Then
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &H80000005
            Else
                .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HC0FFC0
            End If
        End If
        
        For intRow = IIf(.Row > 8, .Row - 8, 1) To IIf(.Row + 8 > .rows - 1, .rows - 1, .Row + 8)
            If Val(.TextMatrix(.Row, .ColIndex("配药ID"))) = Val(.TextMatrix(intRow, .ColIndex("配药ID"))) And .Row > intRow And intBegin = 0 Then
                intBegin = intRow
            ElseIf .Row < intRow And Val(.TextMatrix(.Row, .ColIndex("配药ID"))) = Val(.TextMatrix(intRow, .ColIndex("配药ID"))) Then
                intEnd = intRow
            End If
        Next
        
        intRow = 0
        mintBeginRow = IIf(intBegin = 0, .Row, intBegin)
        mintEndRow = IIf(intEnd = 0, .Row, intEnd)
'        .Cell(flexcpBackColor, mintBeginRow, 1, mintEndRow, .Cols - 1) = &HFFE8D0
        
        .Redraw = flexRDNone

        '初始化框框
        For intRow = 0 To .rows - 1
            .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 0, 0, 0, 0, 0, 0
        Next
        
        intBegin = 0
        intEnd = 0
        '查找选中列的框框范围
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, .ColIndex("配药ID")) = .TextMatrix(.Row, .ColIndex("配药ID")) Then
                If intBegin = 0 Then intBegin = intRow
                intEnd = intRow
            End If
        Next
        
        '对当前选中的行赋个框框
        If intBegin = intEnd Then
            '①只有一行
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 2, 0, 0
        ElseIf intBegin + 1 = intEnd Then
            '②只有2行
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '上部分
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '下部分
        Else
            '③3行及以上
            For intRow = intBegin + 1 To intEnd - 1
                .CellBorderRange intRow, 0, intRow, .Cols - 1, vbBlue, 2, 0, 2, 0, 0, 0     '中间部分
            Next
            
            .CellBorderRange intBegin, 0, intBegin, .Cols - 1, vbBlue, 2, 2, 2, 0, 0, 0     '上部分
            .CellBorderRange intEnd, 0, intEnd, .Cols - 1, vbBlue, 2, 0, 2, 2, 0, 0         '下部分
        End If
        
        '去除选中时的背景色
        .BackColorSel = .Cell(flexcpBackColor, .Row, 1)
        
        .Redraw = flexRDDirect
        .Editable = flexEDNone
        
        intRow = 0
        
        Select Case .Col
            Case .ColIndex("选择")
                .Editable = flexEDKbdMouse
            Case .ColIndex("配药批次")
                If mcondition.strTransStep = M_STR_CALSS_PREPARE And mParams.bln批次设置 = True Then
                    .Editable = flexEDKbdMouse
                End If
        End Select
        
        If mcondition.strTransStep = M_STR_CALSS_PREPARE Then
            '获取数据
            Set rstemp = PIVA_已摆药输液单(mcondition.lngCenterID, CDate(.TextMatrix(.Row, .ColIndex("执行时间"))), _
                Val(.TextMatrix(.Row, .ColIndex("病人id"))), Val(.TextMatrix(.Row, .ColIndex("主页id"))))
            
            If chkSure(1).Value = 1 Then
                rstemp.Filter = "操作状态<>1"
            End If
            
            rstemp.Sort = "配药id"
            With Me.VSFLook
                .rows = 1
                .rows = rstemp.RecordCount + 1
                .RowHeight(0) = 250
                .MergeCells = flexMergeFree
                Do While Not rstemp.EOF
                    intRow = intRow + 1
                    
                    
                    If lng配药id <> rstemp!配药id Then
                        lng配药id = rstemp!配药id
                        If lng配药id <> 0 Then
                            .rows = .rows + 1
                            .Cell(flexcpText, intRow, 0, intRow, .Cols - 1) = 0
                            .RowHidden(intRow) = True
                            intRow = intRow + 1
                        End If
                    Else
                        .MergeCol(.ColIndex("批次")) = True
                        .MergeCol(.ColIndex("摆药单号")) = True
                        .MergeCol(.ColIndex("摆药人")) = True
                        .MergeCol(.ColIndex("摆药时间")) = True
                        .MergeCol(.ColIndex("瓶签号")) = True
                        .MergeCol(.ColIndex("医嘱发送时间")) = True
                        .MergeCol(.ColIndex("执行时间")) = True
                        .MergeCol(.ColIndex("打包")) = True
                        .MergeCol(.ColIndex("操作状态")) = True
                        .MergeCol(.ColIndex("NO")) = True
                    End If
                    
                    .RowHeight(intRow) = 250
                    .TextMatrix(intRow, .ColIndex("批次")) = IIf(zlStr.nvl(rstemp!配药批次) = "", "", zlStr.nvl(rstemp!配药批次) & "#")
                    .TextMatrix(intRow, .ColIndex("药品名称")) = rstemp!通用名
                    .TextMatrix(intRow, .ColIndex("规格")) = rstemp!规格
                    .TextMatrix(intRow, .ColIndex("单量")) = FormatEx(rstemp!单量, 2) & rstemp!剂量单位
                    .TextMatrix(intRow, .ColIndex("数量")) = FormatEx(rstemp!数量, 2) & rstemp!单位
                    .TextMatrix(intRow, .ColIndex("执行时间")) = rstemp!执行时间
                    .TextMatrix(intRow, .ColIndex("瓶签号")) = rstemp!瓶签号
                    .TextMatrix(intRow, .ColIndex("配药id")) = rstemp!配药id
                    .TextMatrix(intRow, .ColIndex("摆药人")) = rstemp!操作人员
                    .TextMatrix(intRow, .ColIndex("摆药时间")) = rstemp!操作时间
                    .TextMatrix(intRow, .ColIndex("摆药单号")) = zlStr.nvl(rstemp!摆药单号, " ")
                    .TextMatrix(intRow, .ColIndex("医嘱发送时间")) = rstemp!医嘱发送时间
                    .TextMatrix(intRow, .ColIndex("操作状态")) = IIf(rstemp!操作状态 = 1, "已确认", IIf(rstemp!操作状态 = 2, "已摆药", IIf(rstemp!操作状态 = 4, "已配药", "已发送")))
                    .TextMatrix(intRow, .ColIndex("打包")) = " "
                    .TextMatrix(intRow, .ColIndex("NO")) = rstemp!NO
                    
'                    .CellPicture = picPacker(Val(.TextMatrix(intRow, .ColIndex("是否打包")))).Picture
                    .Cell(flexcpPicture, intRow, .ColIndex("打包"), intRow, .ColIndex("打包")) = IIf(Val(rstemp!是否打包) = 0, Nothing, picPacker(Val(rstemp!是否打包)).Picture)
                    .Cell(flexcpPictureAlignment, intRow, .ColIndex("打包"), intRow, .ColIndex("打包")) = flexPicAlignCenterCenter
    
                    rstemp.MoveNext
                Loop
            End With
        End If
    End With
    
End Sub

Private Sub vsfTrans_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim Int单据 As Integer
    Dim strNo As String
    
    If Button = 2 Then
        If Me.cbsMain Is Nothing Then Exit Sub
        
        If mParams.intShowPass = 1 And IsInString(gstrprivs, "合理用药监测", ";") And vsfTrans.MouseCol = vsfTrans.ColIndex("审查结果") Then
            'PASS系统弹出菜单
'            If vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("NO")) = "" Then Exit Sub
'            Int单据 = Val(vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("单据")))
'            strNo = vsfTrans.TextMatrix(vsfTrans.MouseRow, vsfTrans.ColIndex("NO"))
'
'            '检查Pass状态
'            If AdviceCheckWarn(Int单据, strNo, 0, vsfTrans.MouseRow) >= 0 Then
'                Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
'                If Not objPopup Is Nothing Then
'                    objPopup.CommandBar.ShowPopup
'                End If
'
'                Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_PASS_Item + 10, , True)
'                If Not objPopup Is Nothing Then objPopup.Visible = False
'            End If
        Else
            '右键操作菜单
            With vsfTrans
                If .Row = 0 Or .Col > .ColIndex("住院号") Then Exit Sub
                If Val(.TextMatrix(.Row, .ColIndex("配药ID"))) = 0 Then Exit Sub
            End With
            
            Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlPopup, conMenu_OperPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.CommandBar.Controls
                    cbrControl.Visible = True
                    If mcondition.strTransStep <> M_STR_CALSS_PREPARE Then
                        If cbrControl.Id = conMenu_Oper_DelBatch Then
                            cbrControl.Visible = False
                        End If
                    End If
                    
                    If cbrControl.Id = conMenu_Oper_PrintLabel Then
                        cbrControl.Visible = mParams.bln瓶签手工打印
                    ElseIf cbrControl.Id = conMenu_Oper_Bag Then
                        cbrControl.Visible = mParams.bln打包设置
                    End If
                    
                    If cbrControl.Id = conMenu_Oper_Look Then
                        cbrControl.Visible = False
                    End If
                Next
                    
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vsfTrans_RowColChange()
    '移动第一栏的标记到当前行！
    With vsfTrans
        .Cell(flexcpText, 0, 0, .rows - 1, 0) = ""
        If .Row > 0 Then
            .Cell(flexcpFontName, , 0) = "Marlett"
            .TextMatrix(.Row, 0) = 4
        End If
        
    End With
    
    
End Sub

Public Sub ChooseType(ByVal intIndex As Integer, ByVal intValue As Integer)
    Me.chkType(intIndex).Value = intValue
    chkType_Click (intIndex)
End Sub

Public Sub chkAllClick(ByVal intValue As Integer)
    mint标志 = 1
    Me.chkAll.Value = intValue
    Chk_all
End Sub

Public Sub Get配药id(ByVal str配药id As String)
    mstr配药id = str配药id
End Sub

Public Sub CheckOne(ByVal str配药id As String, ByVal intValue As Integer)
    Dim i As Integer
    
    If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        With vsfMedis
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("相关id")) = str配药id Then
                    .TextMatrix(i, .ColIndex("标志")) = intValue
                    If intValue = 1 Then
                        .Cell(flexcpPicture, i, .ColIndex("审"), i, .ColIndex("审")) = Me.ImgList.ListImages(3).Picture
                        .Cell(flexcpPictureAlignment, i, .ColIndex("审"), i, .ColIndex("审")) = flexPicAlignCenterCenter
                    ElseIf intValue = 2 Then
                        .Cell(flexcpPicture, i, .ColIndex("审"), i, .ColIndex("审")) = Me.ImgList.ListImages(4).Picture
                        .Cell(flexcpPictureAlignment, i, .ColIndex("审"), i, .ColIndex("审")) = flexPicAlignCenterCenter
                    Else
                        .Cell(flexcpPicture, i, .ColIndex("审"), i, .ColIndex("审")) = Nothing
                    End If
                End If
            Next
        End With
    ElseIf mcondition.strTransStep > M_STR_CALSS_AUDIT And mcondition.strTransStep < M_STR_CALSS_PREPARE Then
        With vsfMedis
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("相关id")) = str配药id Then
                    .TextMatrix(i, .ColIndex("选择")) = intValue
                End If
            Next
        End With
    Else
        With Me.vsfTrans
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("配药id")) = str配药id Then
                    .TextMatrix(i, .ColIndex("选择")) = IIf(intValue <> 0, -1, 0)
                End If
            Next
        End With
        UpdateExeSign -1, Me.tabDeptList.Selected.index
    End If
    
    
    
    
End Sub

Public Sub PackMain(ByVal str配药id As String, ByVal intValue As Integer)
    Dim i As Integer
    
    With Me.vsfTrans
        For i = 1 To .rows - 1
            If .TextMatrix(i, .ColIndex("配药ID")) = str配药id Then
                If intValue = 1 Then
                    .TextMatrix(i, .ColIndex("是否打包")) = 2
                Else
                    .TextMatrix(i, .ColIndex("是否打包")) = 0
                End If
                
                '更新配液(打包)图标
                .Cell(flexcpPicture, i, .ColIndex("打包"), i, .ColIndex("打包")) = IIf(intValue = 2, picPacker(2).Picture, Nothing)
                .Cell(flexcpPictureAlignment, i, .ColIndex("打包"), i, .ColIndex("打包")) = flexPicAlignCenterCenter
                Exit For
            End If
        Next
    End With
End Sub

Public Sub ChangeBatchMain(ByVal lng配药id As Long, ByVal str批次 As String)
    Dim i As Integer
    
    With Me.vsfTrans
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, .ColIndex("配药ID"))) = lng配药id Then
                .TextMatrix(i, .ColIndex("配药批次")) = str批次
                mrsTrans.Filter = "配药ID=" & Val(.TextMatrix(i, .ColIndex("配药ID")))
                Do While Not mrsTrans.EOF
                    mrsTrans!配药批次 = .TextMatrix(i, .ColIndex("配药批次"))
                    mrsTrans.Update
                    mrsTrans.MoveNext
                Loop
                Exit For
            End If
        Next
    End With
End Sub

Public Sub SetTxtFind(ByVal strText As String, ByVal IntKeyAscii As Integer)
    Me.txtFindItem.Text = strText
    txtFinditem_KeyPress IntKeyAscii
End Sub
Private Function InitMedi(ByVal intType As Integer, ByVal strIDS As String, ByVal str配药类型 As String) As Recordset
    Dim rstemp As Recordset
    Dim lng医嘱号 As Long
    Dim lng医嘱id As Long
    Dim bln操作 As Boolean
    Dim i As Integer
    Dim arrExecute As Variant
    Dim rsMedi As Recordset
    
    '输液单记录集
    Set mrsMedi = New ADODB.Recordset
    With mrsMedi
        If .State = 1 Then .Close
        
        '该记录对应的输液配药记录信息
        .Fields.Append "组号", adDouble, 18, adFldIsNullable
        .Fields.Append "id", adDouble, 18, adFldIsNullable
        .Fields.Append "相关id", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "性别", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "年龄", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "病人id", adDouble, 18, adFldIsNullable
        .Fields.Append "主页id", adDouble, 18, adFldIsNullable
        .Fields.Append "科室", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "开单医生", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "病人病区id", adLongVarChar, 18, adFldIsNullable
        .Fields.Append "病人科室id", adLongVarChar, 18, adFldIsNullable
        
        '输液配药记录对应的药品信息
        .Fields.Append "药品名称", adLongVarChar, 50, adFldIsNullable   '编码+通用名/商品名
        .Fields.Append "规格", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单量", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "单位", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "频次", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "数量", adDouble, 18, adFldIsNullable
        .Fields.Append "药名id", adDouble, 18, adFldIsNullable
        .Fields.Append "执行时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "是否皮试", adDouble, 1, adFldIsNullable
        .Fields.Append "审查结果", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "审核标志", adDouble, 1, adFldIsNullable
        .Fields.Append "药师审核原因", adLongVarChar, 500, adFldIsNullable
        .Fields.Append "药品id", adDouble, 18, adFldIsNullable
        .Fields.Append "配药类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "医嘱期效", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "给药途径", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "执行性质", adDouble, 1, adFldIsNullable
        .Fields.Append "执行标记", adDouble, 1, adFldIsNullable
        
        '医嘱信息
        .Fields.Append "开嘱时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "皮试结果", adLongVarChar, 20, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrExecute = GetArrayByStr(strIDS, 3950, ",")
    For i = 0 To UBound(arrExecute)
        Set rsMedi = Piva_GetMedi(intType, CStr(arrExecute(i)), mParams.intCheck)
        With rsMedi
            Do While Not .EOF
                mrsMedi.AddNew
                If lng医嘱id <> !相关id Then
                    lng医嘱id = !相关id
                    lng医嘱号 = lng医嘱号 + 1
                    bln操作 = False
                End If
                
                If !配药类型 = str配药类型 Then bln操作 = True
                
                mrsMedi!组号 = lng医嘱号
                mrsMedi!Id = !Id
                mrsMedi!相关id = !相关id
                mrsMedi!住院号 = !住院号
                mrsMedi!姓名 = !姓名
                mrsMedi!性别 = !性别
                mrsMedi!年龄 = !年龄
                mrsMedi!床号 = !床号
                mrsMedi!病人ID = !病人ID
                mrsMedi!主页id = !主页id
                mrsMedi!科室 = !科室名称
                mrsMedi!开单医生 = !开嘱医生
                mrsMedi!药品名称 = !药品名称
                mrsMedi!规格 = zlStr.nvl(!规格)
                mrsMedi!单量 = nvl(!单次用量, 0)
                mrsMedi!单位 = !计算单位
                mrsMedi!频次 = !执行频次
                mrsMedi!药名ID = !药名ID
                mrsMedi!药品ID = !药品ID
                mrsMedi!病人病区ID = !病人病区ID
                mrsMedi!病人科室id = !病人科室id
                mrsMedi!执行时间 = !执行时间方案
                mrsMedi!是否皮试 = !是否皮试
                mrsMedi!开嘱时间 = !开嘱时间
                mrsMedi!皮试结果 = !皮试结果
                mrsMedi!审查结果 = zlStr.nvl(!审查结果, 0)
                mrsMedi!审核标志 = !审核标志
                mrsMedi!药师审核原因 = !药师审核原因
                mrsMedi!配药类型 = !配药类型
                mrsMedi!医嘱期效 = !医嘱期效
                mrsMedi!给药途径 = !给药途径
                mrsMedi!执行性质 = nvl(!执行性质, 0)
                mrsMedi!执行标记 = nvl(!执行标记, 0)
                mrsMedi.Update
                .MoveNext
                
                If .EOF Then
                    If bln操作 = False And lng医嘱号 <> 0 And str配药类型 <> "" Then
                        mrsMedi.Filter = "组号=" & lng医嘱号
                        Do While Not mrsMedi.EOF
                            mrsMedi.Delete adAffectCurrent
                            mrsMedi.MoveNext
                        Loop
                    End If
                Else
                
                    If lng医嘱id <> !相关id Then
                        If bln操作 = False And lng医嘱号 <> 0 And str配药类型 <> "" Then
                            mrsMedi.Filter = "组号=" & lng医嘱号
                            Do While Not mrsMedi.EOF
                                mrsMedi.Delete adAffectCurrent
                                mrsMedi.MoveNext
                            Loop
                        End If
                    End If
                End If
            Loop
        End With
    Next
End Function

Private Sub LoadVsfMedi(ByVal strIDS As String, Optional ByVal blnFilter As Boolean)
    Dim rstemp As Recordset
    Dim i As Long
    Dim lng医嘱id As Long
    Dim j As Integer
    Dim lng病人id As Long
    Dim intBlackColor As Integer
    Dim intType As Integer
    Dim dateCurrent As Date
    Dim strFilter As String
    Dim str配药类型  As String
    
    mintBeginRow = 0
    mintEndRow = 0
        
    If mcondition.strTransStep = M_STR_CALSS_AUDIT Then
        intType = 0
    ElseIf mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Then
        intType = 1
    ElseIf mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
        intType = 2
    End If
    
    str配药类型 = Me.cboType.Text
    If Me.cboType.ListIndex = 0 Then str配药类型 = ""
    If Not blnFilter Then
        Call InitMedi(intType, strIDS, str配药类型)
    End If
    
    For i = 0 To Me.ImgResult.count - 1
        If Me.chkResult(i).Value = 1 Then
            strFilter = IIf(strFilter = "", "审查结果=" & i, strFilter & " or 审查结果=" & i)
        End If
    Next
    
    If mrsMedi Is Nothing Then Exit Sub
    mrsMedi.Filter = strFilter
    With Me.vsfMedis
        If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
            .ColHidden(.ColIndex("审")) = True
            .ColHidden(.ColIndex("选择")) = False
        ElseIf mcondition.strTransStep = M_STR_CALSS_AUDIT Then
            .ColHidden(.ColIndex("审")) = False
            .ColHidden(.ColIndex("选择")) = True
        End If
        
        .rows = 1
        .rows = 2
        If mrsMedi.RecordCount = 0 Then Exit Sub
        
        dateCurrent = Sys.Currentdate
        
        .Redraw = flexRDNone
        
        .MergeCells = flexMergeFree
        .rows = mrsMedi.RecordCount + 1
        
        i = 1
        mrsMedi.MoveFirst
        Do While Not mrsMedi.EOF
            If lng病人id <> Val(mrsMedi!病人ID) Then
                lng病人id = Val(mrsMedi!病人ID)
                If intBlackColor <> 1 Then
                    intBlackColor = 1
                Else
                    intBlackColor = 2
                End If
            End If
            
            If lng医嘱id <> Val(mrsMedi!相关id) Then
                If lng医嘱id <> 0 Then
                    .rows = .rows + 1
                    .RowHidden(i) = True
                    For j = 0 To .Cols - 1
                        .TextMatrix(i, j) = "00"
                    Next
                    i = i + 1
                End If
                lng医嘱id = Val(mrsMedi!相关id)
            Else
                .MergeCol(.ColIndex("选择")) = True
                .MergeCol(.ColIndex("审")) = True
                .MergeCol(.ColIndex("医嘱id")) = True
                .MergeCol(.ColIndex("姓名")) = True
                .MergeCol(.ColIndex("性别")) = True
                .MergeCol(.ColIndex("年龄")) = True
                .MergeCol(.ColIndex("床号")) = True
                .MergeCol(.ColIndex("住院号")) = True
                .MergeCol(.ColIndex("科室")) = True
                .MergeCol(.ColIndex("开单医生")) = True
                .MergeCol(.ColIndex("背景号")) = True
                .MergeCol(.ColIndex("名字")) = True
                .MergeCol(.ColIndex("病区ID")) = True
                .MergeCol(.ColIndex("科室ID")) = True
                .MergeCol(.ColIndex("病人ID")) = True
                .MergeCol(.ColIndex("主页ID")) = True
                .MergeCol(.ColIndex("警告")) = True
                .MergeCol(.ColIndex("药品id")) = True
                .MergeCol(.ColIndex("药名")) = True
            End If
            
            .TextMatrix(i, .ColIndex("标志")) = "0"
            .TextMatrix(i, .ColIndex("审")) = " "
            .TextMatrix(i, .ColIndex("已审核")) = Val(mrsMedi!审核标志)
            .TextMatrix(i, .ColIndex("背景号")) = intBlackColor
            .TextMatrix(i, .ColIndex("病人id")) = Val(mrsMedi!病人ID)
            .TextMatrix(i, .ColIndex("主页id")) = Val(mrsMedi!主页id)
            .TextMatrix(i, .ColIndex("相关id")) = Val(mrsMedi!相关id)
            .TextMatrix(i, .ColIndex("医嘱id")) = Val(mrsMedi!Id)
            .TextMatrix(i, .ColIndex("姓名")) = zlStr.nvl(mrsMedi!姓名)
            .TextMatrix(i, .ColIndex("性别")) = zlStr.nvl(mrsMedi!性别)
            .TextMatrix(i, .ColIndex("年龄")) = zlStr.nvl(mrsMedi!年龄)
            .TextMatrix(i, .ColIndex("床号")) = IIf(zlStr.nvl(mrsMedi!床号) = "", "<空>", mrsMedi!床号)
            .TextMatrix(i, .ColIndex("住院号")) = zlStr.nvl(mrsMedi!住院号, " ")
            .TextMatrix(i, .ColIndex("名字")) = zlStr.nvl(mrsMedi!姓名, " ")
            .TextMatrix(i, .ColIndex("病区ID")) = zlStr.nvl(mrsMedi!病人病区ID, " ")
            .TextMatrix(i, .ColIndex("科室ID")) = zlStr.nvl(mrsMedi!病人科室id, " ")
            .TextMatrix(i, .ColIndex("科室")) = zlStr.nvl(mrsMedi!科室)
            .TextMatrix(i, .ColIndex("开单医生")) = zlStr.nvl(mrsMedi!开单医生)
            .TextMatrix(i, .ColIndex("药品名称")) = zlStr.nvl(mrsMedi!药品名称) & IIf(zlStr.nvl(mrsMedi!规格) = "", "", "，" & zlStr.nvl(mrsMedi!规格))
            .TextMatrix(i, .ColIndex("药师审核原因")) = zlStr.nvl(mrsMedi!药师审核原因)
            .TextMatrix(i, .ColIndex("警告")) = nvl(mrsMedi!审查结果)
            .TextMatrix(i, .ColIndex("药品id")) = Val(nvl(mrsMedi!药品ID))
            .TextMatrix(i, .ColIndex("药名")) = nvl(mrsMedi!药品名称)
            .TextMatrix(i, .ColIndex("期效")) = nvl(mrsMedi!医嘱期效)
            .TextMatrix(i, .ColIndex("给药途径")) = nvl(mrsMedi!给药途径)
            
            
            If Val(mrsMedi!审核标志) <> 0 And Val(mrsMedi!审核标志) <> 3 And mcondition.strTransStep = M_STR_CALSS_AUDIT Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &HFF0000
            End If
            
            
            '设置合理用药标志 (PASS)
            If Not gobjPass Is Nothing Then
                .Cell(flexcpPicture, i, .ColIndex("警"), i, .ColIndex("警")) = gobjPass.zlPassSetWarnLight_YF(nvl(mrsMedi!审查结果, 0))
                .Cell(flexcpPictureAlignment, i, .ColIndex("警"), i, .ColIndex("警")) = flexPicAlignCenterCenter
            End If

            '显示[自备药]标志
            If mrsMedi!执行性质 = 5 And mrsMedi!执行标记 = 0 Then
                .Cell(flexcpPicture, i, .ColIndex("药品名称"), i, .ColIndex("药品名称")) = Me.ImgPro.ListImages("自备药").Picture
                .Cell(flexcpPictureAlignment, i, .ColIndex("药品名称"), i, .ColIndex("药品名称")) = flexPicAlignLeftCenter
            End If

            If mrsMedi!是否皮试 = 1 Then
                .TextMatrix(i, .ColIndex("皮")) = Get皮试结果(Val(mrsMedi!病人ID), Val(mrsMedi!药名ID), dateCurrent, mrsMedi!开嘱时间, mrsMedi!主页id)
            End If
            
            .TextMatrix(i, .ColIndex("规格")) = zlStr.nvl(mrsMedi!规格)
            .TextMatrix(i, .ColIndex("单量")) = Format(zlStr.nvl(mrsMedi!单量), "#####0.00000;-#####0.00000; ;")
            .TextMatrix(i, .ColIndex("单位")) = zlStr.nvl(mrsMedi!单位)
            .TextMatrix(i, .ColIndex("频率")) = zlStr.nvl(mrsMedi!频次)
            .TextMatrix(i, .ColIndex("执行时间")) = zlStr.nvl(mrsMedi!执行时间)
            
            If mcondition.strTransStep = M_STR_CALSS_PASSEDAUDIT Or mcondition.strTransStep = M_STR_CALSS_FAILAUDIT Then
                .TextMatrix(i, .ColIndex("摆药标志")) = IIf(CheckIs摆药(Val(mrsMedi!相关id)) = True, 1, 0)
            End If
            If Val(.TextMatrix(i, .ColIndex("摆药标志"))) = 1 Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = vbRed
            End If
            
            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = IIf(intBlackColor = 1, &H80000005, &HC0FFC0)
            i = i + 1
            mrsMedi.MoveNext
        Loop
        
        .Cell(flexcpFontBold, 0, .ColIndex("审"), 0, .ColIndex("审")) = True
        .Cell(flexcpForeColor, 0, .ColIndex("审"), 0, .ColIndex("审")) = vbBlue
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Function CheckIs摆药(ByVal lng医嘱id As Long) As Boolean
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select 1 from 输液配药记录 A,病人医嘱记录 B where A.医嘱id=B.id and A.操作状态>1 and A.操作状态<>12 and b.id=[1]"
    
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIs摆药", lng医嘱id)
    
    CheckIs摆药 = (Not rstemp.EOF)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mobjMipModule_ReceiveMessage(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    '1.接收消息：消息类型和上游业务约定
    '2.根据客户机参数设置判断是否是有效消息
    Dim i As Integer
    Const CST_INT_MSGREFRESHINTERVAL As Integer = 1
    Const CST_STR_MSGCODE As String = "ZLHIS_CIS_003,ZLHIS_CIS_013,ZLHIS_CIS_008"
    
    '消息对象为空时退出
    If mobjMipModule Is Nothing Then Exit Sub
    
    '消息服务连接失败时不接收消息
    If mobjMipModule.IsConnect = False Then Exit Sub
        
    '检查输液配置中心接收的消息类型
    If InStr("," & CST_STR_MSGCODE & ",", "," & strMsgItemIdentity & ",") = 0 Then Exit Sub

    '根据客户机参数设置判断是否是有效消息
    Call IsValidMsg(strMsgItemIdentity, strMsgContent)
    
    
    '如果接收到有效消息时立即更新表格
    If Not mrsMsg Is Nothing Then
    lblMsgComment.Caption = "消息提醒(" & mrsMsg.RecordCount & ")"
    mrsMsg.MoveFirst
    With Me.vsfMsg
        .rows = 1
        .rows = mrsMsg.RecordCount + 1
        .RowHeight(0) = 300
        
        For i = 1 To mrsMsg.RecordCount
            .RowHeight(i) = 300
            .TextMatrix(i, .ColIndex("时间")) = mrsMsg!时间
            .TextMatrix(i, .ColIndex("类型")) = mrsMsg!类型
            .TextMatrix(i, .ColIndex("病区")) = mrsMsg!病区
            .TextMatrix(i, .ColIndex("病人")) = mrsMsg!病人
            .TextMatrix(i, .ColIndex("执行时间")) = mrsMsg!执行时间
'            .TextMatrix(i, .ColIndex("操作状态")) = mrsMsg!操作状态
            .TextMatrix(i, .ColIndex("病人病区id")) = mrsMsg!病区ID
            .TextMatrix(i, .ColIndex("科室id")) = mrsMsg!科室ID
            
            mrsMsg.MoveNext
        Next
    End With
    End If
End Sub

Private Sub IsValidMsg(ByVal strMsgItemIdentity As String, ByVal strMsgContent As String)
    Dim rsMsg As Recordset
    Dim strTemp As String
    Dim strSQL As String
    Dim rstemp As Recordset
    Dim blnNext As Boolean
    Dim str科室名称 As String
    Dim str病区名称 As String
    Dim str时间 As String
    Dim str类型 As String
    Dim objXML As New zl9ComLib.clsXML
    
    On Error GoTo ErrHand
    
    If objXML Is Nothing Then Exit Sub

    '打开XML文件
    objXML.OpenXMLDocument strMsgContent
    
    If strMsgItemIdentity = "ZLHIS_CIS_003" Then
'        str类型 = "医嘱作废"
'
'        If objXML.GetMultiNodeRecord("cancel_order", rsMsg) = False Then Exit Sub
'        If rsMsg Is Nothing Then Exit Sub
'        If rsMsg.RecordCount = 0 Then Exit Sub
'
'        '获取医嘱ID,检查该医嘱ID是否产生配药记录
'        If objXML.GetSingleNodeValue("order_id", strTemp, xsString) = False Then Exit Sub
'        If objXML.GetSingleNodeValue("cancel_time", str时间, xsString) = False Then Exit Sub
'
'        strSQL = "select A.ID,A.医嘱id,A.姓名,A.性别,A.年龄,A.床号,A.执行时间,A.操作状态,B.ID 病区ID,B.名称 病区名称,C.ID 科室ID,C.名称 科室名称 from 输液配药记录 A ,部门表 B,部门表 C where B.id=A.病人病区ID And  C.id=A.病人科室ID And A.医嘱id=[1] and A.部门ID=[2]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng配置中心)
'        If Not rsTemp.EOF Then blnNext = True
        
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_013" Then
        str类型 = "销帐申请"
    
        If objXML.GetMultiNodeRecord("cancel_reqeust", rsMsg) = False Then Exit Sub
        If rsMsg Is Nothing Then Exit Sub
        If rsMsg.RecordCount = 0 Then Exit Sub
        
        '获取配药ID,检查该配药记录是否为当前部门的销帐申请记录
        If objXML.GetSingleNodeValue("transfusion_id", strTemp, xsString) = False Then Exit Sub
        If objXML.GetSingleNodeValue("request_time", str时间, xsString) = False Then Exit Sub
        
        strSQL = "select A.ID,A.医嘱id,A.姓名,A.性别,A.年龄,A.床号,A.执行时间,A.操作状态,B.ID 病区ID,B.名称 病区名称,C.ID 科室ID,C.名称 科室名称 from 输液配药记录 A ,部门表 B,部门表 C where B.id=A.病人病区ID And  C.id=A.病人科室ID And A.Id=[1] and A.部门ID=[2] "
        Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng配置中心)
        If Not rstemp.EOF Then blnNext = True
        
    ElseIf strMsgItemIdentity = "ZLHIS_CIS_008" Then
        str类型 = "批次调整"
        str时间 = Now
    
        If objXML.GetMultiNodeRecord("transfusion_info", rsMsg) = False Then Exit Sub
        If rsMsg Is Nothing Then Exit Sub
        If rsMsg.RecordCount = 0 Then Exit Sub
        
        '获取配药ID,检查该配药记录是否为当前部门的输液配药记录
        If objXML.GetSingleNodeValue("transfusion_id", strTemp, xsString) = False Then Exit Sub
        
        strSQL = "select A.ID,A.医嘱id,A.姓名,A.性别,A.年龄,A.床号,A.执行时间,A.操作状态,B.ID 病区ID,B.名称 病区名称,C.ID 科室ID,C.名称 科室名称 from 输液配药记录 A ,部门表 B,部门表 C where B.id=A.病人病区ID And  C.id=A.病人科室ID And A.Id=[1] and A.部门ID=[2]"
        Set rstemp = zlDatabase.OpenSQLRecord(strSQL, "IsValidMsg", strTemp, mParams.lng配置中心)
        If Not rstemp.EOF Then blnNext = True
        
    End If
    
    
    '数据满足条件，更新数据集和界面数据
    If blnNext Then
        Call mobjMipModule.ShowMessage(strMsgItemIdentity, "发现有" & str类型 & ",请操作员注意查看", str类型 & "提醒", "提醒工作人员", "任务id=1234|病人id=344899")
    
        If mrsMsg Is Nothing Then Call InitMsgRs
        With mrsMsg
            .AddNew
            !配药id = rstemp!Id
            !病人 = rstemp!姓名 & " " & rstemp!性别 & " " & rstemp!年龄 & " " & rstemp!科室名称 & " " & rstemp!床号          '姓名，性别，年龄，科室，床位
            !病区 = rstemp!病区名称
            !时间 = str时间
            !类型 = str类型
            !医嘱id = rstemp!医嘱id
            !执行时间 = rstemp!执行时间
            !操作状态 = rstemp!操作状态
            !病区ID = rstemp!病区ID
            !科室ID = rstemp!科室ID
            
            .Update
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitMsgRs()
    Set mrsMsg = New ADODB.Recordset
    With mrsMsg
        If .State = 1 Then .Close
        
        '该记录对应的消息信息
        .Fields.Append "配药ID", adDouble, 18, adFldIsNullable
        .Fields.Append "类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "病区", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "医嘱ID", adDouble, 3, adFldIsNullable
        .Fields.Append "时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "病人", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "执行时间", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "操作状态", adDouble, 3, adFldIsNullable
        .Fields.Append "病区ID", adDouble, 18, adFldIsNullable
        .Fields.Append "科室ID", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Public Sub SendMsgModule()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消息发送处理
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, objDrugXML As New clsXML, objCheckXML As New clsXML
    Dim objTemp As clsXML, str收费时间 As String
    Dim rstemp As ADODB.Recordset, int性质 As Integer
    Dim bln直接收费 As Boolean, p As Long
    Dim lngDrug As Long, lngCheck As Long, blnAddBill As Boolean, blnHaveCheck As Boolean, blnHaveDrug As Boolean
    On Error GoTo errHandle
    Dim i As Integer
    
    If mobjMipModule Is Nothing Then Exit Sub
'    If mobjMipModule.IsConnect = False Then Exit Sub
        
    
'    transfuse_order 医嘱信息
'    order_id 医嘱id
'    order_reason 审核原因
'    send_serial 发送号
'    in_patient 病人信息
'    patient_id 病人id
'    patient_name 姓名
'    in_number 住院号
'    patient_clinic 就诊信息
'    clinic_id 主页id
'    clinic_area_id 就诊病区id
'    clinic_dept_id 就诊科室id
'    clinic_bed 就诊病床


    objDrugXML.ClearXmlText
    objCheckXML.ClearXmlText
    
    With mrsSendMsg
        If .RecordCount = 0 Then Exit Sub
        
        .MoveFirst
        For i = 1 To .RecordCount
            '查询拒绝理由
            gstrSQL = "Select 相关id From 病人医嘱记录 Where ID = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询拒绝理由", !医嘱id)
            
            gstrSQL = "Select Max(药师审核原因) As 审核原因 From 病人医嘱记录 Where 相关id = [1]"
            Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "查询拒绝理由", rstemp!相关id)
        
            '医嘱信息
            Call objDrugXML.AppendNode("transfuse_order")
                Call objDrugXML.AppendData("order_id", !医嘱id)
                Call objDrugXML.AppendData("send_serial", !发送号)
                Call objDrugXML.AppendData("order_reason", rstemp!审核原因)
            
            '病人信息
            Call objDrugXML.AppendNode("in_patient")
                Call objDrugXML.AppendData("patient_id", !病人ID)
                Call objDrugXML.AppendData("patient_name", !姓名)
                Call objDrugXML.AppendData("in_number", !住院号)
            Call objDrugXML.AppendNode("in_patient", True)
            
            '就诊信息
            Call objDrugXML.AppendNode("patient_clinic")
                Call objDrugXML.AppendData("clinic_id", !主页id)
                Call objDrugXML.AppendData("clinic_area_id", !病区ID)
                Call objDrugXML.AppendData("clinic_dept_id", !科室ID)
                Call objDrugXML.AppendData("clinic_bed", IIf(!床号 = "<空>", "", Replace(zlStr.nvl(!床号, ""), "床", "")))
            Call objDrugXML.AppendNode("patient_clinic", True)
            
            
            Call objDrugXML.AppendNode("transfuse_order", True)
            '发送消息
'            Call zlDebugWriteFile(objDrugXML.XmlText)
            Call mobjMipModule.CommitMessage("ZLHIS_TRANSFUSION_001", objDrugXML.XmlText)
            Call zlDatabase.SendMsg("ZLHIS_TRANSFUSION_001", objDrugXML.XmlText)
            objDrugXML.ClearXmlText: objCheckXML.ClearXmlText
            Set objDrugXML = Nothing: Set objCheckXML = Nothing
            
            .MoveNext
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
 
 
 
Private Sub InitSendMsgRs()
    Set mrsSendMsg = New ADODB.Recordset
    With mrsSendMsg
        If .State = 1 Then .Close
        
        '该记录对应的消息信息
        .Fields.Append "医嘱id", adDouble, 18, adFldIsNullable
        .Fields.Append "发送号", adDouble, 18, adFldIsNullable
        .Fields.Append "病人ID", adDouble, 18, adFldIsNullable
        .Fields.Append "姓名", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "住院号", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "主页id", adDouble, 18, adFldIsNullable
        .Fields.Append "病区id", adDouble, 18, adFldIsNullable
        .Fields.Append "科室id", adDouble, 18, adFldIsNullable
        .Fields.Append "床号", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub












