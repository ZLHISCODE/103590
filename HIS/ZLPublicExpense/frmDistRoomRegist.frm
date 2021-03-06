VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\A..\zlIDKind\zlIDKind.vbp"
Begin VB.Form frmDistRoomRegist 
   Caption         =   "门诊分诊挂号"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11685
   Icon            =   "frmDistRoomRegist.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   11685
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdPrice 
      Caption         =   "保存划价单(&J)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5655
      TabIndex        =   53
      Top             =   6420
      Width           =   1725
   End
   Begin VB.TextBox txtGender 
      BackColor       =   &H8000000F&
      Height          =   330
      Left            =   3345
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   525
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9210
      TabIndex        =   12
      Top             =   6420
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7995
      TabIndex        =   11
      Top             =   6420
      Width           =   1100
   End
   Begin VB.Frame fraPay 
      Height          =   750
      Left            =   5625
      TabIndex        =   39
      Top             =   5550
      Width           =   4740
      Begin VB.TextBox txtPayMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   210
         Width           =   1635
      End
      Begin VB.ComboBox cboPayMode 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   210
         Width           =   1500
      End
      Begin VB.Label lblPayMode 
         AutoSize        =   -1  'True
         Caption         =   "支付方式"
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
         Left            =   120
         TabIndex        =   40
         Top             =   270
         Width           =   1200
      End
   End
   Begin VB.Frame fraTotal 
      Height          =   1095
      Left            =   5625
      TabIndex        =   36
      Top             =   4470
      Width           =   4740
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3465
         TabIndex        =   38
         Top             =   330
         Width           =   960
      End
      Begin VB.Label lblSum 
         Caption         =   "合 计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   15
         TabIndex        =   37
         Top             =   135
         Width           =   645
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   3165
      Left            =   5625
      TabIndex        =   28
      Top             =   1305
      Width           =   4740
      Begin VB.TextBox txtSN 
         Height          =   255
         Left            =   2220
         TabIndex        =   52
         Top             =   1035
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboRemark 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         TabIndex        =   9
         Top             =   2760
         Width           =   1500
      End
      Begin VB.ComboBox cboAppointStyle 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2760
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   330
         Left            =   3165
         TabIndex        =   5
         Top             =   615
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192194
         CurrentDate     =   42121
      End
      Begin VB.CheckBox chkBook 
         Caption         =   " 购买病历"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   7
         Top             =   1065
         Width           =   1485
      End
      Begin VB.TextBox txtRegistTime 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   4
         EndProperty
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         TabIndex        =   34
         Top             =   615
         Width           =   1500
      End
      Begin VB.TextBox txtDept 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3165
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1500
      End
      Begin VB.ComboBox cboRoom 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   1500
      End
      Begin VB.ComboBox cboDoctor 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   615
         Width           =   1500
      End
      Begin VB.TextBox txtArrangeNO 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   570
         TabIndex        =   2
         Top             =   180
         Width           =   1500
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfMoney 
         Height          =   1290
         Left            =   75
         TabIndex        =   35
         Top             =   1425
         Width           =   4575
         _cx             =   8070
         _cy             =   2275
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDistRoomRegist.frx":058A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
      Begin VB.Label lblRemark 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   51
         Top             =   2820
         Width           =   420
      End
      Begin VB.Label lblAppointStyle 
         AutoSize        =   -1  'True
         Caption         =   "预约方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   50
         Top             =   2820
         Width           =   840
      End
      Begin VB.Label lblRegistTime 
         AutoSize        =   -1  'True
         Caption         =   "挂号时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2220
         TabIndex        =   33
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2640
         TabIndex        =   32
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         Caption         =   "诊室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   31
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblDoctor 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   30
         Top             =   675
         Width           =   420
      End
      Begin VB.Label lblArrangeNO 
         AutoSize        =   -1  'True
         Caption         =   "号别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   29
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame fraTime 
      Height          =   5805
      Left            =   30
      TabIndex        =   25
      Top             =   975
      Width           =   5520
      Begin VB.ComboBox cboDeptFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2460
         TabIndex        =   45
         Top             =   165
         Width           =   1275
      End
      Begin VB.ComboBox cboDoctorFilter 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmDistRoomRegist.frx":0625
         Left            =   4200
         List            =   "frmDistRoomRegist.frx":0627
         TabIndex        =   44
         Top             =   165
         Width           =   1275
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfArrange 
         Height          =   2700
         Left            =   105
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   540
         Width           =   5385
         _cx             =   9499
         _cy             =   4762
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDistRoomRegist.frx":0629
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   1
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetailTime 
         Height          =   2370
         Left            =   60
         TabIndex        =   26
         Top             =   3360
         Width           =   5385
         _cx             =   9499
         _cy             =   4180
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   18
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
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
         AutoResize      =   0   'False
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
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   330
         Left            =   525
         TabIndex        =   46
         Top             =   165
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   93323267
         CurrentDate     =   42335
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   49
         Top             =   225
         Width           =   420
      End
      Begin VB.Label lblDeptFilter 
         AutoSize        =   -1  'True
         Caption         =   "科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2025
         TabIndex        =   48
         Top             =   225
         Width           =   420
      End
      Begin VB.Label lblDoctorFilter 
         AutoSize        =   -1  'True
         Caption         =   "医生"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3780
         TabIndex        =   47
         Top             =   225
         Width           =   420
      End
   End
   Begin VB.Frame fraSplit 
      Height          =   45
      Left            =   -60
      TabIndex        =   24
      Top             =   945
      Width           =   11000
   End
   Begin VB.TextBox txtAge 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4665
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   525
      Width           =   705
   End
   Begin VB.TextBox txtClinic 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   525
      Width           =   1470
   End
   Begin VB.TextBox txtPatient 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Top             =   525
      Width           =   1500
   End
   Begin VB.TextBox txtFeeType 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8610
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   525
      Width           =   1185
   End
   Begin VB.ComboBox cboNO 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8625
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   1755
   End
   Begin zlIDKind.IDKindNew IDKind 
      Height          =   330
      Left            =   600
      TabIndex        =   23
      Top             =   525
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   582
      Appearance      =   2
      IDKindStr       =   "姓|姓名或就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontSize        =   11.25
      FontName        =   "宋体"
      IDKind          =   -1
      DefaultCardType =   "0"
      BackColor       =   -2147483633
   End
   Begin VB.Label lbl急 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "急"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Left            =   135
      TabIndex        =   43
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblMoney 
      AutoSize        =   -1  'True
      Caption         =   "门诊预交余额:0.00     "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5625
      TabIndex        =   27
      Top             =   1065
      Width           =   2310
   End
   Begin VB.Label lblFeeType 
      AutoSize        =   -1  'True
      Caption         =   "费别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   8130
      TabIndex        =   19
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblPatient 
      AutoSize        =   -1  'True
      Caption         =   "病人"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   18
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblGender 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2835
      TabIndex        =   17
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4170
      TabIndex        =   16
      Top             =   585
      Width           =   420
   End
   Begin VB.Label lblClinic 
      AutoSize        =   -1  'True
      Caption         =   "门诊号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5685
      TabIndex        =   15
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "单据号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7935
      TabIndex        =   0
      Top             =   150
      Width           =   630
   End
End
Attribute VB_Name = "frmDistRoomRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModul As Long, mstrPrivs As String

Private WithEvents mobjIDCard As zlIDCard.clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private mblnCard As Boolean, mblnStartFactUseType As Boolean
Private mstrYBPati As String, mlng挂号ID As Long, mlng领用ID As Long
Private mblnOlnyBJYB As Boolean, mblnSharedInvoice As Boolean
Private mstr险类 As String, mblnChangeFeeType As Boolean
Private mstrPassWord As String, mstrInsure As String
Private mstrDeptIDs As String, mlngRow As Long, msngTime As Single
Private Const SNCOLS = 10
Private Const SnArgCols = 7
Private mrsPlan As ADODB.Recordset
Private mrsSNState As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset, mblnAppointment As Boolean
Private mrsItems As ADODB.Recordset
Private mrs时间段 As ADODB.Recordset
Private mrsInComes As ADODB.Recordset
Private mcolCardPayMode As Collection, mblnInit As Boolean
Private mcur个帐余额 As Currency, mintSysAppLimit As Integer
Private mblnOK As Boolean, mstrCardPass As String
Private mstrNO As String, mintIDKind As Integer
Private mintInsure As Integer, mstrUseType As String
Private mdatLast As Date, mblnUpdateAge As Boolean
Private mblnChangeByCode As Boolean, mblnFilterChange As Boolean
Private mstrCardNO As String
Private mcur个帐透支 As Currency
Private Enum EM_REGISTFEE_MODE  '挂号费用收取方式
        EM_RG_现收 = 0
        EM_RG_划价 = 1
        EM_RG_记帐 = 2
End Enum
Private Enum EM_PATI_CHARGE_MODE    '病人收费模式
    EM_先结算后诊疗 = 0
    EM_先诊疗后结算 = 1
End Enum
Private mRegistFeeMode As EM_REGISTFEE_MODE '挂号费用收取方式
Private mPatiChargeMode As EM_PATI_CHARGE_MODE    '病人收费模式

Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    使用个人帐户   As Boolean  'support挂号使用个人帐户
    不收病历费 As Boolean   'support挂号不收取病历费
    挂号检查项目 As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private Enum ViewMode
     V_普通号
     v_专家号
     v_专家号分时段
     V_普通号分时段
End Enum
Private mViewMode As ViewMode

Private Type ty_ModulePara
    bln姓名模糊查找 As Boolean
    lng姓名查找天数 As Long
    bln默认购买病历 As Boolean
    bln默认输入摘要 As Boolean
    byt挂号模式 As Byte
    bln挂号必须刷卡 As Boolean
    bln优先使用预交 As Boolean
    bln住院病人挂号 As Boolean
    int挂号发票打印 As Integer
    int挂号凭条打印 As Integer
    int预约挂号打印 As Integer
    bln随机序号选择 As Boolean
    lng预约有效时间 As Long
    bln共用收费票据 As Boolean
    bln预约时收款 As Boolean
    bln退号重用 As Boolean
    dbl预存款消费验卡 As Double '预存款消费刷卡控制：0-不进行刷卡控制,1-门诊消费时需要刷卡验证,2-门诊消费时设置密码的，则必须刷卡验证
    int同科限约数           As Integer  '同科室限约
    int同科限挂数           As Integer
    bln同科限挂急诊         As Boolean
    int病人预约科室数       As Integer
    int病人挂号科室数       As Integer
    int专家号挂号限制       As Integer
    int专家号预约限制       As Integer
    bln挂号生成队列 As Boolean
    int预约限制时间 As Integer
    int预约缺省天数  As Integer
    bln严格按时段挂号  As Boolean  '严格按时段挂号

End Type

Private mty_Para As ty_ModulePara
Private mstr家属IDs As String
Private mstrPriceGrade As String, mintPriceGradeStartType As Integer
Private mobjRegistData As clsRegist
Private mblnNotClick As Boolean
Private mstrPreNO As String
Private mlngPreRow As Long
Private mblnUnChange As Boolean
Private mint号长 As Integer   '号别长度

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long, ByVal strDeptIDs As String, _
                    ByRef strOutNO As String, ByVal blnAppointment As Boolean)
    mlngModul = lngModul
    mstrDeptIDs = strDeptIDs
    mblnAppointment = blnAppointment
    Set mobjRegistData = New clsRegist '因为存在oracle连接对象，因此，不用初始化
    mint号长 = mobjRegistData.zlGetRegisterMaxLen_Tradition
    
    If frmMain Is Nothing Then
        Me.Show
    Else
        Me.Show 1, frmMain
    End If
    If mblnOK = True Then
        strOutNO = mstrNO
        Unload Me
    End If
End Sub

Private Sub InitPara()
    Dim strValue As String
    With mty_Para
        .bln姓名模糊查找 = Val(gobjDatabase.GetPara("姓名模糊查找", glngSys, 9000, "0")) = 1
        .lng姓名查找天数 = Val(gobjDatabase.GetPara("姓名查找天数", glngSys, 9000, 0))
        .bln默认购买病历 = Val(gobjDatabase.GetPara("默认购买病历", glngSys, 9000, "0")) = 1
        .bln默认输入摘要 = Val(gobjDatabase.GetPara("默认输入摘要", glngSys, 9000, "1")) = 1
        .byt挂号模式 = Val(gobjDatabase.GetPara("挂号模式", glngSys, 9000, "0"))
        .bln优先使用预交 = Val(gobjDatabase.GetPara("优先使用预交", glngSys, 9000, "0")) = 1
        .bln住院病人挂号 = Val(gobjDatabase.GetPara("允许住院病人挂号", glngSys, 9000, "0")) = 1
        .int挂号发票打印 = Val(gobjDatabase.GetPara("挂号发票打印方式", glngSys, 9000, "0"))
        .int挂号凭条打印 = Val(gobjDatabase.GetPara("挂号凭条打印方式", glngSys, 9000, "0"))
        .int预约挂号打印 = Val(gobjDatabase.GetPara("预约挂号单打印方式", glngSys, 9000, "0"))
        .bln随机序号选择 = Val(gobjDatabase.GetPara("随机序号选择", glngSys, 9000, "0")) = 1
        .bln共用收费票据 = Val(gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121)) = 1
        .bln预约时收款 = Val(gobjDatabase.GetPara("预约时收款", glngSys, 9000, "0")) = 1
        .bln退号重用 = Val(gobjDatabase.GetPara("已退序号允许挂号", glngSys, 1111)) = 1
        .bln挂号必须刷卡 = Val(gobjDatabase.GetPara("挂号必须刷卡", glngSys, 9000)) = 1
        strValue = gobjDatabase.GetPara(28, glngSys, , "1|0")
        If InStr(strValue, "|") = 0 Then strValue = "1|0"
        .dbl预存款消费验卡 = Val(Split(strValue, "|")(0))
        .int同科限约数 = Val(gobjDatabase.GetPara("病人同科限约N个号", glngSys, 1111, 0))
        .int同科限挂数 = Val(Split(gobjDatabase.GetPara("病人同科限挂N个号", glngSys, 1111, 0) & "|", "|")(0))
        .bln同科限挂急诊 = Split(gobjDatabase.GetPara("病人同科限挂N个号", glngSys, 1111, 0) & "|", "|")(1) = "1"
        .int病人挂号科室数 = Val(gobjDatabase.GetPara("病人挂号科室限制", glngSys, 1111, 0))
        .int病人预约科室数 = Val(gobjDatabase.GetPara("病人预约科室数", glngSys, 1111, 0))
        .int专家号挂号限制 = Val(gobjDatabase.GetPara("专家号挂号限制", glngSys, , 0))
        .int专家号预约限制 = Val(gobjDatabase.GetPara("专家号预约限制", glngSys, , 0))
        .bln挂号生成队列 = Val(gobjDatabase.GetPara("排队叫号模式", glngSys, 1113)) <> 0
        If .bln默认输入摘要 Then
            cboRemark.TabStop = True
        Else
            cboRemark.TabStop = False
        End If
        If .bln默认购买病历 Then
            chkBook.Value = 1
        End If
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If .byt挂号模式 = 0 Or .byt挂号模式 = 2 Then
                mRegistFeeMode = EM_RG_现收
            Else
                mRegistFeeMode = EM_RG_划价
            End If
        End If
         strValue = gobjDatabase.GetPara("预约限制时间", glngSys, 1111, "1|60")
        .int预约限制时间 = Val(Split(strValue & "|", "|")(1))
        .int预约缺省天数 = Val(Split(strValue & "|", "|")(0))
        .bln严格按时段挂号 = Val(gobjDatabase.GetPara("严格按时段挂号", glngSys, 1111, 0)) = 1    '62467
    End With
    '刷卡要求输入密码
    mstrCardPass = gobjDatabase.GetPara(46, glngSys, , "0000000000")
    '收费和挂号共用票据
    mblnSharedInvoice = gobjDatabase.GetPara("挂号共用收费票据", glngSys, 1121) = "1"
    mintSysAppLimit = Val(gobjDatabase.GetPara("挂号允许预约天数", glngSys))
    '本地共用挂号批次ID
    If mblnSharedInvoice Then
        mlng挂号ID = Val(gobjDatabase.GetPara("共用收费票据批次", glngSys, 1121, ""))
    Else
        mlng挂号ID = Val(gobjDatabase.GetPara("共用挂号票据批次", glngSys, mlngModul, ""))
    End If
    If mlng挂号ID > 0 Then
        If Not ExistBill(mlng挂号ID, IIf(mblnSharedInvoice, 1, 4)) Then
            If mblnSharedInvoice Then
                gobjDatabase.SetPara "共用收费票据批次", "0", glngSys, 1121
            Else
                gobjDatabase.SetPara "共用挂号票据批次", "0", glngSys, mlngModul
            End If
            mlng挂号ID = 0
        End If
    End If
    '票号严格控制
    strValue = gobjDatabase.GetPara(24, glngSys, , "00000")
    gblnBill挂号 = (Mid(strValue, IIf(mblnSharedInvoice, 1, 4), 1) = "1")
    If mblnSharedInvoice Then
        '挂号用门诊票据:42703
        mblnStartFactUseType = zlStartFactUseType("1")
    End If
    If mblnAppointment Then
        dtpDate.minDate = gobjDatabase.Currentdate
        dtpDate.Value = gobjDatabase.Currentdate
    End If
    
    '价格等级
    mintPriceGradeStartType = GetPriceGradeStartType()
    If mintPriceGradeStartType > 0 Then
        Call GetPriceGrade(gstrNodeNo, 0, 0, "", , , mstrPriceGrade)
    End If
End Sub

Private Function zlStartFactUseType(ByVal int票种 As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否使用了使用类别的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-05-10 16:11:47
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    On Error GoTo errHandle
    strSql = "Select  1 as 存在 From 票据领用记录 where 票种=[1] and nvl(使用类别,'LXH')<>'LXH' and Rownum=1"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查票据是否启用了使用类别的", int票种)
    
    If rsTemp.EOF Then
        Set rsTemp = Nothing: Exit Function
    End If
    Set rsTemp = Nothing
    zlStartFactUseType = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExistBill(lngID As Long, bytKind As Byte) As Boolean
'功能：判断是否存在指定的票据领用
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
        
    On Error GoTo errH
    
    strSql = "Select ID From 票据领用记录 Where ID=[1] And 票种=[2]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "检查领用ID", lngID, bytKind)
    ExistBill = Not rsTmp.EOF
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Function

Private Function zl_GetInvoiceUserType(ByVal lng病人ID As Long, ByVal lng主页Id As Long, Optional intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取发票的使用类别
    '返回:发票的使用类别
    '编制:刘兴洪
    '日期:2011-04-29 11:03:35
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "Select  Zl_Billclass([1],[2],[3]) as 使用类别 From Dual"
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "获取票据使用类别", lng病人ID, lng主页Id, intInsure)
    zl_GetInvoiceUserType = Nvl(rsTemp!使用类别)
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then Resume
End Function

Private Function RefreshFact(Optional ByRef strFact As String) As Boolean
'参数：blnNew=是否新单保存时调用,这时对于非严格控制的票据是保存当前号
    If mblnStartFactUseType Then
        mstrUseType = zl_GetInvoiceUserType(Val(mrsInfo!病人ID), 0, mintInsure)
    End If
    If gblnBill挂号 Then
        mlng领用ID = CheckUsedBill(IIf(mblnSharedInvoice, 1, 4), IIf(mlng领用ID > 0, mlng领用ID, mlng挂号ID), , mstrUseType)
        If mlng领用ID <= 0 Then
            Select Case mlng领用ID
                Case 0 '操作失败
                Case -1
                    MsgBox "你没有自用和共用的挂号票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
            End Select
            strFact = "": Exit Function
        Else
            '严格：取下一个号码
            strFact = GetNextBill(mlng领用ID)
        End If
    Else
        If mblnSharedInvoice Then
            strFact = gobjDatabase.GetPara("当前收费票据号", glngSys, 1121)
        Else
            strFact = gobjDatabase.GetPara("当前挂号票据号", glngSys, 1111)
        End If
        strFact = IncStr(strFact)
        If mblnSharedInvoice Then
            gobjDatabase.SetPara "当前收费票据号", strFact, glngSys, 1121
        Else
            gobjDatabase.SetPara "当前挂号票据号", strFact, glngSys, 1111
        End If
    End If
    RefreshFact = True
End Function

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, strTemp As String
    Dim lngCardID As Long
    If gobjSquare Is Nothing Then CreateSquareCardObject Me, glngModul
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "姓|姓名|0;医|医保号|0;身|身份证号|0;门|门诊号|0;手|手机号|0", txtPatient)
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTemp)
    mintIDKind = Val(strTemp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
End Function


Private Sub cboDeptFilter_Click()
    If mblnNotClick Then Exit Sub
    
    mblnFilterChange = True
    LoadRegPlans (True)
    mblnFilterChange = False
    If mrsPlan.RecordCount <> 0 Then Call vsfArrange_EnterCell
End Sub

Private Sub cboDeptFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intInputType As Integer, i As Integer
    If KeyCode = 13 Then
        If cboDeptFilter.Text = "" Then
            cboDeptFilter.ListIndex = 0
            Exit Sub
        End If
        If IsNumeric(cboDeptFilter.Text) Then
            intInputType = 0
        ElseIf gobjCommFun.IsCharAlpha(cboDeptFilter.Text) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        For i = 1 To cboDeptFilter.ListCount - 1
            Select Case intInputType
            Case 0, 2
                If cboDeptFilter.List(i) Like "*" & cboDeptFilter.Text & "*" Then
                    cboDeptFilter.ListIndex = i
                    Exit For
                End If
            Case 1  '输入的是全字母
                '问题号:116582,焦博,2017/12/5,通过简码调取科室时，提示'运行时错误'9'：下标越界'
                If UCase(gobjCommFun.zlGetSymbol(cboDeptFilter.List(i))) Like "*" & UCase(cboDeptFilter.Text) & "*" Then
                    cboDeptFilter.ListIndex = i
                    Exit For
                End If
            End Select
        Next i
    End If
End Sub

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboDoctorFilter_Click()

    If mblnNotClick Then Exit Sub
    
    mblnFilterChange = True
    LoadRegPlans (True)
    mblnFilterChange = False
    If mrsPlan.RecordCount <> 0 Then Call vsfArrange_EnterCell
End Sub

Private Sub cboDoctorFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intInputType As Integer, i As Integer
    If KeyCode = 13 Then
        If cboDoctorFilter.Text = "" Then
            cboDoctorFilter.ListIndex = 0
            Exit Sub
        End If
        If IsNumeric(cboDoctorFilter.Text) Then
            intInputType = 0
        ElseIf gobjCommFun.IsCharAlpha(cboDoctorFilter.Text) Then
            intInputType = 1
        Else
            intInputType = 2
        End If
        For i = 1 To cboDoctorFilter.ListCount - 1
            Select Case intInputType
            Case 0, 2
                If cboDoctorFilter.List(i) Like "*" & cboDoctorFilter.Text & "*" Then
                    cboDoctorFilter.ListIndex = i
                    Exit For
                End If
            Case 1  '输入的是全字母
                If UCase(gobjCommFun.zlGetSymbol(cboDoctorFilter.List(i))) Like "*" & UCase(cboDoctorFilter.Text) & "*" Then
                    cboDoctorFilter.ListIndex = i
                    Exit For
                End If
            End Select
        Next i
    End If
End Sub

Private Sub cboPayMode_Click()
    If MCPAR.不收病历费 And cboPayMode.Text = mstrInsure Then
        chkBook.Enabled = False
        chkBook.Value = 0
    Else
        chkBook.Enabled = True
    End If
End Sub

Private Sub cboPayMode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cboRemark_Change()
    cboRemark.Tag = ""
End Sub

Private Sub cboRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboRemark.Tag <> "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(cboRemark.Text) = "" Then gobjCommFun.PressKey vbKeyTab: Exit Sub
    If SelectMemo(Trim(cboRemark.Text)) = False Then
        gobjCommFun.PressKey vbKeyTab: Exit Sub
    End If
End Sub

Private Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String

    If Val(gobjDatabase.GetPara("输入匹配")) = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
End Function

Private Function SelectMemo(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择常用摘要
    '入参:strInput-输入串;为空时,表示全部
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-04 16:06:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSql As String, strWhere As String
    Dim rsInfo As ADODB.Recordset
    Dim vRect As RECT, strKey As String
    strKey = GetMatchingSting(strInput, False)
    If strInput <> "" Then
        If gobjCommFun.IsCharChinese(cboRemark.Text) Then
             strWhere = " And  名称 like [1] "
        ElseIf gobjCommFun.IsNumOrChar(cboRemark.Text) Then
             strWhere = " And (简码 like upper([1]) or 编码 like upper([1]))"
        End If
    End If
    
    strSql = "" & _
     "   Select RowNum AS ID,编码,名称,简码  " & _
     "   From 常用挂号摘要 " & _
     "   Where 1=1 " & strWhere & _
     "   Order by 缺省标志"
     vRect = GetControlRect(cboRemark.hWnd)
     On Error GoTo Hd
     Set rsInfo = gobjDatabase.ShowSQLSelect(Me, strSql, 0, "常用挂号摘要", False, _
                    "", "", False, False, True, vRect.Left, vRect.Top, cboRemark.Height, blnCancel, True, False, strKey)
     If blnCancel Then Exit Function
     If rsInfo Is Nothing Then
        If strInput = "" Then
            MsgBox "没有设置常用挂号摘要,请在字典管理中设置", vbOKOnly + vbInformation, gstrSysName
        End If
        gobjCommFun.PressKey vbKeyTab: Exit Function
     End If
     gobjControl.CboSetText Me.cboRemark, Nvl(rsInfo!名称)
     cboRemark.Tag = Nvl(rsInfo!名称)
     gobjCommFun.PressKey vbKeyTab
     SelectMemo = True
     Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then Resume
    gobjComlib.SaveErrLog
End Function

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub chkBook_Click()
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")) = "" Then Exit Sub
    Call LoadFeeItem(Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1, mstrPriceGrade)
End Sub

Private Sub chkBook_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    If txtPatient.Text <> "" Then
        If MsgBox("是否清空当前病人信息？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ClearPatient
        End If
        Exit Sub
    End If
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdPrice_Click()
    Dim bytRegistFeeMode As EM_REGISTFEE_MODE
    bytRegistFeeMode = mRegistFeeMode
    
    mRegistFeeMode = EM_RG_划价
    If SaveData = False Then mRegistFeeMode = bytRegistFeeMode: Exit Sub
    mRegistFeeMode = bytRegistFeeMode
    
    mblnUpdateAge = False
    Call ReloadPage
End Sub

Private Function CheckBrushCard(ByVal dblMoney As Double, ByVal lng医疗卡类别ID As Long, ByVal bln消费卡 As Boolean, _
                                ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMoney As ADODB.Recordset, str年龄 As String, lng病人ID As Long
    On Error GoTo errHandle
    '68991
    If mRegistFeeMode <> EM_RG_现收 Then CheckBrushCard = True: Exit Function
    If dblMoney = 0 Then
        CheckBrushCard = True: Exit Function
    End If
    If Not (cboPayMode.Visible And cboPayMode.Enabled) Then
        CheckBrushCard = True: Exit Function
    End If
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then
        CheckBrushCard = True: Exit Function
    End If
    If lng医疗卡类别ID = 0 Then
        MsgBox cboPayMode.Text & "异常,请检查!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If gobjSquare.objSquareCard Is Nothing Then
        MsgBox "使用" & cboPayMode.Text & "支付必须先初始化接口部件！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call zlGetClassMoney(rsMoney, rsItems, rsIncomes)
    
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean = False, _
    Optional ByVal bln余额不足禁止 As Boolean = True, _
    Optional ByRef varSquareBalance As Variant, _
    Optional ByVal bln转预交 As Boolean = False, _
    Optional ByVal blnAllPay As Boolean = False, _
    Optional ByVal strXmlIn As String = "", _
    Optional ByVal str费用来源 As String, _
    Optional ByVal lng病人ID As Long) As Boolean
    str年龄 = Trim(txtAge.Text)
    If Not mrsInfo Is Nothing Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
   If gobjSquare.objSquareCard.zlBrushCard(Me, glngModul, rsMoney, lng医疗卡类别ID, bln消费卡, _
    txtPatient.Text, NeedName(txtGender.Text), str年龄, dblMoney, mstrCardNO, mstrPassWord, _
    False, True, False, True, Nothing, False, True, "", "1", lng病人ID) = False Then Exit Function
    
    If gobjSquare.objSquareCard.zlPaymentCheck(Me, glngModul, lng医疗卡类别ID, _
        bln消费卡, mstrCardNO, dblMoney, "", "") = False Then Exit Function

    CheckBrushCard = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGetClassMoney(ByRef rsMoney As ADODB.Recordset, ByVal rsItems As ADODB.Recordset, ByVal rsIncomes As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存时,初始化支付类别(收费类别,实收金额)
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-06-10 17:52:18
    '问题:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql  As String
    
    Err = 0: On Error GoTo Errhand:
    
    Set rsMoney = New ADODB.Recordset
    With rsMoney
        If .State = adStateOpen Then .Close
        .Fields.Append "收费类别", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "金额", adDouble, 18, adFldIsNullable
        .CursorLocation = adUseClient
        .Open , , adOpenStatic, adLockOptimistic

        rsItems.Filter = 0
        If rsItems.RecordCount <> 0 Then rsItems.MoveFirst
        Do While Not rsItems.EOF
            rsIncomes.Filter = "项目ID=" & rsItems!项目ID
            rsMoney.Filter = "收费类别='" & Nvl(rsItems!类别, "无") & "'"
            If rsMoney.EOF Then
                .AddNew
            Else
                rsMoney.Filter = 0
            End If
            !收费类别 = Nvl(rsItems!类别, "无")
            Do While Not rsIncomes.EOF
                !金额 = Val(Nvl(!金额)) + Val(Nvl(rsIncomes!实收))
                rsIncomes.MoveNext
            Loop
            .Update
            rsItems.MoveNext
        Loop
    End With
    rsMoney.Filter = 0
    zlGetClassMoney = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If SaveData = False Then
        If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then Call CancelRegNo
        Exit Sub
    End If
    mblnUpdateAge = False
    Call ReloadPage
End Sub
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存挂号数据
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-01 16:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSlipPrint As Boolean, blnInvoicePrint As Boolean, int价格父号 As Integer, blnBalance As Boolean
    Dim k As Integer, i As Integer, j As Integer, strNO As String, strFactNO As String
    Dim cllPro As New Collection, strSql As String, str登记时间 As String, str发生时间 As String
    Dim cur预交 As Currency, cur个帐 As Currency, cur现金 As Currency, str划价NO As String
    Dim lngSN As Long, lng挂号科室ID As Long, lng结帐ID As Long, byt复诊 As Byte
    Dim lng医疗卡类别ID As Long, bln消费卡 As Boolean, blnNoDoc As Boolean, strBalanceStyle As String
    Dim rsItems As ADODB.Recordset, rsIncomes As ADODB.Recordset, cllProAfter As New Collection
    Dim blnTrans As Boolean, blnNotCommit As Boolean, strAdvance As String, blnAppointPrint As Boolean
    Dim lng医生ID As Long, blnOneCard As Boolean, rsTmp As ADODB.Recordset, str付款方式 As String
    Dim cllCardPro As Collection, cllTheeSwap As Collection, strNotValiedNos As String
    Dim dat登记时间 As Date
    Dim bytMode As Byte, rsCheck As ADODB.Recordset, dat预约时间 As Date
    Dim strResult As String, bln专家号 As Boolean
    Dim dblTotal As Double
    
    If CheckValied = False Then Exit Function
    
    If Not mrsInfo Is Nothing Then
        strSql = "Select Zl_Fun_病人挂号记录_Check([1],[2],[3],Null,[4],[5]) As 检查结果 From Dual"
        If mblnAppointment Then
            bytMode = 1
            dat预约时间 = CDate(Format(dtpDate.Value, "yyyy-mm-dd"))
        Else
            bytMode = 0
            dat预约时间 = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd"))
        End If
    
        bln专家号 = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生")) <> ""
        Set rsCheck = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, bytMode, Val(Nvl(mrsInfo!病人ID)), Trim(txtArrangeNO.Text), dat预约时间, IIf(bln专家号, 1, 0))
        If Not rsCheck.EOF Then
            strResult = Nvl(rsCheck!检查结果)
            If Val(Mid(strResult, 1, 1)) <> 0 Then
                MsgBox Mid(strResult, 3), vbInformation, gstrSysName
                Exit Function
            End If
        Else
            MsgBox "有效性检查失败,无法继续！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    strSql = "Select 编号,名称,医院编码,结算方式 From 一卡通目录 Where 启用 = 1 And 结算方式 = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, cboPayMode.Text)
    blnOneCard = rsTmp.RecordCount <> 0
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        blnSlipPrint = False
    Else
        Select Case Val(mty_Para.int挂号凭条打印)
            Case 0    '不打印
                blnSlipPrint = False
            Case 1    '自动打印
                If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                    blnSlipPrint = True
                Else
                    blnSlipPrint = False
                    MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2    '选择打印
                If MsgBox("要打印挂号凭条吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                    If InStr(gstrPrivs, ";病人挂号凭条;") > 0 Then
                        blnSlipPrint = True
                    Else
                        blnSlipPrint = False
                        MsgBox "你没有挂号凭条打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Else
                    blnSlipPrint = False
                End If
        End Select
    End If
    
    If mRegistFeeMode = EM_RG_划价 Or mRegistFeeMode = EM_RG_记帐 Or (mblnAppointment And mty_Para.bln预约时收款 = False) Then
        blnInvoicePrint = False
    Else
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
            Select Case Val(mty_Para.int挂号发票打印)
                Case 0    '不打印
                    blnInvoicePrint = False
                Case 1    '自动打印
                    If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                        blnInvoicePrint = True
                    Else
                        blnInvoicePrint = False
                        MsgBox "你没有挂号发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                    End If
                Case 2    '选择打印
                    If MsgBox("要打印挂号发票吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        If InStr(gstrPrivs, ";挂号发票打印;") > 0 Then
                            blnInvoicePrint = True
                        Else
                            blnInvoicePrint = False
                            MsgBox "你没有挂号发票打印的权限，请联系管理员！", vbInformation, gstrSysName
                        End If
                    Else
                        blnInvoicePrint = False
                    End If
            End Select
        End If
    End If
    
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        Select Case Val(mty_Para.int预约挂号打印)
            Case 0
                blnAppointPrint = False
            Case 1
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    blnAppointPrint = True
                Else
                    blnAppointPrint = False
                    MsgBox "你没有预约挂号单打印的权限，请联系管理员！", vbInformation, gstrSysName
                End If
            Case 2
                If InStr(gstrPrivs, ";预约挂号单;") > 0 Then
                    If MsgBox("要打印预约挂号单吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        blnAppointPrint = True
                    Else
                        blnAppointPrint = False
                    End If
                Else
                    MsgBox "你没有预约挂号单打印的权限，请联系管理员！", vbInformation, gstrSysName
                    blnAppointPrint = False
                End If
        End Select
    Else
        blnAppointPrint = False
    End If
    
    If blnInvoicePrint Or (mintInsure <> 0 And MCPAR.医保接口打印票据) Then
        If RefreshFact(strFactNO) = False Then Exit Function
    End If
    
    If mblnAppointment Then
        If mRegistFeeMode = EM_RG_记帐 And mty_Para.bln预约时收款 Then
            MsgBox "不支持先诊疗后结算病人的预约收款挂号！", vbInformation, gstrSysName
            Exit Function
        End If
        If mty_Para.bln预约时收款 Then
            If Not mRegistFeeMode = EM_RG_划价 Then
                If cboPayMode.Text = "预交金" Then
                    cur预交 = Val(lblTotal.Caption)
                Else
                    If cboPayMode.Text = mstrInsure Then
                        cur个帐 = Val(lblTotal.Caption)
                    Else
                        blnBalance = True
                        cur现金 = Val(lblTotal.Caption)
                    End If
                End If
            End If
        Else
            blnBalance = False
        End If
    Else
        If Not mRegistFeeMode = EM_RG_划价 Then
            If cboPayMode.Text = "预交金" Then
                cur预交 = Val(lblTotal.Caption)
            Else
                If cboPayMode.Text = "个人帐户" Then
                    cur个帐 = Val(lblTotal.Caption)
                Else
                    blnBalance = True
                    cur现金 = Val(lblTotal.Caption)
                End If
            End If
        End If
    End If
    
    If Not mRegistFeeMode = EM_RG_划价 And cboPayMode.Visible Then
        If cboPayMode.Text = "预交金" Then
            cur预交 = Val(lblTotal.Caption)
        Else
            If cboPayMode.Text = mstrInsure Then
                cur个帐 = Val(lblTotal.Caption)
            Else
                blnBalance = True
                cur现金 = Val(lblTotal.Caption)
            End If
        End If
    End If
    
    If Val(cur预交) <> 0 Then
        If Not gobjDatabase.PatiIdentify(Me, glngSys, Nvl(mrsInfo!病人ID), Val(cur预交), mlngModul, 1, , _
                            IIf(-1 * mty_Para.dbl预存款消费验卡 >= Val(cur预交), False, True), True, mstr家属IDs, (mty_Para.dbl预存款消费验卡 <> 0), (mty_Para.dbl预存款消费验卡 = 2)) Then Exit Function
    End If
    
    strSql = "Select 编码 From 医疗付款方式 Where 名称 = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, Nvl(mrsInfo!医疗付款方式))
    If rsTmp.RecordCount <> 0 Then
        str付款方式 = Nvl(rsTmp!编码)
    Else
        strSql = "Select 编码 From 医疗付款方式 Where 缺省标志 = 1"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName)
        If rsTmp.RecordCount <> 0 Then
            str付款方式 = Nvl(rsTmp!编码)
        End If
    End If
    
    ReadRegistPrice Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1, _
        False, txtFeeType.Text, rsItems, rsIncomes, , , , , , mstrPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
        
    If mblnAppointment = False Or (mblnAppointment = True And mty_Para.bln预约时收款) Then
        If zlIsAllowPatiChargeFeeMode(ZVal(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!结算模式))) = False Then Exit Function
    End If
    
    If blnBalance Then
        For i = 1 To mcolCardPayMode.Count
            If cboPayMode.Text = mcolCardPayMode.Item(i)(1) Then
                lng医疗卡类别ID = mcolCardPayMode.Item(i)(3)
                bln消费卡 = Val(mcolCardPayMode.Item(i)(5)) = 1
                strBalanceStyle = mcolCardPayMode.Item(i)(6)
            End If
        Next i
        If CheckBrushCard(Val(cur现金), lng医疗卡类别ID, bln消费卡, rsItems, rsIncomes) = False Then Exit Function
    End If
    
    str登记时间 = "To_Date('" & gobjDatabase.Currentdate & "','yyyy-mm-dd hh24:mi:ss')"
    dat登记时间 = gobjDatabase.Currentdate
    If mblnAppointment Then
        str发生时间 = "To_Date('" & Format(dtpDate.Value, "YYYY-MM-DD") & " " & Format(dtpTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
    Else
        If mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Then
            str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "YYYY-MM-DD") & " " & Format(dtpTime.Value, "hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
        Else
            str发生时间 = "To_Date('" & Format(gobjDatabase.Currentdate, "YYYY-MM-DD hh:mm:ss") & "','yyyy-mm-dd hh24:mi:ss')"
        End If
    End If
    
    lng挂号科室ID = Val(vsfArrange.RowData(vsfArrange.Row))
    If mRegistFeeMode = EM_RG_现收 Then
        lng结帐ID = gobjDatabase.GetNextId("病人结帐记录")
    End If
    byt复诊 = IIf(Check复诊(Val(mrsInfo!病人ID), lng挂号科室ID), 1, 0)
    mstrPreNO = txtArrangeNO.Text
    
    
    
    '票据处理
    If vsfDetailTime.Visible Then
        If mViewMode = v_专家号分时段 Then
            lngSN = Val(Get时段(vsfDetailTime.Row, vsfDetailTime.Col))
        End If
        If mViewMode = v_专家号 Then
            lngSN = Val(vsfDetailTime.TextMatrix(vsfDetailTime.Row, vsfDetailTime.Col))
        End If
    Else
        lngSN = 0
    End If
    '137272:李南春,2019/2/20,专家号也需求提前取出序号，然后进行锁号
    If mViewMode = v_专家号 Or mViewMode = v_专家号分时段 Then
        If ReserveRegNo(txtArrangeNO.Text, True, mViewMode = v_专家号分时段, str发生时间, lngSN, "分诊挂号") = False Then Exit Function
    End If
    
    strNO = gobjDatabase.GetNextNo(12)
    
    rsItems.Filter = ""
    If cboDoctor.ListIndex = -1 Then
        lng医生ID = 0
    Else
        lng医生ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
    End If
    
    dblTotal = 0
    If mRegistFeeMode = EM_RG_划价 Then
        dblTotal = GetRegistMoney(True, False)
        '挂号费存为零且保存为划价单，才产生划价NO
       If dblTotal <> 0 Then str划价NO = gobjDatabase.GetNextNo(13)
    End If
    
    k = 1: rsItems.MoveFirst
    For i = 1 To rsItems.RecordCount
        int价格父号 = k
        rsIncomes.Filter = "项目ID=" & rsItems!项目ID
        For j = 1 To rsIncomes.RecordCount
            strSql = _
            "zl_病人挂号记录_INSERT(" & ZVal(Nvl(mrsInfo!病人ID)) & "," & IIf(txtClinic.Text = "", "NULL", txtClinic.Text) & ",'" & txtPatient.Text & "','" & NeedName(txtGender.Text) & "'," & _
                     "'" & txtAge.Text & "','" & str付款方式 & "','" & txtFeeType.Text & "','" & strNO & "'," & _
                     "'" & IIf(blnInvoicePrint = False, "", "") & "'," & k & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                     "'" & rsItems!类别 & "'," & rsItems!项目ID & "," & rsItems!数次 & "," & rsIncomes!单价 & "," & _
                     rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "','" & IIf(blnBalance, IIf(strBalanceStyle = "", cboPayMode.Text, strBalanceStyle), "") & "'," & _
                     IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!应收) & "," & IIf(mRegistFeeMode = EM_RG_划价, 0, rsIncomes!实收) & "," & _
                     lng挂号科室ID & "," & UserInfo.部门ID & "," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & _
                     str发生时间 & "," & str登记时间 & "," & _
                     "'" & NeedName(cboDoctor.Text) & "'," & ZVal(lng医生ID) & "," & IIf(rsItems!性质 = 3, 1, IIf(rsItems!性质 = 4, 2, 0)) & "," & IIf(lbl急.Visible, 1, 0) & "," & _
                     "'" & txtArrangeNO.Text & "','" & cboRoom.Text & "'," & ZVal(lng结帐ID) & "," & IIf(blnInvoicePrint = False, "NULL", ZVal(mlng领用ID)) & "," & _
                     ZVal(IIf(k = 1, cur预交, 0)) & "," & ZVal(IIf(k = 1, cur现金, 0)) & "," & _
                     ZVal(IIf(k = 1, cur个帐, 0)) & "," & ZVal(Nvl(rsItems!保险大类ID, 0)) & "," & _
                     ZVal(Nvl(rsItems!保险项目否, 0)) & "," & ZVal(Nvl(rsIncomes!统筹金额, 0)) & "," & _
                     "'" & Trim(cboRemark.Text) & "'," & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 0, 1), 0) & "," & IIf(mty_Para.bln共用收费票据, 1, 0) & ",'" & rsItems!保险编码 & "'," & byt复诊 & "," & ZVal(lngSN) & ",Null," & _
                     IIf(mblnAppointment, 1, 0) & ",'" & IIf(cboAppointStyle.Visible, cboAppointStyle.Text, "") & "'," & _
                     IIf(mty_Para.bln挂号生成队列, 1, 0) & ","
            '卡类别id_In   病人预交记录.卡类别id%Type := Null,
            strSql = strSql & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡 = False, lng医疗卡类别ID, "NULL") & ","
            '结算卡序号_In 病人预交记录.结算卡序号%Type := Null,
            strSql = strSql & "" & IIf(lng医疗卡类别ID <> 0 And bln消费卡, lng医疗卡类别ID, "NULL") & ","
            '卡号_In       病人预交记录.卡号%Type := Null,
            strSql = strSql & "'" & mstrCardNO & "',"
            '交易流水号_In 病人预交记录.交易流水号%Type := Null,
            strSql = strSql & " NULL,"
            '交易说明_In   病人预交记录.交易说明%Type := Null,
            strSql = strSql & " NULL,"
            '合作单位_In   病人预交记录.合作单位%Type := Null
            strSql = strSql & " NULL,"
            '  操作类型_In   Number:=0
            strSql = strSql & "0" & ","
            '  险类_IN       病人挂号记录.险类%type:=null,
            strSql = strSql & IIf(mintInsure = 0, "NULL", mintInsure) & ","
            '  结算模式_IN   NUMBER :=0,
            strSql = strSql & IIf(mPatiChargeMode = EM_先诊疗后结算, 1, 0) & ","
            '  记帐费用_IN Number:=0
            strSql = strSql & IIf(mRegistFeeMode = EM_RG_记帐, 1, 0) & ","
            '  退号重用_IN Number:=1
            strSql = strSql & IIf(mty_Para.bln退号重用, 1, 0) & ","
            '  冲预交病人ids_In Varchar2 := Null
            strSql = strSql & "'" & Nvl(mrsInfo!病人ID) & "," & mstr家属IDs & "',"
            '  修正病人费别_In Number := 0
            strSql = strSql & "" & IIf(mblnChangeFeeType, 1, 0) & ","
            '  修正病人年龄_In Number := 0
            strSql = strSql & "" & IIf(mblnUpdateAge, 1, 0) & ","
            '  收费单_In        病人挂号记录.收费单%Type := Null
            strSql = strSql & "'" & str划价NO & "')"
            
            Call zlAddArray(cllPro, strSql)
            '问题:31187:将挂号汇总单独出来
            If txtArrangeNO.Text <> "" And k = 1 Then
                If Nvl(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生"))) = "" Then blnNoDoc = True
                strSql = "zl_病人挂号汇总_Update("
                '  医生姓名_In   挂号安排.医生姓名%Type,
                strSql = strSql & IIf(blnNoDoc, "Null,", "'" & NeedName(cboDoctor.Text) & "',")
                '  医生id_In     挂号安排.医生id%Type,
                strSql = strSql & "" & IIf(blnNoDoc, "0,", ZVal(lng医生ID) & ",")
                '  收费细目id_In 门诊费用记录.收费细目id%Type,
                strSql = strSql & "" & Val(Nvl(rsItems!项目ID)) & ","
                '  执行部门id_In 门诊费用记录.执行部门id%Type,
                strSql = strSql & "" & IIf(Val(Nvl(rsItems!执行科室ID)) = 0, lng挂号科室ID, Val(Nvl(rsItems!执行科室ID))) & ","
                '  发生时间_In   门诊费用记录.发生时间%Type,
                strSql = strSql & "" & str发生时间 & ","
                '  预约标志_In   Number := 0  --是否为预约接收:0-非预约挂号; 1-预约挂号,2-预约接收,3-收费预约
                strSql = strSql & IIf(mblnAppointment, IIf(mty_Para.bln预约时收款, 3, 1), 0) & ","
                '  号码_In       挂号安排.号码%Type := Null
                strSql = strSql & "'" & txtArrangeNO.Text & "')"
                Call zlAddArray(cllProAfter, strSql)
            End If
            If mRegistFeeMode = EM_RG_划价 And dblTotal <> 0 Then
                strSql = _
                "zl_门诊划价记录_Insert('" & str划价NO & "'," & k & "," & ZVal(Nvl(mrsInfo!病人ID)) & ",NULL," & _
                         IIf(txtClinic.Text = "", "NULL", txtClinic.Text) & ",'" & str付款方式 & "'," & _
                         "'" & txtPatient.Text & "','" & txtGender.Text & "','" & txtAge.Text & "'," & _
                         "'" & txtFeeType.Text & "',NULL," & lng挂号科室ID & "," & _
                         IIf(lng挂号科室ID <> 0, lng挂号科室ID, UserInfo.部门ID) & ",'" & UserInfo.姓名 & "'," & IIf(rsItems!性质 = 2, 1, "NULL") & "," & _
                         rsItems!项目ID & ",'" & rsItems!类别 & "','" & rsItems!计算单位 & "'," & _
                         "NULL,1," & rsItems!数次 & ",NULL," & IIf(rsItems!执行科室ID = 0, lng挂号科室ID, rsItems!执行科室ID) & "," & IIf(int价格父号 = k, "NULL", int价格父号) & "," & _
                         rsIncomes!收入项目ID & ",'" & rsIncomes!收据费目 & "'," & rsIncomes!单价 & "," & _
                         rsIncomes!应收 & "," & rsIncomes!实收 & "," & str发生时间 & "," & str登记时间 & ",NULL,'" & UserInfo.姓名 & "','挂号:" & strNO & "')"
                Call zlAddArray(cllPro, strSql)
            End If
            k = k + 1
            rsIncomes.MoveNext
            Next j
        rsItems.MoveNext
    Next i
    
    
    Err = 0: On Error GoTo ErrFirt:
    
    If cllPro.Count > 0 Then
        Err = 0: On Error GoTo ErrFirt:
        zlExecuteProcedureArrAy cllPro, Me.Caption, True, False

        Err = 0: On Error GoTo errH:
        blnTrans = True
        If blnOneCard And lng医疗卡类别ID <> 0 And mRegistFeeMode = EM_RG_现收 And cur现金 <> 0 Then
            If Not mobjICCard.PaymentSwap(Val(cur现金), Val(cur现金), Val(lng医疗卡类别ID), 0, mstrCardNO, "", lng结帐ID, Nvl(mrsInfo!病人ID)) Then
                gcnOracle.RollbackTrans
                MsgBox "一卡通结算挂号费失败", vbInformation, gstrSysName
                Exit Function
            Else
                strSql = "zl_一卡通结算_Update(" & lng结帐ID & ",'" & cboPayMode.Text & "','" & mstrCardNO & "','" & lng医疗卡类别ID & "','" & "" & "'," & cur现金 & ")"
                Call gobjDatabase.ExecuteProcedure(strSql, Me.Caption)
            End If
        End If

        '医保改动
        blnNotCommit = False
        If mintInsure <> 0 And mstrYBPati <> "" And cur个帐 <> 0 Then
            '68991:strAdvance:结算模式(0或1)|挂号费收取方式(0或1) |挂号单号
            strAdvance = ""
            If mRegistFeeMode = EM_RG_记帐 Or mPatiChargeMode = EM_先诊疗后结算 Then
                strAdvance = IIf(mPatiChargeMode = EM_先诊疗后结算, "1", "0")
                strAdvance = strAdvance & "|" & IIf(mRegistFeeMode = EM_RG_记帐, "1", "0")
                strAdvance = strAdvance & "|" & strNO
            End If
            If Not gclsInsure.RegistSwap(lng结帐ID, cur个帐, mintInsure, strAdvance) Then
                gcnOracle.RollbackTrans: cmdOK.Enabled = True: Exit Function
            End If
            blnNotCommit = True
        End If
        '问题:31187 调用医保成功后,最后作一些数据更新:内部过程中已有提交语句,所以不用再写
        zlExecuteProcedureArrAy cllProAfter, Me.Caption, False, False
        Set cllCardPro = New Collection: Set cllTheeSwap = New Collection
        If mRegistFeeMode = EM_RG_现收 And Not blnOneCard And Not mPatiChargeMode = EM_先诊疗后结算 And cur现金 <> 0 Then
            If zlInterfacePrayMoney(lng结帐ID, cllCardPro, cllTheeSwap, Val(cur现金), lng医疗卡类别ID, bln消费卡) = False Then
                gcnOracle.RollbackTrans: If cmdOK.Enabled = False Then cmdOK.Enabled = True
                Exit Function
            End If
            '修正三方交易
            zlExecuteProcedureArrAy cllCardPro, Me.Caption, False, False
        End If
        
        Err = 0: On Error GoTo OthersCommit:
        zlExecuteProcedureArrAy cllTheeSwap, Me.Caption, False, False
OthersCommit:
        gcnOracle.CommitTrans

        If mintInsure > 0 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, True, mintInsure)
        
        blnTrans = False
        On Error GoTo 0
    End If
    '打印单据
    If blnInvoicePrint Then
RePrint:
        If Not (mintInsure <> 0 And MCPAR.医保接口打印票据) And mRegistFeeMode = EM_RG_现收 Then
            Dim blnEnterPrint As Boolean
            blnEnterPrint = True
            Load frmPrint
            Call frmPrint.ReportPrint(1, strNO, "", mlng领用ID, mlng挂号ID, strFactNO, dat登记时间, , , , mintInsure <> 0 And MCPAR.医保接口打印票据, False, mstrUseType)
            If gblnBill挂号 Then
                If zlIsNotSucceedPrintBill(4, strNO, strNotValiedNos) = True Then
                    If MsgBox("挂号单号为[" & strNotValiedNos & "]票据打印未成功,是否重新进行票据打印!", vbYesNo + vbDefaultButton1 + vbQuestion, gstrSysName) = vbYes Then GoTo RePrint:
                End If
            End If
        End If
    End If
    
    If blnSlipPrint Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me, "NO=" & strNO, 2)
    End If
    
    If (blnSlipPrint Or blnInvoicePrint) And Not blnEnterPrint Then
        '记录打印的凭条
        gstrSQL = "Zl_凭条打印记录_Update(4,'" & strNO & "',1,'" & UserInfo.姓名 & "')"
        gobjDatabase.ExecuteProcedure gstrSQL, ""
    End If
    
    SaveData = True
    
    Exit Function
ErrFirt:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
errH:
    If mintInsure > 0 And blnNotCommit Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_RegistSwap, False, mintInsure)
    If blnTrans Then gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
    Exit Function
ErrGo:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Sub ReloadPage()
    Call LoadRegPlans(False)
    Call ClearPatient
    Call ClearRegInfo
End Sub

Private Sub ClearRegInfo()
    mblnChangeByCode = True
    txtArrangeNO.Text = ""
    mblnChangeByCode = False
    txtDept.Text = ""
    cboDoctor.Clear
    cboRoom.Clear
    cboRemark.Text = ""
    chkBook.Value = IIf(mty_Para.bln默认购买病历, 1, 0)
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
    vsfDetailTime.Visible = False
    lbl急.Visible = False
    txtPatient.SetFocus
End Sub

Private Function zlIsNotSucceedPrintBill(ByVal bytType As Byte, ByVal strNos As String, ByRef strOutValidNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否已经正常打印
    '入参:bytType-1-收费,2-预交,3-结帐,4-挂号,5-就诊卡
    '       strNos-本次打印票据的单据,用逗号分离
    '出参:strOutValidNos-打印失败的单据号
    '返回:存在不存功票据的打印,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-16 18:06:01
    '问题:44322,44326,44332,44330
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempNos As String, rsTemp As ADODB.Recordset
    Dim strSql As String, strBillNos As String
    Dim bytBill As Byte
    On Error GoTo errHandle
    strBillNos = Replace(Replace(strNos, "'", ""), " ", "")
    strSql = "" & _
        "Select  /*+ rule */ distinct  B.NO " & _
        " From 票据使用明细 A,票据打印内容 B,Table( f_Str2list([2])) J" & _
        " Where A.打印ID =b.ID And B.数据性质=[1] And B.No=J.Column_value "
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, "检查票据是否打印", bytType, strBillNos)
    
    strTempNos = ""
    With rsTemp
        Do While Not .EOF
            If InStr(1, "," & strBillNos & ",", "," & !NO & ",") = 0 Then
                strTempNos = strTempNos & "," & !NO
            End If
            .MoveNext
        Loop
        If .RecordCount = 0 Then strTempNos = "," & strBillNos
    End With
    If strTempNos <> "" Then strTempNos = Mid(strTempNos, 2)
    rsTemp.Close: Set rsTemp = Nothing
    strOutValidNos = strTempNos
    zlIsNotSucceedPrintBill = strTempNos <> ""
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckValied() As Boolean
    Dim i As Integer
    '保存前检查
    If mrsInfo Is Nothing Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        txtPatient.SetFocus
        Exit Function
    End If
    If mrsInfo.RecordCount = 0 Then
        MsgBox "无法确定病人信息,请先选择一个病人！", vbInformation, gstrSysName
        txtPatient.SetFocus
        Exit Function
    End If
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")) = "" Or txtArrangeNO.Text = "" Then
        txtArrangeNO.SetFocus
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")) <> txtArrangeNO.Text Then
        txtArrangeNO.SetFocus
        MsgBox "无法确定号别信息,请先选择一个号别！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheck限约或限号数(txtArrangeNO.Text) = False Then Exit Function
    If vsfDetailTime.Visible Then
        If vsfDetailTime.Row > vsfDetailTime.Rows - 1 Or vsfDetailTime.Col > vsfDetailTime.Cols - 1 Then
            MsgBox "选择了无效序号，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
        If vsfDetailTime.Cell(flexcpForeColor, vsfDetailTime.Row, vsfDetailTime.Col) <> vbBlack Or vsfDetailTime.Cell(flexcpBackColor, vsfDetailTime.Row, vsfDetailTime.Col) = -2147483633 Then
            MsgBox "选择了无效序号，请检查！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mRegistFeeMode <> EM_RG_划价 Then
        If cboPayMode.Text = "" And cboPayMode.Visible Then
            MsgBox "没有确定可用的结算方式,不能完成挂号!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If mblnAppointment And mty_Para.bln预约时收款 = False Then
        If IsNull(mrsPlan!排班) Then
            MsgBox "预约不收款模式下,不能挂不当班的号别!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If Nvl(mrsInfo!姓名) <> txtPatient.Text Then
        If MsgBox("当前病人姓名已经发生变化,是否重新读取病人信息?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            Call GetPatient(IDKind.GetCurCard, txtPatient.Text, False)
            Exit Function
        Else
            txtPatient.Text = Nvl(mrsInfo!姓名)
        End If
    End If
    
    If InStr(gstrPrivs, ";挂号费别打折;") = 0 Then
        For i = 1 To vsfMoney.Rows - 1
            If Val(vsfMoney.TextMatrix(i, 2)) <> Val(vsfMoney.TextMatrix(i, 1)) Then
                MsgBox "你没有权限给病人使用打折费别,不能完成挂号", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    '服务对象检查
    If Not mrsItems Is Nothing Then
        mrsItems.MoveFirst
        Do While Not mrsItems.EOF
            If Val(Nvl(mrsItems!项目ID)) <> 0 Then
                If CheckServeRange(0, Val(Nvl(mrsItems!项目ID))) = False Then Exit Function
            End If
            mrsItems.MoveNext
        Loop
        mrsItems.MoveFirst
    End If
    
    CheckValied = True
End Function

Private Function CheckServeRange(intType As Integer, lng收费细目ID As Long, Optional intRow As Integer = 0) As Boolean
'功能:检查收费项目的服务对象,intType:0-门诊调用;1-住院调用
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = "Select 名称,Nvl(服务对象,0) As 服务对象 From 收费项目目录 Where ID = [1]"
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, "CheckServeRange", lng收费细目ID)
    If rsTmp.EOF Then
        MsgBox "不能确定" & IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目的服务对象,请检查项目是否正确录入!"
        Exit Function
    Else
        Select Case intType
        Case 0
            If Val(rsTmp!服务对象) = 2 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于门诊,请检查!"
                Exit Function
            End If
        Case 1
            If Val(rsTmp!服务对象) = 1 Or Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于住院,请检查!"
                Exit Function
            End If
        Case Else
            If Val(rsTmp!服务对象) = 0 Then
                MsgBox IIf(intRow = 0, "", "第" & intRow & "行") & "收费项目[" & rsTmp!名称 & "]不适用于病人,请检查!"
                Exit Function
            End If
        End Select
    End If
    CheckServeRange = True
End Function

Private Sub SetControl()
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strTemp As String, i As Integer
    If mblnAppointment Then
        Me.Caption = "门诊预约"
        If mty_Para.bln预约时收款 Then
            fraPay.Visible = True
        Else
            fraPay.Visible = False
        End If
        cboAppointStyle.Clear
        strSql = "Select 名称,缺省标志 From 预约方式"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
        Do While Not rsTmp.EOF
            cboAppointStyle.AddItem Nvl(rsTmp!名称)
            If Val(Nvl(rsTmp!缺省标志)) = 1 Then cboAppointStyle.ListIndex = cboAppointStyle.NewIndex
            rsTmp.MoveNext
        Loop
        strTemp = gobjDatabase.GetPara("缺省预约方式", glngSys, 9000, "")
        For i = 0 To cboAppointStyle.ListCount - 1
            If cboAppointStyle.List(i) = strTemp Then
                cboAppointStyle.ListIndex = i
            End If
        Next i
    Else
        Me.Caption = "门诊挂号"
        lblDeptFilter.Left = lblDate.Left
        cboDeptFilter.Left = dtpDate.Left
        cboDeptFilter.Width = 2055
        lblDoctorFilter.Left = 2805
        cboDoctorFilter.Left = 3315
        cboDoctorFilter.Width = 2055
        lblDate.Visible = False
        dtpDate.Visible = False
        lblAppointStyle.Visible = False
        cboAppointStyle.Visible = False
        lblRemark.Left = lblRoom.Left
        cboRemark.Left = 570
        cboRemark.Width = 4110
        If mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2 Then
            fraPay.Visible = True
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
        Else
            fraPay.Visible = False
            cmdPrice.Visible = False
        End If
        
    End If
End Sub

Private Sub dtpDate_Change()

    mblnNotClick = True
    cboDeptFilter.Text = ""
    cboDoctorFilter.Text = ""
    cboDeptFilter.ListIndex = -1
    cboDoctorFilter.ListIndex = -1
    mblnNotClick = False
    
    Call LoadRegPlans(False)
End Sub

Private Sub Form_Activate()
    If mblnInit Then
        mblnInit = False
    End If
End Sub


Private Function zlInterfacePrayMoney(ByVal lng挂号结帐ID As Long, ByRef cllPro As Collection, _
    ByRef cllThreeSwap As Collection, dblMoney As Double, lng医疗卡类别ID As Long, bln消费卡 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:接口支付金额
    '出参:cllPro-修改三方交易数据
    '        cll三方交易-增加三交方易数据
    '返回:支付成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-17 13:34:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    
    If lng医疗卡类别ID = 0 Or dblMoney = 0 Then zlInterfacePrayMoney = True: Exit Function
    If cboPayMode.ItemData(cboPayMode.ListIndex) <> -1 Then zlInterfacePrayMoney = True: Exit Function
    'zlPaymentMoney(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strBalanceIDs As String, _
    byval  strPrepayNos as string , _
    ByVal dblMoney As Double, _
    ByRef strSwapGlideNO As String, _
    ByRef strSwapMemo As String, _
    Optional ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款交易
    '入参:frmMain-调用的主窗体
    '        lngModule-调用模块号
    '        strBalanceIDs-结帐ID,多个用逗号分离
    '        strPrepayNos-缴预交时有效. 预交单据号,多个用逗号分离
    '       strCardNo-卡号
    '       dblMoney-支付金额
    '出参:strSwapGlideNO-交易流水号
    '       strSwapMemo-交易说明
    '       strSwapExtendInfor-交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
    '返回:扣款成功,返回true,否则返回Flase
    '说明:
    '   在所有需要扣款的地方调用该接口,目前规划在:收费室；挂号室;自助查询机;医技工作站；药房等。
    '   一般来说，成功扣款后，都应该打印相关的结算票据，可以放在此接口进行处理.
    '   在扣款成功后，返回交易流水号和相关备注说明；如果存在其他交易信息，可以放在交易说明中以便退费.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModul, lng医疗卡类别ID, bln消费卡, mstrCardNO, lng挂号结帐ID, "", dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    '更新三交交易数据
     If lng挂号结帐ID <> 0 Then
        '问题:58322
        'mbytMode As Integer '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
        If Not bln消费卡 Then
            '消费卡已经在插入挂号记录时,已经扣款
            Call zlAddUpdateSwapSQL(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapGlideNO, strSwapMemo, cllPro)
        End If
        Call zlAddThreeSwapSQLToCollection(False, lng挂号结帐ID, lng医疗卡类别ID, bln消费卡, mstrCardNO, strSwapExtendInfor, cllThreeSwap)
    End If
    zlInterfacePrayMoney = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlAddThreeSwapSQLToCollection(ByVal bln预交款 As Boolean, _
    ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal str卡号 As String, strExpend As String, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存三方结算数据
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    ' 出参:cllPro-返回SQL集
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-19 10:23:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng结帐ID As Long, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim strSql As String, varData As Variant, varTemp As Variant, i As Long
     
    Err = 0: On Error GoTo Errhand:
    '先提交,这样避免风险,再更新相关的交易信息
    'strExpend:交易扩展信息,格式:项目名称|项目内容||...
    varData = Split(strExpend, "||")
    Dim str交易信息 As String, strTemp As String
    For i = 0 To UBound(varData)
        If Trim(varData(i)) <> "" Then
            varTemp = Split(varData(i) & "|", "|")
            If varTemp(0) <> "" Then
                strTemp = varTemp(0) & "|" & varTemp(1)
                If gobjCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                    str交易信息 = Mid(str交易信息, 3)
                    'Zl_三方结算交易_Insert
                    strSql = "Zl_三方结算交易_Insert("
                    '卡类别id_In 病人预交记录.卡类别id%Type,
                    strSql = strSql & "" & lng卡类别ID & ","
                    '消费卡_In   Number,
                    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
                    '卡号_In     病人预交记录.卡号%Type,
                    strSql = strSql & "'" & str卡号 & "',"
                    '结帐ids_In  Varchar2,
                    strSql = strSql & "'" & strIDs & "',"
                    '交易信息_In Varchar2:交易项目|交易内容||...
                    strSql = strSql & "'" & str交易信息 & "',"
                    '预交款缴款_In Number := 0
                    strSql = strSql & IIf(bln预交款, "1", "0") & ")"
                    zlAddArray cllPro, strSql
                    str交易信息 = ""
                End If
                str交易信息 = str交易信息 & "||" & strTemp
            End If
        End If
    Next
    If str交易信息 <> "" Then
        str交易信息 = Mid(str交易信息, 3)
        'Zl_三方结算交易_Insert
        strSql = "Zl_三方结算交易_Insert("
        '卡类别id_In 病人预交记录.卡类别id%Type,
        strSql = strSql & "" & lng卡类别ID & ","
        '消费卡_In   Number,
        strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
        '卡号_In     病人预交记录.卡号%Type,
        strSql = strSql & "'" & str卡号 & "',"
        '结帐ids_In  Varchar2,
        strSql = strSql & "'" & strIDs & "',"
        '交易信息_In Varchar2:交易项目|交易内容||...
        strSql = strSql & "'" & str交易信息 & "',"
        '预交款缴款_In Number := 0
        strSql = strSql & IIf(bln预交款, "1", "0") & ")"
        zlAddArray cllPro, strSql
    End If
    zlAddThreeSwapSQLToCollection = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlIsAllowPatiChargeFeeMode(ByVal lng病人ID As Long, ByVal int原结算模式 As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否允许改变病人收费模式
    '入参:lng病人ID-病人ID
    '       int原结算模式-0表示先结算后诊疗;1表示先诊疗后结算
    '返回:允许调整收费模式,返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-12-25 10:06:49
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim dtDate As Date, intDay As Integer
    On Error GoTo errHandle
    
'    If mbytMode = 1 Then zlIsAllowPatiChargeFeeMode = True: Exit Function '预约不处理
    '模式未调整，直接返回true
    If int原结算模式 = mPatiChargeMode Then zlIsAllowPatiChargeFeeMode = True: Exit Function
    
      
    If int原结算模式 = 1 Then
        '原为先诊疗后结算且存在未结费用的,则必须采用记帐模式
        strSql = "" & _
        "   Select 1 " & _
        "   From 病人未结费用 " & _
        "   Where 病人id = [1] And (来源途径 In (1, 4) Or 来源途径 = 3 And Nvl(主页id, 0) = 0) And Nvl(金额, 0) <> 0 And Rownum < 2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID)
        If rsTemp.EOF = False Then
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算且存在未结费用，" & _
                                          vbCrLf & "不允许调整该病人的就诊模式,你可以先对未结费用结帐后" & _
                                          vbCrLf & "再挂号或不调整病人的就诊模式", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        intDay = -1 * Val(Left(gobjDatabase.GetPara(21, glngSys, , "01") & "1", 1))
        dtDate = DateAdd("d", intDay, gobjDatabase.Currentdate)
        ' 上次为"先诊疗后结算",本次为"先结算后诊疗"的,同时满足未发生医嘱业务数据的 ,
        '   则不允许更改就诊模式
        strSql = "Select 1 " & _
        " From 病人挂号记录 A, 病人医嘱记录 B " & _
        " Where a.病人id + 0 = b.病人id And a.No || '' = b.挂号单  " & _
        "               And a.记录状态 = 1 And a.记录性质 = 1 And a.登记时间 - 0 >= [2] " & _
        "               And  a.病人id = [1] And rownum<2"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, dtDate)
        If rsTemp.EOF Then
            '未发生医嘱数据
            MsgBox "注意:" & vbCrLf & "  当前病人的就诊模式为先诊疗后结算," & vbCrLf & "  不允许调整该病人的就诊模式!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    zlIsAllowPatiChargeFeeMode = True
    Exit Function
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Function zlAddUpdateSwapSQL(ByVal bln预交 As Boolean, ByVal strIDs As String, ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    str卡号 As String, str交易流水号 As String, str交易说明 As String, _
    ByRef cllPro As Collection, Optional int校对标志 As Integer = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新三方交易流水号和流水说明
    '入参: bln预交款-是否预交款
    '       lngID-如果是预交款,则是预交ID,否则结帐ID
    '出参:cllPro-返回SQL集
    '编制:刘兴洪
    '日期:2011-07-27 10:13:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    strSql = "Zl_三方接口更新_Update("
    '  卡类别id_In   病人预交记录.卡类别id%Type,
    strSql = strSql & "" & lng卡类别ID & ","
    '  消费卡_In     Number,
    strSql = strSql & "" & IIf(bln消费卡, 1, 0) & ","
    '  卡号_In       病人预交记录.卡号%Type,
    strSql = strSql & "'" & str卡号 & "',"
    '  结帐ids_In    Varchar2,
    strSql = strSql & "'" & strIDs & "',"
    '  交易流水号_In 病人预交记录.交易流水号%Type,
    strSql = strSql & "'" & str交易流水号 & "',"
    '  交易说明_In   病人预交记录.交易说明%Type
    strSql = strSql & "'" & str交易说明 & "',"
    '预交款缴款_In Number := 0
    strSql = strSql & "" & IIf(bln预交, 1, 0) & ","
    '退费标志 :1-退费;0-付费
    strSql = strSql & "0,"
    '校对标志
    strSql = strSql & "" & IIf(int校对标志 = 0, "NULL", int校对标志) & ")"
    zlAddArray cllPro, strSql
End Function

Private Sub dtpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        gobjCommFun.PressKeyEx vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    Err = 0
    mblnInit = True
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hWnd)
    Call mobjICCard.SetParent(Me.hWnd)
    Call InitPara
    Call InitIDKind
    Call GetAll医生
    Call RestoreWinState(Me, App.ProductName)
    Call LoadRegPlans(False)
    Call InitFilter
    Call LoadPayMode
    Call SetControl
    glngFormW = 10680: glngFormH = 7425
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    vsfArrange.Height = vsfDetailTime.Top + vsfDetailTime.Height - vsfArrange.Top
    vsfDetailTime.Visible = False
    '137272:李南春,2019/2/20,防止锁号后系统意外崩溃的情况
    Call CancelRegNo
End Sub

Private Sub InitFilter()
    Dim strExists
    On Error GoTo errH
    If mrsPlan Is Nothing Then Exit Sub
    If mrsPlan.State = 0 Then Exit Sub
    If mrsPlan.RecordCount = 0 Then Exit Sub
    mrsPlan.MoveFirst
    strExists = ","
    cboDeptFilter.AddItem " "
    Do While Not mrsPlan.EOF
        If InStr(strExists, "," & Nvl(mrsPlan!科室, "") & ",") = 0 Then
            cboDeptFilter.AddItem Nvl(mrsPlan!科室)
            strExists = strExists & Nvl(mrsPlan!科室) & ","
        End If
        mrsPlan.MoveNext
    Loop
    
    mrsPlan.MoveFirst
    strExists = ","
    cboDoctorFilter.AddItem ""
    Do While Not mrsPlan.EOF
        If InStr(strExists, "," & Nvl(mrsPlan!医生, "") & ",") = 0 And Not IsNull(mrsPlan!医生) Then
            cboDoctorFilter.AddItem Nvl(mrsPlan!医生)
            strExists = strExists & Nvl(mrsPlan!医生) & ","
        End If
        mrsPlan.MoveNext
    Loop
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraSplit.Width = Me.Width + 300
    fraInfo.Height = Me.Height / 2 - 500
    lblRemark.Top = fraInfo.Height - 400
    cboRemark.Top = fraInfo.Height - 450
    lblAppointStyle.Top = lblRemark.Top
    cboAppointStyle.Top = cboRemark.Top
    vsfMoney.Height = fraInfo.Height - vsfMoney.Top - 500
    cboNO.Left = Me.Width - 2100
    lblNO.Left = Me.Width - 3000
    fraPay.Left = Me.Width - 5060
    fraTotal.Left = Me.Width - 5060
    fraInfo.Left = Me.Width - 5060
    lblMoney.Left = fraInfo.Left
    fraTime.Width = fraInfo.Left - 90
    cmdPrice.Left = fraInfo.Left + 150
    cmdOK.Left = Me.Width - 2750
    cmdCancel.Left = Me.Width - 1500
    fraTime.Height = Me.Height - fraTime.Top - 600
    vsfArrange.Width = fraTime.Width - 150
    vsfDetailTime.Width = fraTime.Width - 150
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")) <> "" Then Call GetActiveView
    lblMoney.Top = (Me.Height - fraSplit.Top - fraInfo.Height - 1100) / 2
    fraInfo.Top = lblMoney.Top + lblMoney.Height + 30
    fraTotal.Top = fraInfo.Top + fraInfo.Height + 30
    fraPay.Top = fraTotal.Top + fraTotal.Height + 30
    cmdPrice.Top = fraPay.Top + fraPay.Height + 30
    cmdOK.Top = cmdPrice.Top
    cmdCancel.Top = cmdPrice.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    Call SaveWinState(Me, App.ProductName)
    Set mobjRegistData = Nothing
    mstrPreNO = ""
    
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strExpand As String
    Dim strOutCardNO As String, strOutPatiInforXML As String
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        '系统IC卡
        If Not mobjICCard Is Nothing Then
            txtPatient.Text = mobjICCard.Read_Card()
            If txtPatient.Text <> "" Then
                Call GetPatient(objCard, txtPatient.Text, True)
            End If
        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        Call GetPatient(objCard, txtPatient.Text, True)
    End If
End Sub

Private Function Check复诊(ByVal lng病人ID As Long, ByVal lng执行部门ID As Long) As Boolean
'功能:判断病人是否再次到“相同临床性质的临床科室”挂号
'     包括挂过号的,或住过院的,复诊不好确定时间
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    strSql = "Select a.临床科室id" & vbNewLine & _
    "       From (Select 执行部门id 临床科室id From 病人挂号记录 Where 病人id = [1] and 记录性质=1 and 记录状态=1 " & vbNewLine & _
    "             Union All" & vbNewLine & _
    "             Select 出院科室id 临床科室id From 病案主页 Where 病人id = [1]) a" & vbNewLine & _
    "       Where Exists (Select 1" & vbNewLine & _
    "                    From 临床部门 b" & vbNewLine & _
    "                    Where b.部门id = a.临床科室id And b.工作性质 = (Select 工作性质 From 临床部门 Where 部门id = [2] And Rownum=1))"

    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng执行部门ID)
    Check复诊 = Not rsTmp.EOF
End Function

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Visible Then txtPatient.SetFocus
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.卡号
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
End Sub

Private Sub txtArrangeNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If vsfArrange.Row - 1 >= vsfArrange.FixedRows Then
                KeyCode = 0
                vsfArrange.Row = vsfArrange.Row - 1
                vsfArrange_EnterCell
            End If
        Case vbKeyDown
            If vsfArrange.Row + 1 <= vsfArrange.Rows - 1 Then
                KeyCode = 0
                vsfArrange.Row = vsfArrange.Row + 1
                vsfArrange_EnterCell
            End If
        Case 13
            Call vsfArrange_KeyDown(13, 0)
    End Select
End Sub

Private Sub txtPatient_Change()
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub txtPatient_GotFocus()
    Call gobjControl.TxtSelAll(txtPatient)
    Call gobjCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
End Sub

Private Sub zlInusreIdentify()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：医保身份验卡
    '编制：刘兴洪
    '日期：2010-07-14 11:32:08
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim str病人类型 As String
    Dim rsTmp As ADODB.Recordset
    Dim cur余额 As Currency
    Dim curMoney As Currency
    Dim blnDeposit As Boolean, blnInsure As Boolean
    If mrsInfo Is Nothing Then
        lng病人ID = 0
        str病人类型 = ""
    Else
        lng病人ID = Val(Nvl(mrsInfo!病人ID))
        str病人类型 = Nvl(mrsInfo!病人类型)
    End If

    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard False

    Dim strAdvance As String    '结算模式(0-先结算后诊疗或1-先诊疗后结算)|挂号费收取方式(0-现收或1-记帐)
    Dim varData As Variant
    mstrYBPati = gclsInsure.Identify(3, lng病人ID, mintInsure, strAdvance)
    mRegistFeeMode = EM_RG_现收: mPatiChargeMode = EM_先结算后诊疗
    If txtPatient.Text = "" And Not txtPatient.Locked Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(True)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(True)
        IDKind.SetAutoReadCard True
    End If
    
    If mstrYBPati = "" Then
        If Not txtPatient.Enabled Then txtPatient.Enabled = True
         mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
         Exit Sub
    End If
    
    '空或：0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);8病人ID
    If UBound(Split(mstrYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mstrYBPati, ";")(8)) Then lng病人ID = Val(Split(mstrYBPati, ";")(8))
    End If
        
    If lng病人ID = 0 Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If

    If zlPatiCardCheck(1, lng病人ID, CStr(Split(mstrYBPati, ";")(0)), 2) = False Then
        mstrYBPati = "": mintInsure = 0: txtPatient.SetFocus
        Exit Sub
    End If
    MCPAR.挂号检查项目 = gclsInsure.GetCapability(support挂号检查项目, lng病人ID, mintInsure)
    txtPatient.Text = "-" & lng病人ID
    Call txtPatient_Validate(False)    '其中的Setfocus调用使本事件(txtPatient_KeyPress)执行完后,不会再次自动执行txtPatient_Validate
    Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
    Call SetPatiColor(txtPatient, str病人类型, vbRed)
    txtPatient.BackColor = &HE0E0E0
    txtPatient.Locked = True

    If strAdvance <> "" Then
        varData = Split(strAdvance & "|", "|")
        mPatiChargeMode = IIf(Val(varData(0)) = 1, EM_先诊疗后结算, EM_先结算后诊疗)
        mRegistFeeMode = IIf(Val(varData(1)) = 1, EM_RG_记帐, EM_RG_现收)
    End If
    If mRegistFeeMode = EM_RG_记帐 Then
        fraPay.Visible = False
    End If
    If mRegistFeeMode = EM_RG_现收 Then
        If mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2 Then
            mRegistFeeMode = EM_RG_现收
        Else
            mRegistFeeMode = EM_RG_划价
        End If
    End If
    MCPAR.不收病历费 = gclsInsure.GetCapability(support挂号不收取病历费, lng病人ID, mintInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure)
    mlng领用ID = 0
    curMoney = GetRegistMoney
    Set rsTmp = GetMoneyInfoRegist(lng病人ID, , , 1)
    If Not rsTmp Is Nothing Then cur余额 = rsTmp!预交余额 - rsTmp!费用余额
    If cur余额 > 0 Then
        lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
        If cur余额 >= curMoney Then
            blnDeposit = True
        Else
            blnDeposit = False
        End If
    End If
    mcur个帐余额 = gclsInsure.SelfBalance(lng病人ID, CStr(Split(mstrYBPati, ";")(1)), 10, mcur个帐透支, mintInsure)
    lblMoney.Caption = lblMoney.Caption & "  个人帐户余额:" & Format(mcur个帐余额, "0.00")
    If gclsInsure.GetCapability(support挂号使用个人帐户, lng病人ID, mintInsure) = False Then
        blnInsure = False
    Else
        If mcur个帐余额 + mcur个帐透支 >= curMoney Then
            blnInsure = True
        Else
            blnInsure = False
        End If
    End If
    Call LoadPayMode(blnDeposit, blnInsure)
    If mRegistFeeMode = EM_RG_记帐 Then
        lblSum.Caption = "记 帐"
    End If
    If mRegistFeeMode = EM_RG_现收 Then
        If mblnAppointment Then
            mRegistFeeMode = EM_RG_现收
        Else
            If mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2 Then
                mRegistFeeMode = EM_RG_现收
                fraPay.Visible = True
                cmdPrice.Visible = mty_Para.byt挂号模式 = 2
            Else
                mRegistFeeMode = EM_RG_划价
                fraPay.Visible = False
                cmdPrice.Visible = False
            End If
        End If
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    '0-门诊号,1-姓名,2-挂号单,3-就诊卡号,4-医保号
    Dim blnCard As Boolean
    Dim strKind As String, intLen As Integer
    Static sngBegin As Single
    Dim sngNow As Single
    
    '医保验证
    If txtPatient.Text = "" And KeyAscii = 13 Then
        KeyAscii = 0
        Call zlInusreIdentify
    End If
    
    If KeyAscii <> 0 And KeyAscii > 32 And mty_Para.bln挂号必须刷卡 Then
        sngNow = Timer
        If txtPatient.Text = "" Then
            sngBegin = sngNow
        ElseIf Format((sngNow - sngBegin) / (Len(txtPatient.Text) + 1), "0.000") >= 0.04 Then    '>0.007>=0.01
            txtPatient.Text = Chr(KeyAscii)
            txtPatient.SelStart = 1
            KeyAscii = 0
            sngBegin = sngNow
        End If
    End If
    
    strKind = IDKind.GetCurCard.名称
    txtPatient.PasswordChar = IIf(IDKind.GetCurCard.卡号密文规则 <> "", "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
    
    '取缺省的刷卡方式
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
    Select Case strKind
    Case "姓名"
        blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, gobjSquare.bln缺省卡号密文)
        intLen = gobjSquare.int缺省卡号长度
    Case "门诊号"
        If InStr("0123456789-" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    Case "挂号单"
    Case "医保号"
    Case Else
            If IDKind.GetCurCard.接口序号 <> 0 Then
                blnCard = gobjCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.GetCurCard.卡号密文规则 <> "")
                intLen = IDKind.GetCurCard.卡号长度
            End If
    End Select
    
    '刷卡完毕或输入号码后回车
    If (blnCard And Len(txtPatient.Text) = intLen - 1 And KeyAscii <> 8) Or (KeyAscii = 13) Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0: mblnCard = True
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), blnCard)
        mblnCard = False
        gobjControl.TxtSelAll txtPatient
   End If
End Sub

Private Function CheckNoValied(ByVal lngRow As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：检查指定行的号别是否有效
    '返回：有效,返回true,否则返回False
    '编制：刘兴洪
    '日期：2010-08-17 16:00:11
    '说明：31922
    '------------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivs, ";临时挂号;") > 0 Then
        CheckNoValied = True: Exit Function
    End If
    With vsfArrange
        If Val(.Cell(flexcpData, lngRow, .ColIndex("号别"))) = 1 Then
            MsgBox "号别『" & .TextMatrix(lngRow, .ColIndex("号别")) & "』不在有效范围内或你权限不足,不能挂号,请检查!", vbInformation + vbOKOnly + vbDefaultButton1
            Exit Function
        End If
    End With
    CheckNoValied = True
End Function

Private Sub SetGridTop(intRow As Integer)
    Dim intRows As Integer
    intRows = vsfArrange.Height \ vsfArrange.RowHeight(1) - 2
    If vsfArrange.TopRow + intRows > intRow Then Exit Sub
    vsfArrange.TopRow = intRow
End Sub

Private Sub txtArrangeNo_Change()
    
    '功能：根据输入号别显示内容
    Dim strInfo As String, i As Integer
    Dim blnChkLimit As Boolean
    
    If mblnChangeByCode Then Exit Sub
    

    If mblnUnChange Then Exit Sub
     
    txtSN.Text = ""
   
    '刷新号别直接从缓存中读取数据
    If vsfArrange.Tag = "" Then
        mblnUnChange = True
        Call LoadRegPlans(Len(txtArrangeNO) > 0 And IsNumeric(Trim(txtArrangeNO.Text)), True)
        mblnUnChange = False
    End If
    
    If Trim(txtArrangeNO.Text) = "" Then
'        chkBook.Enabled = mbln病历费
'        lblFree.Visible = False
        Exit Sub
    End If
     
    
    
    
    If Not ((IsNumeric(Trim(txtArrangeNO.Text)) And Len(Trim(txtArrangeNO.Text)) = mint号长 Or vsfArrange.Rows = 2) Or vsfArrange.Tag <> "") Then Exit Sub
    
    With vsfArrange
        If .Tag = "" Then
            If .Rows <> 2 And Trim(txtArrangeNO.Text) <> .TextMatrix(.Row, .ColIndex("号别")) Then
                '当前号别列表只有一行时，如果没有输完整号别，不自动匹配，除非按回车
                Exit Sub
            End If
            
            '定位表格中的号别
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("号别"))) = Trim(txtArrangeNO.Text) Then
                    If CheckNoValied(i) = False Then
                         txtArrangeNO.Text = "": txtArrangeNO.SetFocus: Exit Sub
                    End If
                    
                    Call vsfArrange_LeaveCell
                    .Row = i:  .RowSel = i
                    .Col = 0:  .ColSel = vsfArrange.Cols - 1
                    Call vsfArrange_EnterCell
                    
                    SetGridTop i
                    Exit For
                End If
            Next
            If mrsPlan Is Nothing Then
                txtArrangeNO.Text = ""
                txtArrangeNO.SetFocus: Exit Sub
            End If
            '号表中无安排时要求重输
            If i = vsfArrange.Rows And mrsPlan.RecordCount = 0 Then
                txtArrangeNO.Text = ""
                txtArrangeNO.SetFocus: Exit Sub
            End If
        End If
        If mblnAppointment Then
            blnChkLimit = .TextMatrix(.Row, .ColIndex("限约")) <> ""
            If Not blnChkLimit Then blnChkLimit = .TextMatrix(.Row, .ColIndex("限号")) <> ""
        Else
            blnChkLimit = .TextMatrix(.Row, .ColIndex("限号")) <> ""
        End If

        '限号控制
        If blnChkLimit Then
            If zlCheck限约或限号数(txtArrangeNO.Text) = False Then Exit Sub
        End If
        txtSN.Enabled = .TextMatrix(.Row, .ColIndex("序号控制")) <> ""
        
        If txtSN.Enabled And vsfDetailTime.Tag = "" And vsfDetailTime.Visible Then
            txtSN.Text = GetCurrSN(, Not mty_Para.bln随机序号选择)
            If Val(txtSN.Text) = 0 Then
                txtSN.Text = ""
                If CheckArangement = False Then Exit Sub
            Else
                Call LocateSN(Val(txtSN.Text))
            End If
        End If
    End With
End Sub

Public Function Get失约号(ByVal str号别 As String, ByVal datThis As Date) As Long
   '获取安排在某一天.预约失约数
    Dim strSql  As String
    Dim rsTmp   As ADODB.Recordset
    Dim strDat  As String
'    If mty_Para.bln失约用于挂号 = False Or mty_Para.lng预约有效时间 <= 0 Then Exit Function
    strSql = "                " & " SELECT count(1) AS 失约号 "
    strSql = strSql & vbNewLine & " FROM 挂号序号状态 "
    strSql = strSql & vbNewLine & " WHERE 号码=[1] AND 状态=2 AND 日期-[3]/24/60 <SYSDATE AND To_Char(日期,'yyyy-MM-dd')=[2]"
    strDat = Format(datThis, "yyyy-MM-dd")
    On Error GoTo Hd
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str号别, strDat, mty_Para.lng预约有效时间)
    If rsTmp.EOF Then
        Get失约号 = 0
        Set rsTmp = Nothing
        Exit Function
    End If
    Get失约号 = Val(Nvl(rsTmp!失约号, 0))
    Set rsTmp = Nothing
   Exit Function
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function zlCheck限约或限号数(ByVal str号别 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查限约数或限号数是否合法
    '入参:str号别-号别
    '出参:
    '返回:合法,返回ture,否则返回False
    '编制:刘兴洪
    '日期:2009-12-30 15:15:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, lngTemp As Long, strSql As String, curDate As Date
    Dim lng限约数 As Long, lng限号数 As Long, lng已挂数 As Long, lng已约数 As Long, lng剩余预约数 As Long
    Dim lng失约数 As Long
    Dim bln分时段 As Boolean
    Dim strMsg As String
    Dim lng合作单位数量 As Long
    Dim blnHaveUnitreg As Boolean
    Dim i As Integer, j As Integer
    Err = 0: On Error GoTo Errhand:
    lng限约数 = 0: lng限号数 = 0: lng已挂数 = 0: lng已约数 = 0: lng剩余预约数 = 0

    curDate = CDate(Format(gobjDatabase.Currentdate, "yyyy-MM-dd"))
    strSql = _
      "Select Nvl(C.限号数,0) as 限号数,Nvl(B.已挂数,0)  as 已挂数,Nvl(C.限约数,0) as 限约数,Nvl(B.已约数,0) as 已约数,NVL(B.其中已接收,0) as 已接收" & _
      " From 挂号安排 A,病人挂号汇总 B,挂号安排限制 C " & _
      " Where A.科室ID=B.科室ID(+) And A.项目ID=B.项目ID(+)  " & _
      "       And A.号码=[1] And B.日期(+)=[2] And A.号码=B.号码(+) " & _
      "       And Nvl(A.医生ID,0)=Nvl(B.医生ID(+),0) And Nvl(A.医生姓名,'医生')=Nvl(B.医生姓名(+),'医生') And  A.ID = C.安排id(+)" & _
      "       And Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五','7', '周六', Null) = C.限制项目(+)"

   
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, str号别, curDate, CDate(Format(curDate, "YYYY-MM-DD")))

    lng失约数 = Get失约号(str号别, curDate)

    If Not rsTmp.EOF Then
        lng限约数 = Val(Nvl(rsTmp!限约数)): lng限号数 = Val(Nvl(rsTmp!限号数))
        lng已挂数 = Val(Nvl(rsTmp!已挂数)): lng已约数 = Val(Nvl(rsTmp!已约数)) - Val(Nvl(rsTmp!已接收))
        If lng已约数 < 0 Then lng已约数 = 0
        lng剩余预约数 = IIf(lng限号数 - lng已挂数 - lng已约数 <= 0, 0, lng限约数 - lng已约数): If lng剩余预约数 < 0 Then lng剩余预约数 = 0
        If lng限约数 = 0 Then lng限约数 = lng限号数
        lng已约数 = lng已约数 - lng失约数
    End If
    If lng限号数 <= 0 Then
        '不作限制:返回
        zlCheck限约或限号数 = True: Exit Function
    End If
    
    If lng已挂数 + lng已约数 >= lng限号数 Then
        If InStr(gstrPrivs, ";加号;") > 0 Then
            If MsgBox("该号别今天已达到限号数 " & lng限号数 & "，你是否继续挂号?", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                 mblnChangeByCode = True
                 txtArrangeNO = ""
                 mblnChangeByCode = False
                 If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
                 Exit Function
            End If
            With vsfDetailTime
                For i = 0 To .Rows - 1
                    For j = 0 To .Cols - 1
                        If .Cell(flexcpData, i, j) Like "加*" Then .Select i, j
                    Next j
                Next i
            End With
        Else
            MsgBox "该号别今天已达到限号数 " & lng限号数 & "不能再挂号！", vbInformation, gstrSysName
            mblnChangeByCode = True
            txtArrangeNO = ""
            mblnChangeByCode = False
            If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
            Exit Function
        End If
    End If
    
    zlCheck限约或限号数 = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter = 1 Then Resume
End Function

Private Sub txtArrangeNo_GotFocus()
    Call gobjControl.TxtSelAll(txtArrangeNO)
End Sub

Private Sub txtPatient_LostFocus()
    Call gobjCommFun.OpenIme
    IDKind.SetAutoReadCard False
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    txtPatient.Text = Trim(txtPatient.Text)
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), False)
'        gobjControl.TxtSelAll txtPatient
'    End If
End Sub

Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strID:
        If txtPatient.Text = "" Then
            Call mobjIDCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True, True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
        Case vbKeyF3
            If txtPatient.Visible = True And txtPatient.Enabled Then
                Call txtPatient.SetFocus
            End If
        Case vbKeyEscape
            Call ReloadPage
        Case Else
            IDKind.ActiveFastKey
    End Select
End Sub

Public Sub ActiveIDKindKey()
    IDKind.ActiveFastKey
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strNO As String)
    Dim lngPreIndex As Long
    If txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        txtPatient.Text = strNO
        If txtPatient.Text = "" Then
            Call mobjICCard.SetEnabled(False) '如果不符合发卡条件，禁用继续自动读取
            Exit Sub
        End If
        lngPreIndex = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("IC卡号")
        Call GetPatient(IDKind.GetCurCard, Trim(txtPatient.Text), True)
        IDKind.IDKind = lngPreIndex
    End If
End Sub

Private Sub txtArrangeNo_KeyPress(KeyAscii As Integer)
    cboDeptFilter.ListIndex = 0
    cboDeptFilter.ListIndex = 0
    If KeyAscii = Asc(".") Then
        '相关于按回退键
        KeyAscii = 0: gobjCommFun.PressKey vbKeyBack
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If CheckNoValied(vsfArrange.Row) = False Then
             txtArrangeNO.Text = "": txtArrangeNO.SetFocus: Exit Sub
        End If
        
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        If vsfArrange.Tag <> "" Then
            If txtArrangeNO.Text <> vsfArrange.Tag Then
                txtArrangeNO.Text = vsfArrange.Tag  '自动调用change事件
            Else
                Call txtArrangeNo_Change
            End If
            vsfArrange.Tag = ""
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890+ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        Else
            vsfArrange.Tag = ""
        End If
    End If
End Sub

Private Sub GetPatient(objCard As zlIDKind.Card, ByVal strInput As String, ByVal blnCard As Boolean, Optional blnInputIDCard As Boolean = False, Optional ByRef Cancel As Boolean)
    '功能：获取病人信息
    '参数：blnCard=是否就诊卡刷卡
    '
    '         blnInputIDCard-是否身份证刷卡
    '出参:Cancel-为true表示返回的放弃读取病人信息
    Dim strSql As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    Dim rsTmp As ADODB.Recordset, strTemp As String, rsFeeType As ADODB.Recordset
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim strInputInfo As String '保存传入的输入文本 避免在使用身份证号 对病人进行查找后 被替换成"-" 病人ID的情况
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim bln医保号 As Boolean
    Dim IntMsg As VbMsgBoxResult
    Dim blnOtherType As Boolean '非法卡类别

    strInputInfo = strInput
    
    On Error GoTo errH
    bln医保号 = False
    
    If objCard Is Nothing Then Set objCard = IDKind.GetCurCard

    strSql = "Select  A.病人ID,A.门诊号,A.住院号,A.就诊卡号,A.费别,A.医疗付款方式,A.姓名,A.性别,A.年龄,A.出生日期,A.出生地点,A.身份证号,A.其他证件,A.身份,A.职业,A.民族,A.病人类型, " & _
             "A.国籍,A.籍贯,A.区域,A.学历,A.婚姻状况,A.家庭地址,A.家庭电话,A.家庭地址邮编,A.监护人,A.联系人姓名,A.联系人关系,A.联系人地址,A.联系人电话,A.户口地址, " & _
             "A.户口地址邮编,A.Email,A.QQ,A.合同单位id,A.工作单位,A.单位电话,A.单位邮编,A.单位开户行,A.单位帐号,A.担保人,A.担保额,A.担保性质,A.就诊时间,A.就诊状态, " & _
             "A.就诊诊室,A.住院次数,A.当前科室id,A.当前病区id,A.当前床号,A.入院时间,A.出院时间,A.在院,A.IC卡号,A.健康号,A.医保号,A.险类,A.查询密码,A.登记时间,A.停用时间,A.锁定,A.联系人身份证号, " & _
             "B.名称 险类名称,A.查询密码 As 卡验证码,A.结算模式,a.主页ID From 病人信息 A,保险类别 B Where A.险类 = B.序号(+) And A.停用时间 is NULL  "

    
    If mty_Para.bln住院病人挂号 = False Then
        str非在院 = " And Not Exists(Select 1 From 病案主页 Where 病人ID=A.病人ID   And 主页ID<>0 And 主页ID=A.主页ID And Nvl(病人性质,0)=0 And 出院日期 is Null)"
    End If
   
    If blnCard And objCard.名称 Like "姓名*" And mstrYBPati = "" And InStr("-+*.", Left(strInput, 1)) = 0 Then     '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
'        Else
'            lng卡类别ID = gCurSendCard.lng卡类别ID
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        
        If IDKind.IsMobileNo(strInput) And lng病人ID = 0 Then
            If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        End If
        
        If lng病人ID <= 0 Then GoTo NewPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSql = strSql & " And A.门诊号=[2]" & str非在院
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSql = strSql & " And A.病人ID=[2]" & _
        IIf(mstrYBPati <> "", "", str非在院)
    ElseIf blnInputIDCard Then  '单独的身份证识别
        strInput = UCase(strInput)
        If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        strInput = "-" & lng病人ID
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    ElseIf objCard.名称 Like "姓名*" And IDKind.IsMobileNo(strInput) = True Then
        If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
        strInput = "-" & lng病人ID
        strSql = strSql & " And A.病人ID=[2] " & str非在院
    Else
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                If Not mty_Para.bln姓名模糊查找 Or mty_Para.bln姓名模糊查找 And Len(txtPatient.Text) < 2 Then
                    Set mrsInfo = Nothing: Exit Sub
                End If
                strPati = _
                    " Select distinct 1 as 排序ID,A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄,A.门诊号,A.出生日期,A.身份证号,A.家庭地址,A.工作单位" & _
                    " From 病人信息 A " & _
                    " Where Rownum <101 And A.停用时间 is NULL And A.姓名 Like [1]" & str非在院 & _
                    IIf(mty_Para.lng姓名查找天数 = 0, "", " And Nvl(A.就诊时间,A.登记时间)>Trunc(Sysdate-[2])")
                    
                strPati = strPati & " Union ALL " & _
                        "Select 0,0 as ID,-NULL,'[新病人]',NULL,NULL,-NULL,To_Date(NULL),NULL,NULL,NULL From Dual"
                strPati = strPati & " Order by 排序ID,姓名"
                    
                vRect = GetControlRect(txtPatient.hWnd)
                Set rsTmp = gobjDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", mty_Para.lng姓名查找天数)
                If Not rsTmp Is Nothing Then
                    If rsTmp!ID = 0 Then '当作新病人
                        txtPatient.Text = ""
                        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
                        Set mrsInfo = Nothing: Exit Sub
                    Else '以病人ID读取
                        strInput = rsTmp!病人ID
                        strSql = strSql & " And A.病人ID=[1]"
                    End If
                Else '取消选择
                    txtPatient.Text = ""
                    Set mrsInfo = Nothing: Exit Sub
                End If
            Case "医保号"
                strInput = UCase(strInput)
                bln医保号 = True
                If mblnOlnyBJYB And gobjCommFun.ActualLen(strInput) >= 9 Then
                    strSql = strSql & " And A.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSql = strSql & " And A.医保号=[1]" & str非在院
                End If
            Case "手机号"
                If IDKind.IsMobileNo(strInput) = False Then Exit Sub
                If gobjSquare.objSquareCard.zlGetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then Exit Sub
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
                 
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strSql = strSql & " And A.病人ID=[2] " & str非在院
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSql = strSql & " And A.门诊号=[1]" & str非在院
             Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    blnOtherType = True
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                strSql = strSql & " And A.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
ReadPati:
    If Mid(mstrCardPass, 1, 1) = "1" And strPassWord <> "" Then
        If Not gobjCommFun.VerifyPassWord(Me, "" & strPassWord) Then
            MsgBox "病人身份验证失败！", vbInformation, gstrSysName
            ClearPatient
            Exit Sub
        End If
    End If
    Set mrsInfo = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, strInput, Mid(strInput, 2), strTemp)
    strInput = strInputInfo
    If Not mrsInfo.EOF Then
        txtPatient.Text = Nvl(mrsInfo!姓名) '会调用Change事件
        txtPatient.BackColor = &H80000005
        lblSum.Caption = "合 计"
        If mblnAppointment Then
            fraPay.Visible = False
        Else
            If mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2 Then
                fraPay.Visible = True
                cmdPrice.Visible = mty_Para.byt挂号模式 = 2
            Else
                fraPay.Visible = False
                cmdPrice.Visible = False
            End If
        End If
        '在调用txtPatient_Change事件后在门诊号和病人姓名都为空的情况下 无法识别该病人信息 出现错误
        '对这类数据库数据错误不再进行后续的处理
        If mrsInfo Is Nothing Then Cancel = True: Exit Sub
        Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(Trim(mstr险类) = "", txtPatient.ForeColor, vbRed))
        
        mstrPassWord = strPassWord
        If Not blnHavePassWord Then mstrPassWord = Nvl(mrsInfo!卡验证码)
        txtGender.Text = Nvl(mrsInfo!性别)
        txtPatient.PasswordChar = ""
        
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
        txtFeeType.Text = Nvl(mrsInfo!费别)
        If txtFeeType.Text = "" Then
            strSql = "Select 名称 From 费别 Where 缺省标志 = 1"
            Set rsFeeType = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
            If Not rsFeeType.EOF Then
                txtFeeType.Text = Nvl(rsFeeType!名称)
            End If
        End If
        txtAge.Text = Nvl(mrsInfo!年龄)
        
        mblnUpdateAge = False
        If Not IsNull(mrsInfo!出生日期) Then
            strSql = "Select Zl_Age_Calc([1],[2],Null) As Old From Dual"
            Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, App.ProductName, lng病人ID, CDate(mrsInfo!出生日期))
            If txtAge.Text <> Nvl(rsTmp!old) And Nvl(rsTmp!old) <> "" Then
                mblnUpdateAge = True
                txtAge.Text = Nvl(rsTmp!old)
            End If
        End If
        
        txtClinic.Text = Nvl(mrsInfo!门诊号)
        If txtClinic.Text = "" Then
            txtClinic.Text = gobjDatabase.GetNextNo(3)
            mblnChangeFeeType = True
        Else
            mblnChangeFeeType = False
        End If
        
        '病人预交款信息
        Set rsTmp = GetMoneyInfoRegist(mrsInfo!病人ID, , , 1)
        If Not rsTmp Is Nothing Then cur余额 = rsTmp!预交余额 - rsTmp!费用余额
        If cur余额 > 0 Then
            lblMoney.Caption = "门诊预交余额:" & Format(cur余额, "0.00")
            curMoney = GetRegistMoney
            If cur余额 >= curMoney Then
                Call LoadPayMode(True)
            Else
                Call LoadPayMode
            End If
        Else
            lblMoney.Caption = "门诊预交余额:0.00"
            Call LoadPayMode
        End If
        
        '根据病人重新读取项目费用
        If mintPriceGradeStartType >= 2 Then
            Call GetPriceGrade(gstrNodeNo, Val(Nvl(mrsInfo!病人ID)), Val(Nvl(mrsInfo!主页ID)), Nvl(mrsInfo!医疗付款方式), , , mstrPriceGrade)
            '重新加载费用信息
            Call LoadFeeItem(Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1, mstrPriceGrade)
        End If
        
        If txtArrangeNO.Enabled And txtArrangeNO.Visible Then txtArrangeNO.SetFocus
    Else
NewPati:
        MsgBox "没有找到对应的病人信息，请检查输入信息是否正确或者病人是否建档！", vbInformation, gstrSysName
        ClearPatient
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub ClearPatient()
    txtPatient.Text = ""
    txtPatient.BackColor = &H80000005
    txtPatient.ForeColor = vbBlack
    txtPatient.Locked = False
    txtGender.Text = ""
    txtAge.Text = ""
    txtClinic.Text = ""
    txtFeeType.Text = ""
    lblMoney.Caption = ""
    chkBook.Enabled = True
    lblSum.Caption = "合计"
    If mblnAppointment Then
        mRegistFeeMode = EM_RG_现收
    Else
        If mty_Para.byt挂号模式 = 0 Or mty_Para.byt挂号模式 = 2 Then
            mRegistFeeMode = EM_RG_现收
            fraPay.Visible = True
            cmdPrice.Visible = mty_Para.byt挂号模式 = 2
        Else
            mRegistFeeMode = EM_RG_划价
            fraPay.Visible = False
            cmdPrice.Visible = False
        End If
    End If
    mintInsure = 0
    mlng领用ID = 0
    mstrPreNO = ""
    Set mrsInfo = Nothing
    LoadPayMode False, False
End Sub


Private Function GetTotalFromMshMoney(Optional ByVal str项目名称 As String = "") As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取汇总金额
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:57:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Long
    
    On Error GoTo errHandle
    With vsfMoney
        For i = 1 To .Rows - 1
            If str项目名称 = "" Or Trim(.TextMatrix(i, 0)) = str项目名称 Then
                dblMoney = dblMoney + Val(.TextMatrix(i, 2))
            End If
        Next
    End With
    GetTotalFromMshMoney = dblMoney
    Exit Function
errHandle:
    GetTotalFromMshMoney = 0
End Function



Private Function GetRegistMoney(Optional blnOnlyReg As Boolean = False, Optional blnNoBook As Boolean = False) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前挂号单的合计金额
    '入参:blnOnlyReg-是否仅仅读取挂号费用
    '     blnNoBook-读取病历费
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-11-03 16:53:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl合计 As Double, i As Integer
    Dim k As Integer
    
    If Not blnOnlyReg Then
        dbl合计 = FormatEx(GetTotalFromMshMoney, 5)
    Else
        If mrsItems Is Nothing Then
             GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        mrsItems.Filter = " 性质 <> 4"
        If mrsItems.RecordCount = 0 Then
            mrsItems.Filter = 0
            GetRegistMoney = FormatEx(GetTotalFromMshMoney, 3): Exit Function
        End If
        With mrsItems
            Do While Not .EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                .MoveNext
            Loop
        End With
        mrsItems.Filter = 0
    End If
    If blnNoBook Then
        If Not mrsItems Is Nothing Then
            mrsItems.Filter = " 性质 = 3"
            Do While Not mrsItems.EOF
                dbl合计 = dbl合计 + GetTotalFromMshMoney(Nvl(mrsItems!项目名称, "-"))
                mrsItems.MoveNext
            Loop
            mrsItems.Filter = 0
        End If
    End If
    GetRegistMoney = FormatEx(dbl合计, 5)
End Function




Private Sub LoadPayMode(Optional ByVal blnPrepay As Boolean = False, Optional ByVal blnInsure As Boolean = False)
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSql As String, str性质 As String
    
    strSql = _
        "Select B.编码,B.名称,Nvl(B.性质,1) as 性质,Nvl(A.缺省标志,0) as 缺省" & _
        " From 结算方式应用 A,结算方式 B" & _
        " Where A.应用场合=[1] And B.名称=A.结算方式 And Instr([2] ,','||B.性质||',')>0" & _
        " Order by B.编码"
    On Error GoTo errH
    Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "挂号", ",3,7,8,")
    
    Set mcolCardPayMode = New Collection
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    varData = Split(strPayType, ";")
    
    With cboPayMode
        .Clear: j = 0
'        Do While Not rsTemp.EOF
'            blnFind = False
'            For i = 0 To UBound(varData)
'                varTemp = Split(varData(i) & "|||||", "|")
'                If varTemp(6) = Nvl(rsTemp!名称) Then
'                    blnFind = True
'                    Exit For
'                End If
'            Next
'
'            If Not blnFind Then
'                .AddItem Nvl(rsTemp!名称)
'                mcolCardPayMode.Add Array("", Nvl(rsTemp!名称), 0, 0, 0, 0, Nvl(rsTemp!名称), 0, 0), "K" & j
'                If Val(Nvl(rsTemp!缺省)) = 1 Then
'                    If .ListIndex = -1 Then
'                         .ItemData(.NewIndex) = 1: .ListIndex = .NewIndex
'                    End If
'                End If
'                j = j + 1
'            End If
'            rsTemp.MoveNext
'        Loop
    
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                rsTemp.Filter = "名称='" & varTemp(6) & "'"
                If Not rsTemp.EOF Then
                    mcolCardPayMode.Add varTemp, "K" & j
                    .AddItem varTemp(1): .ItemData(.NewIndex) = -1
                    j = j + 1
                End If
            End If
        Next
    End With
    
    If blnPrepay Then
        cboPayMode.AddItem "预交金"
        If mty_Para.bln优先使用预交 Then
            cboPayMode.ListIndex = cboPayMode.NewIndex
        End If
    End If
    
    If blnInsure Then
        rsTemp.Filter = "性质 = 3"
        If rsTemp.EOF Then
            mstrInsure = ""
            MsgBox "不能加载医保结算方式,请检查!", vbInformation, gstrSysName
        Else
            cboPayMode.AddItem Nvl(rsTemp!名称)
            mstrInsure = Nvl(rsTemp!名称)
            If Not mty_Para.bln优先使用预交 Or blnPrepay = False Then
                cboPayMode.ListIndex = cboPayMode.NewIndex
            End If
            If (mintInsure <> 0 And MCPAR.不收病历费) And cboPayMode.Text = "个人帐户" And cboPayMode.Visible Then
                chkBook.Enabled = False
                chkBook.Value = 0
            Else
                chkBook.Enabled = True
            End If
        End If
    End If
    
    If cboPayMode.ListCount > 0 And cboPayMode.ListIndex = -1 Then
        cboPayMode.ListIndex = 0
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub LoadRegPlans(ByVal blnCache As Boolean, Optional ByVal blnAutoUpdate As Boolean = True)
    Dim strTime As String, strState As String, strWhere As String
    Dim strSql As String, strIF As String
    Dim i As Integer, k As Integer
    Dim DateThis As Date, strZero As String
    Dim str挂号安排 As String, str挂号汇总计划 As String
    Dim str挂号安排计划 As String, str挂号汇总安排 As String
    Dim str排序         As String
    Dim strFilter As String
    On Error GoTo errH
    
    str排序 = "号别,科室,项目,已挂"
    
    If Not blnCache Then
        strSql = "Zl_挂号安排_Autoupdate"
        gobjDatabase.ExecuteProcedure strSql, Me.Caption
    End If
    
    If Not blnCache And blnAutoUpdate Then
    
        If gstrDeptIDs <> "" Then strIF = " And Instr(','||[4]||',',','||P.科室ID||',')>0"
        
        '按输入的号别过滤：仅号别输入过程中才过滤,这时的ActiveControl一定是txtArrangeNo
        If Trim(txtArrangeNO.Text) <> "" And Trim(txtArrangeNO.Text) <> "+" And ActiveControl Is txtArrangeNO Then
            If IsNumeric(Trim(txtArrangeNO.Text)) Then
                strIF = strIF & " And P.号码 Like [2]"
            Else
                strIF = strIF & " And (zlSpellCode(P.医生姓名) Like [2] or B.简码 Like [2])"
            End If
        End If
        
         
        str挂号安排 = "" & _
            "            Select A.ID, A.号码, A.号类, A.科室id, A.项目id, A.医生id, A.医生姓名, A.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
            "                   A.周四 , A.周五, A.周六, A.分诊方式,a.开始时间,a.终止时间, A.序号控制, B.限号数, B.限约数,a.停用日期 " & vbNewLine & _
            "            From 挂号安排 A, 挂号安排限制 B " & vbNewLine & _
            "            Where a.停用日期 Is Null And [5] Between Nvl(a.开始时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
            "                 Nvl(a.终止时间, To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
            "                  And a.ID = B.安排id(+) " & IIf(mblnAppointment, " And [5] <= Trunc(Sysdate)+Nvl(A.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") ", "") & _
            "                  And Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" & vbNewLine & _
            "                  And Decode(To_Char([5], 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null) Is Not Null"
        '挂号安排 限号数限约数 挂号安排限制中获取
        str挂号汇总安排 = str挂号安排 & " And Not Exists (Select 1 From 挂号安排计划 Where 安排id = a.Id) "
        '挂号安排计划 限号数限约数 挂号计划限制中获取
        str挂号汇总计划 = " Union All " & _
            "            Select C.ID, A.号码, C.号类, C.科室id, A.项目id, A.医生id, A.医生姓名, C.病案必须, A. 周日, A.周一, A.周二, A.周三, " & _
            "                   A.周四 , A.周五, A.周六, A.分诊方式,a.生效时间,a.失效时间, A.序号控制, B.限号数, B.限约数,C.停用日期 " & vbNewLine & _
            "            From 挂号安排计划 A, 挂号计划限制 B,挂号安排 C " & vbNewLine & _
            "            Where c.停用日期 Is Null And [5] Between Nvl(a.生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
            "                 a.失效时间 And a.审核时间 Is Not Null And " & _
            "           a.生效时间 = (Select Max(生效时间)" & vbNewLine & _
            "                           From 挂号安排计划" & vbNewLine & _
            "                           Where 安排id = a.安排id And [5] Between" & vbNewLine & _
            "                           Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And 失效时间 And" & vbNewLine & _
            "                           审核时间 Is Not Null)" & _
            "                  And a.ID = B.计划id(+) And a.安排id = c.Id " & IIf(mblnAppointment, " And Trunc(Sysdate)+Nvl(C.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5] ", "") & _
            "                  And Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) = B.限制项目(+)" & vbNewLine & _
            "                  And Decode(To_Char([5], 'D'), '1', a.周日, '2', a.周一, '3', a.周二, '4', a.周三, '5', a.周四, '6', a.周五, '7',a.周六, Null) Is Not Null"
        
        If mblnAppointment Then
            DateThis = Format(dtpDate, "yyyy-mm-dd hh:mm:ss")
        Else
            DateThis = gobjDatabase.Currentdate
        End If
        '取对应日期安排的时间段
        strSql = "Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)"
        
        '该部分语句取现在所对应的时间段
        strTime = _
            "Select 时间段 From 时间段 Where 号类 Is Null And 站点 Is Null And " & _
            "    ('3000-01-10 '||To_Char([5],'HH24:MI:SS') Between" & _
            "               Decode(Sign(开始时间-终止时间),1,'3000-01-09 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS'),'3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS'))" & _
            "               And '3000-01-10 '||To_Char(终止时间,'HH24:MI:SS'))" & _
            " Or" & _
            " ('3000-01-10 '||To_Char([5],'HH24:MI:SS')  Between" & _
            "   '3000-01-10 '||To_Char(Nvl(提前时间,开始时间),'HH24:MI:SS') And" & _
            "     Decode(Sign(开始时间-终止时间),1,'3000-01-11 '||To_Char(终止时间,'HH24:MI:SS'),'3000-01-10 '||To_Char(终止时间,'HH24:MI:SS')))"
            
        '该部分语句当时读取各种安排的挂号情况
        strState = _
        "   Select A.ID as 安排ID,B.已挂数,B.已约数" & _
        "   From (" & str挂号汇总安排 & str挂号汇总计划 & ") A,病人挂号汇总 B" & _
        "   Where A.科室ID = B.科室ID And A.项目ID = B.项目ID" & _
        "               And Nvl(A.医生ID,0)=Nvl(B.医生ID,0) " & _
        "               And Nvl(A.医生姓名,'医生')=Nvl(B.医生姓名,'医生') " & _
        "               And (A.号码=B.号码 or B.号码 is Null )  And B.日期=[6]"
        
        
        If mblnAppointment Then
            str挂号安排计划 = " " & _
                "             Select A.ID,A.ID as 计划ID, A.安排id, A.号码, A.项目id, A.安排人, A.安排时间, A. 周日, A.周一, A.周二, A.周三, A.周四, A.周五," & _
                "                    A.周六 , A.分诊方式, A.序号控制, B.限号数, B.限约数, A.生效时间, A.失效时间 ,A.医生姓名,A.医生ID " & _
                "             From 挂号安排计划 A, 挂号计划限制 B," & vbNewLine & _
                "                  (" & vbNewLine & _
                "                      Select Max(生效时间) As 生效时间, 安排id" & _
                "                      From 挂号安排计划 " & vbNewLine & _
                "                      Where 审核时间 Is Not Null And  [5] Between Nvl(生效时间, To_Date('1900-01-01', 'YYYY-MM-DD')) And " & _
                "                          失效时间  " & vbNewLine & _
                "                       Group By 安排id" & vbNewLine & _
                "                   ) C" & _
                "             Where A.审核时间 Is Not Null And ([5] Between  A.生效时间 + 0 And A.失效时间)" & _
                "                   And A.ID = B.计划id(+) And " & vbNewLine & _
                "                   Decode(To_Char([5], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6'," & _
                "                  '周五', '7', '周六', Null) = B.限制项目(+) And A.生效时间 = C.生效时间 And A.安排id = C.安排id"

            strSql = _
            " Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
            "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班 " & _
            " From (" & str挂号安排 & ") P" & _
            " Where    Not Exists(Select 1 From 挂号安排计划 where 安排ID=P.id And ([5] BETWEEN 生效时间 + 0 and 失效时间)  And 审核时间 is not NULL  ) " & _
            "          And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=P.ID and [5] between 开始停止时间 and 结束停止时间 )" & _
            " Union ALL " & _
            " Select   C.ID,P.计划ID,C.号码,C.号类,C.科室ID,P.项目ID," & _
            "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(C.病案必须,0) as 病案必须," & _
            "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
            "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL)  as 排班 " & _
            " From (" & str挂号安排计划 & ") P, 挂号安排 C" & _
            " Where P.安排ID=C.ID  And C.停用日期 Is  NULL  And Trunc(Sysdate)+Nvl(C.预约天数," & IIf(mintSysAppLimit = 0, 1, mintSysAppLimit) & ") >= [5]  " & _
            "           And Not Exists(Select 1 From 挂号安排停用状态 Where 安排ID=C.ID and [5] between 开始停止时间 and 结束停止时间 )"
            strSql = "(" & strSql & ") P"
        Else
            strSql = _
                        " (Select P.ID,0 as 计划ID,P.号码 ,P.号类,P.科室ID,P.项目ID," & _
                        "       P.医生ID,P.医生姓名,P.限号数,P.限约数,Nvl(P.病案必须,0) as 病案必须," & _
                        "       P.周日,P.周一 ,P.周二 ,P.周三 ,P.周四 ,P.周五 ,P.周六,P.分诊方式,P.序号控制," & _
                        "       Decode(To_Char([5],'D'),'1',P.周日,'2',P.周一,'3',P.周二,'4',P.周三,'5',P.周四,'6',P.周五,'7',P.周六,NULL) as 排班 " & _
                        " From (" & str挂号安排 & ") P "
            strSql = strSql & vbNewLine & "  ) P"
        End If
        
        strSql = _
                    "Select Distinct " & _
                    "       P.ID,p.计划ID,P.号码 as 号别,P.号类,P.科室ID,B.名称 As 科室,P.项目ID,C.名称 As 项目," & _
                    "       P.医生ID,P.医生姓名 as 医生,Nvl(A.已挂数,0) as 已挂,Nvl(A.已约数,0) as 已约," & _
                    "       P.限号数 as 限号,P.限约数 as 限约,Nvl(P.病案必须,0) as 病案,Nvl(C.项目特性,0) as 急诊," & _
                    "       P.周日 as 日,P.周一 as 一,P.周二 as 二,P.周三 as 三,P.周四 as 四,P.周五 as 五,P.周六 as 六," & _
                    "       Decode(P.分诊方式,1,'指定',2,'动态',3,'平均',NULL) as 分诊,P.序号控制,P.排班" & _
                    " From " & strSql & "," & vbCrLf & _
                    "           (" & strState & ") A,部门表 B,收费项目目录 C" & _
                    " Where P.ID=A.安排ID(+) And Nvl(B.撤档时间,To_Date('3000-01-01','YYYY-MM-DD')) > Sysdate And P.科室ID=B.ID And P.项目ID=C.ID" & strIF & strZero & _
                    "           And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & strWhere & _
                    "           And (Nvl(P.医生ID,0)=0 Or Exists(Select 1 From 人员表 Q Where P.医生ID=Q.ID And (Q.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or Q.撤档时间 Is Null)))" & _
                    " Order by " & str排序
                    
        Set mrsPlan = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, _
                UserInfo.姓名, Trim(txtArrangeNO.Text) & "%", "", gstrDeptIDs, DateThis, CDate(Format(DateThis, "yyyy-MM-dd")), CDate(Format(DateThis, "yyyy-MM-dd")) + 1 - 1 / 24 / 60 / 60)
    Else
        '缓存从筛选
        If mrsPlan Is Nothing Then
            LoadRegPlans (False)
            Exit Sub
        End If
        If txtArrangeNO.Text <> "" Or cboDeptFilter.Text <> "" Or cboDoctorFilter.Text <> "" Then
            If txtArrangeNO.Text <> "" And mblnFilterChange = False Then
                strFilter = "号别 like '" & txtArrangeNO.Text & "*'"
            End If
            
            If Trim(cboDeptFilter.Text) <> "" Then
                If strFilter <> "" Then
                    If InStr(cboDeptFilter.Text, "-") > 0 Then
                        strFilter = strFilter & " And 科室 = '" & Split(cboDeptFilter.Text, "-")(1) & "'"
                    Else
                        strFilter = strFilter & " And 科室 = '" & cboDeptFilter.Text & "'"
                    End If
                Else
                    If InStr(cboDeptFilter.Text, "-") > 0 Then
                        strFilter = "科室 = '" & Split(cboDeptFilter.Text, "-")(1) & "'"
                    Else
                        strFilter = "科室 = '" & cboDeptFilter.Text & "'"
                    End If
                End If
            Else
                If mblnFilterChange Then strFilter = ""
            End If
            If Trim(cboDoctorFilter.Text) <> "" Then
                If strFilter <> "" Then
                    strFilter = strFilter & " And 医生 = '" & cboDoctorFilter.Text & "'"
                Else
                    strFilter = "医生 = '" & cboDoctorFilter.Text & "'"
                End If
            Else
                If mblnFilterChange And Trim(cboDeptFilter.Text) = "" Then strFilter = ""
            End If
            mrsPlan.Filter = strFilter
        Else
            LoadRegPlans (False)
            Exit Sub
        End If
        If mrsPlan.RecordCount <> 0 Then
            mrsPlan.MoveFirst
        Else
            vsfArrange.Clear 1
            vsfArrange.Rows = 2
            Exit Sub
        End If
    End If
    If mrsPlan.RecordCount = 0 And mblnAppointment Then
        vsfArrange.Clear 1
        If mblnInit Then MsgBox "当前没有可用的挂号安排，请在挂号安排管理中设置后重试！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsfArrange
        .Redraw = flexRDNone
        If Not mrsPlan.EOF Then
            mblnChangeByCode = True
            .ToolTipText = "共 " & mrsPlan.RecordCount & " 条安排"
            .Clear 1
            .Rows = 2
            .Rows = mrsPlan.RecordCount + 1
            mrsPlan.MoveFirst
            For i = 1 To mrsPlan.RecordCount
                .RowData(i) = Nvl(mrsPlan!科室ID)
                .TextMatrix(i, .ColIndex("IDS")) = mrsPlan!ID & "," & mrsPlan!项目ID & "," & IIf(IsNull(mrsPlan!医生ID), 0, mrsPlan!医生ID)
                .Cell(flexcpData, i, .ColIndex("IDS")) = mrsPlan!ID & "," & Val(Nvl(mrsPlan!计划ID))
                .TextMatrix(i, .ColIndex("号类")) = IIf(IsNull(mrsPlan!号类), "", mrsPlan!号类)
                .TextMatrix(i, .ColIndex("号别")) = mrsPlan!号别
                .TextMatrix(i, .ColIndex("科室")) = mrsPlan!科室
                .TextMatrix(i, .ColIndex("项目")) = mrsPlan!项目
                .Cell(flexcpData, i, .ColIndex("项目")) = Val(Nvl(mrsPlan!急诊))
                .TextMatrix(i, .ColIndex("医生")) = Nvl(mrsPlan!医生)
                .TextMatrix(i, .ColIndex("已约")) = Nvl(mrsPlan!已约)
                .TextMatrix(i, .ColIndex("限约")) = Nvl(mrsPlan!限约)
                
                .TextMatrix(i, .ColIndex("已挂")) = Nvl(mrsPlan!已挂)
                .TextMatrix(i, .ColIndex("限号")) = Nvl(mrsPlan!限号)
                .TextMatrix(i, .ColIndex("日")) = Left(Nvl(mrsPlan!日), 1)
                .Cell(flexcpData, i, .ColIndex("日")) = Nvl(mrsPlan!日)
                .TextMatrix(i, .ColIndex("一")) = Left(Nvl(mrsPlan!一), 1)
                .Cell(flexcpData, i, .ColIndex("一")) = Nvl(mrsPlan!一)
                .TextMatrix(i, .ColIndex("二")) = Left(Nvl(mrsPlan!二), 1)
                .Cell(flexcpData, i, .ColIndex("二")) = Nvl(mrsPlan!二)
                .TextMatrix(i, .ColIndex("三")) = Left(Nvl(mrsPlan!三), 1)
                .Cell(flexcpData, i, .ColIndex("三")) = Nvl(mrsPlan!三)
                .TextMatrix(i, .ColIndex("四")) = Left(Nvl(mrsPlan!四), 1)
                .Cell(flexcpData, i, .ColIndex("四")) = Nvl(mrsPlan!四)
                .TextMatrix(i, .ColIndex("五")) = Left(Nvl(mrsPlan!五), 1)
                .Cell(flexcpData, i, .ColIndex("五")) = Nvl(mrsPlan!五)
                .TextMatrix(i, .ColIndex("六")) = Left(Nvl(mrsPlan!六), 1)
                .Cell(flexcpData, i, .ColIndex("六")) = Nvl(mrsPlan!六)
                .TextMatrix(i, .ColIndex("分诊")) = Nvl(mrsPlan!分诊)
                .TextMatrix(i, .ColIndex("序号控制")) = IIf(mrsPlan!序号控制 = 1, "√", "")
                .Cell(flexcpData, i, .ColIndex("号别")) = ""
                If mrsPlan!号别 = txtArrangeNO.Text Then k = i
                'If k = 0 And mrsPlan!号别 = mstrPreNO And (mSortType = by号别 Or txt号别.Text = "") Then k = i
                mrsPlan.MoveNext
            Next
            mblnChangeByCode = False
        Else
            Set mrsPlan = Nothing
            .Clear 1
            .Rows = 2
            .ToolTipText = ""
        End If
        If k <> 0 Then
            .Row = k
            '53299
            mlngPreRow = k
            Call SetGridTop(k)
        Else
            .Row = .FixedRows
        End If
        Call SetvsfarrangeFiexBackColor
        'If mblnAppointment Then Call SetMshPlanFiexBackColor(False)
'        If blnCache = True Then mblnChangeByCode = True
'
'        Call vsfArrange_EnterCell
        mblnChangeByCode = False
        If txtArrangeNO.Visible And txtArrangeNO.Enabled And Not mblnFilterChange Then txtArrangeNO.SetFocus
'        If mrsPlan.RecordCount = 1 Then gobjCommFun.PressKeyEx vbKeyTab
        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub SetvsfarrangeFiexBackColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置相关固定列的背景色
    '参数:blnCurDate-是否当前日期列,否则就是预约日期列
    '编制:刘兴洪
    '日期:2010-02-04 14:39:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim PreRedaw As RedrawSettings, i As Long, strSql As String, strNow As String
    Dim strKey As String, rsTmp As ADODB.Recordset, strColor As String
    With vsfArrange
         .Redraw = flexRDNone
        strSql = "Select 时间段,开始时间,提前时间,提前颜色 From 时间段"
        Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
        strNow = Format(gobjDatabase.Currentdate, "HH:MM:SS")
        strKey = zlGet当前星期几
        For i = 1 To .Rows - 1
            rsTmp.Filter = "时间段='" & .Cell(flexcpData, i, .ColIndex(strKey)) & "'"
            If Not rsTmp.EOF Then
                If Not IsNull(rsTmp!提前时间) Then
                    strColor = Nvl(rsTmp!提前颜色, "0")
                    If strNow < Format(Nvl(rsTmp!开始时间), "HH:MM:SS") Then
                        .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = strColor
                    End If
                End If
            End If
        Next i
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub GetActiveView()
    '得到当前挂号业务  采取那种类型的流程
    Dim str号码         As String
    Dim dtCurDate            As Date
    Dim blnTimeSlice As Boolean '是否分时段
    Dim blnStartSn As Boolean '是否启用了序号
    
    On Error GoTo errH
    
    
    With vsfArrange
        If .Row > 0 And .ColIndex("号别") >= 0 Then
            str号码 = .TextMatrix(.Row, .ColIndex("号别"))
            blnStartSn = Trim(.TextMatrix(.Row, .ColIndex("序号控制"))) <> ""
        End If
    End With
    
    If mblnAppointment Then
        dtCurDate = dtpDate.Value
    Else
        dtCurDate = gobjDatabase.Currentdate
    End If
    blnTimeSlice = mobjRegistData.zlRegisterIsStartTimeSlice_Tradition(str号码, dtCurDate)
    
    If blnTimeSlice Then
        mViewMode = IIf(blnStartSn, v_专家号分时段, V_普通号分时段)
    Else
        mViewMode = IIf(blnStartSn, v_专家号, V_普通号)
    End If
    
    Call ShowAndHideDetailTime  '显示或隐藏时间段列表
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
         Resume
    End If
    gobjComlib.SaveErrLog
End Sub
Private Sub ShowAndHideDetailTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示或隐藏时间段列表
    '编制:刘兴洪
    '日期:2017-12-19 17:26:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnShowDetailTime As Boolean
    
    On Error GoTo errHandle
    Select Case mViewMode
    Case v_专家号分时段: blnShowDetailTime = True
    Case V_普通号分时段: blnShowDetailTime = IIf(mblnAppointment, True, False)
    Case v_专家号: blnShowDetailTime = True
    Case Else   'V_普通号
        blnShowDetailTime = False
    End Select
      
    If Not blnShowDetailTime Then
        vsfArrange.Height = fraTime.Height - 660
        vsfDetailTime.Visible = False: Exit Sub
    End If

    vsfArrange.Height = fraTime.Height / 2 - 300
    With vsfDetailTime
        .Top = vsfArrange.Top + vsfArrange.Height + 60
        .Height = fraTime.Height - .Top - 90
        .Visible = True
    End With
    If vsfArrange.RowIsVisible(vsfArrange.Row) = False Then Call vsfArrange.ShowCell(vsfArrange.Row, vsfArrange.Col)
    Exit Sub
errHandle:
   If gobjComlib.ErrCenter() = 1 Then
         Resume
   End If
End Sub



Private Function InitTimePlan() As Boolean
    '**************************************
    '加载时段
    '返回时段是否加载成功或是否有分时段
    '**************************************
    Dim dtRegistDate         As Date
    Dim str号码          As String
    str号码 = ""
    With vsfArrange
        If .Row > 0 And .ColIndex("号别") >= 0 Then
            str号码 = .TextMatrix(.Row, .ColIndex("号别"))
        End If
    End With
    
    '获取日期 这里可能要调整
    If mblnAppointment Then
        dtRegistDate = Format(dtpDate, "yyyy-mm-dd")
    Else
        dtRegistDate = gobjDatabase.Currentdate
    End If
    
    If mobjRegistData.zlGetRegistTimes_Tradition(str号码, dtRegistDate, mrs时间段, False) = False Then Exit Function
    InitTimePlan = mrs时间段.EOF = False
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub LoadTimePlan()
    '***************************************
    '加载时间段
    '***************************************
    Dim i               As Integer
    Dim j               As Integer
    Dim blnPre          As Boolean
    Dim lngThis         As Long
    Dim lngMax          As Long
    Dim lngCurrSn       As Long
    Dim lngMaxSn        As Long '预约的最大使用号
    Dim strSql          As String
    Dim rs时段统计      As ADODB.Recordset
    Dim str时间点       As String
    Dim lng预约人数     As Long
    Dim lngTatol        As Long '用于分时段 最后重新计算行数
    Dim strMaxDate      As String  '用于分时段保存大预约时间
    Dim lngCols         As Long
    Dim lngRows         As Long
    Dim strData         As String
    Dim strDate         As String
    Dim blnHave         As Boolean
    Dim datMax          As Date
    Dim dtSysDate       As Date, dtTemp As Date
    Dim bln失约用于挂号 As Boolean
    Dim blnInserted     As Boolean
    Dim lng合作单位人数 As Long
    Dim blnFindSN      As Boolean '是否需要重新定位到上次号别的序号,用于刷新列表时,数据保持
    Dim lngFindSN      As Long '需要查找的序号
    Dim str号别 As String
    Dim rsUnitReg As ADODB.Recordset
    
    
    vsfDetailTime.Redraw = flexRDNone
    vsfDetailTime.Clear
    '***************************************
    '表格信息设置
    '***************************************
    lngMax = Val(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("限号")))
    
    '1.调整位置
    If lngMax > 1000 Then
        vsfDetailTime.FontWidth = 4
    Else
        vsfDetailTime.FontWidth = 0 '恢复缺省字体
    End If
    
    '***************************************
    '初始化时间段
    '***************************************
     If InitTimePlan() = False Then vsfDetailTime.Redraw = flexRDBuffered: Exit Sub
     
     dtSysDate = gobjDatabase.Currentdate
    '***************************************
    '初始化表格
    '***************************************
     
     If mrs时间段 Is Nothing Then vsfDetailTime.Redraw = True: Exit Sub
     
 
    '***************************************
    '序号填充
    '***************************************
    With vsfDetailTime
       .Rows = 1
       .Cols = 1
       .Clear
    End With
    lngCurrSn = -1
    Select Case mViewMode
    Case V_普通号分时段:
        If Not mblnAppointment Then Exit Sub
        str号别 = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        If Not mobjRegistData.zlGetRegistSNStartus_Tradition(str号别, CDate(Format(dtpDate.Value, "yyyy-MM-dd")), rs时段统计, , False) Then Set rs时段统计 = Nothing: Exit Sub
        
        blnHave = False
        str时间点 = ""
        With mrs时间段
            datMax = CDate("00:00:00"): mdatLast = CDate("00:00:00")
            lngRows = -1: lngCols = 0
            Do While Not .EOF
                If datMax < CDate(Nvl(!开始时间, "00:00:00")) Then datMax = CDate(!开始时间)
                If mdatLast < CDate(Nvl(!结束时间, "00:00:00")) Then mdatLast = CDate(!结束时间)
                '预约状态 只填充允许预约的时间段
                '挂号时不区分都填充
                lng预约人数 = 0
                rs时段统计.Filter = " 日期='" & Nvl(!开始时间, "_") & "'"
                If Not rs时段统计.EOF Then lng预约人数 = Val(Nvl(rs时段统计!预约数量))
                
                lng合作单位人数 = 0
                If mblnAppointment And mobjRegistData.zlRegisterIsStartUnitCtrol_Tradition Then
                    If mobjRegistData.zlGetRegistUnitCtrol_Tradition(str号别, CDate(Format(dtpDate.Value, "yyyy-MM-dd")), rsUnitReg, False) Then
                        If Not rsUnitReg Is Nothing And rsUnitReg.RecordCount <> 0 Then
                            rsUnitReg.Filter = "序号=" & Val(Nvl(!序号))
                            lng合作单位人数 = 0
                            If rsUnitReg.RecordCount > 0 Then
                               lng合作单位人数 = Val(Nvl(rsUnitReg!数量))
                            End If
                        End If
                    End If
                End If
                
                If Nvl(!限制数量, 0) <> 0 Then
                    If str时间点 <> Nvl(!时间点) Then
                        lngRows = lngRows + 1
                        str时间点 = Nvl(!时间点)
                        If lngRows > vsfDetailTime.Rows - 1 Then vsfDetailTime.Rows = vsfDetailTime.Rows + 1: lngCols = 0
                        If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                        vsfDetailTime.TextMatrix(lngRows, 0) = str时间点
                     End If
                    lngCols = lngCols + 1
                    If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                    lng预约人数 = Nvl(!限制数量, 0) - lng预约人数 - lng合作单位人数
                    strData = "预约" & IIf(lng预约人数 < 0, 0, lng预约人数) & "人" & vbCrLf & _
                                          !开始时间 & "-" & !结束时间
                    vsfDetailTime.TextMatrix(lngRows, lngCols) = strData
                    If lng预约人数 <= 0 Then
                         vsfDetailTime.Cell(flexcpForeColor, lngRows, lngCols) = vbGreen
                    End If
                      If Format(dtSysDate, "yyyy-mm-dd") = Format(dtpDate, "yyyy-mm-dd") Then
                            If Format(DateAdd("n", mty_Para.int预约限制时间, dtSysDate), "hh:mm:ss") > Format(!结束时间, "hh:mm:ss") Then
                              vsfDetailTime.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                            End If
                      End If
                 End If
                .MoveNext
            Loop
        End With
        Set rs时段统计 = Nothing
    Case v_专家号分时段:
     '*******************************
     '专家号分时段
     '每行以时间点区分
     '*******************************
regHD:
        blnInserted = False
        str时间点 = ""
        With mrs时间段
            lngRows = -1: lngCols = 0
            datMax = CDate("00:00:00")
            Do While Not .EOF
                If datMax < CDate(Nvl(!开始时间, "00:00:00")) Then datMax = CDate(!开始时间)
                '预约状态 只填充允许预约的时间段
                '挂号时不区分都填充
                If blnFindSN Then
                    If Val(Nvl(!序号)) = lngFindSN And lngFindSN > 0 Then
                          lngCurrSn = lngFindSN
                    End If
                End If
                If str时间点 <> Nvl(!时间点) Then
                    lngRows = lngRows + 1
                    str时间点 = Nvl(!时间点)
                    If lngRows > vsfDetailTime.Rows - 1 Then vsfDetailTime.Rows = vsfDetailTime.Rows + 1: lngCols = 0
                    If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                    vsfDetailTime.TextMatrix(lngRows, 0) = str时间点
                    vsfDetailTime.Cell(flexcpForeColor, lngRows, 0, lngRows, 0) = vsfArrange.Cell(flexcpForeColor, vsfArrange.Row, 0, vsfArrange.Row, 0)
                 End If
                 
                lngCols = lngCols + 1
                If lngCols > vsfDetailTime.Cols - 1 Then vsfDetailTime.Cols = vsfDetailTime.Cols + 1
                strData = !序号 & vbCrLf & !开始时间 & "-" & !结束时间
                vsfDetailTime.TextMatrix(lngRows, lngCols) = strData
                
                If Nvl(!实际开始时间) = "" Then
                    dtTemp = CDate(Format(dtSysDate, "yyyy-mm-dd") & " " & Format(!开始时间, "hh:mm:ss"))
                Else
                    dtTemp = CDate(!实际开始时间)
                End If
                If dtSysDate > dtTemp Then  '失效的，恢色字体显示
                    vsfDetailTime.Cell(flexcpFontUnderline, lngRows, lngCols) = True
                    vsfDetailTime.Cell(flexcpForeColor, lngRows, lngCols) = vbGrayText
                End If
            .MoveNext
          Loop
          If blnHave = False And vsfDetailTime.Rows = 1 And vsfDetailTime.Cols = 1 And mrs时间段.RecordCount > 0 Then blnHave = True: mrs时间段.MoveFirst: GoTo regHD
        End With
    End Select
    
    dtpTime.Tag = Format(datMax, "hh:mm:ss")
    '***************************************
    '序号表格状态设置
    '***************************************
    Call SetSnStyle(True)
    '***************************************
    '序号状态 填充
    '现在挂号状态需要填充的只有一种状态
    '***************************************
     If mViewMode = v_专家号分时段 Then
        If mblnAppointment Then              '预约或接收时的日期
            dtTemp = CDate(Format(dtpDate.Value, "yyyy-MM-dd"))
        Else
            dtTemp = dtSysDate
        End If
         
'        If mty_Para.bln失约用于挂号 Then
'           '专家号分时段时  失约的序号用于开放出来挂号
'           bln失约用于挂号 = True
'           Datsys = DateAdd("n", -1 * mty_Para.lng预约有效时间, dtSysDate)
'        End If
        
        str号别 = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        blnHave = mobjRegistData.zlGetRegistSNStartus_Tradition(str号别, dtTemp, mrsSNState)
        If blnHave Then blnHave = mrsSNState.RecordCount > 0
        
        If blnHave Then '存在挂号序号状态数据
            With vsfDetailTime
                For i = 0 To .Rows - 1
                   For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" And Not .Cell(flexcpData, i, j) Like "加*" Then
                        On Error Resume Next
                        .Row = i:  .Col = j
                        
                        On Error GoTo Hd
                        
                        lngFindSN = Val(Get时段(i, j, False))
                        
                        mrsSNState.Filter = "序号=" & lngFindSN
                        If mrsSNState.RecordCount > 0 Then
                            If lngCurrSn = lngFindSN Then lngCurrSn = -1
                                Select Case mrsSNState!状态
                                Case 1  '已挂
                                      If Nvl(mrsSNState!预约, "0") = "0" Then
                                        vsfDetailTime.Cell(flexcpForeColor, i, j) = vbRed
                                      Else
                                        vsfDetailTime.Cell(flexcpForeColor, i, j) = &HC000C0
                                      End If
                                      vsfDetailTime.Cell(flexcpFontStrikethru, i, j) = True
                                Case 2  '已约
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbGreen
                                If lngMaxSn < Val(Nvl(mrsSNState!序号)) Then
                                    lngMaxSn = Val(Nvl(mrsSNState!序号))
                                End If
                                Case 3  '已留
                                  vsfDetailTime.Cell(flexcpForeColor, i, j) = vbBlue
                                Case 4  '退号
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbGrayText
                                    vsfDetailTime.Cell(flexcpFontStrikethru, i, j) = True
                                Case 5  '锁号
                                    vsfDetailTime.Cell(flexcpForeColor, i, j) = vbRed
                                End Select
                            End If
                        End If
                   Next
                Next
            End With
         End If
 
     End If
     '还有可用序号的情况下，屏蔽加号栏
    If CheckAddAvailable = False Then
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 1 To vsfDetailTime.Cols - 1
                If vsfDetailTime.Cell(flexcpData, i, j) Like "加*" Then
                    vsfDetailTime.Cell(flexcpData, i, j) = ""
                    vsfDetailTime.TextMatrix(i, j) = ""
                End If
            Next j
        Next i
    End If
    If vsfDetailTime.Rows > 1 Then
       vsfDetailTime.Cell(flexcpFontBold, 0, 0, vsfDetailTime.Rows - 1, 0) = True
    End If
    
    dtTemp = Format(IIf(mblnAppointment, dtpDate.Value, dtSysDate), "yyyy-mm-dd") & " " & Format(Me.dtpTime.Tag, "hh:mm")
    If dtTemp < dtpTime.minDate Then dtpTime.minDate = dtTemp
    If dtTemp < dtSysDate Then dtTemp = dtSysDate: dtpTime.minDate = dtTemp
    dtpTime.Value = dtTemp
    vsfDetailTime.Redraw = True
    locateSnBy时段 lngCurrSn
    mblnChangeByCode = True
    txtArrangeNO.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
    mblnChangeByCode = False
    txtDept.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("科室"))
    cboDoctor.Clear
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生")) = "" Then
        cboDoctor.Locked = False
        cboDoctor.Enabled = True
        Call LoadDoctor(vsfArrange.RowData(vsfArrange.Row))
    Else
        cboDoctor.AddItem vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生"))
        cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(2))
        cboDoctor.ListIndex = cboDoctor.NewIndex
        cboDoctor.Locked = True
        cboDoctor.Enabled = False
    End If
    Call LoadFeeItem(Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(1)), chkBook.Value = 1, mstrPriceGrade)
    Call vsfDetailTime_DblClick
    Exit Sub
Hd:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub locateSnBy时段(Optional ByVal lngSN As Long = -1, _
    Optional bln强制定位 As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:定位到指定的时段
    '入参:lngSN:>0需要定位的序号上,-1:表示按规则取数
    '出参:bln强制定位-强制定位到指定的数据列上
    '编制:刘兴洪
    '日期:2013-12-07 13:01:55
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngRow As Long, lngCol As Long
    Dim blnFind  As Boolean, blnExit As Boolean, blnMaxSn As Boolean
    Dim lngLastRow As Long, lngLastCol As Long
    Dim dtDate As Date
    
    lngRow = 0: lngCol = 1
    On Error GoTo errH
'    vsfDetailTime.HighLight = flexHighlightAlways
    Select Case mViewMode
    Case V_普通号分时段:
        '****************************
         '普通号分时段 序号定位
         '****************************
         vsfDetailTime.Redraw = False
         blnMaxSn = True
          For i = 0 To vsfDetailTime.Rows - 1
            For j = 1 To vsfDetailTime.Cols - 1
                With vsfDetailTime
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                            If Val(Get时段(i, j, False)) > 0 Then
                                     blnFind = True
                                     lngRow = i: lngCol = j: Exit For
                            End If
                        End If
                        lngLastRow = i
                        lngLastCol = j
                    End If
                End With
            Next
            If blnFind Then Exit For
          Next
         If blnFind Then
           vsfDetailTime.Row = lngRow: vsfDetailTime.Col = lngCol
            If vsfDetailTime.Row > 1 Then
                If vsfDetailTime.RowIsVisible(vsfDetailTime.Row) = False Then
                     vsfDetailTime.TopRow = vsfDetailTime.Row - 1
                End If
            End If
        Else
            vsfDetailTime.Row = lngLastRow: vsfDetailTime.Col = lngLastCol
            If vsfDetailTime.Row > 1 Then
                If vsfDetailTime.RowIsVisible(vsfDetailTime.Row) = False Then
                     vsfDetailTime.TopRow = vsfDetailTime.Row - 1
                End If
            End If
           vsfDetailTime.HighLight = flexHighlightAlways
        End If
        
        dtDate = IIf(blnFind, CDate(Get时段(lngRow, lngCol, True)), CDate(mdatLast))
        If mblnAppointment Then
            dtDate = CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(dtDate, "HH:MM:SS"))
        Else
            dtDate = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd") & " " & Format(dtDate, "HH:MM:SS"))
        End If
        If dtDate < dtpDate.minDate Then dtpDate.minDate = Format(dtDate, "yyyy-mm-dd 00:00:00")
        dtpDate.Value = dtDate
        vsfDetailTime.Redraw = True
 
    Case v_专家号分时段:
        blnMaxSn = True
        With vsfDetailTime
            For i = 0 To .Rows - 1
                For j = 1 To .Cols - 1
                    If .TextMatrix(i, j) <> "" Then
                       
                         If .Cell(flexcpForeColor, i, j) <> vbRed _
                             And .Cell(flexcpForeColor, i, j) <> vbBlue _
                             And .Cell(flexcpForeColor, i, j) <> vbGrayText Then
                             
                            If blnMaxSn = True _
                                And .Cell(flexcpForeColor, i, j) <> vbGreen _
                                And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                                If Not mty_Para.bln随机序号选择 Or lngSN = -1 Then  '66788
                                    blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                    blnExit = True: Exit For
                                End If
                             End If
                             
                             If lngSN <> -1 Then
                                 If lngSN = Val(Get时段(i, j, False)) Then
                                    .Row = i: .Col = j
                                     blnFind = True
                                    lngRow = i: lngCol = j
                                    blnMaxSn = False
                                     dtpTime.Value = CDate(Get时段(i, j, True))
                                     blnExit = True: Exit For
                                 End If
                             End If
                         Else
                              blnMaxSn = True
                         End If
                    End If
                Next
                If blnExit Then Exit For '45768
            Next
        End With
        
        If blnFind And blnMaxSn = False Then
            mblnChangeByCode = True
            vsfDetailTime.Row = lngRow: vsfDetailTime.Col = lngCol
            mblnChangeByCode = False
'            vsfDetailTime.HighLight = flexHighlightAlways
        Else
            vsfDetailTime.Select 0, 0
            vsfDetailTime.HighLight = flexHighlightNever
        End If
        
        If blnFind = False And blnMaxSn Then
            dtDate = Format(CDate(gobjDatabase.Currentdate), "hh:mm:ss")
        Else
            dtDate = CDate(Get时段(lngRow, lngCol, True))
        End If
        
        If mblnAppointment Then
            dtDate = CDate(Format(dtpDate.Value, "yyyy-mm-dd") & " " & Format(dtDate, "HH:MM:SS"))
        Else
            dtDate = CDate(Format(gobjDatabase.Currentdate, "yyyy-mm-dd") & " " & Format(dtDate, "HH:MM:SS"))
        End If
        
        If dtDate < dtpTime.minDate Then dtpTime.minDate = Format(dtDate, "yyyy-mm-dd 00:00:00")
        dtpTime.Value = dtDate
    Case Else: Exit Sub
    End Select
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub
 

Private Sub vsfArrange_DblClick()
     If vsfArrange.MouseRow > 0 Then Call vsfArrange_KeyDown(13, 0)
'
'    mblnChangeByCode = True
'    If txtPatient.Text = "" Then
'        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
'    Else
'        Call vsfArrange_KeyDown(13, 0)
'    End If
'    mblnChangeByCode = False
End Sub

Private Sub vsfArrange_EnterCell()
    Dim i As Integer, j As Integer
    Dim blnPre As Boolean, lngThis As Long, lngMax As Long
    Dim dtThis As Date, lngCurrSn  As Long, lngMaxSn   As Long '预约的最大使用号
    Dim lngRow As Long, lngCol     As Long
    Dim blnChk As Boolean, sngTime As Single
    Dim varTemp As Variant
      
      
      
    '*****************************
    '获取使用那种流程处理挂号
    '******************************
    If mblnChangeByCode Then Exit Sub
    
    
    sngTime = Timer
    If Format(sngTime, "0.000") - Format(msngTime, "0.000") < 0.1 Then
        mblnChangeByCode = True
        If mlngRow <> 0 Then vsfArrange.Select mlngRow, 0
        mblnChangeByCode = False
        Exit Sub
    End If
    
      
     
    lbl急.Visible = Val(vsfArrange.Cell(flexcpData, vsfArrange.Row, vsfArrange.ColIndex("项目"))) = 1
   
    
    msngTime = Timer
    mlngRow = vsfArrange.Row
    
    
   
    Call GetActiveView  '重新设置当前操作的号别类别:专家分时段，普通分时段等
     
    If mViewMode = v_专家号分时段 Or mViewMode = V_普通号分时段 And mblnAppointment Then
       '*************************************************
       '如果存在分时段的情况 使用分时段的处理方法
       '*************************************************
       LoadTimePlan
       Call vsfDetailTime_AfterRowColChange(vsfDetailTime.Row, vsfDetailTime.Col, vsfDetailTime.Row, vsfDetailTime.Col)
       Call ReadRoom
       Exit Sub
    End If
    
    
    vsfDetailTime.Redraw = False
    vsfDetailTime.Clear

    lngMax = Val(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("限号"))) '挂将来的号不当成预约,因为已交费,应当成挂号

    If lngMax > 0 And vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("序号控制")) <> "" Then
        If lngMax = 0 Then GoTo regTab
        '1.调整位置
        If lngMax > 1000 Then
            vsfDetailTime.FontWidth = 4
        Else
            vsfDetailTime.FontWidth = 0 '恢复缺省字体
        End If

        If (lngMax \ SNCOLS) * SNCOLS = lngMax Then
            vsfDetailTime.Rows = lngMax \ SNCOLS
        Else
            vsfDetailTime.Rows = lngMax \ SNCOLS + 1
        End If
        'mblnNotClick = False
        vsfDetailTime.Cols = SNCOLS
        If Not vsfDetailTime.Visible Then
            vsfDetailTime.Visible = True
'            picSplit.Visible = True
        End If
                                
        '填充序号
        lngThis = 1
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 0 To vsfDetailTime.Cols - 1
                vsfDetailTime.TextMatrix(i, j) = lngThis
                lngThis = lngThis + 1
                If lngThis > lngMax Then Exit For
            Next
            If lngThis > lngMax Then Exit For
        Next
        
        
        If mobjRegistData.zlGetRegistSNStartus_Tradition(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")), dtThis, mrsSNState) Then
             lngMaxSn = 0
            For i = 0 To mrsSNState.RecordCount - 1
                 If mrsSNState!序号 <= lngMax Then
                     If (mrsSNState!序号 \ SNCOLS) * SNCOLS = mrsSNState!序号 Then
                        lngRow = (mrsSNState!序号 \ SNCOLS) - 1
                        lngRow = IIf(lngRow < 0, 0, lngRow)
                     Else
                         lngRow = (mrsSNState!序号 \ SNCOLS)
                     End If
                         lngCol = (mrsSNState!序号 - 1) Mod SNCOLS
                         lngCol = IIf(lngCol < 0, 0, lngCol)
                     Select Case mrsSNState!状态
                         Case 1  '已挂
                            If Nvl(mrsSNState!预约, "0") = "0" Then
                               vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                               '用于序号定位最大的有效号后
                               If lngMaxSn < Val(Nvl(mrsSNState!序号)) Then
                                 lngMaxSn = Val(Nvl(mrsSNState!序号))
                               End If
                            Else
                               '预约接收
                               vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = &HC000C0
                            End If
                         Case 2  '已约
                             vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbGreen
                             
                            
                         Case 3  '已留
                             vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
                         Case 4  '退号
                             vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbGrayText
                             vsfDetailTime.Cell(flexcpFontStrikethru, lngRow, lngCol) = True
                         Case 5  '锁号
                             vsfDetailTime.Cell(flexcpForeColor, lngRow, lngCol) = vbRed
                     End Select
                 End If
                 mrsSNState.MoveNext
             Next
             lngCurrSn = GetCurrSN(lngMaxSn)
        End If
    Else
regTab:
        Set mrsSNState = Nothing
        vsfDetailTime.Visible = False
        Call Form_Resize
    End If
    
    vsfDetailTime.Redraw = True
    SetSnStyle
    vsfDetailTime.Select 0, 0
    Call LocateSN(lngCurrSn)
    
    If vsfDetailTime.Row <= vsfDetailTime.Rows - 1 And vsfDetailTime.Col <= vsfDetailTime.Cols - 1 And vsfDetailTime.Cell(flexcpForeColor, vsfDetailTime.Row, vsfDetailTime.Col) = vbBlack Then
        vsfDetailTime.Cell(flexcpBackColor, vsfDetailTime.Row, vsfDetailTime.Col) = &H8000000D
    End If
    txtDept.Text = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("科室"))
    Call ReadRoom
    cboDoctor.Clear
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生")) = "" Then
        cboDoctor.Locked = False
        cboDoctor.Enabled = True
        Call LoadDoctor(vsfArrange.RowData(vsfArrange.Row))
    Else
        cboDoctor.AddItem vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("医生"))
        cboDoctor.ItemData(cboDoctor.NewIndex) = Val(Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")(2))
        cboDoctor.ListIndex = cboDoctor.NewIndex
        cboDoctor.Locked = True
        cboDoctor.Enabled = False
    End If
    If vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")) = "" Then Exit Sub
    varTemp = Split(vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("IDS")), ",")
    Call LoadFeeItem(Val(varTemp(1)), chkBook.Value = 1, mstrPriceGrade)
    Call vsfDetailTime_DblClick
End Sub

Private Sub ReadRoom()
    Dim blnBusy As Boolean, strSql As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    blnBusy = Val(gobjDatabase.GetPara("诊室忙时允许分诊", glngSys, 1113, 0)) = 1
    strSql = _
        " Select b.编码, b.名称, b.位置" & vbNewLine & _
        " From 挂号安排诊室 a, 门诊诊室 b, 挂号安排 c" & vbNewLine & _
        " Where a.门诊诊室 = b.名称 And a.号表id = c.Id And c.号码 = [1] " & _
        IIf(blnBusy, " ", " And b.缺省标志=0 ")
    Set rsTmp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")))
    cboRoom.Clear
    Do While Not rsTmp.EOF
        cboRoom.AddItem Nvl(rsTmp!名称)
        cboRoom.ItemData(cboRoom.NewIndex) = Nvl(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    If cboRoom.ListCount = 1 Then cboRoom.ListIndex = 0
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub LoadFeeItem(ByVal lngItemID As Long, ByVal blnBook As Boolean, ByVal strPriceGrade As String)
    Dim strSql As String, i As Integer, dblTotal As Double, lng病人ID As Long
    Dim strFee As String, str附加项目ID As String, rsFeeTmp As ADODB.Recordset
    Dim cur应收 As Currency, cur实收 As Currency
    Dim j As Integer
    On Error GoTo errH
    
    ReadRegistPrice lngItemID, blnBook, False, txtFeeType.Text, mrsItems, mrsInComes, , , , 1, _
        Val(vsfArrange.RowData(vsfArrange.Row)), strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    If mintInsure <> 0 Then
        If MCPAR.挂号检查项目 = True Then
            If gclsInsure.CheckItem(mintInsure, 2, 0, mrsItems) = False Then
                MsgBox "医保病人收费项目检查失败，不能继续挂号！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    If Not mrsInfo Is Nothing Then
        If mrsInfo.State = 1 Then
            If Not mrsInfo.EOF Then lng病人ID = Val(Nvl(mrsInfo!病人ID))
        End If
    End If
    ReadRegistPrice lngItemID, blnBook, False, txtFeeType.Text, mrsItems, mrsInComes, lng病人ID, _
        mintInsure, vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别")), IIf(mblnAppointment, 1, 0), , strPriceGrade, IIf(mblnAppointment, Format(dtpDate, "yyyy-mm-dd") & " 23:59:59", "")
    
    vsfMoney.Clear 1
    vsfMoney.Rows = 2
    lblTotal.Caption = Format(0, "0.00")
    txtPayMoney.Text = Format(0, "0.00")
    dblTotal = 0
    If mrsItems.RecordCount = 0 Then Exit Sub
    mrsItems.MoveFirst
    Do While Not mrsItems.EOF
        With vsfMoney
            .RowData(.Rows - 1) = Nvl(mrsItems!项目ID)
            .TextMatrix(.Rows - 1, .ColIndex("项目")) = Nvl(mrsItems!项目名称)
            mrsInComes.Filter = "项目ID=" & mrsItems!项目ID
            cur应收 = 0: cur实收 = 0
            For j = 1 To mrsInComes.RecordCount
                cur应收 = cur应收 + mrsInComes!应收
                cur实收 = cur实收 + mrsInComes!实收
                mrsInComes.MoveNext
            Next j
            .TextMatrix(.Rows - 1, .ColIndex("应收金额")) = Format(cur应收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("实收金额")) = Format(cur实收, "0.00")
            .TextMatrix(.Rows - 1, .ColIndex("性质")) = Nvl(mrsItems!性质)
            .Rows = .Rows + 1
        End With
        mrsItems.MoveNext
    Loop
    If vsfMoney.Rows > 2 Then vsfMoney.Rows = vsfMoney.Rows - 1
    For i = 1 To vsfMoney.Rows - 1
        dblTotal = dblTotal + Val(vsfMoney.TextMatrix(i, vsfMoney.ColIndex("实收金额")))
    Next i
    lblTotal.Caption = Format(dblTotal, "0.00")
    txtPayMoney.Text = Format(dblTotal, "0.00")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Sub LoadDoctor(ByVal lng科室id As Long)
'功能：根据科室读取并绑定医生下拉列表
    Dim strSql As String
        
    On Error GoTo errH
    If mrsDoctor Is Nothing Then Call GetAll医生
    If mrsDoctor.State = 1 Then
        mrsDoctor.Filter = "部门id=" & lng科室id
        
        Do While Not mrsDoctor.EOF
            cboDoctor.AddItem IIf(IsNull(mrsDoctor!简码), "", mrsDoctor!简码 & "-") & mrsDoctor!姓名
            cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
            mrsDoctor.MoveNext
        Loop
        cboDoctor.ListIndex = -1
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub GetAll医生()
    Dim strSql As String
    On Error GoTo errH
    
    strSql = "Select a.Id, a.姓名, Upper(a.简码) As 简码,b.部门id,a.编号" & _
            " From 人员表 a, 部门人员 b, 人员性质说明 c" & _
            " Where a.Id = b.人员id And a.Id = c.人员id And c.人员性质 = [1] And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) " & _
            " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
            " Order By a.简码 Desc"
    Set mrsDoctor = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, "医生")
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub


Private Sub LocateSN(lngCurrSn As Long)
'功能:定位到指定序号上
'     如果不是在输号别或序号,则序号表获得焦点
    Dim lngRow          As Long
    Dim i               As Long
    Dim j               As Long
    Dim blnHave         As Boolean
    If lngCurrSn = 0 Then Exit Sub
   
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        '************************************************
        '不分时段 序号定位还是按照以前的方式
        '************************************************
        If (lngCurrSn \ SNCOLS) * SNCOLS = lngCurrSn Then
            lngRow = (lngCurrSn - 1) \ SNCOLS
        Else
            lngRow = (lngCurrSn \ SNCOLS)
        End If
        If Not vsfDetailTime.RowIsVisible(lngRow) Then
            If lngRow >= 1 Then  '保留上一行可见
                vsfDetailTime.TopRow = lngRow - 1
            Else
                vsfDetailTime.TopRow = lngRow
            End If
        End If
        mblnChangeByCode = True
        vsfDetailTime.Select lngRow, (lngCurrSn - 1) Mod SNCOLS
        mblnChangeByCode = False
'        vsfDetailTime.Row = lngRow
'        vsfDetailTime.RowSel = vsfDetailTime.Row
'        vsfDetailTime.Col = (lngCurrSn - 1) Mod SNCOLS
'        vsfDetailTime.ColSel = vsfDetailTime.Col
     
    ElseIf mViewMode = v_专家号分时段 Then
        '*******************************************
        '专家号分时段 序号定位
        '*******************************************
        For i = 0 To vsfDetailTime.Rows - 1
            For j = 1 To vsfDetailTime.Cols - 1
               If vsfDetailTime.TextMatrix(i, j) <> "" Then
                    If lngCurrSn = Val(Get时段(i, j, False)) Then
                     If Not vsfDetailTime.RowIsVisible(i) Then
                        If lngRow >= 1 Then  '保留上一行可见
                             vsfDetailTime.TopRow = i - 1
                        Else
                             vsfDetailTime.TopRow = i
                        End If
                      End If
                      vsfDetailTime.Row = i
                      vsfDetailTime.Col = j
                     blnHave = True
                     Exit For
                    End If
                End If
            Next
            If blnHave Then Exit For
        Next
    End If
'    vsfDetailTime.HighLight = flexHighlightAlways
    If vsfDetailTime.Visible And vsfDetailTime.Enabled _
                And Not Me.ActiveControl Is txtArrangeNO _
                And Not Me.ActiveControl Is vsfArrange Then Call vsfDetailTime.SetFocus     '焦点在号别正在连续输入
End Sub

Private Function Get时段(ByVal lngRow As Long, ByVal lngCol As Long, Optional ByVal blnTime As Boolean = False, Optional ByVal blnLastTime As Boolean = False) As String
    '*****************************************************************
    '功能说明:在挂号专家号分时时 获取 序号,或者 开始时间
    '参数:  blntime 是否获取时间 是则获取时间  否则返回序号
    '*****************************************************************
    Dim strResult       As String, i As Long
    On Error GoTo errH
    If lngRow > vsfDetailTime.Rows - 1 Or lngCol > vsfDetailTime.Cols - 1 Then
        Exit Function
    End If
    If vsfDetailTime.TextMatrix(lngRow, lngCol) = "" Then
        Exit Function
    End If
    
    If blnTime Then
        i = IIf(blnLastTime = False, 0, 1)
        If InStr(vsfDetailTime.TextMatrix(lngRow, lngCol), "-") > 0 Then
            Get时段 = Split(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(1), "-")(i)
        Else
            If InStr(vsfDetailTime.TextMatrix(lngRow, lngCol), "以") = 0 Then Exit Function
            Get时段 = Split(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(1), "以")(i)
        End If
        Exit Function
    End If
    If mViewMode = v_专家号分时段 Then
       strResult = Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(0)
    ElseIf mViewMode = V_普通号分时段 Then
       strResult = Replace(Replace(Split(vsfDetailTime.TextMatrix(lngRow, lngCol), vbCrLf)(0), "预约", ""), "人数", "")
    End If
    Get时段 = strResult
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Sub SetSnStyle(Optional ByVal bln分时段 As Boolean = False)
'****************************************
'对表格样式进行设置
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
    On Error GoTo errH
    Select Case bln分时段
    Case False:
        With vsfDetailTime
            
            .FixedCols = 0
            lngWidth = 570
            lngHeight = 450
            For i = 0 To vsfDetailTime.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            For i = 0 To vsfDetailTime.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
            
        End With
    
    Case True:
        With vsfDetailTime
             If .Cols <= 1 Then Exit Sub
             .FixedCols = 1
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
            lngHeight = 550
            For i = 1 To vsfDetailTime.Cols - 1
                .ColWidth(i) = lngWidth
                .ColAlignment(i) = 4
            Next
            .ColAlignment(0) = 3
            .ColWidth(0) = lngWidth
            For i = 0 To vsfDetailTime.Rows - 1
                 .RowHeight(i) = lngHeight
            Next
           If .Rows > 0 And .Cols > 0 Then
                .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
                .Cell(flexcpFontSize, 0, 1, .Rows - 1, .Cols - 1) = 9
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 18
           End If
        End With
    End Select
    If vsfDetailTime.Rows >= 1 And vsfDetailTime.Cols > 0 Then
       vsfDetailTime.Cell(flexcpFontBold, 0, 0, vsfDetailTime.Rows - 1, vsfDetailTime.Cols - 1) = True
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Sub

Private Function GetCurrSN(Optional ByVal lngCurMaxSN As Long = -1, Optional ByVal blnGetLapseNO As Boolean = False) As Long
'功能:获取当前号别的最大可用序号
'     全部都用完时返回0
'    blngetlapseNo:是否从无效号以后开始算
'     lngCurMaxSN-当明最大使用号
    Dim i           As Integer
    Dim j           As Integer
    Dim lngMaxSn    As Long
    Dim lngSN       As Long
    Dim intStart    As Integer
    Dim lngTmp      As Long
    Dim blnUnitReg  As Boolean
    Dim lngMaxLapse As Long '最大无效号码
    On Error GoTo errH
    If Not mrsSNState Is Nothing Or blnUnitReg Then
ReGet:
        mrsSNState.Filter = ""
        If mrsSNState.RecordCount > 0 Or blnUnitReg Then
            If lngCurMaxSN = -1 And mViewMode = v_专家号分时段 Then
                With vsfDetailTime
                    i = vsfDetailTime.Row
                    j = vsfDetailTime.Col
                    If .TextMatrix(i, j) <> "" Then
                        If .Cell(flexcpForeColor, i, j) <> vbRed And .Cell(flexcpForeColor, i, j) <> vbBlue And .Cell(flexcpForeColor, i, j) <> vbGreen And .Cell(flexcpForeColor, i, j) <> vbGrayText And .Cell(flexcpForeColor, i, j) <> &HC000C0 Then
                           lngTmp = Val(Get时段(i, j, False))
                           mrsSNState.Filter = "序号=" & lngTmp
                            If mrsSNState.RecordCount = 0 And lngTmp > lngMaxLapse Then
                                    GetCurrSN = lngTmp
                                    Exit Function
                            End If
                        End If
                    End If
                End With
            End If
            
            
           If lngCurMaxSN = -1 And mViewMode = v_专家号 Then
               lngTmp = 0
               mrsSNState.Filter = "预约=0 and 状态=1"
                Do While Not mrsSNState.EOF
                   If lngTmp < Val(mrsSNState!序号) Then lngTmp = Val(mrsSNState!序号)
                   mrsSNState.MoveNext
                Loop
                
                'mrsSNState.MoveFirst
                mrsSNState.Filter = 0
               If lngTmp <> 0 Then lngCurMaxSN = lngTmp
            End If
            
            intStart = IIf(mViewMode = v_专家号分时段 Or mViewMode = V_普通号分时段, 1, 0)
            For i = 0 To vsfDetailTime.Rows - 1
                For j = intStart To vsfDetailTime.Cols - 1
                    Select Case mViewMode
                    Case V_普通号, v_专家号:
                        lngSN = Val(vsfDetailTime.TextMatrix(i, j))
                    Case v_专家号分时段:
                        With vsfDetailTime
                            If .Cell(flexcpForeColor, i, j) = vbGrayText Or .Cell(flexcpForeColor, i, j) = &HC000C0 Then
                                lngSN = -1
                            Else
                               lngSN = IIf(Trim(.TextMatrix(i, j)) = "", -1, Val(Get时段(i, j, False)))
                               If lngSN < lngMaxLapse And mty_Para.bln随机序号选择 = False Then lngSN = -1
                            End If
                        End With
                    Case Else
                       Exit Function
                    End Select
                    If lngSN > -1 Then
                        mrsSNState.Filter = "序号=" & lngSN
                        If mrsSNState.RecordCount = 0 Then
                            lngMaxSn = lngSN
                            vsfDetailTime.Select i, j
                            Exit For
                        End If
                    End If
                    
                Next
                
                If lngMaxSn = lngSN Then Exit For
            Next
            If lngCurMaxSN > 0 And lngMaxSn = 0 Then
                '刘兴洪:???
                '主要是解决预约最大+1后,还有预约的情况,所以又从1开始检查是否有未选择的.
                '如:预约从5开始;到了7已经是最大号了,因此再从1开始取.
               ' lngCurMaxSN = -1: GoTo ReGet:
            End If
            GetCurrSN = lngMaxSn
        Else
            Select Case mViewMode
                Case v_专家号分时段:
                     vsfDetailTime.Redraw = False
                    For i = 0 To vsfDetailTime.Rows - 1
                        For j = 1 To vsfDetailTime.Cols - 1
                            If vsfDetailTime.Cell(flexcpForeColor, i, j) <> vbGrayText And vsfDetailTime.Cell(flexcpForeColor, i, j) <> &HC000C0 And vsfDetailTime.TextMatrix(i, j) <> "" Then
                                GetCurrSN = Val(Get时段(i, j, False))
                                vsfDetailTime.Redraw = True
                                Exit Function
                            End If
                        Next
                    Next
                    vsfDetailTime.Redraw = True
                Case Else:
                    GetCurrSN = 1
            End Select
        End If
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

 

Private Sub vsfArrange_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then

        If CheckNoValied(vsfArrange.Row) = False Then
             txtArrangeNO.Text = "": txtArrangeNO.SetFocus: Exit Sub
        End If
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        If txtArrangeNO.Visible And txtArrangeNO.Enabled Then txtArrangeNO.SetFocus
        If txtArrangeNO.Text = vsfArrange.Tag Then
            Call txtArrangeNo_Change
        Else
            txtArrangeNO.Text = vsfArrange.Tag
        End If


        Call vsfArrange_EnterCell
        DoEvents
        If cboDoctor.Enabled Then
            cboDoctor.SetFocus
        Else
            If cboRoom.Enabled Then
                cboRoom.SetFocus
            Else
                chkBook.SetFocus
            End If
        End If
    End If
End Sub

Private Sub vsfArrange_LeaveCell()
   ' Call SetMshPlanColor
End Sub

Private Sub vsfDetailTime_Click()
    Call vsfDetailTime_DblClick
End Sub

 

Private Sub vsfDetailTime_EnterCell()
    '***************************************
    '这里处理预留号
    '预留号处理情况为
    '专家号不分时段 以前的处理方式
    '专家号 分时段 新处理方式
    '普通号分时段 不允许预留
    '***************************************
    Dim lngSN As Long, strTime As String
    If mViewMode = V_普通号分时段 Or mViewMode = V_普通号 Then
'        cmdHold.Enabled = False
'        cmdHold.Caption = "预留(&L)"
        Exit Sub
    End If
    
    With vsfDetailTime
        If .Row <> -1 Then
             If .Cols > .Col And .Rows > .Row Then
                If .TextMatrix(.Row, .Col) = "" Then Exit Sub
             End If
        End If
    End With
'
'    cmdHold.Enabled = Not mrsSNState Is Nothing
'    cmdHold.Caption = "预留(&L)"
'
    If mrsSNState Is Nothing Then Exit Sub

    With vsfDetailTime
        If .Cols > .Col And .Rows > .Row Then
            Select Case mViewMode
            Case v_专家号:
                lngSN = Val(.TextMatrix(.Row, .Col))
                mrsSNState.Filter = "序号=" & lngSN
            Case v_专家号分时段:
                lngSN = Val(Get时段(.Row, .Col, False))
                mrsSNState.Filter = "序号=" & lngSN
            End Select
        End If
    End With
    If mrsSNState.RecordCount = 0 Then mrsSNState.Filter = 0: Exit Sub
    
    If Val(Nvl(mrsSNState!状态)) <> 3 Then mrsSNState.Filter = 0: Exit Sub
    
    If mrsSNState!操作员姓名 = UserInfo.姓名 Then
        '取消预留
       mrsSNState.Filter = 0
       'cmdHold.Caption = "取消预留(&L)"
       Exit Sub
    End If

    'cmdHold.Enabled = False
    If Not Me.ActiveControl Is vsfDetailTime Then mrsSNState.Filter = 0: Exit Sub
    
    Select Case mViewMode
        Case v_专家号:
            MsgBox lngSN & "号已被" & mrsSNState!操作员姓名 & "预留!无法选择.", vbInformation, gstrSysName
        Case v_专家号分时段:
            MsgBox lngSN & "号已被" & mrsSNState!操作员姓名 & "预留!无法选择.", vbInformation, gstrSysName
    End Select
    txtArrangeNo_KeyPress (13)

    mrsSNState.Filter = 0
End Sub

Private Sub vsfDetailTime_KeyDown(KeyCode As Integer, Shift As Integer)
     If mty_Para.bln随机序号选择 Then Exit Sub
     If KeyCode <> 13 Then KeyCode = 0
End Sub

Private Sub vsfDetailTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow > vsfDetailTime.Rows - 1 Or NewCol > vsfDetailTime.Cols - 1 Then Exit Sub

    
    If OldRow <= vsfDetailTime.Rows - 1 And OldCol <= vsfDetailTime.Cols - 1 Then
        vsfDetailTime.Cell(flexcpBackColor, OldRow, OldCol) = &H80000005
        If OldRow = 0 And OldCol = 0 And InStr(vsfDetailTime.TextMatrix(OldRow, OldCol), ":") > 0 Then
            vsfDetailTime.Cell(flexcpBackColor, OldRow, OldCol) = &H8000000F
        End If
    End If
    If NewRow <= vsfDetailTime.Rows - 1 And NewCol <= vsfDetailTime.Cols - 1 Then
        If vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) = vbBlack And vsfDetailTime.Cell(flexcpBackColor, NewRow, NewCol) <> -2147483633 Then
            vsfDetailTime.Cell(flexcpBackColor, NewRow, NewCol) = &H8000000D
        End If
    End If
End Sub

Private Sub vsfDetailTime_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    If mblnAppointment Then Exit Sub    '预约允许选择时间
    If mblnChangeByCode Then Exit Sub
    
    If (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Or mViewMode = v_专家号) And mty_Para.bln随机序号选择 = False _
        And vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then
        Cancel = True
        Exit Sub
    End If
    
    If vsfDetailTime.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
    If vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlack And vsfDetailTime.Cell(flexcpForeColor, NewRow, NewCol) <> vbBlue Then Cancel = True
    If Not CheckAddAvailable Then
        If vsfDetailTime.Cell(flexcpData, NewRow, NewCol) Like "加*" Then Cancel = True
    End If
End Sub

Private Sub vsfDetailTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call vsfDetailTime_DblClick
End Sub

Private Function CheckAddAvailable() As Boolean
'-----------------------------------------------------------------------------------------------------------------------
'功能:检查当前选择的号别加号是否可用
'返回:可用返回True,不可用返回False
'编制:刘尔旋
'日期:2014-01-15
'备注:
'-----------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim intTotal As Integer, intUse As Integer
    On Error GoTo errH
    intTotal = 0
    intUse = 0
    '只对分时段进行处理
    If mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段 Or mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        With vsfDetailTime
            For j = 1 To .Cols - 1
                For i = 0 To .Rows - 1
                    If .TextMatrix(i, j) <> "" And Not .Cell(flexcpData, i, j) Like "加*" Then
                        intTotal = intTotal + 1
                        If .Cell(flexcpForeColor, i, j) <> vbBlack Then
                            intUse = intUse + 1
                        End If
                    End If
                Next i
            Next j
        End With
        If intUse = intTotal Then CheckAddAvailable = True: Exit Function
        CheckAddAvailable = False
        Exit Function
    End If
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

Private Function zlGet当前星期几(Optional strDate As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当日是星期几
    '编制:刘兴洪
    '日期:2010-02-04 14:42:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, bln当前日期 As Boolean, strTemp As String
    On Error GoTo errH
    If strDate = "" Then
        strSql = "Select Decode(To_Char(Sysdate,'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六',NULL) as 星期  From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption)
    Else
        strSql = "Select Decode(To_Char([1],'D'),'1','日','2','一','3','二','4','三','5','四','6','五','7','六','') As 星期 From dual"
        Set rsTemp = gobjDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(strDate))
    End If
    
    If rsTemp.EOF = True Then
        Exit Function
    End If
    strTemp = Nvl(rsTemp!星期)
    zlGet当前星期几 = strTemp
    Exit Function
errH:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function
 

Private Sub vsfDetailTime_DblClick()
    Dim lngSN       As Long
    Dim datThis     As Date
    Dim strTmp      As String
    Dim strSql      As String
    Dim rsTmp       As ADODB.Recordset
    Dim dtSysDate As Date
    
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then
        '*************************************************
        '普通号和没有分时段的专家号 按照以前处理方法
        '*************************************************
        lngSN = Val(vsfDetailTime.TextMatrix(vsfDetailTime.Row, vsfDetailTime.Col))
        If mrsSNState Is Nothing Or lngSN <= 0 Then Exit Sub
        mrsSNState.Filter = "序号=" & lngSN
        If mrsSNState.RecordCount > 0 Then
            If Not (mrsSNState!状态 = 3 And mrsSNState!操作员姓名 = UserInfo.姓名) Then Exit Sub     '自已预留号，自已使用
        Else
            If vsfDetailTime.CellForeColor = &HC000C0 Then Exit Sub
        End If
        
        '满足条件，加载相关数据
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        txtArrangeNO.Text = vsfArrange.Tag
        txtSN.Text = lngSN
        mstrPreNO = txtArrangeNO.Text
        mlngPreRow = vsfArrange.Row
        vsfArrange.Tag = ""
        Exit Sub
    End If
    
    
    '*************************************************
    '分时段 按照新的方式来处理
    '*************************************************
    dtpTime.Enabled = False
    Select Case mViewMode
    Case V_普通号分时段:
        With vsfDetailTime
            If .CellForeColor = vbGrayText Then Exit Sub
            If .TextMatrix(.Row, .Col) = "" Then Exit Sub
            If Val(Get时段(.Row, .Col, False)) = 0 Then Exit Sub
            strTmp = Get时段(.Row, .Col, True)
        End With
        
        dtSysDate = gobjDatabase.Currentdate
        datThis = CDate(Format(IIf(mblnAppointment, dtpDate.Value, dtSysDate), "yyyy-mm-dd") & " " & Format(strTmp, "hh:mm"))
        If dtpTime.minDate >= datThis Then dtpTime.minDate = datThis
        If dtpTime.MaxDate <= datThis Then dtpTime.MaxDate = datThis
        dtpTime.Tag = strTmp
        dtpTime.Value = datThis
        
        If datThis < dtSysDate Then
           dtpTime.Value = dtSysDate
           dtpTime.minDate = dtSysDate
        End If
        
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        txtArrangeNO.Text = txtArrangeNO.Tag
        
        txtSN.Text = ""
        vsfArrange.Tag = ""
        mstrPreNO = txtArrangeNO.Text
        mlngPreRow = vsfArrange.Row
        If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
    
    Case v_专家号分时段:
        '**********************************************
        '如果序号为已挂或者已约的不允许选择
        '**********************************************
        With vsfDetailTime
            If .Row > .Rows - 1 Or .Col > .Cols - 1 Then Exit Sub
            If .TextMatrix(.Row, .Col) = "" Then Exit Sub
            If .CellForeColor = vbRed Or .CellForeColor = vbGreen Or .CellForeColor = vbGrayText Or .CellForeColor = &HC000C0 Then Exit Sub      '--And .CellForeColor <> vbBlue
            strTmp = Get时段(.Row, .Col, True)
        End With
        dtSysDate = gobjDatabase.Currentdate
        
        If strTmp <> "" Then
            datThis = CDate(Format(IIf(mblnAppointment, dtpDate.Value, dtSysDate), "yyyy-mm-dd") & " " & Format(strTmp, "hh:mm"))
        Else
            datThis = dtSysDate
        End If
        
        If dtpTime.minDate >= datThis Then dtpTime.minDate = datThis
        If dtpTime.MaxDate <= datThis Then dtpTime.MaxDate = datThis
        dtpTime.Value = datThis
        If datThis < dtSysDate Then
            dtpTime.Value = dtSysDate
            dtpTime.minDate = dtSysDate
        End If
        vsfArrange.Tag = vsfArrange.TextMatrix(vsfArrange.Row, vsfArrange.ColIndex("号别"))
        txtArrangeNO.Text = vsfArrange.Tag
        txtSN.Text = ""
        vsfArrange.Tag = ""
        mstrPreNO = txtArrangeNO.Text
        mlngPreRow = vsfArrange.Row
        If txtSN.Enabled And txtSN.Visible Then txtSN.SetFocus
    Case Else
        Exit Sub
    End Select
    
End Sub

Private Sub SetDefaultRegistTime()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置缺省的预约时间
    '编制:刘兴洪
    '日期:2012-03-12 15:49:38
    '问题:45509
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str安排 As String, str时间 As String
    Dim dtValue As Date, str号码 As String, strDate As String
    Dim str缺省时间 As String, strWeekName As String
    Dim strDefaultTime As String, strStartTime As String, strEndTime As String
    
    Static str上次号码 As String

    On Error GoTo errHandle
    
    If Not mblnAppointment Then Exit Sub
    
    dtValue = Me.dtpDate.Value
    With vsfArrange
        If .Row < 0 Or .Row > .Rows - 1 Then Exit Sub
        strWeekName = Replace(mobjRegistData.zlGetWeekNameFromDate(dtValue), "周", "")
        If .ColIndex(strWeekName) < 0 Then Exit Sub
        
       str安排 = .Cell(flexcpData, .Row, .ColIndex(strWeekName))
       str号码 = .TextMatrix(.Row, .ColIndex("号别"))
    End With
    
    If str号码 = str上次号码 Then Exit Sub
    If (mViewMode = V_普通号分时段 Or mViewMode = v_专家号分时段) Then str上次号码 = str号码: Exit Sub
    
    If mobjRegistData.zlGetRegisterWorkTime(str安排, "", "", strDefaultTime, strStartTime, strEndTime) = False Then Exit Sub
    
    strDate = Format(dtpDate.Value, "yyyy-mm-dd") & " "
    dtValue = dtpTime.Value
    dtpTime.MaxDate = CDate(strDate & strStartTime)
    dtpTime.minDate = CDate(strDate & strEndTime)
    
    If strDefaultTime > strEndTime Or strDefaultTime < strStartTime Then strDefaultTime = strStartTime
    If strDefaultTime = "" Then Exit Sub
    strDefaultTime = strDate & strDefaultTime
    If dtpTime.minDate > CDate(strDefaultTime) Then dtpTime.minDate = CDate(strDefaultTime)
    dtpTime.Value = CDate(strDefaultTime)
    str上次号码 = str号码
    Exit Sub
errHandle:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
End Sub

'检查安排序号数据是否合法
Private Function CheckArangement() As Boolean
    Dim str号别 As Long, strChkTime As String
    Dim lngSN As Long, i As Long, j As Long
    Dim blnExit As Boolean, strTemp As String
    Dim dtDate As Date
    If mViewMode = V_普通号 Or mViewMode = v_专家号 Then CheckArangement = True: Exit Function
     
    Select Case mViewMode
        Case V_普通号分时段
        '暂时不处理,以后有需求进行补充
        Case v_专家号分时段
            lngSN = Val(txtSN.Text)
            If lngSN = 0 Then
                If mty_Para.bln严格按时段挂号 And InStr(mstrPrivs, ";加号;") = 0 Then
                    MsgBox "该号别的时段已经使用完成,不能再进行挂号!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                CheckArangement = True: Exit Function
            End If
            With vsfDetailTime
                If .TextMatrix(.Row, .Col) Like "加*" Then CheckArangement = True: Exit Function
                If lngSN = Val(Get时段(.Row, .Col)) Then CheckArangement = True: Exit Function
                For i = 0 To .Rows - 1
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                            If lngSN = Val(Get时段(i, j, False)) Then
                               .Row = i: .Col = j
                                strTemp = Get时段(i, j, True)
                                dtDate = CDate(Format(IIf(mblnAppointment, dtpDate, gobjDatabase.Currentdate), "yyyy-mm-dd") & " " & strTemp)
                                If dtpTime.minDate > dtDate Then dtpTime.minDate = dtDate
                                If dtpTime.MaxDate < dtDate Then dtpTime.MaxDate = dtDate
                                dtpTime.Value = dtDate
                                blnExit = True: Exit For
                            End If
                        End If
                    Next
                    If blnExit Then Exit For
                Next
            End With
        Case Else
        CheckArangement = True
        Exit Function
    End Select
    CheckArangement = True
    Exit Function
Errhand:
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    gobjComlib.SaveErrLog
End Function

