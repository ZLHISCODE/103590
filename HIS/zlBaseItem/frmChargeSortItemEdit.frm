VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmChargeSortItemEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "费别单项收费设置"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmChargeSortItemEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdDel 
      Caption         =   "清除(&D)"
      Height          =   350
      Left            =   3600
      TabIndex        =   84
      Top             =   3600
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Frame fraItem 
      Caption         =   "项目选择"
      Height          =   3285
      Left            =   3600
      TabIndex        =   75
      Top             =   120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton cmdMoveAll 
         Caption         =   "移除所有(&C)"
         Height          =   350
         Left            =   4560
         TabIndex        =   83
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "移除(&M)"
         Height          =   350
         Left            =   4560
         TabIndex        =   82
         Top             =   240
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfItemList 
         Height          =   1935
         Left            =   120
         TabIndex        =   81
         Top             =   1200
         Width           =   5655
         _cx             =   9975
         _cy             =   3413
         Appearance      =   1
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
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
      Begin VB.CommandButton cmdFilter 
         Caption         =   "…"
         Height          =   270
         Left            =   3000
         TabIndex        =   80
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         Height          =   270
         Left            =   960
         TabIndex        =   79
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cbo项目类别 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "项目名称"
         Height          =   180
         Left            =   120
         TabIndex        =   78
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "项目类别"
         Height          =   180
         Left            =   120
         TabIndex        =   76
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame fra药品应用 
      Caption         =   "应用范围"
      Height          =   3300
      Left            =   4560
      TabIndex        =   68
      Top             =   3960
      Visible         =   0   'False
      Width           =   4695
      Begin VB.OptionButton opt应用于 
         Caption         =   "仅应用于本规格药品(&0)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   74
         Top             =   480
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“西成药”(&2)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   73
         Top             =   1392
         Width           =   3795
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“片剂”类药品(&3)"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   72
         Top             =   1848
         Width           =   4275
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本品种下所有药品(&1)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   71
         Top             =   936
         Width           =   2955
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于同级的所有药品(&4)"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   70
         Top             =   2304
         Width           =   2955
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本分类的所有药品(&5)"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   69
         Top             =   2760
         Width           =   2955
      End
   End
   Begin VB.Frame fra项目应用 
      Caption         =   "应用范围"
      Height          =   3300
      Left            =   840
      TabIndex        =   63
      Top             =   4200
      Visible         =   0   'False
      Width           =   3495
      Begin VB.OptionButton optApply 
         Caption         =   "应用于该分类下所有项目(&2)"
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   67
         Top             =   1680
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "应用于该类别下所有项目(&3)"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   66
         Top             =   2280
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "应用于同级的所有项目(&1)"
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   65
         Top             =   1080
         Width           =   3075
      End
      Begin VB.OptionButton optApply 
         Caption         =   "仅对本项目起作用(&0)"
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Top             =   480
         Value           =   -1  'True
         Width           =   3075
      End
   End
   Begin VB.Frame fra费别 
      Caption         =   "费别明细"
      Height          =   3300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   735
         TabIndex        =   43
         Text            =   "0.00"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   0
         Left            =   1845
         TabIndex        =   42
         Text            =   "100.000"
         Top             =   2385
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   735
         TabIndex        =   41
         Text            =   "0.00"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   1
         Left            =   1845
         TabIndex        =   40
         Text            =   "100.000"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   735
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   2
         Left            =   1845
         TabIndex        =   38
         Text            =   "100.000"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   735
         TabIndex        =   37
         Text            =   "0.00"
         Top             =   3150
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   3
         Left            =   1845
         TabIndex        =   36
         Text            =   "100.000"
         Top             =   3150
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   4
         Left            =   1845
         TabIndex        =   35
         Text            =   "100.000"
         Top             =   3405
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   4
         Left            =   735
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   3405
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   5
         Left            =   1845
         TabIndex        =   33
         Text            =   "100.000"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   5
         Left            =   735
         TabIndex        =   32
         Text            =   "0.00"
         Top             =   3660
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   6
         Left            =   1845
         TabIndex        =   31
         Text            =   "100.000"
         Top             =   3915
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   6
         Left            =   735
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   3915
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   7
         Left            =   1845
         TabIndex        =   29
         Text            =   "100.000"
         Top             =   4170
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   7
         Left            =   735
         TabIndex        =   28
         Text            =   "0.00"
         Top             =   4170
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   8
         Left            =   1845
         TabIndex        =   27
         Text            =   "100.000"
         Top             =   4425
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   8
         Left            =   735
         TabIndex        =   26
         Text            =   "0.00"
         Top             =   4425
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   1845
         TabIndex        =   25
         Text            =   "100.000"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   9
         Left            =   735
         TabIndex        =   24
         Text            =   "0.00"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   10
         Left            =   1845
         TabIndex        =   23
         Text            =   "100.000"
         Top             =   4935
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   10
         Left            =   735
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   4935
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   11
         Left            =   1845
         TabIndex        =   21
         Text            =   "100.000"
         Top             =   5190
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   11
         Left            =   735
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   5190
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   12
         Left            =   1845
         TabIndex        =   19
         Text            =   "100.000"
         Top             =   5445
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   12
         Left            =   735
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   5445
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   13
         Left            =   1845
         TabIndex        =   17
         Text            =   "100.000"
         Top             =   5700
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   13
         Left            =   735
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   5700
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   14
         Left            =   1845
         TabIndex        =   15
         Text            =   "100.000"
         Top             =   5955
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   14
         Left            =   735
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   5955
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   15
         Left            =   1845
         TabIndex        =   13
         Text            =   "100.000"
         Top             =   6210
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   270
         Index           =   15
         Left            =   735
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   6210
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtStage 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "1"
         Top             =   1800
         Width           =   300
      End
      Begin VB.ComboBox cbo计算方法 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin MSComCtl2.UpDown UdStage 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   1800
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtStage"
         BuddyDispid     =   196626
         OrigLeft        =   2010
         OrigTop         =   1200
         OrigRight       =   2250
         OrigBottom      =   1500
         Max             =   16
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblMoney 
         Caption         =   "应收分段起点"
         Height          =   180
         Left            =   750
         TabIndex        =   62
         Top             =   2160
         Width           =   1080
      End
      Begin VB.Label lblTax 
         Caption         =   "实收比率(%)"
         Height          =   195
         Left            =   1965
         TabIndex        =   61
         Top             =   2160
         Width           =   1035
      End
      Begin VB.Label lblStage 
         Caption         =   "分段号"
         Height          =   225
         Left            =   120
         TabIndex        =   60
         Top             =   2175
         Width           =   540
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   0
         Left            =   225
         TabIndex        =   59
         Top             =   2430
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   58
         Top             =   2685
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   225
         TabIndex        =   57
         Top             =   2940
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   3
         Left            =   225
         TabIndex        =   56
         Top             =   3195
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   4
         Left            =   225
         TabIndex        =   55
         Top             =   3450
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   5
         Left            =   225
         TabIndex        =   54
         Top             =   3705
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "7"
         Height          =   180
         Index           =   6
         Left            =   225
         TabIndex        =   53
         Top             =   3960
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "8"
         Height          =   180
         Index           =   7
         Left            =   225
         TabIndex        =   52
         Top             =   4215
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "9"
         Height          =   180
         Index           =   8
         Left            =   225
         TabIndex        =   51
         Top             =   4470
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "10"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   50
         Top             =   4725
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "11"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   49
         Top             =   4980
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "12"
         Height          =   180
         Index           =   11
         Left            =   180
         TabIndex        =   48
         Top             =   5235
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "13"
         Height          =   180
         Index           =   12
         Left            =   180
         TabIndex        =   47
         Top             =   5490
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "14"
         Height          =   180
         Index           =   13
         Left            =   180
         TabIndex        =   46
         Top             =   5745
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "15"
         Height          =   180
         Index           =   14
         Left            =   180
         TabIndex        =   45
         Top             =   6000
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "16"
         Height          =   180
         Index           =   15
         Left            =   180
         TabIndex        =   44
         Top             =   6255
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblItem 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "收入项目比例分段"
         Height          =   180
         Left            =   1080
         TabIndex        =   11
         Top             =   1860
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmChargeSortItemEdit.frx":000C
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label lblNote 
         Caption         =   "    每一收入项目可按应收金额划分为多段(最多16段)，设置不同的实收比例。"
         Height          =   690
         Left            =   720
         TabIndex        =   8
         Top             =   1125
         Width           =   2595
      End
      Begin VB.Label lblMeasure 
         AutoSize        =   -1  'True
         Caption         =   "计算方法"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "选择费别"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1320
      TabIndex        =   0
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   1
      Top             =   3600
      Width           =   1100
   End
End
Attribute VB_Name = "frmChargeSortItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'接口传入参数
Private mintType As Integer           '类型(由于该窗体可用于多处，用此区分不同使用环境)：0-费别设置中(收入项目)；1-费别设置中(其他项目)；2-收费项目管理中；3-药品管理中
Private mstrGrade As String           '费别：费别管理中为传入值，其他环境为空
Private mlngItemId As Long            '项目ID：费别管理中为0，其他环境为传入值
Private mStrItem As String            '项目名称

'其他变量
Private mintStage As Integer
Private mblnChange As Boolean         '是否改变了
Private mblnOk As Boolean

Private Const mconstListHead = "项目id,7,0|编码,1,1000|名称,1,1500|规格,1,1500|单位,1,800|价格,7,800"
Private Enum 项目列表
    项目id = 0
    编码 = 1
    名称 = 2
    规格 = 3
    单位 = 4
    价格 = 5
    
    列数 = 6
End Enum

Private Const mcstFormHeight As Double = 4600
Private Const mcstFormWidth As Double = 3750
Private Const mcstFormChargeHeight As Double = 3300
Private mstr药品价格等级  As String, mstr卫材价格等级 As String, mstr普通价格等级 As String

Private Sub GetDrugOtherInfo()
    '主要用于药品目录管理中得到当前药品的剂型和材质
    Dim rsTemp As ADODB.Recordset
    Dim str材质 As String
    
    If mintType <> 3 Then Exit Sub
    If mlngItemId = 0 Then Exit Sub
    
    On Error GoTo ErrHandle
    gstrSQL = "Select Decode(A.类别, '5', '西成药', '6', '中成药', '中草药') As 类别, B.药品剂型 " & _
        " From 收费项目目录 A, 药品特性 B, 药品规格 C " & _
        " Where A.ID = C.药品id And B.药名id = C.药名id And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取药品信息", mlngItemId)
    
    If Not rsTemp.EOF Then
        opt应用于(2).Caption = "应用于所有“" & rsTemp!类别 & "”(&2)"
        opt应用于(3).Caption = "应用于所有“" & rsTemp!药品剂型 & "”类药品(&3)"
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub IniItemList()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    strTemp = Split(mconstListHead, "|")
    
    With vsfItemList
        .redraw = flexRDNone
        .Rows = 1
        .Cols = 项目列表.列数
        .SelectionMode = flexSelectionByRow
        .RowHeightMin = 300
        
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .redraw = flexRDDirect
    End With
End Sub
Private Sub GetItemList(ByVal strInput As String, ByVal strItemType As String, Optional ByVal lngItemID As Long = 0)
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim strReturn As String '选择器返回字符串
    Dim strHyID As Long
    Dim strSqlCondition As String
    
    On Error GoTo ErrHandle
    
    rsTmp.CursorLocation = adUseClient

    If InStr(strInput, "'") > 0 Then
        MsgBox "输入了非法字符。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    If lngItemID > 0 Then
        strSqlCondition = " And A.Id = [5] "
    Else
        If strInput <> "" Then
            strSqlCondition = " And (A.编码 like [1] or A.名称 like [1] or  ('['||A.编码||']'||A.名称  =[3])  or  B.简码 like [2]) "
        End If
        If strItemType <> "0" Then
            strSqlCondition = strSqlCondition & " And A.类别 = [4] "
        End If
    End If
    
    Dim strWherePriceGrade As String
    If mstr普通价格等级 = "" And mstr药品价格等级 = "" And mstr卫材价格等级 = "" Then
        strWherePriceGrade = " And d.价格等级 Is Null"
    Else
        strWherePriceGrade = "" & _
            " And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And d.价格等级 = [6])" & vbNewLine & _
            "      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And d.价格等级 = [7])" & vbNewLine & _
            "      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And d.价格等级 = [8])" & vbNewLine & _
            "      Or (d.价格等级 Is Null" & vbNewLine & _
            "          And Not Exists (Select 1" & vbNewLine & _
            "                          From 收费价目" & vbNewLine & _
            "                          Where d.收费细目id = 收费细目id And Sysdate Between 执行日期 And Nvl(终止日期, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
            "                                And ((Instr(';5;6;7;', ';' || a.类别 || ';') > 0 And 价格等级 = [6])" & vbNewLine & _
            "                                      Or (Instr(';4;', ';' || a.类别 || ';') > 0 And 价格等级 = [7])" & vbNewLine & _
            "                                      Or (Instr(';4;5;6;7;', ';' || a.类别 || ';') = 0 And 价格等级 = [8])))))"
    End If
    gstrSQL = _
        "SELECT A.编码,A.名称," & _
        "A.规格,A.计算单位,ltrim(rtrim(to_char(Sum(nvl(D.现价,0)),'9999999990.00'))) 价格,A.ID" & _
        " FROM" & _
        " (Select Distinct A.ID,A.编码,A.名称,A.规格,A.计算单位,a.类别" & _
        "   From 收费项目目录 A,收费项目别名 B" & _
        "   WHERE A.ID = B.收费细目ID" & _
        "       And (A.撤档时间=to_date('3000-01-01','yyyy-mm-dd') or A.撤档时间 is null)" & strSqlCondition & _
        "   ) A,收费价目 D" & vbNewLine & _
        " Where A.ID=D.收费细目ID(+)" & _
        "       And D.执行日期 <= SYSDATE AND (D.终止日期 > SYSDATE OR D.终止日期 IS NULL)" & _
                strWherePriceGrade & vbNewLine & _
        " Group By A.编码,A.名称,A.规格,A.计算单位,A.ID"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strInput & "%", "%" & UCase(strInput) & "%", strInput, _
        strItemType, lngItemID, mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    
    If rsTmp.RecordCount < 1 Then Exit Sub
    If rsTmp.RecordCount > 1 Then
        strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "编码,1200,0,2;名称,1800,0,2;规格,1200,0,2;计算单位,800,0,2;价格,1000,1,2;ID,0,1,2", "收费项目选择器", True, , , 1000 + 1500 + 1500 + 800 + 800 + 2000)
        If Trim(strReturn) = "" Then
            Exit Sub
        End If
    Else
        strReturn = Nvl(rsTmp!编码) & "," & Nvl(rsTmp!名称) & "," & Nvl(rsTmp!规格) & "," & Nvl(rsTmp!计算单位) & "," & Nvl(rsTmp!价格) & "," & Nvl(rsTmp!ID, 0)
    End If
    
    With vsfItemList
        '检查是否重复
        For i = 0 To .Rows - 1
            If Val(.TextMatrix(i, 项目列表.项目id)) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) Then
                Exit Sub
            End If
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 项目列表.编码) = Split(strReturn, ",")(0)
        .TextMatrix(.Rows - 1, 项目列表.名称) = Split(strReturn, ",")(1)
        .TextMatrix(.Rows - 1, 项目列表.规格) = Split(strReturn, ",")(2)
        .TextMatrix(.Rows - 1, 项目列表.单位) = Split(strReturn, ",")(3)
        .TextMatrix(.Rows - 1, 项目列表.价格) = Format(Val(Split(strReturn, ",")(4)), "###0.000;-##0.000;0.000;0.000")
        .TextMatrix(.Rows - 1, 项目列表.项目id) = Split(strReturn, ",")(5)
        
        '调整控件大小
        If .Rows > 3 And .Rows < 11 And .Top + .RowHeightMin * .Rows + 50 > fraItem.Height And UdStage.value > 5 Then
            Me.Height = Me.Height + (.Rows - 3) * .RowHeightMin
            .Height = .Height + (.Rows - 3) * .RowHeightMin
            fraItem.Height = fraItem.Height + (.Rows - 3) * .RowHeightMin
            
            If fra费别.Height < .Height Then
                fra费别.Height = fraItem.Height
                cmdHelp.Top = Me.Height - cmdHelp.Height - 500
                cmdOK.Top = cmdHelp.Top
                cmdCancel.Top = cmdOK.Top
            End If
        End If
        
        If .Rows > 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, 项目列表.名称) & "等]" & "分段数："
        ElseIf .Rows = 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, 项目列表.名称) & "]" & "分段数："
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveCharge() As Boolean
    Dim str比率 As String
    Dim curStart As Currency, curEnd As Currency, dblTax As Double
    Dim i As Long
    Dim blnTrans As Boolean
    Dim int应用 As Integer
    
    curStart = Val(Me.txtMoney(0).Text)
    dblTax = Val(Me.txtTax(0).Text)
    
    Err = 0
    On Error GoTo ErrHand
    
    For mintStage = 0 To Me.UdStage.value - 1
        curStart = Val(Me.txtMoney(mintStage).Text)
        If mintStage >= Me.UdStage.value - 1 Then
            curEnd = Val("10000000000.00")
        Else
            curEnd = Val(Me.txtMoney(mintStage + 1).Text) - 0.01
        End If
        dblTax = Val(Me.txtTax(mintStage).Text)
        str比率 = str比率 & mintStage + 1 & ":" & curStart & ":" & curEnd & ":" & dblTax & ";"
    Next
    
    gcnOracle.BeginTrans
    blnTrans = False
    
    If mintType = 0 Then
        gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & mlngItemId & ",'" & str比率 & "'," & Val(cbo计算方法.Text) & "," & mintType & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ElseIf mintType = 1 Then
        '批量设置收费项目
        For i = 1 To vsfItemList.Rows - 1
            gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & Val(vsfItemList.TextMatrix(i, 项目列表.项目id)) & ",'" & str比率 & "'," & Val(cbo计算方法.Text) & "," & mintType & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
    ElseIf mintType = 2 Then
        '收费项目目录中设置费别
        If optApply(0).value = True Then
            int应用 = 0
        ElseIf optApply(1).value = True Then
            int应用 = 1
        ElseIf optApply(2).value = True Then
            int应用 = 2
        ElseIf optApply(3).value = True Then
            int应用 = 3
        End If
        
        gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & mlngItemId & ",'" & str比率 & "'," & Val(cbo计算方法.Text) & "," & mintType & "," & int应用 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    ElseIf mintType = 3 Then
        '药品目录中设置费别
        If opt应用于(0).value = True Then
            int应用 = 0
        ElseIf opt应用于(1).value = True Then
            int应用 = 1
        ElseIf opt应用于(2).value = True Then
            int应用 = 2
        ElseIf opt应用于(3).value = True Then
            int应用 = 3
        ElseIf opt应用于(4).value = True Then
            int应用 = 4
        ElseIf opt应用于(5).value = True Then
            int应用 = 5
        End If
        
        gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & mlngItemId & ",'" & str比率 & "'," & Val(cbo计算方法.Text) & "," & mintType & "," & int应用 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    
    mblnChange = False
    mblnOk = True
    blnTrans = True
    
    SaveCharge = True
    Exit Function
ErrHand:
    If Not blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Public Function ShowMe(objfrm As Object, ByVal int类型 As Integer, ByVal str费别 As String, ByVal lng项目id As Long, ByVal str项目名称 As String) As Boolean
    mintType = int类型
    mstrGrade = str费别
    mlngItemId = lng项目id
    mStrItem = str项目名称
    
    Me.Show vbModal, objfrm
    
    ShowMe = mblnOk
End Function

Private Sub LoadCharge()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 名称 From 费别 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取费别")
    
    cbo费别.Clear
    
    With rsTemp
        Do While Not .EOF
            cbo费别.AddItem !名称
            
            If !名称 = mstrGrade Then
                intIndex = cbo费别.ListCount - 1
            End If
            
            .MoveNext
        Loop
    End With
    
    If cbo费别.ListCount > 0 Then
        If mintType = 0 Or mintType = 1 Then
            cbo费别.Enabled = False
        End If
        cbo费别.ListIndex = intIndex
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadItemType()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 编码||'-'||名称 As 名称 From 收费项目类别 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取项目类别")
    
    cbo项目类别.Clear
    cbo项目类别.AddItem "0-所有类别"
    With rsTemp
        Do While Not .EOF
            cbo项目类别.AddItem !名称
            
            .MoveNext
        Loop
    End With
    
    If cbo项目类别.ListCount > 0 Then
       cbo项目类别.ListIndex = 0
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo费别_Click()
    mstrGrade = cbo费别.List(cbo费别.ListIndex)
    Call LoadChargeList(mstrGrade, mlngItemId)
End Sub


Private Sub cbo计算方法_Click()
    '1-成本价加收比例计算,不分段
    If cbo计算方法.ListIndex = 1 Then
        txtStage.Text = 1
        UdStage.value = 1
        txtStage.Enabled = False
        UdStage.Enabled = False
        lblNote.Caption = "  药品实收金额=成本价*(1+加收比率)，如果不是药品将忽略此设置，不打折。"
        lblMoney.Caption = "分段起点"
        lblTax.Caption = "加收比率(%)"
    '0-分段比例计算
    Else
       txtStage.Enabled = True
       UdStage.Enabled = True
       lblNote.Caption = "    每一收入项目可按应收金额划分为多段(最多16段)，设置不同的实收比例。"
       lblMoney.Caption = "应收分段起点"
       lblTax.Caption = "实收比率(%)"
    End If
    
End Sub

Private Sub cmdDel_Click()
    Dim int应用 As Integer
    Dim i As Integer
    
    On Error GoTo ErrHandle
    
    If mintType = 1 Then
        With vsfItemList
            If .Rows = 1 Then
                If mStrItem <> "" Then
                    If MsgBox("是否清除[" & mStrItem & "]项目的费别设置？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                    
                    gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & mlngItemId & ",Null,0," & mintType & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            ElseIf .Rows = 2 Then
                If MsgBox("是否清除[" & .TextMatrix(1, 项目列表.名称) & "]项目的费别设置？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                
                gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & Val(.TextMatrix(1, 项目列表.项目id)) & ",Null,0," & mintType & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                If MsgBox("是否清除[" & .TextMatrix(1, 项目列表.名称) & "]等项目的费别设置？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                For i = 1 To .Rows - 1
                    gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & Val(.TextMatrix(i, 项目列表.项目id)) & ",Null,0," & mintType & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                Next
            End If
        End With
    ElseIf mintType = 2 Or mintType = 3 Then
        If mintType = 2 Then
            If optApply(0).value = True Then
                int应用 = 0
            ElseIf optApply(1).value = True Then
                int应用 = 1
            ElseIf optApply(2).value = True Then
                int应用 = 2
            ElseIf optApply(3).value = True Then
                int应用 = 3
            End If
        Else
            If opt应用于(0).value = True Then
                int应用 = 0
            ElseIf opt应用于(1).value = True Then
                int应用 = 1
            ElseIf opt应用于(2).value = True Then
                int应用 = 2
            ElseIf opt应用于(3).value = True Then
                int应用 = 3
            ElseIf opt应用于(4).value = True Then
                int应用 = 4
            ElseIf opt应用于(5).value = True Then
                int应用 = 5
            End If
        End If
        
        If MsgBox("是否清除[" & mStrItem & "]及应用范围下所有项目的费别设置？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        gstrSQL = "zl_费别明细_update('" & mstrGrade & "'," & mlngItemId & ",Null,0," & mintType & "," & int应用 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
    End If
    
    MsgBox "清除成功。", vbExclamation, gstrSysName
    If mintType = 1 Then
        Call IniChargeList
        Call IniItemList
    Else
        Call IniChargeList
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub cmdFilter_Click()
'    If Trim(txtInput.Text) = "" Then Exit Sub
    
    Call GetItemList(txtInput.Text, Mid(cbo项目类别.List(cbo项目类别.ListIndex), 1, 1))
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click()
    With vsfItemList
        If .Row > 0 Then
            .RemoveItem .Row
        End If
        If .Rows > 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, 项目列表.名称) & "等]" & "分段数："
        ElseIf .Rows = 2 Then
            lblItem.Caption = "[" & .TextMatrix(1, 项目列表.名称) & "]" & "分段数："
        End If
    End With
End Sub

Private Sub cmdMoveAll_Click()
    lblItem.Caption = "分段数："
    Call IniItemList
End Sub

Private Sub Form_Load()
    mblnOk = False
    
    Me.Height = mcstFormHeight
    Me.Width = mcstFormWidth
    
    fra费别.Height = mcstFormChargeHeight
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "会员等级单项收费设置"
    End If
    
    '计算方法
    cbo计算方法.AddItem "0-分段比例计算", 0
    cbo计算方法.AddItem "1-成本价加收比例计算", 1
    
    Call GetPriceGrade(mstr药品价格等级, mstr卫材价格等级, mstr普通价格等级)
    '取费别
    Call LoadCharge
    
    '取费别明细
    Call LoadChargeList(mstrGrade, mlngItemId)
    
    fraItem.Visible = False
    fra项目应用.Visible = False
    fra药品应用.Visible = False
    cmdDel.value = False
    
    If mintType = 0 Then
    ElseIf mintType = 1 Then
        Me.Width = Me.Width + fraItem.Width + 100
        fraItem.Visible = True
        fraItem.Top = fra费别.Top
        fraItem.Left = fra费别.Left + fra费别.Width + 100
        fraItem.Height = fra费别.Height
        cmdDel.Visible = True
        
        '取项目类别
        Call LoadItemType
        
        '初始项目列表
        Call IniItemList
        
        '如果传入了项目ID，则提取该项目信息
        If mlngItemId > 0 Then
            Call GetItemList("", "", mlngItemId)
        End If
    ElseIf mintType = 2 Then
        Me.Width = Me.Width + fra项目应用.Width + 100
        fra项目应用.Visible = True
        fra项目应用.Top = fra费别.Top
        fra项目应用.Left = fra费别.Left + fra费别.Width + 100
        fra项目应用.Height = fra费别.Height
        cmdDel.Visible = True
    ElseIf mintType = 3 Then
        Me.Width = Me.Width + fra药品应用.Width + 100
        fra药品应用.Visible = True
        fra药品应用.Top = fra费别.Top
        fra药品应用.Left = fra费别.Left + fra费别.Width + 100
        fra药品应用.Height = fra费别.Height
        cmdDel.Visible = True
        
        '取药品材质，剂型信息
        Call GetDrugOtherInfo
    End If
    cmdOK.Left = Me.Width - cmdCancel.Width - cmdOK.Width - 240
    cmdCancel.Left = cmdOK.Left + cmdOK.Width
    cmdDel.Left = cmdOK.Left - cmdDel.Width - 250
End Sub

Private Sub LoadChargeList(ByVal str费别 As String, ByVal lng项目id As Long)
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    If mintType = 0 Then
        strSQL = " And 收入项目id=[2]"
    Else
        strSQL = " And 收费细目id=[2]"
    End If
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 段号, 应收段首值, 应收段尾值, 实收比率, 计算方法 " & _
        " From 费别明细 Where 费别 = [1] " & strSQL & " Order By 段号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取费别明细", str费别, lng项目id)

    If rsTemp.RecordCount = 0 Then
        Call IniChargeList
        Exit Sub
    End If
    
    cbo计算方法.ListIndex = IIF(rsTemp!计算方法 = 0, 0, 1)
    
    With rsTemp
        txtStage.Text = .RecordCount
        UdStage.value = .RecordCount
        cbo计算方法.ListIndex = Val(.Fields("计算方法").value)     '调用Click事件设置相关控件
        lblItem.Caption = "[" & mStrItem & "]" & "分段数："
        
        For i = 1 To .RecordCount
            If i > 16 Then Exit For
            
            lblNo(.AbsolutePosition - 1).Visible = True
            lblNo(.AbsolutePosition - 1).Caption = .AbsolutePosition
            txtMoney(.AbsolutePosition - 1).Visible = True
            txtMoney(.AbsolutePosition - 1).Text = Format(.Fields("应收段首值").value, "###########0.00;-##########0.00;0.00;0.00")
            txtTax(.AbsolutePosition - 1).Visible = True
            txtTax(.AbsolutePosition - 1).Text = Format(.Fields("实收比率").value, "###0.000;-##0.000;0.000;0.000")
            
            .MoveNext
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub IniChargeList()
    cbo计算方法.ListIndex = 0
    UdStage.Enabled = True
    UdStage.value = 1
    
    lblNo(0).Visible = True
    txtMoney(0).Visible = True
    txtTax(0).Visible = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
'    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
'        Cancel = 1
'    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    '校验
    If IsValidate = False Then Exit Sub
    If mintType = 2 Then
        If optApply(0).value = False Then
            For i = 0 To optApply.UBound
                If optApply(i).value = True Then
                    If MsgBox("费别设置应用范围为“" & optApply(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    End If
    '保存
    If SaveCharge = False Then Exit Sub
    
    If mintType = 0 Then
        MsgBox "设置成功。", vbExclamation, gstrSysName
        Call IniChargeList
    ElseIf mintType = 1 Then
        MsgBox "设置成功。", vbExclamation, gstrSysName
        Call IniChargeList
        Call IniItemList
    Else
        Unload Me
    End If
End Sub

Private Function IsValidate() As Boolean
    Dim intStage As Integer
    Dim str比率 As String
    Dim curStart As Currency, dblTax As Double
    Dim curStartBefore As Currency
    Dim dblTaxBefore As Double
    Dim i As Long
    
    If mintType = 1 And vsfItemList.Rows = 1 Then Exit Function
        
    For intStage = 1 To Me.UdStage.value - 1
        curStart = Val(Me.txtMoney(intStage).Text)
        dblTax = Val(Me.txtTax(intStage).Text)

        If curStart <= Val(Me.txtMoney(intStage - 1).Text) Then
            MsgBox "第" & intStage + 1 & "段错误，应收段值必须由小到大。", vbExclamation, gstrSysName
            txtMoney(intStage).SetFocus
            Exit Function
        End If
        If dblTax = Val(Me.txtTax(intStage - 1).Text) Then
            MsgBox "第" & intStage + 1 & "段错误，相邻段实收比率相同，无意义。", vbExclamation, gstrSysName
            txtTax(intStage).SetFocus
            Exit Function
        End If
    Next
    
    IsValidate = True
End Function

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub opt应用于_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To opt应用于.UBound
        If i = Index Then
            opt应用于(i).FontBold = True
        Else
            opt应用于(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
'    If Trim(txtInput.Text) = "" Then Exit Sub
    
    Call GetItemList(txtInput.Text, Mid(cbo项目类别.List(cbo项目类别.ListIndex), 1, 1))
End Sub


Private Sub txtMoney_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtMoney_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtMoney(Index)
End Sub

Private Sub txtMoney_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii > vbKey9 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtMoney_Validate(Index As Integer, Cancel As Boolean)
    Me.txtMoney(Index).Text = Format(Val(Me.txtMoney(Index).Text), "###########0.00;-##########0.00;0.00;0.00")
    If Val(Me.txtMoney(Index).Text) >= Val("10000000000.00") Or Val(Me.txtMoney(Index).Text) < 0 Then
        MsgBox "应收金额起点只能在 0～10000000000.00之间。", vbExclamation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub txtTax_Change(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtTax_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtTax(Index)
End Sub

Private Sub txtTax_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii > vbKey9 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtTax_Validate(Index As Integer, Cancel As Boolean)
    Me.txtTax(Index).Text = Format(Val(Me.txtTax(Index).Text), "###0.000;-##0.000;0.000;0.000")
    If Val(Me.txtTax(Index).Text) > 500 Or Val(Me.txtTax(Index).Text) < 0 Then
        MsgBox "实收比率只能在 0～500之间。", vbExclamation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UdStage_Change()
    Dim dblRowHeight As Double
    Dim intValue As Integer
    
    intValue = Me.UdStage.value
    dblRowHeight = txtMoney(0).Height
        
    For mintStage = 0 To 15
        Me.lblNo(mintStage).Visible = (Me.UdStage.value > mintStage)
        Me.txtMoney(mintStage).Visible = (Me.UdStage.value > mintStage)
        Me.txtTax(mintStage).Visible = (Me.UdStage.value > mintStage)
    Next
    
    mblnChange = True
     
    If intValue < 4 Then Exit Sub
        
    fra费别.Height = 2750 + (intValue - 1) * dblRowHeight
    Me.Height = 3905 + (intValue - 1) * dblRowHeight
    cmdHelp.Top = Me.Height - cmdHelp.Height - 500
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    cmdDel.Top = cmdHelp.Top
    
    If fraItem.Visible = True Then
        fraItem.Height = fra费别.Height
        vsfItemList.Height = fraItem.Height - vsfItemList.Top - 50
    End If
End Sub

Private Function GetPriceGrade(ByRef str药品价格等级 As String, _
    ByRef str卫材价格等级 As String, ByRef str普通价格等级 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取当前站点价格等级
    '入参:
    '返回:价格等级获取成功返回True，否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    str药品价格等级 = "": str卫材价格等级 = "": str普通价格等级 = ""
    strSQL = "" & _
        "Select Max(Decode(b.是否适用药品, 1, 价格等级, Null)) As 药品等级," & vbNewLine & _
        "       Max(Decode(b.是否适用卫材, 1, 价格等级, Null)) As 卫材等级," & vbNewLine & _
        "       Max(Decode(b.是否适用普通项目, 1, 价格等级, Null)) As 普通等级" & vbNewLine & _
        "From 收费价格等级应用 A, 收费价格等级 B" & vbNewLine & _
        "Where a.价格等级 = b.名称 And a.性质 = 0 And a.站点 = [1]" & vbNewLine & _
        "      And (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'yyyy-mm-dd'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取价格等级", gstrNodeNo)
    If Not rsTemp.EOF Then
        str药品价格等级 = Nvl(rsTemp!药品等级)
        str卫材价格等级 = Nvl(rsTemp!卫材等级)
        str普通价格等级 = Nvl(rsTemp!普通等级)
    End If
    GetPriceGrade = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

