VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.UserControl usrBodyEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   9555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   LockControls    =   -1  'True
   ScaleHeight     =   9555
   ScaleWidth      =   11415
   Begin VB.PictureBox picTmp 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4020
      ScaleHeight     =   300
      ScaleWidth      =   4590
      TabIndex        =   34
      Top             =   8850
      Width           =   4590
      Begin VB.ComboBox cboBaby 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2625
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   0
         Width           =   1920
      End
      Begin VB.OptionButton opt 
         Caption         =   "母亲本人(&0)"
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   36
         Top             =   60
         Value           =   -1  'True
         Width           =   1290
      End
      Begin VB.OptionButton opt 
         Caption         =   "婴儿(&1)"
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   35
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   8280
      Left            =   135
      ScaleHeight     =   8280
      ScaleWidth      =   11220
      TabIndex        =   0
      Top             =   450
      Width           =   11220
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   4815
         TabIndex        =   1
         Top             =   7200
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   7575
         TabIndex        =   2
         Top             =   6300
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin VB.PictureBox picCover 
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   6555
         ScaleHeight     =   660
         ScaleWidth      =   975
         TabIndex        =   3
         Top             =   7185
         Width           =   975
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6810
         Left            =   240
         ScaleHeight     =   6810
         ScaleWidth      =   10920
         TabIndex        =   4
         Top             =   225
         Width           =   10920
         Begin VB.PictureBox picCard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   690
            Index           =   0
            Left            =   90
            ScaleHeight     =   690
            ScaleWidth      =   10560
            TabIndex        =   12
            Top             =   75
            Width           =   10560
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   7
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Text            =   "诊断"
               Top             =   375
               Width           =   2370
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   6
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "年龄"
               Top             =   60
               Width           =   645
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   5
               Left            =   2445
               Locked          =   -1  'True
               TabIndex        =   31
               TabStop         =   0   'False
               Text            =   "性别"
               Top             =   60
               Width           =   420
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   4
               Left            =   3375
               Locked          =   -1  'True
               TabIndex        =   17
               TabStop         =   0   'False
               Text            =   "12"
               Top             =   390
               Width           =   615
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   3
               Left            =   4875
               Locked          =   -1  'True
               TabIndex        =   16
               TabStop         =   0   'False
               Text            =   "入院日期"
               Top             =   60
               Width           =   1140
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   2
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   15
               TabStop         =   0   'False
               Text            =   "科室"
               Top             =   375
               Width           =   2400
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   1
               Left            =   6645
               Locked          =   -1  'True
               TabIndex        =   14
               TabStop         =   0   'False
               Text            =   "1234567"
               Top             =   60
               Width           =   3825
            End
            Begin VB.TextBox txtCard 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   0
               Left            =   465
               Locked          =   -1  'True
               TabIndex        =   13
               TabStop         =   0   'False
               Text            =   "姓无名"
               Top             =   60
               Width           =   1425
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "诊    断:"
               Height          =   180
               Index           =   7
               Left            =   4065
               TabIndex        =   38
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "年龄:"
               Height          =   180
               Index           =   6
               Left            =   2910
               TabIndex        =   32
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "性别:"
               Height          =   180
               Index           =   4
               Left            =   1980
               TabIndex        =   30
               Top             =   60
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "入院日期:"
               Height          =   180
               Index           =   5
               Left            =   4050
               TabIndex        =   22
               Top             =   60
               Width           =   810
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "床号:"
               Height          =   180
               Index           =   3
               Left            =   2910
               TabIndex        =   21
               Top             =   390
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "科室:"
               Height          =   180
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   375
               Width           =   450
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "住院号:"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   19
               Top             =   60
               Width           =   630
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "姓名:"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   60
               Width           =   450
            End
         End
         Begin VB.PictureBox picScale 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   840
            Left            =   1230
            ScaleHeight     =   810
            ScaleWidth      =   5850
            TabIndex        =   26
            Top             =   1350
            Width           =   5880
            Begin VB.Label lblCur 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "△"
               ForeColor       =   &H8000000C&
               Height          =   180
               Left            =   180
               TabIndex        =   27
               Top             =   570
               Width           =   180
            End
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1950
            Left            =   1155
            ScaleHeight     =   1920
            ScaleWidth      =   5520
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   2355
            Width           =   5550
            Begin VB.PictureBox picGraph 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   4140
               Left            =   780
               ScaleHeight     =   4140
               ScaleWidth      =   5445
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   135
               Width           =   5445
               Begin VB.Line linHCur 
                  BorderStyle     =   3  'Dot
                  Visible         =   0   'False
                  X1              =   300
                  X2              =   1635
                  Y1              =   720
                  Y2              =   720
               End
               Begin VB.Line linVCur 
                  BorderStyle     =   3  'Dot
                  Visible         =   0   'False
                  X1              =   1785
                  X2              =   1785
                  Y1              =   15
                  Y2              =   690
               End
            End
         End
         Begin VB.PictureBox picLine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   3420
            Index           =   0
            Left            =   7785
            ScaleHeight     =   3390
            ScaleWidth      =   0
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   1035
            Width           =   15
         End
         Begin zl9BodyEditorHN.VsfGrid vsf 
            Height          =   270
            Left            =   450
            TabIndex        =   5
            Top             =   4440
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   476
         End
         Begin VSFlex8Ctl.VSFlexGrid mshUpTab 
            Height          =   780
            Left            =   240
            TabIndex        =   6
            Top             =   810
            Width           =   5775
            _cx             =   10186
            _cy             =   1376
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   8
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
         Begin VSFlex8Ctl.VSFlexGrid mshDownTab 
            Height          =   1695
            Left            =   300
            TabIndex        =   7
            Top             =   5085
            Width           =   7215
            _cx             =   12726
            _cy             =   2990
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   0   'False
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   4
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
            Begin VB.PictureBox picInput 
               Appearance      =   0  'Flat
               BackColor       =   &H80000001&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Left            =   1860
               ScaleHeight     =   240
               ScaleWidth      =   3360
               TabIndex        =   8
               Top             =   495
               Visible         =   0   'False
               Width           =   3360
               Begin VB.TextBox txtInput 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   0
                  Left            =   0
                  MaxLength       =   12
                  TabIndex        =   10
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.TextBox txtInput 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   1
                  Left            =   1335
                  MaxLength       =   12
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1035
               End
               Begin VB.Label lblInput 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "/E"
                  Height          =   180
                  Left            =   1065
                  TabIndex        =   11
                  Top             =   30
                  Width           =   180
               End
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid mshScale 
            Height          =   3090
            Left            =   150
            TabIndex        =   28
            Top             =   1185
            Width           =   6420
            _cx             =   11324
            _cy             =   5450
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
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   12
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "说明："
            Height          =   180
            Left            =   315
            TabIndex        =   29
            Top             =   6570
            Width           =   540
         End
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrBodyEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'常量
'----------------------------------------------------------------------------------------------------------------------

'存储体温单曲线部份的数据的标志
Private Enum GraphDataRow
    更改标志 = 0
    曲线数据 = 1
    上标说明 = 2
    手术标志 = 3
    部位标志 = 4
    入院标志 = 5
    转科标志 = 6
    换床标志 = 7
    出院标志 = 8
    入科标志 = 9
    复试标志 = 10
    下标说明 = 11
    断开标志 = 12
    出生标志 = 13
    曲线时间 = 14
    未记说明 = 15
End Enum

Private Enum GridDataRow
    修改标志 = 0
End Enum

'操作类型
Private Enum OperateType
    新增操作 = 2                                        '全部是新增的点
    修改操作 = 3                                        '修改的点：可能包含原有的点和新增的点
    删除操作 = 4                                        '删除的点
End Enum

Private Const HOUR_STEP_Twips = 240                 '最小单元格的宽度
Private Const ROWHEIGHT = 39                        '最小单元格的高度*5
Private Const MAXROWS = 47                          '

'自定义类型
'----------------------------------------------------------------------------------------------------------------------
Private Type ITEM_NO
    大便 As Long
    出液 As Long
    心率 As Long
    体温 As Long
    脉搏 As Long
    呼吸 As Long
    血压 As Long
    舒张压 As Long
End Type

Private Type ITEM_SERIAL
    饮入物 As Integer
    饮入量 As Integer
    体温 As Integer
    脉搏 As Integer
    血压 As Integer
    呼吸 As Integer
    心率 As Integer
End Type

Private Type ITEM_STRUCT
    项目名称 As String
    数据类型 As Integer
    数据长度 As Integer
    小数位数 As Integer
    最小值 As String
    最大值 As String
    记录频次 As Integer
    活动项目 As Boolean
    项目序号 As Long
End Type

Private Type GRAPHPOINT
    X As Single
    Y As Single
    符号 As String
    颜色 As Long
    标志 As Byte
End Type

Private Type BODYFLAG
    入院 As Byte
    入科 As Byte
    转出 As Byte
    换床 As Byte
    手术 As Byte
    出院 As Byte
    分娩 As Byte
    出生 As Byte
End Type

'变量定义
'----------------------------------------------------------------------------------------------------------------------
Private mintOpDays As Integer
Private mblnStopFlag As Boolean
Private mbln呼吸曲线 As Boolean
Private mblnBabys As Boolean
Private mint心率应用 As Integer
Private mstr心率符号 As String
Private mstr最小时间 As String      '入院时间或入科时间
Private mbln婴儿体温单显示出院 As Boolean
Private mstrParam As String
Private mlngHourBegin As Long                       '当天刻度开始时间
Private mlngPageCur As Long                         '记录当前是哪一页
Private mblnMoved As Boolean
Private mstrEnterDate As String                      '入院日期
Private mstrSQL As String
Private rsTemp As New ADODB.Recordset
Private intRow As Integer
Private intCol As Integer
Private intCount As Integer                         '行列自由记数器
Private mvarEdit As Boolean                         '是否允许编辑
Private mblnNoneShow As Boolean                     '确定当前体温表是不是显示出来
Private mrsParam As New ADODB.Recordset             '存储格式：病人ID,主页ID,病区ID,科室ID,出院,编辑
Private mstrMsgTitle As String
Private mfrmParent As Object
Private mlngLine As Long
Private mstr体温部位 As String
Private mstr呼吸方式 As String
Private mstr脉搏 As String                          '起搏器或空
Private mlngNo As Long
Private mstrOpsSvr(1 To 7) As String
Private mstrOpsDays(1 To 7) As String
Private mItemSerial As ITEM_SERIAL
Private mItemNo As ITEM_NO
Private mBodyFlag As BODYFLAG
Private mItemStru() As ITEM_STRUCT
Private mItemOtherStru(0 To 1) As ITEM_STRUCT       '其他，0->呼吸;1-舒张压
Private mstrChar(2) As String                       '依次为口温,腋温,肛温
Private mstrBreath As String                        '呼吸
Private mstrPulse As String                         '脉搏
Private mcbrToolBar页面 As CommandBarControl
Private mcbrToolBar As CommandBar
Private WithEvents mfrmCaseTendBodyPrint As frmCaseTendBodyPrint
Attribute mfrmCaseTendBodyPrint.VB_VarHelpID = -1

'事件定义
Public Event Activate()
Public Event DbClickCur()
Public Event DataChanged(ByVal blnChanged As Boolean)
Public Event RButton(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PromptInfo(ByVal strInfo As String)
Public Event SelectScale(ByVal intScale As Integer)
Public Event zlAfterPrint()

Private msinVStep As Single      '滚动条的步长
Private msinHStep As Single      '滚动条的步长

'API引用
'----------------------------------------------------------------------------------------------------------------------
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

'修改说明
'---------------------------------------------------------------
'20090923:呼吸增加呼吸机辅助呼吸功能，当呼吸为表格项录入时，默认为自主呼吸，如果要支持呼吸机辅助呼吸，必须为曲线项


'自定义函数、过程区域
'######################################################################################################################
Public Property Get ParentForm() As Object
    Set ParentForm = mfrmParent
End Property

Public Property Set ParentForm(objParent As Object)
    Set mfrmParent = objParent
End Property

Public Property Get ScrollBarY() As FlatScrollBar
    Set ScrollBarY = vsb
End Property

Public Property Get ScrollBarX() As FlatScrollBar
    Set ScrollBarX = hsb
End Property

Public Property Get 体温项目() As Boolean
    体温项目 = (mItemSerial.体温 = Val(picGraph.Tag))
End Property

Public Property Let 体温部位(vData As String)
    mstr体温部位 = vData
End Property

Public Property Get 呼吸项目() As Boolean
    呼吸项目 = (mItemSerial.呼吸 = Val(picGraph.Tag))
End Property

Public Property Let 呼吸方式(vData As String)
    mstr呼吸方式 = vData
End Property

Public Property Get 脉搏项目() As Boolean
    脉搏项目 = (mItemSerial.脉搏 = Val(picGraph.Tag))
End Property

Public Property Let 脉搏方式(vData As String)
    mstr脉搏 = vData
End Property

Public Property Get CurPostion() As Long
    CurPostion = lblCur.Left \ HOUR_STEP_Twips
End Property

Public Property Get 是否大便项目() As Boolean

    是否大便项目 = (Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.大便)
    
End Property

Public Property Get 是否出液项目() As Boolean

    是否出液项目 = (Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.出液)
    
End Property

Public Property Get GetPicScale() As Object
    Set GetPicScale = picScale
End Property

Public Property Get GetmshScale() As Object
    Set GetmshScale = mshScale
End Property

Public Property Get GetUpObj() As Object
    Set GetUpObj = mshUpTab
End Property

Public Property Get GetpicLine(ByVal intIndex) As Object
    Set GetpicLine = picLine(intIndex)
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get Page() As Long
    Page = mlngPageCur
End Property

Public Property Get LineType() As Long
    LineType = mlngLine
End Property

Public Function ConvertToValue(ByVal intNo As Integer, ByVal Y As Long) As Double
    '******************************************************************************************************************
    '功能： 转换纵坐标为值
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim aryValue() As String
    
    aryValue = Split(picLine(intNo).Tag, ";")
    
    ConvertToValue = aryValue(0) - (Y / mshScale.ROWHEIGHT(1) - aryValue(3) + 1) * aryValue(2)
    
End Function

Public Function ConvertToY(ByVal intCol As Integer, ByVal dbValue As Double) As Long
    '******************************************************************************************************************
    '功能： 转换值为纵坐标
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '获取项目定义:最大值；最小值；单位值；最高行
    aryValue = Split(picLine(intCol).Tag, ";")

    '坐标值=((最大值-当前值)/单位值+最高行-1)*行高度
    ConvertToY = ((Val(aryValue(0)) - dbValue) / Val(aryValue(2)) + Val(aryValue(3)) - 1) * mshScale.ROWHEIGHT(1)
    
End Function

Public Function GetMaxValue(ByVal intCol As Integer) As Double
    '******************************************************************************************************************
    '功能： 获取最大值
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '获取项目定义:最大值；最小值；单位值；最高行
    aryValue = Split(picLine(intCol).Tag, ";")

    GetMaxValue = Val(aryValue(0))

End Function

Public Function GetMinValue(ByVal intCol As Integer) As Double
    '******************************************************************************************************************
    '功能： 获取最大值
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim aryValue() As String
    
    '获取项目定义:最大值；最小值；单位值；最高行
    aryValue = Split(picLine(intCol).Tag, ";")

    GetMinValue = Val(aryValue(1))

End Function


Public Function zlMenuClick(ByVal strMenuItem As String, Optional ByVal strParam As String) As Boolean
    '******************************************************************************************************************
    '功能： 菜单功能处理，主要用于上级窗体接口调用
    '参数： strMenuItem         功能名称
    '       strParam            参数字符串
    '返回： 调用成功返回TRUE；否则FALSE
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim varParam As Variant
    Dim strTmp As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim aryValue As Variant
    Dim intRewrite As Integer
    Dim intNowCol As Integer
    Dim intCol As Integer
    Dim intLoop As Integer
        
    If strParam <> "" Then varParam = Split(strParam, ";")
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "初始数据"
        
        'strParam格式：病人ID;主页ID;病区ID;科室ID;出院;编辑
        
        Set mrsParam = New ADODB.Recordset
    
        Call CreateParam(mrsParam, "病人id", adBigInt)
        Call CreateParam(mrsParam, "主页id", adBigInt)
        Call CreateParam(mrsParam, "病区id", adBigInt)
        Call CreateParam(mrsParam, "科室id", adBigInt)
        Call CreateParam(mrsParam, "出院", adTinyInt)
        Call CreateParam(mrsParam, "婴儿", adTinyInt)
        Call CreateParam(mrsParam, "编辑", adTinyInt)
        Call CreateParam(mrsParam, "病人来源", adTinyInt)
        Call CreateParam(mrsParam, "开始时间", adVarChar, 20)
        Call CreateParam(mrsParam, "结束时间", adVarChar, 20)
        Call CreateParam(mrsParam, "护理等级", adTinyInt)
        
        mrsParam.Open
        mrsParam.AddNew
                        
        mrsParam("病人id").Value = Val(varParam(0))
        mrsParam("主页id").Value = Val(varParam(1))
        mrsParam("病区id").Value = Val(varParam(2))
        mrsParam("科室id").Value = Val(varParam(2))
        If UBound(varParam) >= 3 Then
            mrsParam("出院").Value = Val(varParam(3))
        Else
            mrsParam("出院").Value = 1
        End If
        
        If UBound(varParam) >= 4 Then
            mrsParam("编辑").Value = Val(varParam(4))
        Else
            mrsParam("编辑").Value = 0
        End If
        
        If UBound(varParam) >= 5 Then
            mrsParam("婴儿").Value = Val(varParam(5))
        Else
            mrsParam("婴儿").Value = 0
        End If
        
        mrsParam("病人来源").Value = 2
        mrsParam("护理等级").Value = 3
        
        gstrSQL = "Select a.序号,Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名 From 病人新生儿记录 a,病人信息 b Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id Order By a.序号"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "usrBodyEditor", Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value))
        mblnBabys = (rs.BOF = False)
        picTmp.Visible = mblnBabys
        cboBaby.Clear
        If rs.BOF = False Then
            Do While Not rs.EOF
                cboBaby.AddItem rs("婴儿姓名").Value
                cboBaby.ItemData(cboBaby.NewIndex) = rs("序号").Value
                rs.MoveNext
            Loop
        End If
        If cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0

        
        Call InitData
                
        Call ClearLineSelect
        Call FaceInit
        Call SetBodyMode
        
        If InitBody(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("病区id").Value), Val(mrsParam("婴儿").Value)) = False Then Exit Function
        
    '------------------------------------------------------------------------------------------------------------------
    Case "装载数据"
        
        'strParam格式：起始时间;科室ID;开始时间;结束时间;页号
        
        '判断体温表是否保存，询问保存与否
        strTmp = isSaved()
        If strTmp <> "" Then
            If MsgBox(strTmp & "修改信息丢失！" & vbCrLf & "是否放弃保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        mlngNo = -1
        If strParam = "" Then
            
            '画一个空的体温表,即没有日期、数据及病人信息
            Call DrawScale
            Call DrawPaper
        Else
             mstrParam = strParam
             
'             If InitBody(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("病区id").Value), Val(mrsParam("婴儿").Value)) = False Then Exit Function
             
             mrsParam("病区id").Value = Val(varParam(1))
             mrsParam("科室id").Value = Val(varParam(1))
             mrsParam("开始时间").Value = CStr(varParam(2))
             mrsParam("结束时间").Value = CStr(varParam(3))
             mlngNo = Val(varParam(4))
             
            '调入新的页
            mlngPageCur = Val(varParam(4))
            mstrEnterDate = CStr(varParam(0))
            
            Call ReadBodyData
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
            
            '控制显示控件
'            Call UserControl_Resize
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "刷新数据"
        
        
        '判断体温表是否保存，询问保存与否
        strTmp = isSaved()
        If strTmp <> "" Then
            If MsgBox(strTmp & "修改信息丢失！" & vbCrLf & "是否放弃保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        If mstrParam = "" Then
            
            '画一个空的体温表,即没有日期、数据及病人信息
            Call DrawScale
            Call DrawPaper
        Else
            
            Call ReadBodyData
            
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
            
            '控制显示控件
'            Call UserControl_Resize
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "保存数据"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function

        'strParam格式：
        If SaveData Then
            '需要重新装载
            zlMenuClick = True
            Call ReadBodyData
            Call DrawScale
            Call DrawPaper
            Call DrawGraph
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "恢复数据"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        
        If MsgBox("确实要恢复更改前的数据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        Call ReadBodyData
        Call DrawScale
        Call DrawPaper
        Call DrawGraph
    '------------------------------------------------------------------------------------------------------------------
    Case "操作曲线"
        
        mlngLine = Val(strParam)
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        
        'strParam格式：曲线索引
                
        If Val(varParam(0)) = 0 Then
            If picGraph.Tag <> "" Then
                mshScale_MouseUp 1, 0, Val(picGraph.Tag) * mshScale.ColWidth(0) + 90, 0
                Call ClearLineSelect
            End If
        Else
            mshScale_MouseUp 1, 0, Val(varParam(0)) * mshScale.ColWidth(0) - 90, 0
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "添加项目"
        
        Dim rsData As New ADODB.Recordset
        Dim rsTmp As New ADODB.Recordset
        Dim strNotItem As String
                
        strNotItem = ""
        For intLoop = LBound(mItemStru) To UBound(mItemStru) Step -1
            
            If mItemStru(intLoop).活动项目 Then
                strNotItem = strNotItem & "," & mItemStru(intLoop).项目序号
            End If
            
        Next
        If strNotItem <> "" Then strNotItem = Mid(strNotItem, 2)
                
        Set rsData = GetGridItem(Val(mrsParam("护理等级").Value), Val(mrsParam("科室id").Value), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2), 2, strNotItem)
        
        If rsData.BOF = False Then
            If ShowTxtSelDialog(mfrmParent, Nothing, "名称,1500,0,1;单位,900,0,0;最小值,900,0,0;最大值,900,0,0", mfrmParent.Name & "\护理项目选择", "请从下面选择一个护理项目。", rsData, rsTmp, 6000, 3000, , , 2, False) Then
                If rsTmp.BOF = False Then
                    If AppendGridItem(rsTmp, True) Then
                        Call picPane_Resize
                    End If
                End If
            End If
        End If


    '------------------------------------------------------------------------------------------------------------------
    Case "删除项目"
        
        With mshDownTab
            If .Row <= UBound(mItemStru) Then
                If mItemStru(.Row).活动项目 Then
                    
                    '检查是否有数据，如果无数据时才允许删除
                    '本次保存之前有数据以及当前界面上有数据，则称之为有数据
                    If CheckGridData(.Row) Then
                        ShowSimpleMsg "对不起，你要删除表格行有数据或者以前有数据！"
                        Exit Function
                    End If
                    
                    If MsgBox("确实要删除当前的表格项目吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
                    If DeleteActiveItem(.Row) Then
                        Call picPane_Resize
                    End If
                    
                End If
            End If
        End With
        
    '------------------------------------------------------------------------------------------------------------------
    Case "复试合格"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If picScale.Tag <> "" Then
            With mshScale
                intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
                
                '判断是否是有值
                If .TextMatrix(1, intNowCol) <> "" Then
                    If Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.体温 + 1) <> "" Then
                        aryValue = Split(Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.体温 + 1), ",")
                        
                        If .TextMatrix(10, intNowCol) <> "1" And Val(aryValue(0)) > 0 Then
                            .TextMatrix(10, intNowCol) = "1"
                            RaiseEvent DataChanged(True)
                            
                            aryValue = Split(.TextMatrix(0, intNowCol), ";")
                            
                            intRewrite = Val(aryValue(mItemSerial.体温 + 1))
                            Select Case intRewrite
                            Case 0
                                aryValue(mItemSerial.体温 + 1) = 2
                            Case 1
                                aryValue(mItemSerial.体温 + 1) = 3
                            Case 2
                                aryValue(mItemSerial.体温 + 1) = 2
                            Case 3
                                aryValue(mItemSerial.体温 + 1) = 3
                            Case 4
                                aryValue(mItemSerial.体温 + 1) = 3
                            End Select
                            
                            .TextMatrix(0, intNowCol) = Join(aryValue, ";")
                            
                            Call DrawPaper
                            Call DrawGraph
                                            
                        End If
                    End If
                End If
            End With
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "取消复试"
                
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If picScale.Tag <> "" Then
                
            With mshScale
                intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
                
                
                '判断是否是有升温或降温的情况，即有两个值时
                If .TextMatrix(1, intNowCol) <> "" Then
                    
                    If Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.体温 + 1) <> "" Then
                    
                        aryValue = Split(Split(.TextMatrix(1, intNowCol), ";")(mItemSerial.体温 + 1), ",")
        
                        If .TextMatrix(10, intNowCol) = "1" And Val(aryValue(0)) > 0 Then
                        
                            .TextMatrix(10, intNowCol) = "0"
                            RaiseEvent DataChanged(True)
                            
                            aryValue = Split(.TextMatrix(0, intNowCol), ";")
                            
                            intRewrite = Val(aryValue(mItemSerial.体温 + 1))
                            Select Case intRewrite
                            Case 0
                                aryValue(mItemSerial.体温 + 1) = 2
                            Case 1
                                aryValue(mItemSerial.体温 + 1) = 3
                            Case 2
                                aryValue(mItemSerial.体温 + 1) = 2
                            Case 3
                                aryValue(mItemSerial.体温 + 1) = 3
                            Case 4
                                aryValue(mItemSerial.体温 + 1) = 3
                            End Select
                            
                            .TextMatrix(0, intNowCol) = Join(aryValue, ";")
                            
                            Call DrawPaper
                            Call DrawGraph
                
                        End If
                    End If
                End If
                
            End With
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "填写手术日"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshUpTab.FocusRect = flexFocusLight Then Exit Function
        
        Call mshUpTab_KeyDown(13, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "清除手术日"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        Call mshUpTab_KeyDown(46, 0)
        
    '------------------------------------------------------------------------------------------------------------------
    Case "获取手术日"
        Dim dtOperate As Date
        Dim intStart As Integer
        Dim intEnd As Integer
        For intCol = 1 To mshUpTab.Cols - 1
            Select Case mshUpTab.ColData(intCol)
            Case 0      '非手术日
            Case 1      '原来就是手术日，设置为删除手术日
                mshUpTab.ColData(intCol) = 3
            Case 2      '新手术日，再次设置为非手术日
                mshUpTab.ColData(intCol) = 0
            Case 3      '被删除的的手术日
            End Select
        Next
        
        '从医嘱记录提取手术、分娩数据
        Set rs = GetDataFromHis(Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("婴儿")), CDate(Split(picScale.Tag, ";")(0)), CDate(Split(picScale.Tag, ";")(1)), 1)
        If Not (rs Is Nothing) Then
            If rs.BOF = False Then
                Do While Not rs.EOF
    
                    dtOperate = Int(rs("执行时间").Value)
                    intCol = dtOperate - Int(CDate(CDate(Split(picScale.Tag, ";")(0)))) + 1
                    If intCol >= 1 And intCol <= 7 Then
                    
                        mshUpTab.ColData(intCol) = 1
                        
                        mstrOpsDays(intCol) = Format(rs("执行时间").Value, "yyyy-MM-dd HH:mm:ss")
                        
                        '先清除当前日期内的所有手术文字显示内容
    
                        intStart = GetCurveColumn(CDate(Format(mstrOpsDays(intCol), "yyyy-MM-dd") & " 01:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        intEnd = GetCurveColumn(CDate(Format(mstrOpsDays(intCol), "yyyy-MM-dd") & " 23:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        
                        For intLoop = intStart To intEnd
                            mshScale.TextMatrix(3, intLoop) = ""
                        Next
                        
                        intLoop = GetCurveColumn(CDate(mstrOpsDays(intCol)), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                        
                        Select Case rs("内容").Value
                        Case "手术"
                            If intLoop >= mshScale.FixedCols And intLoop < mshScale.Cols And mBodyFlag.手术 > 0 Then
                                If mBodyFlag.手术 = 2 Then
                                    mshScale.TextMatrix(3, intLoop) = "手术--" & ConvertTimeToChinese(Format(mstrOpsDays(intCol), "HH:mm"))
                                Else
                                    mshScale.TextMatrix(3, intLoop) = "手术"
                                End If
                            End If
                        Case "分娩"
                            If intLoop >= mshScale.FixedCols And intLoop < mshScale.Cols And mBodyFlag.分娩 > 0 Then
                                If mBodyFlag.分娩 = 2 Then
                                    mshScale.TextMatrix(3, intLoop) = "分娩--" & ConvertTimeToChinese(Format(mstrOpsDays(intCol), "HH:mm"))
                                Else
                                    mshScale.TextMatrix(3, intLoop) = "分娩"
                                End If
                            End If
                        End Select
                        
                        mshUpTab.Tag = "填写手术日"
                    End If
                    
                    rs.MoveNext
                Loop
            End If
        End If
        Call ShowOpsDays
        Call DrawPaper
        Call DrawGraph
    '------------------------------------------------------------------------------------------------------------------
    Case "填写记录线"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
                
        If picScale.Tag <> "" Then
            Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
                        
            If frmCaseTendBodySetLine.ShowEdit(UserControl.Extender, lblCur.Left \ HOUR_STEP_Twips, intMinCol, intMaxCol, mrsParam("护理等级").Value, mint心率应用) Then
                RaiseEvent DataChanged(True)
            End If
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "清除记录线"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        
        If picScale.Tag <> "" Then
                        
            If frmCaseTendBodyDelLine.ShowEdit(UserControl.Extender, lblCur.Left \ HOUR_STEP_Twips, Val(mrsParam("护理等级").Value), Val(mrsParam("婴儿").Value)) Then
                RaiseEvent DataChanged(True)
            End If
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "填写表格项"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        Call mshDownTab_DblClick
    '------------------------------------------------------------------------------------------------------------------
    Case "清除表格项"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function

        mshDownTab.SetFocus
        Call mshDownTab_KeyUp(46, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "计算饮入"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
                
        If picScale.Tag <> "" And lblCur.Left >= 0 Then
                
            aryValue = Split(picScale.Tag, ";")
            strTmp = Int(CDate(aryValue(0))) + ((lblCur.Left \ HOUR_STEP_Twips) * 4) / 24
            strTmp = Format(strTmp, "yyyy-MM-DD")
            
            If MsgBox("确实需要计算“" & strTmp & "”内的饮入物和饮入量？", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Function
            
            zlMenuClick = ReadDrink(strTmp)
            
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "假肛"
    
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.大便 Then Exit Function
        If picInput.Visible Then picInput.Visible = False
        
'        mbytSpecChar = 1
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 1
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "*")

    '------------------------------------------------------------------------------------------------------------------
    Case "灌肠"
    
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.大便 Then Exit Function
        
        If picInput.Visible Then picInput.Visible = False
'        mbytSpecChar = 2
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 2
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "E")
        mshDownTab.SetFocus
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "灌肠后排泄"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.大便 Then Exit Function
'        mbytSpecChar = 3
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 3
        
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "/E")
        
'        strTmp = Trim(mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col))

        If picInput.Visible Then
'            If Right(strTmp, 2) <> "/E" And strTmp <> "" Then
'                mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = strTmp & "/E"
'            End If

            mshDownTab.SetFocus

        Else
            Call ShowInput
        End If
        
        If txtInput(0).Visible Then
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 0
        End If
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "导尿"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.出液 Then Exit Function
        
        If picInput.Visible Then picInput.Visible = False
'        mbytSpecChar = 4
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 4
        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "C")
        mshDownTab.SetFocus
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "保留导尿"
        
        If Val(mrsParam("编辑")) = 0 Then Exit Function
        If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Function
        If CheckTimeRange(mshDownTab.Col) = False Then Exit Function
        
        If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.出液 Then Exit Function
'        mbytSpecChar = 5
'        mshDownTab.Cell(flexcpData, mshDownTab.Row, mshDownTab.Col) = 5

        Call WriteDownTab(mshDownTab.Row, mshDownTab.Col, "1/C")
        
        strTmp = Trim(mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col))
        
        If picInput.Visible Then
'            If Right(strTmp, 2) <> "/C" And strTmp <> "" Then
'                If Val(strTmp) > 0 Then
'                    mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = Val(strTmp) & "/C"
'                Else
'                    mshDownTab.TextMatrix(mshDownTab.Row, mshDownTab.Col) = ""
'                End If
'            End If
            mshDownTab.SetFocus
        Else
            Call ShowInput
        End If
        If txtInput(0).Visible Then
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 0
        End If
        
        Call mshDownTab_RowColChange

    '------------------------------------------------------------------------------------------------------------------
    Case "显示病人姓名"
    
        Select Case Val(mrsParam("婴儿").Value)
        Case 0
            txtCard(0).Text = txtCard(0).Tag
            txtCard(7).Text = txtCard(7).Tag
        Case Else
            
            txtCard(5).Text = ""
            txtCard(6).Text = ""
            txtCard(7).Text = ""
            
            gstrSQL = "Select Decode(a.婴儿姓名,Null,b.姓名||'之子'||Trim(To_Char(a.序号,'9')),a.婴儿姓名) As 婴儿姓名,a.婴儿性别,a.出生时间 From 病人新生儿记录 a,病人信息 b Where a.病人id=[1] And a.主页id=[2] And a.病人id=b.病人id And a.序号=[3]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "usrBodyEditor", Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
            If rs.BOF = False Then
            
                txtCard(0).Text = rs("婴儿姓名").Value
                txtCard(5).Text = rs("婴儿性别").Value
                
                txtCard(6).Text = "新生儿"
'                If IsNull(rs("出生时间").Value) = False Then
'                    txtCard(6).Text = DateDiff("d", rs("出生时间").Value, zlDatabase.Currentdate) & "天"
'                End If
                
            End If
            
        End Select
    
    End Select
        
End Function

Public Function PrintState(ByVal intPrintRange As Integer, ByVal blnPrint As Boolean, Optional lngBeginY As Long, _
    Optional ByVal intPageNo As Integer = -1, Optional ByVal strPrintDevice As String) As Boolean
    '******************************************************************************************************************
    '功能:将当前体温表或当前开始的所有体温表输出到打印机上或预览窗体
    '参数:blnCurState = 是否为只打印当前体温表,否则打印从当前开始的所有体温表
    '     blnPrint    = 是否输出到打印机上否则输出到预览窗体里
    '******************************************************************************************************************
    
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim strPrintName As String
    Dim intPage As Integer
    Dim blnYesPrinter As Boolean
    Dim intCol As Integer
    Dim intBeginPage As Integer
    Dim intEndPage As Integer
'    Dim intPageNo As Integer
    Dim byeReturn As Byte
    Dim strArrFromTo() As String
    Dim intOrient As Integer
    Dim intBaby As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim lngIndex As Long
    
    On Error GoTo ErrHandle
    
    intBaby = Val(mrsParam("婴儿").Value)
    
    '------------------------------------------------------------------------------------------------------------------
    '打印机恢复及设置
    If Not ExistsPrinter Then
        MsgBox "系统没有安装任何打印机不能继续打印，程序退出！", vbInformation, gstrSysName
        Exit Function
    End If
    
    If strPrintDevice = "" Then
        If Trim(zlDatabase.GetPara("体温单打印机", glngSys, 1255, "")) = "" Then
            MsgBox "没有设置打印机,将使用系统默认打印机设置！", vbInformation, gstrSysName
        Else
            strPrintName = Trim(zlDatabase.GetPara("体温单打印机", glngSys, 1255, Printer.DeviceName))
            '打印机
            blnYesPrinter = False
            If Printer.DeviceName <> strPrintName Then
                For i = 0 To Printers.Count - 1
                    If Printers(i).DeviceName = strPrintName Then Set Printer = Printers(i): blnYesPrinter = True: Exit For
                Next
                If blnYesPrinter = False Then
                    MsgBox "设置的打印机已不存在,将使用系统默认打印机设置！", vbInformation, gstrSysName
                End If
            End If
        End If
    Else
        strPrintName = strPrintDevice
    End If
        
    intPage = Val(zlDatabase.GetPara("体温单纸张", glngSys, 1255, Printer.PaperSize))
    lngWidth = Val(zlDatabase.GetPara("体温单宽度", glngSys, 1255, Printer.Width))
    lngHeight = Val(zlDatabase.GetPara("体温单高度", glngSys, 1255, Printer.Height))
    lngLeft = Val(zlDatabase.GetPara("体温单左边距", glngSys, 1255, OFFSET_LEFT))
    lngTop = Val(zlDatabase.GetPara("体温单上边距", glngSys, 1255, OFFSET_TOP))
    intOrient = Val(zlDatabase.GetPara("体温单纸向", glngSys, 1255, Printer.Orientation))
    
    On Error Resume Next
    '纸张
    If intPage = 256 Then
        Printer.PaperSize = 256
        Printer.Width = lngWidth
        Printer.Height = lngHeight
    Else
        Printer.PaperSize = intPage
    End If
    Printer.Orientation = intOrient
    
    On Error GoTo ErrHandle
    
    '------------------------------------------------------------------------------------------------------------------
    lngBeginY = IIf(lngTop > lngBeginY, lngTop, lngBeginY)
    lngIndex = mlngNo
    
    
    '读取此病人的体温单总页数
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select 入院时间, 出院时间, 1 + Round((b.出院时间 - b.入院时间) / 7) As 页数" & vbNewLine & _
                "  from (Select Min(开始时间) as 入院时间," & vbNewLine & _
                "               Max(Nvl(终止时间, Sysdate)) as 出院时间" & vbNewLine & _
                "          From 病人变动记录" & vbNewLine & _
                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]) b"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    intCount = 0
    For intCol = 0 To rsTmp("页数").Value - 1
                
        strDateFrom = Format(rsTmp("入院时间").Value + 7 * intCol, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + 7 * (intCol + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")
            
            ReDim Preserve strArrFromTo(intCount)
            strArrFromTo(intCount) = "0;" & intCol + 1 & ";" & intCol + 1
            intCount = intCount + 1
        End If
    Next
        
    '如果只打印当前就只将开始和结束写同一页码
    Set mfrmCaseTendBodyPrint = New frmCaseTendBodyPrint
    Select Case intPrintRange
    Case 0                  '打印当前页
        
        If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), intBaby, _
                Val(mrsParam("科室id").Value), lngBeginY * 56.7, lngLeft, Me, False, _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, , mblnMoved) = True Then
                
                If blnPrint = False Then
                    mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("病人id")), Val(mrsParam("主页id")), _
                        Val(mrsParam("科室id").Value), CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                        CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, lngIndex
                Else
                    Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                    Printer.EndDoc
                End If
        Else
            MsgBox "未知错误，输出体温单失败！", vbExclamation, gstrSysName
        End If
        
    Case 1              '从当前页连续打印
    
        For intCol = lngIndex To UBound(strArrFromTo)
        
            If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("病人id")), Val(mrsParam("主页id")), intBaby, _
                Val(mrsParam("科室id").Value), lngBeginY * 56.7, lngLeft, Me, intCol <> lngIndex, _
                CInt(Split(strArrFromTo(intCol), ";")(1)), CInt(Split(strArrFromTo(intCol), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCol = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("病人id")), Val(mrsParam("主页id")), _
            Val(mrsParam("科室id").Value), CInt(Split(strArrFromTo(lngIndex), ";")(1)), _
                CInt(Split(strArrFromTo(lngIndex), ";")(1)), intPageNo, strArrFromTo, lngIndex
        End If
        
    Case 2          '从第一页连续打印,即全部打印
        
        For intCol = 0 To UBound(strArrFromTo)
        
            If PrintOrPreviewBodyState(mfrmCaseTendBodyPrint, Val(mrsParam("病人id")), Val(mrsParam("主页id")), intBaby, _
                Val(mrsParam("科室id").Value), lngBeginY * 56.7, lngLeft, Me, intCol <> 0, _
                CInt(Split(strArrFromTo(intCol), ";")(1)), CInt(Split(strArrFromTo(intCol), ";")(1)), intPageNo, , mblnMoved) = True Then
            Else
                MsgBox "未知错误，打印失败！", vbExclamation, gstrSysName
                Exit For
            End If
            
            If blnPrint Then
                Printer.PaintPicture mfrmCaseTendBodyPrint.picPage(mfrmCaseTendBodyPrint.picPage.UBound).Image, 0, 0
                If intCol = UBound(strArrFromTo) Then
                    Printer.EndDoc
                Else
                    Printer.NewPage
                End If
            End If
        Next

        If blnPrint = False Then
            mfrmCaseTendBodyPrint.Preview intPrintRange, lngBeginY, lngLeft, Me, Val(mrsParam("病人id")), Val(mrsParam("主页id")), _
            Val(mrsParam("科室id").Value), CInt(Split(strArrFromTo(0), ";")(1)), _
                CInt(Split(strArrFromTo(0), ";")(1)), intPageNo, strArrFromTo, 0
        End If
        
    End Select
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function AllowAudit() As Boolean
    '******************************************************************************************************************
    '功能：检查是否允许复查标志
    '参数：无
    '返回：
    '******************************************************************************************************************
    Dim intNowCol As Integer
    Dim aryValue As Variant

    If picScale.Tag <> "" Then

        With mshScale
            intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
            
            If Split(.TextMatrix(GraphDataRow.曲线数据, intNowCol), ";")(mItemSerial.体温 + 1) <> "" Then
                aryValue = Split(Split(.TextMatrix(GraphDataRow.曲线数据, intNowCol), ";")(mItemSerial.体温 + 1), ",")
                AllowAudit = (Val(.TextMatrix(GraphDataRow.复试标志, intNowCol)) = 0 And Val(aryValue(0)) > 0)
            End If

        End With
    End If
End Function

Public Function AllowUnAudit() As Boolean
    '******************************************************************************************************************
    '功能：检查是否允许撤消复查标志
    '参数：无
    '返回：
    '******************************************************************************************************************
    Dim intNowCol As Integer
    Dim aryValue As Variant
    
    If picScale.Tag <> "" Then

        With mshScale
            intNowCol = .FixedCols + lblCur.Left \ HOUR_STEP_Twips
            
            If Split(.TextMatrix(GraphDataRow.曲线数据, intNowCol), ";")(mItemSerial.体温 + 1) <> "" Then
                aryValue = Split(Split(.TextMatrix(GraphDataRow.曲线数据, intNowCol), ";")(mItemSerial.体温 + 1), ",")
                AllowUnAudit = (.TextMatrix(GraphDataRow.复试标志, intNowCol) = "1" And Val(aryValue(0)) > 0)
            End If
            
        End With
    End If
End Function

Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    '------------------------------------------------------------------------------------------------------------------
    '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义:包括公共部份，请对xtpControlPopup类型的命令ID重新赋值

    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义:包括公共部份
    
    Set mcbrToolBar = cbsMain.Add("婴儿", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap
    
    Set objCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Option, "")
    picTmp.Visible = True
    objCustom.Handle = picTmp.hWnd

End Function

Private Function InitBody(ByVal lng病人id As Long, ByVal lng主页id As Long, ByVal lng病区id As Long, ByVal int婴儿 As Integer) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim cbrItem As CommandBarControl
    Dim intCount As Integer
    Dim strDateFrom As String
    Dim strDateTo As String
    Dim strEnterDate As String
    Dim intCol As Integer
    Dim strCaption As String
    Dim strParameter As String
    Dim strSvrCaption As String
    Dim strNow As String
    Dim strCut As String
    Dim lngLoop As Long
    Dim strTmp As String
    Dim lnglast科室id As Long
    
    If lng病人id = 0 Then Exit Function
    strCut = "123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    '删除操作页面菜单项
    
    If Not mcbrToolBar页面 Is Nothing Then mcbrToolBar页面.Delete
    Set mcbrToolBar页面 = mcbrToolBar.Controls.Add(xtpControlPopup, conMenu_Edit_NewItem, "页面"):  mcbrToolBar页面.BeginGroup = True
    mcbrToolBar页面.IconId = conMenu_Edit_Modify
    mcbrToolBar页面.Style = xtpButtonIconAndCaption
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select Decode(c.出生时间,Null,b.入院时间,c.出生时间) As 入院时间, 出院时间, 1 + Round((b.出院时间 - Decode(c.出生时间,Null,b.入院时间,c.出生时间)) / 7) As 页数" & vbNewLine & _
                "  from (Select 病人ID,主页id,Min(开始时间) as 入院时间," & vbNewLine & _
                "               Max(Nvl(终止时间, Sysdate)) as 出院时间" & vbNewLine & _
                "          From 病人变动记录" & vbNewLine & _
                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2] Group By 病人ID,主页id) b," & vbNewLine & _
                "       (Select 病人ID,主页id,出生时间 From 病人新生儿记录 Where 病人ID = [1] And 主页ID = [2] And 序号=[3]) c Where b.病人id=c.病人id(+) And b.主页id=c.主页id(+)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng病人id, lng主页id, int婴儿)
    If rsTmp.BOF Then
        MsgBox "无病人本次住院记录！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    strEnterDate = Format(rsTmp!入院时间, "yyyy-MM-dd HH:mm:ss")

    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select 1 + Round((a.开始时间 - b.入院时间) / 7) As 开始页码,1 + Round((a.终止时间 - b.入院时间) / 7) As 结束页码,b.入院时间," & vbNewLine & _
                "       病区id,c.名称," & vbNewLine & _
                "       开始时间," & vbNewLine & _
                "       终止时间" & vbNewLine & _
                "  from (Select 病区id," & vbNewLine & _
                "               Min(开始时间) as 开始时间," & vbNewLine & _
                "               Max(Nvl(终止时间, Sysdate)) as 终止时间" & vbNewLine & _
                "          From 病人变动记录" & vbNewLine & _
                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2]" & vbNewLine & _
                "         Group by 病区id) a," & vbNewLine & _
                "       (Select Decode(y.出生时间,Null,x.入院时间,y.出生时间) As 入院时间 From (Select 病人ID,主页id,Min(开始时间) as 入院时间" & vbNewLine & _
                "          From 病人变动记录" & vbNewLine & _
                "         Where 开始时间 is Not Null And 病人ID = [1] And 主页ID = [2] Group By 病人id,主页id) x,(Select 病人ID,主页id,出生时间 From 病人新生儿记录 Where 病人ID = [1] And 主页ID = [2] And 序号=[3]) y Where x.病人id=y.病人id(+) And x.主页id=y.主页id(+) ) b,部门表 c Where c.ID=a.病区id " & vbNewLine & _
                " order by a.开始时间"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, "usrBodyEditor", lng病人id, lng主页id, int婴儿)
        
    For lngLoop = 0 To rsTmp("页数").Value - 1
                
        strDateFrom = Format(rsTmp("入院时间").Value + 7 * lngLoop, "yyyy-MM-dd") & " 00:00:00"
        strDateTo = Format(rsTmp("入院时间").Value + 7 * (lngLoop + 1) - 1, "yyyy-MM-dd") & " 23:59:59"
        If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
            strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
        End If
        
        If strDateFrom < Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then
        
            If strDateFrom < Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateFrom = Format(rsTmp("入院时间").Value, "yyyy-MM-dd HH:mm:ss")
            If strDateTo > Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss") Then strDateTo = Format(rsTmp("出院时间").Value, "yyyy-MM-dd HH:mm:ss")
    
            rs.Filter = ""
            rs.Filter = "开始页码<=" & lngLoop + 1 & " And 结束页码>=" & lngLoop + 1
            If rs.RecordCount > 0 Then rs.MoveFirst
            For intCol = 1 To rs.RecordCount
                
                If strDateFrom < Format(rs("开始时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strTmp = Format(rs("开始时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strTmp = strDateFrom
                End If
                
                If strDateTo > Format(rs("终止时间").Value, "yyyy-MM-dd HH:mm:ss") Then
                    strCaption = Format(rs("终止时间").Value, "yyyy-MM-dd HH:mm:ss")
                Else
                    strCaption = strDateTo
                End If
                
                strCaption = Format(strTmp, "yyyy-MM-dd") & "～" & Format(strCaption, "yyyy-MM-dd")
                strCaption = "第" & lngLoop + 1 & "页：" & strCaption & "(" & rs("名称").Value & ")"
                
                '入院时间;科室id;开始时间;结束时间;
                Set cbrItem = mcbrToolBar页面.CommandBar.Controls.Add(xtpControlButton, conMenu_View_Jump, strCaption, -1, False)
                cbrItem.Parameter = strEnterDate & ";" & rs!病区ID & ";" & strDateFrom & ";" & strDateTo & ";" & lngLoop
                
                lnglast科室id = rs("病区ID").Value
                
                rs.MoveNext
                
                strParameter = cbrItem.Parameter
                strSvrCaption = strCaption
            Next
        End If
        
    Next
    
    Call picPane_Resize
    
    If strParameter <> "" Then
        mcbrToolBar页面.Caption = strSvrCaption
        Call zlMenuClick("装载数据", strParameter)
    End If
    
    InitBody = True
End Function


Private Function CheckTimeRange(ByVal intCol As Integer) As Boolean
    Dim strTime As String
    Dim strFrom As String
    Dim strTo As String

    Dim strEnd As String
    Dim strStart As String
    
    If picScale.Tag = "" Then Exit Function
    If InStr(picScale.Tag, ";") = 0 Then Exit Function
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    If strTo > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then strTo = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    strTime = GetEditDateTime(intCol - mshDownTab.FixedCols + 1, CDate(strFrom))
    strStart = Split(strTime, ",")(0)
    strEnd = Split(strTime, ",")(1)
    
    CheckTimeRange = False
    
    If strStart <= strFrom And strEnd >= strTo Then
        CheckTimeRange = True
    End If
    
    If strStart >= strFrom And strStart <= strTo Then
        CheckTimeRange = True
    End If
    
    If strEnd > strFrom And strEnd < strTo Then
        CheckTimeRange = True
    End If

End Function

Private Function GetTextPos(ByVal lngHwnd As Long) As Long

    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngFirst As Long
    
    lngFirst = SendMessage(lngHwnd, EM_GETFIRSTVISIBLELINE, lngRow, lngCol) + 1 '以0行开始
    Call GetCaretPos(lngHwnd, lngRow, lngCol)
    GetTextPos = lngCol
    
End Function

Private Sub GetCaretPos(ByVal TextHwnd As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long, k As Long
    Dim lParam As Long, wParam As Long

    '首先向文本框传递EM_GETSEL消息以获取从起始位置到
    '光标所在位置的字符数
    i = SendMessage(TextHwnd, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16
    
    '再向文本框传递EM_LINEFROMCHAR消息根据获得的字符
    '数确定光标以获取所在行数
    LineNo = SendMessage(TextHwnd, EM_LINEFROMCHAR, j, 0) '
    LineNo = LineNo + 1
    
    '向文本框传递EM_LINEINDEX消息以获取所在列数
    k = SendMessage(TextHwnd, EM_LINEINDEX, -1, 0)
    ColNo = j - k + 1
End Sub

Private Function ReadDrink(ByVal strDate As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim intCol As Long
    Dim strTmp As String
    Dim strFrom As String
    
    Dim strStart As String
    Dim strEnd As String
    Dim lng饮入物id As Long
    Dim lng饮入量id As Long
    Dim strValue As String
    Dim int记录法 As Integer
    Dim intMax As Integer
    Dim lngCol As Long
    
    On Error GoTo errHand
    
    strFrom = CStr(mrsParam("开始时间"))
    
    strSQL = "Select A.项目id From 护理记录项目 A Where A.项目序号=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, 6)
    If rs.BOF = False Then lng饮入物id = zlCommFun.NVL(rs("项目id"), 0)
    
    strSQL = "Select A.项目id,B.记录法 From 护理记录项目 A,体温记录项目 B Where A.项目序号=B.项目序号 AND A.项目序号=[1] "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, 7)
    If rs.BOF = False Then
        lng饮入量id = zlCommFun.NVL(rs("项目id"), 0)
        int记录法 = zlCommFun.NVL(rs("记录法"), 1)
    End If
    
    If int记录法 = 1 Then
        intMax = 6
    Else
        intMax = 2
    End If
    
    For intCol = 0 To intMax - 1
            
        strStart = Format(Int(CDate(strDate)) + intCol / intMax - (4 - mlngHourBegin) / 24, "YYYY-MM-DD hh:mm:ss")
        strEnd = Format(Int(CDate(strDate)) + intCol / intMax - (4 - mlngHourBegin) / 24 + 1 / intMax, "YYYY-MM-DD hh:mm:ss")
        
        If Int(CDate(strStart)) < Int(CDate(strDate)) Then
            strStart = Format(strDate, "yyyy-MM-dd HH:mm:ss")
        End If
        
        strSQL = "Select zl_PatitDrink([1],[2],[3],[4]) As 饮入 From Dual"
        
        Set rs = zlDatabase.OpenSQLRecord(strSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(strStart), CDate(strEnd))
        If rs.BOF = False Then
            
            strTmp = zlCommFun.NVL(rs("饮入"))
            
            If strTmp <> "" Then
                strValue = Trim(Split(strTmp, ";")(0))
                
                mstrSQL = "ZL_电子护理记录_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & "7,"
                mstrSQL = mstrSQL & "0,"
                    
                mstrSQL = mstrSQL & IIf(Val(strValue) = 0, "NULL", "'" & Val(strValue) & "'")
                
                mstrSQL = mstrSQL & ")"

                Call zlDatabase.ExecuteProcedure(mstrSQL, mstrMsgTitle)
                
                If mItemSerial.饮入量 >= 0 Then
                    
                    If int记录法 = 2 Then
                        
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshDownTab.FixedCols
                        
                        Call WriteDownTab(mItemSerial.饮入量, lngCol, strValue)
                    Else
                                                
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshScale.FixedCols
                        
                        Call WriteScaleTab(mItemSerial.饮入量, lngCol, strValue)
                    End If
                    
                End If
                
                strValue = Trim(Split(strTmp, ";")(1))
                If UBound(Split(strTmp, ";")) > 1 Then strValue = strValue & "等"
                            
                mstrSQL = "ZL_电子护理记录_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & "6,"
                mstrSQL = mstrSQL & "0,"
                    
                mstrSQL = mstrSQL & IIf(strValue = "", "NULL", "'" & strValue & "'")
                
                mstrSQL = mstrSQL & ")"

                Call zlDatabase.ExecuteProcedure(mstrSQL, mstrMsgTitle)
                
                If mItemSerial.饮入物 >= 0 Then
                    If int记录法 = 2 Then
                    
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshDownTab.FixedCols
                        Call WriteDownTab(mItemSerial.饮入物, lngCol, strValue)
                        
                    Else
                        
                        lngCol = intCol + (Int(CDate(strStart)) - Int(CDate(strFrom)) + (4 - mlngHourBegin) / 24) * intMax + mshScale.FixedCols
                        Call WriteScaleTab(mItemSerial.饮入量, lngCol, strValue)
                        
                    End If
                End If
                
            End If
        End If
    Next
        
    ReadDrink = True
                            
    Exit Function
    
errHand:
'    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回： 调用成功返回TRUE；否则FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (pic.Width - picPane.Width) / 100
    msinVStep = (pic.Height - picPane.Height) / 100
    
    hsb.Max = 0 - Int(0 - ((pic.Width - picPane.Width) / 300))
    vsb.Max = 0 - Int(0 - ((pic.Height - picPane.Height) / 300))
    hsb.Enabled = (hsb.Max > 0)
    vsb.Enabled = (vsb.Max > 0)
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then hsb.Max = 100
    If vsb.Enabled Then vsb.Max = 100
    
    CalcScrollBarSize = True
    
End Function

Private Function Check是否包含(strSource As String, strTarge As String) As Boolean
    '检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    Check是否包含 = False
    
    Select Case strTarge
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    Check是否包含 = True
End Function

Private Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
    '功能: 清除指定网格的指定行指定列的数据
    '参数: obj=要操作的网格控件
    '      intRow=要清除的行号
    '      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Private Sub SetColumnText(fgd As Object, intRow As Integer, ByVal varColText As Variant)
    '功能: 设置指定网格控件的列头文本
    '参数: fgd=网格控件
    '      intRow=行号
    '      varColText=列头文本数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.TextMatrix(intRow, i) = varColText(i)
    Next
End Sub

Private Sub SetColAlignment(fgd As Object, varColAlignment As Variant)
    '功能: 设置指定网格控件的列对齐方式
    '参数: fgd=网格控件
    '      varColAlignment=列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varColAlignment)
        fgd.ColAlignment(i) = varColAlignment(i)
    Next
End Sub

Private Sub SetColData(fgd As Object, varColData As Variant)
    '功能: 设置指定网格控件的列数据来源方式
    '参数: fgd=网格控件
    '      varColData=列数据来源方式数组
    Dim i As Long
    For i = 0 To UBound(varColData)
        fgd.ColData(i) = varColData(i)
    Next
End Sub

Private Sub SetFixColAlignment(fgd As Object, varFixColAlignment As Variant)
    '功能: 设置指定网格控件的固定列对齐方式
    '参数: fgd=网格控件
    '      varColAlignment=固定列对齐方式数组
    Dim i As Long
    For i = 0 To UBound(varFixColAlignment)
        fgd.ColAlignmentFixed(i) = varFixColAlignment(i)
    Next
End Sub

Private Sub SetColumnWidth(fgd As Object, ByVal varColWidth As Variant)
    '功能: 设置指定网格控件的列宽
    '参数: fgd=网格控件
    '      varColWidth=列宽数组
    Dim i As Integer
    For i = 0 To fgd.Cols - 1
        fgd.ColWidth(i) = varColWidth(i)
    Next
End Sub

Public Function SetDispMode(Optional blnReadOnly As Boolean) As Boolean
    
    '用来设置体温表当前是编辑模式还是显示模式
    
    Call SetBodyMode
    
End Function


Private Function InitData() As Boolean
    '******************************************************************************************************************
    '功能：用来调用体温单的接口函数
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    '------------------------------------------------------------------------------------------------------------------
    '变量初始化
    
    mlngLine = 0
    mlngPageCur = 1
    
    '------------------------------------------------------------------------------------------------------------------
    With pic
        .Left = 0
        .Top = 0
        .Width = UserControl.Width
        .Height = UserControl.Height
    End With
    
    mstrMsgTitle = "体温表"
    UserControl.BackColor = RGB(255, 255, 255) '将背景置白色
    mblnNoneShow = False
    
    
    
    '读取体温表一天开始时间
    '------------------------------------------------------------------------------------------------------------------
    mlngHourBegin = zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4)
    mbln婴儿体温单显示出院 = (zlDatabase.GetPara("婴儿体温单显示出院信息", glngSys, 1255, 1) = 1)
    
    '病人变动标记显示方法
    '------------------------------------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("体温单标记", glngSys, 1255, "1;1;1;1;1;1;1;1")
    If UBound(Split(strTmp, ";")) >= 5 Then
        mBodyFlag.入院 = Val(Split(strTmp, ";")(0))
        mBodyFlag.入科 = Val(Split(strTmp, ";")(1))
        mBodyFlag.转出 = Val(Split(strTmp, ";")(2))
        mBodyFlag.换床 = Val(Split(strTmp, ";")(3))
        mBodyFlag.手术 = Val(Split(strTmp, ";")(4))
        mBodyFlag.出院 = Val(Split(strTmp, ";")(5))
        If UBound(Split(strTmp, ";")) >= 6 Then mBodyFlag.分娩 = Val(Split(strTmp, ";")(6))
        If UBound(Split(strTmp, ";")) >= 7 Then mBodyFlag.出生 = Val(Split(strTmp, ";")(7))
    End If
    
    '读取护理等级
    '------------------------------------------------------------------------------------------------------------------
    mrsParam("护理等级").Value = 3
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rs.BOF = False Then
        mrsParam("护理等级").Value = zlCommFun.NVL(rs("护理等级"), 3)
    End If
    
    '检查是否有曲线体温项目
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = " Select 1 From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
                "Where C.项目序号=A.项目序号 " & _
                        "AND C.项目ID=B.ID(+) " & _
                        "AND C.护理等级>=[1] " & _
                        "And A.记录法=1 And RowNum<2 And C.项目序号<>-1 "
                
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("护理等级")))
    If rs.EOF Then
        ShowSimpleMsg "至少要有一个已记录的曲线项目！"
        
        If Val(mrsParam("编辑")) = 0 Then
            mblnNoneShow = True
            Exit Function   '显示模式不允许增加项目
        End If
    End If
    
    '判断病人是否已转出
    '因为该函数内外都在调用,参数不好变,直接读取
    '------------------------------------------------------------------------------------------------------------------
    mblnMoved = False
    If Val(mrsParam("病人id")) > 0 And Val(mrsParam("出院")) = 1 Then
        mstrSQL = "Select 数据转出 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
        mblnMoved = NVL(rs!数据转出, 0) <> 0
    End If
    If mblnMoved Or Val(mrsParam("编辑")) = 0 Then Call SetDispMode(True)
    
    
    vsf.Body.Appearance = flexFlat
    vsf.Body.RowHidden(0) = True
    vsf.Body.ColHidden(0) = True
    vsf.Body.ScrollBars = flexScrollBarNone
    vsf.Body.BorderStyle = flexBorderNone
    vsf.FixedCols = 1
    
    vsf.Rows = 2
    
    InitData = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ReadPatiInfo() As Boolean
    '******************************************************************************************************************
    '功能： 提取当前病人mlng病人ID的住院情况（住院变动记录），整体为若干页体温表
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHead
    
    If Val(mrsParam("病人id")) = 0 Then Exit Function
    If Val(mrsParam("主页id")) = 0 Then Exit Function
    
    '填写病人姓名、住院号
    gstrSQL = "Select A.姓名,B.住院号 From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And B.病人id=[1] And B.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rsTmp.BOF Then
        ShowSimpleMsg "指定的病人不存在！"
        Exit Function
    End If
    
    txtCard(0).Tag = zlCommFun.NVL(rsTmp("姓名").Value)
    
    Call zlMenuClick("显示病人姓名")
    
    txtCard(0).Tag = Val(mrsParam("病人id"))
    txtCard(1).Text = zlCommFun.NVL(rsTmp("住院号").Value)
    txtCard(1).Tag = Val(mrsParam("婴儿").Value)
    
    gstrSQL = "Select Zl_Replace_Element_Value([1],[2],[3],2) As 最后诊断 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, "最后诊断", Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rsTmp.BOF = False Then
        If Val(mrsParam("婴儿").Value) = 0 Then
            txtCard(7).Text = zlCommFun.NVL(rsTmp("最后诊断").Value)
        Else
            txtCard(7).Text = ""
        End If
    Else
        txtCard(7).Text = ""
    End If
    txtCard(7).Tag = txtCard(7).Text
    
    mstrSQL = " Select D.ID,D.名称,开始,终止" & _
                " From 部门表 D," & _
                "   (Select 病区id,Min(开始时间) as 开始,Max(Nvl(终止时间,Sysdate)) as 终止" & _
                "    From 病人变动记录" & _
                "    Where 开始时间 is Not Null And 病人ID=[1] And 主页ID=[2]" & _
                "    Group by 病区id) L" & _
                " Where L.病区id=D.ID" & _
                " Order by 开始"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rsTmp.BOF Then
        
        'ShowSimpleMsg "无病人本次住院记录！"
        mblnNoneShow = True
        
        Exit Function
    End If
        
    ReadPatiInfo = True
    
    Exit Function
    
None:
    
    SetVisible
    Exit Function
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ClearLineSelect() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 清除当前选择的曲线项目
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim intLoop As Integer
    
    picGraph.Tag = ""
    picGraph.MousePointer = 0
    linHCur.Visible = False
    linVCur.Visible = False
    
    ClearLineSelect = True
    
End Function

Private Function AddCrlf(ByVal strText As String) As String
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Dim intLoop As Integer
    Dim strTmp As String
    
    For intLoop = 1 To Len(strText)
        strTmp = strTmp & Mid(strText, intLoop, 1) & vbCrLf
    Next
    
    AddCrlf = strTmp
    
End Function

Private Function FaceInit() As Boolean
    '******************************************************************************************************************
    '功能： 根据体温表设置，调整体温表的布局
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim i As Long
    Dim rs As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo errHand
    
    mItemSerial.体温 = -1
    mItemSerial.脉搏 = -1
    mItemSerial.呼吸 = -1
    mItemSerial.心率 = -1

'    Erase mvarDataType
    
    lblComment.Caption = "说明:"
    mbln呼吸曲线 = False
    mstr最小时间 = ""
    Call Get入院入科时间
    
    If mblnNoneShow Then Exit Function
    
    '其它初始化
    mshUpTab.Rows = 3
    mshUpTab.Cell(flexcpAlignment, 0, 0, mshUpTab.Rows - 1, 0) = 4
    mshUpTab.Cell(flexcpText, 0, mshUpTab.FixedCols, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = ""
    mshUpTab.Cell(flexcpData, 0, mshUpTab.FixedCols, mshUpTab.Rows - 1, mshUpTab.Cols - 1) = ""
    mshUpTab.Cell(flexcpForeColor, 0, mshUpTab.FixedCols, 1, mshUpTab.Cols - 1) = 16711680
    mshUpTab.Cell(flexcpForeColor, 2, mshUpTab.FixedCols, 2, mshUpTab.Cols - 1) = 255

    mshDownTab.RowHidden(0) = True
    
    mshUpTab.Redraw = False
    mshScale.Redraw = False
    mshDownTab.Redraw = False
    mvarEdit = False
    If picLine.Count > 1 Then
        For i = 1 To picLine.Count - 1
            Unload picLine(i)
        Next
    End If
    
    '读取本地打印开始页号
    UserControl.BackColor = RGB(255, 255, 255)
    Call ClearSpecRowCol(mshScale, 0, Array())
    
    '为避免在picture上绘出的线条不完整，先将picScale和picGraph设置为最大
    picScale.Width = Screen.Width
    picScale.Height = Screen.Height
    picGraph.Left = 0
    picGraph.Top = 0
    picGraph.Width = Screen.Width
    picGraph.Height = Screen.Height
    lblCur.Top = 350
    
    '根据项目设置调整显示内容
    
    mItemSerial.饮入量 = -1
    mItemSerial.饮入物 = -1
    mItemNo.出液 = 0
    mItemNo.大便 = 0
    mItemNo.心率 = 0
    mItemNo.脉搏 = 0
    mItemNo.体温 = 0
    
    mstrSQL = " Select 项目序号 From 体温记录项目 A Where A.记录名=[1]"
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "心率")
    If rs.BOF = False Then
        mItemNo.心率 = rs("项目序号").Value
    End If

    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "脉搏")
    If rs.BOF = False Then
        mItemNo.脉搏 = rs("项目序号").Value
    End If
        
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "体温")
    If rs.BOF = False Then
        mItemNo.体温 = rs("项目序号").Value
    End If
        
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, "呼吸")
    If rs.BOF = False Then
        mItemNo.呼吸 = rs("项目序号").Value
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    '求出最高行
    mstrSQL = " Select Max((A.最大值-A.最小值)/Decode(A.单位值,0,1,A.单位值)+A.最高行) From 体温记录项目 A Where A.记录法=1 "
    Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle)
    If rs.BOF = False Then
        If IsNull(rs.Fields(0).Value) = False Then
            mshDownTab.Tag = "1"
            mvarEdit = True
            
            mshScale.Rows = MAXROWS
            
            mint心率应用 = 2
            mstr心率符号 = ""
            mstrSQL = "Select a.应用方式,b.记录符 From 护理记录项目 a,体温记录项目 b Where a.项目序号=-1 And a.项目序号=b.项目序号"
            Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle)
            If rs.BOF = False Then
                mint心率应用 = zlCommFun.NVL(rs("应用方式").Value, 2)
                mstr心率符号 = zlCommFun.NVL(rs("记录符").Value, "○")
            End If
            
            '得到所有曲线项目
                        
            mstrSQL = " Select A.记录法,A.记录名 as 项目名,A.项目序号 as 项目号,Nvl(B.ID,0) as 项目ID," & _
                        " C.项目单位 As 单位,记录符,最小值,最大值,记录色,1 as 记录否,单位值,最高行,Nvl(B.类型,1) as 存储类型 " & _
                        " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
                        " Where c.项目ID=B.ID(+) And A.项目序号=C.项目序号 And A.记录法=1 And Nvl(C.应用方式,0)=1 AND C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
                        " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2]))) " & _
                        " Order by A.排列序号"
                        
            Set rs = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("护理等级").Value), Val(mrsParam("科室id").Value), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2))
            If rs.RecordCount > 0 Then rs.MoveFirst
            
            'mshScale列数=固定列 + 7天 * 6格
            
            mshScale.Cols = rs.RecordCount + (mshUpTab.Cols - 1) * 6
            mshScale.FixedCols = rs.RecordCount
            mshScale.RowHeightMin = ROWHEIGHT
            
            mshScale.Tag = ""
            Do While Not rs.EOF
                
                If rs!项目名 = "呼吸" Then mbln呼吸曲线 = True
                
                If zlCommFun.NVL(rs!项目号, 0) = 7 Then mItemSerial.饮入量 = rs.AbsolutePosition - 1
                If zlCommFun.NVL(rs!项目号, 0) = 6 Then mItemSerial.饮入物 = rs.AbsolutePosition - 1
                
                If rs.AbsolutePosition > picLine.Count Then
                    Load picLine(rs.AbsolutePosition - 1)
                End If
                picLine(rs.AbsolutePosition - 1).Tag = rs!最大值 & ";" & rs!最小值 & ";" & rs!单位值 & ";" & rs!最高行
                picLine(rs.AbsolutePosition - 1).Visible = True
                picLine(rs.AbsolutePosition - 1).ZOrder
                
                If rs!项目名 = "体温" Then mItemSerial.体温 = rs.AbsolutePosition - 1
                If rs!项目名 = "脉搏" Then mItemSerial.脉搏 = rs.AbsolutePosition - 1
                If rs!项目名 = "呼吸" Then mItemSerial.呼吸 = rs.AbsolutePosition - 1
                If rs!项目名 = "心率" Then mItemSerial.心率 = rs.AbsolutePosition - 1
                
                '设置表格内项目
                mshScale.ColWidth(rs.AbsolutePosition - 1) = IIf(mshScale.FixedCols < 4, 1200 / mshScale.FixedCols, 450)
                mshScale.ColData(rs.AbsolutePosition - 1) = Val(rs("项目号").Value)
                If zlCommFun.NVL(rs("单位").Value) <> "" Then
                    mshScale.TextMatrix(0, rs.AbsolutePosition - 1) = rs("项目名").Value & " (" & zlCommFun.NVL(rs("单位").Value) & ")"
                Else
                    mshScale.TextMatrix(0, rs.AbsolutePosition - 1) = rs("项目名").Value
                End If
                mshScale.Cell(flexcpAlignment, 0, rs.AbsolutePosition - 1) = flexAlignCenterTop

                mshScale.Row = 0
                mshScale.Col = rs.AbsolutePosition - 1
                mshScale.CellForeColor = rs!记录色
                If mItemSerial.体温 = rs.AbsolutePosition - 1 Then
                    mshScale.Tag = mshScale.Tag & " "
                Else
                    mshScale.Tag = mshScale.Tag & zlCommFun.NVL(rs!记录符, " ")
                End If
                
                mstrChar(0) = ""
                mstrChar(0) = ""
                mstrChar(0) = ""
                If mItemSerial.体温 = rs.AbsolutePosition - 1 Then

                    Dim varTmp As Variant
                                        
                    gstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, mstrMsgTitle, 1)
                    If rsTmp.BOF = False Then
                        varTmp = Split(zlCommFun.NVL(rsTmp("记录符").Value, "·,×,○"), ",")
                    Else
                        varTmp = Split("·,×,○", ",")
                    End If
                    mstrChar(0) = CStr(varTmp(0))
                    mstrChar(1) = CStr(varTmp(1))
                    mstrChar(2) = CStr(varTmp(2))
        
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "、") & rs!项目名 & "(口温" & mstrChar(0) & ",腋温" & mstrChar(1) & ",肛温" & mstrChar(2) & ")"
                ElseIf mItemSerial.呼吸 = rs.AbsolutePosition - 1 Then
                    mstrBreath = rs!记录符
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "、") & rs!项目名 & "(自主呼吸" & rs!记录符 & ",呼吸机R)"
                ElseIf mItemSerial.脉搏 = rs.AbsolutePosition - 1 Then
                    mstrPulse = rs!记录符
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "、") & rs!项目名 & "(缺省记录符" & rs!记录符 & ",起搏器H)"
                Else
                    lblComment.Caption = lblComment.Caption & IIf(rs.AbsolutePosition = 1, "", "、") & rs!项目名 & "(" & rs!记录符 & ")"
                End If
                
                For intRow = 1 To mshScale.Rows - 1
                    mshScale.Row = intRow
                    mshScale.CellForeColor = rs!记录色
                    mshScale.ROWHEIGHT(intRow) = ROWHEIGHT * 5

                    If intRow >= rs!最高行 And rs!最大值 - (intRow - rs!最高行) * rs!单位值 >= rs!最小值 Then
                    
                        '刚好为整数时求出下一刻度值
                        
                        If Int(rs("最大值").Value - (intRow - rs!最高行) * rs("单位值").Value) = rs!最大值 - (intRow - rs!最高行) * rs!单位值 Then
                        
                            Select Case rs!项目名
                            Case "脉搏", "心率", "呼吸"
                                If (rs!最大值 - (intRow - rs!最高行) * rs!单位值) Mod 10 = 0 Then
                                    mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = rs!最大值 - (intRow - rs!最高行) * rs!单位值
                                End If
                            Case "体温"
                                mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = CStr(rs!最大值 - (intRow - rs!最高行) * rs!单位值) & "°"
                            'Case "呼吸"
                            '    mshScale.TextMatrix(intRow, rs.AbsolutePosition - 1) = rs!最大值 - (intRow - rs!最高行) * rs!单位值
                            End Select
                            
                        End If
                    End If
                Next
                rs.MoveNext
            Loop
        End If
    End If
    
    mshScale.Cell(flexcpAlignment, 0, 0, mshScale.Rows - 1, mshScale.FixedCols - 1) = 4
    
    With vsf
        .Cols = 0
        .NewColumn "", 0, 1
        .NewColumn "项目", mshScale.FixedCols * mshScale.ColWidth(0) + 15, 1
        For intCol = 1 To 42
            .NewColumn intCol, HOUR_STEP_Twips, 1, , 1
        Next
        .FixedCols = 2
        .Cell(flexcpAlignment, 1, 1) = flexAlignCenterCenter
        .Cell(flexcpFontName, 1, 2, 1, .Cols - 1) = "Times New Roman"
        .Cell(flexcpFontSize, 1, 2, 1, .Cols - 1) = 7.5
        .Body.Select 1, 1
        .Body.CellBorder 0, 1, 0, 0, 0, 0, 0
        .Body.Select 1, vsf.Cols - 1
        .Body.CellBorder 0, 0, 0, 1, 0, 0, 0
        .Body.BackColorFixed = .Body.BackColor
        
        For intCol = 3 To .Cols - 1 Step 2
            .Cell(flexcpBackColor, 1, intCol, 1, intCol) = &HF7ECE6
        Next
        
    End With
    
    '初始化直接录入项目,mshDownTab的前面四列用于记录项目有关属性:显示名称;项目名称;最大值;最小值
    '------------------------------------------------------------------------------------------------------------------
    Set rsTmp = GetGridItem(Val(mrsParam("护理等级").Value), Val(mrsParam("科室id").Value), IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2), 1)
    With rsTmp
        If rsTmp.RecordCount > 0 Then
            rsTmp.Sort = "排列序号"
            rsTmp.MoveFirst
            
            mshDownTab.Tag = "1"

            '初始数组
            ReDim mItemStru(0 To 1)
            mshDownTab.Rows = 1
            
            For i = 1 To rsTmp.RecordCount
                
                Call AppendGridItem(rsTmp)

                rsTmp.MoveNext
            Next
        Else
            mshDownTab.Rows = 2
        End If
    End With
   
   '------------------------------------------------------------------------------------------------------------------
    '找到系统特殊项目的项目号
    '设置三个系统固定的项目
    mshUpTab.RowData(0) = -1
    mshUpTab.RowData(1) = -2
    mshUpTab.RowData(2) = -3
    
    '初始上面表格
    With mshUpTab
        .ColWidth(0) = mshScale.FixedCols * mshScale.ColWidth(0)
        .TextMatrix(0, 0) = "日    期"
        .TextMatrix(1, 0) = "住院天数"
        .TextMatrix(2, 0) = "术/娩后日数"
        For intCol = 1 To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips * 6
        Next
    End With
    
    '初始数据表格体
    With mshScale
        .ROWHEIGHT(0) = 400 '600
        For intCol = .FixedCols To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips
        Next
    End With
    
    '初始下面的表格
    With mshDownTab
        .ColWidth(0) = mshScale.FixedCols * mshScale.ColWidth(0)
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        
        For intCol = .FixedCols To .Cols - 1
            .ColWidth(intCol) = HOUR_STEP_Twips * 3
            .ColAlignment(intCol) = 1
        Next
        
        
    End With
    
    '再根据读出的数据画图
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    mshDownTab.Redraw = True
    
    Call ReadPatiInfo
    Call SetVisible
    
    FaceInit = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AppendGridItem(ByVal rsTmp As ADODB.Recordset, Optional ByVal blnAppend As Boolean) As Boolean
    '******************************************************************************************************************
    '功能：填写表格项目的标题等
    '参数：rsTmp：要添加的表格项目
    '返回：
    '******************************************************************************************************************
    Dim intTmp As Integer
    Dim intRow As Integer
    
    On Error GoTo errHand
            
    Select Case rsTmp("名称").Value
    '------------------------------------------------------------------------------------------------------------------
    Case "呼吸"
        
        vsf.TextMatrix(1, 1) = "呼吸" & IIf(Not IsNull(rsTmp!单位), "(" & rsTmp!单位 & ")", "")
        
        mItemOtherStru(0).项目名称 = zlCommFun.NVL(rsTmp!名称, "")
        mItemOtherStru(0).项目序号 = zlCommFun.NVL(rsTmp!项目号, 0)
        mItemOtherStru(0).数据类型 = zlCommFun.NVL(rsTmp!存储类型, 1)
        mItemOtherStru(0).数据长度 = zlCommFun.NVL(rsTmp!项目长度, 0)
        mItemOtherStru(0).小数位数 = zlCommFun.NVL(rsTmp!项目小数, 0)
        mItemOtherStru(0).最小值 = zlCommFun.NVL(rsTmp!最小值, "")
        mItemOtherStru(0).最大值 = zlCommFun.NVL(rsTmp!最大值, "")
        mItemOtherStru(0).记录频次 = zlCommFun.NVL(rsTmp!记录频次, 0)
        mItemOtherStru(0).活动项目 = False
        
        mItemNo.呼吸 = zlCommFun.NVL(rsTmp!项目号, 0)
    '------------------------------------------------------------------------------------------------------------------
    Case "舒张压"
        
        mItemNo.舒张压 = zlCommFun.NVL(rsTmp!项目号, 0)
        
        mItemOtherStru(1).项目名称 = zlCommFun.NVL(rsTmp!名称, "")
        mItemOtherStru(1).项目序号 = zlCommFun.NVL(rsTmp!项目号, 0)
        mItemOtherStru(1).数据类型 = zlCommFun.NVL(rsTmp!存储类型, 1)
        mItemOtherStru(1).数据长度 = zlCommFun.NVL(rsTmp!项目长度, 0)
        mItemOtherStru(1).小数位数 = zlCommFun.NVL(rsTmp!项目小数, 0)
        mItemOtherStru(1).最小值 = zlCommFun.NVL(rsTmp!最小值, "")
        mItemOtherStru(1).最大值 = zlCommFun.NVL(rsTmp!最大值, "")
        mItemOtherStru(1).记录频次 = zlCommFun.NVL(rsTmp!记录频次, 0)
        mItemOtherStru(1).活动项目 = False
                
    '------------------------------------------------------------------------------------------------------------------
    Case Else
                
        mshDownTab.Rows = mshDownTab.Rows + 1
        intRow = mshDownTab.Rows - 1
        
        If rsTmp("名称").Value = "收缩压" Then
            mItemNo.血压 = zlCommFun.NVL(rsTmp!项目号, 0)
            mItemSerial.血压 = intRow
            mshDownTab.TextMatrix(intRow, 0) = "血压" & IIf(Not IsNull(rsTmp!单位), "(" & rsTmp!单位 & ")", "")
        Else
            mshDownTab.TextMatrix(intRow, 0) = zlCommFun.NVL(rsTmp!名称, "") & IIf(Not IsNull(rsTmp!单位), "(" & rsTmp!单位 & ")", "")
        End If
        
        mshDownTab.RowData(intRow) = rsTmp("项目号").Value
        mshDownTab.ROWHEIGHT(intRow) = 255
        mshDownTab.TextMatrix(intRow, 1) = mshDownTab.TextMatrix(intRow, 0)
        mshDownTab.TextMatrix(intRow, 2) = zlCommFun.NVL(rsTmp!最大值, "")
        mshDownTab.TextMatrix(intRow, 3) = zlCommFun.NVL(rsTmp!最小值, "")
                
        ReDim Preserve mItemStru(intRow)
        
        mItemStru(intRow).项目名称 = zlCommFun.NVL(rsTmp!名称, "")
        mItemStru(intRow).项目序号 = zlCommFun.NVL(rsTmp!项目号, 0)
        mItemStru(intRow).数据类型 = zlCommFun.NVL(rsTmp!存储类型, 1)
        mItemStru(intRow).数据长度 = zlCommFun.NVL(rsTmp!项目长度, 0)
        mItemStru(intRow).小数位数 = zlCommFun.NVL(rsTmp!项目小数, 0)
        mItemStru(intRow).最小值 = zlCommFun.NVL(rsTmp!最小值, "")
        mItemStru(intRow).最大值 = zlCommFun.NVL(rsTmp!最大值, "")
        mItemStru(intRow).记录频次 = zlCommFun.NVL(rsTmp!记录频次, 0)
        mItemStru(intRow).活动项目 = (zlCommFun.NVL(rsTmp!项目性质, 1) = 2)
        
        Select Case zlCommFun.NVL(rsTmp!项目号, 0)
        Case 6
            mItemSerial.饮入物 = intRow
        Case 7
            mItemSerial.饮入量 = intRow
        Case 9
            mItemNo.出液 = 9
        Case 10
            mItemNo.大便 = 10
        End Select
                    
    End Select
    
    '如果是后来添加表格项目，则增加处理修改标志
    If blnAppend Then

        With mshDownTab
            For intCol = .FixedCols To .Cols - 1
                .TextMatrix(GridDataRow.修改标志, intCol) = .TextMatrix(GridDataRow.修改标志, intCol) & ";"
            Next
            
        End With
        
    End If
    
    AppendGridItem = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ShowOpsDays() As Boolean
    '******************************************************************************************************************
    '显示当前区域段内的手术日标记
    '******************************************************************************************************************
    Dim lng次数 As Long
    Dim intCol As Integer
    Dim intLoop As Integer
    Dim strTmp As String
    Dim rsTmp As New ADODB.Recordset
    Dim strFrom As String
    
    On Error GoTo errHand
    
    '先清除当前区划内的手术表现标记
    
        
    For intCol = 1 To mshUpTab.Cols - 1
        mshUpTab.TextMatrix(2, intCol) = mstrOpsSvr(intCol)
    Next
    
    strFrom = Split(picScale.Tag, ";")(0)
    
    '找开始日期-14天前的手术次数
    mstrSQL = "SELECT Nvl(Count(a.发生时间),0) As 次数 " & _
                "FROM 病人护理记录 a,病人护理内容 c " & _
                "Where a.ID = c.记录ID " & _
                    "AND a.病人来源=2 " & _
                    "AND Nvl(a.婴儿,0)=[4] " & _
                    "AND a.病人id=[1] " & _
                    "AND a.主页id=[2] " & _
                    "AND c.记录类型=4 " & _
                    "AND a.发生时间<[3] And c.终止版本 Is Null "
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
        mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(strFrom), Val(mrsParam("婴儿")))
    If rsTmp.BOF = False Then lng次数 = rsTmp("次数").Value
    
    For intCol = 1 To mshUpTab.Cols - 1
    
        If DateDiff("d", CDate(strFrom), CDate(Split(picScale.Tag, ";")(1))) + 1 >= intCol Then

            If mshUpTab.ColData(intCol) = 1 Or mshUpTab.ColData(intCol) = 2 And lng次数 < 12 Then
                
                lng次数 = lng次数 + 1
                
                strTmp = Switch(lng次数 = 1, "Ⅰ", _
                                lng次数 = 2, "Ⅱ", _
                                lng次数 = 3, "Ⅲ", _
                                lng次数 = 4, "Ⅳ", _
                                lng次数 = 5, "Ⅴ", _
                                lng次数 = 6, "Ⅵ", _
                                lng次数 = 7, "Ⅶ", _
                                lng次数 = 8, "Ⅷ", _
                                lng次数 = 9, "Ⅸ", _
                                lng次数 = 10, "Ⅹ", _
                                lng次数 = 11, "Ⅺ", _
                                lng次数 = 12, "Ⅻ")
                
                If mblnStopFlag Then
                    
                    If strTmp = "Ⅰ" Then
                        mshUpTab.TextMatrix(2, intCol) = "0"
                    Else
                        mshUpTab.TextMatrix(2, intCol) = strTmp & "- 0"
                    End If
                Else
                    
                    If mshUpTab.TextMatrix(2, intCol) <> "" Then
                        mshUpTab.TextMatrix(2, intCol) = mshUpTab.TextMatrix(2, intCol) & "/" & strTmp
                    Else
                        mshUpTab.TextMatrix(2, intCol) = strTmp
                    End If
                End If
                
                For intLoop = intCol + 1 To mshUpTab.Cols - 1
                    strTmp = intLoop - intCol
                    
                    If Val(strTmp) <= mintOpDays Then
                        
                        If mblnStopFlag Then
                            mshUpTab.TextMatrix(2, intLoop) = strTmp
                        Else
                            If mshUpTab.TextMatrix(2, intLoop) <> "" Then
                                mshUpTab.TextMatrix(2, intLoop) = mshUpTab.TextMatrix(2, intLoop) & "/" & strTmp
                            Else
                                mshUpTab.TextMatrix(2, intLoop) = strTmp
                            End If
                        End If
                    End If
                    
                Next
                            
            End If
        End If
    Next
    
    ShowOpsDays = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadBodyData() As Boolean
    '******************************************************************************************************************
    '功能： 根据当前病人体温表，提取病人的体温表数据，填写到相应的单元中
    '参数： lng病区id : 科室
    '       strFrom : 开始时间
    '       strTo : 终止时间
    '返回：
    '注意： 病人病历所见单的行为0时“所见内容”用来保存着时间数据并且“所见项ID”为空，对于说明、手术、记录人等系统项
    '       目的值将与其它项目一样保存在“病人病历所见单”的所内内容中可以从“所见项ID”找到是那个项目
    '******************************************************************************************************************
    Dim lngValue As Long
    Dim lngValue2 As Long
    Dim intSvrCol As Integer
    Dim lng单位个数 As Long
    Dim dbl最大 As Double
    Dim dbl最小 As Double
    Dim dbl单位值 As Double
    Dim lng最高行 As Long
    Dim rsTmp As New ADODB.Recordset
    Dim aryValue() As String
    Dim aryPart() As String
    Dim intMinCol As Integer
    Dim intMaxCol As Integer
    Dim blnOperate As Boolean
    Dim dtOperate As Date
    Dim strEnd As String '是否手术、手术日期
    Dim i As Long
    Dim lng病区id As Long, strFrom As String, strTo As String
    Dim intColTmp As Integer
    Dim lngColor As Long
    Dim strTime As String
    Dim strStart1 As String
    Dim strEnd1 As String
    Dim strTmp As String
    Dim blnShow As Boolean          '是否显示入出院等信息
    
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHead
    
    '变量初始化
    '------------------------------------------------------------------------------------------------------------------
    lng病区id = Val(mrsParam("病区id"))
    strFrom = CStr(mrsParam("开始时间"))
    
    If zlDatabase.GetPara("体温单显示诊断", glngSys, 1255, 1) = 0 Then
        lbl(7).Visible = False
        txtCard(7).Visible = False
    Else
        lbl(7).Visible = True
        txtCard(7).Visible = True
    End If
    
    If CStr(mrsParam("结束时间")) > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
        strTo = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Else
        strTo = CStr(mrsParam("结束时间"))
    End If
        
    txtCard(3).Text = ""
    
    '如果是新生儿，则重新计算时间，即婴儿体温单的开始时间
    If Val(mrsParam("婴儿").Value) > 0 Then
        mstrSQL = " Select  b.出生时间 From 病人新生儿记录 B Where 病人id=[1] And 主页id=[2] And 序号=[3] "
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value))
        If rsTmp.BOF = False Then
            mstrEnterDate = Format(zlCommFun.NVL(rsTmp("出生时间").Value), "yyyy-MM-dd HH:mm:ss")
            txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("出生时间").Value), "yyyy-MM-dd")
            strFrom = mstrEnterDate
        End If
    End If
    
    '按照4小时精确度，整理开始时间和终止时间
    
    intCol = GetCurveColumn(CDate(strFrom), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
    strFrom = Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(0)
    
    If Int(CDate(strFrom)) < Int(CDate(mrsParam("开始时间"))) Then
        strFrom = Format(Int(CDate(mrsParam("开始时间"))), "yyyy-MM-dd HH:mm:ss")
    End If

    intCol = GetCurveColumn(CDate(strTo), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
    strTo = Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(1)
    
    '将允许的开始时间和终止时间填写到标尺tag中
    picScale.Tag = strFrom & ";" & strTo
    
    
    '读取病人基本信息
    '------------------------------------------------------------------------------------------------------------------
    '填写病人病区、床号等，由于病人可能在科内换床等，本处填写为病人的在指定时间的最后床号
'    lblTime.Caption = "日期:" & Format(strFrom, "MM-DD") & "～" & Format(strTo, "MM-DD")
    
    '入院时间(以入科时间为准)
    mstrSQL = "select 开始时间 from 病人变动记录 where 病人id=[1] And 主页id=[2] and 开始原因=2 order by 开始时间"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rsTmp.BOF = False Then
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("开始时间").Value), "yyyy-MM-dd")
    End If
    
    '读取病人基本信息
    mstrSQL = " Select  b.姓名,A.住院号,b.入院时间,b.性别,b.年龄 From 病人信息 B,病案主页 A Where A.病人ID=B.病人ID And A.病人id=[1] And A.主页ID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")))
    If rsTmp.BOF = False Then
        txtCard(0).Text = zlCommFun.NVL(rsTmp("姓名").Value)
        txtCard(0).Tag = zlCommFun.NVL(rsTmp("姓名").Value)
        txtCard(1).Text = zlCommFun.NVL(rsTmp("住院号").Value)
        txtCard(5).Text = zlCommFun.NVL(rsTmp("性别").Value)
        txtCard(6).Text = zlCommFun.NVL(rsTmp("年龄").Value)
        If txtCard(3).Text = "" Then txtCard(3).Text = Format(zlCommFun.NVL(rsTmp("入院时间").Value), "yyyy-MM-dd")
    End If
    
    Call zlMenuClick("显示病人姓名")

    '读取病人科室、床号等信息
    
    txtCard(2).Text = ""
    txtCard(4).Text = ""
    
    mstrSQL = " Select  c.名称 As 科室,b.名称 As 病区,a.床号,a.开始原因 " & _
                "From 病人变动记录 a,部门表 b,部门表 c " & _
                "Where a.病人id=[1] And a.主页id=[2] And a.科室id Is Not Null And a.病区id=b.id and a.科室id=c.id And a.开始时间-4/24<=[3] And Nvl(a.终止时间,Sysdate)>=[4] Order By a.开始时间"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(strTo), CDate(strFrom))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            
            If zlCommFun.NVL(rsTmp("科室").Value) <> strTmp And zlCommFun.NVL(rsTmp("科室").Value) <> "" Then
            
                strTmp = zlCommFun.NVL(rsTmp("科室").Value)
                
                If txtCard(2).Text = "" Then
                    txtCard(2).Text = strTmp
                Else
                    txtCard(2).Text = txtCard(2).Text & "->" & strTmp
                End If
                
            End If

            If zlCommFun.NVL(rsTmp("床号").Value) <> strTime And zlCommFun.NVL(rsTmp("床号").Value) <> "" Then
            
                strTime = zlCommFun.NVL(rsTmp("床号").Value)
                
                If txtCard(4).Text = "" Then
                    txtCard(4).Text = strTime
                Else
                    txtCard(4).Text = txtCard(4).Text & "->" & strTime
                End If
                
            End If
                        
            rsTmp.MoveNext
        Loop
        
        If Left(txtCard(2).Text, 2) = "->" Then txtCard(2).Text = Mid(txtCard(2).Text, 3)
        If Left(txtCard(4).Text, 2) = "->" Then txtCard(4).Text = Mid(txtCard(4).Text, 3)
    End If
    
    mshUpTab.Redraw = False
    mshScale.Redraw = False
        
    '填写日期和住院天数等信息
    '------------------------------------------------------------------------------------------------------------------
    With mshUpTab
        
        intSvrCol = .Col
        
        lngValue = 0
        mstrSQL = "Select zl_CalcInDays([1],[2],[3],[4]) As 开始天数 From Dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("婴儿").Value), (Int(CDate(strFrom))))
        If rsTmp.BOF = False Then
            lngValue = rsTmp("开始天数").Value
        End If
        
        For intCol = 1 To .Cols - 1

            .ColData(intCol) = 0
            .ColAlignment(intCol) = 4
            
            strTmp = Format(CDate(strFrom) + intCol - 1, "yyyy-MM-dd")
                        
            If Right(strTmp, 5) = "01-01" Then
                '一年的第一天
                .TextMatrix(0, intCol) = strTmp
            ElseIf strTmp = Format(mstrEnterDate, "yyyy-MM-dd") Then
                '入院第一天，写上年份
                .TextMatrix(0, intCol) = strTmp
            ElseIf intCol = 1 Then
                .TextMatrix(0, intCol) = strTmp
            ElseIf Right(strTmp, 2) = "01" Then
                .TextMatrix(0, intCol) = Right(strTmp, 5)
            Else
                .TextMatrix(0, intCol) = Right(strTmp, 2)
            End If

            .TextMatrix(1, intCol) = lngValue + (intCol - 1)
            .TextMatrix(2, intCol) = ""
        Next

    End With

    
    Dim intDays As Integer
    
    For intCol = 1 To mshUpTab.Cols - 1
         mstrOpsSvr(intCol) = ""
         mstrOpsDays(intCol) = ""
    Next
        
    '1.提取病人体温图形数据，换算为坐标值填写到表中，为图形显示作准备
    '------------------------------------------------------------------------------------------------------------------
    With mshScale
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(GraphDataRow.更改标志, intCol) = String(.FixedCols, ";")    '更改标志
            .TextMatrix(GraphDataRow.曲线数据, intCol) = String(.FixedCols, ";")    '曲线数据
            .TextMatrix(GraphDataRow.上标说明, intCol) = ""                         '说明(上标)
            
            .Cell(flexcpData, GraphDataRow.手术标志, intCol, GraphDataRow.手术标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.部位标志, intCol, GraphDataRow.部位标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.入院标志, intCol, GraphDataRow.入院标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.转科标志, intCol, GraphDataRow.转科标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.换床标志, intCol, GraphDataRow.换床标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.出院标志, intCol, GraphDataRow.出院标志, intCol) = ""
            .Cell(flexcpData, GraphDataRow.入科标志, intCol, GraphDataRow.入科标志, intCol) = ""
            
            .TextMatrix(GraphDataRow.手术标志, intCol) = ""                         '手术
            .TextMatrix(GraphDataRow.入院标志, intCol) = ""                         '入院
            .TextMatrix(GraphDataRow.转科标志, intCol) = ""                         '转科
            .TextMatrix(GraphDataRow.换床标志, intCol) = ""                         '换床
            .TextMatrix(GraphDataRow.出院标志, intCol) = ""                         '出院
            .TextMatrix(GraphDataRow.入科标志, intCol) = ""                         '入科
            .TextMatrix(GraphDataRow.复试标志, intCol) = ""                         '体温复试合格
            .TextMatrix(GraphDataRow.下标说明, intCol) = ""                         '说明(下标)
            .TextMatrix(GraphDataRow.断开标志, intCol) = ""                         '无曲线数据，断开
            .TextMatrix(GraphDataRow.出生标志, intCol) = ""                         '出生
            .TextMatrix(GraphDataRow.曲线时间, intCol) = String(.FixedCols, ";")    '曲线时间
            .TextMatrix(GraphDataRow.未记说明, intCol) = String(.FixedCols, ";")    '未记说明
            .TextMatrix(GraphDataRow.部位标志, intCol) = String(.FixedCols, ";")    '体温部位
            
        Next
        
    End With
    
    mintOpDays = Val(zlDatabase.GetPara("手术后标注天数", glngSys, 1255, "10"))
    mblnStopFlag = (Val(zlDatabase.GetPara("再次手术停止前次标注", glngSys, 1255, "0")) = 1)
    
    '显示当前区域段前的手术日标记
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "SELECT a.发生时间 As 时间,c.项目名称 " & _
                "FROM 病人护理记录 a,病人护理内容 c " & _
                "Where a.ID = c.记录ID " & _
                    "AND a.病人来源=2 " & _
                    "AND Nvl(a.婴儿,0)=[5] " & _
                    "AND a.病人id=[1] " & _
                    "AND a.主页id=[2] " & _
                    "AND c.记录类型=4 And c.终止版本 Is Null " & _
                    "AND a.发生时间 Between [3] And [4] Order By a.发生时间 "

    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(strFrom) - 14, CDate(strTo), Val(mrsParam("婴儿")))
    If rsTmp.BOF = False Then
        Do While Not rsTmp.EOF
            dtOperate = Int(rsTmp.Fields(0).Value)
            For intCol = 1 To mshUpTab.Cols - 1
                
                If DateDiff("d", CDate(strFrom), CDate(Split(picScale.Tag, ";")(1))) + 1 >= intCol Then
                    intDays = Val(Int(CDate(strFrom)) + intCol - 1 - dtOperate)
    
                    Select Case intDays
                    Case 0
                    
                        mshUpTab.ColData(intCol) = 1
                        mstrOpsDays(intCol) = rsTmp.Fields(0).Value
                        
                    Case 1 To mintOpDays
                    
                        If intDays >= intCol Then
                        
                            If mshUpTab.TextMatrix(2, intCol) <> "" And Not mblnStopFlag Then
                                mshUpTab.TextMatrix(2, intCol) = mshUpTab.TextMatrix(2, intCol) & "/" & intDays
                            Else
                                mshUpTab.TextMatrix(2, intCol) = intDays
                            End If
                            
                        End If
                    End Select
                    
                    mstrOpsSvr(intCol) = mshUpTab.TextMatrix(2, intCol)
                    
                    intCount = GetCurveColumn(rsTmp("时间").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    
                    Select Case rsTmp("项目名称").Value
                    Case "分娩"
                        If intCount > 0 And intCount < mshScale.Cols And mBodyFlag.分娩 > 0 Then
                            If mBodyFlag.分娩 = 2 Then
                                mshScale.TextMatrix(3, intCount) = rsTmp("项目名称").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Else
                                mshScale.TextMatrix(3, intCount) = rsTmp("项目名称").Value
                            End If
                            
                             mshScale.Cell(flexcpData, 3, intCount, 3, intCount) = Format(rsTmp("时间").Value, "HH:mm:ss")
    
                        End If
                    Case Else
                        If intCount > 0 And intCount < mshScale.Cols And mBodyFlag.手术 > 0 Then
                            If mBodyFlag.手术 = 2 Then
                                mshScale.TextMatrix(3, intCount) = rsTmp("项目名称").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Else
                                mshScale.TextMatrix(3, intCount) = rsTmp("项目名称").Value
                            End If
                            
                             mshScale.Cell(flexcpData, 3, intCount, 3, intCount) = Format(rsTmp("时间").Value, "HH:mm:ss")
    
                        End If
                    End Select
                End If
            Next
            rsTmp.MoveNext
        Loop
    End If

    '显示当前区域段内的手术日标记
    '------------------------------------------------------------------------------------------------------------------
    Call ShowOpsDays
    
    '2.读取入出转等标志数据
    '------------------------------------------------------------------------------------------------------------------
    Dim bytShow As Byte
    
    Set rsTmp = GetDataFromHis(Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("婴儿")), CDate(strFrom), CDate(strTo), 2)
    If Not (rsTmp Is Nothing) Then
        If rsTmp.BOF = False Then
            Do While Not rsTmp.EOF

                intCol = GetCurveColumn(rsTmp("时间").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                
                If zlCommFun.NVL(rsTmp("内容")) <> "" Then
                    
                    bytShow = 0
                    
                    Select Case Val(rsTmp("行号").Value)
                    Case 5
                        bytShow = mBodyFlag.入院
                    Case 6
                        bytShow = mBodyFlag.转出
                    Case 7
                        bytShow = mBodyFlag.换床
                    Case 8
                        bytShow = mBodyFlag.出院
                    Case 9
                        bytShow = mBodyFlag.入科
                    End Select
                    
                    If intCol >= mshScale.FixedCols And intCol < mshScale.Cols And bytShow > 0 Then
                        blnShow = True
                        If Val(rsTmp("行号").Value) = 8 And Val(mrsParam("婴儿")) > 0 Then
                            blnShow = mbln婴儿体温单显示出院
                        End If
                        
                        If blnShow Then
                            mshScale.Cell(flexcpData, Val(rsTmp("行号").Value), intCol, Val(rsTmp("行号").Value), intCol) = Format(rsTmp("时间").Value, "HH:mm:ss")
                            Select Case bytShow
                            Case 1
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value
                            Case 2
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Case 3
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & rsTmp("科室").Value
                            Case 4
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & rsTmp("科室").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            End Select
                        End If
                    End If
                End If
                                            
                rsTmp.MoveNext
            Loop
        End If
    End If
    
    If Val(mrsParam("婴儿")) > 0 Then
        Set rsTmp = GetDataFromHis(Val(mrsParam("病人id")), Val(mrsParam("主页id")), Val(mrsParam("婴儿")), CDate(strFrom), CDate(strTo), 3)
        
        If Not (rsTmp Is Nothing) Then
            If rsTmp.BOF = False Then
                Do While Not rsTmp.EOF
    
                    intCol = GetCurveColumn(rsTmp("时间").Value, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    
                    If zlCommFun.NVL(rsTmp("内容")) <> "" Then
                                               
                        If intCol >= mshScale.FixedCols And intCol < mshScale.Cols And mBodyFlag.出生 > 0 Then
                            
                            mshScale.Cell(flexcpData, Val(rsTmp("行号").Value), intCol, Val(rsTmp("行号").Value), intCol) = Format(rsTmp("时间").Value, "HH:mm:ss")
                            Select Case mBodyFlag.出生
                            Case 1
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value
                            Case 2
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            Case 3
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & rsTmp("科室").Value
                            Case 4
                                mshScale.TextMatrix(Val(rsTmp("行号").Value), intCol) = rsTmp("内容").Value & rsTmp("科室").Value & "--" & ConvertTimeToChinese(Format(rsTmp("时间").Value, "HH:mm"))
                            End Select
                            
                        End If
                    End If
                                                
                    rsTmp.MoveNext
                Loop
            End If
        End If
    End If
    
    '3.拒查等标注部份
    '------------------------------------------------------------------------------------------------------------------
    mstrSQL = "SELECT c.记录类型,a.发生时间 As 时间,c.记录内容 As 说明,c.记录标记 " & _
                "FROM 病人护理记录 a,病人护理内容 C " & _
                "Where a.ID = c.记录ID " & _
                    "AND Nvl(a.婴儿,0)=[5] " & _
                    "AND a.病人id=[1] " & _
                    "AND a.主页id=[2] " & _
                    "AND C.记录类型 In (2,6) " & _
                    "AND a.病人来源=2 And c.终止版本 Is Null " & _
                    "AND a.发生时间 BETWEEN [3] And [4] " & _
                "Order By a.发生时间"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
        mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
    End If
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, _
                                        Val(mrsParam("病人id")), _
                                        Val(mrsParam("主页id")), _
                                        CDate(Format(strFrom, "YYYY-MM-DD")), _
                                        CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), Val(mrsParam("婴儿")))
    With rsTmp
        Do While Not .EOF

            intCol = GetCurveColumn(!时间, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
            
            If intCol >= mshScale.FixedCols And intCol < mshScale.Cols Then
                If zlCommFun.NVL(rsTmp("记录类型").Value, 0) = 2 Then
                    mshScale.TextMatrix(GraphDataRow.上标说明, intCol) = zlCommFun.NVL(!说明)
                Else
                    mshScale.TextMatrix(GraphDataRow.下标说明, intCol) = zlCommFun.NVL(!说明)
                End If
                
                aryValue() = Split(mshScale.TextMatrix(GraphDataRow.更改标志, intCol), ";")
                aryValue(0) = 1
                mshScale.TextMatrix(GraphDataRow.更改标志, intCol) = Join(aryValue, ";")
                mshScale.TextMatrix(GraphDataRow.断开标志, intCol) = IIf(IsNull(!记录标记), "0", !记录标记)
                
            End If
            
            .MoveNext
        Loop
    End With
    
    '4.曲线数据部份
    '------------------------------------------------------------------------------------------------------------------
    Dim int列号 As Integer
    Dim aryItemName() As String
    Dim aryItemOrder() As Integer
    
    Dim strItemOrder As String
    Dim strItemName As String
    
    ReDim aryItemName(0 To mshScale.FixedCols - 1)
    ReDim aryItemOrder(0 To mshScale.FixedCols - 1)
    
    For intCol = 0 To mshScale.FixedCols - 1
        If InStr(mshScale.TextMatrix(0, intCol), "(") > 0 Then
            strTmp = Trim(Left(mshScale.TextMatrix(0, intCol), InStr(mshScale.TextMatrix(0, intCol), "(") - 1))
        Else
            strTmp = Trim(mshScale.TextMatrix(0, intCol))
        End If
        
        aryItemName(intCol) = strTmp
        aryItemOrder(intCol) = intCol + 1

    Next
    
    '45987,刘鹏飞,2012-09-10,湖南需求
    '1.呼吸显示为曲线；2、呼吸统一用黑色实行点（●）表示，不在体温单上显示呼吸数字
    '3. 使用呼吸机的患者，呼吸以黑R表示，在相应时间内呼吸30次横线下顶格用黑笔划R，相邻的R之间以及R和自主呼吸之间不连线
    mstrSQL = "SELECT a.发生时间 As 时间,Decode(D.项目序号," & mItemNo.呼吸 & ",Decode(C.体温部位,'呼吸机','29',C.记录内容) ,c.记录内容) As 数值,c.体温部位,c.复试合格,D.记录名,E.保留项目,D.项目序号,C.记录标记,C.未记说明 " & _
                "FROM 病人护理记录 A,病人护理内容 C,体温记录项目 D,护理记录项目 E " & _
                "Where a.ID = c.记录ID " & _
                    "AND a.病人来源=2 " & _
                    "AND Nvl(a.婴儿,0)=[5] " & _
                    "AND a.病人id=[1] " & _
                    "AND a.主页id=[2] " & _
                    "AND D.项目序号=c.项目序号 " & _
                    "AND c.记录类型=1 " & _
                    "AND E.项目序号=D.项目序号 " & _
                    "AND E.护理等级>=[6]  " & _
                    "AND a.发生时间 BETWEEN [3] And [4] And c.终止版本 Is Null " & _
                    "AND D.记录法=1  " & _
                "Order By a.发生时间,c.记录标记"
    If mblnMoved Then
        mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
        mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
    End If
    
    
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, Val(mrsParam("病人id")), Val(mrsParam("主页id")), CDate(Format(strFrom, "YYYY-MM-DD")), CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), Val(mrsParam("婴儿")), Val(mrsParam("护理等级").Value))
    With rsTmp
        
        Dim dtTmp As Date
        Dim blnAllow As Boolean
        Dim rsOffset As ADODB.Recordset
        
        Call InitOffset(rsOffset)
                
        Do While Not .EOF

            intCol = GetCurveColumn(!时间, CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
            
            If (intCol - mshScale.FixedCols) < 42 Then

                strTmp = !记录名
                Select Case mint心率应用
                Case 1      '单独应用
                    If zlCommFun.NVL(!记录标记, 0) = 1 And strTmp = "脉搏" Then
                        strTmp = "心率"
                    End If
                Case 2      '共享使用
                    If strTmp = "心率" Then strTmp = "脉搏"
                End Select
                    
                '根据列名计算出列号
                For int列号 = 0 To mshScale.FixedCols - 1
                    If strTmp = aryItemName(int列号) Then
                        int列号 = aryItemOrder(int列号)
                        Exit For
                    End If
                Next
                                
                '如果同一列中有多个值，则取最靠中点的值作为本列的值
                dtTmp = CDate(Split(GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin), ",")(0))
                dtTmp = DateAdd("h", 2, dtTmp)
                blnAllow = IsCenterValue(rsOffset, int列号, intCol, !时间, dtTmp)
                
                If blnAllow Then
                    
                    '获取项目定义:最大值；最小值；单位值；最高行
                    aryValue = Split(picLine(int列号 - 1).Tag, ";")
                    dbl最大 = Val(Split(picLine(int列号 - 1).Tag, ";")(0))
                    dbl最小 = Val(Split(picLine(int列号 - 1).Tag, ";")(1))
                    dbl单位值 = Val(Split(picLine(int列号 - 1).Tag, ";")(2))
                    lng最高行 = Val(Split(picLine(int列号 - 1).Tag, ";")(3))
                    
                    '求出总共有多少个单位
                    lng单位个数 = (dbl最大 - dbl最小) / dbl单位值
                    '如果单位个数加上最高行大于20，就取20-最高行，否则就取 单位个数+最高行
                    lng单位个数 = IIf(lng单位个数 + lng最高行 > (MAXROWS - 1), (MAXROWS - 1) - lng最高行, lng单位个数 + lng最高行)
                    '坐标值=((最大值-当前值)/单位值+最高行-1)*行高度
                    
                    If zlCommFun.NVL(!数值) <> "" Then
                        lngValue2 = 0
                        If InStr(!数值, ",") > 0 Then
                            lngValue = ConvertToY(int列号 - 1, Val(Mid(!数值, 1, InStr(!数值, ",") - 1)))
                            lngValue2 = ConvertToY(int列号 - 1, Val(Mid(!数值, InStr(!数值, ",") + 1)))
                        Else
                            lngValue = ConvertToY(int列号 - 1, Val(!数值))
                        End If
                    
                        aryValue() = Split(mshScale.TextMatrix(GraphDataRow.曲线数据, intCol), ";")
    
                        '记录同一时间两个点的曲线
                        
                        strStart1 = Format(Int(CDate(strFrom)) + ((intCol - mshScale.FixedCols) * 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        strEnd1 = Format(Int(CDate(strFrom)) + ((intCol - mshScale.FixedCols) * 4 + 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        
                        If Val(aryValue(int列号)) > 0 And zlCommFun.NVL(!记录标记, 0) = 1 Then
                            If strTmp = "心率" And mint心率应用 = 1 Then
                                aryValue(int列号) = lngValue
                            Else
                                aryValue(int列号) = aryValue(int列号) & "," & lngValue
                            End If
                                                    
                        Else
                            aryValue(int列号) = lngValue
                            
                            If !项目序号 = mItemNo.体温 Then
                                '保存体温部位
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.部位标志, intCol), ";")
                                aryPart(int列号) = zlCommFun.NVL(!体温部位, "腋温")
                                mshScale.TextMatrix(GraphDataRow.部位标志, intCol) = Join(aryPart, ";")
                                mshScale.TextMatrix(GraphDataRow.复试标志, intCol) = zlCommFun.NVL(!复试合格, "0")
                            ElseIf !项目序号 = mItemNo.呼吸 Then
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.部位标志, intCol), ";")
                                aryPart(int列号) = zlCommFun.NVL(!体温部位, "自主呼吸")
                                mshScale.TextMatrix(GraphDataRow.部位标志, intCol) = Join(aryPart, ";")
                            ElseIf !项目序号 = mItemNo.脉搏 Then
                                aryPart = Split(mshScale.TextMatrix(GraphDataRow.部位标志, intCol), ";")
                                aryPart(int列号) = zlCommFun.NVL(!体温部位, "")
                                mshScale.TextMatrix(GraphDataRow.部位标志, intCol) = Join(aryPart, ";")
                            End If
                    
                        End If
                        mshScale.TextMatrix(GraphDataRow.曲线数据, intCol) = Join(aryValue, ";")
                    End If
                    
                    '填写更改标志
                    aryValue() = Split(mshScale.TextMatrix(GraphDataRow.更改标志, intCol), ";")
                    aryValue(int列号) = 1
                    mshScale.TextMatrix(GraphDataRow.更改标志, intCol) = Join(aryValue, ";")
                    
                    '填写曲线时间
                    aryValue() = Split(mshScale.TextMatrix(GraphDataRow.曲线时间, intCol), ";")
                    aryValue(int列号) = Format(!时间, "yyyy-MM-dd HH:mm:ss")
                    mshScale.TextMatrix(GraphDataRow.曲线时间, intCol) = Join(aryValue, ";")
                    
                    '填写未记说明
                    If zlCommFun.NVL(!数值) = "不升" Then
                        aryValue() = Split(mshScale.TextMatrix(GraphDataRow.未记说明, intCol), ";")
                        aryValue(int列号) = "不升"
                        mshScale.TextMatrix(GraphDataRow.未记说明, intCol) = Join(aryValue, ";")
                    Else
                        If zlCommFun.NVL(!未记说明) <> "" Then
                            aryValue() = Split(mshScale.TextMatrix(GraphDataRow.未记说明, intCol), ";")
                            aryValue(int列号) = !未记说明
                            mshScale.TextMatrix(GraphDataRow.未记说明, intCol) = Join(aryValue, ";")
                        End If
                    End If
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    '提取病人体温表格直接录入数据，填写到表中
    '------------------------------------------------------------------------------------------------------------------
    Call ReadGridData(Val(mrsParam("护理等级").Value), _
                        Val(mrsParam("科室id").Value), _
                        IIf(Val(mrsParam("婴儿").Value) = 0, 1, 2), _
                        Val(mrsParam("病人id").Value), _
                        Val(mrsParam("主页id").Value), _
                        CDate(Format(strFrom, "YYYY-MM-DD")), _
                        CDate(Format(strTo, "YYYY-MM-DD") & " 23:59:59"), _
                        Val(mrsParam("婴儿").Value), _
                        mblnMoved)
    
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    
    '如果是源程序调试,输出表格内容
    On Error Resume Next
    Err = 0
    Debug.Print 1 / 0
    If Err <> 0 Then
        Call OutputDadaForDebug
    End If
    
    Screen.MousePointer = vbDefault
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckGridData(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '功能： 检查指定的表格录入项目是否有数据
    '参数： intIndex : 索引
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim aryTmp As Variant
    Dim intCol As Integer
    Dim intCount As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    CheckGridData = True
    
    With mshDownTab
        For intLoop = .FixedCols To .Cols - 1
            If Trim(.TextMatrix(intIndex, intLoop)) <> "" Then
                Exit Function
            End If
            
            aryTmp = Split(.TextMatrix(GridDataRow.修改标志, intLoop), ";")
            Select Case Val(aryTmp(intIndex - 1))
            Case OperateType.新增操作, OperateType.修改操作, OperateType.删除操作
                Exit Function
            End Select
        Next
    End With
        
    CheckGridData = False
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function DeleteActiveItem(ByVal intIndex As Integer) As Boolean
    '******************************************************************************************************************
    '功能： 清除指定的活动项目
    '参数： intIndex : 活动项目索引
    '返回：
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim aryTmp As Variant
    Dim intCol As Integer
    Dim intCount As Integer
    Dim strTmp As String
    
    On Error GoTo errHand
    
    For intLoop = UBound(mItemStru) To LBound(mItemStru) Step -1
        If mItemStru(intLoop).活动项目 And intIndex = intLoop Then

            With mshDownTab
                
                '移动相关的数据
                For intCol = .FixedCols To .Cols - 1
                                        
                    aryTmp = Split(.TextMatrix(GridDataRow.修改标志, intCol), ";")
                    
                    For intCount = intIndex To .Rows - 2
                        aryTmp(intCount) = aryTmp(intCount + 1)
                        mItemStru(intCount).项目名称 = mItemStru(intCount + 1).项目名称
                        mItemStru(intCount).数据类型 = mItemStru(intCount + 1).数据类型
                        mItemStru(intCount).数据长度 = mItemStru(intCount + 1).数据长度
                        mItemStru(intCount).小数位数 = mItemStru(intCount + 1).小数位数
                        mItemStru(intCount).最小值 = mItemStru(intCount + 1).最小值
                        mItemStru(intCount).最大值 = mItemStru(intCount + 1).最大值
                        mItemStru(intCount).记录频次 = mItemStru(intCount + 1).记录频次
                        mItemStru(intCount).活动项目 = mItemStru(intCount + 1).活动项目
                        mItemStru(intCount).项目序号 = mItemStru(intCount + 1).项目序号
                    Next
                    aryTmp(.Rows - 2) = ""
                    
                    strTmp = Join(aryTmp, ";")
                    .TextMatrix(GridDataRow.修改标志, intCol) = Left(strTmp, Len(strTmp) - 1)
                Next
                
                '删除行及相关数组
                .RemoveItem intIndex
                
                intCount = UBound(mItemStru) - 1
                ReDim Preserve mItemStru(intCount)
            End With

            DeleteActiveItem = True

            Exit For
        End If
    Next
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function ReadGridData(ByVal byt护理等级 As Byte, _
                                ByVal lng科室id As Long, _
                                ByVal byt适用病人 As Byte, _
                                ByVal lng病人id As Long, _
                                ByVal lng主页id As Long, _
                                ByVal dt开始时间 As Date, _
                                ByVal dt结束时间 As Date, ByVal byt婴儿 As Byte, Optional ByVal blnMoved As Boolean) As Boolean
    '******************************************************************************************************************
    '功能： 提取病人体温表格直接录入数据，填写到表中
    '参数： lng病区id : 科室
    '       strFrom : 开始时间
    '       strTo : 终止时间
    '返回：
    '******************************************************************************************************************
    Dim strItemOrder As String
    Dim strItemName As String
    Dim i As Long
    Dim aryValue() As String
    Dim aryTmp As Variant
    Dim intRow As Integer
    Dim lngColor As Long
    Dim strTime As String
    Dim intColTmp As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim sgl饮入量() As Single
    Dim sgl排出量() As Single
    Dim sgl饮入量2() As Single
    Dim sgl排出量2() As Single
    Dim blnChanged As Boolean
    Dim intCount As Integer
    
    On Error GoTo errHand
        
    mshDownTab.Redraw = False
    
    '活动项目处理
    '------------------------------------------------------------------------------------------------------------------
    '先清除所有活动项目
'    intCount = 0
    For i = UBound(mItemStru) To LBound(mItemStru) Step -1
        If mItemStru(i).活动项目 Then
'            intCount = intCount + 1
      
            Call DeleteActiveItem(i)
            blnChanged = True
        Else
            Exit For
        End If
    Next
'    If blnChanged And mshDownTab.Rows - intCount > 0 Then
'        '删除表格行
'        mshDownTab.Rows = mshDownTab.Rows - intCount
'
'        '删除相关的数组项
'        intCount = UBound(mItemStru) - intCount
'        ReDim Preserve mItemStru(intCount)
'    End If
    
    '自动添加活动项目（只加当前页中有数据的）
    Set rsTmp = GetGridDataItem(byt护理等级, lng科室id, byt适用病人, lng病人id, lng主页id, dt开始时间, dt结束时间, byt婴儿, blnMoved)
    If rsTmp.BOF = False Then
        blnChanged = True
        Do While Not rsTmp.EOF
            Call AppendGridItem(rsTmp)
            rsTmp.MoveNext
        Loop
    End If
    
    '重新调整界面控件位置
    If blnChanged Then Call picPane_Resize
    
    '初始修改标志及结果内容
    '------------------------------------------------------------------------------------------------------------------
    With mshDownTab
        For intCol = .FixedCols To .Cols - 1
            .TextMatrix(GridDataRow.修改标志, intCol) = String(.Rows - 2, ";")
        Next
        For intRow = 1 To .Rows - 1
            For intCol = .FixedCols To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
        Next
    End With
    
    '读取项目及序号清单
    '------------------------------------------------------------------------------------------------------------------
    strItemOrder = ""
    strItemName = ""
    For intRow = mshDownTab.FixedRows To mshDownTab.Rows - 1
        If Val(mshDownTab.RowData(intRow)) <> mItemNo.血压 Then
            i = InStr(1, mshDownTab.TextMatrix(intRow, 0), "(", vbTextCompare)
            If i > 0 Then
                strItemName = strItemName & ",'" & Left(mshDownTab.TextMatrix(intRow, 0), i - 1) & "'"
            Else
                strItemName = strItemName & ",'" & mshDownTab.TextMatrix(intRow, 0) & "'"
            End If
        End If
    Next
    
    
    '读取数据
    '------------------------------------------------------------------------------------------------------------------
    If strItemName <> "" Then
        strItemName = Mid(strItemName, 2) & ",'呼吸','收缩压','舒张压'"
                        
        mstrSQL = "SELECT a.发生时间 As 时间,C.记录内容 As 结果,E.保留项目,D.记录名,D.项目序号,D.记录频次 " & _
                    "FROM 病人护理记录 A,病人护理内容 C,体温记录项目 D,护理记录项目 E " & _
                    "Where A.ID = c.记录ID " & _
                        "AND A.病人来源=2 " & _
                        "AND Nvl(a.婴儿,0)=[6] " & _
                        "AND A.病人id=[1] " & _
                        "AND A.主页id=[2] " & _
                        "AND INSTR([5],','''||D.记录名||''',')>0 " & _
                        "AND D.项目序号=C.项目序号 " & _
                        "AND c.记录类型=1 " & _
                        "AND E.项目序号=D.项目序号 " & _
                        "AND E.护理等级>=[7]  " & _
                        "AND a.发生时间 BETWEEN [3] And [4] And c.终止版本 Is Null " & _
                        "AND D.记录法=2 " & _
                    "Order By Decode(D.记录名,'收缩压',0,1)," & strItemName & ",a.发生时间"
                    
        If blnMoved Then
            mstrSQL = Replace(mstrSQL, "病人护理记录", "H病人护理记录")
            mstrSQL = Replace(mstrSQL, "病人护理内容", "H病人护理内容")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, mstrMsgTitle, lng病人id, lng主页id, dt开始时间, dt结束时间, "," & strItemName & ",", byt婴儿, byt护理等级)
                                                    
        Dim intColFirst1 As Integer
        Dim intColFirst2 As Integer
        
        intColFirst1 = 0
        intColFirst2 = 0
        
        With rsTmp
            vsf.Cell(flexcpText, 1, 2, 1, vsf.Cols - 1) = ""
            vsf.Cell(flexcpData, 1, 2, 1, vsf.Cols - 1) = ""
            vsf.Cell(flexcpForeColor, 1, 2, 1, vsf.Cols - 1) = 200
                
            If rsTmp.RecordCount > 0 Then
                    
                ReDim sgl饮入量(0 To mshDownTab.Cols)
                ReDim sgl排出量(0 To mshDownTab.Cols)
                
                rsTmp.MoveFirst
                
                For i = 0 To rsTmp.RecordCount - 1
                                        
                    Select Case !项目序号
                    Case mItemNo.呼吸
                        intCol = GetCurveColumn(!时间, dt开始时间, mlngHourBegin) + vsf.FixedCols - 1
                        
                        If intCol < vsf.Cols Then
                            vsf.TextMatrix(1, intCol) = zlCommFun.NVL(!结果, "")
                        End If
                        
                    Case mItemNo.舒张压
                        
                        intCol = Int((!时间 - Int(dt开始时间)) * 24) \ 12 + mshDownTab.FixedCols
                        strTime = Format(Int(dt开始时间) + (intCol - mshDownTab.FixedCols) \ 2, "YYYY-MM-DD")
                        intRow = mItemSerial.血压
                        
                        
                        '入院当天的第一次以接近入院时间为准，也就是最早的为准，不是取最后一次
                        If Format(!时间, "yyyy-MM-dd") = Format(txtCard(3).Text, "yyyy-MM-dd") Then
                            If intColFirst2 = 0 Then
                                If intCol < mshDownTab.Cols Then
                                    If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!结果, "") <> "" Then
                                        
                                        If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                            mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!结果, "")
                                        Else
                                            aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                            aryTmp(1) = zlCommFun.NVL(!结果, "")
                                            mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                        End If
                                        
                                    End If
            
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                                
                                intColFirst2 = intCol
                            ElseIf intColFirst2 <> intCol Then
                                If intCol < mshDownTab.Cols Then
                                    If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!结果, "") <> "" Then
                                        
                                        If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                            mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!结果, "")
                                        Else
                                            aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                            aryTmp(1) = zlCommFun.NVL(!结果, "")
                                            mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                        End If
                                        
                                    End If
            
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                            End If
                            
                        Else
                            intColFirst2 = intCol
                            
                            If intCol < mshDownTab.Cols Then
                                If mshDownTab.TextMatrix(intRow, intCol) <> "" Or zlCommFun.NVL(!结果, "") <> "" Then
                                    
                                    If InStr(mshDownTab.TextMatrix(intRow, intCol), "/") = 0 Then
                                        mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/" & zlCommFun.NVL(!结果, "")
                                    Else
                                        aryTmp = Split(mshDownTab.TextMatrix(intRow, intCol), "/")
                                        aryTmp(1) = zlCommFun.NVL(!结果, "")
                                        mshDownTab.TextMatrix(intRow, intCol) = aryTmp(0) & "/" & aryTmp(1)
                                    End If
                                    
                                End If
        
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                            End If
                        End If

                    Case Else

                        For intRow = 1 To mshDownTab.Rows - 1
                            If Val(mshDownTab.RowData(intRow)) = !项目序号 Then
                                Exit For
                            End If
                        Next
                        
                        intCol = Int((!时间 - Int(dt开始时间)) * 24) \ 12 + mshDownTab.FixedCols
                        strTime = Format(Int(dt开始时间) + (intCol - mshDownTab.FixedCols) \ 2, "YYYY-MM-DD")
                                                                                   
                        If intCol < mshDownTab.Cols Then
                            Select Case !项目序号
                            Case 7          '入液量
                                
                                If !记录频次 = 1 Then
                                    intColTmp = IIf(intCol Mod 2 = 0, intCol + 1, intCol)
                                Else
                                    intColTmp = intCol
                                End If
                                
                                sgl饮入量(intColTmp) = sgl饮入量(intColTmp) + Val(zlCommFun.NVL(!结果, ""))
                                mshDownTab.TextMatrix(intRow, intColTmp) = sgl饮入量(intColTmp)
                                
                                mshDownTab.Cell(flexcpData, intRow, intColTmp) = 0
                            Case 9          '出液量
                            
                                If !记录频次 = 1 Then
                                    intColTmp = IIf(intCol Mod 2 = 0, intCol + 1, intCol)
                                Else
                                    intColTmp = intCol
                                End If
                                
                                sgl排出量(intColTmp) = sgl排出量(intColTmp) + Val(zlCommFun.NVL(!结果))
                                
                                If Right(zlCommFun.NVL(!结果), 2) = "/C" Then
                                    mshDownTab.TextMatrix(intRow, intColTmp) = sgl排出量(intColTmp) & "/C"
                                ElseIf Right(zlCommFun.NVL(!结果), 1) = "C" Then
                                    mshDownTab.TextMatrix(intRow, intColTmp) = "C"
                                Else
                                    mshDownTab.TextMatrix(intRow, intColTmp) = sgl排出量(intColTmp)
                                End If
                                
                                mshDownTab.Cell(flexcpData, intRow, intColTmp) = 0
                                
                                If Right(mshDownTab.TextMatrix(intRow, intColTmp), 1) = "C" Then
                                    mshDownTab.Cell(flexcpData, intRow, intColTmp) = 4
                                End If
                                
                                If Right(mshDownTab.TextMatrix(intRow, intColTmp), 2) = "/C" Then
                                    mshDownTab.Cell(flexcpData, intRow, intColTmp) = 5
                                End If
                            Case 10         '大便次数
                            
                                mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!结果, "")
                                
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                If Right(mshDownTab.TextMatrix(intRow, intCol), 2) = "/E" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 3
                                ElseIf Right(mshDownTab.TextMatrix(intRow, intCol), 1) = "E" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 2
                                ElseIf Right(mshDownTab.TextMatrix(intRow, intCol), 1) = "*" Then
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 1
                                End If
                            
                            Case mItemNo.血压
                                                            
                                '入院当天的第一次以接近入院时间为准，也就是最早的为准，不是取最后一次
                                If Format(!时间, "yyyy-MM-dd") = Format(txtCard(3).Text, "yyyy-MM-dd") Then
                                    If intColFirst1 = 0 Then
                                        mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!结果, "")
                                        If zlCommFun.NVL(!结果, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                        mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                        
                                        intColFirst1 = intCol
                                        
                                    ElseIf intColFirst1 <> intCol Then
                                    
                                        mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!结果, "")
                                        If zlCommFun.NVL(!结果, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                        mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                        
                                    End If
                                    
                                Else
                                    intColFirst1 = intCol
                                    mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!结果, "")
                                    If zlCommFun.NVL(!结果, "") <> "" Then mshDownTab.TextMatrix(intRow, intCol) = mshDownTab.TextMatrix(intRow, intCol) & "/"
                                    mshDownTab.Cell(flexcpData, intRow, intCol) = 0
                                End If
                                
                            Case Else
                                
                                lngColor = GridTextColor(!记录名, zlCommFun.NVL(!结果, ""))

                                mshDownTab.Cell(flexcpForeColor, intRow, intCol, intRow, intCol) = lngColor
                                mshDownTab.TextMatrix(intRow, intCol) = zlCommFun.NVL(!结果, "")
                                mshDownTab.Cell(flexcpData, intRow, intCol) = 0

                            End Select
    
                            If InStr(mshDownTab.TextMatrix(0, intCol), ";") = 0 Then
                                mshDownTab.TextMatrix(0, intCol) = "1"
                            Else
                                aryValue() = Split(mshDownTab.TextMatrix(0, intCol), ";")
                                aryValue(intRow - 1) = 1
                                mshDownTab.TextMatrix(0, intCol) = Join(aryValue, ";")
                            End If
                        End If
                    End Select
                    
                    .MoveNext
                Next
            End If
        End With
    End If
    
    mshDownTab.Redraw = True
    Call mshDownTab_RowColChange
    
    ReadGridData = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    mshDownTab.Redraw = True
    Call SaveErrLog
    
End Function

Private Function DrawScale() As Boolean
    '******************************************************************************************************************
    '功能： 在picture上标尺，用于进入时使用
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim i As Long
    Dim strTmp As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim X0 As Long
    Dim Y0 As Long
    Dim X1 As Long
    Dim Y1 As Long
    Dim lngColor As Long
    
    picScale.Cls
    
    
    '获得最小最大时间范围
    
    If picScale.Tag <> "" Then
        Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
        lblCur.Left = intMinCol * HOUR_STEP_Twips
    End If
    
    X0 = 0
    'Y0 = mshScale.ROWHEIGHT(0) / 2 '原来  mshScale.ROWHEIGHT(0)=600 现在为300
    Y0 = mshScale.ROWHEIGHT(0)
    X1 = X0 + 15000
    'DrawLine picScale, X0, Y0, X1, Y0, &H8000000A
    For intCol = 1 To mshUpTab.Cols - 1
        X0 = intCol * HOUR_STEP_Twips * 6 - 15
        'Y0 = mshScale.ROWHEIGHT(0) / 2
        Y0 = 0
        Y1 = Y0 + 800
        
        DrawLine picScale, X0 - HOUR_STEP_Twips * 5, Y0, X0 - HOUR_STEP_Twips * 5, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 4, Y0, X0 - HOUR_STEP_Twips * 4, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 2, Y0, X0 - HOUR_STEP_Twips * 2, Y1, &H8000000A
        DrawLine picScale, X0 - HOUR_STEP_Twips * 1, Y0, X0 - HOUR_STEP_Twips * 1, Y1, &H8000000A
        Y0 = 0
        'DrawLine picScale, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picScale, X0, Y0, X0, Y1, &H8000000A, , 2
        '此处的225是随HOUR_STEP_Twips改为260而人为修改有
        'DrawText picScale, X0 - HOUR_STEP_Twips * 6 + 225, 80, "上午", &H80000012
        'DrawText picScale, X0 - HOUR_STEP_Twips * 3 + 225, 80, "下午", &H80000012
        For i = 6 To 1 Step -1
            Select Case i
            Case 6
                strTmp = mlngHourBegin + 4 * 0
                lngColor = &H8080FF
            Case 5
                strTmp = mlngHourBegin + 4 * 1
                lngColor = &H8080FF
            Case 4
                strTmp = mlngHourBegin + 4 * 2
                lngColor = &H80000012
            Case 3
                lngColor = &H80000012
                strTmp = mlngHourBegin + 4 * 3
            Case 2
                lngColor = &H80000012
                strTmp = mlngHourBegin + 4 * 4
            Case 1
                lngColor = &H8080FF
                strTmp = mlngHourBegin + 4 * 5
            End Select
            
            '此处的135是随HOUR_STEP_Twips改为260而人为修改有
            If picScale.Tag <> "" Then
                DrawText picScale, X0 - HOUR_STEP_Twips * i + 135 - picScale.TextWidth(strTmp) / 2, 100, strTmp, IIf(intCol * 6 - i >= intMinCol And intCol * 6 - i <= intMaxCol, lngColor, &H8000000A)
            End If
        Next
    Next
End Function

Public Function DrawPaper() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 在picture上画坐标纸，用于进入和刷新数据之前使用
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim X0 As Long
    Dim Y0 As Long
    Dim X1 As Long
    Dim Y1 As Long
    
    picGraph.Cls
    
    '画纵向坐标图线
    For intCol = 1 To mshUpTab.Cols - 1
        
        X0 = intCol * HOUR_STEP_Twips * 6 - 15
        Y0 = 0
        Y1 = Y0 + 15000
        
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 5, Y0, X0 - HOUR_STEP_Twips * 5, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 4, Y0, X0 - HOUR_STEP_Twips * 4, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 3, Y0, X0 - HOUR_STEP_Twips * 3, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 2, Y0, X0 - HOUR_STEP_Twips * 2, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 1, Y0, X0 - HOUR_STEP_Twips * 1, Y1, &H8000000A
        DrawLine picGraph, X0 - HOUR_STEP_Twips * 0, Y0, X0 - HOUR_STEP_Twips * 0, Y1, &H8000000A, , 2
    Next
    
    '画横向坐标图线
    For intRow = 1 To mshScale.Rows - 1
        
        X0 = 0
        Y0 = (intRow - 1) * ROWHEIGHT * 5
        
        X1 = X0 + 15000

        If (intRow - 1) Mod 5 = 0 Then
            If Int((intRow - 1) / 5) = 5 Then
                DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8080FF, 0, 2
            Else
                DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8000000A, 0, 2
            End If
        Else
            DrawLine picGraph, X0, Y0 + ROWHEIGHT * 5, X1, Y0 + ROWHEIGHT * 5, &H8000000A, 0
        End If
    Next
    

End Function

Private Function Ceil(ByVal dbValue As Double) As Integer
    '******************************************************************************************************************
    '功能： 转换时间为列值
    '参数：
    '返回：
    '******************************************************************************************************************
    
    Ceil = (0 - Int(0 - dbValue))
    Ceil = Int(dbValue + 0.5)
End Function

'Public Function DrawGraph() As Boolean
'    '******************************************************************************************************************
'    '功能： 根据已经填写到表中的数据作图
'    '参数：
'    '返回：
'    '******************************************************************************************************************
'    Dim strComment As String
'    Dim strChar As String
'    Dim dblHeight As Double
'    Dim X0 As Single, Y0 As Single
'    Dim X1 As Single, Y1 As Single
'    Dim y As Single
'    Dim aryValue() As String
'    Dim aryNote() As String
'    Dim aryDots() As String
'    Dim lngColor As Long
'    Dim dblValues As Double
'    Dim strFrom As String, i As Long
'    Dim strDate0 As String, strDate1 As String
'    Dim strtmp As String
'    Dim intPointCount As Integer
'    Dim blnStop As Boolean
'    Dim byt未记显示位置 As Byte
'    Dim mpt脉搏() As POINTAPI
'    Dim mpt心率() As POINTAPI
'    ReDim mpt脉搏(0 To mshScale.Cols - mshScale.FixedCols - 1)
'    ReDim mpt心率(0 To mshScale.Cols - mshScale.FixedCols - 1)
'    Dim rsPoint As ADODB.Recordset
'    Dim rs As New ADODB.Recordset
'    Dim varNote As Variant
'
'    On Error GoTo errHand
'
'    byt未记显示位置 = Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0"))
'
'
'    Call PointInit(rsPoint)
'
'    strFrom = Split(picScale.Tag, ";")(0)
'
'    With mshScale
'        '画线条
'        .Row = 0
'        For intCol = 0 To .FixedCols - 1
'            intPointCount = -1
'            .Col = intCol
'            strChar = Mid(.Tag, intCol + 1, 1)
'            X0 = 0: Y0 = 0: strDate0 = ""
'            blnStop = False
'
'
'            For intCount = 0 To .Cols - .FixedCols - 1
'
'                strDate1 = Format(Int(CDate(strFrom)) + (intCount * 4 + 2) / 24, "yyyy-MM-dd")
'                aryValue = Split(.TextMatrix(GraphDataRow.曲线数据, intCount + .FixedCols), ";")
'
'                If Trim(aryValue(intCol + 1)) <> "" And Val(aryValue(intCol + 1)) > 0 Then
'
'                    X1 = HOUR_STEP_Twips * intCount + (HOUR_STEP_Twips / 2)
'                    aryDots = Split(aryValue(intCol + 1), ",")
'
'                    For i = 0 To UBound(aryDots)
'                        Y1 = aryDots(i)
'                        strChar = Mid(.Tag, intCol + 1, 1)
'                        lngColor = .CellForeColor
'
'                        If X0 <> 0 Then
'                            If i = 0 Then
'
'                                Select Case intCol
'                                Case mItemSerial.体温
'                                    Select Case Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1)
'                                    Case "口温"
'                                        strChar = mstrChar(0)
'                                    Case "腋温"
'                                        strChar = mstrChar(1)
'                                    Case "肛温"
'                                        strChar = mstrChar(2)
'                                    Case Else
'                                        strChar = mstrChar(1)
'                                    End Select
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        '体温35度以下
'                                        strChar = "·"
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'
'                                    Case Is >= GetMaxValue(intCol)
'                                        '体温42度以上
'                                        strChar = "·"
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
'                                        '复试合格
'                                        Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
'                                    End If
'
'                                Case mItemSerial.脉搏
'                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "" Then
'                                        strChar = mstrPulse
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt脉搏(intCount).x = X1
'                                    mpt脉搏(intCount).y = Y1
'
'                                Case mItemSerial.心率
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt心率(intCount).x = X1
'                                    mpt心率(intCount).y = Y1
'
'                                Case mItemSerial.呼吸
'                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "自主呼吸" Then
'                                        strChar = mstrBreath
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End Select
'
'                                '间隔一天无数据才不连线;2.如果中间有说明并要求断线的
'                                If (DateDiff("d", CDate(strDate0), CDate(strDate1)) <= 1) Then
'                                    If blnStop = False Then
'                                        '如果未记说明是"不升",则不与上个结点画连接线
'                                        If intCol = mItemSerial.体温 And Split(.TextMatrix(GraphDataRow.未记说明, intCount + .FixedCols), ";")(intCol + 1) = "不升" Then
'                                            'nothing to do
'                                        Else
'                                            DrawLine picGraph, X0, Y0, X1, Y1, .CellForeColor
'                                        End If
'                                    End If
'                                    blnStop = False
'                                End If
'
'                            Else
'                                Select Case intCol
'                                Case mItemSerial.体温
'
'                                    '物理降温
'                                    lngColor = &HFF&
'                                    strChar = "○"
'                                    If Y1 < Y0 Then
'
'                                        '物理降温失败，画带箭头的红色实线，字符固定用○
'                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, , , True)
'
'                                    ElseIf Y1 > Y0 Then
'
'                                        '物理降温成功，画红色虚线，字符固定用○
'                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, 2)
'
'                                    End If
'
'                                Case mItemSerial.脉搏
'                                    If Y1 <> Y0 Then
'                                        lngColor = &HFF&
'                                        strChar = mstr心率符号
'
'                                        Select Case ConvertToValue(intCol, Y1)
'                                        Case Is <= GetMinValue(intCol)
'                                            Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                        Case Is >= GetMaxValue(intCol)
'                                            Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                        End Select
'
'                                        mpt心率(intCount).x = X1
'                                        mpt心率(intCount).y = Y1
'
'                                    End If
'                                Case mItemSerial.心率
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'
'                                    mpt心率(intCount).x = X1
'                                    mpt心率(intCount).y = Y1
'
'                                Case mItemSerial.呼吸
'                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "自主呼吸" Then
'                                        strChar = mstrBreath
'                                    Else
'                                        strChar = ""
'                                    End If
'
'                                    Select Case ConvertToValue(intCol, Y1)
'                                    Case Is <= GetMinValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                    Case Is >= GetMaxValue(intCol)
'                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                    End Select
'                                End Select
'                            End If
'                        Else
'
'                            Select Case intCol
'                            Case mItemSerial.体温
'                                Select Case Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1)
'                                Case "口温"
'                                    strChar = mstrChar(0)
'                                Case "腋温"
'                                    strChar = mstrChar(1)
'                                Case "肛温"
'                                    strChar = mstrChar(2)
'                                Case Else
'                                    strChar = mstrChar(1)
'                                End Select
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    '体温35度以下
'                                    strChar = "·"
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'
'                                Case Is >= GetMaxValue(intCol)
'                                    '体温42度以上
'                                    strChar = "·"
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'
'                                End Select
'
'                                If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
'                                    '复试合格
'                                    Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
'                                End If
'
'                            Case mItemSerial.脉搏
'                                If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "" Then
'                                    strChar = mstrPulse
'                                Else
'                                    strChar = ""
'                                End If
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                                mpt脉搏(intCount).x = X1
'                                mpt脉搏(intCount).y = Y1
'
'                            Case mItemSerial.心率
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                                mpt心率(intCount).x = X1
'                                mpt心率(intCount).y = Y1
'
'                            Case mItemSerial.呼吸
'                                If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "自主呼吸" Then
'                                    strChar = mstrBreath
'                                Else
'                                    strChar = ""
'                                End If
'
'                                Select Case ConvertToValue(intCol, Y1)
'                                Case Is <= GetMinValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
'                                Case Is >= GetMaxValue(intCol)
'                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
'                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
'                                End Select
'
'                            End Select
'                        End If
'
'                        If intCol = mItemSerial.体温 Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols))
'                        ElseIf intCol = mItemSerial.呼吸 Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "BREATH", ""))
'                        ElseIf intCol = mItemSerial.脉搏 Then
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "PACEMAKER", ""))
'                        Else
'                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, "")
'                        End If
'
'                        '记录上次画点位置和日期
'                        If X0 <> 0 And i <> 0 Then X1 = X0: Y1 = Y0 '从第一点与下一点连接
'                        X0 = X1: Y0 = Y1: strDate0 = strDate1
'
'                        blnStop = False
'                    Next i
'                End If
'
''                If blnStop = False Then
''                    If (.TextMatrix(GraphDataRow.上标说明, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.下标说明, intCount + .FixedCols) <> "") Then
''                        blnStop = (Val(.TextMatrix(GraphDataRow.断开标志, intCount + .FixedCols)) = 1)
''                    End If
''                End If
'
'                If blnStop = False Then
'
'                    aryNote = Split(.TextMatrix(GraphDataRow.未记说明, intCount + .FixedCols), ";")
'                    blnStop = (Trim(aryNote(intCol + 1)) <> "")
'
'                    '以下本来是可以不要了的，但考虑到以前的数据
'                    If blnStop = False Then
'                        If (.TextMatrix(GraphDataRow.上标说明, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.下标说明, intCount + .FixedCols) <> "") Then
'                            blnStop = (Val(.TextMatrix(GraphDataRow.断开标志, intCount + .FixedCols)) = 1)
'                        End If
'                    End If
'
'                End If
'
'            Next intCount
'
'        Next intCol
'
'
'        '画点的字符或图形
'        '--------------------------------------------------------------------------------------------------------------
'        Call DrawPoint(picGraph, rsPoint)
'
'        '根据脉搏和心率坐标形成多边形，并进行连线和填充
'        '--------------------------------------------------------------------------------------------------------------
'        Call DrawPoly(picGraph, mpt脉搏, mpt心率)
'
'        Dim lngYMax As Long
'        lngYMax = ConvertToY(mItemSerial.体温, 34.2)
'
'        '打印入出转标志
'        '--------------------------------------------------------------------------------------------------------------
'        Dim intLoop As Integer
'        Dim rsTmp As ADODB.Recordset
'
'        '20090926:必须在40-42度间打印,单独一条信息如果超长就缩小字体,有多条信息则延后面一格打印,如果是最后一格就直接全部打印
'        Set rsTmp = New ADODB.Recordset
'        rsTmp.Fields.Append "列号", adVarChar, 30
'        rsTmp.Fields.Append "时间", adVarChar, 30
'        rsTmp.Fields.Append "结果", adVarChar, 50
'        '20090926--
'        rsTmp.Fields.Append "打印列", adVarChar, 30
'        rsTmp.Fields.Append "字体", adVarChar, 50
'        '----------
'        rsTmp.Open
'
'        Dim intCharNumber As Integer
'
'        For intCol = 0 To .Cols - .FixedCols - 1
'
'            X1 = HOUR_STEP_Twips * intCol + HOUR_STEP_Twips / 2
''            Y1 = ConvertToY(mItemSerial.体温, 42)
'            Y1 = 195
'            dblHeight = lngYMax - Y1
'
'            '行号:=3表示手术;=5表示入院;=6表示转科;=7表示换床;=8表示出院,=13出生
'            rsTmp.Filter = ""
'            For intLoop = 5 To 9
'                rsTmp.AddNew
'                rsTmp.Fields("列号").Value = intCol + .FixedCols
'                rsTmp.Fields("时间").Value = .Cell(flexcpData, intLoop, intCol + .FixedCols, intLoop, intCol + .FixedCols)
'                rsTmp.Fields("结果").Value = .TextMatrix(intLoop, intCol + .FixedCols)
'            Next
'            rsTmp.AddNew
'            rsTmp.Fields("列号").Value = intCol + .FixedCols
'            rsTmp.Fields("时间").Value = .Cell(flexcpData, 3, intCol + .FixedCols, 3, intCol + .FixedCols)
'            rsTmp.Fields("结果").Value = .TextMatrix(3, intCol + .FixedCols)
'
'            rsTmp.AddNew
'            rsTmp.Fields("列号").Value = intCol + .FixedCols
'            rsTmp.Fields("时间").Value = .Cell(flexcpData, 13, intCol + .FixedCols, 13, intCol + .FixedCols)
'            rsTmp.Fields("结果").Value = .TextMatrix(13, intCol + .FixedCols)
'
'            '一条信息打一列
'            strComment = ""
'            rsTmp.Filter = "列号=" & intCol + .FixedCols
'            If rsTmp.RecordCount > 0 Then
'                rsTmp.Sort = "时间"
'                rsTmp.MoveFirst
'                Do While Not rsTmp.EOF
'                    If strComment = "" Then
'                        strComment = rsTmp.Fields("结果").Value
'                    Else
'                        strComment = Trim(strComment) & " " & rsTmp.Fields("结果").Value
'                    End If
'                    rsTmp.MoveNext
'                Loop
'            End If
'
'            If Trim(strComment) <> "" Then
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'
'                    If Y1 < lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '红色
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, 255)
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'            '未记说明
'            '----------------------------------------------------------------------------------------------------------
'            If byt未记显示位置 = 0 Then
'                strComment = IIf(Trim(strComment) = "", "", " ")
'                strtmp = ""
'                varNote = Split(.TextMatrix(GraphDataRow.未记说明, intCol + .FixedCols), ";")
'                For intCount = 0 To UBound(varNote)
'                    If varNote(intCount) <> "不升" Then
'                        If InStr(";" & strtmp & ";", ";" & varNote(intCount) & ";") = 0 Then
'                            strtmp = strtmp & ";" & varNote(intCount)
'                        End If
'                    End If
'                Next
'                If strtmp <> "" Then
'                    varNote = Split(strtmp, ";")
'                    For intCount = 0 To UBound(varNote)
'                        If strComment = "" Or strComment = " " Then
'                            strComment = strComment & varNote(intCount)
'                        Else
'                            strComment = strComment & " " & varNote(intCount)
'                        End If
'                    Next
'                End If
'
'                If Trim(strComment) <> "" Then
'
'                    intCharNumber = 0
'                    For intCount = 1 To Len(strComment)
'                        If Y1 <= lngYMax Then
'                            strChar = Mid(strComment, intCount, 1)
'                            '蓝色
'
'                            If Asc(strChar) < 0 Then
'                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                            End If
'
'                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                            If Asc(strChar) < 0 Then
'                                intCharNumber = 0
'                                Y1 = Y1 + ROWHEIGHT * 5
'                            Else
'                                Y1 = Y1 + ROWHEIGHT * 2.5
'                                intCharNumber = intCharNumber + 1
'                            End If
'                        End If
'                    Next
'                End If
'            End If
'
'            '上标说明
'            '----------------------------------------------------------------------------------------------------------
'            strComment = IIf(Trim(strComment) = "", "", " ") & Trim(.TextMatrix(GraphDataRow.上标说明, intCol + .FixedCols))
'            If Trim(strComment) <> "" Then
'
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'                    If Y1 <= lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '蓝色
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'            '下标说明
'            '----------------------------------------------------------------------------------------------------------
'
''            Y1 = ConvertToY(mItemSerial.体温, 35)
'            Y1 = 7020
'            strComment = ""
'
'            '未记说明
'            '----------------------------------------------------------------------------------------------------------
'            If byt未记显示位置 = 1 Then
'                strComment = IIf(Trim(strComment) = "", "", " ")
'                strtmp = ""
'                varNote = Split(.TextMatrix(GraphDataRow.未记说明, intCol + .FixedCols), ";")
'                For intCount = 0 To UBound(varNote)
'                    If varNote(intCount) <> "不升" Then
'                        If InStr(";" & strtmp & ";", ";" & varNote(intCount) & ";") = 0 Then
'                            strtmp = strtmp & ";" & varNote(intCount)
'                        End If
'                    End If
'                Next
'                If strtmp <> "" Then
'                    varNote = Split(strtmp, ";")
'                    For intCount = 0 To UBound(varNote)
'                        If strComment = "" Or strComment = " " Then
'                            strComment = strComment & varNote(intCount)
'                        Else
'                            strComment = strComment & " " & varNote(intCount)
'                        End If
'                    Next
'                End If
'
'                If Trim(strComment) <> "" Then
'
'                    intCharNumber = 0
'                    For intCount = 1 To Len(strComment)
'                        If Y1 <= lngYMax Then
'                            strChar = Mid(strComment, intCount, 1)
'                            '蓝色
'
'                            If Asc(strChar) < 0 Then
'                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                            End If
'
'                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                            If Asc(strChar) < 0 Then
'                                intCharNumber = 0
'                                Y1 = Y1 + ROWHEIGHT * 5
'                            Else
'                                Y1 = Y1 + ROWHEIGHT * 2.5
'                                intCharNumber = intCharNumber + 1
'                            End If
'                        End If
'                    Next
'                End If
'            End If
'
'            strComment = IIf(Trim(strComment) = "", "", " ") & .TextMatrix(GraphDataRow.下标说明, intCol + .FixedCols)
'            If Trim(strComment) <> "" Then
'                intCharNumber = 0
'                For intCount = 1 To Len(strComment)
'                    If Y1 <= lngYMax Then
'                        strChar = Mid(strComment, intCount, 1)
'                        '蓝色
'
'                        If Asc(strChar) < 0 Then
'                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
'                        End If
'
'                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
'
'                        If Asc(strChar) < 0 Then
'                            intCharNumber = 0
'                            Y1 = Y1 + ROWHEIGHT * 5
'                        Else
'                            Y1 = Y1 + ROWHEIGHT * 2.5
'                            intCharNumber = intCharNumber + 1
'                        End If
'                    End If
'                Next
'            End If
'
'        Next
'
'    End With
'
'    Exit Function
'
'    '------------------------------------------------------------------------------------------------------------------
'errHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'
'End Function

Public Function DrawGraph() As Boolean
    '******************************************************************************************************************
    '功能： 根据已经填写到表中的数据作图
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strComment As String
    Dim strChar As String, strChar1 As String
    Dim dblHeight As Double         '40-42度之间的有效打印高度
    Dim X0 As Single, Y0 As Single
    Dim X1 As Single, Y1 As Single
    Dim Y As Single
    Dim aryValue() As String
    Dim aryNote() As String
    Dim aryDots() As String
    Dim lngColor As Long
    Dim dblValues As Double
    Dim strFrom As String, i As Long
    Dim strDate0 As String, strDate1 As String
    Dim strTmp As String
    Dim intPointCount As Integer
    Dim blnStop As Boolean
    Dim byt未记显示位置 As Byte
    Dim mpt脉搏() As POINTAPI
    Dim mpt心率() As POINTAPI
    ReDim mpt脉搏(0 To mshScale.Cols - mshScale.FixedCols - 1)
    ReDim mpt心率(0 To mshScale.Cols - mshScale.FixedCols - 1)
    Dim rsPoint As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim varNote As Variant
    
    On Error GoTo errHand
    
    byt未记显示位置 = Val(zlDatabase.GetPara("未记说明显示位置", glngSys, 1255, "0"))
           

    Call PointInit(rsPoint)
    
    strFrom = Split(picScale.Tag, ";")(0)
    
    With mshScale
        '画线条
        .Row = 0
        For intCol = 0 To .FixedCols - 1
            intPointCount = -1
            .Col = intCol
            strChar = Mid(.Tag, intCol + 1, 1)
            X0 = 0: Y0 = 0: strDate0 = "": strChar1 = mstrBreath
            blnStop = False
            
            
            For intCount = 0 To .Cols - .FixedCols - 1
                
                strDate1 = Format(Int(CDate(strFrom)) + (intCount * 4 + 2) / 24, "yyyy-MM-dd")
                aryValue = Split(.TextMatrix(GraphDataRow.曲线数据, intCount + .FixedCols), ";")

                If Trim(aryValue(intCol + 1)) <> "" And Val(aryValue(intCol + 1)) > 0 Then
                
                    X1 = HOUR_STEP_Twips * intCount + (HOUR_STEP_Twips / 2)
                    aryDots = Split(aryValue(intCol + 1), ",")
                    
                    For i = 0 To UBound(aryDots)
                        Y1 = aryDots(i)
                        strChar = Mid(.Tag, intCol + 1, 1)
                        lngColor = .CellForeColor
                        
                        If X0 <> 0 Then
                            If i = 0 Then
                                
                                Select Case intCol
                                Case mItemSerial.体温
                                    Select Case Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1)
                                    Case "口温"
                                        strChar = mstrChar(0)
                                    Case "腋温"
                                        strChar = mstrChar(1)
                                    Case "肛温"
                                        strChar = mstrChar(2)
                                    Case Else
                                        strChar = mstrChar(1)
                                    End Select

                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        '体温35度以下
                                        strChar = "·"
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                        
                                    Case Is >= GetMaxValue(intCol)
                                        '体温42度以上
                                        strChar = "·"
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
                                        '复试合格
                                        Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
                                    End If
                                    
                                Case mItemSerial.脉搏
                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "" Then
                                        strChar = mstrPulse
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt脉搏(intCount).X = X1
                                    mpt脉搏(intCount).Y = Y1
                                    
                                Case mItemSerial.心率
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt心率(intCount).X = X1
                                    mpt心率(intCount).Y = Y1
                                    
                                Case mItemSerial.呼吸
                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) <> "呼吸机" Then
                                        strChar = mstrBreath
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End Select
                                
                                '间隔一天无数据才不连线;2.如果中间有说明并要求断线的
                                '45987,刘鹏飞,2012-09-10,湖南需求
                                '1.呼吸显示为曲线；2、呼吸统一用黑色实行点（●）表示，不在体温单上显示呼吸数字
                                '3. 使用呼吸机的患者，呼吸以黑R表示，在相应时间内呼吸30次横线下顶格用黑笔划R，相邻的R之间以及R和自主呼吸之间不连线
                                If (DateDiff("d", CDate(strDate0), CDate(strDate1)) <= 1) Then
                                    If blnStop = False Then
                                        '如果未记说明是"不升",则不与上个结点画连接线
                                        If intCol = mItemSerial.体温 And Split(.TextMatrix(GraphDataRow.未记说明, intCount + .FixedCols), ";")(intCol + 1) = "不升" Then
                                            'nothing to do
                                        ElseIf intCol = mItemSerial.呼吸 And (strChar = "" Or strChar1 = "") Then
                                            'nothing to do
                                        Else
                                            DrawLine picGraph, X0, Y0, X1, Y1, .CellForeColor
                                        End If
                                    End If
                                    blnStop = False
                                End If
    
                            Else
                                Select Case intCol
                                Case mItemSerial.体温
                                    
                                    '物理降温
                                    lngColor = &HFF&
                                    strChar = "○"
                                    If Y1 < Y0 Then
                                    
                                        '物理降温失败，画带箭头的红色实线，字符固定用○
                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, , , True)
                                        
                                    ElseIf Y1 > Y0 Then
                                    
                                        '物理降温成功，画红色虚线，字符固定用○
                                        Call DrawLine(picGraph, X0, Y0, X1, Y1, lngColor, 2)
                                        
                                    End If
                                    
                                Case mItemSerial.脉搏
                                    If Y1 <> Y0 Then
                                        lngColor = &HFF&
                                        strChar = mstr心率符号

                                        Select Case ConvertToValue(intCol, Y1)
                                        Case Is <= GetMinValue(intCol)
                                            Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                        Case Is >= GetMaxValue(intCol)
                                            Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                            Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                        End Select
                                        
                                        mpt心率(intCount).X = X1
                                        mpt心率(intCount).Y = Y1
                                        
                                    End If
                                Case mItemSerial.心率

                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                    
                                    mpt心率(intCount).X = X1
                                    mpt心率(intCount).Y = Y1
                                
                                Case mItemSerial.呼吸
                                    If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) <> "呼吸机" Then
                                        strChar = mstrBreath
                                    Else
                                        strChar = ""
                                    End If
                                    
                                    Select Case ConvertToValue(intCol, Y1)
                                    Case Is <= GetMinValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    Case Is >= GetMaxValue(intCol)
                                        Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                        Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    End Select
                                End Select
                            End If
                        Else

                            Select Case intCol
                            Case mItemSerial.体温
                                Select Case Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1)
                                Case "口温"
                                    strChar = mstrChar(0)
                                Case "腋温"
                                    strChar = mstrChar(1)
                                Case "肛温"
                                    strChar = mstrChar(2)
                                Case Else
                                    strChar = mstrChar(1)
                                End Select

                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    '体温35度以下
                                    strChar = "·"
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                    
                                Case Is >= GetMaxValue(intCol)
                                    '体温42度以上
                                    strChar = "·"
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                    
                                End Select
                                
                                If Val(.TextMatrix(10, intCount + .FixedCols)) = 1 Then
                                    '复试合格
                                    Call DrawText(picGraph, X1 - 50, Y1 - 250, "v", lngColor)
                                End If
                                    
                            Case mItemSerial.脉搏
                                If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) = "" Then
                                    strChar = mstrPulse
                                Else
                                    strChar = ""
                                End If
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                
                                mpt脉搏(intCount).X = X1
                                mpt脉搏(intCount).Y = Y1
                                
                            Case mItemSerial.心率
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                    
                                mpt心率(intCount).X = X1
                                mpt心率(intCount).Y = Y1
                                
                            Case mItemSerial.呼吸
                                If Split(.TextMatrix(GraphDataRow.部位标志, intCount + .FixedCols), ";")(intCol + 1) <> "呼吸机" Then
                                    strChar = mstrBreath
                                Else
                                    strChar = ""
                                End If
                                
                                Select Case ConvertToValue(intCol, Y1)
                                Case Is <= GetMinValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMinValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 + 200, lngColor, , , True)
                                Case Is >= GetMaxValue(intCol)
                                    Y1 = ConvertToY(intCol, GetMaxValue(intCol))
                                    Call DrawLine(picGraph, X1, Y1, X1, Y1 - 200, lngColor, , , True)
                                End Select
                                
                            End Select
                        End If
                        
                        If intCol = mItemSerial.体温 Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols))
                        ElseIf intCol = mItemSerial.呼吸 Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "BREATH", ""))
                        ElseIf intCol = mItemSerial.脉搏 Then
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, mshScale.TextMatrix(4, intCount + .FixedCols), IIf(strChar = "", "PACEMAKER", ""))
                        Else
                            Call PointAdd(rsPoint, X1, Y1, mshScale.ColData(intCol), strChar, lngColor, intCount, "")
                        End If
                                                   
                        '记录上次画点位置和日期
                        If X0 <> 0 And i <> 0 Then X1 = X0: Y1 = Y0 '从第一点与下一点连接
                        X0 = X1: Y0 = Y1: strDate0 = strDate1: strChar1 = strChar 'strChar1变量目前只针对呼吸
                        
                        blnStop = False
                    Next i
                End If
                    
'                If blnStop = False Then
'                    If (.TextMatrix(GraphDataRow.上标说明, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.下标说明, intCount + .FixedCols) <> "") Then
'                        blnStop = (Val(.TextMatrix(GraphDataRow.断开标志, intCount + .FixedCols)) = 1)
'                    End If
'                End If
                
                If blnStop = False Then
                    
                    aryNote = Split(.TextMatrix(GraphDataRow.未记说明, intCount + .FixedCols), ";")
                    blnStop = (Trim(aryNote(intCol + 1)) <> "")
                    
                    '以下本来是可以不要了的，但考虑到以前的数据
                    If blnStop = False Then
                        If (.TextMatrix(GraphDataRow.上标说明, intCount + .FixedCols) <> "" Or .TextMatrix(GraphDataRow.下标说明, intCount + .FixedCols) <> "") Then
                            blnStop = (Val(.TextMatrix(GraphDataRow.断开标志, intCount + .FixedCols)) = 1)
                        End If
                    End If
                    
                End If
                
            Next intCount
            
        Next intCol
        
        
        '画点的字符或图形
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoint(picGraph, rsPoint, mItemSerial.体温)
        
        '根据脉搏和心率坐标形成多边形，并进行连线和填充
        '--------------------------------------------------------------------------------------------------------------
        Call DrawPoly(picGraph, mpt脉搏, mpt心率)

        Dim lngYMax As Long
        If mItemSerial.体温 <> -1 Then
            lngYMax = ConvertToY(mItemSerial.体温, 33.4)
        Else
            lngYMax = 8580
        End If
        
        '打印入出转标志
        '--------------------------------------------------------------------------------------------------------------
        Dim intLoop As Integer
        Dim rsTmp As ADODB.Recordset
        
        '20090926:必须在40-42度间打印,单独一条信息如果超长就缩小字体,有多条信息则延后面一格打印,如果是最后一格就直接全部打印
        Set rsTmp = New ADODB.Recordset
        rsTmp.Fields.Append "列号", adDouble, 30
        rsTmp.Fields.Append "时间", adVarChar, 30
        rsTmp.Fields.Append "结果", adVarChar, 50
        '20090926--
        rsTmp.Fields.Append "类型", adVarChar, 50       '记录是入出转,手术出院,还是未记说明,上标说明
        rsTmp.Fields.Append "打印列", adVarChar, 30
        rsTmp.Fields.Append "坐标", adVarChar, 30
        rsTmp.Fields.Append "高度", adVarChar, 30       '未记说明及上标说明不用管高度
        rsTmp.Fields.Append "字体大小", adVarChar, 50
        '----------
        rsTmp.Open

        Dim intCharNumber As Integer
        
        For intCol = 0 To .Cols - .FixedCols - 1
            
            X1 = HOUR_STEP_Twips * intCol + HOUR_STEP_Twips / 2
'            Y1 = ConvertToY(mItemSerial.体温, 42)
            Y1 = 195
            If mItemSerial.体温 <> -1 Then
                dblHeight = ConvertToY(mItemSerial.体温, 40) - Y1
            Else
                dblHeight = 2145 - Y1
            End If
            
            '行号:=3表示手术;=5表示入院;=6表示转科;=7表示换床;=8表示出院,=13出生
            rsTmp.Filter = ""
            For intLoop = 5 To 9
                If .TextMatrix(intLoop, intCol + .FixedCols) <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("类型").Value = intLoop
                    rsTmp.Fields("坐标").Value = X1 & ";" & Y1
                    rsTmp.Fields("列号").Value = intCol
                    rsTmp.Fields("时间").Value = .Cell(flexcpData, intLoop, intCol + .FixedCols, intLoop, intCol + .FixedCols)
                    rsTmp.Fields("结果").Value = .TextMatrix(intLoop, intCol + .FixedCols)
                End If
            Next
            If .TextMatrix(手术标志, intCol + .FixedCols) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("类型").Value = 手术标志
                rsTmp.Fields("坐标").Value = X1 & ";" & Y1
                rsTmp.Fields("列号").Value = intCol
                rsTmp.Fields("时间").Value = .Cell(flexcpData, 手术标志, intCol + .FixedCols, 手术标志, intCol + .FixedCols)
                rsTmp.Fields("结果").Value = .TextMatrix(手术标志, intCol + .FixedCols)
            End If
            If .TextMatrix(出生标志, intCol + .FixedCols) <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("类型").Value = 出生标志
                rsTmp.Fields("坐标").Value = X1 & ";" & Y1
                rsTmp.Fields("列号").Value = intCol
                rsTmp.Fields("时间").Value = .Cell(flexcpData, 出生标志, intCol + .FixedCols, 出生标志, intCol + .FixedCols)
                rsTmp.Fields("结果").Value = .TextMatrix(出生标志, intCol + .FixedCols)
            End If
            

            '未记说明
            '----------------------------------------------------------------------------------------------------------
            If byt未记显示位置 = 0 Then
                strComment = ""
                strTmp = ""
                varNote = Split(.TextMatrix(GraphDataRow.未记说明, intCol + .FixedCols), ";")
                For intCount = 0 To UBound(varNote)
                    If varNote(intCount) <> "不升" Then
                        If InStr(";" & strTmp & ";", ";" & varNote(intCount) & ";") = 0 Then
                            strTmp = strTmp & ";" & varNote(intCount)
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    varNote = Split(strTmp, ";")
                    For intCount = 0 To UBound(varNote)
                        If strComment = "" Or strComment = " " Then
                            strComment = strComment & varNote(intCount)
                        Else
                            strComment = strComment & " " & varNote(intCount)
                        End If
                    Next
                End If
                If strComment <> "" Then
                    rsTmp.AddNew
                    rsTmp.Fields("类型").Value = 未记说明
                    rsTmp.Fields("坐标").Value = X1 & ";" & Y1
                    rsTmp.Fields("列号").Value = intCol
                    rsTmp.Fields("时间").Value = .Cell(flexcpData, 未记说明, intCol + .FixedCols, 未记说明, intCol + .FixedCols)
                    rsTmp.Fields("结果").Value = strComment
                End If
            End If
            
            '上标说明
            '----------------------------------------------------------------------------------------------------------
            strComment = Trim(.TextMatrix(GraphDataRow.上标说明, intCol + .FixedCols))
            If strComment <> "" Then
                rsTmp.AddNew
                rsTmp.Fields("类型").Value = 上标说明
                rsTmp.Fields("坐标").Value = X1 & ";" & Y1
                rsTmp.Fields("列号").Value = intCol
                rsTmp.Fields("时间").Value = .Cell(flexcpData, 上标说明, intCol + .FixedCols, 上标说明, intCol + .FixedCols)
                rsTmp.Fields("结果").Value = strComment
            End If
            
            '下标说明
            '----------------------------------------------------------------------------------------------------------
            
'            Y1 = ConvertToY(mItemSerial.体温, 35)
            Y1 = 7020
            strComment = ""
            
            '未记说明
            '----------------------------------------------------------------------------------------------------------
            If byt未记显示位置 = 1 Then
                strComment = ""
                strTmp = ""
                varNote = Split(.TextMatrix(GraphDataRow.未记说明, intCol + .FixedCols), ";")
                For intCount = 0 To UBound(varNote)
                    If varNote(intCount) <> "不升" Then
                        If InStr(";" & strTmp & ";", ";" & varNote(intCount) & ";") = 0 Then
                            strTmp = strTmp & ";" & varNote(intCount)
                        End If
                    End If
                Next
                If strTmp <> "" Then
                    varNote = Split(strTmp, ";")
                    For intCount = 0 To UBound(varNote)
                        If strComment = "" Or strComment = " " Then
                            strComment = strComment & varNote(intCount)
                        Else
                            strComment = strComment & " " & varNote(intCount)
                        End If
                    Next
                End If

                If Trim(strComment) <> "" Then
                    
                    intCharNumber = 0
                    For intCount = 1 To Len(strComment)
                        If Y1 <= lngYMax Then
                            strChar = Mid(strComment, intCount, 1)
                            '蓝色
                            
                            If Asc(strChar) < 0 Then
                                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
                            End If
                            
                            Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
                            
                            If Asc(strChar) < 0 Then
                                intCharNumber = 0
                                Y1 = Y1 + ROWHEIGHT * 5
                            Else
                                Y1 = Y1 + ROWHEIGHT * 2.5
                                intCharNumber = intCharNumber + 1
                            End If
                        End If
                    Next
                End If
            End If

            strComment = IIf(Trim(strComment) = "", "", " ") & .TextMatrix(GraphDataRow.下标说明, intCol + .FixedCols)
            If Trim(strComment) <> "" Then
                intCharNumber = 0
                For intCount = 1 To Len(strComment)
                    If Y1 <= lngYMax Then
                        strChar = Mid(strComment, intCount, 1)
                        '蓝色
                        
                        If Asc(strChar) < 0 Then
                            If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
                        End If
                        
                        Call DrawRotateText(picGraph, X1 - picGraph.TextWidth(strChar) / 2, Y1 + 15, strChar, -2147483635)
                        
                        If Asc(strChar) < 0 Then
                            intCharNumber = 0
                            Y1 = Y1 + ROWHEIGHT * 5
                        Else
                            Y1 = Y1 + ROWHEIGHT * 2.5
                            intCharNumber = intCharNumber + 1
                        End If
                    End If
                Next
            End If
            
        Next

    End With
    
    Call OutputNote(picGraph, dblHeight, rsTmp)
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function GetFontSize(ByVal objDraw As Object, ByVal dblHeight As Double, ByVal strText As String, ByRef Y1 As Single) As Single
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    '计算合理的字体大小
    
    sinFontSize_Bak = objDraw.FontSize
    For sinFontSize = objDraw.FontSize To 5 Step -1
        Y1 = 0
        intCharNumber = 0
        For intCount = 1 To Len(strText)
            strChar = Mid(strText, intCount, 1)
            
            If Asc(strChar) < 0 Then
                If intCharNumber Mod 2 = 1 Then Y1 = Y1 + ROWHEIGHT * 2.5
            End If
            
            If Asc(strChar) < 0 Then
                intCharNumber = 0
                Y1 = Y1 + ROWHEIGHT * 5
            Else
                Y1 = Y1 + ROWHEIGHT * 2.5
                intCharNumber = intCharNumber + 1
            End If
        Next
        'If Y1 <= dblHeight Then Exit For
        Exit For
    Next
    
    objDraw.FontSize = sinFontSize_Bak
    GetFontSize = sinFontSize
End Function

Private Sub OutputNote(ByVal objDraw As Object, ByVal dblHeight As Double, ByRef rsNote As ADODB.Recordset)
    '输出以下信息:入院,入科,转科,出院,手术分娩,未记说明,上标说明及出生
    '未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印
    '除未记说明及上标说明外,入出转等信息当一个刻度发生多个时,依次写入各个刻度中,如其它刻度也有信息,顺移
    Dim intCol As Integer                   '记录当前列号
    Dim intMax As Integer                   '总列数
    Dim intCur As Integer                   '当前记录的位置
    Dim bln上标 As Boolean
    Dim sinX1 As Single, sinY1 As Single, sinHeight As Single, sinMaxY1 As Single
    Dim rsTarget As New ADODB.Recordset
    
    '输出字符相关变量定义
    Dim sinFontSize As Single
    Dim sinFontSize_Bak As Single
    Dim intCharNumber As Integer
    Dim intCount As Integer
    Dim strChar As String
    
    intMax = mshScale.Cols - mshScale.FixedCols - 1
    sinFontSize_Bak = objDraw.FontSize
    Set rsTarget = rsNote.Clone
    With rsNote
        If .RecordCount = 0 Then Exit Sub
        .Sort = "列号,时间"
        intCol = !列号
        
        '先在入出转手术等中循环
        Do While Not .EOF
            If Trim(NVL(!结果)) <> "" Then
                If Not (!类型 = 未记说明 Or !类型 = 上标说明) Then
                    '检查待打印列是否已存在输出,如果存在则校正坐标
                    If intCol > intMax Then intCol = intMax
                    
                    '计算得到合适的字体大小及高度
                    !字体大小 = GetFontSize(objDraw, dblHeight, NVL(!结果), sinY1)
                    !高度 = sinY1
                    !打印列 = IIf(intCol < !列号, !列号, intCol)
                    .Update
                    If intCol <= !列号 Then intCol = !列号
                    intCol = intCol + 1
                Else
                    Call GetFontSize(objDraw, dblHeight, NVL(!结果), sinY1)
                    !高度 = sinY1
                    .Update
                End If
            End If
            
            .MoveNext
        Loop
        .MoveFirst
        
        '调整入出转等的纵坐标(只有最后一列才存在一格打完的情况)
        sinY1 = 195
        .Filter = "打印列='" & intMax & "'"
        .Sort = "列号,时间"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '只有入出转手术才更新了打印列
            !坐标 = Split(!坐标, ";")(0) & ";" & sinY1
            .Update
            sinY1 = sinY1 + !高度 + 100
            
            .MoveNext
        Loop
        .Filter = 0
        .MoveFirst
        
        '重新校正未记说明以及上标说明的高度(未记说明及上标说明,在没有入出转手术分娩及出生的信息时,打印在42-40之间;否则从40开始向下打印)
        Set rsTarget = .Clone
        intCol = 0
        Do While Not .EOF
            If (!类型 = 未记说明 Or !类型 = 上标说明) Then
                bln上标 = False
                Set rsTarget = .Clone
                rsTarget.Filter = "打印列='" & !列号 & "'"
                If rsTarget.RecordCount <> 0 Then
                    '已存在打印内容的才校正纵坐标
                    sinMaxY1 = Split(rsTarget!坐标, ";")(1)
                    Do While Not rsTarget.EOF
                        If bln上标 = False Then
                            '考虑到上标有可能是在40度开始打的,所以需校正一下sinMaxY1的坐标
                            bln上标 = (rsTarget!类型 = 未记说明 Or rsTarget!类型 = 上标说明)
                            If bln上标 Then sinMaxY1 = Split(rsTarget!坐标, ";")(1)
                        End If
                        sinMaxY1 = sinMaxY1 + rsTarget!高度 + 100
                        rsTarget.MoveNext
                    Loop
                    If mItemSerial.体温 <> -1 Then
                        sinY1 = ConvertToY(mItemSerial.体温, 40)
                    Else
                        sinY1 = 2145
                    End If
                    If sinY1 < sinMaxY1 Or bln上标 Then sinY1 = sinMaxY1
                    sinHeight = !高度
                    intCol = !列号
                Else
                    sinY1 = 195
                    intCol = !列号
                    sinHeight = !高度
                End If
                rsTarget.Filter = 0
                
                !坐标 = Split(!坐标, ";")(0) & ";" & sinY1
                !打印列 = !列号                                 '此时更新打印列,以便上面的循环过滤
                .Update
            End If
            .MoveNext
        Loop
    
        '开始按数据输出内容
        .MoveFirst
        Do While Not .EOF
            If Trim(NVL(!结果)) <> "" Then
                'If (!类型 = 未记说明 Or !类型 = 上标说明) Then Stop
                sinX1 = HOUR_STEP_Twips * (IIf(!打印列 = "", Val(!列号), Val(!打印列))) + HOUR_STEP_Twips / 2
                sinY1 = Split(!坐标, ";")(1)
                intCharNumber = 0
                objDraw.FontSize = IIf(!字体大小 = "", 9, !字体大小)
                
                For intCount = 1 To Len(!结果)
                    strChar = Mid(!结果, intCount, 1)
                    
                    If Asc(strChar) < 0 Then
                        If intCharNumber Mod 2 = 1 Then sinY1 = sinY1 + ROWHEIGHT * 2.5
                    End If
                    Call DrawRotateText(objDraw, sinX1 - objDraw.TextWidth(strChar) / 2, sinY1 + 15, strChar, IIf(!类型 = 未记说明 Or !类型 = 上标说明, -2147483635, 255))
                    If Asc(strChar) < 0 Then
                        intCharNumber = 0
                        sinY1 = sinY1 + ROWHEIGHT * 5
                    Else
                        sinY1 = sinY1 + ROWHEIGHT * 2.5
                        intCharNumber = intCharNumber + 1
                    End If
                Next
            End If
            
            .MoveNext
        Loop
    End With
    objDraw.FontSize = sinFontSize_Bak
End Sub

Private Function isSaved() As String
    '------------------------------------------------------------------------------------------------------------------
    '功能： 判断体温表是否保存，如果未保存返回提示信息，已经保存返回零长度字符串
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String
    Dim strItem As String
    On Error Resume Next
    
    With mshUpTab
        For intCol = .FixedCols To .Cols - 1
            If .ColData(intCol) = 2 Or .ColData(intCol) = 3 Then
                isSaved = "改变了“手术日”尚未保存。"
                Exit Function
            End If
        Next
    End With
    With mshScale
        For intCol = .FixedCols To .Cols - 1
            aryValue = Split(.TextMatrix(0, intCol), ";")
            For intCount = 0 To UBound(aryValue)
                If intCount = 0 Then
                    strItem = "说明"
                Else
                    strItem = .TextMatrix(0, intCount - 1)
                End If
                If aryValue(intCount) = "2" Or aryValue(intCount) = "3" Or aryValue(intCount) = "4" Then
                    isSaved = "改变了“" & strItem & "”数据尚未保存。"
                    Exit Function
                End If
            Next
        Next
    End With
    With mshDownTab
        For intCol = .FixedCols To .Cols - 1
            aryValue = Split(.TextMatrix(0, intCol), ";")
            For intCount = 0 To UBound(aryValue)
                strItem = .TextMatrix(intCount + 1, 1)
                If aryValue(intCount) = "2" Or aryValue(intCount) = "3" Or aryValue(intCount) = "4" Then
                    isSaved = "改变了“" & strItem & "”数据尚未保存"
                    Exit Function
                End If
            Next
        Next
    End With
End Function


Private Function CalcMinMaxCol(ByVal strDate As String, MinCol As Long, MaxCol As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 获得最小最大时间范围
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim aryValue() As String
    Dim dtTmp As Date
    Dim strTmp As String
    
    If mvarEdit = False Then Exit Function
    
    aryValue = Split(strDate, ";")
    
    MinCol = GetCurveColumn(CDate(aryValue(0)), CDate(aryValue(0)), mlngHourBegin) - 1
    MaxCol = GetCurveColumn(CDate(aryValue(1)), CDate(aryValue(0)), mlngHourBegin) - 1
    
End Function

Private Function SetColBkColor(ByVal Col As Long, ByVal COLOR As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    mshScale.Redraw = False
    mshScale.Col = Col
    For i = 0 To mshScale.Rows - 1
        mshScale.Row = i
        mshScale.CellBackColor = COLOR
    Next
    mshScale.Redraw = True
End Function

Private Function SetVisible() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("编辑")) = 0 Then
        mshUpTab.Enabled = False
        mshDownTab.Enabled = False
        picBack.Enabled = False
        picScale.Enabled = False
    Else
        mshUpTab.Enabled = True
        mshDownTab.Enabled = True
        picBack.Enabled = True
        picScale.Enabled = True
    End If
End Function

Private Function SetBodyMode() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能： 设置体温单是显示模式还是编辑模式
    '参数：
    '返回：
    '------------------------------------------------------------------------------------------------------------------
    If Val(mrsParam("编辑")) = 0 Then

        mshDownTab.Enabled = False
        mshUpTab.Enabled = False
    Else

        mshDownTab.Enabled = True
        mshUpTab.Enabled = True
    End If
    
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '1、保存手术日期设置
    '2、保存体温图线数据
    '3、保存体温表格数据
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long, i As Long
    Dim aryValue() As String, aryData() As String
    Dim aryMakeTime() As String
    Dim strFrom As String, strTo As String
    Dim strItem As String, strTime As String
    Dim strMakeTime As String
    Dim mvarStrValue As String, dblValues As Double
    Dim lngItemCode As Long, intMode As Integer
    Dim rs As New ADODB.Recordset
    Dim intLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim strValues As String
    Dim intTmp As Integer, intMax As Integer
    Dim lng病历文件id As Long
    Dim lng病历内容id As Long
    Dim strSQL() As String
    Dim strTmp As String
    Dim str删除心率短拙 As String           '保证一列只删除一次
    Dim blnHistoryData As Boolean           '历史数据保存的有时间,以历史时间为准更新数据
    Dim blnTrans As Boolean
    
    If Val(mrsParam("编辑")) = 0 Then Exit Function
        
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    Screen.MousePointer = 11
    
    ReDim Preserve strSQL(1 To 1)
    
    On Error GoTo ErrHead

    '1.保存手术日期设置
    With mshUpTab
        'mshDownTab是按1为标志进行判断，mshUpTab是按"删除手术日"和"填写手术日"是否进行了编辑
        If .Tag = "填写手术日" Or .Tag = "删除手术日" Then
            For intCol = .FixedCols To .Cols - 1
                
                For intLoop = (intCol - 1) * 6 + mshScale.FixedCols To intCol * 6 + mshScale.FixedCols
                    
                    strTmp = GetCurveDateTime(intLoop - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
                    strStart = Split(strTmp, ",")(0)
                    strEnd = Split(strTmp, ",")(1)

                    If Int(CDate(strStart)) < Int(CDate(strFrom)) Then
                        strStart = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
                    End If
                    
                    mstrSQL = "ZL_电子护理记录_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("婴儿").Value) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "4,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "NULL"
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL
                Next

                strStart = mstrOpsDays(intCol)
                strEnd = mstrOpsDays(intCol)
                If strStart <> "" Then
                
                    strTmp = ""
                    
                    intTmp = GetCurveColumn(CDate(strStart), CDate(strFrom), mlngHourBegin) + mshScale.FixedCols - 1
                    If Left(mshScale.TextMatrix(3, intTmp), 4) = "手术分娩" Then
                        strTmp = "手术分娩"
                    ElseIf Left(mshScale.TextMatrix(3, intTmp), 2) = "手术" Then
                        strTmp = "手术"
                    ElseIf Left(mshScale.TextMatrix(3, intTmp), 2) = "分娩" Then
                        strTmp = "分娩"
                    End If
                    
                    mstrSQL = "ZL_电子护理记录_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("婴儿").Value) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "4,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "'" & strTmp & "'"
                    mstrSQL = mstrSQL & ")"

                    strSQL(ReDimArray(strSQL)) = mstrSQL
                End If

            Next
        End If
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '2.保存体温图线数据
    With mshScale
    
        '注释说明保存
        For intCol = .FixedCols To .Cols - 1
            strTmp = GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
            strTime = Split(strTmp, ",")(0)
            strEnd = Split(strTmp, ",")(1)
            
            If Int(CDate(strTime)) < Int(CDate(strFrom)) Then strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")

            aryValue = Split(.TextMatrix(GraphDataRow.更改标志, intCol), ";")
                            
                            

            If Mid(Format(strTime, "yyyy-MM-dd HH:mm"), 12, 5) = "00:00" Then
                strMakeTime = Format(DateAdd("h", -2, CDate(strEnd)), "yyyy-MM-dd HH:mm")
                strMakeTime = Format(DateAdd("n", 1, CDate(strMakeTime)), "yyyy-MM-dd HH:mm:ss")
            Else
                strMakeTime = Format(DateAdd("h", 2, CDate(strTime)), "yyyy-MM-dd HH:mm:ss")
            End If
            strMakeTime = "To_Date('" & strMakeTime & "','yyyy-mm-dd hh24:mi:ss')"
                
            mstrSQL = "ZL_电子护理记录_UPDATE("
            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "2,"
            mstrSQL = mstrSQL & "0,"
            mstrSQL = mstrSQL & Val(.TextMatrix(GraphDataRow.断开标志, intCol)) & ","
            mstrSQL = mstrSQL & IIf(Val(aryValue(0)) = 4, "NULL", "'" & .TextMatrix(GraphDataRow.上标说明, intCol) & "'")
            mstrSQL = mstrSQL & ",Null,1,1,0,0," & strMakeTime & ",Null"
            mstrSQL = mstrSQL & ")"
            strSQL(ReDimArray(strSQL)) = mstrSQL

            mstrSQL = "ZL_电子护理记录_UPDATE("
            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
            mstrSQL = mstrSQL & "6,"
            mstrSQL = mstrSQL & "0,"
            mstrSQL = mstrSQL & Val(.TextMatrix(GraphDataRow.断开标志, intCol)) & ","
            mstrSQL = mstrSQL & IIf(Val(aryValue(0)) = 4, "NULL", "'" & .TextMatrix(GraphDataRow.下标说明, intCol) & "'")
            mstrSQL = mstrSQL & ",Null,1,1,0,0," & strMakeTime & ",Null"
            mstrSQL = mstrSQL & ")"
            
            strSQL(ReDimArray(strSQL)) = mstrSQL
            
        Next
        
        '项目数值保存
        For intCount = 0 To .FixedCols - 1
            Dim dbl分钟差 As Double                                         '保存当前格的分钟差
            '提取指定项目定义：最大值；最小值；单位值；最高行
            strItem = .TextMatrix(0, intCount)              '提取名称
            aryValue = Split(picLine(intCount).Tag, ";")    '提取项目属性列
            lngItemCode = mshScale.ColData(intCount)        '提取项目序号
            
            '从数据列中读出数据
            For intCol = .FixedCols To .Cols - 1
                
                '得到时间,如果没有时间的,以计算出来的中间时间为准(新数据);否则以历史时间为准进行数据更新
                strTime = Split(mshScale.TextMatrix(GraphDataRow.曲线时间, intCol), ";")(intCount + 1)
                If strTime = "" Then
                    blnHistoryData = False
                    strTmp = GetCurveDateTime(intCol - mshScale.FixedCols + 1, CDate(strFrom), mlngHourBegin)
                    strTime = Split(strTmp, ",")(0)
                    strEnd = Split(strTmp, ",")(1)
                    If Int(CDate(strTime)) < Int(CDate(strFrom)) Then strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
                Else
                    strEnd = strTime
                    strMakeTime = strTime
                    blnHistoryData = True
                End If
                
                intMode = Val(Split(.TextMatrix(GraphDataRow.更改标志, intCol), ";")(intCount + 1))
                
                If intMode = OperateType.删除操作 Or intMode = OperateType.修改操作 Or intMode = OperateType.新增操作 Then
                    '取每格的中间时间点
                    If blnHistoryData = False Then
                        dbl分钟差 = DateDiff("n", CDate(Split(strTmp, ",")(0)), CDate(Split(strTmp, ",")(1)))
                        dbl分钟差 = dbl分钟差 \ 2
                        strMakeTime = DateAdd("n", dbl分钟差, CDate(Split(strTmp, ",")(0)))
                        If strMakeTime < mstrEnterDate Then strMakeTime = mstrEnterDate
                    End If
                    strMakeTime = "To_Date('" & strMakeTime & "','yyyy-mm-dd hh24:mi:ss')"
                End If
                
                If intMode = OperateType.删除操作 Or intMode = OperateType.修改操作 Then

                    '删除物理降温的数据
                    mstrSQL = "ZL_电子护理记录_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL

                    '删除脉搏短绌的数据
                    If mint心率应用 = 2 And InStr(1, str删除心率短拙 & ",", "," & intCol & ",") = 0 Then
                        mstrSQL = "ZL_电子护理记录_UPDATE("
                        mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                        mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & mItemNo.心率 & ","
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                        mstrSQL = mstrSQL & ")"
                        strSQL(ReDimArray(strSQL)) = mstrSQL
                        str删除心率短拙 = str删除心率短拙 & "," & intCol
                    End If
                    
                    mstrSQL = "ZL_电子护理记录_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "0,"
                    mstrSQL = mstrSQL & "Null,"
                    mstrSQL = mstrSQL & IIf(lngItemCode = mItemNo.体温 Or lngItemCode = mItemNo.呼吸 Or lngItemCode = mItemNo.脉搏, "'" & Split(.TextMatrix(GraphDataRow.部位标志, intCol), ";")(intCount + 1) & "'", "''")
                    mstrSQL = mstrSQL & ",1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.未记说明, intCol), ";")(intCount + 1)) & "')"
                    
                    strSQL(ReDimArray(strSQL)) = mstrSQL
                End If
                
                If intMode = OperateType.新增操作 Or intMode = OperateType.修改操作 Then
                    
                    '删除物理降温的数据
                    mstrSQL = "ZL_电子护理记录_UPDATE("
                    mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                    mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                    mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & lngItemCode & ","
                    mstrSQL = mstrSQL & "1,"
                    mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                    mstrSQL = mstrSQL & ")"
                    strSQL(ReDimArray(strSQL)) = mstrSQL

                    '删除脉搏短绌的数据
                    If mint心率应用 = 2 And InStr(1, str删除心率短拙 & ",", "," & intCol & ",") = 0 Then
                        mstrSQL = "ZL_电子护理记录_UPDATE("
                        mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                        mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                        mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & mItemNo.心率 & ","
                        mstrSQL = mstrSQL & "1,"
                        mstrSQL = mstrSQL & "Null,Null,1,1,0,0," & strMakeTime
                        mstrSQL = mstrSQL & ")"
                        strSQL(ReDimArray(strSQL)) = mstrSQL
                        str删除心率短拙 = str删除心率短拙 & "," & intCol
                    End If
                    
                    strValues = ""
                    aryData = Split(Split(.TextMatrix(GraphDataRow.曲线数据, intCol), ";")(intCount + 1), ",")
                    If UBound(aryData) = -1 Then
                    
                            mstrSQL = "ZL_电子护理记录_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            
                            Select Case lngItemCode
                            Case mItemNo.脉搏
                                If mint心率应用 = 2 Then
                                    mstrSQL = mstrSQL & IIf(strValues = "", mItemNo.脉搏, mItemNo.心率) & ","
                                Else
                                    mstrSQL = mstrSQL & mItemNo.脉搏 & ","
                                End If
                                                            
                            Case Else
                                mstrSQL = mstrSQL & lngItemCode & ","
                            End Select
                            
                            If lngItemCode = mItemNo.心率 Then
                                mstrSQL = mstrSQL & "1,"
                            Else
                                mstrSQL = mstrSQL & IIf(strValues = "", "0", "1") & ","
                            End If
    
                            mstrSQL = mstrSQL & "'',"
                            
                            mstrSQL = mstrSQL & "'',1,1,0"
                                                        
                            mstrSQL = mstrSQL & ",0," & strMakeTime
                            mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.未记说明, intCol), ";")(intCount + 1)) & "')"
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                    Else
                        For i = 0 To UBound(aryData)
                            
                            dblValues = ConvertToValue(intCount, aryData(i))
                            If intCount = mItemSerial.体温 Then
                                dblValues = Format(dblValues, "0.00")
                            Else
                                dblValues = Format(dblValues, "0")
                            End If
                            
                            mstrSQL = "ZL_电子护理记录_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            
                            Select Case lngItemCode
                            Case mItemNo.脉搏
                                If mint心率应用 = 2 Then
                                    mstrSQL = mstrSQL & IIf(strValues = "", mItemNo.脉搏, mItemNo.心率) & ","
                                Else
                                    mstrSQL = mstrSQL & mItemNo.脉搏 & ","
                                End If
                                                            
                            Case Else
                                mstrSQL = mstrSQL & lngItemCode & ","
                            End Select
                            
                            If lngItemCode = mItemNo.心率 Then
                                mstrSQL = mstrSQL & "1,"
                            Else
                                mstrSQL = mstrSQL & IIf(strValues = "", "0", "1") & ","
                            End If
                            
                            '如果是体温项目,其数值为零且标记说明为"不升",需将值保存为"不升",标记说明保存为空
                            If CStr(Val(dblValues)) = "0" And lngItemCode = mItemNo.体温 And Trim(Split(.TextMatrix(GraphDataRow.未记说明, intCol), ";")(intCount + 1)) = "不升" Then
                                mstrSQL = mstrSQL & "'不升',"
                            Else
                                mstrSQL = mstrSQL & "'" & dblValues & "',"
                            End If
                            
                            mstrSQL = mstrSQL & IIf((lngItemCode = mItemNo.体温 Or lngItemCode = mItemNo.呼吸 Or lngItemCode = mItemNo.脉搏) And strValues = "", "'" & Split(.TextMatrix(GraphDataRow.部位标志, intCol), ";")(intCount + 1) & "'", "''") & ",1,1,"
                            mstrSQL = mstrSQL & IIf(lngItemCode = mItemNo.体温 And i = 0, Val(.TextMatrix(GraphDataRow.复试标志, intCol)), "0")
                            
                            mstrSQL = mstrSQL & ",0," & strMakeTime
                            '如果是体温项目,其数值为零且标记说明为"不升",需将值保存为"不升",标记说明保存为空
                            If CStr(Val(dblValues)) = "0" And lngItemCode = mItemNo.体温 And Trim(Split(.TextMatrix(GraphDataRow.未记说明, intCol), ";")(intCount + 1)) = "不升" Then
                                mstrSQL = mstrSQL & ",'')"
                            Else
                                mstrSQL = mstrSQL & ",'" & Trim(Split(.TextMatrix(GraphDataRow.未记说明, intCol), ";")(intCount + 1)) & "')"
                            End If
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                            If strValues = "" Then
                                strValues = dblValues
                            Else
                                strValues = strValues & "," & dblValues
                            End If
                        Next
                    End If
                End If
            Next
        Next
    End With
    '------------------------------------------------------------------------------------------------------------------
    '3.保存呼吸表格数据
    With vsf
        For intCol = 2 To .Cols - 1
            
            strTmp = GetCurveDateTime(intCol - 2 + 1, CDate(strFrom), mlngHourBegin)
            strTime = Split(strTmp, ",")(0)
            strEnd = Split(strTmp, ",")(1)
            
            If Int(CDate(strTime)) < Int(CDate(strFrom)) Then
                strTime = Format(strFrom, "yyyy-MM-dd HH:mm:ss")
            End If
            
            intMode = Val(vsf.ColData(intCol))
            mvarStrValue = Trim(vsf.TextMatrix(1, intCol))
            
            If intMode = OperateType.删除操作 Or intMode = OperateType.修改操作 Then
            
                mstrSQL = "ZL_电子护理记录_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & mItemNo.呼吸 & ","
                mstrSQL = mstrSQL & "0,"
                mstrSQL = mstrSQL & "NULL"
                mstrSQL = mstrSQL & ")"
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
                
            End If
            
            If intMode = OperateType.新增操作 Or intMode = OperateType.修改操作 And mvarStrValue <> "" Then
                mstrSQL = "ZL_电子护理记录_UPDATE("
                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                mstrSQL = mstrSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                mstrSQL = mstrSQL & "1,"
                mstrSQL = mstrSQL & mItemNo.呼吸 & ","
                mstrSQL = mstrSQL & "0,"
                mstrSQL = mstrSQL & "'" & mvarStrValue & "'"
                mstrSQL = mstrSQL & ")"
                
                strSQL(ReDimArray(strSQL)) = mstrSQL
            End If
            
        Next
    End With

    '------------------------------------------------------------------------------------------------------------------
    '3.保存体温表格数据
    With mshDownTab
        If .Tag = "1" And .RowData(1) > 0 Then
            For intCol = .FixedCols To .Cols - 1 Step 2
                
                '求出时间
                
                strTmp = GetEditDateTime(intCol - .FixedCols + 1, CDate(strFrom))
                
                strTime = Split(strTmp, ",")(0)
                strEnd = Split(strTmp, ",")(1)
                
                '求出属性列表
                
                For intCount = 0 To .Rows - 2
                    '求出名称
                    strItem = .TextMatrix(intCount + 1, 1)
                    
                    '操作数据
                    
                    For intLoop = intCol To intCol + 1
                        
                        aryValue = Split(.TextMatrix(0, intLoop), ";")
                        intMode = Val(aryValue(intCount))
                        
                        mvarStrValue = .TextMatrix(intCount + 1, intLoop)
                                                
                        If intLoop = intCol + 1 Then
                            strEnd = Format(DateAdd("d", 1, CDate(Left(strTime, 10))), "yyyy-MM-dd") & " 00:00:00"
                            strTime = Left(strTime, 10) & " 12:00:00"
                        Else
                            strTime = Left(strTime, 10) & " 00:00:00"
                            strEnd = Left(strTime, 10) & " 12:00:00"
                        End If
                        
                        strStart = Format(CDate(strTime) - (4 - 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        If Int(CDate(strStart)) <> Int(CDate(strFrom)) Then strStart = Format(strTime, "yyyy-MM-dd HH:mm:ss")
                        If strStart < mstr最小时间 Then strStart = mstr最小时间
                        strEnd = Format(CDate(strEnd) - (4 - 4) / 24, "YYYY-MM-DD hh:mm:ss")
                        
                        If intMode = 4 Or intMode = 3 Then
                            
                            mstrSQL = "ZL_电子护理记录_UPDATE("
                            mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                            mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                            mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                            mstrSQL = mstrSQL & "1,"
                            mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                            mstrSQL = mstrSQL & "0,"
                            mstrSQL = mstrSQL & "NULL"
                            mstrSQL = mstrSQL & ")"
                            
                            strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                            If Val(.RowData(intCount + 1)) = mItemNo.血压 Then
                                mstrSQL = "ZL_电子护理记录_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & mItemNo.舒张压 & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "NULL"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            End If
                            
                        End If
                        If (intMode = 2 Or intMode = 3) And mvarStrValue <> "" Then
                            
                            
                            If Val(.RowData(intCount + 1)) = mItemNo.血压 Then
                            
                                mstrSQL = "ZL_电子护理记录_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & Split(mvarStrValue, "/")(0) & "'"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            
                                mstrSQL = "ZL_电子护理记录_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & mItemNo.舒张压 & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & Split(mvarStrValue, "/")(1) & "'"
                                mstrSQL = mstrSQL & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            Else
                                mstrSQL = "ZL_电子护理记录_UPDATE("
                                mstrSQL = mstrSQL & Val(mrsParam("病人id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("主页id")) & ","
                                mstrSQL = mstrSQL & Val(mrsParam("婴儿")) & ","
                                mstrSQL = mstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss'),"
                                mstrSQL = mstrSQL & "1,"
                                mstrSQL = mstrSQL & Val(.RowData(intCount + 1)) & ","
                                mstrSQL = mstrSQL & "0,"
                                mstrSQL = mstrSQL & "'" & mvarStrValue & "'"
                                mstrSQL = mstrSQL & ",Null,1,1,0," & IIf(IsNumeric(mvarStrValue), 0, 1) & ")"
                                
                                strSQL(ReDimArray(strSQL)) = mstrSQL
                            End If
                        End If
                                            
                    Next
                Next
            Next
        End If
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '循环执行SQL保存数据
    gcnOracle.BeginTrans
    blnTrans = True
    intMax = UBound(strSQL)
    For intTmp = 1 To intMax
        If strSQL(intTmp) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(intTmp), "保存体温数据")
    Next
    gcnOracle.CommitTrans
    blnTrans = False
    SaveData = True
    
    Screen.MousePointer = 0
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    
    If blnTrans Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Private Sub cboBaby_Click()
    
    If opt(1).Value = False Then Exit Sub
    
    If Val(mrsParam("婴儿").Value) = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    mrsParam("婴儿").Value = cboBaby.ItemData(cboBaby.ListIndex)
    
    If InitBody(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("病区id").Value), Val(mrsParam("婴儿").Value)) = False Then Exit Sub

    Call zlMenuClick("显示病人姓名")
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
    '--------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        mcbrToolBar页面.Caption = Control.Caption
        Call zlMenuClick("装载数据", Control.Parameter)
        cbsMain.RecalcLayout
        
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long
    Dim lngTop  As Long
    Dim lngRight  As Long
    Dim lngBottom  As Long

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next

    '窗体其它控件Resize处理
    picPane.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    picPane.BackColor = pic.BackColor
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.Id
    Case conMenu_View_Option
        Control.Visible = mblnBabys
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_View_Jump
        
        If Control.Parameter = "" Then
            Control.Checked = True
        Else
            Control.Checked = (Val(Split(Control.Parameter, ";")(4)) = Page)
        End If
        
        
    End Select
    
End Sub

'######################################################################################################################
'事件

Private Sub hsb_Change()
    
    On Error Resume Next
    
'    pic.Left = 60 - hsb.Value * 300
    pic.Left = -60 - hsb.Value * msinHStep
End Sub

Private Sub mfrmCaseTendBodyPrint_AfterPrint()
    RaiseEvent zlAfterPrint
End Sub

Private Sub mshDownTab_DblClick()
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub

    txtInput(1).Text = ""
    Call ShowInput

End Sub

Private Function ShowInput() As Boolean
    Dim strTmp As String
    
    With mshDownTab
        
        picInput.Move .Cell(flexcpLeft, .Row, .Col), .Cell(flexcpTop, .Row, .Col), .Cell(flexcpWidth, .Row, .Col) - 15, .Cell(flexcpHeight, .Row, .Col) - 15
        picInput.Visible = True
        picInput.BackColor = .Cell(flexcpBackColor, .Row, .Col)
        picInput.Tag = .Row & ";" & .Col & ";" & ""
        txtInput(0).BackColor = picInput.BackColor
        
        If mItemNo.血压 = Val(.RowData(.Row)) Then
            txtInput(1).BackColor = picInput.BackColor
            lblInput.Caption = "/"
            
            txtInput(0).Move 0, 0, (picInput.Width - lblInput.Width) / 2, picInput.Height
            lblInput.Left = txtInput(0).Left + txtInput(0).Width
            txtInput(1).Move lblInput.Left + lblInput.Width, 0, (picInput.Width - lblInput.Width) / 2, picInput.Height
            
            strTmp = .TextMatrix(.Row, .Col)
            If InStr(strTmp, "/") > 0 Then
                
                txtInput(0).Text = Left(strTmp, InStr(strTmp, "/") - 1)
                txtInput(1).Text = Mid(strTmp, InStr(strTmp, "/") + 1)
                
            Else
                txtInput(0).Text = strTmp
            End If
            txtInput(0).Alignment = 2
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 3
            txtInput(1).Visible = True
            txtInput(1).SelStart = 0
            txtInput(1).SelLength = 3
        Else
            txtInput(1).Visible = False
            
            lblInput.Caption = ""
            txtInput(0).Alignment = 1
            txtInput(0).Move 0, 0, picInput.Width, picInput.Height
            txtInput(0).Text = .TextMatrix(.Row, .Col)
            txtInput(0).MaxLength = IIf(mItemStru(.Row).数据长度 < 12, 12, mItemStru(.Row).数据长度)
            lblInput.Move txtInput(0).Left + txtInput(0).Width, txtInput(0).Top
        
            txtInput(0).SelStart = 0
            txtInput(0).SelLength = 100
        End If

        txtInput(0).SetFocus
    End With
    
End Function

Private Sub mshDownTab_KeyPress(KeyAscii As Integer)
    If Val(mrsParam("编辑")) = 0 Then Exit Sub

    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If mshDownTab.RowData(mshDownTab.Row) <= 0 Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
    
    Select Case KeyAscii
    Case 13                 'Enter移动单元格
        With mshDownTab
            If .Row = .Rows - 1 Then
                If .Col < .Cols - 1 Then
                    .Col = .Col + 1
                End If
                .Row = .FixedRows
            Else
                If .RowHidden(.Row + 1) Then
                    .Row = .Row + 2
                Else
                    .Row = .Row + 1
                End If
            End If
        End With
        Call mshDownTab_RowColChange
    Case 32                 '空格键进入编辑
        Call mshDownTab_DblClick
    Case vbKeyDelete
        Call zlMenuClick("删除项目")
    Case Else

        Select Case mItemStru(mshDownTab.Row).数据类型
        Case 0 '数值型
            
            Select Case Val(mshDownTab.RowData(mshDownTab.Row))
            Case mItemNo.大便
                If Check是否包含(UCase(Chr(KeyAscii)), "0123456789+/E*") = False Then KeyAscii = 0
            Case mItemNo.出液
                If Check是否包含(UCase(Chr(KeyAscii)), "0123456789/C") = False Then KeyAscii = 0
            Case Else
                If Check是否包含(UCase(Chr(KeyAscii)), "正小数") = True Then KeyAscii = 0
            End Select

        Case 1 '字符型
            If Check是否包含(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
        End Select
        
        If KeyAscii <> 0 Then
            Call mshDownTab_DblClick
            txtInput(0).Text = Chr(KeyAscii)
            txtInput(0).SelStart = Len(txtInput(0).Text)
        End If
    End Select
End Sub

Private Sub mshDownTab_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim aryValue() As String
    Dim intRewrite As Integer
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
    
    If KeyCode = 46 Then        'Delete清除单元
        With mshDownTab
            '清除单元数值
            If .TextMatrix(.Row, .Col) = "" Then Exit Sub
            .TextMatrix(.Row, .Col) = ""
            
            '处理删除标记
            aryValue() = Split(.TextMatrix(0, .Col), ";")
            intRewrite = Val(aryValue(.Row - 1))
            If ((.Col + 1) - mshDownTab.FixedCols) / 2 = ((.Col + 1) - mshDownTab.FixedCols) \ 2 Then
                '如果为下午
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 0
                Case 1
                    aryValue(.Row - 1) = 4
                Case 2
                    aryValue(.Row - 1) = 0
                Case 3
                    aryValue(.Row - 1) = 4
                Case 4
                    aryValue(.Row - 1) = 4
                End Select
            Else
                '否则为上午
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 2
                Case 1
                    aryValue(.Row - 1) = 3
                Case 2
                    aryValue(.Row - 1) = 2
                Case 3
                    aryValue(.Row - 1) = 3
                Case 4
                    aryValue(.Row - 1) = 3
                End Select
            End If
            .TextMatrix(0, .Col) = Join(aryValue, ";")
            '如果为下午就检查更新上午的操作标志
            If ((.Col + 1) - mshDownTab.FixedCols) / 2 = ((.Col + 1) - mshDownTab.FixedCols) \ 2 Then
                aryValue() = Split(.TextMatrix(0, .Col - 1), ";")
                intRewrite = Val(aryValue(.Row - 1))
                '增加或修改操作
                Select Case intRewrite
                Case 0
                    aryValue(.Row - 1) = 2
                Case 1
                    aryValue(.Row - 1) = 3
                Case 2
                    aryValue(.Row - 1) = 2
                Case 3
                    aryValue(.Row - 1) = 3
                Case 4
                    aryValue(.Row - 1) = 3
                End Select
                .TextMatrix(0, .Col - 1) = Join(aryValue, ";")
            End If
        End With
        picInput.Visible = False
        
    End If
End Sub

Private Sub mshDownTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    
    If mshDownTab.Tag = "" Or mvarEdit = False Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then Exit Sub
End Sub

Private Sub mshDownTab_RowColChange()
    
    Dim strFrom As String, strTo As String
    Dim intNowRow As Integer, intNowCol As Integer
    Dim strInfo As String
    On Error GoTo ErrHead
    
    If mshDownTab.Tag = "" Or mvarEdit = False Or picScale.Tag = "" Then Exit Sub
    If CheckTimeRange(mshDownTab.Col) = False Then
        mshDownTab.FocusRect = flexFocusLight
        Exit Sub
    Else
        mshDownTab.FocusRect = flexFocusSolid
    End If

    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    With mshDownTab

        For intRow = .FixedRows To .Rows - 1
            For intCol = .FixedCols To .Cols - 1
                If intCol / 2 <> intCol \ 2 Then
                    '双数列时为蓝色
                    .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &HF7ECE6
                Else
                    '单数列为白色
                    .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &H80000005
                End If
            Next
        Next
    End With
    
    If (Val(mshDownTab.TextMatrix(mshDownTab.Row, 2)) <> 0 Or Val(mshDownTab.TextMatrix(mshDownTab.Row, 3)) <> 0) And mItemStru(mshDownTab.Row).数据类型 = 0 Then
        strInfo = "“" & mshDownTab.TextMatrix(mshDownTab.Row, 1) & "”项目范围：" & Val(mshDownTab.TextMatrix(mshDownTab.Row, 3)) & "～" & Val(mshDownTab.TextMatrix(mshDownTab.Row, 2)) & " " & strInfo
    End If
    
    RaiseEvent PromptInfo(strInfo)
    
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshScale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    If Not mvarEdit Then Exit Sub
    
    With mshScale
        intCol = (X - .Left) \ .ColWidth(0)
        If Button = 1 Then
            If intCol >= .FixedCols Then Exit Sub
            
            .Cell(flexcpBackColor, 0, 0, .Rows - 1, .FixedCols - 1) = RGB(255, 255, 255)
            
            If picGraph.Tag = CStr(intCol) Then
                Call ClearLineSelect
                mlngLine = 0
                RaiseEvent PromptInfo("")
            Else
                RaiseEvent PromptInfo("")
                mlngLine = intCol + 1

                .Cell(flexcpBackColor, 0, intCol, .Rows - 1, intCol) = RGB(0, 255, 255)
                
                picGraph.Tag = intCol
                
                picGraph.MousePointer = 2

                linHCur.BorderColor = .Cell(flexcpForeColor, 0, intCol)
                linVCur.BorderColor = linHCur.BorderColor
                
                linHCur.X1 = 0: linHCur.X2 = 0: linHCur.Y1 = 0: linHCur.Y2 = 0
                linHCur.Visible = True
                
                linVCur.X1 = 0: linVCur.X2 = 0: linVCur.Y1 = 0: linVCur.Y2 = 0
                linVCur.Visible = True
            End If
            RaiseEvent SelectScale(intCol)
        ElseIf mItemSerial.体温 = intCol Or mItemSerial.呼吸 = intCol Or mItemSerial.脉搏 = intCol Then

            RaiseEvent RButton(Button, Shift, X, Y)
            
        End If
    End With
End Sub

Private Sub lblCur_Click()
    picScale.SetFocus
End Sub

Private Sub lblCur_DblClick()
    picScale_KeyDown 13, 0
End Sub

Private Sub mshUpTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strDay As String
    Dim strFrom As String
    Dim strTo As String
    
    mshUpTab.FocusRect = flexFocusLight
    If picScale.Tag = "" Then Exit Sub
    If InStr(picScale.Tag, ";") = 0 Then Exit Sub
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    strDay = Format(Int(CDate(strFrom) + NewCol - mshUpTab.FixedCols), "yyyy-MM-dd")
    
    If strDay >= Format(strFrom, "yyyy-MM-dd") And strDay <= Format(strTo, "yyyy-MM-dd") Then
        mshUpTab.FocusRect = flexFocusSolid
    Else
        mshUpTab.FocusRect = flexFocusLight
    End If
    
End Sub

Private Sub mshUpTab_DblClick()
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
        
    Call zlMenuClick("填写手术日")
End Sub

Private Sub mshUpTab_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCol As Integer
    Dim strTime As String
    Dim intStart As Integer
    Dim intEnd As Integer
    Dim intLoop As Integer
    Dim strCaption As String
    
    On Error GoTo ErrHead
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyReturn ' 13     '设置手术日
        With mshUpTab
        
            If Trim(.TextMatrix(2, .Col)) = "0" Then Exit Sub       '已经是手术日期，直接退出

            strTime = mstrOpsDays(.Col)
            If strTime = "" Then

                strTime = GetCurveDateTime((.Col - .FixedCols) * 6 + mshScale.FixedCols - 1, CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin)
                strTime = Split(strTime, ",")(0)
                
            End If
            
            intCol = GetCurveColumn(CDate(strTime), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
             
            strCaption = mshScale.TextMatrix(3, intCol)
            If frmInputDate.ShowMe(strTime, Split(picScale.Tag, ";")(0), Split(picScale.Tag, ";")(1), strCaption) Then
                
                mshUpTab.Tag = "填写手术日"
                
                intCol = GetCurveColumn(CDate(strTime), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                
                .Col = Int((intCol - mshScale.FixedCols) / 6) + 1

                If Trim(.TextMatrix(2, .Col)) <> "" And Trim(.TextMatrix(2, .Col)) <> "0" Then
                    If MsgBox("病人" & .TextMatrix(2, .Col) & "天前曾经手术/分娩，是否再次手术/分娩？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                mstrOpsDays(.Col) = strTime
                
                '先清除当前日期内的所有手术文字显示内容
                intStart = GetCurveColumn(CDate(Format(strTime, "yyyy-MM-dd") & " 00:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
                intEnd = GetCurveColumn(CDate(Format(strTime, "yyyy-MM-dd") & " 23:00:00"), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1

                For intLoop = intStart To intEnd
                    If Left(mshScale.TextMatrix(3, intLoop), 2) = "手术" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                    If Left(mshScale.TextMatrix(3, intLoop), 2) = "分娩" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                  If Left(mshScale.TextMatrix(3, intLoop), 2) = "手术分娩" Then
                        mshScale.TextMatrix(3, intLoop) = ""
                    End If
                Next
                                                         
                Select Case strCaption
                Case "分娩"
                    If mBodyFlag.分娩 = 2 Then
                        mshScale.TextMatrix(3, intCol) = strCaption & "--" & ConvertTimeToChinese(Format(strTime, "HH:mm"))
                    Else
                        mshScale.TextMatrix(3, intCol) = strCaption
                    End If
                Case Else
                    If mBodyFlag.手术 = 2 Then
                        mshScale.TextMatrix(3, intCol) = strCaption & "--" & ConvertTimeToChinese(Format(strTime, "HH:mm"))
                    Else
                        mshScale.TextMatrix(3, intCol) = strCaption
                    End If
                End Select
                
                mshScale.Cell(flexcpData, 3, intCol, 3, intCol) = Format(strTime, "HH:mm:ss")
                
                Select Case .ColData(.Col)
                Case 0      '非手术日，填写为新手术日
                    .ColData(.Col) = 2
                Case 1      '原来就是手术日
                Case 2      '已经设置为新手术日
                Case 3      '被删除的的手术日，再次设置为手术日
                    .ColData(.Col) = 1
                End Select
                
                Call ShowOpsDays
                Call DrawPaper
                Call DrawGraph
            
            End If

        End With
        
    Case vbKeyDelete ' 46     '清除手术日
        With mshUpTab
'            If Trim(.TextMatrix(2, .Col)) <> "0" Then Exit Sub          '当前并非手术日
            If .ColData(.Col) <> 1 And .ColData(.Col) <> 2 Then Exit Sub
            
            If MsgBox("是否清除病人" & .TextMatrix(0, .Col) & "日的手术登记？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            mshUpTab.Tag = "删除手术日"
            Select Case .ColData(.Col)
            Case 0      '非手术日
            Case 1      '原来就是手术日，设置为删除手术日
                .ColData(.Col) = 3
            Case 2      '新手术日，再次设置为非手术日
                .ColData(.Col) = 0
            Case 3      '被删除的的手术日
            End Select
            
            intCol = GetCurveColumn(CDate(mstrOpsDays(.Col)), CDate(Split(picScale.Tag, ";")(0)), mlngHourBegin) + mshScale.FixedCols - 1
            
            If intCol > 0 Then mshScale.TextMatrix(3, intCol) = ""
            mstrOpsDays(.Col) = ""
            
            Call ShowOpsDays
            Call DrawPaper
            Call DrawGraph

        End With
        
    End Select
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub opt_Click(Index As Integer)
    
    cboBaby.Enabled = (opt(1).Value = True)
    
    Select Case Index
    Case 0                  '病人本人
        
        If Val(mrsParam("婴儿").Value) = 0 Then Exit Sub
        mrsParam("婴儿").Value = 0
        
        If InitBody(Val(mrsParam("病人id").Value), Val(mrsParam("主页id").Value), Val(mrsParam("病区id").Value), Val(mrsParam("婴儿").Value)) = False Then Exit Sub
        
        Call zlMenuClick("显示病人姓名")
        
    Case 1                  '婴儿
        
        Call cboBaby_Click
        
    End Select
        
End Sub

Private Sub picCard_Paint(Index As Integer)
    Dim intLoop As Integer
    
    On Error Resume Next
    
    picCard(Index).Cls
    For intLoop = 0 To txtCard.UBound
        txtCard(intLoop).Height = 180
        If txtCard(intLoop).Visible Then
            DrawLine picCard(Index), txtCard(intLoop).Left, txtCard(intLoop).Top + txtCard(intLoop).Height + 15, txtCard(intLoop).Left + txtCard(intLoop).Width, txtCard(intLoop).Top + txtCard(intLoop).Height + 15, &H8000000C
        End If
    Next
End Sub

Private Sub picCard_Resize(Index As Integer)
    On Error Resume Next
    
    txtCard(1).Move txtCard(1).Left, txtCard(1).Top, picCard(Index).Width - txtCard(1).Left - 45
    txtCard(7).Move txtCard(7).Left, txtCard(7).Top, picCard(Index).Width - txtCard(7).Left - 45
    
End Sub

Private Sub picGraph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo ErrHead
    '-------------------------------------------------
    '1、标尺线随鼠标移动
    '2、当鼠标移动超过横纵向数据限制，鼠标由十字变成不可移动，标尺不在移动
    '-------------------------------------------------
    Dim intMinCol As Long
    Dim intMaxCol As Long
    Dim sglLeft As Single
    Dim sglRight As Single
    
    Dim aryValue() As String
    Dim aryNote() As String
    
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    If picGraph.Tag = "" Or picScale.Tag = "" Then Exit Sub
    
    Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)

    If picGraph.MousePointer <> 2 Then picGraph.MousePointer = 2
    
    '获得最小最大时间范围
    aryValue = Split(picScale.Tag, ";")
    
    sglLeft = intMinCol * HOUR_STEP_Twips + 30
    sglRight = (intMaxCol + 1) * HOUR_STEP_Twips - 30
    
    If X < sglLeft Then
        X = sglLeft
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    If X > sglRight Then
        X = sglRight
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
        
    
    '获取项目定义:最大值；最小值；单位值；最高行
    aryValue = Split(picLine(Val(picGraph.Tag)).Tag, ";")
    If Y < (aryValue(3) - 1) * mshScale.ROWHEIGHT(1) Then
        Y = (aryValue(3) - 1) * mshScale.ROWHEIGHT(1)
        
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
        
    End If
    If Y > (aryValue(3) - 1 + (aryValue(0) - aryValue(1)) / aryValue(2)) * mshScale.ROWHEIGHT(1) Then
        Y = (aryValue(3) - 1 + (aryValue(0) - aryValue(1)) / aryValue(2)) * mshScale.ROWHEIGHT(1)
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    With linHCur
        .X1 = 0: .X2 = X:
        .Y1 = Y: .Y2 = Y
    End With
    With linVCur
        .X1 = X: .X2 = X:
        .Y1 = 0: .Y2 = Y
    End With
    
    '状态提示坐标显示处理
    '------------------------------------------------------------------------------------------------------------------
    Dim intNowCol As Integer
    Dim dblValues As Single
    Dim strTmp As String
    
    intNowCol = (linVCur.X1 \ HOUR_STEP_Twips) + 1
    
    '如果当前列是断开的，则不允许作图
    If Val(mshScale.TextMatrix(GraphDataRow.断开标志, intNowCol + mshScale.FixedCols - 1)) = 1 Then
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
       
    aryNote = Split(mshScale.TextMatrix(GraphDataRow.未记说明, intNowCol + mshScale.FixedCols - 1), ";")
    If aryNote(Val(picGraph.Tag) + 1) <> "" Then
        If picGraph.MousePointer <> 12 Then picGraph.MousePointer = 12
    End If
    
    aryValue = Split(picScale.Tag, ";")
    
    strTmp = GetCurveDateTime(intNowCol, CDate(aryValue(0)), mlngHourBegin)
    dblValues = ConvertToValue(Val(picGraph.Tag), Y)
    
    If mItemSerial.体温 = Val(picGraph.Tag) Then
        dblValues = Format(dblValues, "0.00")
    Else
        dblValues = Format(dblValues, "0")
    End If
    
    If strTmp <> "" Then
        strTmp = "日期：" & Format(Split(strTmp, ",")(0), "yyyy-MM-dd") & " 时间：" & Format(Split(strTmp, ",")(0), "HH时mm分") & "～" & Format(Split(strTmp, ",")(1), "HH时mm分")
    End If
    
    RaiseEvent PromptInfo(strTmp & " " & mshScale.TextMatrix(0, Val(picGraph.Tag)) & "：" & dblValues)
    
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picGraph_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHead
    '-------------------------------------------------
    '根据按下鼠标状态进行数值记录：鼠标左键——增加修改操作；鼠标右键——删除操作
    '1、如果在指定时间没有数值，则记录数据，分别检测前和后时间内是否有数据，有则进行线条打印；
    '2、如果在指定时间已经有数值,则保存数据，同时调用整页重新绘画的方法
    '-------------------------------------------------
    Dim aryValue() As String
    Dim aryPart() As String
    Dim intRewrite As Integer
    Dim X0 As Long, Y0 As Long
    Dim strChar As String
    
    Dim aryData() As String
    Dim i As Long
    Dim intHave As Integer
    Dim intDots As Integer
    
    If picGraph.MousePointer <> 2 Then Exit Sub
    intCol = Int(X / HOUR_STEP_Twips)

    If Val(mshScale.TextMatrix(GraphDataRow.断开标志, intCol + mshScale.FixedCols)) = 1 Then Exit Sub
    
    X = intCol * HOUR_STEP_Twips + HOUR_STEP_Twips / 2
    
    aryValue = Split(mshScale.TextMatrix(GraphDataRow.更改标志, mshScale.FixedCols + intCol), ";")
    intRewrite = Val(aryValue(picGraph.Tag + 1))
    '------------------------------------------------------------------------------------------------------------------
    If Button = 1 Then
        Dim dblY As Double
        '增加或修改操作
        Select Case intRewrite
        Case 0 '在原本无的基础上删除的
            aryValue(picGraph.Tag + 1) = 2
        Case 1 '原本就有的
            aryValue(picGraph.Tag + 1) = 3
        Case 2 '新增的
            aryValue(picGraph.Tag + 1) = 2
        Case 3 '修改的
            aryValue(picGraph.Tag + 1) = 3
        Case 4 '在原本有的基础上删除的
            aryValue(picGraph.Tag + 1) = 3
        End Select
        
        mshScale.TextMatrix(GraphDataRow.更改标志, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
        aryValue = Split(mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol), ";")
        If 呼吸项目 Then
            '因呼吸项目刻度太小容易引起误差
            dblY = Format(ConvertToValue(mItemSerial.呼吸, Y), "0")
            Y = ConvertToY(mItemSerial.呼吸, dblY)
        End If
        aryValue(picGraph.Tag + 1) = Y
        mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
        If Val(picGraph.Tag) = mItemSerial.体温 Then
            aryPart() = Split(mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr体温部位
            mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol) = Join(aryPart, ";")
      
        End If
        If Val(picGraph.Tag) = mItemSerial.呼吸 Then
            aryPart = Split(mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr呼吸方式
            mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol) = Join(aryPart, ";")
        End If
        If Val(picGraph.Tag) = mItemSerial.脉搏 Then
            aryPart = Split(mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol), ";")
            aryPart(picGraph.Tag + 1) = mstr脉搏
            mshScale.TextMatrix(GraphDataRow.部位标志, mshScale.FixedCols + intCol) = Join(aryPart, ";")
        End If
    '------------------------------------------------------------------------------------------------------------------
    ElseIf Not (intRewrite = 0 Or intRewrite = 4) Then
    
        '检查鼠标位置是否在一个点上(附近)
        
        aryData = Split(mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol), ";")
        aryData = Split(aryData(picGraph.Tag + 1), ",")
        intDots = UBound(aryData) + 1 '已画的点数
        If Abs(Val(aryData(0)) - Y) <= 60 Then intHave = 1 '在第一个点上
        For i = 1 To UBound(aryData)
            If Abs(Val(aryData(i)) - Y) <= 60 Then intHave = i + 1: Exit For '在第i个点上
        Next
        
        If intHave = 0 Then
            '新增点
            Select Case intRewrite
            Case 1
                aryValue(picGraph.Tag + 1) = 3
            Case 2
                aryValue(picGraph.Tag + 1) = 2
            Case 3
                aryValue(picGraph.Tag + 1) = 3
            End Select
            mshScale.TextMatrix(GraphDataRow.更改标志, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        
            '新增点数据
            aryValue = Split(mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol), ";")
            If intDots = 1 Then
                '新增第二点
                Select Case Val(picGraph.Tag)
                Case mItemSerial.体温, mItemSerial.脉搏
                
                    '如果是体温项目，表示物理降温；脉搏项目表示心率，即脉搏短绌
                    If Y <> Val(aryValue(picGraph.Tag + 1)) Then
                        aryValue(picGraph.Tag + 1) = aryValue(picGraph.Tag + 1) & "," & Y
                    End If
                    
                End Select
                
            Else
                '修改第二点
                aryData = Split(aryValue(picGraph.Tag + 1), ",")
                If Y <> Val(aryData(0)) Then
                    aryData(intDots - 1) = Y
                    aryValue(picGraph.Tag + 1) = Join(aryData, ",")
                End If
                
            End If
            mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        Else
            '在某个点附近则删除该点
            Select Case intRewrite
            Case 1
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 3, 4)
            Case 2
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 2, 0)
            Case 3
                aryValue(picGraph.Tag + 1) = IIf(intDots > 1, 3, 4)
            End Select
            mshScale.TextMatrix(GraphDataRow.更改标志, mshScale.FixedCols + intCol) = Join(aryValue, ";")
            
            '删除该点数据(是删除,不是置空)
            aryValue = Split(mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol), ";")
            aryData = Split(aryValue(picGraph.Tag + 1), ",")
            aryData(intHave - 1) = " "
            aryValue(picGraph.Tag + 1) = Mid(Replace("," & Join(aryData, ","), ", ", ""), 2)
            mshScale.TextMatrix(GraphDataRow.曲线数据, mshScale.FixedCols + intCol) = Join(aryValue, ";")
        End If
    End If
        
    Call DrawPaper
    Call DrawGraph
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picPane_Resize()
    On Error GoTo errHand
    
    With vsb
        .Left = picPane.Width - .Width
        .Top = 0
        .Height = picPane.Height - hsb.Height
    End With
    
    With hsb
        .Left = 0
        .Top = picPane.Height - .Height
        .Width = picPane.Width - vsb.Width
    End With
    
    With picCover
        .Left = vsb.Left
        .Top = hsb.Top
    End With
    
    picCard(0).Move 60, 0
    
    '-------------------------------------------------
    '页面调整
    mshUpTab.Redraw = False
    mshScale.Redraw = False
    mshDownTab.Redraw = False
    
    With mshUpTab
        .Left = 60
        .Width = .ColWidth(0) + .ColWidth(1) * (.Cols - 1) + 15
        .Top = picCard(0).Top + picCard(0).Height + 30
        .Refresh
    End With
    
    picCard(0).Left = mshUpTab.Left
    picCard(0).Width = mshUpTab.Width
    
    With mshScale
        .Left = mshUpTab.Left
        .Width = mshUpTab.Width
        .Top = mshUpTab.Top + mshUpTab.Height - 15
        .Height = .Rows * .ROWHEIGHT(.Rows - 1) + 600
        .Refresh
    End With

    With vsf
        .Left = mshUpTab.Left
        .Top = mshScale.Top + mshScale.Height
        .Width = mshUpTab.Width
        .Visible = Not mbln呼吸曲线
    End With
        
    With mshDownTab
        .RowHeightMin = 255
        .Left = mshUpTab.Left
        .Top = IIf(mbln呼吸曲线, vsf.Top, vsf.Top + vsf.Height) - 15
        .Width = mshUpTab.Width
        .Height = (.Rows - 1) * .ROWHEIGHT(1) + 15
    End With
    
    lblComment.Left = mshUpTab.Left
    lblComment.Top = mshDownTab.Top + mshDownTab.Height + 45
        
    For intCol = 0 To picLine.UBound
        picLine(intCol).Move mshScale.Left + mshScale.ColWidth(0) * (intCol + 1), mshScale.Top, 15, mshScale.Height - 15
    Next
    
    With picScale
        .Left = mshUpTab.ColWidth(0) + mshUpTab.Left
        .Width = mshUpTab.ColWidth(1) * (mshUpTab.Cols - 1) + 15
        .Top = mshUpTab.Top + mshUpTab.Height - 15
        .Height = mshScale.ROWHEIGHT(0) + mshScale.ROWHEIGHT(1) / 2
    End With
    
    With picBack
        .Left = picScale.Left
        .Width = picScale.Width
        .Top = picScale.Top + picScale.Height - 15
        .Height = mshScale.Top + mshScale.Height - .Top
    End With
    
    mshUpTab.Redraw = True
    mshScale.Redraw = True
    mshDownTab.Redraw = True
    
    pic.Width = mshScale.Width + vsb.Width + 45
    pic.Height = lblComment.Top + lblComment.Height + 45 + hsb.Height
    
    '计算滚动条
    Call CalcScrollBarSize
    
    hsb.Value = 0
    vsb.Value = 0
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picScale_GotFocus()
    
    If mvarEdit = False Then Exit Sub
    
    lblCur.ForeColor = &H80000012
End Sub

Private Sub picScale_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo ErrHead
    
    Dim aryValue() As String
    Dim intMinCol As Long
    Dim intMaxCol As Long
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    
    '获得最小最大时间范围
    Call CalcMinMaxCol(picScale.Tag, intMinCol, intMaxCol)
    
    Select Case KeyCode
    Case 37     '左移动
        If lblCur.Left - HOUR_STEP_Twips >= intMinCol * HOUR_STEP_Twips And (lblCur.Left - HOUR_STEP_Twips) >= 0 Then lblCur.Left = lblCur.Left - HOUR_STEP_Twips
    Case 39     '右移动
        If lblCur.Left + HOUR_STEP_Twips <= intMaxCol * HOUR_STEP_Twips Then lblCur.Left = lblCur.Left + HOUR_STEP_Twips
    Case 13     'Enter输入数据
        
        RaiseEvent DbClickCur

    End Select
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrHead
    Dim aryValue() As String
    '获得最小最大时间范围
    If mvarEdit = False Then Exit Sub
    
    aryValue = Split(picScale.Tag, ";")
    If X < (Int((CDate(aryValue(0)) - Int(CDate(aryValue(0)))) * 24) \ 4 - 1) * HOUR_STEP_Twips + 15 Then Exit Sub
    If X > (Int((CDate(aryValue(1)) - Int(CDate(aryValue(0)))) * 24) \ 4 + 1) * HOUR_STEP_Twips - 15 Then Exit Sub
    lblCur.Left = Int(X / HOUR_STEP_Twips) * HOUR_STEP_Twips
    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtCard_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtCard(Index)
        
End Sub

Private Sub txtCard_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txtCard_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtCard(Index).Locked Then
        glngTXTProc = GetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtCard_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtCard(Index).Locked Then
        Call SetWindowLong(txtCard(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtInput_Change(Index As Integer)
    Dim blnCancel As Boolean
    
    If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.血压 Then Exit Sub
    If Index <> 0 Then Exit Sub
    
    If Len(txtInput(Index).Text) = 3 And GetTextPos(txtInput(Index).hWnd) = 4 Then
        
        If CheckBlood(0) Then
            txtInput(1).SetFocus
            zlControl.TxtSelAll txtInput(1)
        Else
            txtInput(1).Text = ""
            picInput.Visible = False
            mshDownTab.SetFocus
        End If
        
    End If

End Sub

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    
    If Val(mshDownTab.RowData(mshDownTab.Row)) <> mItemNo.血压 Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyLeft
        If Index = 1 Then
            If GetTextPos(txtInput(Index).hWnd) = 1 Then
            
                If CheckBlood(1) Then
                    txtInput(0).SetFocus
                    zlControl.TxtSelAll txtInput(0)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
        
            End If
        End If
    Case vbKeyRight
        If Index = 0 Then
            If GetTextPos(txtInput(Index).hWnd) >= Len(txtInput(0).Text) Then
                If CheckBlood(0) Then
                    txtInput(1).SetFocus
                    zlControl.TxtSelAll txtInput(1)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
            End If
        End If
    Case vbKeyBack
        If Index = 1 Then
            If GetTextPos(txtInput(Index).hWnd) = 1 Then

                If CheckBlood(1) Then
                    txtInput(0).SetFocus
                    If Len(txtInput(0).Text) > 0 Then
                       txtInput(0).Text = Left(txtInput(0).Text, Len(txtInput(0).Text) - 1)
                    End If
                    txtInput(0).SelStart = Len(txtInput(0).Text)
                Else
                    txtInput(1).Text = ""
                    picInput.Visible = False
                    mshDownTab.SetFocus
                End If
            End If
        End If
    End Select
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim intRow As Long
    Dim intCol As Long
    Dim blnCancel As Boolean
    
    If Val(mrsParam("编辑")) = 0 Then Exit Sub
    
    If KeyAscii = Asc("'") Or KeyAscii = Asc(";") Then
        KeyAscii = 0
    End If
    
    If mshDownTab.Tag = "" Then Exit Sub
    On Error Resume Next
    intRow = Split(picInput.Tag, ";")(0)
    intCol = Split(picInput.Tag, ";")(1)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtInput_Validate(Index, blnCancel)
        If blnCancel Then
            picInput.Visible = False
            Exit Sub
        End If
        
        If picInput.Visible Then
            mshDownTab.SetFocus
            Exit Sub
        Else
            mshDownTab.SetFocus
            Call mshDownTab_KeyPress(13)
        End If
        
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        picInput.Visible = False
        mshDownTab.SetFocus
    Else
    
        If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.血压 Then
            Select Case KeyAscii
            Case 191, vbKeyDivide, Asc("/")
                KeyAscii = 0
                If Index = 0 Then
                
                    If CheckBlood(0) Then
                        txtInput(1).SetFocus
                        zlControl.TxtSelAll txtInput(1)
                    Else
                        txtInput(1).Text = ""
                        picInput.Visible = False
                        mshDownTab.SetFocus
                    End If
                
                    End If

            End Select
        End If
    
'        Select Case mItemStru(intRow).数据类型
'        Case 0 '数值型
'            If Check是否包含(UCase(Chr(KeyAscii)), "正小数") = True Then KeyAscii = 0
'        Case 1 '字符型
            If Check是否包含(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
'        End Select
    End If

End Sub

Private Sub txtInput_LostFocus(Index As Integer)
    If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.血压 Then Exit Sub
    
    txtInput(Index).Text = Replace(txtInput(Index).Text, "'", "")
    picInput.Visible = False
End Sub

Private Function WriteScaleTab(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strInput As String) As Boolean

    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryPara() As String


    On Error GoTo ErrHead
        
    With mshScale
        
        '保存线条数据
        aryValue = Split(.TextMatrix(GraphDataRow.更改标志, intCol), ";")
        intRewrite = Val(aryValue(intRow + 1))
        
        If strInput <> "" Then
            '存在内容，相当于增加或修改操作
            Select Case intRewrite
            Case 0
                aryValue(intRow + 1) = 2
            Case 1
                aryValue(intRow + 1) = 3
            Case 2
                aryValue(intRow + 1) = 2
            Case 3
                aryValue(intRow + 1) = 3
            Case 4
                aryValue(intRow + 1) = 3
            End Select
        Else
            '没有内容，相当于删除操作
            Select Case intRewrite
            Case 0
                aryValue(intRow + 1) = 0
            Case 1
                aryValue(intRow + 1) = 4
            Case 2
                aryValue(intRow + 1) = 0
            Case 3
                aryValue(intRow + 1) = 4
            Case 4
                aryValue(intRow + 1) = 4
            End Select
        End If
        .TextMatrix(0, intCol) = Join(aryValue, ";")
        
        aryValue = Split(.TextMatrix(1, intCol), ";")
        If strInput <> "" Then
            '提取指定项目定义：最大值；最小值；单位值；最高行
            aryPara = Split(picLine(intRow).Tag, ";")
            aryValue(intRow + 1) = ((aryPara(0) - Val(strInput)) / aryPara(2) + aryPara(3) - 1) * .ROWHEIGHT(1)
        Else
            aryValue(intRow + 1) = ""
        End If
        .TextMatrix(1, intCol) = Join(aryValue, ";")
        
    End With
    
    '调用上级窗体进行图形处理
    Call DrawPaper
    Call DrawGraph
    
    WriteScaleTab = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Function WriteDownTab(ByVal intRow As Integer, ByVal intCol As Integer, ByVal strInput As String) As Boolean
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim strUnit As String
    Dim lngColor As Long
    
    On Error GoTo ErrHead
    
    With mshDownTab

        If .TextMatrix(intRow, intCol) = strInput Then Exit Function
        
        lngColor = GridTextColor(.TextMatrix(intRow, 0), strInput)
        .Cell(flexcpForeColor, intRow, intCol, intRow, intCol) = lngColor
             
        .TextMatrix(intRow, intCol) = strInput
        
        '处理增删改标志(输入框中没有任何数值，认为进行删除操作)
        If InStr(.TextMatrix(GridDataRow.修改标志, intCol), ";") = 0 Then
            ReDim aryValue(0 To 0) As String
            aryValue(0) = .TextMatrix(GridDataRow.修改标志, intCol)
        Else
            aryValue() = Split(.TextMatrix(GridDataRow.修改标志, intCol), ";")
        End If
        
        intRewrite = Val(aryValue(intRow - 1))
        
        If Trim(strInput) <> "" Then
            '增加或修改操作
            Select Case intRewrite
            Case 0
                aryValue(intRow - 1) = 2
            Case 1
                aryValue(intRow - 1) = 3
            Case 2
                aryValue(intRow - 1) = 2
            Case 3
                aryValue(intRow - 1) = 3
            Case 4
                aryValue(intRow - 1) = 3
            End Select
        Else
            If ((intCol + 1) - mshDownTab.FixedCols) / 2 <> ((intCol + 1) - mshDownTab.FixedCols) \ 2 And .TextMatrix(intRow, intCol) = .TextMatrix(intRow, intCol + 1) And .TextMatrix(intRow, intCol + 1) = "" Then
                '如果当前这个单元格为上午时
                '删除操作<---取消以前的删除为修改操作
                Select Case intRewrite
                Case 0
                    aryValue(intRow - 1) = 0
                Case 1
                    aryValue(intRow - 1) = 4
                Case 2
                    aryValue(intRow - 1) = 0
                Case 3
                    aryValue(intRow - 1) = 4
                Case 4
                    aryValue(intRow - 1) = 4
                End Select
            Else
                '操作类型：2-全部是新增的点,3-修改的点：可能包含原有的点和新增的点,4-删除的点
                '删除操作<---取消以前的删除为修改操作
                Select Case intRewrite
                Case 0
                    aryValue(intRow - 1) = 0
                Case 1
                    aryValue(intRow - 1) = 3
                Case 2
                    aryValue(intRow - 1) = 2
                Case 3
                    aryValue(intRow - 1) = 3
                Case 4
                    aryValue(intRow - 1) = 3
                End Select
            End If
        End If
        .TextMatrix(GridDataRow.修改标志, intCol) = Join(aryValue, ";")
        
        '修改单数单元格的操作标志
        '如果当前这个单元格为下午那么就修改上午的操作标志为修改或插入标志
        '否则如果当前单元格为上午并且为删除时就将上午的删除操作标志改为修改标志
        
        If ((intCol + 1) - mshDownTab.FixedCols) / 2 = ((intCol + 1) - mshDownTab.FixedCols) \ 2 Then
            If Trim(.TextMatrix(GridDataRow.修改标志, intCol - 1)) = "" Then
                .TextMatrix(GridDataRow.修改标志, intCol - 1) = " "
            End If
            aryValue() = Split(.TextMatrix(GridDataRow.修改标志, intCol - 1), ";")
            intRewrite = Val(aryValue(intRow - 1))
            
            '增加或修改操作
            Select Case intRewrite
            Case 0
                aryValue(intRow - 1) = 2
            Case 1
                aryValue(intRow - 1) = 3
            Case 2
                aryValue(intRow - 1) = 2
            Case 3
                aryValue(intRow - 1) = 3
            Case 4
                aryValue(intRow - 1) = 3
            End Select
            
            .TextMatrix(GridDataRow.修改标志, intCol - 1) = Join(aryValue, ";")
            
        End If
    End With
    
    WriteDownTab = True
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBlood(Index As Integer) As Boolean
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryBeforeValue() As String
    Dim strInput As String
    Dim dbMin As Double
    Dim dbMax As Double
    Dim strName As String
    
    CheckBlood = True
    
    If picInput.Visible = False Then Exit Function

    intRow = Split(picInput.Tag, ";")(0)
    intCol = Split(picInput.Tag, ";")(1)
        
    If Index = 0 Then
        '收缩压
        dbMin = Val(mItemStru(intRow).最小值)
        dbMax = Val(mItemStru(intRow).最大值)
        strName = "收缩压"
    Else
        '
        dbMin = Val(mItemOtherStru(1).最小值)
        dbMax = Val(mItemOtherStru(1).最大值)
        strName = "舒张压"
    End If
            
    If Trim(txtInput(Index).Text) <> "" And (Val(txtInput(Index).Text) > dbMax And dbMax <> 0 Or Val(txtInput(Index).Text) < dbMin And dbMin) And mItemStru(intRow).数据类型 = 0 Then
        
        picInput.Visible = True
        mstrSQL = "输入数值超过“" & strName & "”的允许范围：" & dbMin & "～" & dbMax
        ShowSimpleMsg mstrSQL

        CheckBlood = False
        Exit Function
    End If
    
    
End Function

Private Sub txtInput_Validate(Index As Integer, Cancel As Boolean)
    Dim aryValue() As String
    Dim intRewrite As Integer
    Dim aryBeforeValue() As String
    Dim strInput As String
    Dim dbMin As Double
    Dim dbMax As Double
    Dim strName As String
    Dim strTmp As String
    Dim intPos As Integer
    
    If txtInput(Index).Visible = False Then Exit Sub
    If mvarEdit = False Or mshDownTab.Tag = "" Then Exit Sub
    
    Cancel = Not StrIsValid(txtInput(Index).Text, txtInput(Index).MaxLength)
    If Cancel Then
        On Error Resume Next
        txtInput(Index).SetFocus
        Exit Sub
    End If

    On Error GoTo ErrHead

    With mshDownTab

        If picInput.Visible = False Then Exit Sub

        intRow = Split(picInput.Tag, ";")(0)
        intCol = Split(picInput.Tag, ";")(1)
        
        dbMin = Val(.TextMatrix(intRow, 3))
        dbMax = Val(.TextMatrix(intRow, 2))
        strName = mItemStru(intRow).项目名称
            
        Select Case Val(mshDownTab.RowData(mshDownTab.Row))
        Case mItemNo.血压
            If Index = 0 Then
                '收缩压
                dbMin = Val(mItemStru(intRow).最小值)
                dbMax = Val(mItemStru(intRow).最大值)
                strName = "收缩压"
            Else
                '
                dbMin = Val(mItemOtherStru(1).最小值)
                dbMax = Val(mItemOtherStru(1).最大值)
                strName = "舒张压"
            End If
        Case mItemNo.大便
            If Index = 0 Then
                
                txtInput(Index).Text = UCase(txtInput(Index).Text)
                
                strTmp = txtInput(Index).Text
                
                If strTmp <> "" Then
                    If Check是否包含(strTmp, "0123456789+/E*") = False Then
                        txtInput(Index).Text = ""
                    Else
                        intPos = InStr(strTmp, "E")
                        
                        If intPos > 0 Then
                            If Right(strTmp, 1) <> "E" Then
                                txtInput(Index).Text = ""
                            Else
                                If InStr(Mid(strTmp, 1, intPos - 1), "E") > 0 And intPos > 1 Then
                                    txtInput(Index).Text = ""
                                End If
                            End If
                            
                            intPos = InStr(strTmp, "/")
                            If intPos > 0 Then
                                If InStr(Mid(strTmp, 1, intPos - 1), "/") > 0 Then
                                    txtInput(Index).Text = ""
                                ElseIf InStr(Mid(strTmp, intPos + 1), "/") > 0 Then
                                    txtInput(Index).Text = ""
                                End If
                            End If
                            
                        ElseIf InStr(strTmp, "*") > 0 Then
                            If strTmp <> "*" Then
                                txtInput(Index).Text = ""
                            End If
                        End If
                        
                        If strTmp = "/E" Then txtInput(Index).Text = ""
                    End If
                End If
            End If
            
        Case mItemNo.出液
            If Index = 0 Then
            
                txtInput(Index).Text = UCase(txtInput(Index).Text)
                strTmp = txtInput(Index).Text
                
                If strTmp <> "" Then
                    If Check是否包含(strTmp, "0123456789/C") = False Then
                        txtInput(Index).Text = ""
                    Else
                    
                        intPos = InStr(strTmp, "/C")
                        If intPos > 0 Then
                        
                            If strTmp = "/C" Then
                            
                                txtInput(Index).Text = ""
                                
                            ElseIf InStr(Mid(strTmp, 1, intPos - 2), "/") > 0 And intPos > 2 Then
                                
                                txtInput(Index).Text = ""
                                
                            ElseIf InStr(Mid(strTmp, 1, intPos - 2), "C") > 0 And intPos > 2 Then
                                
                                txtInput(Index).Text = ""
                                
                            ElseIf Right(strTmp, 2) <> "/C" Then
                                
                                txtInput(Index).Text = ""
                            
                            End If
                        ElseIf InStr(strTmp, "C") > 0 Then
                            If strTmp <> "C" Then
                                txtInput(Index).Text = ""
                            End If
                        End If
                    End If
                End If
            End If
        End Select
        
        If CheckStrType(txtInput(Index).Text, 99, "0123456789.") Then
            If Trim(txtInput(Index).Text) <> "" And (Val(txtInput(Index).Text) > dbMax And dbMax <> 0 Or Val(txtInput(Index).Text) < dbMin And dbMin) And mItemStru(intRow).数据类型 = 0 Then
                Cancel = True
                
                picInput.Visible = True
                mstrSQL = "输入数值超过“" & strName & "”的允许范围：" & dbMin & "～" & dbMax
                ShowSimpleMsg mstrSQL
                txtInput(Index) = ""
                txtInput(Index).SetFocus
                mshDownTab.SetFocus
                
                Exit Sub
            End If
            
            If CheckNumber(Val(txtInput(Index).Text), mItemStru(intRow).数据长度, mItemStru(intRow).小数位数) = False Then
                ShowSimpleMsg "“" & strName & "”的整数位最长:" & mItemStru(intRow).数据长度 - mItemStru(intRow).小数位数 & "；小数位最长:" & mItemStru(intRow).小数位数
                mshDownTab.SetFocus
                Exit Sub
            End If
        End If
        
        '填写单元数值
        If mItemStru(intRow).数据类型 = 0 Then
            If Val(mshDownTab.RowData(mshDownTab.Row)) = mItemNo.血压 Then
                strInput = txtInput(0).Text & "/" & txtInput(1).Text
                If strInput = "/" Then strInput = ""
            Else
                If Trim(txtInput(Index).Text) = "" Then
                    strInput = ""
                Else
                    If CheckStrType(txtInput(Index).Text, 99, "0123456789.") Then
                        If mItemStru(intRow).小数位数 > 0 Then
                            strInput = Format(Val(txtInput(Index).Text), "0." & String(mItemStru(intRow).小数位数, "0"))
                        Else
                            strInput = Format(Val(txtInput(Index).Text), "0")
                        End If
                    Else
                        strInput = txtInput(Index).Text
                    End If
                    
                End If
            End If
        ElseIf mItemStru(intRow).数据类型 = 1 Then
            strInput = Trim(txtInput(Index).Text)
        End If
        
        picInput.Visible = False
        
        Call WriteDownTab(intRow, intCol, strInput)
        
    End With
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub UserControl_GotFocus()
    RaiseEvent Activate

End Sub

Private Sub UserControl_Initialize()
    mstr体温部位 = "腋温"
    mstr呼吸方式 = "自主呼吸"
    mstr脉搏 = ""
    Call InitCommandBar
End Sub

Private Sub vsb_Change()
    On Error Resume Next
    
'    pic.Top = 0 - vsb.Value * 300
    pic.Top = 0 - vsb.Value * msinVStep
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim intRewrite As Integer
    Dim strInput As String
    
    '操作类型：2-全部是新增的点,3-修改的点：可能包含原有的点和新增的点,4-删除的点
    
    strInput = vsf.TextMatrix(Row, Col)
     
    intRewrite = Val(vsf.ColData(Col))
    
    If Trim(strInput) <> "" Then
        '增加或修改操作
        Select Case intRewrite
        Case 0
            vsf.ColData(Col) = 2
        Case 1
            vsf.ColData(Col) = 3
        Case 2
            vsf.ColData(Col) = 2
        Case 3
            vsf.ColData(Col) = 3
        Case 4
            vsf.ColData(Col) = 3
        End Select
    Else
        vsf.ColData(Col) = 4
    End If
    
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    
    On Error Resume Next
    
    If NewCol <= 1 Then Exit Sub
    

    strInfo = "“" & vsf.TextMatrix(NewRow, 1) & "”项目范围：" & mItemOtherStru(0).最小值 & "～" & mItemOtherStru(0).最大值

    
    RaiseEvent PromptInfo(strInfo)
End Sub

Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim aryValue() As String
    Dim strStart As String
    Dim strEnd As String
    Dim strFrom As String
    Dim strTo As String
    Dim strTmp As String
    
    On Error Resume Next
    
    If Val(mrsParam("编辑")) = 0 Then
        vsf.EditMode(NewCol) = 0
        Exit Sub
    End If
    If picScale.Tag = "" Then
        vsf.EditMode(NewCol) = 0
        Exit Sub
    End If
    
    '获得最小最大时间范围
    
    strFrom = Split(picScale.Tag, ";")(0)
    strTo = Split(picScale.Tag, ";")(1)
    
    strTmp = GetCurveDateTime(NewCol - vsf.FixedCols + 1, CDate(strFrom), mlngHourBegin)
    strStart = Split(strTmp, ",")(0)
    strEnd = Split(strTmp, ",")(1)
    
    If (strStart >= strFrom And strStart <= strTo) Or (strEnd >= strFrom And strEnd <= strTo) Then
        vsf.EditMode(NewCol) = 1
    Else
        vsf.EditMode(NewCol) = 0
    End If
    
End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        If CheckStrType(Chr(KeyAscii), 99, "0123456789") = False Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsf_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then        'Delete清除单元
        If vsf.Body.Editable = flexEDKbdMouse And vsf.Row = 1 And vsf.Col > 1 Then
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
        End If
    End If
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Call vsf_BeforeRowColChange(0, 0, Row, Col, False)
    If vsf.EditMode(Col) = 0 Then Cancel = True
    
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsf.EditText <> "" Then
        If Val(vsf.EditText) < mItemOtherStru(0).最小值 Or Val(vsf.EditText) > mItemOtherStru(0).最大值 Then
            vsf.EditText = ""
            Cancel = True
        End If
        
    End If
End Sub

Private Sub Get入院入科时间()
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHand
    '数据的发生时间不能小于入科时间,更不能小于入院时间
    
    gstrSQL = " Select MIN(开始时间) AS 时间 From 病人变动记录 Where 病人ID=[1] And 主页ID=[2] And 病区ID=[3]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", CLng(mrsParam!病人ID), CLng(mrsParam!主页ID), CLng(mrsParam!科室ID))
    mstr最小时间 = Format(DateAdd("n", 1, rsCheck!时间), "yyyy-MM-dd HH:mm:00")    '入院时间有可能存在秒数,需进位到分
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub OutputDadaForDebug()
    Dim strRow As String
    Dim intRow As Integer, intCol As Integer
    Dim intRows As Integer, intCols As Integer
    '输出表格内记录的数据
    
    intRows = mshScale.Rows - 1
    intCols = mshScale.Cols - 1
    
    For intRow = 0 To intRows
        strRow = ""
        For intCol = 0 To intCols
            strRow = strRow & "," & mshScale.TextMatrix(intRow, intCol)
        Next
        'Debug.Print "Row:" & intRow & Space(4) & Mid(strRow, 2)
    Next
End Sub
