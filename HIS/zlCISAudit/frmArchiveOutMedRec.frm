VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmArchiveOutMedRec 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "门诊首页"
   ClientHeight    =   8580
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   7905
   Icon            =   "frmArchiveOutMedRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   8235
      Left            =   615
      TabIndex        =   25
      Top             =   150
      Width           =   6570
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   20
         Left            =   4635
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   24
         Top             =   7875
         Width           =   1725
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   18
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   9
         Top             =   1410
         Width           =   2310
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   17
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   8
         Top             =   1410
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   16
         Left            =   4050
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   7
         Top             =   1050
         Width           =   2310
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   15
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   6
         Top             =   1050
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   14
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   3
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   13
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin VB.CheckBox chkEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "复诊(&R)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   285
         TabIndex        =   22
         Top             =   7890
         Width           =   930
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   12
         Left            =   315
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   7290
         Width           =   6045
      End
      Begin VB.CheckBox chkEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "传染病上传(&U)"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   23
         Top             =   7890
         Width           =   1470
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   18
         Top             =   4080
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   5
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2280
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   17
         Top             =   3720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         Top             =   3720
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   9
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3360
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   4965
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   14
         Top             =   3000
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   13
         Top             =   3000
         Width           =   3030
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   6
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   12
         Top             =   2640
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   10
         Top             =   1920
         Width           =   5400
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   3
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   4
         Top             =   540
         Width           =   1230
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         MaxLength       =   64
         TabIndex        =   0
         Top             =   180
         Width           =   1575
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5220
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1140
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiag 
         Height          =   915
         Left            =   135
         TabIndex        =   20
         Top             =   5940
         Width           =   6225
         _cx             =   10980
         _cy             =   1614
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
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmArchiveOutMedRec.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   115
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
      Begin VSFlex8Ctl.VSFlexGrid vsAller 
         Height          =   915
         Left            =   135
         TabIndex        =   19
         Top             =   4665
         Width           =   6225
         _cx             =   10980
         _cy             =   1614
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
         BackColorSel    =   4210752
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   8421504
         GridColorFixed  =   8421504
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   225
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmArchiveOutMedRec.frx":0072
         ScrollTrack     =   -1  'True
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
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   20
         X1              =   4560
         X2              =   6365
         Y1              =   8070
         Y2              =   8070
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   19
         X1              =   135
         X2              =   6365
         Y1              =   7785
         Y2              =   7785
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   18
         X1              =   4890
         X2              =   6360
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   17
         X1              =   4890
         X2              =   6360
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   16
         X1              =   885
         X2              =   3990
         Y1              =   4275
         Y2              =   4275
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   15
         X1              =   885
         X2              =   3990
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   14
         X1              =   885
         X2              =   3990
         Y1              =   3195
         Y2              =   3195
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   13
         X1              =   885
         X2              =   6360
         Y1              =   3555
         Y2              =   3555
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   12
         X1              =   885
         X2              =   6360
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   11
         X1              =   885
         X2              =   6360
         Y1              =   2475
         Y2              =   2475
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   10
         X1              =   885
         X2              =   6360
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   9
         X1              =   3990
         X2              =   6360
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   3990
         X2              =   6360
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   885
         X2              =   3255
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   6
         X1              =   885
         X2              =   3255
         Y1              =   1245
         Y2              =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   5160
         X2              =   6360
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   5160
         X2              =   6360
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   3030
         X2              =   4315
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   3030
         X2              =   4315
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   885
         X2              =   2550
         Y1              =   735
         Y2              =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   885
         X2              =   2550
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊摘要："
         Height          =   180
         Index           =   20
         Left            =   120
         TabIndex        =   48
         Top             =   7005
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "过敏记录："
         Height          =   180
         Index           =   18
         Left            =   120
         TabIndex        =   47
         Top             =   4455
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断记录："
         Height          =   180
         Index           =   19
         Left            =   120
         TabIndex        =   46
         Top             =   5730
         Width           =   900
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发病时间"
         Height          =   180
         Index           =   21
         Left            =   3810
         TabIndex        =   45
         Top             =   7890
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "监护人"
         Height          =   180
         Index           =   22
         Left            =   300
         TabIndex        =   44
         Top             =   4095
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   42
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付款方式"
         Height          =   180
         Index           =   5
         Left            =   4380
         TabIndex        =   41
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭邮编"
         Height          =   180
         Index           =   17
         Left            =   4095
         TabIndex        =   40
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   39
         Top             =   3735
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Index           =   15
         Left            =   120
         TabIndex        =   38
         Top             =   3375
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Index           =   14
         Left            =   4095
         TabIndex        =   37
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Index           =   13
         Left            =   120
         TabIndex        =   36
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位名称"
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   1935
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   33
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Index           =   7
         Left            =   3555
         TabIndex        =   32
         Top             =   1065
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Index           =   9
         Left            =   3555
         TabIndex        =   31
         Top             =   1425
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   1425
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Index           =   2
         Left            =   2610
         TabIndex        =   29
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   28
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   27
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Index           =   3
         Left            =   4560
         TabIndex        =   26
         Top             =   195
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmArchiveOutMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long
Private mlng挂号ID As Long
Private mblnMoved As Boolean
Private mblnCheck As Boolean

Private Enum TXT_ENUM
    txt姓名 = 0
    txt性别 = 13
    txt年龄 = 3
    txt国籍 = 15
    txt民族 = 16
    txt婚姻 = 17
    txt职业 = 18
    txt门诊号 = 1
    txt监护人 = 2
    txt出生日期 = 14
    txt身份证号 = 4
    txt出生地点 = 5
    txt工作单位 = 6
    txt单位电话 = 7
    txt单位邮编 = 8
    txt家庭地址 = 9
    txt家庭电话 = 10
    txt家庭邮编 = 11
    txt就诊摘要 = 12
    txt付款方式 = 19
    txt发病时间 = 20
End Enum
Private Enum CHK_ENUM
    chk复诊 = 0
    chk传染病上传 = 1
End Enum
Private Enum COL_ENUM
    col类型 = 0
    col诊断 = 1
    col疑诊 = 2
End Enum

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng挂号id As Long, ByVal blnMoved As Boolean) As Boolean
'功能：刷新或清除医嘱清单
    mlng病人ID = lng病人ID: mlng挂号ID = lng挂号id: mblnMoved = blnMoved
    zlRefresh = LoadMedRec
End Function

Private Function LoadMedRec() As Boolean
'功能：读取门诊首页的各种信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim lngRow As Long, bln中医 As Boolean
    
    mblnCheck = True
    
    On Error GoTo errH
    
    '基本信息
    strSQL = "Select B.执行部门ID as 科室ID,B.摘要,B.复诊," & _
        " B.传染病上传,B.发病时间,A.险类,A.门诊号,A.姓名,A.性别,A.年龄,A.出生日期,A.医疗付款方式," & _
        " A.国籍,A.民族,A.婚姻状况,A.职业,A.身份证号,A.出生地点,A.监护人,A.家庭地址,A.家庭电话," & _
        " A.家庭地址邮编,A.工作单位,A.单位电话,A.单位邮编" & _
        " From 病人信息 A,病人挂号记录 B Where A.病人ID=B.病人ID And B.ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人挂号记录", "H病人挂号记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng挂号ID)
    If rsTmp.EOF Then Exit Function
    
    bln中医 = Have部门性质(rsTmp!科室ID, "中医科")
        
    txtEdit(txt姓名).Text = NVL(rsTmp!姓名)
    txtEdit(txt性别).Text = NVL(rsTmp!性别)
    txtEdit(txt年龄).Text = NVL(rsTmp!年龄)
    txtEdit(txt门诊号).Text = NVL(rsTmp!门诊号)
    
    txtEdit(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd")
    If Format(rsTmp!出生日期, "HH:mm") <> "00:00" Then
        txtEdit(txt出生日期).Text = Format(rsTmp!出生日期, "yyyy-MM-dd HH:mm")
    End If
    
    txtEdit(txt付款方式) = NVL(rsTmp!医疗付款方式)
    txtEdit(txt国籍) = NVL(rsTmp!国籍)
    txtEdit(txt民族) = NVL(rsTmp!民族)
    txtEdit(txt婚姻) = NVL(rsTmp!婚姻状况)
    txtEdit(txt职业) = NVL(rsTmp!职业)
    txtEdit(txt监护人).Text = NVL(rsTmp!监护人)
    txtEdit(txt身份证号).Text = NVL(rsTmp!身份证号)
    txtEdit(txt出生地点).Text = NVL(rsTmp!出生地点)
    txtEdit(txt工作单位).Text = NVL(rsTmp!工作单位)
    txtEdit(txt单位电话).Text = NVL(rsTmp!单位电话)
    txtEdit(txt单位邮编).Text = NVL(rsTmp!单位邮编)
    txtEdit(txt家庭地址).Text = NVL(rsTmp!家庭地址)
    txtEdit(txt家庭电话).Text = NVL(rsTmp!家庭电话)
    txtEdit(txt家庭邮编).Text = NVL(rsTmp!家庭地址邮编)
    txtEdit(txt就诊摘要).Text = NVL(rsTmp!摘要)
    chkEdit(chk复诊).Value = NVL(rsTmp!复诊, 0)
    chkEdit(chk传染病上传).Value = NVL(rsTmp!传染病上传, 0)

    txtEdit(txt发病时间).Text = Format(rsTmp!发病时间, "yyyy-MM-dd")
    If Format(rsTmp!发病时间, "HH:mm") <> "00:00" Then
        txtEdit(txt发病时间).Text = Format(rsTmp!发病时间, "yyyy-MM-dd HH:mm")
    End If
    
    '过敏信息:本次挂号的,过敏的
    strSQL = "Select 记录来源,Decode(过敏时间,Null ,记录时间,过敏时间) as 过敏时间,药物ID,药物名 From 病人过敏记录 A" & _
        " Where 结果=1 And 病人ID=[1] And 主页ID=[2]" & _
        " And Not Exists(Select 药物ID From 病人过敏记录" & _
            " Where (Nvl(药物ID,0)=Nvl(A.药物ID,0) Or Nvl(药物名,'Null')=Nvl(A.药物名,'Null'))" & _
            " And Nvl(结果,0)=0 And 记录时间>=A.记录时间 And 病人ID=[1] And 主页ID=[2])" & _
        " Order by 过敏时间,药物名"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人过敏记录", "H病人过敏记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    vsAller.Rows = vsAller.FixedRows
    If Not rsTmp.EOF Then
        rsTmp.Filter = "记录来源=3" '首页本身填写的
        If rsTmp.EOF Then rsTmp.Filter = "记录来源<>3" '其它来源的作为缺省显示
        With vsAller
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                '其它来源的可能有重复
                lngRow = -1
                If Not IsNull(rsTmp!药物ID) Then
                    lngRow = .FindRow(CLng(rsTmp!药物ID))
                ElseIf Not IsNull(rsTmp!药物名) Then
                    lngRow = .FindRow(CStr(rsTmp!药物名), , 1)
                End If
                If lngRow = -1 Then
                    .RowData(i) = CLng(NVL(rsTmp!药物ID, 0))
                    .TextMatrix(i, 0) = Format(rsTmp!过敏时间, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, 1) = NVL(rsTmp!药物名)
                End If
                rsTmp.MoveNext
            Next
            .Row = 1: .Col = 1
        End With
    End If
    
    '诊断信息:本次挂号的
    strSQL = "Select 记录来源,诊断类型,疾病ID,诊断ID,证候ID,诊断描述,是否疑诊 From 病人诊断记录" & _
        " Where 记录来源 IN(1,3) And 诊断类型 IN(1,11)" & _
        " And 取消时间 is Null And 病人ID=[1] And 主页ID=[2]" & _
        " Order by 诊断类型,诊断次序"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人诊断记录", "H病人诊断记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng挂号ID)
    
    vsDiag.Rows = vsDiag.FixedRows
    If Not rsTmp.EOF Then
        With vsDiag
            '西医诊断
            rsTmp.Filter = "诊断类型=1 And 记录来源=3" '首页本身填写的
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=1 And 记录来源<>3" '其它来源的作为缺省显示
            .Rows = .Rows + rsTmp.RecordCount
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, col类型) = "西医"
                .TextMatrix(.Rows - 1, col诊断) = NVL(rsTmp!诊断描述)
                .TextMatrix(.Rows - 1, col疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                rsTmp.MoveNext
            Loop
            
            '中医诊断
            rsTmp.Filter = "诊断类型=11 And 记录来源=3"
            If rsTmp.EOF Then rsTmp.Filter = "诊断类型=11 And 记录来源<>3"
            If rsTmp.EOF Then .ColHidden(col类型) = True
            .Rows = .Rows + rsTmp.RecordCount
            Do While Not rsTmp.EOF
                .TextMatrix(.Rows - 1, col类型) = "中医"
                .TextMatrix(.Rows - 1, col诊断) = NVL(rsTmp!诊断描述)
                .TextMatrix(.Rows - 1, col疑诊) = IIf(NVL(rsTmp!是否疑诊, 0) = 1, "？", "")
                rsTmp.MoveNext
            Loop
            
            .Cell(flexcpForeColor, .FixedRows, col疑诊, .Rows - 1, col疑诊) = vbRed
            .Row = .FixedRows: .Col = col诊断
        End With
    End If
    
    mblnCheck = False
    LoadMedRec = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub chkEdit_Click(Index As Integer)
    If Not mblnCheck Then
        mblnCheck = True
        chkEdit(Index).Value = IIf(chkEdit(Index).Value = 1, 0, 1)
        mblnCheck = False
    End If
End Sub

Private Sub Form_Load()
    Me.BackColor = fraBack.BackColor
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraBack.Top = 0
    fraBack.Left = 0
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtEdit(Index))
End Sub
