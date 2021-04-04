VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmRegistPlanInvalidation 
   Caption         =   "停用时间设置"
   ClientHeight    =   8475
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "frmRegistPlanInvalidation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   11865
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picCmd 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   570
      ScaleHeight     =   600
      ScaleWidth      =   9225
      TabIndex        =   52
      Top             =   7890
      Width           =   9225
      Begin VB.Frame fraSplit 
         Height          =   60
         Index           =   1
         Left            =   -30
         TabIndex        =   56
         Top             =   -15
         Width           =   11805
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   6780
         TabIndex        =   55
         Top             =   75
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   8040
         TabIndex        =   54
         Top             =   60
         Width           =   1100
      End
      Begin VB.CheckBox chkClearHistory 
         Caption         =   "清除所有已失效的停用时间(&S)"
         Height          =   285
         Left            =   2190
         TabIndex        =   53
         Top             =   75
         Width           =   2835
      End
   End
   Begin VB.PictureBox picStop 
      BorderStyle     =   0  'None
      Height          =   3390
      Left            =   5160
      ScaleHeight     =   3390
      ScaleWidth      =   6645
      TabIndex        =   40
      Top             =   195
      Width           =   6645
      Begin VB.CommandButton cmdSel 
         Caption         =   "&P"
         Height          =   285
         Left            =   4560
         TabIndex        =   58
         Top             =   435
         Width           =   315
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "恢复停用安排(&Y)"
         Height          =   345
         Left            =   4905
         TabIndex        =   57
         Top             =   30
         Width           =   1710
      End
      Begin VB.CommandButton cmdDeleteTime 
         Caption         =   "删除(&R)"
         Height          =   345
         Left            =   5775
         TabIndex        =   51
         Top             =   405
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   345
         Left            =   4905
         TabIndex        =   48
         Top             =   405
         Width           =   855
      End
      Begin VB.TextBox txtMemo 
         Height          =   315
         Left            =   840
         MaxLength       =   100
         TabIndex        =   41
         Top             =   420
         Width           =   4050
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   840
         TabIndex        =   42
         Top             =   75
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109445123
         CurrentDate     =   40427.6041666667
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   3045
         TabIndex        =   43
         Top             =   75
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         CalendarTitleBackColor=   8421504
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   109445123
         CurrentDate     =   40427.0416666667
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   2415
         Left            =   0
         TabIndex        =   44
         Top             =   795
         Width           =   6345
         _cx             =   11192
         _cy             =   4260
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanInvalidation.frx":030A
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
         ExplorerBar     =   7
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
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "停用时间"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   47
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   2730
         TabIndex        =   46
         Top             =   105
         Width           =   225
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "备注"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   45
         Top             =   480
         Width           =   360
      End
   End
   Begin VB.PictureBox picOthers 
      BorderStyle     =   0  'None
      Height          =   3330
      Left            =   5115
      ScaleHeight     =   3330
      ScaleWidth      =   7815
      TabIndex        =   34
      Top             =   4305
      Width           =   7815
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&U)"
         Height          =   345
         Left            =   5325
         TabIndex        =   50
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "删除(&D)"
         Height          =   345
         Left            =   4425
         TabIndex        =   49
         Top             =   30
         Width           =   855
      End
      Begin VB.TextBox txtCode 
         Height          =   315
         Left            =   525
         TabIndex        =   36
         Top             =   45
         Width           =   2580
      End
      Begin VB.CommandButton cmdOthers 
         Caption         =   "其他条件(&O)"
         Height          =   345
         Left            =   3165
         TabIndex        =   35
         Top             =   30
         Width           =   1230
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOthers 
         Height          =   2430
         Left            =   45
         TabIndex        =   38
         Top             =   465
         Width           =   6405
         _cx             =   11298
         _cy             =   4286
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   9
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   23
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmRegistPlanInvalidation.frx":03F9
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
         ExplorerBar     =   7
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
         Begin VB.PictureBox picImgList 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   45
            ScaleHeight     =   225
            ScaleWidth      =   210
            TabIndex        =   39
            Top             =   60
            Width           =   210
            Begin VB.Image imgColList 
               Height          =   195
               Left            =   0
               Picture         =   "frmRegistPlanInvalidation.frx":06AE
               ToolTipText     =   "选择需要显示的列(ALT+C)"
               Top             =   0
               Width           =   195
            End
         End
      End
      Begin VB.Label lblCon 
         AutoSize        =   -1  'True
         Caption         =   "号别"
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   37
         Top             =   105
         Width           =   360
      End
   End
   Begin VB.PictureBox picBill 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      ScaleHeight     =   7815
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.Frame Frame1 
         Caption         =   "基本信息"
         Height          =   1860
         Left            =   60
         TabIndex        =   12
         Top             =   105
         Width           =   4890
         Begin VB.TextBox txt号别 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            MaxLength       =   5
            TabIndex        =   21
            Top             =   270
            Width           =   960
         End
         Begin VB.TextBox txt限号 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   20
            Top             =   660
            Width           =   1215
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1035
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   660
            TabIndex        =   18
            Top             =   1410
            Width           =   2115
         End
         Begin VB.ComboBox cbo科室 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   660
            Width           =   2115
         End
         Begin VB.TextBox txt限约 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            MaxLength       =   5
            TabIndex        =   16
            Top             =   1035
            Width           =   1215
         End
         Begin VB.CheckBox chk病案 
            Caption         =   "挂号时必须建病案"
            Height          =   195
            Left            =   2985
            TabIndex        =   15
            Top             =   1463
            Width           =   1755
         End
         Begin VB.ComboBox cbo号类 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3540
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   1230
         End
         Begin VB.CheckBox chk序号控制 
            Caption         =   "序号控制"
            Height          =   255
            Left            =   1750
            TabIndex        =   13
            Top             =   293
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "号别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   210
            TabIndex        =   28
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "科室"
            Height          =   180
            Left            =   240
            TabIndex        =   27
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "项目"
            Height          =   180
            Left            =   240
            TabIndex        =   26
            Top             =   1110
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "医生"
            Height          =   180
            Left            =   240
            TabIndex        =   25
            Top             =   1485
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "限号"
            Height          =   180
            Left            =   3105
            TabIndex        =   24
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "限约"
            Height          =   180
            Left            =   3105
            TabIndex        =   23
            Top             =   1095
            Width           =   360
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "号类"
            Height          =   180
            Left            =   3105
            TabIndex        =   22
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "应诊诊室:"
         Height          =   2730
         Left            =   75
         TabIndex        =   6
         Top             =   5070
         Width           =   4860
         Begin VB.OptionButton opt分诊 
            Caption         =   "不分诊"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "指定诊室"
            Height          =   180
            Index           =   1
            Left            =   1020
            TabIndex        =   9
            Top             =   300
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "动态分诊"
            Height          =   180
            Index           =   2
            Left            =   2115
            TabIndex        =   8
            Top             =   300
            Width           =   1020
         End
         Begin VB.OptionButton opt分诊 
            Caption         =   "平均分诊"
            Height          =   180
            Index           =   3
            Left            =   3135
            TabIndex        =   7
            Top             =   315
            Width           =   1020
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   2040
            Left            =   105
            TabIndex        =   11
            Top             =   615
            Width           =   4650
            _ExtentX        =   8202
            _ExtentY        =   3598
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "应诊时间"
         Height          =   2835
         Left            =   60
         TabIndex        =   1
         Top             =   2070
         Width           =   4890
         Begin VSFlex8Ctl.VSFlexGrid vsPlan1 
            Height          =   660
            Left            =   1200
            TabIndex        =   33
            Top             =   1305
            Width           =   3510
            _cx             =   6191
            _cy             =   1164
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanInvalidation.frx":0BFC
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
         Begin VB.CheckBox chk有效期 
            Caption         =   "有效期"
            Height          =   195
            Left            =   285
            TabIndex        =   29
            Top             =   2085
            Width           =   855
         End
         Begin VB.OptionButton opt天 
            Caption         =   "每天(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   5
            Top             =   285
            Width           =   960
         End
         Begin VB.OptionButton opt周 
            Caption         =   "每周(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   4
            Top             =   630
            Width           =   930
         End
         Begin VB.ComboBox cbo天 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   270
            Width           =   1110
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   660
            Left            =   1200
            TabIndex        =   2
            Top             =   690
            Width           =   3510
            _cx             =   6191
            _cy             =   1164
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
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanInvalidation.frx":0C43
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
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1200
            TabIndex        =   30
            Top             =   2040
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   109445123
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   1215
            TabIndex        =   31
            Top             =   2415
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   109445123
            CurrentDate     =   38091
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   930
            TabIndex        =   32
            Top             =   2475
            Width           =   180
         End
      End
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRegistPlanInvalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng安排ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
Private mlngModule As Long, mstrPrivs As String
Private mrsRoom As ADODB.Recordset
Private mstrDelete序号 As String   '删除序号

Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
     Optional lng安排ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示所要修改的计划安排
    '入参:mfrmMain-调用的主窗口
    '     lngModule-模块号
    '     strPrivs-权限串
    '     lng安排ID-挂号安排ID.
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 10:05:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
     mlngModule = lngModule: mstrPrivs = strPrivs:  mblnSucces = False: mlng安排ID = lng安排ID
    Me.Show 1, mfrmMain
    ShowCard = mblnSucces
End Function
Private Function InitPanel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化条件区哉
    '编制:刘兴洪
    '日期:2010-09-08 11:41:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, strKey As String
    Dim objBill As Pane, lngHeight As Long, lngBillWidth As Long
    lngHeight = picCmd.Height \ Screen.TwipsPerPixelY
    lngBillWidth = picBill.Width \ Screen.TwipsPerPixelX
    With dkpMan
        Set objPane = .CreatePane(3, 400, 400, DockBottomOf, Nothing)
        objPane.Title = "按钮项"
        objPane.Options = PaneNoCloseable Or PaneNoHideable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
        objPane.Handle = picCmd.Hwnd
        objPane.MaxTrackSize.Height = lngHeight
        objPane.MinTrackSize.Height = lngHeight
        
        Set objBill = .CreatePane(1, 300, 100, DockTopOf, objPane)
        objBill.Title = "当前挂号安排": objBill.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objBill.Handle = picBill.Hwnd
        objBill.MaxTrackSize.Width = lngBillWidth: objBill.MinTrackSize.Width = lngBillWidth:
         Set objPane = .CreatePane(2, 400, 400, DockRightOf, objBill)
        objPane.Title = "停用时间设置"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStop.Hwnd
         Set objPane = .CreatePane(3, 400, 400, DockBottomOf, objPane)
        objPane.Title = "应用于其他挂号项目"
        objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picOthers.Hwnd
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
    dkpMan.RecalcLayout: DoEvents
    'zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Function



Private Function LoadData(Optional blnRestore As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载计划安排数据信息
    '编制:刘兴洪
    '日期:2009-09-14 14:40:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String, i As Long
    Dim strCurDate As String

    Err = 0: On Error GoTo Errhand:
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    mstrDelete序号 = ""
    dtpStartDate.Value = Format(CDate(strCurDate) + 1, "yyyy-mm-dd 00:00:00")
    dtpEndDate.Value = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 23:59:59")
    dtpStartDate.MinDate = CDate(strCurDate)
    dtpEndDate.MinDate = dtpStartDate.MinDate
    
   strSQL = " " & _
    "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id,  F.限号数,  F.限约数,   " & _
    "           A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六, " & _
    "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
    "   From 挂号安排 A,收费项目目录 B,挂号安排计划 C,部门表 D,挂号安排限制 F " & _
    "   Where A.Id=C.安排ID(+) And A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
    "         And A.Id=[1]  And a.Id = f.安排id(+) And" & _
    "  Decode(To_Char([2], 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =f.限制项目(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, CDate(strCurDate))
    If rsTemp.EOF Then
        MsgBox "注意:" & vbCrLf & _
        "    挂号安排可能已经被他人删除,不能再进行停用时间设置!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '加载数据到控件中
    txt号别.Text = Nvl(rsTemp!号码)
    cbo号类.AddItem Nvl(rsTemp!号类): cbo号类.ListIndex = cbo号类.NewIndex
    txt限号.Text = Nvl(rsTemp!限号数): txt限约.Text = Nvl(rsTemp!限约数)
    chk序号控制.Value = IIf(Val(Nvl(rsTemp!序号控制)) = 1, 1, 0)
    chk病案.Value = IIf(Val(Nvl(rsTemp!病案必须)) = 1, 1, 0)
    With cbo科室
        .AddItem Nvl(rsTemp!科室): .ItemData(.NewIndex) = Val(Nvl(rsTemp!科室ID)): .ListIndex = .NewIndex
    End With
    With cboItem
        .AddItem Nvl(rsTemp!项目): .ItemData(.NewIndex) = Val(Nvl(rsTemp!项目ID)): .ListIndex = .NewIndex
    End With
    With cboDoctor
        .AddItem Nvl(rsTemp!医生姓名): .ItemData(.NewIndex) = Val(Nvl(rsTemp!医生ID)): .ListIndex = .NewIndex
    End With
    If Nvl(rsTemp!周日) <> Nvl(rsTemp!周一) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周二) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周三) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周四) _
        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周五) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周六) Then
        '每周
        opt周.Value = True
        With vsPlan
            For i = 0 To 4
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("周" & Replace(.ColKey(i), "日日", "日")))  '不知什么原因,将.colkey(i)的日,要更改成日日了.
            Next
        End With
        With vsPlan1
            For i = 0 To 1
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("周" & Replace(.ColKey(i), "日日", "日")))  '不知什么原因,将.colkey(i)的日,要更改成日日了.
            Next
        End With
    Else
        '每天
        opt天.Value = True:  cbo天.ListIndex = cbo.FindIndex(cbo天, Nvl(rsTemp!周日), True): cbo天.Enabled = True
    End If
    '有效时间范围
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!开始时间) Then
        chk有效期.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!开始时间, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!终止时间) Then
            dtpEnd.Value = CDate(Format(rsTemp!终止时间, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If
        
    Select Case Val(Nvl(rsTemp!分诊方式))     '0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
        Case 0  '"不分诊"
            opt分诊(0).Value = True
        Case 1  ' "指定诊室"
            opt分诊(1).Value = True
        Case 2 '"动态分诊"
            opt分诊(2).Value = True
        Case 3 ' "平均分诊"
            opt分诊(3).Value = True
    End Select
    
    strSQL = "Select 号表ID,门诊诊室　From 挂号安排诊室 Where 号表ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    Dim objItem As ListItem
    
    lvwDept.ListItems.Clear: i = 1
    Do While Not rsTemp.EOF
       Set objItem = lvwDept.ListItems.Add(, "K" & i, Nvl(rsTemp!门诊诊室))
        objItem.Checked = True
        i = i + 1
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '加载该挂号项目的的停用时间信息
    strSQL = "Select 安排ID,序号,开始停止时间,结束停止时间,制订人,制订日期,备注 From 挂号安排停用状态 where 安排ID=[1] Order by 开始停止时间,制订日期"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    With vsList
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("序号")) = i
            .Cell(flexcpData, i, .ColIndex("序号")) = Val(Nvl(rsTemp!序号))
            .TextMatrix(i, .ColIndex("开始停用时间")) = Format(rsTemp!开始停止时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("结束停用时间")) = Format(rsTemp!结束停止时间, "yyyy-mm-dd HH:MM")
            .TextMatrix(i, .ColIndex("制订人")) = Nvl(rsTemp!制订人)
            .TextMatrix(i, .ColIndex("制订日期")) = Format(rsTemp!制订日期, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("备注")) = Nvl(rsTemp!备注)
            If Format(rsTemp!结束停止时间, "yyyy-mm-dd HH:MM:SS") < strCurDate Then
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = Me.BackColor
            Else
                .Cell(flexcpForeColor, i, 1, i, .Cols - 1) = &H8000000C
            End If
            .RowData(i) = 1
            i = i + 1
            rsTemp.MoveNext
        Loop
       zl_vsGrid_Para_Restore mlngModule, vsList, Me.Caption, "停用安排-停用计划", True, InStr(1, mstrPrivs, ";参数设置;") > 0
    End With
    If blnRestore = False Then
        vsOthers.Clear 1
        vsOthers.Rows = 2
       zl_vsGrid_Para_Restore mlngModule, vsOthers, Me.Caption, "停用安排-挂号安排", True, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    
    '问题:43148
    gstrSQL = " Select  名称   From 安排停用原因 where 缺省标志=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If Not rsTemp.EOF Then txtMemo.Text = Nvl(rsTemp!名称)
    rsTemp.Close
    Set rsTemp = Nothing
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdAdd_Click()
    Dim lngRow As Long
    If CheckStopValied = False Then Exit Sub
    With vsList
        If .TextMatrix(.Row, .ColIndex("开始停用时间")) <> "" Then
            .Rows = .Rows + 1
            .Row = .Rows - 1: lngRow = .Row
        Else
            lngRow = .Row
        End If
        .RowData(lngRow) = 0
        .TextMatrix(lngRow, .ColIndex("序号")) = lngRow
        .TextMatrix(lngRow, .ColIndex("开始停用时间")) = Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM")
        .TextMatrix(lngRow, .ColIndex("结束停用时间")) = Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM")
        .TextMatrix(lngRow, .ColIndex("备注")) = Trim(txtMemo.Text)
    End With
End Sub
Private Sub SetCmdEnable()
    '设置按钮控件的Enabled属性
    With vsPlan
        If Trim(.TextMatrix(.Row, .ColIndex("开始停用时间"))) <> "" Then
            cmdDeleteTime.Enabled = True
        Else
            cmdDeleteTime.Enabled = False
        End If
    End With
    
End Sub

Private Sub cmdClear_Click()
    If MsgBox("注意:" & vbCrLf & "  你是否全清所有的挂号安排?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    vsOthers.Clear 1
    vsOthers.Rows = 2
    vsOthers.Cell(flexcpData, 1, 0, 1, vsOthers.Cols - 1) = ""
End Sub

Private Sub cmdDel_Click()
        '删除行
        Dim lngRow As Long
        With vsOthers
            If .TextMatrix(.Row, .ColIndex("号别")) <> "" Then
                If MsgBox("注意:" & vbCrLf & " 你是否真的要移除号别为『" & .TextMatrix(.Row, .ColIndex("号别")) & "』的挂号安排吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If .Rows - 1 <= 1 Then
                .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            Else
                lngRow = .Row
                .RemoveItem lngRow
                
                If lngRow > .Rows - 1 Then
                    .Row = lngRow - 1
                Else
                    '.Row = lngRow + 1
                End If
            End If
        End With
End Sub


Private Sub cmdDeleteTime_Click()
    '删除行
    Dim lngRow As Long
    With vsList
        If .TextMatrix(.Row, .ColIndex("开始停用时间")) <> "" Then
            If MsgBox("注意:" & vbCrLf & " 你是否真的要移除时间范围为" & vbCrLf & .TextMatrix(.Row, .ColIndex("开始停用时间")) & "～" & .TextMatrix(.Row, .ColIndex("结束停用时间")) & vbCrLf & "的停用计划安排吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("序号"))) <> 0 Then
                mstrDelete序号 = mstrDelete序号 & "," & Val(.Cell(flexcpData, .Row, .ColIndex("序号")))
        End If
        
        If .Rows - 1 <= 1 Then
            .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
            .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        Else
            lngRow = .Row
            .RemoveItem lngRow
            If lngRow > .Rows - 1 Then
                .Row = lngRow - 1
            Else
                '.Row = lngRow + 1
            End If
        End If
    End With
    Call ReFreshNo
End Sub
Private Sub ReFreshNo()
    '重新刷新序号
    Dim i As Long
    With vsList
        For i = 1 To .Rows - 1
            If Not .RowHidden(i) Then
                .TextMatrix(i, .ColIndex("序号")) = i
            End If
        Next
    End With
End Sub
Public Function GetSplitStrUnionTable(ByVal strInputSplit As String, ByVal blnNum As Boolean, _
    ByVal intBandStart As Integer, ByVal strNotSplitTable As String, ByVal strNotSplitFieldName As String, _
    ByRef OutSplitValue() As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定值,分解成相关的表和数据
    '入参:intBandStart-绑定的启始数
    '出参:OutSplitValue:返回0-10的值,未分配完的,直接填成IN方式
    '       strNotSplitValue:未分配完时,返回值
    '返回:返回SQL
    '编制:刘兴洪
    '日期:2010-09-08 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, i As Long, j As Long, strTemp As String
    Dim strSubItem  As String
    strSubItem = ""
    '先分解出来,再查找
    varData = Split(strInputSplit, ",")
    j = intBandStart: strTemp = ""
    For i = 0 To UBound(varData)
        If Len(strTemp) > 1990 And j - intBandStart <= 10 Then
            OutSplitValue(j - intBandStart) = Mid(strTemp, 2)
            If blnNum Then
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Num2List([" & j & "]))   "
            Else
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Str2List([" & j & "]))  "
            End If
            j = j + 1: strTemp = ""
        End If
        strTemp = strTemp & "," & IIf(blnNum, Val(varData(i)), varData(i))
    Next
    
    If strTemp <> "" Then
        If j - intBandStart > 10 Then
            If blnNum Then
                strSubItem = strSubItem & vbCrLf & " UNION ALL Select ID From " & strNotSplitTable & " Where " & strNotSplitFieldName & " in (" & Mid(strTemp, 2) & ")"
            Else
                strTemp = "'" & Replace(Mid(strTemp, 2), ",", "','") & "'"
                strSubItem = strSubItem & vbCrLf & " UNION ALL Select ID From " & strNotSplitTable & " Where " & strNotSplitFieldName & " in (" & strTemp & ")"
            End If
        Else
            OutSplitValue(j - intBandStart) = Mid(strTemp, 2)
            If blnNum Then
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Num2List([" & j & "]))  "
            Else
                 strSubItem = strSubItem & vbCrLf & " Union ALL " & _
                " Select Column_Value From Table(f_Str2List([" & j & "]))   "
            End If
        End If
    End If
    If strSubItem <> "" Then strSubItem = Mid(strSubItem, 13)
    GetSplitStrUnionTable = strSubItem
End Function
Private Sub cmdOthers_Click()
    Dim strType  As String, strDept   As String, str项目   As String, str医生 As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, strTable As String, strTemp As String
    Dim strVarDept(0 To 10) As String, strVar项目(0 To 10) As String
    Dim strVar医生(0 To 10) As String, strVar医生1(0 To 10) As String
    Dim i As Long, lngRow As Long, blnFind As Boolean, blnNotMsg As Boolean
    Dim lngCount As Long
    Dim varData As Variant
    If frmRegistPlanInvalidationCons.ShowCons(Me, mlngModule, mstrPrivs, strType, strDept, str项目, str医生) = False Then
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strTable = "": strWhere = ""
    If strType <> "" Then
        strTable = strTable & ",(" & " Select Column_Value as 号类 From Table(f_Str2List([1]))) J "
        strWhere = strWhere & " And A.号类=J.号类"
    End If
    If strDept <> "" Then
        strTemp = GetSplitStrUnionTable(strDept, True, 2, "部门表", "ID", strVarDept)
        strTable = strTable & vbCrLf & ",(" & strTemp & ") M "
        strWhere = strWhere & " And A.科室ID=M.Column_Value"
    End If
    
    If str项目 <> "" Then
        strTemp = GetSplitStrUnionTable(str项目, True, 13, "收费项目目录", "ID", strVar项目)
        strTable = strTable & vbCrLf & ",(" & strTemp & ") Q "
        strWhere = strWhere & " And A.项目ID=Q.Column_Value"
    End If
    If str医生 <> "" Then
        varData = Split(str医生, "||")
        For i = 0 To UBound(varData)
            If i = 0 Then
                strTemp = GetSplitStrUnionTable(varData(i), True, 24, "人员表", "ID", strVar医生)
                strTable = strTable & vbCrLf & ",(" & strTemp & ") H "
                strWhere = strWhere & " And (A.医生ID=H.Column_Value  "
            ElseIf i = 1 Then '院外医生
                strTemp = GetSplitStrUnionTable(varData(i), False, 24, "挂号安排", "医生姓名", strVar医生1)
                strTable = strTable & vbCrLf & ",(" & strTemp & ") M "
                strWhere = strWhere & " Or A.医生姓名=M.Column_Value  "
            End If
        Next
        If strWhere <> "" Then strWhere = strWhere & ")"
    End If
     

  
   strSQL = "" & _
    "   Select /*+ rule */ A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id,  F.限号数,  F.限约数,   " & _
    "           A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六, " & _
    "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
    "   From 挂号安排 A,收费项目目录 B,部门表 D,挂号安排限制 F " & strTable & _
    "   Where  A.项目id=b.Id(+) And A.科室id =d.Id(+) And a.id=F.安排ID(+) And " & vbNewLine & _
    "          Decode(To_Char(Sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =f.限制项目(+)" & vbNewLine & _
    "          And A.ID <>" & mlng安排ID & strWhere & _
    "   Order by A.号类,A.号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strType, _
        strVarDept(0), strVarDept(1), strVarDept(2), strVarDept(3), strVarDept(4), strVarDept(5), strVarDept(6), strVarDept(7), strVarDept(8), strVarDept(9), strVarDept(10), _
        strVar项目(0), strVar项目(1), strVar项目(2), strVar项目(3), strVar项目(4), strVar项目(5), strVar项目(6), strVar项目(7), strVar项目(8), strVar项目(9), strVar项目(10), _
        strVar医生(0), strVar医生(1), strVar医生(2), strVar医生(3), strVar医生(4), strVar医生(5), strVar医生(6), strVar医生(7), strVar医生(8), strVar医生(9), strVar医生(10), _
        strVar医生1(0), strVar医生1(1), strVar医生1(2), strVar医生1(3), strVar医生1(4), strVar医生1(5), strVar医生1(6), strVar医生1(7), strVar医生1(8), strVar医生1(9), strVar医生1(10), _
        "")
    With rsTemp
        blnNotMsg = False
        lngCount = .RecordCount
        Do While Not .EOF
                With vsOthers
                    blnFind = False
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, .ColIndex("ID"))) = Val(Nvl(rsTemp!安排ID)) Then
                            .Row = i
                            If Not blnNotMsg Then
                                If lngCount > 1 Then
                                    If MsgBox("注意:" & vbCrLf & "    号码『" & .TextMatrix(i, .ColIndex("号别")) & "』已经存在," & vbCrLf & _
                                                "此号别将不再加入,如果存在相同情况,是否不再提示?" & vbCrLf & _
                                                "『是』表示如果还存在重复的号别,则不再提示。" & vbCrLf & _
                                                "『否』表示如果还存在重复的号别，则继续提示。", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then
                                            blnNotMsg = True
                                    End If
                                Else
                                    Call MsgBox("注意:" & vbCrLf & "    号码『" & .TextMatrix(i, .ColIndex("号别")) & "』已经存在,不能再加入", vbInformation + vbDefaultButton1, gstrSysName)
                                End If
                            End If
                            
                            '加载数据
                            blnFind = True: Exit For
                        End If
                    Next
                    If blnFind = False Then
                       If .TextMatrix(.Rows - 1, .ColIndex("ID")) <> "" Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!安排ID)
                    .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsTemp!号类)
                    .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTemp!号码)
                    .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
                    .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsTemp!项目)
                    .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsTemp!医生姓名)
                    .TextMatrix(lngRow, .ColIndex("限号")) = Nvl(rsTemp!限号数)
                    .TextMatrix(lngRow, .ColIndex("限约")) = Nvl(rsTemp!限约数)
                    .TextMatrix(lngRow, .ColIndex("周日")) = Nvl(rsTemp!周日)
                    .TextMatrix(lngRow, .ColIndex("周一")) = Nvl(rsTemp!周一)
                    .TextMatrix(lngRow, .ColIndex("周二")) = Nvl(rsTemp!周二)
                    .TextMatrix(lngRow, .ColIndex("周三")) = Nvl(rsTemp!周三)
                    .TextMatrix(lngRow, .ColIndex("周四")) = Nvl(rsTemp!周四)
                    .TextMatrix(lngRow, .ColIndex("周五")) = Nvl(rsTemp!周五)
                    .TextMatrix(lngRow, .ColIndex("周六")) = Nvl(rsTemp!周六)
                    .TextMatrix(lngRow, .ColIndex("建病案")) = IIf(Val(Nvl(rsTemp!病案必须)) = 0, "", "√")
                    .TextMatrix(lngRow, .ColIndex("分诊方式")) = Nvl(rsTemp!分诊方式)
                    .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!科室ID) & "_" & Nvl(rsTemp!项目ID) & "_" & Nvl(rsTemp!医生ID)
                    .TextMatrix(lngRow, .ColIndex("应诊诊室")) = Read安排应诊诊室(Val(Nvl(rsTemp!安排ID)))    ' Nvl(rsTemp!门诊诊室)
                    
                    If Not IsNull(rsTemp!开始时间) Then
                        .TextMatrix(lngRow, .ColIndex("有效范围")) = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & _
                            "至" & Format(rsTemp!终止时间, "yyyy-MM-dd HH:mm:ss")
                        .TextMatrix(lngRow, .ColIndex("有效范围")) = Replace(.TextMatrix(lngRow, .ColIndex("有效范围")), " 00:00:00", "")
                    End If
                    .TextMatrix(lngRow, .ColIndex("序号控制")) = IIf(Val(Nvl(rsTemp!序号控制)) = 0, "", "√")
                    .Row = lngRow
                    lngRow = lngRow + 1
                    End If
                End With
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRestore_Click()
    If MsgBox("注意:" & vbCrLf & "   执行恢复功能后,将会取消当前的设置,是否继续?", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    Call LoadData(True)
End Sub

Private Sub cmdSel_Click()
    If SelectStopMemo(txtMemo, "") = False Then Exit Sub
End Sub

Private Sub Form_Load()
    Call InitPanel
    Call RestoreWinState(Me, App.ProductName)
    mblnFirst = True
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadData = False Then Unload Me: Exit Sub
    
    Call SetCtrlEnabled
    zlControl.ControlSetFocus dtpStartDate
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub
Private Sub SetCtrlEnabled()
    '设置控件的Enabled属性
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            If ctl Is txtMemo Or ctl Is txtCode Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
            zlSetCtrolBackColor ctl
        Case UCase("ComboBox")
            ctl.Enabled = False
            zlSetCtrolBackColor ctl
        Case UCase("ListView")
            ctl.Enabled = False
            zlSetCtrolBackColor ctl
        Case UCase("DTPicker")
            If ctl Is dtpStartDate Or ctl Is dtpEndDate Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
           
        Case UCase("optionbutton"), UCase("CheckBox")
            If ctl Is chkClearHistory Then
                ctl.Enabled = True
            Else
                ctl.Enabled = False
            End If
        Case Else
        End Select
    Next
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function CheckStopValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查安排时间停用的合法性
    '返回:合法,返回True,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 14:06:13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If dtpStartDate.Value > dtpEndDate.Value Then
        ShowMsgbox "注意:" & vbCrLf & "    开始停用日期大于了结束停用日期,请检查!"
        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
        Exit Function
    End If
    
    If dtpStartDate.Value < zlDatabase.Currentdate Then
        ShowMsgbox "注意:" & vbCrLf & "    开始停用日期小于了当前系统时间,请检查!"
        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
        Exit Function
    End If
    If zlCommFun.ActualLen(txtMemo.Text) > 100 Then
        ShowMsgbox "注意:" & vbCrLf & "   备注输入过长，只能输入100个字符或50个汉字,请检查!"
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    If InStr(1, txtMemo.Text, "'") > 0 Then
        ShowMsgbox "注意:" & vbCrLf & "   备注不能输入单引号,请检查!"
        If txtMemo.Enabled And txtMemo.Visible Then txtMemo.SetFocus
        Exit Function
    End If
    '检查表格中的日期是否合法:
    With vsList
        For i = 1 To .Rows - 1
            If Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM") >= Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) _
               And Format(dtpStartDate.Value, "yyyy-mm-dd HH:MM") <= .TextMatrix(i, .ColIndex("结束停用时间")) Then
               ShowMsgbox "注意:" & vbCrLf & "    开始停用时间已经在第" & i & "行中存在,请检查!"
               If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
               Exit Function
            End If
            If Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM") >= Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) _
               And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM") <= .TextMatrix(i, .ColIndex("结束停用时间")) Then
               ShowMsgbox "注意:" & vbCrLf & "    结束停用时间已经在第" & i & "行中存在,请检查!"
               If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
               Exit Function
            End If
            If Format(Trim(.TextMatrix(i, .ColIndex("开始停用时间"))), "yyyy-mm-dd hh:mm") <= Format(dtpStartDate.Value, "yyyy-mm-dd HH:mm") Then
                If Format(Trim(.TextMatrix(i, .ColIndex("结束停用时间"))), "yyyy-mm-dd hh:mm") >= Format(dtpStartDate.Value, "yyyy-mm-dd HH:mm") Then
                    ShowMsgbox "注意:" & vbCrLf & "    第" & i & "行中的停用时间范围已经包含在当前所设置的停用范围中,请检查!"
                    If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
                    Exit Function
                End If
            Else
                If Format(Trim(.TextMatrix(i, .ColIndex("开始停用时间"))), "yyyy-mm-dd hh:mm") <= Format(dtpEndDate.Value, "yyyy-mm-dd HH:mm") Then
                    ShowMsgbox "注意:" & vbCrLf & "    第" & i & "行中的停用时间范围已经包含在当前所设置的停用范围中,请检查!"
                    If dtpEnd.Enabled And dtpEnd.Visible Then dtpEnd.SetFocus
                    Exit Function
                End If
            End If
        Next
    End With
     CheckStopValied = True
End Function
Private Function CheckOtherPlan(ByVal str安排ID As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查应用于其他安排的日期是否合法
    '入参:str安排ID-多个安排时,用逗号分隔
    '出参:
    '返回:合法,返回true, 否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 14:22:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strValues(0 To 10) As String, strSubItem As String, varData As Variant
    Dim strTemp As String, i As Long, j As Long, strEndDate As String, strStartDate As String
    Dim strValue(0 To 10)  As String
    On Error GoTo errHandle
     '先分解出来,再查找
    varData = Split(str安排ID, ",")
    strTemp = "": j = 1
    For i = 0 To UBound(varData)
        If Len(strTemp) > 1990 And j <= 10 Then
            strValue(j - 1) = Mid(strTemp, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 安排ID From Table(f_Num2List([" & j & "])) B "
            strTemp = "," & Val(varData(i)): j = j + 1
        Else
            strTemp = strTemp & "," & Val(varData(i))
        End If
    Next
    
    If strTemp <> "" Then
        If j - 1 > 10 Then
             strSubItem = strSubItem & " UNION ALL Select ID From 挂号安排停用状态 Where 安排ID in (" & Mid(strTemp, 2) & ")"
        Else
            strValue(j - 1) = Mid(strTemp, 2)
            strSubItem = strSubItem & " Union ALL " & _
            " Select Column_Value as 安排ID From Table(f_Num2List([" & j & "])) B "
        End If
    End If
    strSQL = "" & _
       "   Select /*+ Rule*/ B.号别,A.开始停止时间,A.结束停止时间  " & _
       "   From 挂号安排停用状态 A,挂号安排 B, (" & Mid(strSubItem, 11) & ") D" & _
       "   Where A.安排ID = D.安排ID and A.安排ID=b.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str安排ID, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    With rsTemp
        Do While Not .EOF
            strStartDate = Format(rsTemp!开始停止时间, "yyyy-mm-dd HH:MM")
            strEndDate = Format(rsTemp!结束停止时间, "yyyy-mm-dd HH:MM")
            '检查表格中的日期是否合法:
            With vsList
                For i = 1 To .Rows - 1
                    If Val(.RowData(i)) <> 1 Then
                        If Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) >= strStartDate _
                           And .TextMatrix(i, .ColIndex("开始停用时间")) <= strEndDate Then
                           ShowMsgbox "注意:" & vbCrLf & "    号别为『" & Nvl(rsTemp!号别) & "』的停用时间(" & Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) & " ~ " & Trim(.TextMatrix(i, .ColIndex("结束停用时间"))) & ")已经存在,请检查!"
                           .Row = i: .Col = .ColIndex("开始停用时间")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        If Trim(.TextMatrix(i, .ColIndex("结束停用时间"))) >= strStartDate _
                           And .TextMatrix(i, .ColIndex("结束停用时间")) <= strEndDate Then
                           ShowMsgbox "注意:" & vbCrLf & "    结束停用时间已经在第" & i & "行中存在,请检查!"
                           .Row = i: .Col = .ColIndex("结束停用时间")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        
                        If strStartDate >= Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) And _
                          strStartDate >= Trim(.TextMatrix(i, .ColIndex("结束停用时间"))) Then
                           ShowMsgbox "注意:" & vbCrLf & "    开始停用时间已经在第" & i & "行中存在,请检查!"
                           .Row = i: .Col = .ColIndex("结束停用时间")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                        If strEndDate >= Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) And _
                          strEndDate >= Trim(.TextMatrix(i, .ColIndex("结束停用时间"))) Then
                           ShowMsgbox "注意:" & vbCrLf & "    结束停用时间已经在第" & i & "行中存在,请检查!"
                           .Row = i: .Col = .ColIndex("结束停用时间")
                           If vsList.Enabled And vsList.Visible Then vsList.SetFocus
                           Exit Function
                        End If
                    End If
                Next
            End With
            .MoveNext
        Loop
    End With
    CheckOtherPlan = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存所有的数据
    '返回:成功,返回true,否则返回false
    '编制:刘兴洪
    '日期:2010-09-07 14:54:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String, i As Long, j As Long, str安排ID As String
    Dim strStartDate As String, strEndDate As String, str备注 As String
    Dim cll安排 As Collection
    
    Set cllPro = New Collection: Set cll安排 = New Collection
    With vsOthers
        str安排ID = "," & mlng安排ID
        For j = 1 To .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("ID"))) <> 0 Then
                If Len(str安排ID) > 1920 Then
                    str安排ID = Mid(str安排ID, 2)
                   cll安排.Add str安排ID
                    str安排ID = ""
                End If
                str安排ID = str安排ID & "," & Val(.TextMatrix(j, .ColIndex("ID")))
            End If
        Next
        If str安排ID <> "" Then
            str安排ID = Mid(str安排ID, 2)
            cll安排.Add str安排ID
        End If
    End With
    
    With vsList
        '先处理删除数据
        If mstrDelete序号 <> "" Then
            mstrDelete序号 = Mid(mstrDelete序号, 2)
            For j = 1 To cll安排.Count
                'Zl_挂号安排停用状态_Delete
                strSQL = "Zl_挂号安排停用状态_Delete("
                '  安排id_In     In 挂号安排停用状态.安排id%Type,
                strSQL = strSQL & "" & mlng安排ID & ","
                '  序号_In       In Varchar2, --用逗号分隔
                strSQL = strSQL & "'" & mstrDelete序号 & "',"
                '  其他安排id_In In Varchar2 --用逗号分隔
                strSQL = strSQL & "'" & cll安排(j) & "')"
                zlAddArray cllPro, strSQL
            Next
 
        End If
        '处理增加的数据
        For i = 1 To .Rows - 1
            If Val(.RowData(i)) = 0 And Trim(.TextMatrix(i, .ColIndex("开始停用时间"))) <> "" Then '0代表本次增加的停用日期
                str安排ID = "," & mlng安排ID
                strStartDate = .TextMatrix(i, .ColIndex("开始停用时间"))
                strEndDate = .TextMatrix(i, .ColIndex("结束停用时间"))
                str备注 = .TextMatrix(i, .ColIndex("备注"))
                For j = 1 To cll安排.Count
                    If chkClearHistory.Value = 1 Then
                        '    Zl_挂号安排停用状态_Clear(安排id_In Varchar2) Is
                        strSQL = "Zl_挂号安排停用状态_Clear('" & cll安排(j) & "')"
                        zlAddArray cllPro, strSQL
                    End If
                     'Zl_挂号安排停用状态_Insert
                     strSQL = "Zl_挂号安排停用状态_Insert("
                    '开始停止时间_In In 挂号安排停用状态.开始停止时间%Type,
                    strSQL = strSQL & "to_date('" & strStartDate & "','yyyy-mm-dd HH24:mi'),"
                     '结束停止时间_In In 挂号安排停用状态.结束停止时间%Type,
                    strSQL = strSQL & "to_date('" & strEndDate & "','yyyy-mm-dd HH24:mi'),"
                     '制订人_In       In 挂号安排停用状态.制订人%Type,
                    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                     '备注_In         In 挂号安排停用状态.备注%Type,
                    strSQL = strSQL & "'" & str备注 & "',"
                     '安排id_In       In Varchar2 --用逗号分隔
                    strSQL = strSQL & "'" & cll安排(j) & "')"
                    zlAddArray cllPro, strSQL
                Next
            End If
        Next
    End With
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function IsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否合法
    '入参:
    '出参:
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 16:03:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, blnFind As Boolean
    On Error GoTo errHandle
    With vsList
        For lngRow = 1 To .Rows - 1
             If .TextMatrix(lngRow, .ColIndex("开始停用时间")) <> "" And InStr("0,2", Val(.RowData(lngRow))) > 0 Then
                    blnFind = True: Exit For
             End If
        Next
        If blnFind = False And mstrDelete序号 = "" Then
            MsgBox "没有加入停用时间，不能继续!", vbInformation + vbOKOnly, gstrSysName
            If dtpStartDate.Enabled And dtpStartDate.Visible Then dtpStartDate.SetFocus
            Exit Function
        End If
    End With
    IsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOk_Click()
     If IsValied = False Then Exit Sub
     If SaveData = False Then Exit Sub
     mblnSucces = True
    Unload Me
End Sub
Public Function SelectItem(ByVal strInput As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据输入的值，选择相关的数据(存在多选)
    '入参:intIndex-索引
    '       strInput-输入的值
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-09-07 10:21:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strCode As String, blnCancel As Boolean, rsTemp As ADODB.Recordset
    Dim strDept As String, strDeptWhere As String, strTable As String
    Dim strLike As String, strWhere As String, bytCode As Byte, lngRow As Long
    Dim strTittle As String, strValue(0 To 10) As String
    Dim vRect As RECT, j As Long, i As Long
    If Trim(strInput) = "" Then Exit Function
     On Error GoTo Hd
    bytCode = Val(zlDatabase.GetPara("简码方式", , , 0))
    strLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    '功能：多功能选择器,使用ADO.Command打开,允许使用[x]参数
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '     arrInput=对应的各个SQL参数值,按顺序传入,必须为明确类型
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等
    If strInput <> "" Then
        strCode = strLike & strInput & "%"
        If zlCommFun.IsCharAlpha(strInput) Then
            strWhere = " And ( B.简码 Like upper([1]))"
        ElseIf IsNumeric(strInput) Or zlCommFun.IsNumOrChar(strInput) Then
            strWhere = " And A.号码 Like upper([1])"
        Else
            strWhere = " And (A.医生姓名 Like [1] Or exists(Select 1 From 人员表 where 姓名 like [1] and a.医生ID=C.id )  or B.名称 like [1] or  A.号类 like [1]  )"
        End If
    Else
        strWhere = ""
    End If

    strSQL = "" & _
    "   Select Distinct A.id, A.号类,A.号码,b.名称 as 科室,C.名称 as 项目,nvl(D.姓名,A.医生姓名) as 医生 " & _
    "   From 挂号安排 A,部门表 B,收费项目目录 C,人员表 D" & _
    "   Where A.科室ID=B.id  And A.项目ID=C.id and A.医生ID=D.id(+) And A.ID <>" & mlng安排ID & strWhere & _
    "           And rownum<101 " & _
    "   Order by 号类,号码"
    
    strTittle = "挂号安排"
    
    vRect = zlControl.GetControlRect(txtCode.Hwnd)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, strTittle & "选择", False, "", "请选择", False, False, True, vRect.Left, vRect.Top, txtCode.Height, blnCancel, True, True, strCode)
    
    
    If blnCancel = True Then
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        MsgBox "没有找到满足条件的" & strTittle & "，请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Call txtCode_GotFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        MsgBox "没有找到满足条件的" & strTittle & "，请检查!", vbInformation + vbOKOnly, gstrSysName
        If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
        Call txtCode_GotFocus
        Exit Function
    End If
    Dim strValues As String, strSubItem As String, blnFind As Boolean, blnNotMsg As Boolean, lngCount As Long
    
    
    With rsTemp
        strValues = "": j = 1: lngCount = .RecordCount
        blnNotMsg = False
        Do While Not .EOF
            '先检查是否存在在网格中已经存在了
            With vsOthers
                blnFind = False
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ID"))) = Val(Nvl(rsTemp!ID)) Then
                            If Not blnNotMsg Then
                                If lngCount > 1 Then
                                    If MsgBox("注意:" & vbCrLf & "    号码『" & .TextMatrix(i, .ColIndex("号别")) & "』已经存在," & vbCrLf & _
                                                "此号别将不再加入,如果存在相同情况,是否不再提示?" & vbCrLf & _
                                                "『是』表示如果还存在重复的号别,则不再提示。" & vbCrLf & _
                                                "『否』表示如果还存在重复的号别，则继续提示。", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName) = vbYes Then
                                            blnNotMsg = True
                                    End If
                                Else
                                    Call MsgBox("注意:" & vbCrLf & "    号码『" & .TextMatrix(i, .ColIndex("号别")) & "』已经存在,不能再加入", vbInformation + vbDefaultButton1, gstrSysName)
                                    .Row = i
                                End If
                            End If
                           blnFind = True: Exit For
                    End If
                Next
            End With
            
            If blnFind = False Then
                If Len(strValues) > 1990 And j <= 10 Then
                    strValue(j - 1) = Mid(strValues, 2)
                    strSubItem = strSubItem & " Union ALL " & _
                    " Select Column_Value as 安排ID From Table(f_Num2List([" & j & "])) B "
                    strValues = "," & Val(Nvl(rsTemp!ID)): j = j + 1
                Else
                    strValues = strValues & "," & Val(Nvl(rsTemp!ID))
                End If
            End If
            .MoveNext
        Loop
        If strValues <> "" Then
            If j - 1 > 10 Then
                 strSubItem = strSubItem & " UNION ALL Select ID From 挂号安排 Where ID in (" & Mid(strValues, 2) & ")"
            Else
                strValue(j - 1) = Mid(strValues, 2)
                strSubItem = strSubItem & " Union ALL " & _
                "   Select Column_Value as 安排ID From Table(f_Num2List([" & j & "])) B "
            End If
        End If
    End With
    If strSubItem = "" Then Exit Function
    
    strSQL = "" & _
        "   Select /*+ rule */ A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id, F.限号数,  F.限约数,   " & _
        "           A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六, " & _
        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室  " & _
        "   From 挂号安排 A,收费项目目录 B,部门表 D , 挂号安排限制 F,(" & Mid(strSubItem, 11) & ") M" & _
        "   Where  A.项目id=b.Id(+) And A.科室id =d.Id(+) And A.id=M.安排ID  And a.Id = f.安排id(+) And" & _
        "   Decode(To_Char(sysdate, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五', '7', '周六', Null) =f.限制项目(+)" & vbNewLine & _
        "   Order by A.号类,A.号码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strValue(0), strValue(1), strValue(2), strValue(3), strValue(4), strValue(5), strValue(6), strValue(7), strValue(8), strValue(9), strValue(10))
    With rsTemp
        Do While Not .EOF
                With vsOthers
                    If .TextMatrix(.Rows - 1, .ColIndex("ID")) <> "" Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!安排ID)
                    .TextMatrix(lngRow, .ColIndex("号类")) = Nvl(rsTemp!号类)
                    .TextMatrix(lngRow, .ColIndex("号别")) = Nvl(rsTemp!号码)
                    .TextMatrix(lngRow, .ColIndex("科室")) = Nvl(rsTemp!科室)
                    .TextMatrix(lngRow, .ColIndex("项目")) = Nvl(rsTemp!项目)
                    .TextMatrix(lngRow, .ColIndex("医生")) = Nvl(rsTemp!医生姓名)
                    .TextMatrix(lngRow, .ColIndex("限号")) = Nvl(rsTemp!限号数)
                    .TextMatrix(lngRow, .ColIndex("限约")) = Nvl(rsTemp!限约数)
                    .TextMatrix(lngRow, .ColIndex("周日")) = Nvl(rsTemp!周日)
                    .TextMatrix(lngRow, .ColIndex("周一")) = Nvl(rsTemp!周一)
                    .TextMatrix(lngRow, .ColIndex("周二")) = Nvl(rsTemp!周二)
                    .TextMatrix(lngRow, .ColIndex("周三")) = Nvl(rsTemp!周三)
                    .TextMatrix(lngRow, .ColIndex("周四")) = Nvl(rsTemp!周四)
                    .TextMatrix(lngRow, .ColIndex("周五")) = Nvl(rsTemp!周五)
                    .TextMatrix(lngRow, .ColIndex("周六")) = Nvl(rsTemp!周六)
                    .TextMatrix(lngRow, .ColIndex("建病案")) = IIf(Val(Nvl(rsTemp!病案必须)) = 0, "", "√")
                    .TextMatrix(lngRow, .ColIndex("分诊方式")) = Nvl(rsTemp!分诊方式)
                    .TextMatrix(lngRow, .ColIndex("IDS")) = Nvl(rsTemp!科室ID) & "_" & Nvl(rsTemp!项目ID) & "_" & Nvl(rsTemp!医生ID)
                    .TextMatrix(lngRow, .ColIndex("应诊诊室")) = Read安排应诊诊室(Val(Nvl(rsTemp!安排ID)))    ' Nvl(rsTemp!门诊诊室)
                    
                    If Not IsNull(rsTemp!开始时间) Then
                        .TextMatrix(lngRow, .ColIndex("有效范围")) = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & _
                            "至" & Format(rsTemp!终止时间, "yyyy-MM-dd HH:mm:ss")
                        .TextMatrix(lngRow, .ColIndex("有效范围")) = Replace(.TextMatrix(lngRow, .ColIndex("有效范围")), " 00:00:00", "")
                    End If
                    .TextMatrix(lngRow, .ColIndex("序号控制")) = IIf(Val(Nvl(rsTemp!序号控制)) = 0, "", "√")
                   ' .TextMatrix(lngRow, .ColIndex("停用日期")) = Nvl(rsTemp!停用日期)
'                    If Trim(.TextMatrix(lngRow, .ColIndex("停用日期"))) <> "" Then
'                        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
'                    End If
                    .Row = lngRow
                    lngRow = lngRow + 1
                End With
            .MoveNext
        Loop
    End With
   Call txtCode_GotFocus
    If txtCode.Enabled And txtCode.Visible Then txtCode.SetFocus
    SelectItem = True
    Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function



Private Function Read安排应诊诊室(ByVal lngID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取指定诊室
    '入参:lngID-ID
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-14 22:39:14
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim strSQL As String
    
    On Error GoTo errH
    If lngID = 0 Then Exit Function
    
    If mrsRoom Is Nothing Then
        strSQL = "Select 门诊诊室,号表ID From 挂号安排诊室"
        Set mrsRoom = New Recordset
        Call zlDatabase.OpenRecordset(mrsRoom, strSQL, Me.Caption)
    End If
    With mrsRoom
        .Filter = "号表ID=" & lngID
        If .RecordCount = 0 Then Exit Function
        Do While Not .EOF
            Read安排应诊诊室 = Read安排应诊诊室 & ";" & !门诊诊室
            .MoveNext
        Loop
    End With
    Read安排应诊诊室 = Mid(Read安排应诊诊室, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "停用安排-停用计划", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "停用安排-挂号安排", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
    
End Sub

Private Sub imgColList_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = zlControl.GetControlRect(picImgList.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgList.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsList, lngLeft, lngTop, imgColList.Height)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "停用安排-挂号安排", True, , InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
Private Sub picImgList_Click()
    Call picImgList_Click
End Sub
 
Private Sub picCmd_Resize()
    Err = 0: On Error Resume Next:
    With picCmd
        fraSplit(1).Left = .ScaleLeft: fraSplit(1).Width = .ScaleWidth
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
        chkClearHistory.Left = cmdOK.Left - chkClearHistory.Width * 2
        If chkClearHistory.Left < 0 Then chkClearHistory.Left = .ScaleLeft
    End With
End Sub
Private Sub picOthers_Resize()
   Err = 0: On Error Resume Next:
    With picOthers
        vsOthers.Left = .ScaleLeft + 50: vsOthers.Width = .ScaleWidth - vsList.Left * 2
        vsOthers.Height = .ScaleHeight - vsOthers.Top
    End With
End Sub

Private Sub picStop_Resize()
   Err = 0: On Error Resume Next:
    With picStop
        vsList.Left = .ScaleLeft + 50: vsList.Width = .ScaleWidth - vsList.Left * 2
        vsList.Height = .ScaleHeight - vsList.Top
    End With
End Sub

Private Sub txtCode_GotFocus()
    zlControl.TxtSelAll txtCode
End Sub
Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If SelectItem(Trim(txtCode.Text)) = False Then
        Exit Sub
    End If
End Sub

Private Sub txtMemo_Change()
    txtMemo.Tag = ""
End Sub

Private Sub txtMemo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtMemo.Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtMemo.Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If SelectStopMemo(txtMemo, Trim(txtMemo.Text)) = False Then Exit Sub
End Sub

Private Sub vsList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "停用安排-停用计划", True, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsList, Me.Caption, "停用安排-停用计划", True, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsOthers_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "停用安排-挂号安排", True, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub

Private Sub vsOthers_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsOthers, Me.Caption, "停用安排-挂号安排", True, InStr(1, mstrPrivs, ";参数设置;") > 0
End Sub
Private Function SelectStopMemo(ByVal objCtl As Control, Optional strKey As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择停用原因
    '入参:strKey-输入值
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-11-08 15:00:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim vRect As RECT, strWhere As String, bytStyle As Byte
    Dim sngX As Single, sngY As Single, lngH As Long
    Dim blnCancel As Boolean
 
    On Error GoTo errH
    bytStyle = 0
    strWhere = " "
    If strKey <> "" Then
        strWhere = " Where 1=1 "
        If zlCommFun.IsCharChinese(strKey) Then
            strWhere = strWhere & " And 名称 like [1]  Order by 名称"
        ElseIf zlCommFun.IsCharAlpha(strKey) Then
            strWhere = strWhere & " And 简码 like upper([1]) Order by 简码"
        ElseIf zlCommFun.IsNumOrChar(strKey) Then
            strWhere = strWhere & " And 编码 like upper([1])  Order by 编码"
        Else
            strWhere = strWhere & " And  (名称 like [1] or 编码 like upper([1]) or 简码 like upper([1])) Order by 编码"
        End If
        bytStyle = 0
        strKey = gstrLike & strKey & "%"
    End If
    
    strSQL = "" & _
    "   Select Rownum as ID,编码,名称,简码,decode(缺省标志,1,'√','') as 缺省" & _
    "   From 安排停用原因" & _
        strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strKey)
    
    'ShowSelect:
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    vRect = zlControl.GetControlRect(objCtl.Hwnd)
    lngH = objCtl.Height
    sngX = vRect.Left - 15: sngY = vRect.Top
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, bytStyle, "停用原因选择", IIf(bytStyle = 2, True, False), "", "请选择符合条件的停用原因", IIf(bytStyle = 2, True, False), True, True, sngX, sngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgbox "不存在符合条件的停用原因,请检查!"
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
        zlControl.TxtSelAll objCtl
        Exit Function
        Exit Function
    End If
    With rsTemp
        objCtl.Text = Nvl(!名称): objCtl.Tag = Nvl(!ID)
    End With
    If objCtl.Enabled And objCtl.Visible Then objCtl.SetFocus
    zlControl.TxtSelAll objCtl
    zlCommFun.PressKey vbKeyTab
    SelectStopMemo = True
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
End Function


