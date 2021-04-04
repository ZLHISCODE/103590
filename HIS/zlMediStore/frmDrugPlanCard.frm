VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugPlanCard 
   Caption         =   "药品采购计划"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmDrugPlanCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk是否显示库存情况 
      Caption         =   "显示库存数量详情"
      Height          =   240
      Left            =   6600
      TabIndex        =   58
      Top             =   6360
      Width           =   1932
   End
   Begin VB.PictureBox pic库房 
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   6600
      ScaleHeight     =   2385
      ScaleWidth      =   3855
      TabIndex        =   53
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CheckBox chk库房 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   20
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1983
         Width           =   675
      End
      Begin VB.CommandButton cmd取消 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   2720
         TabIndex        =   55
         Top             =   1920
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   1560
         TabIndex        =   54
         Top             =   1920
         Width           =   1100
      End
      Begin MSComctlLib.ListView lvw存储库房 
         Height          =   1935
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   3413
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ImageList img16 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugPlanCard.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugPlanCard.frx":69AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugPlanCard.frx":D20E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.PictureBox picStock 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   2520
      ScaleHeight     =   1680
      ScaleWidth      =   8775
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   8775
      Begin VB.PictureBox picHeadStock 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         ScaleHeight     =   300
         ScaleWidth      =   8535
         TabIndex        =   46
         Tag             =   "0"
         Top             =   0
         Width           =   8535
         Begin VB.CheckBox chk所有库房 
            BackColor       =   &H00FFEDDD&
            Caption         =   "所有库房"
            Height          =   180
            Left            =   1080
            TabIndex        =   50
            Top             =   48
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chk来源库房 
            BackColor       =   &H00FFEDDD&
            Caption         =   "来源库房"
            Height          =   180
            Left            =   2400
            TabIndex        =   49
            Top             =   48
            Width           =   1095
         End
         Begin VB.CheckBox chk来源药房 
            BackColor       =   &H00FFEDDD&
            Caption         =   "来源药房"
            Height          =   180
            Left            =   3720
            TabIndex        =   48
            Top             =   48
            Width           =   1095
         End
         Begin VB.CommandButton cmd库房 
            Caption         =   "…"
            Height          =   285
            Left            =   5040
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   -4
            Width           =   285
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "库存详情"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   0
            TabIndex        =   52
            Top             =   48
            Width           =   720
         End
         Begin VB.Label lbl自定义库房 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "自定义库房"
            Height          =   180
            Left            =   5280
            TabIndex        =   51
            Top             =   45
            Width           =   975
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   1200
         Left            =   10
         TabIndex        =   45
         Top             =   330
         Width           =   8295
         _cx             =   14631
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
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDrugPlanCard.frx":13A70
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
   End
   Begin VB.CheckBox chk隐藏近期采购计划 
      Caption         =   "隐藏近期采购计划"
      Height          =   240
      Left            =   4440
      TabIndex        =   43
      Top             =   6360
      Width           =   1932
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   8
      Top             =   5850
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   7
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   6
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   4
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   5
      Top             =   5760
      Width           =   1100
   End
   Begin VB.PictureBox Pic单据 
      BackColor       =   &H80000004&
      Height          =   5325
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin VB.PictureBox picHis 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   240
         ScaleHeight     =   1755
         ScaleWidth      =   8775
         TabIndex        =   32
         Top             =   1800
         Width           =   8775
         Begin VB.PictureBox picHscSend 
            BackColor       =   &H00FFEDDD&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            ScaleHeight     =   300
            ScaleWidth      =   8535
            TabIndex        =   34
            Tag             =   "0"
            Top             =   0
            Width           =   8535
            Begin VB.PictureBox picColor 
               BackColor       =   &H000000FF&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   2
               Left            =   7080
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   41
               Top             =   45
               Width           =   255
            End
            Begin VB.PictureBox picColor 
               BackColor       =   &H008080FF&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   1
               Left            =   5640
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   39
               Top             =   45
               Width           =   255
            End
            Begin VB.PictureBox picColor 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   200
               Index           =   0
               Left            =   4200
               ScaleHeight     =   195
               ScaleWidth      =   255
               TabIndex        =   37
               Top             =   45
               Width           =   255
            End
            Begin VB.CheckBox chkMore 
               BackColor       =   &H00FFEDDD&
               Caption         =   "更多"
               Height          =   240
               Left            =   2880
               TabIndex        =   36
               Top             =   25
               Width           =   855
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "没有执行"
               Height          =   180
               Index           =   2
               Left            =   7440
               TabIndex        =   42
               Top             =   45
               Width           =   720
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "部分执行"
               Height          =   180
               Index           =   1
               Left            =   6000
               TabIndex        =   40
               Top             =   45
               Width           =   720
            End
            Begin VB.Label lblColorTxt 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "完全执行"
               Height          =   180
               Index           =   0
               Left            =   4560
               TabIndex        =   38
               Top             =   45
               Width           =   720
            End
            Begin VB.Image imgDown 
               Height          =   240
               Left            =   0
               Picture         =   "frmDrugPlanCard.frx":13B3A
               Top             =   0
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image imgUp 
               Height          =   240
               Left            =   0
               Picture         =   "frmDrugPlanCard.frx":13E7C
               Top             =   0
               Width           =   240
            End
            Begin VB.Label lblDiag 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFEDDD&
               Caption         =   "近期采购计划执行情况"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   480
               TabIndex        =   35
               Top             =   50
               Width           =   1800
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfHisPlane 
            Height          =   1400
            Left            =   0
            TabIndex        =   33
            Top             =   315
            Width           =   2760
            _cx             =   4868
            _cy             =   2469
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   0
            GridColorFixed  =   0
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
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmDrugPlanCard.frx":141BE
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
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   1
         Top             =   950
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txt摘要 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   3
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lbl复核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复核人"
         Height          =   180
         Left            =   8745
         TabIndex        =   31
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label txt复核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   30
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label lbl复核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "复核日期"
         Height          =   180
         Left            =   8520
         TabIndex        =   29
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label txt复核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   28
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label lbl编制方法 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制方法:"
         Height          =   180
         Left            =   8070
         TabIndex        =   25
         Top             =   660
         Width           =   810
      End
      Begin VB.Label txt编制方法 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "临近期间平均参照法"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9000
         TabIndex        =   24
         Top             =   660
         Width           =   2355
      End
      Begin VB.Label txt计划类型 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Height          =   180
         Left            =   1080
         TabIndex        =   23
         Top             =   660
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "金额合计："
         Height          =   180
         Left            =   240
         TabIndex        =   22
         Top             =   3840
         Width           =   900
      End
      Begin VB.Label Txt审核人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5400
         TabIndex        =   20
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label Txt审核日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5400
         TabIndex        =   19
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt填制日期 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4800
         Width           =   1890
      End
      Begin VB.Label Txt填制人 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Top             =   4440
         Width           =   1890
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   16
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   15
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lbl摘要 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "摘要(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "药品采购计划单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   14
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label Lbl计划类型 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "计划类型:"
         Height          =   180
         Left            =   180
         TabIndex        =   0
         Top             =   660
         Width           =   810
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制人"
         Height          =   180
         Left            =   300
         TabIndex        =   13
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl填制日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "编制日期"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Lbl审核人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核人"
         Height          =   180
         Left            =   4800
         TabIndex        =   11
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl审核日期 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审核日期"
         Height          =   180
         Left            =   4560
         TabIndex        =   10
         Top             =   4860
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14352
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":1456C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14786
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":149A0
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14BBA
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14DD4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":14FEE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15208
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15422
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":1563C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15856
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15A70
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15C8A
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":15EA4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":160BE
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanCard.frx":162D8
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   6615
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPlanCard.frx":164F2
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14393
            Key             =   "STOCK"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrugPlanCard.frx":16D86
            Key             =   "PY"
            Object.ToolTipText     =   "拼音(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmDrugPlanCard.frx":17288
            Key             =   "WB"
            Object.ToolTipText     =   "五笔(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf供应商选择 
      Height          =   2565
      Left            =   5850
      TabIndex        =   27
      Top             =   1890
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblCode 
      AutoSize        =   -1  'True
      Caption         =   "编码"
      Height          =   180
      Left            =   3240
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Menu mnuCol 
      Caption         =   "列名"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(编码和名称)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅编码)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "药名(仅名称)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmDrugPlanCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint编辑状态 As Integer             '1.新增；2、修改；3、验收；4、查看；5、复核；6、修改执行数量
Private mstr单据号 As String                '具体的单据号;
Private mint记录状态 As Integer             '1:正常记录;2-冲销记录;3-已经冲销的原记录
Private mblnSuccess As Boolean              '只要有一张成功，即为True，否则为False
Private mblnFirst As Boolean                '第一次显示
Private mblnSave As Boolean                 '是否存盘和审核   TURE：成功。
Private mblnStart As Boolean
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '是否可以修改
Private mblnChange As Boolean               '是否进行过编辑
Private mintParallelRecord As Integer       '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintPriceUnit As Integer            '指导批发价定价单位:缺省为0-按售价单位定价，可选为1-按药库单位定价；
Private mbln下限 As Boolean                 '仅提取低于储备下限的药品
Private mint上限 As Integer
Private mint下限 As Integer
Private mbln计划数量 As Boolean
Private mbln不考虑库存 As Boolean
Private mlng计划ID As Long
Private mlng库房ID As Long
Private mint计划类型 As Integer
Private mint编制方法 As Integer
Private mstr供应商ID As String
Private mbln中标单位 As Boolean
Private mbln数量方式 As Boolean             'false 按上限计划数量 true 按下限计划数量
Private Str期间  As String                  '月以六位表示,季以五位表示,年以四位表示
Private mstrPrivs As String                     '权限
Private mblnCheckRefresh    As Boolean      '审核时是否改变计划数或者说明
Private mblnClearZeroPlan  As Boolean       '是否删除计划数量为0的记录
Private mblnBaseMedi As Boolean             '是否包含基本药物
Private mblnOnlyBaseMedi As Boolean         '仅仅包含基本药物
Private mintStock As Integer                '常备药选择：0-只提取常备药；1-只提取非常备药；2-不区别是否常备药；
Private mblnEnter As Boolean                '是否进入单元格
Private Const MStrCaption As String = "药品计划管理"
Private mintPlanPoint As Integer            '全院计划不管站点 0-要管站点，1-不管站点
Private mstrToxicologyClass As String       '毒理分类
Private mbln按销量产生计划 As Boolean
Private mstr来源药房 As String               '格式:药房id1,药房id2...
Private mstr来源库房 As String               '格式:药房id1,药房id2...
Private mstrAll来源药房 As String            '所有来源药房。格式:药房id1,药房id2...
Private mstrAll来源库房 As String            '所有来源药房。格式:药房id1,药房id2...
Private mstr自定义库房 As String            '纪录库房ID

Private marrFrom As Variant                   '纪录用户恢复窗体表列格宽度
Private marrInitGrid As Variant                '纪录初始化窗体表列格宽度

Private mstrBeginDate As String
Private mstrEndDate As String
Private mstrNow As String

Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
'从参数表中取药品价格、数量、金额小数位数
Private mintShowCostDigit As Integer            '成本价小数位数
Private mintShowPriceDigit As Integer           '售价小数位数
Private mintShowNumberDigit As Integer          '数量小数位数
Private mintShowMoneyDigit As Integer           '金额小数位数

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mdatThisMondyDate As Date
Private mint价格显示 As Integer             '0-显示成本价;1-显示售价;2-显示成本价和售价
Private mblnViewCost As Boolean             '查看成本价 true-可以查看成本价 false-不可以查看成本价
Private mint供应商选择 As Integer           '0-取上次入库供应商；1-取合同单位
Private mint供应商范围 As Integer           '0-所有供应商；1-中标单位

Private mintDrugNameShow As Integer         '药品显示：0－显示编码和名称；1－仅显示编码；2－仅显示名称

Private mrsHisPlane As ADODB.Recordset      '记录历史采购计划
Private mintLastCol As Integer              '记录表格显示的最后列
Private mstrColumn_UnSelected As String     '记录哪些列被设置为不显示

Private mlng生产商长度 As Long                 '生产商字段长度
Private mlng原产地长度 As Long                 '原产地字段长度

'=========================================================================================
Private mconIntCol序号 As Integer
Private mconIntCol药名 As Integer
Private mconIntCol商品名 As Integer
Private mconIntCol来源 As Integer
Private mconIntCol规格 As Integer
Private mconIntCol生产商 As Integer
Private mconIntCol原产地 As Integer
Private mconIntCol单位 As Integer
Private mconIntCol比例系数 As Integer
Private mconIntcol医保类型 As Integer
Private mconIntCol前期数量 As Integer
Private mconIntCol上期数量 As Integer
Private mconIntCol库存上限 As Integer
Private mconIntCol库存下限 As Integer
Private mconintCol库存数量 As Integer
Private mconintCol上期销量 As Integer
Private mconintCol本期销量 As Integer
Private mconintCol计划数量 As Integer
Private mconintCol执行数量 As Integer
Private mconintCol原执行数量 As Integer
Private mconintCol送货单位 As Integer
Private mconintCol送货数量 As Integer
Private mconintCol送货包装 As Integer
Private mconintCol成本价 As Integer
Private mconIntCol成本金额 As Integer
Private mconIntCol售价 As Integer
Private mconIntCol售价金额 As Integer
Private mconintCol上次供应商 As Integer
Private mconintCol说明 As Integer
Private mconIntCol药品编码和名称 As Integer
Private mconIntCol药品编码 As Integer
Private mconIntCol药品名称 As Integer
Private mconIntCol基本药物 As Integer
Private mconIntCol批准文号 As Integer
Private mconIntColS   As Integer     '总列数
'=========================================================================================

Private Sub ClearZeroPlan()
    Dim n As Integer
    Dim i As Integer
    
    '清除计划数为0的记录（根据计划条件的设置来判断）
    If mblnClearZeroPlan = False Then Exit Sub
    With mshBill
        For n = .rows - 1 To 1 Step -1
            If n = 1 And .rows = 2 And Val(.TextMatrix(n, mconintCol计划数量)) = 0 Then
                For i = 0 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                Exit For
            End If
            If Val(.TextMatrix(n, mconintCol计划数量)) = 0 Then
                .MsfObj.RemoveItem n
            End If
        Next
    End With
End Sub

Private Sub GegReg()
    mint价格显示 = Val(zlDataBase.GetPara("价格显示方式", glngSys, 模块号.药品计划))
    mint供应商选择 = Val(zlDataBase.GetPara("供应商默认选择", glngSys, 模块号.药品计划))
    mint供应商范围 = Val(zlDataBase.GetPara("供应商选择范围", glngSys, 模块号.药品计划))
End Sub


Private Sub IniHisPlaneRec()
    Set mrsHisPlane = New ADODB.Recordset
    With mrsHisPlane
        If .State = 1 Then .Close
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "计划数量", adDouble, 18, adFldIsNullable
        .Fields.Append "执行数量", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "计划类型", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "编制方法", adLongVarChar, 50, adFldIsNullable '
        .Fields.Append "编制人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "编制日期", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "审核人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "审核日期", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "复核人", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "复核日期", adLongVarChar, 20, adFldIsNullable
       
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub LoadHisPlane(ByVal lng库房ID As Long, ByVal lng药品id As Long, ByVal lngRow As Long)
    '装载历史采购计划记录
    Dim rsTmp As ADODB.Recordset
    Dim intExe As Integer
    Dim j As Integer
    
    On Error GoTo errHandle
    If Not mrsHisPlane Is Nothing Then
        mrsHisPlane.Filter = "药品ID=" & lng药品id
        If Not mrsHisPlane.EOF Then Exit Sub
    End If
    
    gstrSQL = "Select B.计划数量, B.执行数量, A.NO, Decode(A.计划类型, 1, '月度计划', 2, '季度计划', 3, '年度计划', '周计划') As 计划类型, " & _
        " Decode(A.编制方法, 1, '往年同期线形参照法', 2, '临近期间平均参照法', 3, '药品储备定额参照法', 4, '药品日销售量参照法', '自定义区间参照法') As 编制方法, A.编制人, A.编制日期, " & _
        " A.审核人 , A.审核日期, A.复核人, A.复核日期 " & _
        " From 药品采购计划 A, 药品计划内容 B " & _
        " Where A.Id = B.计划id And A.审核人 Is Not Null And B.计划数量>0 And A.库房id + 0 = [1] And B.药品id = [2] "
    If Trim(txtNo.Caption) <> "" Then
        gstrSQL = gstrSQL & " And A.NO <> [3]"
    End If
    gstrSQL = gstrSQL & " Order By No Desc "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "LoadHisPlane", lng库房ID, lng药品id, Trim(txtNo.Caption))
    
    With rsTmp
        If Not .EOF Then
            If NVL(!执行数量, 0) = 0 Then
                intExe = 1
            ElseIf NVL(!执行数量, 0) >= NVL(!计划数量, 0) Then
                intExe = 2
            Else
                intExe = 3
            End If
        End If

        Do While Not .EOF
            mrsHisPlane.AddNew
            
            mrsHisPlane!药品id = lng药品id
            mrsHisPlane!计划数量 = zlStr.FormatEx(NVL(!计划数量, 0) / Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
            mrsHisPlane!执行数量 = zlStr.FormatEx(NVL(!执行数量, 0) / Val(mshBill.TextMatrix(lngRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
            mrsHisPlane!NO = !NO
            mrsHisPlane!计划类型 = !计划类型
            mrsHisPlane!编制方法 = !编制方法
            mrsHisPlane!编制人 = NVL(!编制人, "")
            mrsHisPlane!编制日期 = IIf(IsNull(!编制日期), "", Format(!编制日期, "YYYY-MM-DD"))
            mrsHisPlane!审核人 = NVL(!审核人, "")
            mrsHisPlane!审核日期 = IIf(IsNull(!审核日期), "", Format(!审核日期, "YYYY-MM-DD"))
            mrsHisPlane!复核人 = NVL(!复核人, "")
            mrsHisPlane!复核日期 = IIf(IsNull(!复核日期), "", Format(!复核日期, "YYYY-MM-DD"))
            
            .MoveNext
        Loop
    End With
    
    '根据上次计划完成情况对当前药品上色
    If intExe > 0 Then
        mblnEnter = False
        With mshBill
            .Row = lngRow
            .Col = mconIntCol药名
            j = .ColData(mconIntCol药名)
            If .ColData(mconIntCol药名) = 5 Then .ColData(mconIntCol药名) = 0
            
            If intExe = 1 Then
                '未执行
                .MsfObj.CellForeColor = vbRed
            ElseIf intExe = 2 Then
                '完全执行
                .MsfObj.CellForeColor = vbBlue
            Else
                '部分执行
                .MsfObj.CellForeColor = &H8080FF
            End If
    
            .ColData(mconIntCol药名) = j
        End With
        mblnEnter = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ResizeHisPlane()
    On Error Resume Next
    
    With picHis
        If Val(picHscSend.Tag) = 1 Then
            .Height = 1755
        Else
            .Height = 300
        End If
        
        .Top = lblPurchasePrice.Top - .Height - 100
        .Left = mshBill.Left
        .Width = mshBill.Width
    End With
    
    With picStock
        If picHis.Visible Then
            .Top = picHis.Top - .Height - 60
        Else
            .Top = lblPurchasePrice.Top - .Height - 60
        End If
        
        .Left = mshBill.Left
        .Width = mshBill.Width
    End With
    
    With picHeadStock
        .Width = picStock.Width
    End With

    With vsfStock
        .Width = picStock.Width
        .Height = picStock.Height - 330
    End With
    
    With picHscSend
        .Width = picHis.Width
    End With
    
    With vsfHisPlane
        .Width = picHis.Width
    End With

    With pic库房
        .Top = picStock.Top + cmd库房.Height
        .Left = cmd库房.Left + 160
    End With
    
    With mshBill
        If picHis.Visible And picStock.Visible Then
            .Height = picStock.Top - .Top - 60
        ElseIf picHis.Visible And Not picStock.Visible Then
            .Height = picHis.Top - .Top - 60
        ElseIf Not picHis.Visible And picStock.Visible Then
            .Height = picStock.Top - .Top - 60
        Else
            .Height = lblPurchasePrice.Top - .Top - 60
        End If
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, _
        ByVal int编辑状态 As Integer, Optional blnSuccess As Boolean = False, Optional lng库房ID As Long)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mintParallelRecord = 1
    mlng库房ID = lng库房ID
    mstrPrivs = GetPrivFunc(glngSys, 1330)

    mblnSuccess = blnSuccess
    mblnChange = False
    mblnFirst = True

    Set mfrmMain = FrmMain

    If mint编辑状态 = 1 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 2 Then
        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        mblnEdit = False
        CmdSave.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        mblnEdit = False
        CmdSave.Caption = "打印(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "采购计划打印") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If

    End If

    Me.Show vbModal, FrmMain
    blnSuccess = mblnSuccess
    str单据号 = mstr单据号

End Sub

Private Sub ShowHisPlane(ByVal lngRow As Long, ByVal lng药品id As Long)
    '显示历史采购计划
    
    vsfHisPlane.rows = 1
    vsfHisPlane.rows = 2
    
    lblDiag.Caption = "近期采购计划执行情况"
    
    If mrsHisPlane Is Nothing Then Exit Sub
    
    mrsHisPlane.Filter = "药品ID=" & lng药品id
    If mrsHisPlane.EOF Then Exit Sub
    
    lblDiag.Caption = lblDiag.Caption & "(" & mrsHisPlane.RecordCount & ")"
    
    With vsfHisPlane
        .Redraw = flexRDNone
        Do While Not mrsHisPlane.EOF
            .TextMatrix(.rows - 1, .ColIndex("计划数量")) = zlStr.FormatEx(mrsHisPlane!计划数量, mintShowNumberDigit, , True)
            .TextMatrix(.rows - 1, .ColIndex("执行数量")) = zlStr.FormatEx(NVL(mrsHisPlane!执行数量, 0), mintShowNumberDigit, , True)
            .TextMatrix(.rows - 1, .ColIndex("NO")) = mrsHisPlane!NO
            .TextMatrix(.rows - 1, .ColIndex("计划类型")) = mrsHisPlane!计划类型
            .TextMatrix(.rows - 1, .ColIndex("编制方法")) = mrsHisPlane!编制方法
            .TextMatrix(.rows - 1, .ColIndex("编制人")) = mrsHisPlane!编制人
            .TextMatrix(.rows - 1, .ColIndex("编制日期")) = mrsHisPlane!编制日期
            .TextMatrix(.rows - 1, .ColIndex("审核人")) = mrsHisPlane!审核人
            .TextMatrix(.rows - 1, .ColIndex("审核日期")) = mrsHisPlane!审核日期
            .TextMatrix(.rows - 1, .ColIndex("复核人")) = mrsHisPlane!复核人
            .TextMatrix(.rows - 1, .ColIndex("复核日期")) = mrsHisPlane!复核日期
            
            If NVL(mrsHisPlane!执行数量, 0) = 0 Then
                '未执行
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbRed
            ElseIf NVL(mrsHisPlane!执行数量, 0) >= mrsHisPlane!计划数量 Then
                '完全执行
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = vbBlue
            Else
                '部分执行
                .Cell(flexcpForeColor, .rows - 1, 0, .rows - 1, .Cols - 1) = &H8080FF
            End If
            
            .Cell(flexcpFontBold, .rows - 1, .ColIndex("计划数量")) = True
            .Cell(flexcpFontBold, .rows - 1, .ColIndex("执行数量")) = True
            
            .rows = .rows + 1
                        
            If chkMore.Value = 0 And .rows > 4 Then
                .Redraw = flexRDDirect
                Exit Sub
            End If
            
            mrsHisPlane.MoveNext
        Loop
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub chkMore_Click()
    If mshBill.TextMatrix(mshBill.Row, 0) <> "" Then
        Call ShowHisPlane(mshBill.Row, Val(mshBill.TextMatrix(mshBill.Row, 0)))
    End If
End Sub

Private Sub chk是否显示库存情况_Click()
    With picStock
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
End Sub
Private Sub chk来源库房_Click()
    If chk来源库房.Value = 1 Then
        If chk所有库房.Value = 1 Then chk所有库房.Value = 0
    End If
    Call 显示库存
End Sub
Private Sub chk来源药房_Click()
    If chk来源药房.Value = 1 Then
        If chk所有库房.Value = 1 Then chk所有库房.Value = 0
    End If
    Call 显示库存
End Sub
Private Sub chk所有库房_Click()
    If chk所有库房.Value = 1 Then
        If chk来源药房.Value = 1 Then chk来源药房.Value = 0
        If chk来源库房.Value = 1 Then chk来源库房.Value = 0
        If mstr自定义库房 <> "" Then mstr自定义库房 = ""
    End If
    Call 显示库存
End Sub
Private Sub chk隐藏近期采购计划_Click()
    With picHis
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("你确定要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Unload Me
End Sub

'查找
Private Sub cmdFind_Click()

    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntCol药品编码和名称, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    Dim intItem As Integer
    Dim intItems As Integer
    '取得存储库房
    mstr自定义库房 = ""
    intItems = Me.lvw存储库房.ListItems.count
    For intItem = 1 To intItems
        If lvw存储库房.ListItems(intItem).Checked Then
            mstr自定义库房 = mstr自定义库房 & "," & Mid(lvw存储库房.ListItems(intItem).Key, 2)
        End If
    Next
    mstr自定义库房 = Mid(mstr自定义库房, 2)
    
    With pic库房
        .Visible = False
    End With
    
    If mstr自定义库房 <> "" Then
        If chk所有库房.Value = 1 Then chk所有库房.Value = 0
    End If
    Call 显示库存
End Sub

Private Sub cmd取消_Click()
    With pic库房
        .Visible = False
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntCol药名, txtCode.Text, False
    ElseIf KeyCode = vbKeyEscape Then
        If Msf供应商选择.Visible Then
            Msf供应商选择.ZOrder 1
            Msf供应商选择.Visible = False
            Exit Sub
        End If
'        Call cmdCancel_Click
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub


Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), 0, 0, 1330, "药品采购计划单", txtNo.Tag, mint价格显示)
        '退出
        Unload Me
        Exit Sub
    End If

    If mint编辑状态 = 3 Then        '审核
        If mblnCheckRefresh Then
            If Not SaveCard Then
                Exit Sub
            End If
        End If
        If SaveCheck = True Then
            If Val(zlDataBase.GetPara("审核打印", glngSys, 模块号.药品计划)) = 1 Then
                '打印
                If zlStr.IsHavePrivs(mstrPrivs, "采购计划打印") Then
                    ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & txtNo.Tag, IIf(mint价格显示 = 0, "ReportFormat=1", IIf(mint价格显示 = 1, "ReportFormat=2", "ReportFormat=3")), 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 5 Then        '复核
        If SaveReCheck = True Then
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint编辑状态 = 6 Then        '修改执行数量
        If SaveExeAmount = True Then
            Unload Me
        End If
        Exit Sub
    End If

    If ValidData = False Then Exit Sub
    blnSuccess = SaveCard

    If blnSuccess = True Then

        If Val(zlDataBase.GetPara("存盘打印", glngSys, 模块号.药品计划)) = 1 Then
            '打印
            If zlStr.IsHavePrivs(mstrPrivs, "采购计划打印") Then
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "单据编号=" & txtNo.Tag, IIf(mint价格显示 = 0, "ReportFormat=1", IIf(mint价格显示 = 1, "ReportFormat=2", "ReportFormat=3")), 2
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    txt摘要.Text = ""
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "上一张单据的NO号：" & txtNo.Tag
End Sub

Private Sub Form_Activate()
    Dim intMonth As Integer
    Dim datCurrDate As Date
    Dim intWeekDay As Integer
    Const intMonday As Integer = vbMonday
    Dim intCountDay As Integer
    
    If mblnFirst = False Then Exit Sub

    If Not zlStr.IsHavePrivs(mstrPrivs, "所有库房") Then
        chk是否显示库存情况.Visible = False
    Else
        chk是否显示库存情况.Visible = True
        Call Init存储库房
    End If
    
    '初始化简码方式
    If (mint编辑状态 = 1 Or mint编辑状态 = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint简码方式 = Val(zlDataBase.GetPara("简码方式", , , 0))    '默认拼音简码
        Logogram staThis, gint简码方式
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    If mint编辑状态 = 1 Then
        Dim str用途ID As String, str剂型编码 As String
        Dim lng库房ID As Long, int计划类型 As Integer, int编制方法 As Integer, bln数量方式 As Boolean
        Dim strToxicologyClass As String
        
        If frmDrugPlanCondition.GetCondition(mfrmMain, str用途ID, str剂型编码, lng库房ID, int计划类型, _
                int编制方法, mbln下限, mint上限, mint下限, mbln计划数量, _
                mstr供应商ID, mbln中标单位, mstrBeginDate, mstrEndDate, mbln不考虑库存, _
                mblnClearZeroPlan, mblnBaseMedi, mintStock, bln数量方式, mblnOnlyBaseMedi, _
                strToxicologyClass, mbln按销量产生计划, mstr来源药房, mstr来源库房, mstrAll来源药房, mstrAll来源库房) = True Then
            mlng库房ID = lng库房ID
            mint计划类型 = int计划类型
            mint编制方法 = int编制方法
            mbln数量方式 = bln数量方式
            mstrToxicologyClass = strToxicologyClass
            Select Case mint计划类型
                Case 1       '月计划
                    Str期间 = Format(DateAdd("m", 1, Sys.Currentdate), "yyyyMM")
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Right(Str期间, 2) & "月" & ") " & LblTitle.Tag & "采购计划"
                    
                    mshBill.TextMatrix(0, mconintCol上期销量) = "上月销量"
                    mshBill.TextMatrix(0, mconintCol本期销量) = "本月销量"
                Case 2       '季计划
                    intMonth = Month(DateAdd("Q", 1, Sys.Currentdate))
                    Str期间 = Format(DateAdd("Q", 1, Sys.Currentdate), "yyyy") & IIf(intMonth <= 3, 1, IIf(intMonth >= 10, 4, IIf(intMonth <= 9 And intMonth >= 7, 3, 2)))
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Right(Str期间, 1) & "季" & ")" & LblTitle.Tag & "采购计划"
                    
                    mshBill.TextMatrix(0, mconintCol上期销量) = "上季度销量"
                    mshBill.TextMatrix(0, mconintCol本期销量) = "本季度销量"
                Case 3      '年计划
                    Str期间 = Format(DateAdd("yyyy", 1, Sys.Currentdate), "yyyy")
                    LblTitle.Caption = GetUnitName & "(" & Str期间 & "年" & ")" & LblTitle.Tag & "采购计划"
                    
                    mshBill.TextMatrix(0, mconintCol上期销量) = "上年销量"
                    mshBill.TextMatrix(0, mconintCol本期销量) = "本年销量"
                Case 4      '周计划
                    datCurrDate = Sys.Currentdate
                    intWeekDay = Weekday(datCurrDate)
                    If intWeekDay = 1 Then
                        intCountDay = -6
                    Else
                        intCountDay = intMonday - intWeekDay
                    End If
                    mdatThisMondyDate = DateAdd("d", intCountDay, datCurrDate)
                    Str期间 = Format(DateAdd("d", 7, mdatThisMondyDate), "yyyyMMDD")
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Mid(Str期间, 5, 2) & "月" & Right(Str期间, 2) & "日" & ")" & LblTitle.Tag & "采购计划"
                    
                    mshBill.TextMatrix(0, mconintCol上期销量) = "上周销量"
                    mshBill.TextMatrix(0, mconintCol本期销量) = "本周销量"
            End Select
            
            If mint编制方法 = 5 Then
                '自定义区间编制法
                mshBill.TextMatrix(0, mconIntCol前期数量) = "本期数量"
                mshBill.TextMatrix(0, mconIntCol上期数量) = "本期销量"
                mshBill.TextMatrix(0, mconintCol上期销量) = "上月销量"
                mshBill.TextMatrix(0, mconintCol本期销量) = "本月销量"
            End If

            ReFreshALLDrug str用途ID, str剂型编码, lng库房ID, int计划类型, int编制方法, bln数量方式
        Else
            Unload Me
            Exit Sub
        End If
        If mshBill.Visible = True Then
            mshBill.SetFocus
        End If
    Else
        mblnChange = False
        Select Case mintParallelRecord
            Case 1
                '正常
            Case 2
                '单据已被删除
                MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
            Case 3
                '修改的单据已被审核
                MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
                Unload Me
                Exit Sub
        End Select
    End If

    mblnStart = True
End Sub

Private Sub ReFreshALLDrug(ByVal str用途ID, ByVal str剂型 As String, _
    ByVal lng库房ID As Long, ByVal int计划类型 As Integer, ByVal int编制方法 As Integer, ByVal bln数量方式 As Boolean)
        '---------------------------------------------------
        '--功能:对所有药品进行计划编制
        '--参数:
        '---------------------------------------------------
    Dim rsAllDrug As New ADODB.Recordset
    Dim rspurchase As New ADODB.Recordset
    Dim intRecord As Long
    Dim intRow  As Long
    Dim rsData As ADODB.Recordset
    Dim lng供应商ID As Long
    Dim str供应商 As String
    Dim str产地 As String
    Dim str原产地 As String
    Dim dbl库存数量 As Double
    Dim blnOK As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim rs合同单位 As ADODB.Recordset
    Dim str药名 As String
    Dim str剂型串 As String
    Dim str送货单位 As String
    Dim dbl送货包装 As Double
    
    On Error GoTo errHandle
    Me.Refresh
    Me.MousePointer = vbHourglass
    mshBill.Redraw = False
    staThis.Panels(2).Text = "正在计算"
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    Pic单据.Enabled = False

    str剂型串 = Replace(str剂型, "'", "")
        
    '取指定条件的药品信息
    gstrSQL = "" & _
         " SELECT DISTINCT A.药品ID,'[' || F.编码 || ']' As 药品编码, F.名称 As 通用名, B.名称 As 商品名,A.药品来源,f.规格,a.基本药物," & _
         " Decode(" & mintUnit & ", 1, f.计算单位, 2, a.门诊单位, 3, a.住院单位, a.药库单位) As 单位," & _
         " DECODE(A.成本价,NULL,NVL(A.指导批发价,0),0,NVL(A.指导批发价,0),NVL(A.成本价,0)) AS 单价,F.产地,A.原产地," & _
         " Decode(" & mintUnit & ", 1, 1, 2, a.门诊包装, 3, a.住院包装, a.药库包装) As 比例系数," & _
         " Nvl(G.现价, 0) 售价,a.上次售价,Nvl(F.是否变价,0) 是否变价, a.送货单位, a.送货包装,f.费用类型, a.上次供应商id As 供应商id, d.名称 As 供应商,nvl(a.上次批准文号,a.批准文号) as 批准文号 " & _
         " FROM 药品规格 A,收费项目别名 B,诊疗项目目录 C,诊疗分类目录 L,收费项目目录 F,药品特性 T, 收费价目 G, 供应商 D "
    
    gstrSQL = gstrSQL & " WHERE A.药品ID=F.ID And A.药名ID=C.ID and A.药名ID=T.药名ID And C.分类ID=L.ID and L.类型 in (1,2,3)" & _
         " And A.药品ID = B.收费细目ID(+) And B.性质(+)=3 " & _
         " AND (F.撤档时间>=TO_DATE('3000-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS') OR F.撤档时间 IS NULL)" & _
         " And A.药品id = G.收费细目id And (G.终止日期 Is Null Or Sysdate Between G.执行日期 And Nvl(G.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
         GetPriceClassString("G") & " And a.上次供应商id = d.Id(+) "
    
    '毒理
    If mstrToxicologyClass <> "" Then
            gstrSQL = gstrSQL & " And " & mstrToxicologyClass
    End If
    
    If mintStock <> 2 Then
        gstrSQL = gstrSQL & " And Nvl(A.是否常备, 0) = [4] "
    End If
    
    If mblnOnlyBaseMedi = True Then
        gstrSQL = gstrSQL & " and a.基本药物 is not null "
    End If
    
    If mblnBaseMedi = False And mblnOnlyBaseMedi = False Then
        gstrSQL = gstrSQL & " And A.基本药物 Is Null "
    End If
    
    If str用途ID = "" Then
        gstrSQL = gstrSQL & " And L.ID Is NULL "
    ElseIf str用途ID <> "所有分类" Then
        gstrSQL = gstrSQL & " And L.ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) "
    End If
    
    If str剂型串 <> "" Then
        gstrSQL = gstrSQL & " And T.药品剂型 in (select * from Table(Cast(f_Str2list([3]) As zlTools.t_Strlist))) "
    End If

    If lng库房ID = 0 Then
        '如果是全库房，取所有库房库存，并从药品规格中取上次供应商和上次产地
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select A.药品ID, C.ID As 供应商ID,C.名称 As 供应商, B.上次产地,B.原产地, A.库存数量, A.平均售价 " & _
                  " From (Select 药品id, Sum(实际数量) As 库存数量, " & _
                  " Decode(Sign(Sum(实际数量)), 1, Decode(Sign(Sum(实际金额)), 1, Sum(实际金额), 0) / Sum(实际数量), 0) 平均售价 " & _
                  " From 药品库存 " & _
                  " Where 性质 = 1 " & _
                  " Group By 药品id) A, " & _
                  " 药品规格 B, " & _
                  " (SELECT ID,名称 FROM 供应商 WHERE SUBSTR(类型,1,1)=1) C " & _
                  " Where A.药品id = B.药品id And B.上次供应商id = C.ID(+)) E "
    
    Else
        '取库存数量，及最大批次的供应商，上次产地
        gstrSQL = "( " & gstrSQL & ") D," & _
                  " (Select DISTINCT A.药品ID, C.ID As 供应商ID,C.名称 As 供应商, B.上次产地, B.原产地, A.库存数量, A.平均售价 " & _
                  " From (Select 药品id, Sum(实际数量) As 库存数量, " & _
                  " Decode(Sign(Sum(实际数量)), 1, Decode(Sign(Sum(实际金额)), 1, Sum(实际金额), 0) / Sum(实际数量), 0) 平均售价 " & _
                  " From 药品库存 " & _
                  " Where 性质 = 1 "
        If mstr来源库房 <> "" Then
            gstrSQL = gstrSQL & " And 库房id In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)))"
        Else
            gstrSQL = gstrSQL & " AND 库房ID=[1] "
        End If
        
        gstrSQL = gstrSQL & " Group By 药品id) A, " & _
                  " (Select 药品id, max(上次供应商id) as 上次供应商id, max(上次产地) as 上次产地, max(原产地) as 原产地  From 药品库存 " & _
                  " Where 性质 = 1 AND (库房ID=[1] or 库房id In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList)))) " & _
                  " And (药品id,Nvl(批次, 0)) In " & _
                  " (Select 药品id,Nvl(Max(Nvl(批次, 0)), 0) 批次 From 药品库存 Where 性质 = 1 AND (库房ID=[1] or 库房id In(select * from Table(Cast(f_Num2List([5]) As zlTools.t_NumList))))  Group By 药品id ) Group By 药品id) B, " & _
                  " (SELECT ID,名称 FROM 供应商 WHERE SUBSTR(类型,1,1)=1) C " & _
                  " Where A.药品id = B.药品id And B.上次供应商id = C.ID(+)) E, " & _
                  " (Select distinct 收费细目id From 收费执行科室 Where 执行科室id = [1]) Z "
                  
    End If
    
    '加上提取药品储备限额的SQL
    If mbln下限 Then
        gstrSQL = gstrSQL & _
                ",(Select 药品ID,sum(下限) 下限  " & _
                "     From 药品储备限额  " & _
                "" & IIf(lng库房ID = 0, "", " Where 库房ID=[1]") & _
                "     Group By 药品ID)     F"
    End If

    '联合所有（在最外层加上取药品储备限额.下限）
    gstrSQL = "SELECT d.药品id, d.药品编码, d.通用名, d.商品名, d.药品来源,d.规格, " _
            & "DECODE (e.上次产地, NULL, d.产地, e.上次产地) AS 产地," _
            & "DECODE (e.原产地, NULL, d.原产地, e.原产地) AS 原产地," _
            & "d.单位,nvl(e.库存数量,0)/d.比例系数 as 库存数量" & IIf(mbln下限, ",nvl(F.下限,0)/d.比例系数 as 下限", "") & " , d.单价 as 单价 ,Nvl(e.供应商id, d.供应商id) As 供应商id, Nvl(e.供应商, d.供应商) As 供应商,d.比例系数, " _
            & " Decode(D.是否变价, 0, D.售价, Decode(nvl(d.上次售价,0), 0, Decode(Nvl(E.平均售价, 0), 0, D.售价, E.平均售价), d.上次售价)) 售价,d.送货单位,d.送货包装,d.费用类型,d.基本药物,d.批准文号 from " _
            & gstrSQL _
            & " WHERE d.药品id = e.药品id (+) "
    If lng库房ID <> 0 Then
        gstrSQL = gstrSQL & " And d.药品id = z.收费细目id "
    End If
    
    If mbln下限 Then
        '加上条件判断
        '加上储备限额的判断，低于储备限额的药品才提取出来做采购计划
        gstrSQL = gstrSQL & " And d.药品ID=F.药品ID(+)"
        gstrSQL = "Select * From (" & gstrSQL & ") Where (库存数量<下限 and 下限<>0)"
    End If
    gstrSQL = gstrSQL & " Order by 药品编码"

    If rsAllDrug.State = 1 Then rsAllDrug.Close
    
    Set rsAllDrug = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, str用途ID, str剂型串, mintStock, mstr来源库房)

    With rsAllDrug
        intRecord = .RecordCount

        If intRecord = 0 Then
            mshBill.Redraw = True
            Me.MousePointer = vbDefault
            CmdSave.Enabled = True
            CmdCancel.Enabled = True
            Pic单据.Enabled = True
            Me.staThis.Panels(2).Text = ""
            Exit Sub
        End If
        .MoveFirst
        Me.Refresh
        DoEvents
        Do While Not .EOF
            dbl库存数量 = IIf(IsNull(!库存数量), 0, !库存数量)
            lng供应商ID = NVL(!供应商id, 0)
            str供应商 = IIf(IsNull(!供应商), "", !供应商)
            str产地 = IIf(IsNull(!产地), "", !产地)
            str原产地 = IIf(IsNull(!原产地), "", !原产地)
            
            blnOK = True
            
            '如果无库存，则从药品规格中取供应商，上次产地
            If IIf(IsNull(!库存数量), 0, !库存数量) = 0 Then
                If mstr供应商ID = "" Then
                    gstrSQL = "Select B.id 供应商ID, B.名称 供应商, C.上次产地, C.原产地, 0 库存数量 from " & _
                          " (Select id,名称 From 供应商 Where Substr(类型, 1, 1) = 1)  B, 药品规格 C " & _
                          " Where C.上次供应商id = B.ID(+) And 药品id = [1] "
                Else
                    gstrSQL = "Select B.id 供应商ID, B.名称 供应商, C.上次产地, C.原产地, 0 库存数量 from " & _
                          " (Select A.id,A.名称 From 供应商 A Where Substr(A.类型, 1, 1) = 1 And A.id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))))  B, 药品规格 C " & _
                          " Where C.上次供应商id = B.ID And 药品id = [1] "
                End If
                
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次供应商及产地信息]", Val(rsAllDrug!药品id), mstr供应商ID)
                If rsData.RecordCount > 0 Then
                    blnOK = True
                    lng供应商ID = NVL(rsData!供应商id, 0)
                    str产地 = IIf(IsNull(rsData!上次产地), "", rsData!上次产地)
                    str原产地 = IIf(IsNull(rsData!原产地), "", rsData!原产地)
                    str供应商 = IIf(IsNull(rsData!供应商), "", rsData!供应商)
                    dbl库存数量 = IIf(IsNull(rsData!库存数量), 0, rsData!库存数量)
                Else
                    blnOK = False
                End If
            End If
            If mstr供应商ID <> "" Then
                If InStr("," & mstr供应商ID & ",", "," & lng供应商ID & ",") = 0 Then
                    If mbln中标单位 Then
                        gstrSQL = "Select b.名称 from 药品中标单位 a,供应商 b where  a.药品ID=[1] and   a.单位id=b.id and a.单位ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))"
                        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, Val(rsAllDrug!药品id), mstr供应商ID)
                        blnOK = (rsTmp.RecordCount > 0)
                        If blnOK = True Then str供应商 = rsTmp!名称
                    Else
                        blnOK = False
                    End If
                End If
            End If
            If blnOK Then
                intRow = intRow + 1
                mshBill.TextMatrix(intRow, 0) = !药品id

                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    str药名 = !通用名
                Else
                    str药名 = IIf(IsNull(!商品名), !通用名, !商品名)
                End If
                
                mshBill.TextMatrix(intRow, mconIntCol药品编码和名称) = !药品编码 & str药名
                mshBill.TextMatrix(intRow, mconIntCol药品编码) = !药品编码
                mshBill.TextMatrix(intRow, mconIntCol药品名称) = str药名
                
                If mintDrugNameShow = 1 Then
                    mshBill.TextMatrix(intRow, mconIntCol药名) = mshBill.TextMatrix(intRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    mshBill.TextMatrix(intRow, mconIntCol药名) = mshBill.TextMatrix(intRow, mconIntCol药品名称)
                Else
                    mshBill.TextMatrix(intRow, mconIntCol药名) = mshBill.TextMatrix(intRow, mconIntCol药品编码和名称)
                End If
                
                mshBill.TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(!商品名), "", !商品名)

                mshBill.TextMatrix(intRow, mconIntCol来源) = NVL(!药品来源)
                mshBill.TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(!规格), "", !规格)
                mshBill.TextMatrix(intRow, mconIntCol单位) = IIf(IsNull(!单位), "", !单位)
                mshBill.TextMatrix(intRow, mconIntcol医保类型) = IIf(IsNull(!费用类型), "", !费用类型)
                mshBill.TextMatrix(intRow, mconIntCol生产商) = str产地
                mshBill.TextMatrix(intRow, mconIntCol原产地) = str原产地
                
                mshBill.TextMatrix(intRow, mconintCol上次供应商) = str供应商
                If mint供应商选择 = 1 Then
                    gstrSQL = "Select B.名称 From 药品规格 A, 供应商 B Where A.合同单位id = B.ID And A.药品id = [1] "
                    Set rs合同单位 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, Val(rsAllDrug!药品id))
                    If Not rs合同单位.EOF Then
                        mshBill.TextMatrix(intRow, mconintCol上次供应商) = rs合同单位!名称
                    End If
                End If
                
                mshBill.TextMatrix(intRow, mconintCol库存数量) = zlStr.FormatEx(dbl库存数量, mintShowNumberDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol比例系数) = !比例系数
                
                mshBill.TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(IIf(IsNull(!单价), "0", !单价 * !比例系数), mintShowPriceDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(IIf(IsNull(!售价), "0", !售价 * !比例系数), mintShowPriceDigit, , True)
                mshBill.TextMatrix(intRow, mconIntCol基本药物) = IIf(IsNull(!基本药物), "", !基本药物)
                mshBill.TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(!批准文号), "", !批准文号)
                
                SetNumer !药品id, lng库房ID, IIf(IsNull(!库存数量), 0, !库存数量), intRow, int计划类型, int编制方法, bln数量方式
                
                str送货单位 = IIf(IsNull(!送货单位), "", !送货单位)
                dbl送货包装 = IIf(IsNull(!送货包装), 0, !送货包装)
                If dbl送货包装 <> 0 Then
                    mshBill.TextMatrix(intRow, mconintCol送货包装) = dbl送货包装
                    mshBill.TextMatrix(intRow, mconintCol送货单位) = str送货单位 & "(1" & str送货单位 & "=" & zlStr.FormatEx(dbl送货包装 / !比例系数, 1, , True) & mshBill.TextMatrix(intRow, mconIntCol单位) & ")"
                    mshBill.TextMatrix(intRow, mconintCol送货数量) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconintCol计划数量)) / dbl送货包装, 1, , True)
                End If
                
                If intRow = mshBill.rows - 1 Then mshBill.rows = mshBill.rows + 1
                Call zlControl.StaShowPercent(intRow / intRecord, staThis.Panels(2), frmDrugPlanCard)
            End If
            .MoveNext
        Loop
    End With
    Call ClearZeroPlan
    Call RefreshRowNO(mshBill, mconIntCol序号, 1)
    Call 显示合计金额
    Me.MousePointer = vbDefault
    mshBill.Redraw = True
    CmdSave.Enabled = True
    Pic单据.Enabled = True
    CmdCancel.Enabled = True
    mshBill.Col = mconintCol计划数量
    Me.staThis.Panels(2).Text = ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDate(ByVal int模式 As Integer, ByVal datCurrent As Date, _
        ByRef strBegin As String, ByRef strEnd As String) As Boolean
    Dim rsdate As New Recordset

    'int模式=1,月计划，2：季计划
    On Error GoTo errHandle
    GetDate = False
    Select Case int模式
    Case 1
        strBegin = Year(datCurrent) & "-" & String(2 - Len(Month(datCurrent)), "0") & Month(datCurrent) & "-01"
        gstrSQL = "select last_day(to_date([1],'yyyy-mm-dd')) from dual"
        Set rsdate = zlDataBase.OpenSQLRecord(gstrSQL, "GetDate", Format(datCurrent, "yyyy-mm-dd"))
        strEnd = Format(rsdate.Fields(0), "yyyy-mm-dd")
        rsdate.Close
    Case 2
        Select Case DatePart("Q", datCurrent)
            Case 1
                strBegin = Year(datCurrent) & "-01-01"
                strEnd = Year(datCurrent) & "-03-31"
            Case 2
                strBegin = Year(datCurrent) & "-04-01"
                strEnd = Year(datCurrent) & "-06-30"
            Case 3
                strBegin = Year(datCurrent) & "-07-01"
                strEnd = Year(datCurrent) & "-09-30"
            Case 4
                strBegin = Year(datCurrent) & "-10-01"
                strEnd = Year(datCurrent) & "-12-31"
        End Select
    Case 4
        strBegin = Format(datCurrent, "yyyy-mm-dd")
        strEnd = Format(DateAdd("d", 6, datCurrent), "yyyy-mm-dd")
    End Select
    GetDate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'设置前期数量，上期数量，计划数量,金额等
Private Sub SetNumer(ByVal lng药品id As Long, ByVal lng库房ID As Long, _
        ByVal num库存数量 As Double, ByVal intCurrentRow As Integer, _
        ByVal int计划类型 As Integer, ByVal int编制方法 As Integer, ByVal bln数量方式 As Boolean)
    '---------------------------------------------------------------------------
    '--功能:确定耗用数量和计划数量
    '   1 )往年同期线性参照法：根据去前年同期药品的消耗情况，按线性规划原则预测消耗，对比库存产生采购计划供用户修改调整
    '   2 )临近期间平均参照法：以同年临近期间(前期、上期)的平均消耗预测消耗对比库存产生采购计划供用户修改调整；
    '   3 )药品储备参照法：根据药品储务下限与库存相减所得的差额为药品计划采购数；

    '--参数:
    '       int计划类型:1:月度计划,2.季度计划,3.年度计划,4.周计划
    '       int编制方法:1 表示往年同期线性参照法,2 临近期间平均参照法,3.储备限额;4.日销售量;5-自定义区间
    '       bln数量方式:false 按上限指定计划数量  计划数量=上限数量-库存数量;  true 按下限计划数量 计划数量=下限数量-库存数量
    '--返回:
    '---------------------------------------------------------------------------
    Dim num前期数量 As Double
    Dim num上期数量 As Double
    Dim num上期销量 As Double
    Dim num计划数量 As Double
    Dim num上限 As Double, num下限 As Double
    Dim lng天数 As Long

    Dim dat前期 As Date
    Dim dat上期 As Date
    Dim strBegin As String
    Dim strEnd As String
    Dim rsNum As New Recordset
    
    Dim str汇总最大日期 As String
    Dim str收发结束时间 As String
    
    On Error GoTo errHandle
    num库存数量 = IIf(mbln不考虑库存 = True, 0, num库存数量)
    
    With mshBill
        Select Case int编制方法
            Case 1      '往年同期线形参照   只有月度、季度计划
                dat前期 = DateAdd("m", Choose(int计划类型, 1, 3), DateAdd("yyyy", -2, mstrNow))
                dat上期 = DateAdd("m", Choose(int计划类型, 1, 3), DateAdd("yyyy", -1, mstrNow))
                If lng库房ID = 0 Then
                    GetDate int计划类型, dat前期, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                            & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0 = [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd))
                            
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int计划类型, dat上期, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0 = [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    GetDate int计划类型, dat前期, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and 库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    GetDate int计划类型, dat上期, strBegin, strEnd
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and 库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                num前期数量 = num前期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                '计划数量=2×上期数量－前期数量－库存数量
                If mbln计划数量 Then
                    num计划数量 = 2 * num上期数量 - num前期数量 - num库存数量
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
                .TextMatrix(intCurrentRow, mconIntCol前期数量) = zlStr.FormatEx(num前期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
            Case 2      '临近期间平均参照法
                dat前期 = Choose(int计划类型, DateAdd("m", -2, mstrNow), DateAdd("m", -6, mstrNow), DateAdd("yyyy", -2, mstrNow), DateAdd("d", -14, mdatThisMondyDate))
                dat上期 = Choose(int计划类型, DateAdd("m", -1, mstrNow), DateAdd("m", -3, mstrNow), DateAdd("yyyy", -1, mstrNow), DateAdd("d", -7, mdatThisMondyDate))
                If lng库房ID = 0 Then
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0= [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat前期), "yyyy-mm-dd hh:mm:ss")), _
                        CDate(Format(dat前期, "yyyy-mm-dd hh:mm:ss")))
                    
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
    
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0= [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat上期), "yyyy-mm-dd hh:mm:ss")), _
                            CDate(Format(dat上期, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 前期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and a.库房id+0=[1] " _
                            & "  AND 药品id+0= [2]" _
                            & "  AND 日期 BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat前期), "yyyy-mm-dd hh:mm;ss")), _
                            CDate(Format(dat前期, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        num前期数量 = 0
                    Else
                        num前期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                    
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and a.库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(Format(DateAdd(Choose(int计划类型, "m", "m", "m", "d"), Choose(int计划类型, -1, -3, -12, -7), dat上期), "yyyy-mm-dd hh:mm:ss")), _
                            CDate(Format(dat上期, "yyyy-mm-dd hh:mm:ss")))
                            
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                num前期数量 = num前期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                '计划数量 = (前期数量 + 上期数量) / 2 - 库存数量
                If mbln计划数量 Then
                    num计划数量 = (num上期数量 + num前期数量) / 2 - num库存数量
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
                .TextMatrix(intCurrentRow, mconIntCol前期数量) = zlStr.FormatEx(num前期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
            
            Case 3      '药品储备定额参照法
                If mstrBeginDate = "" Or mstrEndDate = "" Then
                    mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
                    mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Select Max(日期) As 日期 From 药品收发汇总"
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
                If NVL(rsNum!日期, "") = "" Then
                    str汇总最大日期 = Format(DateAdd("d", 1, CDate(mstrBeginDate)), "yyyy-mm-dd")
                Else
                    str汇总最大日期 = Format(DateAdd("d", 1, rsNum!日期), "yyyy-mm-dd")
                End If
                
                str收发结束时间 = Format(DateAdd("d", 1, CDate(mstrEndDate)), "yyyy-mm-dd")
                
                If lng库房ID = 0 Then
                    gstrSQL = " Select Sum(上期数量) As 上期数量 " _
                            & " From (SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                            & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0= [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.入出系数 * Nvl(A.实际数量, 0) * Nvl(A.付数, 1))) As 上期数量 " _
                            & " From 药品收发记录 A, 药品入出类别 B " _
                            & " Where A.单据<>6 And A.入出类别id = B.ID And B.系数 = -1 And 药品id + 0 = [1] And " _
                            & " 审核日期 >= [2] " _
                            & " And 审核日期 Between [4] And [5])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str汇总最大日期), CDate(str收发结束时间))
                            
                    If Not rsNum.EOF Then
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0)) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数))
                         .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                    End If
                    rsNum.Close
                Else
                    gstrSQL = " Select Sum(上期数量) As 上期数量 " _
                            & " From (SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                            & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and a.库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.入出系数 * Nvl(A.实际数量, 0) * Nvl(A.付数, 1))) As 上期数量 " _
                            & " From 药品收发记录 A, 药品入出类别 B " _
                            & " Where A.入出类别id = B.ID And B.系数 = -1 And A.库房id + 0 = [1] And 药品id + 0 = [2] And " _
                            & " 审核日期 >= [3] " _
                            & " And 审核日期 Between [5] And [6])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str汇总最大日期), CDate(str收发结束时间))
                            
                    If Not rsNum.EOF Then
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0)) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数))
                         .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                    End If
                    rsNum.Close
                End If
                
                If lng库房ID = 0 Then
                    gstrSQL = "select sum(Nvl(上限,0)) as  上限,sum(Nvl(下限,0)) as  下限 " _
                            & " from 药品储备限额 " _
                           & " where 药品id=[1] "
    
                Else
                    gstrSQL = "select Nvl(上限,0) As 上限,Nvl(下限,0) as 下限 " _
                            & " from 药品储备限额 " _
                           & " where 药品id=[1] " _
                           & "   and 库房id=[2]"
    
                End If
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, lng库房ID)
                
                If rsNum.EOF Then
                    num上限 = 0
                    num下限 = 0
                Else
                    num上限 = IIf(IsNull(rsNum!上限), 0, rsNum!上限)
                    num下限 = IIf(IsNull(rsNum!下限), 0, rsNum!下限)
                End If
    
                '把各单位转换成药库单位先
                num上限 = num上限 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                num下限 = num下限 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                '计划数量=储备上限－库存数量
                If mbln计划数量 Then
                    If bln数量方式 = False Then
                        num计划数量 = IIf(num上限 > num库存数量, num上限 - num库存数量, 0)
                    Else
                        num计划数量 = IIf(num下限 > num库存数量, num下限 - num库存数量, 0)
                    End If
                End If
                
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                            , "" _
                            , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
            Case 4  '日销售量
                dat前期 = Choose(int计划类型, DateAdd("m", -2, mstrNow), DateAdd("m", -6, mstrNow), DateAdd("yyyy", -2, mstrNow), DateAdd("d", -14, mdatThisMondyDate))
                dat上期 = Choose(int计划类型, DateAdd("m", -1, mstrNow), DateAdd("m", -3, mstrNow), DateAdd("yyyy", -1, mstrNow), DateAdd("d", -7, mdatThisMondyDate))
                GetDate int计划类型, dat上期, strBegin, strEnd
                lng天数 = CDate(Format(strEnd, "yyyy-MM-DD")) - CDate(Format(strBegin, "yyyy-MM-DD")) + 1
                If lng天数 <= 0 Then lng天数 = 1
                
                If lng库房ID = 0 Then
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0 = [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = "SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                           & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and 库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] "
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(strBegin), CDate(strEnd))
                    
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
                num上限 = num上期数量 / lng天数 * mint上限
                num下限 = num上期数量 / lng天数 * mint下限
                '计划数量=2×上期数量－前期数量－库存数量
                
                If mbln计划数量 Then
                    If num库存数量 < num下限 Then
                        num计划数量 = num上限 - num库存数量
                    Else
                        num计划数量 = 0
                    End If
                    If num计划数量 < 0 Then num计划数量 = 0
                End If
                .TextMatrix(intCurrentRow, mconIntCol前期数量) = zlStr.FormatEx(num前期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
            Case 5  '自定义区间
                If mstrBeginDate = "" Or mstrEndDate = "" Then
                    mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
                    mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Select Max(日期) As 日期 From 药品收发汇总"
                Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption)
                If NVL(rsNum!日期, "") = "" Then
                    str汇总最大日期 = Format(DateAdd("d", 1, CDate(mstrBeginDate)), "yyyy-mm-dd")
                Else
                    str汇总最大日期 = Format(DateAdd("d", 1, rsNum!日期), "yyyy-mm-dd")
                End If
                
                str收发结束时间 = Format(DateAdd("d", 1, CDate(mstrEndDate)), "yyyy-mm-dd")
                
                If lng库房ID = 0 Then
                    gstrSQL = " Select Sum(上期数量) As 上期数量 " _
                            & " From (SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                            & " Where a.类别id = b.id " _
                            & "  and 单据 <>6 AND b.系数 = -1 " _
                            & "  AND 药品id+0= [1] " _
                            & "  AND 日期 BETWEEN [2] and [3] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.入出系数 * Nvl(A.实际数量, 0) * Nvl(A.付数, 1))) As 上期数量 " _
                            & " From 药品收发记录 A, 药品入出类别 B " _
                            & " Where A.单据<>6 And A.入出类别id = B.ID And B.系数 = -1 And 药品id + 0 = [1] And " _
                            & " 审核日期 >= [2] " _
                            & " And 审核日期 Between [4] And [5])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str汇总最大日期), CDate(str收发结束时间))
                            
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                Else
                    gstrSQL = " Select Sum(上期数量) As 上期数量 " _
                            & " From (SELECT ABS(SUM(NVL(数量, 0))) AS 上期数量 " _
                            & " FROM 药品收发汇总 a, 药品入出类别 b " _
                            & " Where a.类别id = b.id " _
                            & "  AND b.系数 = -1 " _
                            & "  and a.库房id+0=[1] " _
                            & "  AND 药品id+0= [2] " _
                            & "  AND 日期 BETWEEN [3] and [4] " _
                            & " Union All " _
                            & " Select Abs(Sum(A.入出系数 * Nvl(A.实际数量, 0) * Nvl(A.付数, 1))) As 上期数量 " _
                            & " From 药品收发记录 A, 药品入出类别 B " _
                            & " Where A.入出类别id = B.ID And B.系数 = -1 And A.库房id + 0 = [1] And 药品id + 0 = [2] And " _
                            & " 审核日期 >= [3] " _
                            & " And 审核日期 Between [5] And [6])"
                    Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng库房ID, lng药品id, CDate(mstrBeginDate), CDate(mstrEndDate), CDate(str汇总最大日期), CDate(str收发结束时间))
                            
                    If rsNum.EOF Then
                        num上期数量 = 0
                    Else
                        num上期数量 = IIf(IsNull(rsNum.Fields(0)), 0, rsNum.Fields(0))
                    End If
                    rsNum.Close
                End If
    
                '把各单位转换成药库单位先
                num上期数量 = num上期数量 / .TextMatrix(intCurrentRow, mconIntCol比例系数)
    
                If mbln计划数量 Then
                    If num上期数量 > num库存数量 Then
                        num计划数量 = num上期数量 - num库存数量
                    End If
                End If
                .TextMatrix(intCurrentRow, mconIntCol前期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
        End Select
        
        '取库房上、下限
        If lng库房ID = 0 Then
            gstrSQL = "select sum(Nvl(上限,0)) as  上限,sum(Nvl(下限,0)) as 下限 " _
                    & " from 药品储备限额 " _
                    & " where 药品id=[1] "
        
        Else
            gstrSQL = "select Nvl(上限,0) As 上限,Nvl(下限,0) As 下限 " _
                    & " from 药品储备限额 " _
                    & " where 药品id=[1] " _
                    & "   and 库房id=[2]"
        End If
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, lng库房ID)
        
        If Not rsNum.EOF Then
            .TextMatrix(intCurrentRow, mconIntCol库存上限) = zlStr.FormatEx(NVL(rsNum!上限, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
            .TextMatrix(intCurrentRow, mconIntCol库存下限) = zlStr.FormatEx(NVL(rsNum!下限, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
        End If
        
        '分别计算上期和本期的销售量
        '取上期的区间范围
        Select Case int计划类型
            '1:月度计划,2.季度计划,3.年度计划,4.周计划
            Case 1
                '上月时间范围
                strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
                strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
            Case 2
                '上季度时间范围
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-10-01"
                        strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                        strEnd = Format(mstrNow, "YYYY") & "-03-31 23:59:59"
                     Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                        strEnd = Format(mstrNow, "YYYY") & "-06-30 23:59:59"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                        strEnd = Format(mstrNow, "YYYY") & "-09-30 23:59:59"
                End Select
            Case 3
                '上年度时间范围
                strBegin = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-01-01"
                strEnd = Format(DateAdd("yyyy", -1, CDate(mstrNow)), "YYYY") & "-12-31 23:59:59"
            Case 4
                '上周时间范围（按中国习惯，周一是一个星期的第一天）
                strBegin = Format(DateAdd("d", -DatePart("w", CDate(mstrNow), vbMonday) + 1, DateAdd("d", -7, CDate(mstrNow))), "YYYY-MM-DD")
                strEnd = Format(DateAdd("d", 6, CDate(strBegin)), "YYYY-MM-DD") & " 23:59:59"
        End Select
        
        '计算上期销售量（不要求精确值，用药品收发汇总统计）
        gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
            " From 药品收发汇总" & _
            " Where 单据 + 0 In (8, 9, 10) And 药品id+0=[1] And 日期 Between [2] And [3] "
        If mstr来源药房 <> "" Then
            gstrSQL = gstrSQL & " And 库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
        End If
        
        If int编制方法 = 5 Then
            '自定义区间，上期销量改为上月销量
            strBegin = Format(DateAdd("m", -1, CDate(mstrNow)), "YYYY-MM") & "-01"
            strEnd = Format(DateAdd("d", -1, CDate(Format(CDate(mstrNow), "YYYY-MM") & "-01")), "YYYY-MM-DD") & " 23:59:59"
        End If
        
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd), mstr来源药房)
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mconintCol上期销量) = zlStr.FormatEx(NVL(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
        End If
        
        '往年同期线形参照,取去年同期销量
        If int编制方法 = 1 Then
            dat上期 = DateAdd("m", Choose(int计划类型, 1, 3), DateAdd("yyyy", -1, mstrNow))
            GetDate int计划类型, dat上期, strBegin, strEnd
            
            gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
                " From 药品收发汇总" & _
                " Where 单据 + 0 In (8, 9, 10) And 药品id+0=[1] And 日期 Between [2] And [3] "
            If mstr来源药房 <> "" Then
                gstrSQL = gstrSQL & " And 库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
            End If
            
            Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd), mstr来源药房)
            If rsNum.RecordCount > 0 Then
                num上期销量 = zlStr.FormatEx(NVL(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
            End If
        End If
        
        '取本期的区间范围
        Select Case int计划类型
            '1:月度计划,2.季度计划,3.年度计划,4.周计划
            Case 1
                '本月时间范围
                strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            Case 2
                '本季度时间范围
                Select Case DatePart("Q", CDate(mstrNow))
                    Case 1
                        strBegin = Format(mstrNow, "YYYY") & "-01-01"
                    Case 2
                        strBegin = Format(mstrNow, "YYYY") & "-04-01"
                    Case 3
                        strBegin = Format(mstrNow, "YYYY") & "-07-01"
                    Case 4
                        strBegin = Format(mstrNow, "YYYY") & "-10-01"
                End Select
            Case 3
                '本年度时间范围
                strBegin = Format(mstrNow, "YYYY") & "-01-01"
            Case 4
                '本周时间范围（按中国习惯，周一是一个星期的第一天）
                strBegin = Format(DateAdd("d", -DatePart("w", CDate(mstrNow), vbMonday) + 1, CDate(mstrNow)), "YYYY-MM-DD")
        End Select
        
        '本期结束时间截止到当日
        strEnd = Format(mstrNow, "YYYY-MM-DD") & " 23:59:59"
            
        '计算本期销售量（不要求精确值，用药品收发汇总统计）
        gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
            " From 药品收发汇总" & _
            " Where 单据 + 0 In (8, 9, 10) And 药品id+0=[1] And 日期 Between [2] And [3] "
        If mstr来源药房 <> "" Then
            gstrSQL = gstrSQL & " And 库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
        End If
        
        If int编制方法 = 5 Then
            '自定义区间，本期销量改为本月销量
            strBegin = Format(mstrNow, "YYYY-MM") & "-01"
            strEnd = Format(mstrNow, "YYYY-MM-DD")
        End If
        
        Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(strBegin), CDate(strEnd), mstr来源药房)
        If rsNum.RecordCount > 0 Then
            .TextMatrix(intCurrentRow, mconintCol本期销量) = zlStr.FormatEx(NVL(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
        End If
        
        '自定义区间法，单独计算上期销量
        If int编制方法 = 5 Then
            If mstrBeginDate = "" Or mstrEndDate = "" Then
                mstrBeginDate = Format(DateAdd("m", -1, mstrNow), "yyyy-mm-dd")
                mstrEndDate = Format(mstrNow, "yyyy-mm-dd")
            End If
            
            gstrSQL = "Select -Sum(Nvl(数量, 0)) As 销售数量 " & _
                " From 药品收发汇总" & _
                " Where 单据 + 0 In (8, 9, 10) And 药品id+0=[1] And 日期 Between [2] And [3] "
            If mstr来源药房 <> "" Then
                gstrSQL = gstrSQL & " And 库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList)))"
            End If
            
            '自定义区间法，上期数量改为上期销量
            Set rsNum = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id, CDate(mstrBeginDate), CDate(mstrEndDate), mstr来源药房)
            If rsNum.RecordCount > 0 Then
                .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(NVL(rsNum!销售数量, 0) / Val(.TextMatrix(intCurrentRow, mconIntCol比例系数)), mintShowNumberDigit, , True)
            End If
        End If
        
        '按销量产生计划数量
        If mbln按销量产生计划 Then
            If mbln计划数量 Then
                If int编制方法 = 5 Then
                    num上期数量 = Val(.TextMatrix(intCurrentRow, mconIntCol上期数量))
                ElseIf int编制方法 = 1 Then
                    num上期数量 = num上期销量
                Else
                    num上期数量 = Val(.TextMatrix(intCurrentRow, mconintCol上期销量))
                End If
                
                If num上期数量 > num库存数量 Then
                    num计划数量 = num上期数量 - num库存数量
                Else
                    num计划数量 = 0
                End If
                               
'                .TextMatrix(intCurrentRow, mconIntCol上期数量) = zlStr.FormatEx(num上期数量, mintShowNumberDigit, , True)
                .TextMatrix(intCurrentRow, mconintCol计划数量) = IIf(zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(num计划数量, mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol成本金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconintCol成本价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconintCol成本价)), mintShowNumberDigit, , True))
                .TextMatrix(intCurrentRow, mconIntCol售价金额) = _
                        IIf(zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True) = 0 _
                        , "" _
                        , zlStr.FormatEx(num计划数量 * IIf(mshBill.TextMatrix(intCurrentRow, mconIntCol售价) = "", 0, mshBill.TextMatrix(intCurrentRow, mconIntCol售价)), mintShowNumberDigit, , True))
            End If
        End If
        
        '取历史采购计划
        Call LoadHisPlane(lng库房ID, lng药品id, intCurrentRow)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsDetail As New Recordset
    Dim bln中药库房 As Boolean
    Dim str库房性质 As String
    Dim str屏蔽列 As String
    Dim i As Integer, j As Integer
    
    On Error GoTo errHandle
    
    marrFrom = Array()
    marrInitGrid = Array()
    Call GetDefineSize
    Call GetDrugDigit(mlng库房ID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")
    '组织格式化串
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    mintPlanPoint = Val(zlDataBase.GetPara("全院计划不管站点", glngSys, 1330, 0))

    mintPriceUnit = GetUnit()
    txtNo = mstr单据号
    txtNo.Tag = txtNo
    mblnEnter = True
    mstrNow = Format(Sys.Currentdate, "yyyy-mm-dd")
        
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品计划管理", "药品名称显示方式", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GegReg
    
    IniHisPlaneRec
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    initCard
    
    '只有中药类库房才显示"原产地"列
    str库房性质 = ""
    gstrSQL = "Select 工作性质 From 部门性质说明 Where 部门id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "判断库房性质", mlng库房ID)
    Do While Not rsDetail.EOF
        str库房性质 = str库房性质 & "," & rsDetail!工作性质
        rsDetail.MoveNext
    Loop
    If str库房性质 Like "*中药*" Or str库房性质 Like "*制剂室*" Then bln中药库房 = True
    str屏蔽列 = zlDataBase.GetPara("屏蔽列", glngSys, 模块号.药品计划)
    If InStr(1, "|" & str屏蔽列 & "|", "|原产地|") = 0 Then mshBill.ColWidth(mconIntCol原产地) = IIf(bln中药库房, 800, 0)
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    '恢复个性化设置
    RestoreWinState Me, App.ProductName, MStrCaption

    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next
    
    '恢复个性化设置后需要重新根据权限控制列是否显示
    With mshBill
        If mblnViewCost = False Then
            .ColWidth(mconintCol成本价) = 0
            .ColWidth(mconIntCol成本金额) = 0
        End If
    End With
    
    If (mint编辑状态 = 4 Or mint编辑状态 = 6) And Trim(Txt审核人.Caption) <> "" Then
        If mshBill.ColWidth(mconintCol执行数量) = 0 And InStr(1, "|" & mstrColumn_UnSelected & "|", "|执行数量|") = 0 Then mshBill.ColWidth(mconintCol执行数量) = 1100
    Else
        mshBill.ColWidth(mconintCol执行数量) = 0
    End If
    
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If
    
'    If grsMaster.State = adStateClosed Then
'        Call SetSelectorRS(1, mstrCaption, mlng库房ID, mlng库房ID)
'    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim intRecordCount As Integer
    Dim str单位 As String
    Dim strOrder As String, strCompare As String
    Dim str药名 As String
    Dim strSqlOrder As String
    Dim str送货单位 As String
    Dim dbl送货包装 As Double
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("排序", glngSys, 模块号.药品计划)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "序号"
    
    If strCompare = "0" Then
        strSqlOrder = "序号"
    ElseIf strCompare = "1" Then
        strSqlOrder = "药品编码"
    ElseIf strCompare = "2" Then
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            strSqlOrder = "通用名"
        Else
            strSqlOrder = "Nvl(商品名, 通用名)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")

    '库房
    Select Case mint编辑状态
        Case 1
            Txt填制人 = UserInfo.用户姓名
            Txt填制日期 = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 3, 4, 5, 6
            Select Case mintUnit
            Case 1
                str单位 = ",d.售价单位 单位,d.售价包装 比例系数"
            Case 2
                str单位 = ",d.门诊单位 单位,d.门诊包装 比例系数"
            Case 3
                str单位 = ",d.住院单位 单位,d.住院包装 比例系数"
            Case 4
                str单位 = ",d.药库单位 单位,d.药库包装 比例系数"
            End Select
            
            gstrSQL = "SELECT a.id,nvl(a.库房id,0) as 库房id,nvl(c.名称,'全院') AS 库房,a.药房id, a.no, a.计划类型,a.期间, a.编制方法, a.编制人," _
                    & "TO_CHAR (a.编制日期, 'yyyy-mm-dd HH24:MI:SS') AS 编制日期, a.审核人," _
                    & "TO_CHAR (a.审核日期, 'yyyy-mm-dd HH24:MI:SS') AS 审核日期,a.复核人,TO_CHAR (a.复核日期, 'yyyy-mm-dd HH24:MI:SS') AS 复核日期,a.编制说明," _
                    & "b.序号,b.药品id,d.药品编码,d.通用名,d.商品名,d.药品来源, d.规格,d.基本药物" & str单位 & ", nvl(b.前期数量,0) as 前期数量, nvl(b.上期数量,0) as 上期数量, " _
                    & " nvl(b.上期销量,0) as 上期销量,nvl(b.本期销量,0) as 本期销量,b.库存数量, b.计划数量,nvl(b.执行数量,0) as 执行数量,b.送货数量,d.送货单位,d.送货包装, b.单价, b.金额, b.上次供应商,b.上次生产商,d.原产地,b.说明,b.售价,b.售价金额,d.费用类型,b.批准文号 " _
                    & " FROM 药品采购计划 a, 药品计划内容 b,部门表 c," _
                    & " (SELECT DISTINCT a.药品id," _
                    & " '[' || C.编码 || ']' As 药品编码, C.名称 As 通用名, a.原产地,B.名称 As 商品名,a.药品来源,c.规格,a.药库单位,A.药库包装,a.基本药物,a.门诊单位,A.门诊包装,a.住院单位,a.住院包装,a.送货单位,a.送货包装,C.计算单位 售价单位,1 售价包装,c.费用类型 " _
                    & " FROM 药品规格 a, 收费项目别名 b, 收费项目目录 c " _
                    & " WHERE a.药品id = b.收费细目ID(+) and B.性质(+)=3 " _
                    & "   AND a.药品id = c.ID) d " _
                    & "Where a.id = b.计划id " _
                    & "  and nvl(a.库房id,0)=c.id(+) " _
                    & "  and b.药品id=d.药品id " _
                    & "  AND a.no = [1] " & _
                    " Order by " & strSqlOrder
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr单据号)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If

            intRecordCount = rsInitCard.RecordCount

            Txt填制人 = rsInitCard!编制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户姓名
            End If
            Txt填制日期 = Format(rsInitCard!编制日期, "yyyy-mm-dd hh:mm:ss")

            Txt审核人 = IIf(IsNull(rsInitCard!审核人), "", rsInitCard!审核人)
            Txt审核日期 = IIf(IsNull(rsInitCard!审核日期), "", Format(rsInitCard!审核日期, "yyyy-mm-dd hh:mm:ss"))
            
            txt复核人 = IIf(IsNull(rsInitCard!复核人), "", rsInitCard!复核人)
            txt复核日期 = IIf(IsNull(rsInitCard!复核日期), "", Format(rsInitCard!复核日期, "yyyy-mm-dd hh:mm:ss"))
            
            txt摘要.Text = IIf(IsNull(rsInitCard!编制说明), "", rsInitCard!编制说明)
            txt摘要.Tag = NVL(rsInitCard!药房id)
            txt计划类型 = Choose(rsInitCard!计划类型 + 1, "临时", "月度计划", "季度计划", "年度计划", "周计划")
            txt编制方法 = Choose(rsInitCard!编制方法 + 1, "根据申领产生", "往年同期线形参照法", "临近期间平均参照法", "药品储备定额参照法", "药品日销售量参照法", "自定义区间参照法")
            mint计划类型 = rsInitCard!计划类型
            mint编制方法 = rsInitCard!编制方法
            mlng库房ID = rsInitCard!库房id
            mlng计划ID = rsInitCard!Id

            Str期间 = IIf(IsNull(rsInitCard!期间), "", rsInitCard!期间)
            Select Case mint计划类型
                Case 0       '临时计划
                    LblTitle.Caption = GetUnitName & rsInitCard!库房 & "采购计划"
                Case 1       '月计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Right(Str期间, 2) & "月" & ") " & rsInitCard!库房 & "采购计划"
                Case 2       '季计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Right(Str期间, 1) & "季" & ")" & rsInitCard!库房 & "采购计划"
                Case 3       '年计划
                    LblTitle.Caption = GetUnitName & "(" & Str期间 & "年" & ")" & rsInitCard!库房 & "采购计划"
                Case 4       '周计划
                    LblTitle.Caption = GetUnitName & "(" & Mid(Str期间, 1, 4) & "年" & Mid(Str期间, 5, 2) & "月" & Right(Str期间, 2) & "日" & ")" & LblTitle.Tag & "采购计划"
            End Select

            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            'If IIf(IsNull(rsInitCard!单价), 0, rsInitCard!单价) = 0 Then
            '    mint价格显示 = 1
            'Else
            '    mint价格显示 = 0
            'End If
            
            initGrid
            
            
            If mint编制方法 = 5 Then
                '自定义区间编制法
                mshBill.TextMatrix(0, mconIntCol前期数量) = "本期数量"
                mshBill.TextMatrix(0, mconIntCol上期数量) = "本期销量"
                mshBill.TextMatrix(0, mconintCol上期销量) = "上月销量"
                mshBill.TextMatrix(0, mconintCol本期销量) = "本月销量"
            End If
            
            With mshBill
                For intRow = 1 To intRecordCount

                    .TextMatrix(intRow, 0) = rsInitCard!药品id
                    
                    If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                        str药名 = rsInitCard!通用名
                    Else
                        str药名 = IIf(IsNull(rsInitCard!商品名), rsInitCard!通用名, rsInitCard!商品名)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol药品编码和名称) = rsInitCard!药品编码 & str药名
                    .TextMatrix(intRow, mconIntCol药品编码) = rsInitCard!药品编码
                    .TextMatrix(intRow, mconIntCol药品名称) = str药名
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品名称)
                    Else
                        .TextMatrix(intRow, mconIntCol药名) = .TextMatrix(intRow, mconIntCol药品编码和名称)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol商品名) = IIf(IsNull(rsInitCard!商品名), "", rsInitCard!商品名)

                    .TextMatrix(intRow, mconIntCol来源) = NVL(rsInitCard!药品来源)
                    .TextMatrix(intRow, mconIntCol规格) = IIf(IsNull(rsInitCard!规格), "", rsInitCard!规格)
                    .TextMatrix(intRow, mconintCol上次供应商) = IIf(IsNull(rsInitCard!上次供应商), "", rsInitCard!上次供应商)
                    .TextMatrix(intRow, mconIntCol生产商) = IIf(IsNull(rsInitCard!上次生产商), "", rsInitCard!上次生产商)
                    .TextMatrix(intRow, mconIntCol原产地) = IIf(IsNull(rsInitCard!原产地), "", rsInitCard!原产地)
                    .TextMatrix(intRow, mconIntCol单位) = rsInitCard!单位
                    .TextMatrix(intRow, mconIntcol医保类型) = IIf(IsNull(rsInitCard!费用类型), "", rsInitCard!费用类型)
                    .TextMatrix(intRow, mconIntCol比例系数) = rsInitCard!比例系数
                    .TextMatrix(intRow, mconIntCol前期数量) = zlStr.FormatEx(rsInitCard!前期数量 / rsInitCard!比例系数, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol上期数量) = zlStr.FormatEx(rsInitCard!上期数量 / rsInitCard!比例系数, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol上期销量) = zlStr.FormatEx(rsInitCard!上期销量 / rsInitCard!比例系数, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol本期销量) = zlStr.FormatEx(rsInitCard!本期销量 / rsInitCard!比例系数, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol库存数量) = zlStr.FormatEx(rsInitCard!库存数量 / rsInitCard!比例系数, mintShowNumberDigit, , True)
                    .TextMatrix(intRow, mconintCol计划数量) = IIf(zlStr.FormatEx(rsInitCard!计划数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!计划数量 / rsInitCard!比例系数, mintShowNumberDigit, , True))
                    .TextMatrix(intRow, mconintCol执行数量) = IIf(zlStr.FormatEx(rsInitCard!执行数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!执行数量 / rsInitCard!比例系数, mintShowNumberDigit, , True))
                    .TextMatrix(intRow, mconintCol原执行数量) = IIf(zlStr.FormatEx(rsInitCard!执行数量, mintShowNumberDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!执行数量 / rsInitCard!比例系数, mintShowNumberDigit, , True))
                    
                    dbl送货包装 = IIf(IsNull(rsInitCard!送货包装), 0, rsInitCard!送货包装)
                    str送货单位 = IIf(IsNull(rsInitCard!送货单位), "", rsInitCard!送货单位)
                    If dbl送货包装 <> 0 Then
                        .TextMatrix(intRow, mconintCol送货数量) = IIf(IsNull(rsInitCard!送货数量), "", zlStr.FormatEx(rsInitCard!送货数量, 1, , True))
                        .TextMatrix(intRow, mconintCol送货单位) = str送货单位 & "(1" & str送货单位 & "=" & zlStr.FormatEx(dbl送货包装 / rsInitCard!比例系数, 1, , True) & rsInitCard!单位 & ")"
                        .TextMatrix(intRow, mconintCol送货包装) = dbl送货包装
                    End If
                    
                    .TextMatrix(intRow, mconintCol成本价) = zlStr.FormatEx(rsInitCard!单价 * rsInitCard!比例系数, mintShowPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol成本金额) = IIf(zlStr.FormatEx(rsInitCard!金额, mintShowMoneyDigit, , True) = 0, "", zlStr.FormatEx(rsInitCard!金额, mintShowMoneyDigit, , True))
                    .TextMatrix(intRow, mconIntCol售价) = zlStr.FormatEx(rsInitCard!售价 * rsInitCard!比例系数, mintShowPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol售价金额) = IIf(zlStr.FormatEx(rsInitCard!售价金额, mintShowMoneyDigit) = 0, "", zlStr.FormatEx(rsInitCard!售价金额, mintShowMoneyDigit, , True))
                    
                    .TextMatrix(intRow, mconintCol说明) = IIf(IsNull(rsInitCard!说明), "", rsInitCard!说明)
                    .TextMatrix(intRow, mconIntCol基本药物) = IIf(IsNull(rsInitCard!基本药物), "", rsInitCard!基本药物)
                    .TextMatrix(intRow, mconIntCol批准文号) = IIf(IsNull(rsInitCard!批准文号), "", rsInitCard!批准文号)
                    If intRow = .rows - 1 Then .rows = .rows + 1
                    rsInitCard.MoveNext
                    
                    Call LoadHisPlane(mlng库房ID, Val(.TextMatrix(intRow, 0)), intRow)
                Next
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol序号, 1)
    Call 显示合计金额
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'初始化编辑控件
Private Sub initGrid()
    Dim intCol As Integer

    Call SetColumnByUserDefine '列设置
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 2

        .TextMatrix(0, mconIntCol序号) = "序号"
        .TextMatrix(0, mconIntCol药名) = "药品名称与编码"
        .TextMatrix(0, mconIntCol商品名) = "商品名"
        .TextMatrix(0, mconIntCol来源) = "药品来源"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol生产商) = "生产商"
        .TextMatrix(0, mconIntCol原产地) = "原产地"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconIntCol比例系数) = "比例系数"
        .TextMatrix(0, mconIntcol医保类型) = "医保类型"
        .TextMatrix(0, mconIntCol前期数量) = "前期数量"
        .TextMatrix(0, mconIntCol上期数量) = "上期数量"
        .TextMatrix(0, mconIntCol库存上限) = "库存上限"
        .TextMatrix(0, mconIntCol库存下限) = "库存下限"
        .TextMatrix(0, mconintCol库存数量) = "库存数量"
        .TextMatrix(0, mconintCol上期销量) = "上期销量"
        .TextMatrix(0, mconintCol本期销量) = "本期销量"
        .TextMatrix(0, mconintCol计划数量) = "计划数量"
        .TextMatrix(0, mconintCol执行数量) = "执行数量"
        .TextMatrix(0, mconintCol原执行数量) = "原执行数量"
        .TextMatrix(0, mconintCol送货单位) = "送货单位"
        .TextMatrix(0, mconintCol送货数量) = "送货数量"
        
        .TextMatrix(0, mconintCol成本价) = "成本价"
        .TextMatrix(0, mconIntCol成本金额) = "成本金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        
        .TextMatrix(0, mconintCol上次供应商) = "上次供应商"
        .TextMatrix(0, mconintCol说明) = "说明"
        .TextMatrix(0, mconIntCol药品编码和名称) = "药品编码和名称"
        .TextMatrix(0, mconIntCol药品编码) = "药品编码"
        .TextMatrix(0, mconIntCol药品名称) = "药品名称"
        .TextMatrix(0, mconIntCol基本药物) = "基本药物"
        .TextMatrix(0, mconIntCol批准文号) = "批准文号"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol序号) = "1"
    
        .ColWidth(mconIntCol序号) = 500
        .ColWidth(mconIntCol药名) = 2000
        .ColWidth(mconIntCol商品名) = 2000
        .ColWidth(mconIntCol来源) = 1000
        .ColWidth(mconIntCol规格) = 900
        .ColWidth(mconIntCol生产商) = 800
        .ColWidth(mconIntCol原产地) = 0
        .ColWidth(mconIntCol单位) = 500
        .ColWidth(mconIntcol医保类型) = 1000
        .ColWidth(mconIntCol前期数量) = 1100
        .ColWidth(mconIntCol上期数量) = 1100
        .ColWidth(mconIntCol库存上限) = 1100
        .ColWidth(mconIntCol库存下限) = 1100
        .ColWidth(mconintCol库存数量) = 1100
        .ColWidth(mconintCol上期销量) = 1100
        .ColWidth(mconintCol本期销量) = 1100
        .ColWidth(mconintCol计划数量) = 1100
        .ColWidth(mconintCol执行数量) = IIf(mint编辑状态 = 6, 1100, 0)
        .ColWidth(mconintCol原执行数量) = 0
        .ColWidth(mconintCol送货单位) = 1500
        .ColWidth(mconintCol送货数量) = 1100
        .ColWidth(mconintCol送货包装) = 0
        
        If mint价格显示 = 0 Then
            .ColWidth(mconintCol成本价) = 1000
            .ColWidth(mconIntCol成本金额) = 1200
            .ColWidth(mconIntCol售价) = 0
            .ColWidth(mconIntCol售价金额) = 0
        ElseIf mint价格显示 = 1 Then
            .ColWidth(mconintCol成本价) = 0
            .ColWidth(mconIntCol成本金额) = 0
            .ColWidth(mconIntCol售价) = 1000
            .ColWidth(mconIntCol售价金额) = 1200
        Else
            .ColWidth(mconintCol成本价) = 1000
            .ColWidth(mconIntCol成本金额) = 1200
            .ColWidth(mconIntCol售价) = 1000
            .ColWidth(mconIntCol售价金额) = 1200
        End If
        If mblnViewCost = False Then
            .ColWidth(mconintCol成本价) = 0
            .ColWidth(mconIntCol成本金额) = 0
        End If
        
        .ColWidth(mconintCol上次供应商) = 2000
        .ColWidth(mconIntCol比例系数) = 0
        .ColWidth(mconintCol说明) = 3000
        .ColWidth(0) = 0
        .ColWidth(mconIntCol药品编码和名称) = 0
        .ColWidth(mconIntCol药品编码) = 0
        .ColWidth(mconIntCol药品名称) = 0
        .ColWidth(mconIntCol基本药物) = 2000
        .ColWidth(mconIntCol批准文号) = 2000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择
        For intCol = 0 To .Cols - 1
            .ColData(intCol) = 5
        Next
        
        .ColData(mconintCol送货单位) = 5
        .ColData(mconintCol送货数量) = 5
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            txt摘要.Enabled = True
            .ColData(mconIntCol药名) = 1
            .ColData(mconintCol计划数量) = 4
            .ColData(mconintCol成本价) = 4
            .ColData(mconIntCol生产商) = 4
            .ColData(mconIntCol原产地) = 4
            .ColData(mconintCol上次供应商) = 1
            .ColData(mconintCol说明) = 4
            .ColData(mconintCol送货数量) = 4
            .ColData(mconIntCol批准文号) = 4
        ElseIf mint编辑状态 = 4 Then
            txt摘要.Enabled = False
            .ColData(mconintCol计划数量) = 0
        ElseIf mint编辑状态 = 3 Then
            txt摘要.Enabled = False
            .ColData(mconintCol计划数量) = 4
            .ColData(mconintCol说明) = 4
        ElseIf mint编辑状态 = 6 Then
            .ColData(mconintCol执行数量) = 4
        End If
        
        .ColAlignment(mconIntCol药名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol商品名) = flexAlignLeftCenter
        .ColAlignment(mconIntCol来源) = flexAlignLeftCenter
        .ColAlignment(mconIntCol规格) = flexAlignLeftCenter
        .ColAlignment(mconIntCol生产商) = flexAlignLeftCenter
        .ColAlignment(mconIntCol原产地) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconIntcol医保类型) = flexAlignLeftCenter
        .ColAlignment(mconIntCol前期数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol上期数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol库存上限) = flexAlignRightCenter
        .ColAlignment(mconIntCol库存下限) = flexAlignRightCenter
        .ColAlignment(mconintCol库存数量) = flexAlignRightCenter
        .ColAlignment(mconintCol上期销量) = flexAlignRightCenter
        .ColAlignment(mconintCol本期销量) = flexAlignRightCenter
        .ColAlignment(mconintCol计划数量) = flexAlignRightCenter
        .ColAlignment(mconintCol执行数量) = flexAlignRightCenter
        .ColAlignment(mconintCol送货单位) = flexAlignCenterCenter
        .ColAlignment(mconintCol送货数量) = flexAlignRightCenter
        .ColAlignment(mconintCol成本价) = flexAlignRightCenter
        .ColAlignment(mconIntCol成本金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        .ColAlignment(mconintCol上次供应商) = flexAlignLeftCenter
        .ColAlignment(mconintCol说明) = flexAlignLeftCenter
        .ColAlignment(mconIntCol基本药物) = flexAlignLeftCenter
        .ColAlignment(mconIntCol批准文号) = flexAlignLeftCenter
        
        .PrimaryCol = mconIntCol药名
        .LocateCol = mconIntCol药名
        Call SetColumnByUserDefine '列设置
        If InStr(1, "34", mint编辑状态) <> 0 Then .ColData(mconIntCol药名) = 0
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub

    With Pic单据
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 100
    End With

    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic单据.Width
    End With


    With mshBill
        .Left = 200
        .Width = Pic单据.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With

    txt编制方法.Left = mshBill.Left + mshBill.Width - txt编制方法.Width
    lbl编制方法.Left = txt编制方法.Left - lbl编制方法.Width - 100


    Lbl计划类型.Left = mshBill.Left

    txt计划类型.Left = Lbl计划类型.Left + Lbl计划类型.Width + 100

    With Lbl填制日期
        .Top = Pic单据.Height - 100 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt填制日期
        .Top = Lbl填制日期.Top - 60
        .Left = Lbl填制日期.Left + Lbl填制日期.Width + 100
    End With
    
    With Lbl填制人
        .Top = Lbl填制日期.Top - Lbl填制日期.Height - 180
        .Left = Lbl填制日期.Left
    End With
    
    With Txt填制人
        .Top = Lbl填制人.Top - 60
        .Left = Txt填制日期.Left
    End With
    
    With Lbl审核日期
        .Top = Lbl填制日期.Top
        .Left = mshBill.Left + (mshBill.Width - .Width - Txt审核日期.Width - 100) / 2
    End With
    
    With Txt审核日期
        .Top = Txt填制日期.Top
        .Left = Lbl审核日期.Left + Lbl审核日期.Width + 100
    End With
    
    With Lbl审核人
        .Top = Lbl填制人.Top
        .Left = Lbl审核日期.Left
    End With
    
    With Txt审核人
        .Top = Txt填制人.Top
        .Left = Txt审核日期.Left
    End With
    
    With txt复核日期
        .Top = Txt填制日期.Top
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With lbl复核日期
        .Top = Lbl填制日期.Top
        .Left = txt复核日期.Left - 100 - .Width
    End With
    
    With txt复核人
        .Top = Txt填制人.Top
        .Left = txt复核日期.Left
    End With
    
    With lbl复核人
        .Top = Lbl填制人.Top
        .Left = lbl复核日期.Left
    End With
    
    
    With txt摘要
        .Top = Lbl填制人.Top - 140 - .Height
        .Left = Txt填制人.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With

    With lbl摘要
        .Top = txt摘要.Top + 50
        .Left = txt摘要.Left - .Width - 200
    End With

    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txt摘要.Top - 60 - .Height
        .Width = mshBill.Width
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With

    With CmdCancel
        .Left = Pic单据.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic单据.Top + Pic单据.Height + 100
    End With

    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With

    With cmdHelp
        .Left = Pic单据.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With

    With cmdFind
        .Left = cmdHelp.Left + cmdHelp.Width + 200
        .Top = CmdCancel.Top
    End With

    With lblCode
        .Left = cmdFind.Left + cmdFind.Width + 50
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Left = lblCode.Left + lblCode.Width + 50
        .Top = CmdCancel.Top + 30
    End With
    
    With chk隐藏近期采购计划
        .Left = txtCode.Left + txtCode.Width + 150
        .Top = CmdCancel.Top + 30
    End With

    With chk是否显示库存情况
        .Left = txtCode.Left + txtCode.Width + 150 + chk隐藏近期采购计划.Width
        .Top = CmdCancel.Top + 30
    End With
    
    Call ResizeHisPlane
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mblnChange = False Or mint编辑状态 = 4 Or mint编辑状态 = 3) Then
        If MsgBox("数据可能已改变，但未存盘，真要退出吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品计划管理", "药品名称显示方式", mintDrugNameShow)
    
    mstr自定义库房 = ""
    mstr来源库房 = ""
    mstr来源药房 = ""
    mstrAll来源库房 = ""
    mstrAll来源药房 = ""
    mblnStart = False
    
    Call ReleaseSelectorRS

    Set mfrmMain = Nothing
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    '药品名称显示：
    'intType：0－显示编码和名称；1－仅显示编码；2－仅显示名称
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol药名) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品名称)
                Else
                    .TextMatrix(lngRow, mconIntCol药名) = .TextMatrix(lngRow, mconIntCol药品编码和名称)
                End If
            End If
        Next
    End With
End Sub
Private Function SaveCheck() As Boolean
    Dim str审核人 As String

    mblnSave = False
    SaveCheck = False

    str审核人 = UserInfo.用户姓名

    On Error GoTo errHandle
    'zl_药品计划管理_VERIFY( /*ID_IN*/, /*审核人_IN*/ );
    gstrSQL = "zl_药品计划管理_VERIFY('" & mlng计划ID & "','" & str审核人 & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "审核失败！", vbInformation, gstrSysName
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function SaveReCheck() As Boolean
    Dim str审核人 As String

    mblnSave = False
    SaveReCheck = False

    str审核人 = UserInfo.用户姓名

    On Error GoTo errHandle
    'zl_药品计划管理_VERIFY( /*ID_IN*/, /*审核人_IN*/ );
    gstrSQL = "zl_药品计划管理_VERIFY('" & mlng计划ID & "','" & str审核人 & "',1)"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveReCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Function SaveExeAmount() As Boolean
    '修改执行数量
    Dim strInput As String
    Dim strNo As String
    Dim intRow As Integer

    mblnSave = False
    SaveExeAmount = False
    
    strNo = Trim(txtNo.Caption)
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconintCol执行数量)) - Val(.TextMatrix(intRow, mconintCol原执行数量)) <> 0 Then
                    strInput = IIf(strInput = "", "", strInput & "|") & Val(.TextMatrix(intRow, 0)) & "," & (Val(.TextMatrix(intRow, mconintCol执行数量)) - Val(.TextMatrix(intRow, mconintCol原执行数量))) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                End If
            End If
        Next
    End With
    
    On Error GoTo errHandle
    
    'Zl_药品计划内容_修改执行数量( /*No_In*/, /*Input_In*/ );
    '执行数量为数量差（相对于原执行数量，经本次录入后增加或减少的数量）
    gstrSQL = "Zl_药品计划内容_修改执行数量("
    'No_In
    gstrSQL = gstrSQL & "'" & strNo & "'"
    'Input_In  --格式:"药品ID1,执行数量1|药品ID2,执行数量2|....."
    gstrSQL = gstrSQL & ",'" & strInput & "'"
    gstrSQL = gstrSQL & ")"
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveExeAmount = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub Image1_Click()

End Sub

Private Sub imgDown_Click()
    picHscSend.Tag = 0
    imgUp.Visible = True
    imgDown.Visible = False
    
    Call ResizeHisPlane
End Sub

Private Sub imgUp_Click()
    picHscSend.Tag = 1
    imgUp.Visible = False
    imgDown.Visible = True
    
    Call ResizeHisPlane
End Sub


Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
End Sub

Private Sub Msf供应商选择_DblClick()
    Dim blnCancel As Boolean
    With mshBill
        .Text = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
        .TextMatrix(.Row, mconintCol上次供应商) = Msf供应商选择.TextMatrix(Msf供应商选择.Row, 2)
    End With
    Msf供应商选择.Visible = False
    mshBill.SetFocus
    If mshBill.Col <> mshBill.Cols - 1 Then
        mshBill.Col = mshBill.Col + 1
    End If
End Sub

Private Sub Msf供应商选择_GotFocus()
    If Msf供应商选择.rows - 1 = 1 Then Call Msf供应商选择_DblClick
End Sub

Private Sub Msf供应商选择_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Msf供应商选择_DblClick
    End If
End Sub

Private Sub Msf供应商选择_LostFocus()
    Msf供应商选择.ZOrder 1
    Msf供应商选择.Visible = False
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol序号, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call RefreshRowNO(mshBill, mconIntCol序号, mshBill.Row)
    Call 显示合计金额
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntCol药名) = 0 Then
        'Cancel = True    '等待加CANCEL参数
        Exit Sub
    End If
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint编辑状态) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("你确实要删除该行药品？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
            vsfStock.rows = 1
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim sngLeft As Single, sngTop As Single
    Dim RecReturn As Recordset
    Dim strUnit As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim lngRow As Long
    Dim strTemp As String
    
    intOldRow = mshBill.Row
    
    On Error GoTo errHandle
    If mshBill.Col = mconIntCol药名 Then
        mblnChange = True
'        Set RecReturn = Frm药品选择器.ShowME(Me, 1, , mlng库房ID)
        If grsMaster.State = adStateClosed Then
           Call SetSelectorRS(1, MStrCaption, mlng库房ID, mlng库房ID)
        End If
        Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , , mlng库房ID, , , , , , , , mstrPrivs)
        
        If RecReturn.RecordCount > 0 Then
            '检查重复记录 并将重复记录的药品id返回回来
            If RecReturn.RecordCount = 1 Then
                lngRow = CheckDouData(RecReturn)
                If lngRow > 0 Then
                    If MsgBox("该药品已经存在，是否跳转到记录行？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        mshBill.Row = lngRow
                        mshBill.Col = 0
                        mshBill.SetFocus
                    End If
                    Exit Sub
                End If
            Else
                Set RecReturn = CheckData(RecReturn)
            End If
        End If
                
        If RecReturn.RecordCount > 0 Then
            RecReturn.MoveFirst
            For i = 1 To RecReturn.RecordCount
                With mshBill
                    intCurRow = .Row
                    Select Case mintUnit
                    Case 1
                        strUnit = "售价单位"
                    Case 2
                        strUnit = "门诊单位"
                    Case 3
                        strUnit = "住院单位"
                    Case Else
                        strUnit = "药库单位"
                    End Select
                    
                    SetDrugRows RecReturn!药品id, "[" & RecReturn!药品编码 & "]", RecReturn!通用名, IIf(IsNull(RecReturn!商品名), "", RecReturn!商品名), NVL(RecReturn!药品来源), _
                                IIf(IsNull(RecReturn!规格), "", RecReturn!规格), IIf(IsNull(RecReturn!产地), "", RecReturn!产地), _
                                Switch(strUnit = "售价单位", RecReturn!售价单位, strUnit = "门诊单位", RecReturn!门诊单位, strUnit = "住院单位", RecReturn!住院单位, strUnit = "药库单位", RecReturn!药库单位), RecReturn!指导批发价, _
                                Switch(strUnit = "售价单位", 1, strUnit = "门诊单位", RecReturn!门诊包装, strUnit = "住院单位", RecReturn!住院包装, _
                                strUnit = "药库单位", RecReturn!药库包装), RecReturn!原产地
                    If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                End With
                RecReturn.MoveNext
            Next
            RecReturn.Close
            mshBill.Row = intCurRow
            mshBill.Col = mconintCol计划数量
        End If
    Else
        '药品供应商的选择
        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
        If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100

        Set RecReturn = New ADODB.Recordset
        
        '1、非全院计划时要管站点
        '全院计划未勾选参数时要管站点
        If mlng库房ID <> 0 Or (mlng库房ID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
            strTemp = "(站点 = [2] Or 站点 is Null) And "
        End If
        
        If mint供应商范围 = 1 Then
            gstrSQL = "Select A.ID,A.编码,A.名称,A.简码 From 供应商 A,药品中标单位 B " & _
                      "Where " & strTemp & " A.ID=B.单位ID And B.药品ID=[1] " & _
                      "  And (To_Char(B.撤档时间,'yyyy-MM-dd')='3000-01-01' or B.撤档时间 is null) " & _
                      "  And A.末级=1 And (substr(A.类型,1,1)=1 Or Nvl(A.末级,0)=0) " & _
                      "  And (To_Char(A.撤档时间,'yyyy-MM-dd')='3000-01-01' or A.撤档时间 is null) " & _
                      "Order By A.编码 "
        Else
            gstrSQL = "Select ID,编码,名称,简码 From 供应商 " & _
                      "Where " & strTemp & " 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                      "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                      "Order By 编码 "
        End If
        Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "读取药应商", Val(mshBill.TextMatrix(mshBill.Row, 0)), gstrNodeNo)
        If RecReturn.RecordCount = 0 Then
            If mint供应商范围 = 1 Then
                '如果没有设置中标单位，则提取所有供应商
                gstrSQL = "Select ID,编码,名称,简码 From 供应商 " & _
                          "Where " & strTemp & " 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                          "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                          "Order By 编码 "
                Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "读取药应商", Val(mshBill.TextMatrix(mshBill.Row, 0)), gstrNodeNo)
                
                If RecReturn.RecordCount = 0 Then
                    MsgBox "请先初始化药品供应商！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                MsgBox "请先初始化药品供应商！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        With Msf供应商选择
            .Clear
            Set .DataSource = RecReturn
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800

            .Row = 1
            .ColSel = .Cols - 1
        End With
        With Msf供应商选择
            .Left = sngLeft
            .Top = sngTop
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDouData(ByVal rsData As ADODB.Recordset) As Long
    '检查数据是否重复并范围重复数据所在行
    Dim lngRow As Long
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) = rsData!药品id Then
                CheckDouData = lngRow
                Exit Function
            End If
        Next
    End With
End Function

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
        If strkey = "" Then
            strkey = .TextMatrix(.Row, .Col)
        End If
        
        If .Col = mconintCol计划数量 Or .Col = mconintCol成本价 Then
            Select Case .Col
                Case mconintCol计划数量
                    intDigit = mintShowNumberDigit
                Case mconintCol成本价
                    intDigit = mintShowCostDigit
            End Select
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strkey) Then Exit Sub
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
        
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnter Then Exit Sub

    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntCol药名
                .txtCheck = False
                .MaxLength = 40
                '只在药名列才显示合计信息和库存数
                Call 显示合计金额
            Case mconIntCol生产商
                .txtCheck = False
                .MaxLength = mlng生产商长度
            Case mconIntCol原产地
                .txtCheck = False
                .MaxLength = mlng原产地长度
            Case mconintCol上次供应商
                .MaxLength = 40
                .txtCheck = False
            Case mconintCol计划数量, mconintCol执行数量
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mconintCol成本价, mconIntCol售价
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconintCol说明
                .txtCheck = False
                .MaxLength = 50
            Case mconIntCol批准文号
                .txtCheck = False
                .MaxLength = 40
            Case mconintCol送货数量
                .txtCheck = True
                .MaxLength = 10
                .TextMask = ".1234567890"
                If .TextMatrix(Row, mconintCol送货单位) = "" Then
                    .ColData(Col) = 5
                Else
                    If mint编辑状态 = 1 Or mint编辑状态 = 2 Or mint编辑状态 = 6 Then
                        .ColData(Col) = 4
                    Else
                        .ColData(Col) = 5
                    End If
                End If
        End Select
        
        If Row > 0 Then
            If .TextMatrix(Row, 0) <> "" Then
                Call ShowHisPlane(Row, Val(.TextMatrix(Row, 0)))
            End If
        End If
    End With
    
    Call 显示库存
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim lngRow As Long
    
    Dim rsTemp As Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim i As Integer
    Dim strTemp As String
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    intOldRow = mshBill.Row
    
    With mshBill
        If .Col = mconIntCol药名 Then
            .Text = UCase(Trim(.Text))
        Else
            .Text = Trim(.Text)
        End If
        strkey = .Text

        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col

            Case mconIntCol药名
                If strkey <> "" Then

                    sngLeft = Me.Left + Pic单据.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic单据.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If

'                    Set rsTemp = Frm药品多选选择器.ShowME(Me, 1, , mlng库房ID, , strkey, sngLeft, sngTop)
                    If grsMaster.State = adStateClosed Then
                       Call SetSelectorRS(1, MStrCaption, mlng库房ID, mlng库房ID)
                    End If
                    Set rsTemp = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, , mlng库房ID, , , , , , , , mstrPrivs)
                    
                    If rsTemp.RecordCount > 0 Then
                        '检查重复记录 并将重复记录的药品id返回回来
                        If rsTemp.RecordCount = 1 Then
                            lngRow = CheckDouData(rsTemp)
                            If lngRow > 0 Then
                                If MsgBox("该药品已经存在，是否跳转到记录行？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                                    mshBill.Row = lngRow
                                    mshBill.Col = 0
                                    mshBill.SetFocus
                                End If
                                Exit Sub
                            End If
                        Else
                            Set rsTemp = CheckData(rsTemp)
                        End If
                    End If
                    
                    If rsTemp.RecordCount > 0 Then
                        rsTemp.MoveFirst
                        For i = 1 To rsTemp.RecordCount
                            With mshBill
                                intCurRow = .Row
                                Select Case mintUnit
                                Case 1
                                    strUnit = "售价单位"
                                Case 2
                                    strUnit = "门诊单位"
                                Case 3
                                    strUnit = "住院单位"
                                Case Else
                                    strUnit = "药库单位"
                                End Select
                                Call SetDrugRows(rsTemp!药品id, _
                                        "[" & rsTemp!药品编码 & "]", _
                                        rsTemp!通用名, _
                                        IIf(IsNull(rsTemp!商品名), "", rsTemp!商品名), _
                                        IIf(IsNull(rsTemp!药品来源), "", rsTemp!药品来源), _
                                        IIf(IsNull(rsTemp!规格), "", rsTemp!规格), _
                                        IIf(IsNull(rsTemp!产地), "", rsTemp!产地), _
                                        Switch(strUnit = "售价单位", rsTemp!售价单位, strUnit = "门诊单位", rsTemp!门诊单位, _
                                               strUnit = "住院单位", rsTemp!住院单位, strUnit = "药库单位", rsTemp!药库单位), _
                                        rsTemp!指导批发价, _
                                        Switch(strUnit = "售价单位", 1, strUnit = "门诊单位", rsTemp!门诊包装, strUnit = "住院单位", _
                                               rsTemp!住院包装, strUnit = "药库单位", rsTemp!药库包装), rsTemp!原产地)
                                .Text = .TextMatrix(.Row, .Col)
                                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                    .rows = .rows + 1
                                End If
                                .Row = .rows - 1
                            End With
                            rsTemp.MoveNext
                        Next
                        mshBill.Row = intCurRow
                        mshBill.Col = mconintCol计划数量
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                        Cancel = True
                    End If
                End If
            Case mconintCol计划数量
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，计划数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintCol计划数量) = ""
                    End If
                    If .TextMatrix(.Row, mconintCol计划数量) <> "" Then
                        strkey = .TextMatrix(.Row, mconintCol计划数量)
                        If .TextMatrix(.Row, mconintCol成本价) <> "" Then
                            .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol成本价) * strkey, mintShowMoneyDigit, , True)
                        End If
                        If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                            .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strkey, mintShowMoneyDigit, , True)
                        End If
                    End If
                    .Col = mconintCol说明
                    Cancel = True
                End If
                
                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintShowNumberDigit, , True)
                    If Val(.TextMatrix(.Row, mconintCol计划数量)) <> Val(strkey) And Not mblnCheckRefresh Then
                        mblnCheckRefresh = True
                    End If
                    .Text = strkey
                    If .TextMatrix(.Row, mconintCol成本价) <> "" Then
                        .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol成本价) * strkey, mintShowMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol售价) <> "" Then
                        .TextMatrix(.Row, mconIntCol售价金额) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol售价) * strkey, mintShowMoneyDigit, , True)
                    End If
                    If Val(.TextMatrix(.Row, mconintCol送货包装)) <> 0 Then
                        .TextMatrix(.Row, mconintCol送货数量) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol比例系数)) / Val(.TextMatrix(.Row, mconintCol送货包装)) * Val(strkey), 1)
                    End If
                End If
                
                Call 显示合计金额
            Case mconintCol执行数量
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，计划数量必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mconintCol成本价
                If .TxtVisible = False Then Exit Sub
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "对不起，单价必须为数字型,请重输！", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconintCol成本价) = ""
                    End If
                    Cancel = True
                    Exit Sub
                End If

                If strkey <> "" Then
                    strkey = zlStr.FormatEx(strkey, mintShowPriceDigit, , True)
                    .Text = strkey
                    If .TextMatrix(.Row, mconintCol计划数量) <> "" Then
                        .TextMatrix(.Row, mconIntCol成本金额) = zlStr.FormatEx(.TextMatrix(.Row, mconintCol计划数量) * strkey, mintShowMoneyDigit, , True)
                    End If

                End If
                Call 显示合计金额
            Case mconIntCol生产商
                If strkey = "" And .TextMatrix(.Row, mconIntCol生产商) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntCol生产商) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, mlng生产商长度) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntCol原产地
                If strkey = "" And .TextMatrix(.Row, mconIntCol原产地) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntCol原产地) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, mlng原产地长度) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconintCol上次供应商
                If .TxtVisible = False Then Exit Sub
                If strkey = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol上次供应商) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strkey = UCase(strkey)
                    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft
                    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngLeft + Msf供应商选择.Width > Me.ScaleWidth Then sngLeft = Me.ScaleWidth - Msf供应商选择.Width - 100
                    
                    '1、非全院计划时要管站点
                    '全院计划未勾选参数时要管站点
                    If mlng库房ID <> 0 Or (mlng库房ID = 0 And mintPlanPoint = 0 And (gstrNodeNo <> "-" Or gstrNodeNo <> "0")) Then
                        strTemp = "(A.站点 = '" & gstrNodeNo & "' Or A.站点 is Null) And "
                    End If
                    
                    If mint供应商范围 = 1 Then
                        gstrSQL = "Select A.ID,A.编码,A.名称,A.简码 From 供应商 A,药品中标单位 B Where " & strTemp & " A.末级=1 And (substr(A.类型,1,1)=1 Or Nvl(A.末级,0)=0) ANd (To_Char(A.撤档时间,'yyyy-MM-dd')='3000-01-01' or A.撤档时间 is null) " & _
                            " And A.ID=B.单位ID And B.药品ID=[2] And (To_Char(B.撤档时间,'yyyy-MM-dd')='3000-01-01' or B.撤档时间 is null) " & _
                            " And (upper(A.编码) Like [1] Or Upper(A.名称) Like [1] Or Upper(A.简码) Like [1]) " & _
                            " Order By A.编码 "
                    Else
                        gstrSQL = "Select A.ID,A.编码,A.名称,A.简码 From 供应商 A Where  " & strTemp & " A.末级=1 And (substr(A.类型,1,1)=1 Or Nvl(A.末级,0)=0) ANd (To_Char(A.撤档时间,'yyyy-MM-dd')='3000-01-01' or A.撤档时间 is null) " & _
                            " And (upper(A.编码) Like [1] Or Upper(A.名称) Like [1] Or Upper(A.简码) Like [1]) " & _
                            " Order By A.编码 "
                    End If
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取药应商]", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", Val(mshBill.TextMatrix(mshBill.Row, 0)))
                    
                    If rsTemp.RecordCount = 0 Then
                        If mint供应商范围 = 1 Then
                            '如果没有设置中标单位，则提取所有供应商
                            gstrSQL = "Select A.ID,A.编码,A.名称,A.简码 From 供应商 A Where  " & strTemp & " A.末级=1 And (substr(A.类型,1,1)=1 Or Nvl(A.末级,0)=0) ANd (To_Char(A.撤档时间,'yyyy-MM-dd')='3000-01-01' or A.撤档时间 is null) " & _
                                " And (upper(A.编码) Like [1] Or Upper(A.名称) Like [1] Or Upper(A.简码) Like [1]) " & _
                                " Order By A.编码 "
                            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[读取药应商]", IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", Val(mshBill.TextMatrix(mshBill.Row, 0)))
                            
                            If rsTemp.RecordCount = 0 Then
                                MsgBox "没有找到符合条件的供应商！", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            ElseIf rsTemp.RecordCount = 1 Then
                                .Text = rsTemp!名称
                                Exit Sub
                            End If
                        Else
                            MsgBox "没有找到符合条件的供应商！", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If

                    ElseIf rsTemp.RecordCount = 1 Then
                        .Text = rsTemp!名称
                        Exit Sub
                    End If
                    
                    With Msf供应商选择
                        .Clear
                        Set .DataSource = rsTemp
                        .ColWidth(0) = 0
                        .ColWidth(1) = 800
                        .ColWidth(2) = 3000
                        .ColWidth(3) = 800
            
                        .Row = 1
                        .ColSel = .Cols - 1
                    End With
                    With Msf供应商选择
                        .Left = sngLeft
                        .Top = sngTop
                        .Visible = True
                        .ZOrder 0
                        .SetFocus
                    End With
                    Cancel = True
                End If
            Case mconintCol说明
                If strkey = "" And .TextMatrix(.Row, mconintCol说明) = "" Then
                    strkey = " "
                    If Trim(.TextMatrix(.Row, mconintCol说明)) <> Trim(strkey) And Not mblnCheckRefresh Then
                        mblnCheckRefresh = True
                    End If
                    .Text = strkey
                    .TextMatrix(.Row, mconintCol说明) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 50) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
            Case mconIntCol批准文号
                If strkey = "" And .TextMatrix(.Row, mconIntCol批准文号) = "" Then
                    strkey = " "
                    .Text = strkey
                    .TextMatrix(.Row, mconIntCol批准文号) = strkey
                Else
                    If zlCommFun.StrIsValid(strkey, 40) = False Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntCol药名 Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub picHscSend_Click()
    If Val(picHscSend.Tag) = "1" Then
        Call imgDown_Click
    Else
        Call imgUp_Click
    End If
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer

    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '先判有否数据

            If LenB(StrConv(txt摘要.Text, vbFromUnicode)) > 40 Then
                MsgBox "摘要超长,最多能输入20个汉字或40个字符!", vbInformation + vbOKOnly, gstrSysName
                txt摘要.SetFocus
                Exit Function
            End If

            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntCol药名)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintCol计划数量))) <> "" Then
                        If Not IsNumeric(.TextMatrix(intLop, mconintCol计划数量)) Then
                            MsgBox "第" & intLop & "行药品的计划数量不为数字型，请检查！", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mconintCol计划数量
                            Exit Function
                        End If

                    End If
                    If Val(.TextMatrix(intLop, mconintCol计划数量)) > 9999999999# Then
                        MsgBox "第" & intLop & "行药品的计划数量大于了数据库能够保存的" & vbCrLf & "最大范围9999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol计划数量
                        Exit Function
                    End If

                    If Val(.TextMatrix(intLop, mconIntCol成本金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol计划数量
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol售价金额)) > 9999999999999# Then
                        MsgBox "第" & intLop & "行药品的售价金额大于了数据库能够保存的" & vbCrLf & "最大范围9999999999999，请检查！", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconintCol计划数量
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With

    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim lng序号 As Long
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim 计划类型_IN As Integer
    Dim 期间_IN As String
    Dim 库房ID_IN As Long
    Dim 编制方法_IN As Integer
    Dim 编制人_IN As String
    Dim 编制日期_IN As String
    Dim 编制说明_IN As String

    Dim 药品ID_IN As Long
    Dim 计划数量_IN As Double
    Dim 单价_IN As Double
    Dim 金额_IN As Double
    Dim 前期数量_IN As Double
    Dim 上期数量_IN As Double
    Dim 库存数量_IN As Double
    Dim 上次供应商_IN As String
    Dim 上次生产商_IN As String
    Dim 说明_IN As String
    Dim intRow As Integer
    Dim 售价_IN As Double
    Dim 售价金额_IN As Double
    Dim 上期销量_IN As Double
    Dim 本期销量_IN As Double
    Dim 药房ID_IN As Double
    Dim i As Integer
    Dim arrSql As Variant
    Dim 送货数量_in As Double
    Dim 批准文号_IN As String
    
    SaveCard = False
    arrSql = Array()

    On Error GoTo errHandle
    With mshBill
        ID_IN = Sys.NextId("药品采购计划")
        NO_IN = Trim(txtNo)
        If NO_IN = "" Then NO_IN = Sys.GetNextNo(32, mlng库房ID)
         
        If IsNull(NO_IN) Then Exit Function
        Me.txtNo.Tag = NO_IN
        计划类型_IN = mint计划类型
        编制方法_IN = mint编制方法
        库房ID_IN = mlng库房ID
        编制人_IN = UserInfo.用户姓名
        编制日期_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        编制说明_IN = Trim(txt摘要.Text)
        期间_IN = Str期间
        药房ID_IN = Val(txt摘要.Tag)
        
        If mint编辑状态 = 2 Or (mint编辑状态 = 3 And mblnCheckRefresh) Then      '修改
            gstrSQL = "zl_药品计划管理_DELETE('" & mlng计划ID & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If

        gstrSQL = "zl_药品计划管理主表_INSERT("
        '计划ID
        gstrSQL = gstrSQL & ID_IN
        'NO
        gstrSQL = gstrSQL & ",'" & NO_IN & "'"
        '计划类型
        gstrSQL = gstrSQL & "," & 计划类型_IN
        '期间
        gstrSQL = gstrSQL & ",'" & 期间_IN & "'"
        '库房ID
        gstrSQL = gstrSQL & "," & IIf(库房ID_IN = 0, "Null", 库房ID_IN)
        '药房ID
        gstrSQL = gstrSQL & "," & IIf(药房ID_IN = 0, "Null", 药房ID_IN)
        '编制方法
        gstrSQL = gstrSQL & "," & 编制方法_IN
        '编制人
        gstrSQL = gstrSQL & ",'" & 编制人_IN & "'"
        '编制日期
        gstrSQL = gstrSQL & ",to_date('" & 编制日期_IN & "','yyyy-mm-dd HH24:MI:SS')"
        '编制说明
        gstrSQL = gstrSQL & ",'" & 编制说明_IN & "'"
        '来源库房ID
        gstrSQL = gstrSQL & ",'" & IIf(mstr来源库房 = "", IIf(mlng库房ID = 0, mstrAll来源库房, mlng库房ID), mstr来源库房) & "'"
        '来源药房ID
        gstrSQL = gstrSQL & ",'" & IIf(mstr来源药房 = "", mstrAll来源药房, mstr来源药房) & "'"
        gstrSQL = gstrSQL & ")"

        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL

        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng序号 = .TextMatrix(intRow, mconIntCol序号)
                药品ID_IN = .TextMatrix(intRow, 0)
                
                单价_IN = .TextMatrix(intRow, mconintCol成本价) / Val(.TextMatrix(intRow, mconIntCol比例系数))
                金额_IN = Val(.TextMatrix(intRow, mconIntCol成本金额))
                售价_IN = .TextMatrix(intRow, mconIntCol售价) / Val(.TextMatrix(intRow, mconIntCol比例系数))
                售价金额_IN = Val(.TextMatrix(intRow, mconIntCol售价金额))
            
                前期数量_IN = Val(.TextMatrix(intRow, mconIntCol前期数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                上期数量_IN = Val(.TextMatrix(intRow, mconIntCol上期数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                库存数量_IN = Val(.TextMatrix(intRow, mconintCol库存数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                上期销量_IN = Val(.TextMatrix(intRow, mconintCol上期销量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                本期销量_IN = Val(.TextMatrix(intRow, mconintCol本期销量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                计划数量_IN = Val(.TextMatrix(intRow, mconintCol计划数量)) * Val(.TextMatrix(intRow, mconIntCol比例系数))
                上次供应商_IN = .TextMatrix(intRow, mconintCol上次供应商)
                上次生产商_IN = .TextMatrix(intRow, mconIntCol生产商)
                说明_IN = .TextMatrix(intRow, mconintCol说明)
                送货数量_in = Val(.TextMatrix(intRow, mconintCol送货数量))
                批准文号_IN = .TextMatrix(intRow, mconIntCol批准文号)
                
                gstrSQL = "zl_药品计划管理次表_INSERT("
                '计划ID
                gstrSQL = gstrSQL & ID_IN
                '药品ID
                gstrSQL = gstrSQL & "," & 药品ID_IN
                '序号
                gstrSQL = gstrSQL & "," & lng序号
                '计划数量
                gstrSQL = gstrSQL & "," & 计划数量_IN
                '单价
                gstrSQL = gstrSQL & "," & 单价_IN
                '金额
                gstrSQL = gstrSQL & "," & 金额_IN
                '前期数量
                gstrSQL = gstrSQL & "," & 前期数量_IN
                '上期数量
                gstrSQL = gstrSQL & "," & 上期数量_IN
                '库存数量
                gstrSQL = gstrSQL & "," & 库存数量_IN
                '供应商
                gstrSQL = gstrSQL & ",'" & 上次供应商_IN & "'"
                '生产商
                gstrSQL = gstrSQL & ",'" & 上次生产商_IN & "'"
                '说明
                gstrSQL = gstrSQL & ",'" & 说明_IN & "'"
                '售价
                gstrSQL = gstrSQL & "," & 售价_IN
                '售价金额
                gstrSQL = gstrSQL & "," & 售价金额_IN
                '上期销量
                gstrSQL = gstrSQL & "," & 上期销量_IN
                '本期销量
                gstrSQL = gstrSQL & "," & 本期销量_IN
                '送货数量
                gstrSQL = gstrSQL & "," & 送货数量_in
                '批准文号
                gstrSQL = gstrSQL & ",'" & 批准文号_IN & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    
    If mint编辑状态 = 3 And mblnCheckRefresh Then
        mlng计划ID = ID_IN
    End If
        
    SaveCard = True
    vsfStock.rows = 1
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "存盘失败！请检查！", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub 显示合计金额()
    Dim Dbl金额 As Double, dbl售价金额 As Double
    Dim intLop As Integer

    Dbl金额 = 0: dbl售价金额 = 0

    With mshBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                Dbl金额 = Dbl金额 + Val(.TextMatrix(intLop, mconIntCol成本金额))
                dbl售价金额 = dbl售价金额 + Val(.TextMatrix(intLop, mconIntCol售价金额))
            End If
        Next
    End With
    If mint价格显示 = 0 Then
        lblPurchasePrice.Caption = "金额合计：" & zlStr.FormatEx(Dbl金额, mintShowMoneyDigit, , True)
    ElseIf mint价格显示 = 1 Then
        lblPurchasePrice.Caption = "金额合计：" & zlStr.FormatEx(dbl售价金额, mintShowMoneyDigit, , True)
    Else
        lblPurchasePrice.Caption = "成本金额合计：" & zlStr.FormatEx(Dbl金额, mintShowMoneyDigit, , True) & "      售价金额合计：" & zlStr.FormatEx(dbl售价金额, mintShowMoneyDigit, , True)
    End If
End Sub


Private Sub 显示库存()
    Dim rsData As ADODB.Recordset
    Dim lng药品id As Long
    Dim str单位 As String
    Dim dbl包装 As Double
    Dim strSql As String
    
    If mblnStart = False Then Exit Sub
    
    On Error GoTo errHandle
    Me.staThis.Panels(2).Text = ""
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then vsfStock.rows = 1: Exit Sub
    
    lng药品id = Val(mshBill.TextMatrix(mshBill.Row, 0))
    str单位 = mshBill.TextMatrix(mshBill.Row, mconIntCol单位)
    dbl包装 = Val(mshBill.TextMatrix(mshBill.Row, mconIntCol比例系数))
    vsfStock.Tag = 0
    vsfStock.rows = 1
    
    If txtNo <> "" Then
        strSql = "Select t.来源库房, t.来源药房 From 药品采购计划 T Where t.No =[1]"
        Set rsData = zlDataBase.OpenSQLRecord(strSql, "", txtNo)
        
        mstr来源库房 = IIf(NVL(rsData!来源库房, 0) = 0, "", NVL(rsData!来源库房))
        mstr来源药房 = NVL(rsData!来源药房)
    End If
    
    If chk来源药房.Value = 1 Or chk来源库房.Value = 1 Or chk所有库房.Value = 1 Or mstr自定义库房 <> "" Then
        
        gstrSQL = "Select B.名称, A.药品id, Nvl(Sum(A.可用数量),0) As 可用数量, Nvl(Sum(A.实际数量),0) As 实际数量 " & _
            " From 药品库存 A, 部门表 B " & _
            " Where A.性质 = 1 And A.库房id + 0 = B.ID And A.药品id = [1] "
            
        If chk所有库房.Value = 0 Then
            If chk来源库房.Value = 1 And chk来源药房.Value = 1 And mstr自定义库房 <> "" And (mstr来源库房 <> "" Or mstrAll来源库房 <> "") _
                                                                                                                                     And (mstr来源药房 <> "" Or mstrAll来源药房 <> "") Then
                gstrSQL = gstrSQL & " and ( A.库房id In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                 " or A.库房id In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) " & _
                                                 " or A.库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                
            ElseIf chk来源库房.Value = 1 And chk来源药房.Value = 1 And mstr自定义库房 = "" And (mstr来源库房 <> "" Or mstrAll来源库房 <> "") _
                                                                                                                                        And (mstr来源药房 <> "" Or mstrAll来源药房 <> "") Then
                gstrSQL = gstrSQL & " and ( A.库房id In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                                 " or A.库房id In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) )"
            
            ElseIf chk来源库房.Value = 1 And chk来源药房.Value = 0 And mstr自定义库房 <> "" And (mstr来源库房 <> "" Or mstrAll来源库房 <> "") Then
                gstrSQL = gstrSQL & " and ( A.库房id In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) " & _
                                                                 " or A.库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                                                                          
            ElseIf chk来源库房.Value = 0 And chk来源药房.Value = 1 And mstr自定义库房 <> "" And (mstr来源药房 <> "" Or mstrAll来源药房 <> "") Then
                gstrSQL = gstrSQL & " and ( A.库房id In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) " & _
                                                                 " or A.库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) )"
                                                                                                          
            ElseIf chk来源库房.Value = 1 And chk来源药房.Value = 0 And mstr自定义库房 = "" And (mstr来源库房 <> "" Or mstrAll来源库房 <> "") Then
                gstrSQL = gstrSQL & " and A.库房id In(select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList))) "
                
            ElseIf chk来源库房.Value = 0 And chk来源药房.Value = 1 And mstr自定义库房 = "" And (mstr来源药房 <> "" Or mstrAll来源药房 <> "") Then
                gstrSQL = gstrSQL & " and A.库房id In(select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) "
                
            ElseIf chk来源库房.Value = 0 And chk来源药房.Value = 0 And mstr自定义库房 <> "" Then
                gstrSQL = gstrSQL & " and A.库房id In(select * from Table(Cast(f_Num2List([4]) As zlTools.t_NumList))) "
                
            End If
        End If
        
        gstrSQL = gstrSQL & " Group By B.名称, A.药品id " & _
        " Order By B.名称"
        
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "显示库存", lng药品id, _
                                                                        IIf(mstr来源库房 = "", IIf(mlng库房ID = 0, mstrAll来源库房, mlng库房ID), mstr来源库房), _
                                                                        IIf(mstr来源药房 = "", mstrAll来源药房, mstr来源药房), mstr自定义库房)
        
        Do While Not rsData.EOF
            vsfStock.rows = vsfStock.rows + 1
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("库房")) = rsData!名称
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("可用数量")) = zlStr.FormatEx(rsData!可用数量 / dbl包装, mintShowNumberDigit, , True)
            vsfStock.TextMatrix(vsfStock.rows - 1, vsfStock.ColIndex("实际数量")) = zlStr.FormatEx(rsData!实际数量 / dbl包装, mintShowNumberDigit, , True)
            
            rsData.MoveNext
        Loop
    End If
            
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt摘要_Change()
    mblnChange = True
End Sub

Private Sub txt摘要_GotFocus()
    OS.OpenIme True
    With txt摘要
        .SelStart = 0
        .SelLength = Len(txt摘要.Text)
    End With
End Sub

Private Sub txt摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txt摘要_LostFocus()
    OS.OpenIme
End Sub

Private Function SetDrugRows(ByVal lng药品id As Long, ByVal str药品编码 As String, ByVal str通用名 As String, ByVal str商品名 As String, ByVal str药品来源 As String, _
        ByVal str规格 As String, ByVal str产地 As String, ByVal str单位 As String, _
        ByVal dbl指导批发价 As Double, ByVal dbl比例系数 As Double, ByVal str原产地 As String) As Boolean
    Dim rsData As New Recordset
    Dim intCount As Integer
    Dim intRow As Integer
    Dim intCol As Integer

    Dim lng批次 As Long
    Dim dbl库存数量 As Double
    Dim dbl成本单价 As Double, dbl销售单价 As Double
    Dim rs合同单位 As ADODB.Recordset
    Dim str药名 As String
    Dim rsTemp As ADODB.Recordset
    Dim dbl送货包装 As Double
    Dim str送货单位 As String
    Dim str供应商 As String
    Dim str批准文号 As String
    
    On Error GoTo errH
    SetDrugRows = False

    With mshBill
        .TextMatrix(.Row, mconIntCol序号) = .Row
        .TextMatrix(.Row, mconIntCol生产商) = str产地
        .TextMatrix(.Row, mconIntCol原产地) = str原产地
        .TextMatrix(.Row, 0) = lng药品id
        .TextMatrix(.Row, mconIntCol比例系数) = dbl比例系数
        
        gstrSQL = "Select a.成本价,a.指导批发价,b.名称 as 供应商,nvl(a.上次批准文号,a.批准文号) as 批准文号 From 药品规格 a ,供应商 b Where a.上次供应商id=b.id(+) and a.药品ID=[1]"
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取平均成本价]", lng药品id)
        dbl成本单价 = NVL(rsData!成本价, 0)
        If dbl成本单价 = 0 Then dbl成本单价 = NVL(rsData!指导批发价, 0)
        str供应商 = NVL(rsData!供应商, "")
        str批准文号 = NVL(rsData!批准文号, "")
        
        gstrSQL = "Select Decode(Nvl(a.是否变价, 0), 0, Nvl(b.现价, 0), Decode(nvl(d.上次售价,0), 0, Decode(Nvl(c.平均售价, 0), 0, b.现价, c.平均售价), d.上次售价)) 售价 " & _
                 " From 收费项目目录 A, 收费价目 B, 药品规格 D, " & _
                 " (Select 药品id, " & _
                 " Decode(Sign(Sum(实际数量)), 1, Decode(Sign(Sum(实际金额)), 1, Sum(实际金额), 0) / Sum(实际数量), 0) 平均售价 " & _
                 " From 药品库存 " & _
                 " Where 性质 = 1 " & IIf(mlng库房ID = 0, "", " AND 库房ID=[2] ") & " And 药品id = [1] " & _
                 " Group By 药品id) C " & _
                 " Where A.ID = B.收费细目id And A.ID = C.药品id(+) And A.ID = D.药品id And A.ID = [1] And " & _
                 " (A.撤档时间 >= To_Date('3000-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS') Or A.撤档时间 Is Null) And " & _
                 " (B.终止日期 Is Null Or Sysdate Between B.执行日期 And Nvl(B.终止日期, To_Date('3000-01-01', 'yyyy-MM-dd'))) " & _
                 GetPriceClassString("B")
                 
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取售价]", lng药品id, mlng库房ID)
        dbl销售单价 = rsData!售价
        
        If mlng库房ID = 0 Then
            '如果是全库房，则从药品库存中取库存数量，从药品规格中取供应商，上次产地
            gstrSQL = "Select B.名称 供应商, C.上次产地, C.原产地,Nvl(A.库存数量,0) 库存数量" & _
                      " From (Select 药品id, Sum(实际数量) As 库存数量 From 药品库存 " & _
                      " Where 性质 = 1 And 药品id = [1] Group By 药品id) A, " & _
                      " (Select id,名称 From 供应商 Where Substr(类型, 1, 1) = 1) B, 药品规格 C " & _
                      " Where C.药品id = A.药品id(+) And C.上次供应商id = B.ID(+) And C.药品id = [1] "
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次供应商及产地信息]", lng药品id, mlng库房ID)
        Else
            '如果指定库房，则从指定库房取库存数量，及最大批次的供应商，上次产地
            gstrSQL = " Select C.名称 As 供应商, B.上次产地, B.原产地,A.库存数量 " & _
                      " From (Select 库房id, 药品id, Sum(实际数量) As 库存数量 " & _
                      " From 药品库存 " & _
                      " Where 性质 = 1 And 药品id = [1] And 库房ID=[2] " & _
                      " Group By 库房id, 药品id) A, " & _
                      " (Select 库房id,药品id,上次供应商ID,上次产地,原产地 From 药品库存 " & _
                      " Where 性质 = 1 And 药品id = [1] And 库房ID=[2] " & _
                      " And Nvl(批次, 0) = " & _
                      " (Select Nvl(Max(Nvl(批次, 0)), 0) 批次 From 药品库存 Where 性质 = 1 And 药品id = [1] And 库房ID=[2] )) B, " & _
                      " (SELECT id,名称 FROM 供应商 WHERE SUBSTR(类型,1,1)=1) C " & _
                      " Where A.库房id = B.库房id And A.药品id = B.药品id And B.上次供应商id = C.ID(+) " & _
                      " And A.药品id = [1] And A.库房ID=[2] "
            Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次供应商及产地信息]", lng药品id, mlng库房ID)
            
            '如果指定库房无库存，则从药品规格中取供应商，上次产地
            If rsData.RecordCount = 0 Then
                gstrSQL = "Select B.名称 供应商, C.上次产地, C.原产地, 0 库存数量 from " & _
                          " (Select ID,名称 From 供应商 Where Substr(类型, 1, 1) = 1) B, 药品规格 C " & _
                          " Where C.上次供应商id = B.ID(+) And 药品id = [1] "
                Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[取上次供应商及产地信息]", lng药品id, mlng库房ID)
            End If
        End If
       
        If Not rsData.EOF Then
            .TextMatrix(.Row, mconintCol库存数量) = zlStr.FormatEx(IIf(IsNull(rsData!库存数量), 0, rsData!库存数量) / dbl比例系数, mintShowNumberDigit, , True)
            
            .TextMatrix(.Row, mconintCol上次供应商) = IIf(IsNull(rsData!供应商), str供应商, rsData!供应商)
            If mint供应商选择 = 1 Then
                gstrSQL = "Select B.名称 From 药品规格 A, 供应商 B Where A.合同单位id = B.ID And A.药品id = [1] "
                Set rs合同单位 = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lng药品id)
                If Not rs合同单位.EOF Then
                    .TextMatrix(.Row, mconintCol上次供应商) = rs合同单位!名称
                End If
            End If
            
            .TextMatrix(.Row, mconIntCol生产商) = IIf(IsNull(rsData!上次产地), str产地, rsData!上次产地)
            .TextMatrix(.Row, mconIntCol原产地) = IIf(IsNull(rsData!原产地), str产地, rsData!原产地)
            SetNumer lng药品id, mlng库房ID, .TextMatrix(.Row, mconintCol库存数量), .Row, mint计划类型, mint编制方法, mbln数量方式
        End If
        
        '加载大包装入库信息
        gstrSQL = "select a.送货单位,a.送货包装,a.基本药物,b.费用类型 from 药品规格 a,收费项目目录 b where a.药品id=b.id and a.药品id=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "入库信息", lng药品id)
        dbl送货包装 = IIf(IsNull(rsTemp!送货包装), 0, rsTemp!送货包装)
        str送货单位 = IIf(IsNull(rsTemp!送货单位), "", rsTemp!送货单位)
        .TextMatrix(.Row, mconIntcol医保类型) = IIf(IsNull(rsTemp!费用类型), "", rsTemp!费用类型)
        .TextMatrix(.Row, mconIntCol基本药物) = IIf(IsNull(rsTemp!基本药物), "", rsTemp!基本药物)
        If dbl送货包装 <> 0 Then
            .TextMatrix(.Row, mconintCol送货单位) = str送货单位 & "(1" & str送货单位 & "=" & zlStr.FormatEx(dbl送货包装 / dbl比例系数, 1, , True) & str单位 & ")"
            .TextMatrix(.Row, mconintCol送货包装) = dbl送货包装
            If Val(.TextMatrix(.Row, mconintCol计划数量)) <> 0 Then
                .TextMatrix(.Row, mconintCol送货数量) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconintCol计划数量)) / dbl送货包装, 1, , True)
            End If
        End If
        
        If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
            str药名 = str通用名
        Else
            str药名 = IIf(str商品名 <> "", str商品名, str通用名)
        End If
        
        .TextMatrix(.Row, mconIntCol药品编码和名称) = str药品编码 & str药名
        .TextMatrix(.Row, mconIntCol药品编码) = str药品编码
        .TextMatrix(.Row, mconIntCol药品名称) = str药名
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(.Row, mconIntCol药名) = .TextMatrix(.Row, mconIntCol药品编码)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(.Row, mconIntCol药名) = .TextMatrix(.Row, mconIntCol药品名称)
        Else
            .TextMatrix(.Row, mconIntCol药名) = .TextMatrix(.Row, mconIntCol药品编码和名称)
        End If
        
        .TextMatrix(.Row, mconIntCol商品名) = str商品名
        
        .TextMatrix(.Row, mconIntCol来源) = str药品来源
        .TextMatrix(.Row, mconIntCol规格) = str规格
        .TextMatrix(.Row, mconIntCol单位) = str单位
        .TextMatrix(.Row, mconintCol成本价) = zlStr.FormatEx(dbl成本单价 * dbl比例系数, mintShowPriceDigit, , True)
        .TextMatrix(.Row, mconIntCol售价) = zlStr.FormatEx(dbl销售单价 * dbl比例系数, mintShowPriceDigit, , True)
        .TextMatrix(.Row, mconIntCol批准文号) = str批准文号
    End With
    rsData.Close
    SetDrugRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'取指导批发价定价单位的设置值，缺省为0-按售价单位定价，可选为1-按药库单位定价；
Private Function GetUnit() As Integer
    GetUnit = gtype_UserSysParms.P29_指导批发价定价单位
End Function

Private Sub vsfHisPlane_EnterCell()
    Dim lngColor As Long
    
    With vsfHisPlane
        .ForeColorSel = &H80000008
        
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("NO")) = "" Then Exit Sub
        
        lngColor = .Cell(flexcpForeColor, .Row, .ColIndex("计划数量"))
        
        .ForeColorSel = lngColor
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '功能：用来检查列表中已有药品与新选择的药品是否重复和时价药品是否有库存

    Dim i As Integer
    Dim strTemp As String
    Dim strInfo As String
    Dim strSql As String
    Dim strDub As String    '重复药品
    Dim str重复药名 As String   '用来记录重复选择了的药品名称
    
    rsTemp.MoveFirst
    strTemp = ""
    Do While Not rsTemp.EOF
        If InStr(1, strTemp, rsTemp!药品id) = 0 Then
            strTemp = strTemp & rsTemp!药品id & "," & rsTemp!通用名 & "|"
        End If
        rsTemp.MoveNext
    Loop
        
    With mshBill    '把重复的查询出来
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & ",") > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol药名) & "|"
            End If
        Next
        
        If strInfo <> "" Then   '为过滤数据拼接sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "药品id<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str重复药名, ",")) <= 2 Then
                    str重复药名 = str重复药名 & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        '判断以什么方式拼接sql
        If str重复药名 <> "" Then
            MsgBox str重复药名 & "列表中已经含有了！" & vbCrLf & "以上药品不再添加！", vbInformation, gstrSysName
            strSql = strDub
        End If
        If strSql <> "" Then
            rsTemp.Filter = strSql
        End If
        
        Set CheckData = rsTemp
    End With
End Function

Private Sub SetColumnByUserDefine()
    Dim intColumns As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected, arr总列, arr可设置列
    Dim intCol As Integer, intCols As Integer
    Dim strAllCol As String
    
    On Error GoTo ErrHand
    
    strColumn_Selected = zlDataBase.GetPara("选择列", glngSys, 模块号.药品计划)
    mstrColumn_UnSelected = zlDataBase.GetPara("屏蔽列", glngSys, 模块号.药品计划)
    strColumn_All = "药名,2|商品名,3|药品来源,4|规格,5|生产商,6|原产地,7|单位,8|医保类型,10|前期数量,11|上期数量,12|库存上限,13|库存下限,14|" & _
                        "库存数量,15|上期销量,16|本期销量,17|计划数量,18|执行数量,19|送货单位,21|送货数量,22|成本价,24|成本金额,25|售价,26|售价金额,27|上次供应商,28|说明,29|基本药物,33|批准文号,34"
    
    If strColumn_Selected <> "" Then
        If mstrColumn_UnSelected <> "" Then
            strAllCol = strColumn_Selected & "|" & mstrColumn_UnSelected
        Else
            strAllCol = strColumn_Selected
        End If
        arr总列 = Split(strColumn_All, "|")
        arr可设置列 = Split(strAllCol, "|")
        If UBound(arr总列) <> UBound(arr可设置列) Then
            strColumn_Selected = "药名|商品名|药品来源|规格|生产商|原产地|单位|医保类型|前期数量|上期数量|库存上限|库存下限|库存数量|上期销量|本期销量|计划数量|执行数量|送货单位|送货数量|成本价|成本金额|售价|售价金额|上次供应商|说明|基本药物|批准文号"
            mstrColumn_UnSelected = ""
            zlDataBase.SetPara "选择列", strColumn_Selected, glngSys, 模块号.药品计划
            zlDataBase.SetPara "屏蔽列", mstrColumn_UnSelected, glngSys, 模块号.药品计划
        End If
    Else
        strColumn_Selected = "药名|商品名|药品来源|规格|生产商|原产地|单位|医保类型|前期数量|上期数量|库存上限|库存下限|库存数量|上期销量|本期销量|计划数量|执行数量|送货单位|送货数量|成本价|成本金额|售价|售价金额|上次供应商|说明|基本药物|批准文号"
        mstrColumn_UnSelected = ""
        zlDataBase.SetPara "选择列", strColumn_Selected, glngSys, 模块号.药品计划
        zlDataBase.SetPara "屏蔽列", mstrColumn_UnSelected, glngSys, 模块号.药品计划
    End If
    
    '设置默认值
    mconIntCol序号 = 1
    mconIntCol药名 = 2
    mconIntCol商品名 = 3
    mconIntCol来源 = 4
    mconIntCol规格 = 5
    mconIntCol生产商 = 6
    mconIntCol原产地 = 7
    mconIntCol单位 = 8
    mconIntCol比例系数 = 9
    mconIntcol医保类型 = 10
    mconIntCol前期数量 = 11
    mconIntCol上期数量 = 12
    mconIntCol库存上限 = 13
    mconIntCol库存下限 = 14
    mconintCol库存数量 = 15
    mconintCol上期销量 = 16
    mconintCol本期销量 = 17
    mconintCol计划数量 = 18
    mconintCol执行数量 = 19
    mconintCol原执行数量 = 20
    mconintCol送货单位 = 21
    mconintCol送货数量 = 22
    mconintCol送货包装 = 23
    mconintCol成本价 = 24
    mconIntCol成本金额 = 25
    mconIntCol售价 = 26
    mconIntCol售价金额 = 27
    mconintCol上次供应商 = 28
    mconintCol说明 = 29
    mconIntCol药品编码和名称 = 30
    mconIntCol药品编码 = 31
    mconIntCol药品名称 = 32
    mconIntCol基本药物 = 33
    mconIntCol批准文号 = 34
    mconIntColS = 35      '总列数
    mshBill.Cols = 35
    '商品名列处理
    If gint药品名称显示 = 2 Then
        '显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = IIf(mshBill.ColWidth(mconIntCol商品名) = 0, 2000, mshBill.ColWidth(mconIntCol商品名))
    Else
        '不单独显示商品名列
        mshBill.ColWidth(mconIntCol商品名) = 0
    End If

    '根据用户设置调整列顺序
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(strColumn_Selected, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    '将未选择的列的列宽设置为零，且列数据为5——不可选择
    If mstrColumn_UnSelected = "" Then Exit Sub
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstrColumn_UnSelected, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    Exit Sub
ErrHand:
    MsgBox "恢复列设置时发生错误，请重新进行列设置！", vbInformation, gstrSysName
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str列名
        Case "序号"
            mconIntCol序号 = intValue
        Case "药名"
            mconIntCol药名 = intValue
        Case "商品名"
            mconIntCol商品名 = intValue
        Case "药品来源"
            mconIntCol来源 = intValue
        Case "规格"
            mconIntCol规格 = intValue
        Case "生产商"
            mconIntCol生产商 = intValue
        Case "原产地"
            mconIntCol原产地 = intValue
        Case "单位"
            mconIntCol单位 = intValue
        Case "医保类型"
            mconIntcol医保类型 = intValue
        Case "前期数量"
            mconIntCol前期数量 = intValue
        Case "上期数量"
            mconIntCol上期数量 = intValue
        Case "库存上限"
            mconIntCol库存上限 = intValue
        Case "库存下限"
            mconIntCol库存下限 = intValue
        Case "库存数量"
            mconintCol库存数量 = intValue
        Case "上期销量"
            mconintCol上期销量 = intValue
        Case "本期销量"
            mconintCol本期销量 = intValue
        Case "计划数量"
            mconintCol计划数量 = intValue
        Case "执行数量"
            mconintCol执行数量 = intValue
        Case "送货单位"
            mconintCol送货单位 = intValue
        Case "送货数量"
            mconintCol送货数量 = intValue
        Case "成本价"
            mconintCol成本价 = intValue
        Case "成本金额"
            mconIntCol成本金额 = intValue
        Case "售价"
            mconIntCol售价 = intValue
        Case "售价金额"
            mconIntCol售价金额 = intValue
        Case "上次供应商"
            mconintCol上次供应商 = intValue
        Case "说明"
            mconintCol说明 = intValue
        Case "基本药物"
            mconIntCol基本药物 = intValue
        Case "批准文号"
            mconIntCol批准文号 = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Sub GetDefineSize()
    '功能：得到数据库的表字段的长度
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.上次产地 as 生产商, t.原产地 as 原产地 From 药品规格 T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng生产商长度 = rsTmp.Fields("生产商").DefinedSize
    mlng原产地长度 = rsTmp.Fields("原产地").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Init存储库房()
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where c.工作性质 = b.名称 " _
            & "  AND Instr('HIJKLMN',b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " _
            & IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[2])")
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.用户ID)
    
    If rsDepend.EOF Then
        MsgBox "没有设置药库性质的部门,请查看部门管理！", vbInformation, gstrSysName
        rsDepend.Close
        Exit Sub
    End If
    
    lvw存储库房.ListItems.Clear

    With rsDepend
        Do While Not .EOF
            lvw存储库房.ListItems.Add , "K" & !Id, !名称, , 2
            .MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Change存储库房()
    Dim i As Integer, j As Integer
    Dim intSelect As Integer
    Dim strArr存储库房() As String
    strArr存储库房 = Split(mstr自定义库房, ",")
    
    For i = LBound(strArr存储库房) To UBound(strArr存储库房)
        For intSelect = 1 To lvw存储库房.ListItems.count
            If strArr存储库房(i) = Mid(lvw存储库房.ListItems(intSelect).Key, 2) Then
                lvw存储库房.ListItems(intSelect).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvw存储库房.ListItems.count Then
        chk库房.Value = 1
    ElseIf j > 0 And j < lvw存储库房.ListItems.count Then
        chk库房.Value = 2
    End If
End Sub
Private Sub cmd库房_Click()
    With pic库房
        .Visible = Not .Visible
    End With
    
    Call ResizeHisPlane
    Call Change存储库房
End Sub

Private Sub chk库房_Click()
'库房全选按钮
    If chk库房.Value = 2 Then Exit Sub
    Call SetSelect(lvw存储库房, chk库房.Value)
End Sub
Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
'全选功能
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub lvw存储库房_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'具体选择的存储库房
    Call ItemCheck(lvw存储库房, Item, chk库房)
End Sub
Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
'纪录选择的库房
    Dim lngCheck As Long, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub

