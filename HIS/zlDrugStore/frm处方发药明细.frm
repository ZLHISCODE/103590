VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm处方发药明细 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picRecipt 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      ScaleHeight     =   6375
      ScaleWidth      =   11775
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.PictureBox picRecInfo 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   1515
         Left            =   0
         ScaleHeight     =   1515
         ScaleWidth      =   10755
         TabIndex        =   24
         Top             =   0
         Width           =   10755
         Begin VB.PictureBox picRecipeColor 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   460
            Left            =   120
            ScaleHeight     =   465
            ScaleWidth      =   1095
            TabIndex        =   27
            Top             =   50
            Width           =   1095
            Begin VB.Label lblRecipeType 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "普通"
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
               Left            =   240
               TabIndex        =   28
               Top             =   105
               Width           =   600
            End
         End
         Begin VB.ComboBox TxtNo 
            Height          =   300
            ItemData        =   "frm处方发药明细.frx":0000
            Left            =   8085
            List            =   "frm处方发药明细.frx":0002
            TabIndex        =   26
            Top             =   280
            Width           =   1965
         End
         Begin VB.TextBox txt诊断内容 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   2400
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   715
            Width           =   7695
         End
         Begin VB.Label LblTel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "15310625533"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6000
            TabIndex        =   52
            Tag             =   "年龄："
            Top             =   45
            Width           =   1155
         End
         Begin VB.Label LblTel 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "电话："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   5400
            TabIndex        =   51
            Tag             =   "年龄："
            Top             =   45
            Width           =   585
         End
         Begin VB.Label lblNotice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "禁忌药品说明："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   1440
            TabIndex        =   50
            Tag             =   "中药煎法:"
            Top             =   1140
            Width           =   1365
         End
         Begin VB.Label lbl诊断 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "临床诊断："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   1440
            TabIndex        =   49
            Tag             =   "临床诊断："
            Top             =   715
            Width           =   975
         End
         Begin VB.Label Lbl科室 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "科室："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   48
            Tag             =   "科室："
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "性别："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   3120
            TabIndex        =   47
            Tag             =   "性别："
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "年龄："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   4200
            TabIndex        =   46
            Tag             =   "年龄："
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl住院号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "标识号："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   5400
            TabIndex        =   45
            Tag             =   "标识号："
            Top             =   330
            Width           =   780
         End
         Begin VB.Label Lbl床号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "床号："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   3120
            TabIndex        =   44
            Tag             =   "床号："
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbl姓名 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "姓名："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   43
            Tag             =   "姓名："
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl收费员 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "收费员："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   9165
            TabIndex        =   42
            Tag             =   "收费员:"
            Top             =   45
            Width           =   780
         End
         Begin VB.Label LblNo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "单据号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Left            =   7320
            TabIndex        =   41
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbl药房 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "药房"
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
            Left            =   120
            TabIndex        =   40
            Top             =   675
            Width           =   390
         End
         Begin VB.Label lbl就诊卡号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "就诊卡："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   7320
            TabIndex        =   39
            Tag             =   "就诊卡："
            Top             =   45
            Width           =   780
         End
         Begin VB.Label Lbl科室 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "五官科"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   38
            Top             =   330
            Width           =   585
         End
         Begin VB.Label lbl姓名 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "张三疯"
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
            Index           =   1
            Left            =   1980
            TabIndex        =   37
            Top             =   45
            Width           =   585
         End
         Begin VB.Label Lbl性别 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "男"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3720
            TabIndex        =   36
            Tag             =   "性别："
            Top             =   45
            Width           =   195
         End
         Begin VB.Label Lbl床号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "33"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   3720
            TabIndex        =   35
            Tag             =   "床号："
            Top             =   330
            Width           =   210
         End
         Begin VB.Label Lbl年龄 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "22岁"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4800
            TabIndex        =   34
            Tag             =   "年龄："
            Top             =   45
            Width           =   405
         End
         Begin VB.Label Lbl住院号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "1234567"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   6120
            TabIndex        =   33
            Tag             =   "标识号："
            Top             =   330
            Width           =   735
         End
         Begin VB.Label lbl就诊卡号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "123456789"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   8040
            TabIndex        =   32
            Tag             =   "就诊卡："
            Top             =   45
            Width           =   945
         End
         Begin VB.Label Lbl收费员 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "孙大圣 "
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   9960
            TabIndex        =   31
            Tag             =   "收费员:"
            Top             =   45
            Width           =   690
         End
         Begin VB.Label LblWeight 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "55kg"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4800
            TabIndex        =   30
            Tag             =   "体重："
            Top             =   330
            Width           =   420
         End
         Begin VB.Label LblWeight 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "体重："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   180
            Index           =   0
            Left            =   4200
            TabIndex        =   29
            Tag             =   "床号："
            Top             =   330
            Width           =   585
         End
      End
      Begin VB.PictureBox picMark1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   10920
         ScaleHeight     =   855
         ScaleWidth      =   855
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   855
         Begin VB.PictureBox picMark2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   695
            Left            =   80
            ScaleHeight     =   690
            ScaleWidth      =   690
            TabIndex        =   19
            Top             =   80
            Width           =   695
            Begin VB.Label lblMark 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "记"
               BeginProperty Font 
                  Name            =   "微软雅黑"
                  Size            =   24
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   600
               Left            =   100
               TabIndex        =   20
               Top             =   0
               Width           =   615
            End
         End
      End
      Begin VB.PictureBox picHscSend 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   4575
         TabIndex        =   16
         Tag             =   "0"
         Top             =   3480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.Label lblDiag 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "抗菌药物相关信息"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   1200
            TabIndex        =   17
            Top             =   30
            Width           =   2400
         End
         Begin VB.Image imgDown 
            Height          =   240
            Left            =   0
            Picture         =   "frm处方发药明细.frx":0004
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgUp 
            Height          =   240
            Left            =   0
            Picture         =   "frm处方发药明细.frx":0346
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frm处方发药明细.frx":0688
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfNoList 
         Height          =   1005
         Left            =   7560
         TabIndex        =   12
         Top             =   2760
         Visible         =   0   'False
         Width           =   2820
         _cx             =   4974
         _cy             =   1773
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm处方发药明细.frx":0BD6
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
         ExplorerBar     =   1
         PicturesOver    =   -1  'True
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
         VirtualData     =   0   'False
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
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   1296
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
         FormatString    =   $"frm处方发药明细.frx":0D2E
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
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1800
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   3720
         _cx             =   6562
         _cy             =   3175
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
         BackColorSel    =   16769992
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm处方发药明细.frx":0D7C
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
      Begin VB.PictureBox picProcess 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   0
         ScaleHeight     =   735
         ScaleWidth      =   10935
         TabIndex        =   4
         Top             =   5400
         Width           =   10935
         Begin VB.CommandButton cmdSendByNoTake 
            Caption         =   "病人未取药发药(&T)"
            Height          =   350
            Left            =   7320
            TabIndex        =   23
            ToolTipText     =   "热键：F2"
            Top             =   360
            Width           =   1815
         End
         Begin VB.ComboBox cbo配药人 
            Enabled         =   0   'False
            Height          =   300
            Left            =   810
            TabIndex        =   22
            Text            =   "cbo配药人"
            Top             =   400
            Width           =   1815
         End
         Begin VB.ComboBox cbo开单医生 
            Enabled         =   0   'False
            Height          =   300
            Left            =   810
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   30
            Width           =   1815
         End
         Begin VB.ComboBox cbo核查人 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            TabIndex        =   13
            Text            =   "cbo核查人"
            Top             =   400
            Width           =   1695
         End
         Begin VB.CommandButton CmdSend 
            Caption         =   "发药(&S)"
            Height          =   350
            Left            =   9360
            TabIndex        =   6
            ToolTipText     =   "热键：F2"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox Chk全退 
            Appearance      =   0  'Flat
            Caption         =   "全退"
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5400
            TabIndex        =   5
            Top             =   445
            Value           =   1  'Checked
            Width           =   765
         End
         Begin VB.Label lbl核查人 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "核查人"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2880
            TabIndex        =   14
            Top             =   460
            Width           =   540
         End
         Begin VB.Label Lbl开单医生 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "开单医生"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   8
            Top             =   85
            Width           =   720
         End
         Begin VB.Label Lbl配药人 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "配药人"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   7
            Top             =   460
            Width           =   540
         End
      End
      Begin VB.TextBox txt用药理由 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3840
         Width           =   9975
      End
      Begin VB.PictureBox picRecInfo_CM 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   7455
         TabIndex        =   1
         Top             =   4800
         Width           =   7455
         Begin VB.Label lbl原始付数 
            AutoSize        =   -1  'True
            Caption         =   "原始付数："
            Height          =   180
            Left            =   0
            TabIndex        =   3
            Tag             =   "原始付数:"
            Top             =   60
            Width           =   900
         End
         Begin VB.Label lbl中药煎法 
            AutoSize        =   -1  'True
            Caption         =   "中药煎法："
            Height          =   180
            Left            =   1830
            TabIndex        =   2
            Tag             =   "中药煎法:"
            Top             =   60
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   1920
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":0DF1
            Key             =   "打印11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":118B
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":79ED
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":E24F
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":E7E9
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":EB83
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":EF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":F2B7
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":F651
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":10063
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":168C5
            Key             =   "未检"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":1D127
            Key             =   "在检"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":23989
            Key             =   "已检"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":2A1EB
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":30A4D
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":30DE7
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":31181
            Key             =   "套餐"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":379E3
            Key             =   "类型"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":3E245
            Key             =   "照片"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":44AA7
            Key             =   "参数"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":4B309
            Key             =   "指标"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":51B6B
            Key             =   "体检"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":583CD
            Key             =   "病历样式"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":5EC2F
            Key             =   "病历文件"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":65491
            Key             =   "规则"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":6BCF3
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":6C705
            Key             =   "诊断"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":72F67
            Key             =   "创建"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":797C9
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8002B
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8688D
            Key             =   "结束"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8D0EF
            Key             =   "部份"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8D489
            Key             =   "全部"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8D823
            Key             =   "部份总检"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8DBBD
            Key             =   "全部总检"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8DF57
            Key             =   "总检"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8E2F1
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8ED03
            Key             =   "已经打印"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":8F715
            Key             =   "药品"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":95F77
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   1200
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":96511
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":96AAB
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药明细.frx":97045
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   6240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm处方发药明细"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'工具栏菜单
Private Const conMenu_Tool_ShowPlug = 101           '调用插件：合理用药
Private mInt可操作 As Integer

Private mlngMode As Long

'条件
Private Type Type_ShowListCondition
    intListType As Integer                          '0-待配药;1-已配药;2-待发药;3-退药
    bln显示付数 As Boolean
    intShowPass As Integer
    bln显示大小单位 As Boolean
    bln显示过程单据 As Boolean
    bln显示重量 As Boolean
    bln校验处方 As Boolean                            '是否需要校验处方
    lng药房ID As Long
    bln医嘱作废 As Boolean
    str配药人 As String
    str核查人 As String
    bln是否需要配药过程 As Boolean
    bln自动配药 As Boolean
    int库存检查 As Integer
    bln过滤模式 As Boolean
    int金额显示 As Integer                          '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
    bln取药确认 As Boolean          '是否启用病人实际取药确认模式：0-不启用，1-启用
    bln处方审查 As Boolean
    bln允许核查人和配药人相同 As Boolean
    bln显示原产地 As Boolean
End Type
Private mcondition As Type_ShowListCondition

Private mbln中药处方 As Boolean
Private mstrDosUser As String
Private mstr核查人 As String

Private mblnAllowClick As Boolean                        '允许执行Click事件
Public mblnInput As Boolean                             '是否是通过录入方式来过滤处方
Private mbln未取药发药 As Boolean                       '未取药发药模式

Private mstrPrivs As String
Private mstrRecipeInfo As String                         '处方信息：单据;处方号;记录性质;门诊标志;药房ID

'列表类型
Private Enum mListType
    配药确认 = 0
    待配药 = 1
    已配药 = 2
    待发药 = 3
    超时未发 = 4
    退药 = 5
End Enum

'处方类型：普通、儿科、急诊、精二、精一、麻醉
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

'用户定义的处方颜色，从注册表取的字符串，用;分隔
Private mstrUserRecipeColor As String

Private mrsDetail As ADODB.Recordset

Private Const mlng紫色 As Long = &HC000C0
Private mblnResize As Boolean

'从参数表中取药品价格、数量、金额小数位数
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库

Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

Private mstrOracleMoneyForamt As String                 'ORACLE中金额格式
Private mstrVBMoneyForamt As String                     'VB中金额格式

Private mstrUnallowSetColHide As String         '不允许设置隐藏的列
Private mstrUnallowShow As String                   '不允许显示的列

Private mblnAllBack As Boolean

Private Const mconIntCol列数 = 68
Private mIntCol当前行 As Integer
Private mIntCol顺序号 As Integer
Private mIntCol审查结果 As Integer
Private mIntCol药品名称 As Integer
Private mintCol皮试结果 As Integer
Private mIntCol其它名 As Integer
Private mIntCol英文名 As Integer
Private mIntCol配方名称 As Integer
Private mintcol序号 As Integer
Private mintcol规格 As Integer
Private mintcol批号 As Integer
Private mintcol效期 As Integer
Private mIntColId As Integer
Private mintcol药品id As Integer
Private mintcol批次 As Integer
Private mintcol单位 As Integer
Private mIntCol单价 As Integer
Private mIntCol付数 As Integer
Private mintcol数量 As Integer
Private mIntCol金额 As Integer
Private mIntCol实收金额 As Integer
Private mIntCol重量 As Integer
Private mIntCol单量 As Integer
Private mIntCol用法 As Integer
Private mIntCol频次 As Integer
Private mIntCol用药目的 As Integer
Private mIntCol超量说明 As Integer
Private mIntCol医生嘱托 As Integer
Private mIntCol费别 As Integer
Private mIntCol库存数 As Integer
Private mIntCol货位 As Integer
Private mIntCol已退数 As Integer
Private mIntCol准退数 As Integer
Private mIntCol准退数大 As Integer
Private mIntCol准退数小 As Integer
Private mIntCol退药数 As Integer
Private mIntCol退药数大 As Integer
Private mIntCol单位大 As Integer
Private mIntCol退药数小 As Integer
Private mIntCol单位小 As Integer
Private mIntCol分批 As Integer
Private mIntCol新批号 As Integer
Private mIntCol新效期 As Integer
Private mIntCol新产地 As Integer
Private mIntCol备注 As Integer
Private mIntCol医嘱id As Integer
Private mIntCol实际数量 As Integer
Private mIntCol包装 As Integer
Private mIntCol单据 As Integer
Private mIntColNO As Integer
Private mIntCol门诊标志 As Integer
Private mIntCol记录性质 As Integer
Private mIntCol库房ID As Integer
Private mIntCol用药理由 As Integer
Private mIntCol相关id As Integer
Private mIntCol开嘱医生 As Integer
Private mIntCol频率间隔 As Integer
Private mIntCol间隔单位 As Integer
Private mIntCol医嘱标志 As Integer
Private mIntCol开始时间 As Integer
Private mIntCol结束时间 As Integer
Private mIntCol频率次数 As Integer
Private mIntCol警告 As Integer
Private mIntCol门诊号 As Integer
Private mIntCol住院号 As Integer
Private mIntCol禁忌药品说明 As Integer
Private mintcol生产商 As Integer
Private mintcol原产地 As Integer
Public Sub CmdProcess()
    If CmdSend.Enabled Then CmdSend_Click
End Sub

Public Sub FormClear()
    Me.lbl姓名(1).Caption = ""
    Me.Lbl床号(1).Caption = ""
    Me.Lbl床号(1).Caption = ""
    Me.Lbl科室(1).Caption = ""
    Me.Lbl年龄(1).Caption = ""
    Me.lbl就诊卡号(1).Caption = ""
    Me.Lbl收费员(1).Caption = ""
    Me.Lbl性别(1).Caption = ""
    Me.Lbl住院号(1).Caption = ""
    Me.LblWeight(1).Caption = ""
    LblTel(1).Caption = ""
    
    Me.txt诊断内容.Text = ""
    Me.txt诊断内容.Tag = ""
    
    txtNo.Clear
    txtNo.Tag = ""
    Lbl药房.Caption = ""
    
    vsfList.rows = 1
    vsfList.rows = 2
    
    Me.lbl原始付数.Caption = Me.lbl原始付数.Tag
    Me.lbl中药煎法.Caption = Me.lbl中药煎法.Tag
    lblNotice.Caption = "禁忌药品说明："
    
    CmdSend.Enabled = False
End Sub

Public Function GetDetailList() As VSFlexGrid
    Set GetDetailList = vsfList
End Function

Private Sub GetRecipeByNO()
    If mblnAllowClick = False Then Exit Sub
    If txtNo.ListIndex = -1 Then Exit Sub
    
    If mcondition.intListType <> mListType.退药 Then
        If zlStr.IsHavePrivs(mstrPrivs, "发其它药房的处方") = True Then
            If mcondition.lng药房ID <> Val(Split(txtNo.Tag, "|")(0)) Then
                If Val(Split(txtNo.Tag, "|")(5)) <> 1 Then
                    MsgBox "[" & Mid(txtNo.Text, 1, 8) & "]已经进行过发药操作,不能进行代发操作，请到" & Split(txtNo.Tag, "|")(1) & "取药！", vbInformation + vbOKOnly, gstrSysName
                    DoEvents
                    txtNo.Clear
                    txtNo.Text = ""
                    txtNo.Tag = ""
                    Lbl药房.Caption = ""
                    txtNo.SetFocus
                    Exit Sub
                End If
                If CDate(Format(Split(txtNo.Tag, "|")(4), "yyyy-MM-dd")) < CDate(Format(Sys.Currentdate, "yyyy-MM-dd")) - 30 Then
                    MsgBox "[" & Mid(txtNo.Text, 1, 8) & "]不是" & Split(txtNo.Tag, "|")(1) & "30天以内的的处方,不能进行代发操作，请到" & Split(txtNo.Tag, "|")(1) & "取药！", vbInformation + vbOKOnly, gstrSysName
                    Exit Sub
                End If
                
                If MsgBox("[" & Mid(txtNo.Text, 1, 8) & "]是" & Split(txtNo.Tag, "|")(1) & "的处方，是否继续！", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    DoEvents
                    txtNo.Clear
                    txtNo.Text = ""
                    txtNo.Tag = ""
                    Lbl药房.Caption = ""
                    txtNo.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    
    If (frm药品处方发药New.Dtp开始时间 > CDate(Split(txtNo.Tag, "|")(4)) Or frm药品处方发药New.cbo时间范围.ListIndex <> 3 Or (Split(txtNo.Tag, "|")(5) <> 1 And frm药品处方发药New.Chk显示退药待发单据.Value <> 1)) And Not mcondition.bln过滤模式 And mcondition.lng药房ID = Val(Split(txtNo.Tag, "|")(0)) Then
        If zlStr.IsHavePrivs(mstrPrivs, "允许查询所有时间范围单据") And Not (Not zlStr.IsHavePrivs(mstrPrivs, "修改过滤日期") And mcondition.intListType = mListType.退药) Then
            frm药品处方发药New.cbo时间范围.ListIndex = 3
            If frm药品处方发药New.Dtp开始时间 > CDate(Split(txtNo.Tag, "|")(4)) Then frm药品处方发药New.Dtp开始时间 = CDate(Split(txtNo.Tag, "|")(4))
            If frm药品处方发药New.Chk显示退药待发单据.Value <> 1 And Split(txtNo.Tag, "|")(5) <> 1 And mcondition.intListType <> mListType.退药 Then frm药品处方发药New.Chk显示退药待发单据.Value = 1
            frm药品处方发药New.GetCondition
            If mcondition.intListType = mListType.退药 Then
                frm药品处方发药New.RefreshList_Return Mid(txtNo.Text, 1, 8), True
            ElseIf mcondition.intListType = mListType.待发药 Then
                frm药品处方发药New.RefreshList_Send Mid(txtNo.Text, 1, 8), True
            ElseIf mcondition.intListType = mListType.超时未发 Then
                frm药品处方发药New.RefreshList_OverTime Mid(txtNo.Text, 1, 8), True
            Else
                frm药品处方发药New.RefreshList_Dosage Mid(txtNo.Text, 1, 8), True
            End If
        End If
    End If
    
    frm药品处方发药New.FindListRow 1, Mid(txtNo.Text, 1, 8), Mid(txtNo.Text, 11)

    DoEvents

    If mcondition.intListType = mListType.退药 Then
        frm药品处方发药New.RefreshDetail_Return Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), "", 1, Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3)), True
    Else
        frm药品处方发药New.RefreshDetail_Send Val(Split(txtNo.Tag, "|")(0)), Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3))
    End If

    If CmdSend.Enabled = True Then
        CmdSend.SetFocus
        mblnInput = True
    End If
End Sub

Public Function Get开单医生() As String
    If cbo开单医生.ListIndex = -1 Then
        Get开单医生 = ""
    ElseIf InStr(cbo开单医生.Text, "-") > 0 Then
        Get开单医生 = Mid(cbo开单医生.Text, InStr(cbo开单医生.Text, "-") + 1)
    Else
        Get开单医生 = cbo开单医生.Text
    End If
End Function

Public Function Get配药人() As String
    If Cbo配药人.ListIndex = -1 Then
        Get配药人 = ""
    ElseIf InStr(Cbo配药人.Text, "-") > 0 Then
        Get配药人 = Mid(Cbo配药人.Text, InStr(Cbo配药人.Text, "-") + 1)
    Else
        Get配药人 = Cbo配药人.Text
    End If
End Function
Public Function Get核查人() As String
    If cbo核查人.ListIndex = -1 Then
        Get核查人 = ""
    ElseIf InStr(cbo核查人.Text, "-") > 0 Then
        Get核查人 = Mid(cbo核查人.Text, InStr(cbo核查人.Text, "-") + 1)
    Else
        Get核查人 = cbo核查人.Text
    End If
End Function


Private Sub Load医生()
    Dim rsData As ADODB.Recordset
    
    Set rsData = RecipeSendWork_Get医生
    
    Me.cbo开单医生.Clear
    cbo开单医生.AddItem ""
    Do While Not rsData.EOF
        cbo开单医生.AddItem rsData!医生
        rsData.MoveNext
    Loop
    cbo开单医生.ListIndex = 0
End Sub

Private Sub SetCmdSendPrivs(ByVal int审查结果 As Integer)
    '权限控制
    Select Case mcondition.intListType
    Case mListType.配药确认
       '配药确认
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "配药确认")
    Case mListType.待配药
        '配药
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "配药") And mcondition.bln自动配药 = False And (mcondition.bln处方审查 = False Or (mcondition.bln处方审查 = True And int审查结果 = 1)))
    Case mListType.已配药
        '取消
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "配药")
    Case mListType.待发药, mListType.超时未发
        '发药
        CmdSend.Enabled = (zlStr.IsHavePrivs(mstrPrivs, "发药") And (mcondition.bln处方审查 = False Or (mcondition.bln处方审查 = True And int审查结果 = 1)))
    Case mListType.退药
        CmdSend.Enabled = zlStr.IsHavePrivs(mstrPrivs, "退药")
        If mcondition.intListType = mListType.退药 And mInt可操作 = 1 And zlStr.IsHavePrivs(mstrPrivs, "退药") Then
            CmdSend.Enabled = True
        Else
            CmdSend.Enabled = False
        End If
    End Select
End Sub

Public Sub SetParams()
    Dim bln是否配药确认 As Boolean
    
    With mcondition
        .bln显示付数 = (Val(zldatabase.GetPara("显示付数", glngSys, 1341)) = 1)
        .intShowPass = gintPass
        .bln显示大小单位 = (Val(zldatabase.GetPara("显示大小单位", glngSys, 1341)) = 1)
        .bln校验处方 = IsInString(gstrprivs, "校验处方", ";")
        .bln医嘱作废 = (gtype_UserSysParms.P68_门诊药嘱先作废后退药 = 0)
        .str配药人 = zldatabase.GetPara("配药人", glngSys, 1341)
        .bln自动配药 = (Val(zldatabase.GetPara("自动配药", glngSys, 1341)) = 1)
        .bln过滤模式 = (Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "界面定位", 0)) = 1)
        .int金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
        .bln取药确认 = (Val(zldatabase.GetPara("启用病人实际取药确认模式", glngSys, 1341, 0)) = 1)
        .bln处方审查 = ((gtype_UserSysParms.P240_药房处方审查 = 1 Or gtype_UserSysParms.P240_药房处方审查 = 3) And gtype_UserSysParms.P241_处方审查时机 = 2)
        .str核查人 = zldatabase.GetPara("核查人", glngSys, 1341)
        .bln允许核查人和配药人相同 = (Val(zldatabase.GetPara("允许核查人和配药人相同", glngSys, 1341, 0)) = 1)
        
        If mcondition.str配药人 = "|当前操作员|" Then
            mstrDosUser = gstrUserName
        Else
            mstrDosUser = mcondition.str配药人
        End If
        
        If mcondition.str核查人 = "|当前操作员|" Then
            mstr核查人 = gstrUserName
        Else
            mstr核查人 = mcondition.str核查人
        End If
    
        If .lng药房ID <> Val(zldatabase.GetPara("发药药房", glngSys, 1341)) Then
            .lng药房ID = Val(zldatabase.GetPara("发药药房", glngSys, 1341))
            .bln是否需要配药过程 = RecipeSendWork_DispensingMedi(.lng药房ID, bln是否配药确认)
            Call Load配药人(.lng药房ID)
            Call GetDrugDigit(.lng药房ID, "药品处方发药", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            
            '金额显示精度取费用精度
            mintMoneyDigit = Val(zldatabase.GetPara("费用金额保留位数", glngSys, 0))
            
            .int库存检查 = MediWork_GetCheckStockRule(.lng药房ID)
        End If
        
        .bln显示原产地 = Is中药库房(.lng药房ID)
        
        If .intListType = mListType.退药 Then
            Cbo配药人.Enabled = False
            cbo核查人.Enabled = False
        Else
            If zlStr.IsHavePrivs(mstrPrivs, "配药") = True Then
                If .bln自动配药 = False Then
                    Cbo配药人.Enabled = True
                Else
                    Cbo配药人.Enabled = False
                End If
            Else
                Cbo配药人.Enabled = False
            End If
            cbo核查人.Enabled = True
        End If
        
        Call Load核查人(.lng药房ID)
        
        cmdSendByNoTake.Visible = (.bln取药确认 And .intListType = mListType.待发药)
        
        Call Load配药人(.lng药房ID)
        Call Load核查人(.lng药房ID)
    
    End With
    
    
    
    mstrUserRecipeColor = zldatabase.GetPara("处方颜色", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal bln显示过程单据 As Boolean)
    Dim i As Integer
    
    With mcondition
        If .intListType <> intType Then
            SaveListColState
            .intListType = intType
            
            If .intListType = mListType.待发药 Or .intListType = mListType.超时未发 Then
                cbo开单医生.Enabled = True
            Else
                cbo开单医生.Enabled = False
            End If
            
            If .intListType <> mListType.退药 Then
                For i = 0 To Cbo配药人.ListCount - 1
                    If mstrDosUser = Cbo配药人.List(i) Then
                        Cbo配药人.ListIndex = i
                        Exit For
                    End If
                Next
                
                For i = 0 To cbo核查人.ListCount - 1
                    If mstr核查人 = cbo核查人.List(i) Then
                        cbo核查人.ListIndex = i
                        Exit For
                    End If
                Next
            End If
            
            If .intListType = mListType.退药 Then
                Cbo配药人.Enabled = False
                cbo核查人.Enabled = False
            Else
                If zlStr.IsHavePrivs(mstrPrivs, "配药") = True Then
                    If .bln自动配药 = False Then
                        Cbo配药人.Enabled = True
                    Else
                        Cbo配药人.Enabled = False
                    End If
                Else
                    Cbo配药人.Enabled = False
                End If
                cbo核查人.Enabled = True
            End If
        End If
        .bln显示过程单据 = bln显示过程单据
    End With
   
    Call SetComandBars(intType)
    
    InitList mcondition.intListType
    
    Call InitColSelList(intType)
    
    Select Case mcondition.intListType
        Case mListType.配药确认
            Me.cbo开单医生.Enabled = False
            Me.Cbo配药人.Enabled = False
            Me.cbo核查人.Enabled = False
            CmdSend.Caption = "配药确认(&O)"
            picRecipeColor.Visible = True
        Case mListType.待配药
            CmdSend.Caption = "配药(&V)"
            picRecipeColor.Visible = True
        Case mListType.已配药
            CmdSend.Caption = "取消配药(&C)"
            picRecipeColor.Visible = True
        Case mListType.待发药, mListType.超时未发
            CmdSend.Caption = "发药(&S)"
            picRecipeColor.Visible = True
        Case mListType.退药
            CmdSend.Caption = "退药(&R)"
            picRecipeColor.Visible = True
    End Select
    
    cmdSendByNoTake.Visible = (mcondition.bln取药确认 And mcondition.intListType = mListType.待发药)
    Chk全退.Enabled = (mcondition.intListType = mListType.退药)
    
    SetCmdSendPrivs mcondition.bln处方审查
    
    DoEvents
    Call Form_Resize
End Sub

Private Sub cbo核查人_Click()
    Dim i As Integer
    
    If mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发 Then
        mstr核查人 = Me.cbo核查人.Text
    End If
End Sub

Private Sub cbo核查人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo配药人.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(cbo核查人.hWnd) '获取位置
    gstrSQL = "Select ID, 编号, 姓名, 简码" & _
               " From 人员表" & _
               " Where ID In (Select 人员id From 部门人员 Where 部门id = [1]) And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) And" & _
                     " (编号 Like [2] Or 姓名 Like [2] Or 简码 Like [2])"
    
    Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "查询人员", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 500, blnCancel, False, True, mcondition.lng药房ID, IIf(gstrMatchMethod = 0, "%", "") & UCase(cbo核查人.Text) & "%")
        
    If rstemp Is Nothing Then
        Exit Sub
    End If
    For i = 0 To cbo核查人.ListCount
        If Mid(cbo核查人.List(i), InStr(1, cbo核查人.List(i), "-") + 1) = rstemp!姓名 Then
            cbo核查人.ListIndex = i
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo配药人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstemp As ADODB.Recordset
    Dim i As Integer
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo配药人.Text = "" Then Exit Sub
    
    On Error GoTo errHandle
    vRect = zlControl.GetControlRect(Cbo配药人.hWnd) '获取位置
    gstrSQL = "Select id, 编号,姓名,简码" & _
              "  From 人员表" & _
               " Where ID In (Select Distinct 人员id" & _
                            " From 人员性质说明 " & _
                            " Where 人员性质 = '药房发药人' And 人员id In (Select 人员id From 部门人员 Where 部门id = [1])) And" & _
                     " (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) and  (编号 like [2] or 姓名 like [2] or 简码 like [2])"
    
    Set rstemp = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "查询人员", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 500, blnCancel, False, True, mcondition.lng药房ID, IIf(gstrMatchMethod = 0, "%", "") & UCase(Cbo配药人.Text) & "%")
        
    If rstemp Is Nothing Then
        Exit Sub
    End If
    For i = 0 To Cbo配药人.ListCount
        If Mid(Cbo配药人.List(i), InStr(1, Cbo配药人.List(i), "-") + 1) = rstemp!姓名 Then
            Cbo配药人.ListIndex = i
        End If
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo配药人_Click()
    If mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发 Then
        mstrDosUser = Me.Cbo配药人.Text
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
   Dim Int单据 As Integer
    Dim strNo As String
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str挂号单 As String
    Dim lng主页ID As Long
    Dim lngCurrAdviceID As Long
    
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.Row = vsfList.rows - 1 Then Exit Sub
    
    Int单据 = vsfList.TextMatrix(vsfList.Row, mIntCol单据)
    strNo = vsfList.TextMatrix(vsfList.Row, mIntColNO)
    lngCurrAdviceID = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("医嘱id")))
    
    
    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
            strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!病人ID
            str挂号单 = NVL(rsTmp!挂号单)
            lng主页ID = rsTmp!主页id
            
    Select Case Control.Id
        Case conMenu_Tool_ShowPlug
            '功能：对病人过敏史/病生状态进行管理
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng主页ID, str挂号单)
        '弹出菜单：PASS命令
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng主页ID, str挂号单, lngCurrAdviceID)
    End Select
End Sub


Private Function AdviceCheckWarn(ByVal Int单据 As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号，lngCmd=0时需要
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String, lng药品id As Long, str单量单位 As String, str频率 As String
    Dim strsql As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lngPassPati As Long
    Dim lng主页ID As Long
    Dim str挂号单 As String
    Dim lngCount As Long
    Dim blnDo As Boolean
    

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
    strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] " & _
        " Union All " & _
        " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)

    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If

    lngPatiID = rsTmp!病人ID
    str挂号单 = zlStr.NVL(rsTmp!挂号单)
    lng主页ID = rsTmp!主页id
    

    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    If lngPatiID <> lngPassPati Then
        If str挂号单 <> "" Then               '门诊病人
            strsql = "Select 病人ID,Count(Distinct Trunc(登记时间)) as 就诊次数 From 病人挂号记录 Where 记录性质=1 And 记录状态=1 And 病人ID=[1] Group by 病人ID"
            strsql = "Select D.就诊次数,A.姓名,A.性别,A.出生日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,(" & strsql & ") D,人员表 E" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=D.病人ID" & _
                " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, str挂号单)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, rsTmp!就诊次数, rsTmp!姓名, zlStr.NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.NVL(rsTmp!医生码) & "/" & zlStr.NVL(rsTmp!医生名), ""), "")
        Else                                    '住院病人
            strsql = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And A.主页id=b.主页id And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, lng主页ID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, lng主页ID, rsTmp!姓名, zlStr.NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.NVL(rsTmp!医生码) & "/" & zlStr.NVL(rsTmp!医生名), ""), _
                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        End If
        lngPassPati = lngPatiID
    End If
    
    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        '取药品名称
        str药品 = vsfList.TextMatrix(lngRow, mIntCol药品名称)
        lng药品id = vsfList.TextMatrix(lngRow, mintcol药品id)
        str单量单位 = Mid(vsfList.TextMatrix(lngRow, mIntCol单量), InStr(vsfList.TextMatrix(lngRow, mIntCol单量), "(") + 1)
        If InStr(str单量单位, ")") > 0 Then str单量单位 = Replace(str单量单位, ")", "")
        '取药品给药途径
        str用法 = vsfList.TextMatrix(lngRow, mIntCol用法)

        If InStr(str药品, " ") > 0 Then str药品 = Left(str药品, InStr(str药品, " ") - 1)
        If InStr(str药品, "]") > 0 Then str药品 = Mid(str药品, InStr(str药品, "]") + 1, Len(str药品) - InStr(str药品, "]"))
        '传入查询药品信息
        Call PassSetQueryDrug(lng药品id, str药品, str单量单位, str用法)

        '设置菜单可用状态
        Call SetPassMenuState

        AdviceCheckWarn = 1 '表示可以弹出菜单

        Screen.MousePointer = 0: Exit Function
    ElseIf lngCmd = 6 Then
        Call PassSetWarnDrug(Val(vsfList.TextMatrix(lngRow, mIntCol相关id))) '单药警告(已警告的医嘱唯一码)
    Else
        With Me.vsfList
            '用药审核或用药研究
            lngCount = 0
            str药品 = "": str用法 = "": str频率 = ""
            i = 1
            If .TextMatrix(i, mIntCol开嘱医生) <> "" Then
                strsql = "select 编号 from 人员表 where 姓名=[1]"
                Set rsTmp = zldatabase.OpenSQLRecord(strsql, "", .TextMatrix(i, mIntCol开嘱医生))
            End If
            For i = .FixedRows To .rows - 1
                blnDo = Val(.TextMatrix(i, mIntCol医嘱id)) <> 0 And Val(.TextMatrix(i, mintcol药品id)) <> 0
                If blnDo Then
                    '取药品名称
                    str药品 = .TextMatrix(i, mIntCol药品名称)
                    If InStr(str药品, "]") > 0 Then str药品 = Mid(str药品, InStr(str药品, "]") + 1, Len(str药品) - InStr(str药品, "]"))
                    
                    '取药品给药途径
                    str用法 = .TextMatrix(i, mIntCol用法)
                    
                    '取用药频率(次/天),都为整数四舍五入
                    If .TextMatrix(i, mIntCol间隔单位) = "天" Then
                        str频率 = .TextMatrix(i, mIntCol频率次数) & "/" & .TextMatrix(i, mIntCol频率间隔)
                    ElseIf .TextMatrix(i, mIntCol间隔单位) = "周" Then
                        str频率 = .TextMatrix(i, mIntCol频率次数) & "/7"
                    ElseIf .TextMatrix(i, mIntCol间隔单位) = "小时" Then
                        If Val(.TextMatrix(i, mIntCol频率间隔)) <= 24 Then
                            str频率 = Format(24 / Val(.TextMatrix(i, mIntCol频率间隔)) * Val(.TextMatrix(i, mIntCol频率次数)), "0") & "/1"
                        Else
                            str频率 = Val(.TextMatrix(i, mIntCol频率次数)) & "/" & Format(Val(.TextMatrix(i, mIntCol频率间隔)) / 24, "0")
                        End If
                    ElseIf .TextMatrix(i, mIntCol间隔单位) = "分钟" Then
                        str频率 = Format((24 * 60) / Val(.TextMatrix(i, mIntCol频率间隔)) * Val(.TextMatrix(i, mIntCol频率次数)), "0") & "/1"
                    End If
                    
'                    MsgBox "医嘱id：" & .TextMatrix(i, mIntCol医嘱id) & "；药品id:" & .TextMatrix(i, mIntCol药品ID) & ";药品：" & str药品 & "；单量：" & Mid(.TextMatrix(i, mIntCol单量), 1, InStr(1, .TextMatrix(i, mIntCol单量), "(") - 1) & ";" & _
'                            "单位：" & Mid(.TextMatrix(i, mIntCol单量), InStr(1, .TextMatrix(i, mIntCol单量), "(") + 1, InStr(1, .TextMatrix(i, mIntCol单量), ")") - InStr(1, .TextMatrix(i, mIntCol单量), "(") - 1) & ";用药频率" & str频率 & ";" & _
'                            "开始时间：" & Format(.TextMatrix(i, mIntCol开始时间), "yyyy-MM-dd") & ";结束时间：" & Format(.TextMatrix(i, mIntCol结束时间), "yyyy-MM-dd") & ";用法：" & str用法 & _
'                            "相关id：" & .TextMatrix(i, mIntCol相关id) & ";医嘱标志：" & .TextMatrix(i, mIntCol医嘱标志) & ";医生情况：" & rsTmp!编号 & "\" & .TextMatrix(i, mIntCol开嘱医生)
                    '传入医嘱信息
                    Call PassSetRecipeInfo(.TextMatrix(i, mIntCol医嘱id), .TextMatrix(i, mintcol药品id), str药品, _
                        Mid(.TextMatrix(i, mIntCol单量), 1, InStr(1, .TextMatrix(i, mIntCol单量), "(") - 1), Mid(.TextMatrix(i, mIntCol单量), InStr(1, .TextMatrix(i, mIntCol单量), "(") + 1, InStr(1, .TextMatrix(i, mIntCol单量), ")") - InStr(1, .TextMatrix(i, mIntCol单量), "(") - 1), str频率, _
                        Format(.TextMatrix(i, mIntCol开始时间), "yyyy-MM-dd"), Format(.TextMatrix(i, mIntCol结束时间), "yyyy-MM-dd"), str用法, _
                        .TextMatrix(i, mIntCol相关id), .TextMatrix(i, mIntCol医嘱标志), rsTmp!编号 & "\" & .TextMatrix(i, mIntCol开嘱医生))
                    lngCount = lngCount + 1
                End If
            Next
            
            '无可审查的药品
            If (lngCmd = 1 Or lngCmd = 2 Or lngCmd = 3) And lngCount = 0 Then
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
Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    On Error Resume Next
    
    If cbsMain.count > 1 Then
        Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

        picRecipt.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop
    End If
End Sub


Sub InitList(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    Dim bln列设置无效 As Boolean
    
    '''初始化列顺序
    '默认列顺序
    mIntCol当前行 = 0
    mIntCol顺序号 = 1
    mIntCol审查结果 = 2
    mIntCol药品名称 = 3
    mintCol皮试结果 = 4
    mIntCol其它名 = 5
    mIntCol英文名 = 6
    mIntCol配方名称 = 7
    mintcol序号 = 8
    mintcol规格 = 9
    mintcol生产商 = 10
    mintcol原产地 = 11
    mintcol批号 = 12
    mintcol效期 = 13
    mintcol单位 = 14
    mIntCol单价 = 15
    mIntCol付数 = 16
    mintcol数量 = 17
    mIntCol金额 = 18
    mIntCol实收金额 = 19
    mIntCol重量 = 20
    mIntCol单量 = 21
    mIntCol用法 = 22
    mIntCol频次 = 23
    mIntCol超量说明 = 24
    mIntCol用药目的 = 25
    mIntCol医生嘱托 = 26
    mIntCol费别 = 27
    mIntCol库存数 = 28
    mIntCol货位 = 29
    mIntCol已退数 = 30
    mIntCol准退数 = 31
    mIntCol准退数大 = 32
    mIntCol准退数小 = 33
    mIntCol退药数 = 34
    mIntCol退药数大 = 35
    mIntCol单位大 = 36
    mIntCol退药数小 = 37
    mIntCol单位小 = 38
    mIntCol备注 = 39
    '--------------以下列为不可见--------------
    mIntCol分批 = 40
    mIntCol新批号 = 41
    mIntCol新效期 = 42
    mIntCol新产地 = 43
    mIntCol医嘱id = 44
    mIntCol实际数量 = 45
    mIntCol包装 = 46
    mIntCol单据 = 47
    mIntColNO = 48
    mIntCol门诊标志 = 49
    mIntCol记录性质 = 50
    mIntCol库房ID = 51
    mIntCol用药理由 = 52
    mIntCol相关id = 53
    mIntCol开嘱医生 = 54
    mIntCol频率间隔 = 55
    mIntCol间隔单位 = 56
    mIntCol医嘱标志 = 57
    mIntCol开始时间 = 58
    mIntCol结束时间 = 59
    mIntCol频率次数 = 60
    mIntCol警告 = 61
    mIntCol门诊号 = 62
    mIntCol住院号 = 63
    mIntColId = 64
    mintcol药品id = 65
    mintcol批次 = 66
    mIntCol禁忌药品说明 = 67
    
    
    '恢复用户自定义列顺序
    str列设置 = LoadListColState
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                If Split(arr列设置(n), ",")(0) = "" Then
                    bln列设置无效 = True
                    Exit For
                End If
            Next
            
            If bln列设置无效 = True Then
                str列设置 = ""
            Else
                For n = 0 To UBound(arr列设置)
                    SetColumnValue Split(arr列设置(n), ",")(0), n
                Next
            End If
        End If
    End If
    
    '初始化未发药清单
    With vsfList
        .Redraw = flexRDNone
        
        .rows = 1
        .rows = 2
        .Cols = mconIntCol列数
        
        .Cell(flexcpPicture, 1, mIntCol当前行, 1, mIntCol当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol当前行, .rows - 1, mIntCol当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol当前行, "", 250, flexAlignCenterCenter, "当前行"
        VsfGridColFormat vsfList, mIntCol顺序号, "序号", 450, flexAlignRightCenter, "顺序号"
        
        If IsInString(gstrprivs, "合理用药监测", ";") And Not gobjPass Is Nothing Then
            VsfGridColFormat vsfList, mIntCol审查结果, "警", 280, flexAlignCenterCenter, "审查结果"
        Else
            VsfGridColFormat vsfList, mIntCol审查结果, "警", 0, flexAlignCenterCenter, "审查结果"
        End If
        
        VsfGridColFormat vsfList, mIntCol药品名称, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfList, mintCol皮试结果, "", 400, flexAlignLeftCenter, "皮试结果"
        VsfGridColFormat vsfList, mIntCol其它名, "其它名", 2000, flexAlignLeftCenter, "其它名"
        VsfGridColFormat vsfList, mIntCol英文名, "英文名", 2000, flexAlignLeftCenter, "英文名"
        VsfGridColFormat vsfList, mIntCol配方名称, "配方名称", 1800, flexAlignLeftCenter, "配方名称"
        VsfGridColFormat vsfList, mintcol序号, "序号", 0, flexAlignCenterCenter, "序号"
        VsfGridColFormat vsfList, mintcol规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfList, mintcol批号, "批号", 1500, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList, mintcol效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfList, mIntColId, "Id", 0, flexAlignCenterCenter, "Id"
        VsfGridColFormat vsfList, mintcol药品id, "药品ID", 0, flexAlignCenterCenter, "药品ID"
        
        VsfGridColFormat vsfList, mintcol批次, "批次", 0, flexAlignRightCenter, "批次"
        VsfGridColFormat vsfList, mintcol单位, "单位", IIf(intType = mListType.退药 And mcondition.bln显示大小单位, 0, 500), flexAlignCenterCenter, "单位"
        VsfGridColFormat vsfList, mIntCol单价, "单价", 1000, flexAlignRightCenter, "单价"
        VsfGridColFormat vsfList, mIntCol付数, "付数", IIf(mcondition.bln显示付数, 800, 0), flexAlignRightCenter, "付数"
        VsfGridColFormat vsfList, mintcol数量, "数量", 1200, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList, mIntCol金额, "应收金额", IIf(mcondition.int金额显示 = 1, 0, 1000), flexAlignRightCenter, "应收金额"
        VsfGridColFormat vsfList, mIntCol实收金额, "实收金额", IIf(mcondition.int金额显示 = 0, 0, 1000), flexAlignRightCenter, "实收金额"
        VsfGridColFormat vsfList, mIntCol重量, "重量", 1200, flexAlignRightCenter, "重量"
        VsfGridColFormat vsfList, mIntCol单量, "单量", 1200, flexAlignCenterCenter, "单量"
        VsfGridColFormat vsfList, mIntCol用法, "用法", 1500, flexAlignLeftCenter, "用法"
        VsfGridColFormat vsfList, mIntCol频次, "频次", 1500, flexAlignLeftCenter, "频次"
        VsfGridColFormat vsfList, mIntCol超量说明, "超量说明", 1500, flexAlignLeftCenter, "超量说明"
        VsfGridColFormat vsfList, mIntCol用药目的, "用药目的", 0, flexAlignLeftCenter, "用药目的"
        
       
        VsfGridColFormat vsfList, mIntCol医生嘱托, "医生嘱托", IIf(intType = mListType.退药, 0, 1500), flexAlignLeftCenter, "医生嘱托"
        VsfGridColFormat vsfList, mIntCol费别, "费别", 1000, flexAlignLeftCenter, "费别"
        VsfGridColFormat vsfList, mIntCol库存数, "库存数", IIf(intType = mListType.退药, 0, 1200), flexAlignRightCenter, "库存数"
        VsfGridColFormat vsfList, mIntCol货位, "库房货位", IIf(intType = mListType.退药, 0, 1200), flexAlignLeftCenter, "库房货位"
        VsfGridColFormat vsfList, mIntCol已退数, "已退数", IIf(intType = mListType.退药, 1200, 0), flexAlignRightCenter, "已退数"
        VsfGridColFormat vsfList, mIntCol准退数, "准退数", IIf(intType = mListType.退药, 1200, 0), flexAlignRightCenter, "准退数"
        VsfGridColFormat vsfList, mIntCol准退数大, "准退数大", 0, flexAlignCenterCenter, "准退数大"
        VsfGridColFormat vsfList, mIntCol准退数小, "准退数小", 0, flexAlignCenterCenter, "准退数小"
        VsfGridColFormat vsfList, mIntCol退药数, "退药数", IIf(intType = mListType.退药 And mcondition.bln显示大小单位 = False, 1200, 0), flexAlignRightCenter, "退药数"
        VsfGridColFormat vsfList, mIntCol退药数大, "退药数(大包装)", IIf(intType = mListType.退药 And mcondition.bln显示大小单位 = True, 1500, 0), flexAlignRightCenter, "退药数(大包装)"
        
        VsfGridColFormat vsfList, mIntCol单位大, "单位(大)", IIf(intType = mListType.退药 And mcondition.bln显示大小单位 = True, 500, 0), flexAlignCenterCenter, "单位(大)"
        VsfGridColFormat vsfList, mIntCol退药数小, "退药数(小包装)", IIf(intType = mListType.退药 And mcondition.bln显示大小单位 = True, 1500, 0), flexAlignRightCenter, "退药数(小包装)"
        VsfGridColFormat vsfList, mIntCol单位小, "单位(小)", IIf(intType = mListType.退药 And mcondition.bln显示大小单位 = True, 500, 0), flexAlignCenterCenter, "单位(小)"
        VsfGridColFormat vsfList, mIntCol分批, "分批", 0, flexAlignCenterCenter, "分批"
        VsfGridColFormat vsfList, mIntCol新批号, "新批号", 0, flexAlignCenterCenter, "新批号"
        VsfGridColFormat vsfList, mIntCol新效期, "新效期", 0, flexAlignCenterCenter, "新效期"
        VsfGridColFormat vsfList, mIntCol新产地, "新产地", 0, flexAlignCenterCenter, "新产地"
        VsfGridColFormat vsfList, mIntCol备注, "备注", 1200, flexAlignLeftCenter, "备注"
        VsfGridColFormat vsfList, mintcol生产商, "生产商", 1200, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfList, mintcol原产地, "原产地", 1200, flexAlignLeftCenter, "原产地"
        VsfGridColFormat vsfList, mIntCol医嘱id, "医嘱id", 0, flexAlignCenterCenter, "医嘱id"
        VsfGridColFormat vsfList, mIntCol实际数量, "实际数量", 0, flexAlignCenterCenter, "实际数量"
        
        VsfGridColFormat vsfList, mIntCol包装, "包装", 0, flexAlignLeftCenter, "包装"
        VsfGridColFormat vsfList, mIntCol单据, "单据", 0, flexAlignLeftCenter, "单据"
        VsfGridColFormat vsfList, mIntColNO, "No", 0, flexAlignLeftCenter, "NO"
        
        VsfGridColFormat vsfList, mIntCol门诊标志, "门诊标志", 0, flexAlignLeftCenter, "门诊标志"
        VsfGridColFormat vsfList, mIntCol记录性质, "记录性质", 0, flexAlignLeftCenter, "记录性质"
        VsfGridColFormat vsfList, mIntCol库房ID, "库房ID", 0, flexAlignLeftCenter, "库房ID"
        VsfGridColFormat vsfList, mIntCol用药理由, "用药理由", 0, flexAlignLeftCenter, "用药理由"
        VsfGridColFormat vsfList, mIntCol相关id, "相关id", 0, flexAlignLeftCenter, "相关id"
        VsfGridColFormat vsfList, mIntCol开嘱医生, "开嘱医生", 0, flexAlignLeftCenter, "开嘱医生"
        VsfGridColFormat vsfList, mIntCol频率间隔, "频率间隔", 0, flexAlignLeftCenter, "频率间隔"
        VsfGridColFormat vsfList, mIntCol间隔单位, "间隔单位", 0, flexAlignLeftCenter, "间隔单位"
        VsfGridColFormat vsfList, mIntCol医嘱标志, "医嘱标志", 0, flexAlignLeftCenter, "医嘱标志"
        VsfGridColFormat vsfList, mIntCol开始时间, "开始时间", 0, flexAlignLeftCenter, "开始时间"
        VsfGridColFormat vsfList, mIntCol结束时间, "结束时间", 0, flexAlignLeftCenter, "结束时间"
        VsfGridColFormat vsfList, mIntCol频率次数, "频率次数", 0, flexAlignLeftCenter, "频率次数"
        VsfGridColFormat vsfList, mIntCol警告, "警告", 0, flexAlignLeftCenter, "警告"
        VsfGridColFormat vsfList, mIntCol门诊号, "门诊号", 0, flexAlignLeftCenter, "门诊号"
        VsfGridColFormat vsfList, mIntCol住院号, "住院号", 0, flexAlignLeftCenter, "住院号"
        VsfGridColFormat vsfList, mIntCol禁忌药品说明, "禁忌药品说明", 0, flexAlignLeftCenter, "禁忌药品说明"
        
        mstrUnallowShow = "当前行;序号;皮试结果;Id;药品ID;批次;用药目的;分批;新批号;新效期;新产地;医嘱id;实际数量;包装;单据;NO;门诊标志;记录性质;库房ID;用药理由;相关id;开嘱医生;频率间隔;间隔单位;医嘱标志;开始时间;结束时间;频率次数;警告;门诊号;住院号;禁忌药品说明"
        If mcondition.int金额显示 = 0 Then mstrUnallowShow = mstrUnallowShow & ";实收金额"
        If mcondition.int金额显示 = 1 Then mstrUnallowShow = mstrUnallowShow & ";应收金额"
        
        If intType <> mListType.退药 Then
            mstrUnallowSetColHide = "药品名称;数量"
            mstrUnallowShow = mstrUnallowShow & ";" & "退药数(大包装);退药数(小包装);已退数;准退数;准退数大;准退数小;退药数;退药数(大包装);单位(大);退药数(小包装);单位(小)"
        Else
            mstrUnallowShow = mstrUnallowShow & ";" & "医生嘱托;库存数;库房货位"
            If mcondition.bln显示大小单位 Then
                mstrUnallowSetColHide = "药品名称;数量;已退数;准退数;退药数(大包装);退药数(小包装)"
                mstrUnallowShow = mstrUnallowShow & ";" & "单位;退药数;准退数大;准退数小"
            Else
                mstrUnallowSetColHide = "药品名称;数量;已退数;准退数;退药数"
                mstrUnallowShow = mstrUnallowShow & ";" & "准退数大;准退数小;退药数(大包装);退药数(小包装);单位(大);单位(小)"
            End If
        End If
        
        If mcondition.intShowPass <> 0 Or Not IsInString(gstrprivs, "合理用药监测", ";") Then mstrUnallowShow = mstrUnallowShow & ";" & "审查结果"
        If mcondition.bln显示付数 = False Then mstrUnallowShow = mstrUnallowShow & ";" & "付数"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow, Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr列设置(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If intType = mListType.退药 And mcondition.bln显示大小单位 = True Then
            VsfGridColFormat vsfList, mIntCol退药数, "退药数", 0, flexAlignRightCenter, "退药数"
            VsfGridColFormat vsfList, mIntCol退药数大, "退药数(大包装)", 1500, flexAlignRightCenter, "退药数(大包装)"
            
            VsfGridColFormat vsfList, mIntCol单位大, "单位(大)", 500, flexAlignCenterCenter, "单位(大)"
            VsfGridColFormat vsfList, mIntCol退药数小, "退药数(小包装)", 1500, flexAlignRightCenter, "退药数(小包装)"
            VsfGridColFormat vsfList, mIntCol单位小, "单位(小)", 500, flexAlignCenterCenter, "单位(小)"
        End If
        
        '只有中药类库房才显示"原产地"列
        If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList, mintcol原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        
        '重新生成网格
        .Select 0, 0, .rows - 1, .Cols - 1
        .CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        .Select 0, mIntCol药品名称, vsfList.rows - 1, mintCol皮试结果
        .CellBorder &H9D9D9D, -1, -1, -1, -1, 0, 1
        
        .RowSel = 1
        
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub SaveListColState()
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    If vsfList.Cols <> mconIntCol列数 Then Exit Sub
    
    Select Case mcondition.intListType
        Case mListType.配药确认
            strType = "配药确认"
        Case mListType.待配药
            strType = "待配药"
        Case mListType.已配药
            strType = "已配药"
        Case mListType.待发药
            strType = "待发药"
        Case mListType.超时未发
            strType = "超时未发"
        Case mListType.退药
            strType = "退药"
    End Select
    
    With vsfList
        For i = 0 To .Cols - 1
            If vsfList.ColKey(i) = "" Then
                MsgBox "AA"
            End If
            str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & vsfList.ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, str列设置)
End Sub

Private Function LoadListColState() As String
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    
    Select Case mcondition.intListType
        Case mListType.配药确认
            strType = "配药确认"
        Case mListType.待配药
            strType = "待配药"
        Case mListType.已配药
            strType = "已配药"
        Case mListType.待发药
            strType = "待发药"
        Case mListType.超时未发
            strType = "超时未发"
        Case mListType.退药
            strType = "退药"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function

Public Function RefreshList(ByVal RecData As ADODB.Recordset, Optional ByVal strWeight As String, Optional ByVal int可操作 As Integer = 0, Optional ByVal int排队状态 As Integer, Optional ByVal int审查结果 As Integer) As Boolean
    Dim dbl应收金额, dbl实收金额 As Double
    Dim IntLocate As Integer
    Dim str操作员 As String
    Dim dbl总重量 As Double
    Dim str重量单位 As String
    Dim lng整数量 As Long
    Dim dbl小数量 As Double
    Dim int门诊 As Integer
    Dim intRow As Integer
    Dim strDiag As String
    Dim strSum As String
    Dim i As Integer
    Dim bln皮试 As Boolean
    Dim bln中药处方 As Boolean
    Dim dateCurrent As Date
    Dim dbl实际数量 As Double
    
    dateCurrent = Sys.Currentdate
    
    CmdSend.Enabled = False
    mInt可操作 = int可操作
    
    If Chk全退.Enabled = True Then Chk全退.Value = 1
    
    mcondition.bln显示过程单据 = (Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "显示退药过程单据", 1)) = 1)
    
    Set mrsDetail = RecData
    
    SaveListColState
    InitList mcondition.intListType
    
    Lbl配药人.Caption = IIf(mcondition.intListType = mListType.退药, IIf(int可操作 <> 3, "发药人", "退药人"), "配药人")
    
    RefreshList = False

    dbl应收金额 = 0
    dbl实收金额 = 0
    txtNo.Clear
    
    '如果不是原始记录，则不显示列（已退数、准退数）
    vsfList.ColWidth(mIntCol已退数) = 0
    vsfList.ColWidth(mIntCol准退数) = 0
    vsfList.ColWidth(mIntCol退药数) = 0
    If mcondition.intListType = mListType.退药 And int可操作 = 1 Then
        vsfList.ColWidth(mIntCol已退数) = 1000
        vsfList.ColWidth(mIntCol准退数) = 1000
        vsfList.ColWidth(mIntCol退药数) = IIf(mcondition.bln显示大小单位 = False, 1000, 0)
    End If
    
    '填充单据内容
    With mrsDetail
        If .EOF Then
            Call FormClear
        Else
            If !记录性质 = 1 Or (!记录性质 = 2 And (!门诊标志 = 1 Or !门诊标志 = 4)) Then
                int门诊 = 1
            Else
                int门诊 = 2
            End If
            
            '确定病人未取药按钮的显示状态
            cmdSendByNoTake.Visible = (mcondition.bln取药确认 And mcondition.intListType = mListType.待发药 And int门诊 = 1)
                
            '填充表头
            Me.lbl姓名(1).Caption = IIf(IsNull(!姓名), "", !姓名)
            Me.lbl姓名(1).ForeColor = zldatabase.GetPatiColor(IIf(IsNull(!病人类型), "", !病人类型))
            
            If mcondition.intListType <> mListType.退药 Then
                If zlStr.NVL(!结算模式, 0) = 1 Then
                    Me.picMark1.Visible = True
                Else
                    Me.picMark1.Visible = False
                End If
            Else
                Me.picMark1.Visible = False
            End If
            
            Me.Lbl床号(1).Caption = IIf(IsNull(!床号), "", !床号)
            If !单据 = 8 Then Me.Lbl床号(1).Caption = ""
            Me.cbo开单医生.ListIndex = 0
            If (mcondition.bln校验处方 = False) And zlStr.IsHavePrivs(gstrprivs, "医生查询") Then
                str操作员 = IIf(IsNull(!开单人), "", !开单人)
            Else
                If mcondition.intListType = mListType.退药 And mcondition.bln校验处方 = True Then
                    str操作员 = IIf(IsNull(!填制人), "", !填制人)
                Else
                    str操作员 = ""
                End If
            End If
            If str操作员 <> "" Then
                '定位医生
                For IntLocate = 1 To cbo开单医生.ListCount
                    If Mid(cbo开单医生.List(IntLocate), InStr(1, cbo开单医生.List(IntLocate), "-") + 1) = str操作员 Then
                        cbo开单医生.ListIndex = IntLocate
                        Exit For
                    End If
                Next
            End If
            
            cbo开单医生.Enabled = ((mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发) And mcondition.bln校验处方 = True And cbo开单医生.ListIndex = 0)
            
            Me.Lbl科室(1).Caption = IIf(IsNull(!科室), "", !科室)
            Me.Lbl年龄(1).Caption = IIf(IsNull(!年龄), "", !年龄)
            Me.lbl就诊卡号(1).Caption = IIf(IsNull(!就诊卡号), "", !就诊卡号)
            Me.LblTel(1).Caption = IIf(IsNull(!联系人电话), "", !联系人电话)
            
            If mcondition.intListType = mListType.退药 Then
                If IIf(IsNull(!发药人), "", !发药人) <> "" Then
                    Me.Cbo配药人 = IIf(IsNull(!发药人), "", !发药人)
                End If
            Else
            
                If IIf(IsNull(!配药人), "", !配药人) <> "" Then
                    Me.Cbo配药人 = IIf(IsNull(!配药人), "", !配药人)
                End If
            End If
            
            
            If mcondition.intListType = mListType.退药 Then
                If IIf(IsNull(!核查人), "", !核查人) <> "" Then
                    Me.cbo核查人 = IIf(IsNull(!核查人), "", !核查人)
                End If
            End If
            
            Me.Lbl收费员(1).Caption = IIf(IsNull(!操作员姓名), "", !操作员姓名)
            Me.Lbl性别(1).Caption = IIf(IsNull(!性别), "", !性别)
            Me.Lbl住院号(1).Caption = IIf(IsNull(!住院号), "", !住院号)
            
'            If mcondition.intListType <> mListType.退药 Then
                picRecipeColor.BackColor = Val(Split(mstrUserRecipeColor, ";")(Val(!处方类型)))
                lblRecipeType.Caption = Split(gconstrRecipeType, ";")(Val(!处方类型))
'            Else
'                picRecipeColor.BackColor = &HFFFFFF
'                lblRecipeType.Caption = "处方"
'            End If
            '82922,显示体重
            Me.LblWeight(1).Caption = IIf(IsNumeric(strWeight), strWeight & "kg", strWeight)
            
            '诊断信息
            txt诊断内容.Text = ""
            txt诊断内容.Tag = ""
            txt诊断内容.Height = 180
            
            Call picRecInfo_Resize
            Call picRecipt_Resize
            
            txtNo.AddItem !NO & "--" & !姓名
            txtNo.ItemData(txtNo.NewIndex) = !单据
            txtNo.Tag = !药房ID & "|" & !药房 & "|" & !门诊标志 & "|" & !记录性质 & "|" & !填制日期 & "|" & !状态
            Lbl药房.Caption = !药房
            
            If 判断是否中药处方(!药房ID, !单据, !NO) Then
                bln中药处方 = True
            End If

            mblnAllowClick = False
            txtNo.ListIndex = 0
            mblnAllowClick = True

            '是否显示合理用药
            If (mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发) And mcondition.intShowPass = 1 And IsInString(gstrprivs, "合理用药监测", ";") Then
                Dim cbrControl As CommandBarControl
                
'                Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
'                If Not cbrControl Is Nothing Then cbrControl.Enabled = Check是否存在医嘱(Val(!单据), !NO)
            End If
            
            vsfList.rows = 1
            vsfList.Redraw = False
            
            Do While Not .EOF
                '60022诊断信息显示
                strDiag = RecipeSendWork_GetDiagnosis(int门诊, IIf(int门诊 = 1, Val(!医嘱id), Val(!病人ID)), Val(!主页id), IIf(bln中药处方, 1, 2))
                '修正住院医生工作站将医嘱发送到门诊收费,导致诊断信息为空的情况
                If int门诊 = 1 And !在院 = 1 And strDiag = "" Then
                    int门诊 = 2
                    strDiag = RecipeSendWork_GetDiagnosis(int门诊, IIf(int门诊 = 1, Val(!医嘱id), Val(!病人ID)), Val(!主页id), IIf(bln中药处方, 1, 2))
                End If
                
                If strDiag <> "" Then
                    strDiag = strDiag & "|"
                    For i = 0 To UBound(Split(strDiag, "|"))
                        If Split(strDiag, "|")(i) <> "" Then
                            If InStr(1, txt诊断内容.Text & " ※", "※" & Split(strDiag, "|")(i) & " ※") < 1 Then
                                txt诊断内容.Text = IIf(txt诊断内容.Text = "", " ※", txt诊断内容.Text & " ※") & Split(strDiag, "|")(i)
                                txt诊断内容.Tag = IIf(txt诊断内容.Tag = "", "※ ", txt诊断内容.Tag & vbCrLf & "※ ") & Split(strDiag, "|")(i)
                            End If
                        End If
                    Next
                End If
            
                intRow = intRow + 1
                vsfList.rows = intRow + 1

                vsfList.TextMatrix(intRow, mIntCol顺序号) = intRow
                If Val(!高危药品) <> 0 Then
                    vsfList.Cell(flexcpPicture, intRow, mIntCol药品名称) = Me.ImgList.ListImages(40).Picture
                    vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol药品名称) = flexPicAlignLeftCenter
                Else
                    If Val(!抗生素) <> 0 Then
                        vsfList.Cell(flexcpPicture, intRow, mIntCol药品名称) = Me.ImgList.ListImages(39).Picture
                        vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol药品名称) = flexPicAlignLeftCenter
                    End If
                End If
                vsfList.TextMatrix(intRow, mIntCol药品名称) = !品名
                
                If Not bln中药处方 And !是否皮试 = 1 Then
                    vsfList.TextMatrix(intRow, mintCol皮试结果) = Get皮试结果(!病人ID, !药名ID, dateCurrent, !开嘱时间)
                    If vsfList.TextMatrix(intRow, mintCol皮试结果) <> "" Then
                        bln皮试 = True
                    End If
                End If
                
                vsfList.TextMatrix(intRow, mIntCol其它名) = IIf(IsNull(!其它名), "", !其它名)
                vsfList.TextMatrix(intRow, mIntCol英文名) = IIf(IsNull(!英文名), "", !英文名)
                vsfList.TextMatrix(intRow, mIntCol配方名称) = IIf(IsNull(!配方名称), "", !配方名称)
                vsfList.TextMatrix(intRow, mintcol序号) = !序号
                vsfList.TextMatrix(intRow, mintcol规格) = IIf(IsNull(!规格), "", !规格)
                vsfList.TextMatrix(intRow, mintcol批号) = IIf(IsNull(!批号), "", !批号)
                vsfList.TextMatrix(intRow, mintcol效期) = IIf(IsNull(!效期), "", !效期)
                vsfList.TextMatrix(intRow, mIntColId) = !收发ID
                vsfList.TextMatrix(intRow, mintcol药品id) = !药品ID
                vsfList.TextMatrix(intRow, mintcol批次) = !批次
                vsfList.TextMatrix(intRow, mintcol单位) = IIf(IsNull(!单位), "", !单位)
                vsfList.TextMatrix(intRow, mIntCol单价) = Format(!单价, "#0." & String(mintPriceDigit, "0"))
                vsfList.TextMatrix(intRow, mIntCol付数) = Format(!付数, "#####0;-#####0; ;")
                vsfList.TextMatrix(intRow, mIntCol单据) = !单据
                vsfList.TextMatrix(intRow, mIntColNO) = !NO
                vsfList.TextMatrix(intRow, mIntCol住院号) = zlStr.NVL(!住院号)
                vsfList.TextMatrix(intRow, mIntCol门诊号) = zlStr.NVL(!门诊号)
                vsfList.TextMatrix(intRow, mIntCol禁忌药品说明) = zlStr.NVL(!禁忌药品说明)
                vsfList.TextMatrix(intRow, mintcol生产商) = zlStr.NVL(!产地)
                vsfList.TextMatrix(intRow, mintcol原产地) = zlStr.NVL(!原产地)
                
                If mcondition.bln显示大小单位 = True Then
                    '按大小包装显示数量
                    lng整数量 = Int(!数量)
                    If !售价单位 = !单位 Then
                        '售价单位和门诊单位名称相同，则可以直接用门诊单位，无需单位的换算
                        vsfList.TextMatrix(intRow, mintcol数量) = !数量 & IIf(IsNull(!售价单位), "", !售价单位)
                    Else
                        '售价单位和门诊单位不同，需要进行换算
                        If !实际数量 = 0 Then
                            dbl实际数量 = !小单位数量
                        Else
                            dbl实际数量 = !实际数量
                        End If
                        
                        If dbl实际数量 < 0 Then
                            lng整数量 = -Int(Abs(dbl实际数量) / !包装)
                        Else
                            lng整数量 = Int(dbl实际数量 / !包装)
                        End If

                        If lng整数量 = 0 Then
                            '门诊数量小于1时，用售价单位的实际数量直接显示
                            vsfList.TextMatrix(intRow, mintcol数量) = Abs(dbl实际数量) & IIf(IsNull(!售价单位), "", !售价单位)
                        Else
                            '门诊数量大于1时，需要考虑是否有零头的情况。例如:1板3片
                            If (dbl实际数量 / !包装) = lng整数量 Then
                                '没有零头
                                vsfList.TextMatrix(intRow, mintcol数量) = Abs(lng整数量) & IIf(IsNull(!单位), "", !单位)
                            Else
                                '有零头
                                vsfList.TextMatrix(intRow, mintcol数量) = Abs(lng整数量) & IIf(IsNull(!单位), "", !单位) & Abs((dbl实际数量 - (lng整数量 * !包装))) & IIf(IsNull(!售价单位), "", !售价单位)
                            End If
                        End If
                    End If
                    
                    If !数量 < 0 Then
                        vsfList.TextMatrix(intRow, mintcol数量) = "（退）" & vsfList.TextMatrix(intRow, mintcol数量)
                    End If
                    vsfList.TextMatrix(intRow, mIntCol包装) = Val(!包装)
                Else
                    vsfList.TextMatrix(intRow, mintcol数量) = zlStr.FormatEx(!数量, 5)
                End If
                
                vsfList.TextMatrix(intRow, mIntCol金额) = zlStr.FormatEx(Val(!零售金额), mintMoneyDigit, , True)
                vsfList.TextMatrix(intRow, mIntCol实收金额) = zlStr.FormatEx(Val(!实收金额), mintMoneyDigit, , True)
                vsfList.TextMatrix(intRow, mIntCol重量) = !重量 & !计算单位
                
                dbl总重量 = dbl总重量 + !重量
                str重量单位 = !计算单位
                vsfList.TextMatrix(intRow, mIntCol频次) = IIf(IsNull(!频次), "", !频次)
                vsfList.TextMatrix(intRow, mIntCol用药目的) = zlStr.NVL(!用药目的)
                If Not IsNull(!单量) Then
                    vsfList.TextMatrix(intRow, mIntCol单量) = zlStr.FormatEx(!单量, mintNumberDigit) & "(" & zlStr.NVL(!计算单位) & ")"
                End If
                vsfList.TextMatrix(intRow, mIntCol用法) = zlStr.NVL(!用法)
                vsfList.TextMatrix(intRow, mIntCol门诊标志) = Val(!门诊标志)
                vsfList.TextMatrix(intRow, mIntCol记录性质) = Val(!记录性质)
                vsfList.TextMatrix(intRow, mIntCol用药理由) = zlStr.NVL(!用药理由)
                vsfList.TextMatrix(intRow, mIntCol相关id) = Val(!相关id)
                vsfList.TextMatrix(intRow, mIntCol开嘱医生) = zlStr.NVL(!开嘱医生)
                vsfList.TextMatrix(intRow, mIntCol频率间隔) = zlStr.NVL(!频率间隔)
                vsfList.TextMatrix(intRow, mIntCol间隔单位) = zlStr.NVL(!间隔单位)
                vsfList.TextMatrix(intRow, mIntCol医嘱标志) = zlStr.NVL(!医嘱标志)
                vsfList.TextMatrix(intRow, mIntCol开始时间) = zlStr.NVL(!开始时间)
                vsfList.TextMatrix(intRow, mIntCol结束时间) = zlStr.NVL(!结束时间)
                vsfList.TextMatrix(intRow, mIntCol频率次数) = zlStr.NVL(!频率次数)
                vsfList.TextMatrix(intRow, mIntCol超量说明) = zlStr.NVL(!超量说明)
                
                If mcondition.intListType = mListType.退药 Then
                    vsfList.TextMatrix(intRow, mIntCol包装) = Val(!包装)
                    If mcondition.bln显示大小单位 = True Then
                        '按大小包装显示数量：分别处理已退数量、准退数量、退药数量
                        '已退数量、准退数量列显示模式为"大包装数量+大包装单位+小包装数量+售价单位"；退药数分两列显示，且只显示数值
                        lng整数量 = Int(!已退数量)
                        If !售价单位 = !单位 Or lng整数量 = !已退数量 Then
                            vsfList.TextMatrix(intRow, mIntCol已退数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                        Else
                            
'                            lng整数量 = Int(!小单位已退数量 / !包装)
                            If lng整数量 = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol已退数) = !小单位已退数 & IIf(IsNull(!售价单位), "", !售价单位)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol已退数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & (!小单位已退数 Mod !包装) & IIf(IsNull(!售价单位), "", !售价单位)
                            End If
                        End If
                        
                        lng整数量 = Int(!准退数)
                        If !售价单位 = !单位 Or lng整数量 = !准退数 Then
                            vsfList.TextMatrix(intRow, mIntCol准退数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                        Else
'                            lng整数量 = Int(!小单位准退数 / !包装)
                            If lng整数量 = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol准退数) = !小单位准退数 & IIf(IsNull(!售价单位), "", !售价单位)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol准退数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & (!小单位准退数 Mod !包装) & IIf(IsNull(!售价单位), "", !售价单位)
                            End If
                        End If
                        
                        lng整数量 = Int(!准退数)
                        If !售价单位 = !单位 Then
                            vsfList.TextMatrix(intRow, mIntCol准退数小) = zlStr.FormatEx(lng整数量, mintNumberDigit)
                        ElseIf lng整数量 = !准退数 Then
                            vsfList.TextMatrix(intRow, mIntCol准退数大) = zlStr.FormatEx(lng整数量, mintNumberDigit)
                            vsfList.TextMatrix(intRow, mIntCol准退数小) = zlStr.FormatEx(0, mintNumberDigit)
                        Else
'                            dbl小数量 = (Val(!准退数) - lng整数量) * !包装
                            If lng整数量 = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol准退数小) = zlStr.FormatEx(!小单位准退数, mintNumberDigit)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol准退数大) = zlStr.FormatEx(lng整数量, mintNumberDigit)
                                vsfList.TextMatrix(intRow, mIntCol准退数小) = zlStr.FormatEx((!小单位准退数 Mod !包装), mintNumberDigit)
                            End If
                        End If
                        
                        vsfList.TextMatrix(intRow, mIntCol退药数) = zlStr.FormatEx(!准退数, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol退药数大) = vsfList.TextMatrix(intRow, mIntCol准退数大)
                        vsfList.TextMatrix(intRow, mIntCol退药数小) = vsfList.TextMatrix(intRow, mIntCol准退数小)
                        vsfList.TextMatrix(intRow, mIntCol单位大) = IIf(IsNull(!单位), "", !单位)
                        vsfList.TextMatrix(intRow, mIntCol单位小) = IIf(IsNull(!售价单位), "", !售价单位)
                    Else
                        vsfList.TextMatrix(intRow, mIntCol已退数) = zlStr.FormatEx(!已退数量, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol准退数) = zlStr.FormatEx(!准退数, mintNumberDigit)
                        vsfList.TextMatrix(intRow, mIntCol退药数) = zlStr.FormatEx(!准退数, mintNumberDigit)
                    End If
                
                    vsfList.TextMatrix(intRow, mIntCol实际数量) = !实际数量
                Else
                    If mcondition.bln显示大小单位 = True Then
                        '按大小包装显示数量
                        lng整数量 = Int(!库存数)
                        If !售价单位 = !单位 Or lng整数量 = !库存数 Then
                            vsfList.TextMatrix(intRow, mIntCol库存数) = lng整数量 & IIf(IsNull(!单位), "", !单位)
                        Else
                            lng整数量 = Int(!库存实际数量 / !包装)
                            If lng整数量 = 0 Then
                                vsfList.TextMatrix(intRow, mIntCol库存数) = !库存实际数量 & IIf(IsNull(!售价单位), "", !售价单位)
                            Else
                                vsfList.TextMatrix(intRow, mIntCol库存数) = lng整数量 & IIf(IsNull(!单位), "", !单位) & (!库存实际数量 Mod !包装) & IIf(IsNull(!售价单位), "", !售价单位)
                            End If
                        End If
                    Else
                        vsfList.TextMatrix(intRow, mIntCol库存数) = zlStr.FormatEx(zlStr.NVL(!库存数, 0), mintNumberDigit)
                    End If
                
                    vsfList.TextMatrix(intRow, mIntCol货位) = zlStr.NVL(!库房货位)
                    vsfList.TextMatrix(intRow, mIntCol医生嘱托) = zlStr.NVL(!医生嘱托)
                    vsfList.TextMatrix(intRow, mIntCol医嘱id) = zlStr.NVL(!医嘱id)
                    
                    If Not gobjPass Is Nothing Then
                        vsfList.Cell(flexcpPicture, intRow, mIntCol审查结果, intRow, mIntCol审查结果) = gobjPass.zlPassSetWarnLight_YF(Val(!审查结果))
                        vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol审查结果, intRow, mIntCol审查结果) = flexPicAlignCenterCenter
                    End If
                    
                    vsfList.TextMatrix(intRow, mIntCol警告) = Val(!审查结果)
'                    '以下用于测试PASS
'                    vsfList.Cell(flexcpPicture, intRow, mIntCol审查结果, intRow, mIntCol审查结果) = frmPublic.imgPass.ListImages(Val(!审查结果) + 2).Picture
'                    vsfList.Cell(flexcpPictureAlignment, intRow, mIntCol审查结果, intRow, mIntCol审查结果) = flexPicAlignCenterCenter
                End If
                
                vsfList.TextMatrix(intRow, mIntCol分批) = IIf(IsNull(!分批), 0, !分批)
                vsfList.TextMatrix(intRow, mIntCol新批号) = ""
                vsfList.TextMatrix(intRow, mIntCol新效期) = ""
                vsfList.TextMatrix(intRow, mIntCol新产地) = ""
                vsfList.TextMatrix(intRow, mIntCol备注) = ""
                vsfList.TextMatrix(intRow, mIntCol费别) = IIf(IsNull(!费别), "", !费别)
                
                dbl应收金额 = dbl应收金额 + Val(!零售金额)
                dbl实收金额 = dbl实收金额 + Val(!实收金额)
                
                '对低于库存下限的药品上色
                vsfList.Redraw = flexRDNone
                If !库存下限 = 0 Then
                    vsfList.Cell(flexcpForeColor, intRow, 1, intRow, vsfList.Cols - 1) = mlng紫色
                Else
                    vsfList.Cell(flexcpForeColor, intRow, 1, intRow, vsfList.Cols - 1) = vbBlack
                End If
                            
                '特殊药品粗体显示
                If InStr(";毒性药;麻醉药;精神I类;精神II类;", zlStr.NVL(!毒理分类)) > 0 And zlStr.NVL(!毒理分类) <> "" Then
                    vsfList.Cell(flexcpFontBold, intRow, mIntCol药品名称, intRow, mIntCol药品名称) = True
                End If
            
                .MoveNext
            Loop
        End If
        
        '退药数粗体显示
        If mcondition.intListType = mListType.退药 Then
            If mcondition.bln显示大小单位 = True Then
                vsfList.Cell(flexcpFontBold, 1, mIntCol退药数大, intRow, mIntCol退药数大) = True
                vsfList.Cell(flexcpFontBold, 1, mIntCol退药数小, intRow, mIntCol退药数小) = True
            Else
                vsfList.Cell(flexcpFontBold, 1, mIntCol退药数, intRow, mIntCol退药数) = True
            End If
        End If
        
        '最后空白行显示金额合计
        intRow = intRow + 1
        vsfList.rows = intRow + 1

        If mcondition.int金额显示 = 1 Then
            strSum = "实收金额：" & Format(dbl实收金额, mstrVBMoneyForamt) & "元" & "(" & zlStr.ChineseMoney(dbl实收金额) & ")"
        ElseIf mcondition.int金额显示 = 2 Then
            strSum = "应收金额：" & Format(dbl应收金额, mstrVBMoneyForamt) & "元" & "  实收金额：" & Format(dbl实收金额, mstrVBMoneyForamt) & "元" & "(" & zlStr.ChineseMoney(dbl实收金额) & ")"
        Else
            strSum = "应收金额：" & Format(dbl应收金额, mstrVBMoneyForamt) & "元" & "(" & zlStr.ChineseMoney(dbl应收金额) & ")"
        End If
        
        If mcondition.bln显示重量 And mbln中药处方 Then
            strSum = strSum & "  总重量：" & dbl总重量 & str重量单位
        End If
        
        vsfList.Cell(flexcpText, intRow, mIntCol顺序号, intRow, vsfList.Cols - 1) = strSum
        vsfList.Cell(flexcpAlignment, intRow, mIntCol顺序号, intRow, vsfList.Cols - 1) = flexAlignLeftCenter
        vsfList.Cell(flexcpFontBold, intRow, mIntCol顺序号, intRow, vsfList.Cols - 1) = True
        
        vsfList.MergeCells = flexMergeRestrictRows
        vsfList.MergeRow(vsfList.rows - 1) = True
        
        '中药处方
        picRecInfo_CM.Visible = False
        If txtNo.ListIndex <> -1 Then
            If txtNo.Tag <> "" Then
                If bln中药处方 Then
                    picRecInfo_CM.Visible = True
                    Call 中药处方特别处理(Val(Split(txtNo.Tag, "|")(0)), Val(txtNo.ItemData(txtNo.ListIndex)), Mid(txtNo.Text, 1, 8), Val(Split(txtNo.Tag, "|")(2)), Val(Split(txtNo.Tag, "|")(3)))
                    vsfList.ColWidth(mIntCol重量) = 1200
                Else
                    vsfList.ColWidth(mIntCol重量) = 0
                End If
            End If
        End If
        Call picRecipt_Resize
        
        '重新生成网格
        vsfList.Select 0, 0, vsfList.rows - 1, vsfList.Cols - 1
        vsfList.CellBorder &H9D9D9D, 1, 1, 1, 1, 1, 1
        
        '如果有皮试，处理皮试结果
        If bln皮试 = True Then
            vsfList.ColWidth(mintCol皮试结果) = 800
            vsfList.Select 0, mIntCol药品名称, vsfList.rows - 1, mintCol皮试结果
            vsfList.CellBorder &H9D9D9D, -1, -1, -1, -1, 0, 1
            
            For i = 1 To vsfList.rows - 1
                If vsfList.TextMatrix(i, mintCol皮试结果) = "(+)" Then
                    vsfList.Cell(flexcpForeColor, i, mintCol皮试结果, i, mintCol皮试结果) = vbRed
                ElseIf vsfList.TextMatrix(i, mintCol皮试结果) = "(-)" Then
                    vsfList.Cell(flexcpForeColor, i, mintCol皮试结果, i, mintCol皮试结果) = vbBlue
                Else
                    vsfList.Cell(flexcpForeColor, i, mintCol皮试结果, i, mintCol皮试结果) = &H80000008
                End If
            Next
        Else
            vsfList.ColWidth(mintCol皮试结果) = 0
        End If
        
        vsfList.Row = vsfList.rows - 1
        
        vsfList.Redraw = flexRDDirect
    End With
    
    Form_Resize
    Call picProcess_Resize
    Call InitColSelList(mcondition.intListType)
    RefreshList = True
    

    SetCmdSendPrivs int审查结果
    
    If Me.CmdSend.Caption = "配药确认(&O)" And int排队状态 = 1 Then
        Me.CmdSend.Caption = "取消确认(&C)"
    ElseIf Me.CmdSend.Caption = "取消确认(&C)" And int排队状态 = 0 Then
        Me.CmdSend.Caption = "配药确认(&O)"
    End If
End Function

Private Function 判断是否中药处方(ByVal lngNO药房id As Long, ByVal BillType As Integer, ByVal BillNo As String) As Boolean
    '通过药品id判断是否是中药
    Dim strsql As String
    Dim rs As New ADODB.Recordset
    Dim lng药房ID As Long
    Dim blnMoved As Boolean
    
    On Error GoTo errHandle
    
    lng药房ID = lngNO药房id
    If lngNO药房id = 0 Then lng药房ID = mcondition.lng药房ID
    
    strsql = "Select a.类别 as 类别 From 收费项目目录 a ,药品收发记录 b Where b.药品id=a.Id And b.单据=[2] and b.No=[1] And (b.记录状态=1 Or Mod(b.记录状态,3)=0) and (b.库房ID+0=[3] OR b.库房ID IS NULL) " _
   
    '如果数据转出，则直接从后备表中提取数据
    blnMoved = Sys.IsMovedByNO("药品收发记录", BillNo, " 单据 = ", BillType)
    If blnMoved Then
        gstrSQL = Replace(gstrSQL, "药品收发记录", "H药品收发记录")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(strsql, Me.Caption & "[判断是否中药处方]", BillNo, BillType, lng药房ID)
    
    mbln中药处方 = IIf(rs!类别 = 7, True, False)
    rs.Close
    
    判断是否中药处方 = mbln中药处方
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub 中药处方特别处理(ByVal lngNO药房id As Long, ByVal BillStyle As Integer, ByVal BillNo As String, ByVal int记录性质 As Integer, ByVal int门诊标志 As Integer)
    '中药处方显示原始付数和中药煎法
    Dim rs As New ADODB.Recordset
    Dim lng药房ID As Long
    
    On Error GoTo errHandle
    lng药房ID = lngNO药房id
    If lngNO药房id = 0 Then lng药房ID = mcondition.lng药房ID

    gstrSQL = "Select a.外观,b.付数 From 药品收发记录 a ,门诊费用记录 b Where a.费用id=b.Id " _
        & " And a.单据=[2] And a.No=[1] " _
        & " And (a.记录状态=1 Or Mod(a.记录状态,3)=0) and (a.库房ID+0=[3] OR a.库房ID IS NULL) "
    If int记录性质 = 1 Or (int记录性质 = 2 And (int门诊标志 = 1 Or int门诊标志 = 4)) Then
    Else
        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
    End If
    
    Set rs = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[中药处方特别处理]", BillNo, BillStyle, lng药房ID)
    
    lbl原始付数.Caption = lbl原始付数.Tag & CStr(IIf(IsNull(rs!付数), 1, rs!付数))
    lbl中药煎法.Caption = lbl中药煎法.Tag & IIf(IsNull(rs!外观), "", rs!外观)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function GetRecord(Optional ByRef int可操作 As Integer) As ADODB.Recordset
    int可操作 = mInt可操作
    Set GetRecord = mrsDetail
End Function

Private Sub GetMoneyFormat()
    Dim n As Integer
    Dim strOracleTmp As String
    Dim strVbTmp As String
    
    strOracleTmp = "999999990."
    strVbTmp = "########0."
    For n = 1 To mintMoneyDigit
        strOracleTmp = strOracleTmp & "0"
        strVbTmp = strVbTmp & "0"
    Next
    
    mstrOracleMoneyForamt = strOracleTmp
    mstrVBMoneyForamt = strVbTmp
    
End Sub

Private Sub Chk全退_Click()
    Dim intRow As Integer
    Dim lng整数量 As Long
    Dim dbl小数量 As Double
    
    If mcondition.intListType <> mListType.退药 Then Exit Sub
    
    If Not Chk全退.Enabled Then Exit Sub
    With vsfList
        For intRow = 1 To .rows - 2
            If mcondition.bln显示大小单位 = True Then
                If Chk全退.Value = 1 Then
                    .TextMatrix(intRow, mIntCol退药数大) = .TextMatrix(intRow, mIntCol准退数大)
                    .TextMatrix(intRow, mIntCol退药数小) = .TextMatrix(intRow, mIntCol准退数小)
                    
                    .TextMatrix(intRow, mIntCol退药数) = zlStr.FormatEx(Val(.TextMatrix(intRow, mIntCol实际数量)) / Val(.TextMatrix(intRow, mIntCol包装)), 5)
                Else
                    .TextMatrix(intRow, mIntCol退药数) = ""
                    .TextMatrix(intRow, mIntCol退药数大) = ""
                    .TextMatrix(intRow, mIntCol退药数小) = ""
                End If
            Else
                .TextMatrix(intRow, mIntCol退药数) = IIf(Chk全退.Value = 1, .TextMatrix(intRow, mIntCol准退数), "")
            End If
        Next
        mblnAllBack = (Chk全退.Value = 1)
    End With
End Sub

Private Sub CmdSend_Click()
    Dim blnmsg As Boolean
    
    If (mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发) And mcondition.bln允许核查人和配药人相同 = False Then
        If Me.cbo核查人.Text = Me.Cbo配药人.Text Then
            If MsgBox("当前发药处方的核查人和配药人相同，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                blnmsg = True
            End If
        End If
        
        If InStr(1, Me.Cbo配药人.Text, "-") < 1 And Not blnmsg Then
            If Mid(Me.cbo核查人.Text, InStr(1, Me.cbo核查人.Text, "-") + 1) = Me.Cbo配药人.Text Then
                If MsgBox("当前发药处方的核查人和配药人相同，是否继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If mcondition.intListType = mListType.退药 Then
        If frm药品处方发药New.RecipeWork(mcondition.intListType, False, vsfList) = False Then
            RefreshList mrsDetail
        End If
    Else
        If mcondition.lng药房ID <> Val(Split(txtNo.Tag, "|")(0)) Then
            If CDate(Format(Split(txtNo.Tag, "|")(4), "yyyy-MM-dd")) <> CDate(Format(Sys.Currentdate, "yyyy-MM-dd")) Then
                If MsgBox("        代发非当天单据，会删除汇总数据重新汇总，" & vbCrLf & "如果已经出了报表的可能需要重新出报表，是否继续操作？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
        
        Call FormClear
        
        If frm药品处方发药New.RecipeWork(mcondition.intListType, mblnInput, vsfList, mbln未取药发药) = False Then
            RefreshList mrsDetail
        End If
        
        mbln未取药发药 = False
    End If
    
    If mblnInput = True Then
        txtNo.SetFocus
        Call zlControl.TxtSelAll(txtNo)
        mblnInput = False
    End If
End Sub

Private Sub cmdSendByNoTake_Click()
    mbln未取药发药 = True
    Call CmdSend_Click
End Sub


Private Sub imgDown_Click()
    imgDown.Visible = False
    imgUp.Visible = True
    
    picRecipt_Resize
End Sub
Private Sub imgUp_Click()
    imgDown.Visible = True
    imgUp.Visible = False
    
    picRecipt_Resize
End Sub

Private Sub lblDiag_Click()
    If Me.imgDown.Visible Then
        imgDown_Click
    Else
        imgUp_Click
    End If
End Sub

Private Sub picHscSend_Click()
    If Me.imgDown.Visible Then
        imgDown_Click
    Else
        imgUp_Click
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveListColState
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    On Error Resume Next
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - picRecipt.Top - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - picRecipt.Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList.Cols - 1
            '不在不允许显示列表的列才能加入列选择列表
            If IsInString(mstrUnallowShow, vsfList.ColKey(i), ";") = False Then
                If (mcondition.bln显示原产地 And vsfList.ColKey(i) = "原产地") Or vsfList.ColKey(i) <> "原产地" Then
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, 1) = vsfList.TextMatrix(0, i)
                    .RowData(.rows - 1) = i
                End If
                
                '列宽为空或者隐藏的列设置为不勾选
                If Not (vsfList.ColWidth(i) = 0 Or vsfList.ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                '指定的列设置为不能设置隐藏
                If IsInString(mstrUnallowSetColHide, vsfList.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub
Private Sub Form_Load()
    mblnAllowClick = True
    
    mstrPrivs = gstrprivs
    
    mlngMode = glngModul
    
    Lbl药房.Caption = ""
    
    '取金额位数
    mintMoneyDigit = Val(zldatabase.GetPara("费用金额保留位数", glngSys, 0))
    
    '设置金额格式
    Call GetMoneyFormat
    
    Call Load医生
    
    Call SetParams

    Call InitComandBars
    
    Call FormClear
    
    picRecInfo_CM.BackColor = &H8000000F
    picProcess.BackColor = &H8000000F
    picRecInfo.BackColor = &H8000000F
End Sub

Private Sub Load配药人(ByVal lng药房ID As Long)
    '配药人
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select 简码||'-'||姓名 As 姓名,姓名 As 名称 From 人员表  Where ID in " & _
             " (Select Distinct 人员ID From 人员性质说明 Where 人员性质='药房发药人' " & _
             " And 人员ID IN (Select 人员ID From 部门人员 Where 部门ID=[1]))" & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取配药人", lng药房ID)
    
    With rsData
        Me.Cbo配药人.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            Cbo配药人.AddItem !姓名
            
            If InStr(1, mstrDosUser, "-") > 0 Then
                mstrDosUser = Mid(mstrDosUser, InStr(1, mstrDosUser, "-") + 1)
            End If
            
            If mstrDosUser = !名称 Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        Cbo配药人.Enabled = Not Cbo配药人.ListCount = 0

        If intIndex <> -1 Then Cbo配药人.ListIndex = intIndex
        mstrDosUser = Me.Cbo配药人.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load核查人(ByVal lng药房ID As Long)
    '核查人
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = " Select 简码||'-'||姓名 As 姓名,姓名 As 名称 From 人员表  Where ID in " & _
             " (Select Distinct 人员ID From 人员性质说明 Where 人员性质='药房发药人' " & _
             " And 人员ID IN (Select 人员ID From 部门人员 Where 部门ID=[1]))" & _
             " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
             
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取审核人", lng药房ID)
    
    With rsData
        Me.cbo核查人.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo核查人.AddItem !姓名
            
            If InStr(1, mstr核查人, "-") > 0 Then
                mstr核查人 = Mid(mstr核查人, InStr(1, mstr核查人, "-") + 1)
            End If
            
            If mstr核查人 = !名称 Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        cbo核查人.Enabled = Not cbo核查人.ListCount = 0

        If intIndex <> -1 Then cbo核查人.ListIndex = intIndex
        mstr核查人 = Me.cbo核查人.Text
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If cbsMain.count = 1 Then
        picRecipt.Move 0, 0, Me.Width, Me.Height
    End If
End Sub


Private Sub picProcess_Resize()
    On Error Resume Next
    
    With CmdSend
        .Left = picProcess.Width - .Width - 100
    End With
    
    With cmdSendByNoTake
        .Left = CmdSend.Left - .Width - 100
    End With
    
    With Chk全退
        .Left = IIf(cmdSendByNoTake.Visible, cmdSendByNoTake.Left, CmdSend.Left) - .Width - 100
    End With
End Sub


Private Sub picRecInfo_Resize()
    Dim objTmp As Object
    
    On Error Resume Next

    With txt诊断内容
        .Width = picRecInfo.Width - .Left - 50
    End With
    
    With lblNotice
        .Top = txt诊断内容.Top + txt诊断内容.Height + 200
    End With

    
    With picRecInfo
        .Height = lblNotice.Top + lblNotice.Height + 100
    End With
End Sub


Private Sub picRecipt_Resize()
    On Error Resume Next
    
    With picRecInfo
        .Top = 0
        .Left = 0
        .Width = picRecipt.Width
    End With
    
    With picProcess
        .Top = picRecipt.Height - .Height
        .Left = 0
        .Width = picRecipt.Width
    End With
        
    With picRecInfo_CM
        If .Visible Then
            .Top = picProcess.Top - .Height - 100
            .Left = 0
            .Width = picRecipt.Width
        End If
    End With
    
    With vsfList
        .Top = picRecInfo.Top + picRecInfo.Height + 100
        .Left = 0
        .Width = picRecipt.Width
        .Height = IIf(picRecInfo_CM.Visible, picRecInfo_CM.Top, picProcess.Top) - picRecInfo.Height - 100 - IIf(Me.picHscSend.Visible, Me.picHscSend.Height, 0) - IIf(imgDown.Visible, Me.txt用药理由.Height, 0)
    End With
    
    With picHscSend
        .Top = Me.vsfList.Top + Me.vsfList.Height - IIf(Me.picHscSend.Visible, 0, Me.picHscSend.Height)
        .Left = Me.vsfList.Left
        .Width = Me.vsfList.Width - 20
    End With
    
    lblDiag.Left = (picHscSend.Width - lblDiag.Width) / 2
    
    If imgDown.Visible Then
        With Me.txt用药理由
            .Visible = True
            .Top = Me.picHscSend.Top + Me.picHscSend.Height
            .Left = Me.picHscSend.Left
            .Width = Me.picHscSend.Width
        End With
    Else
        txt用药理由.Visible = False
    End If
    
    With fraColSel
        .Left = vsfList.Left + vsfList.ColWidth(0) - .Width - 30
        .Top = vsfList.Top + (vsfList.RowHeight(0) - .Height) / 2 + 30
        .ZOrder
    End With
End Sub

Private Sub InitComandBars()
    Dim cbrControl As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    
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
        .LargeIcons = False
'        .SetIconSize False, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.ActiveMenuBar.Visible = False
    Me.cbsMain.AddImageList Me.imgCheck
End Sub

Private Sub SetComandBars(ByVal intListType As Integer)
    Dim cbrControl As CommandBarControl
    Dim cbrControlSub As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    Dim objMenu As CommandBarPopup
        
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    If (intListType <> mListType.待发药 And intListType <> mListType.超时未发) Then Exit Sub
    
    Select Case intListType
        Case mListType.待发药, mListType.超时未发
            '设置工具栏菜单
            If Not gobjPass Is Nothing And IsInString(gstrprivs, "合理用药监测", ";") Then
'                Set objCmdBar = cbsMain.Add("条件", xtpBarTop)
'                objCmdBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
'                objCmdBar.ModifyStyle XTP_CBRS_GRIPPER, 0
'                objCmdBar.ContextMenuPresent = False
'
'                Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "过敏史/病生状态")
'                cbrControl.BeginGroup = True
'                cbrControl.ToolTipText = "提示：显示过敏史/病生状态"
'                cbrControl.Style = xtpButtonIconAndCaption
'                cbrControl.IconId = 3

                If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, conMenu_Tool_ShowPlug, 3)
            End If
            
'            设置弹出菜单，PASS
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS（&P)", 1, False)
            objMenu.Id = mconMenu_PASS
'            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, mconMenu_PASS, 1)
'            With objMenu.CommandBar.Controls
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 0, "药物临床信息参考(&C)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 1, "药品说明书(&D)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 2, "中国药典(&N)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 3, "病人用药教育(&S)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 4, "检验值(&T)")
'
'                Set cbrControl = .Add(xtpControlPopup, mconMenu_PASS_Item + 5, "专项信息(&P)")
'                cbrControl.BeginGroup = True
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 0, "药物-药物相互作用(&D)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 1, "药物-食物相互作用(&F)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 2, "国内注射剂配伍(&M)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 3, "国外注射剂配伍(&T)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 4, "禁忌症(&C)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 5, "副作用(&S)", -1, False)
'
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 6, "老年人用药(&G)", -1, False)
'                cbrControlSub.BeginGroup = True
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 7, "儿童用药(&P)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 8, "妊娠期用药(&E)", -1, False)
'                Set cbrControlSub = cbrControl.CommandBar.Controls.Add(xtpControlButton, mconMenu_PASS_Spec + 9, "哺乳期用药(&L)", -1, False)
'
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 6, "医药信息中心(&I)")
'                cbrControl.BeginGroup = True
'
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 7, "药品配对信息(&M)")
'                cbrControl.BeginGroup = True
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 8, "给药途径配对信息(&R)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 9, "医院药品信息(&F)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 10, "警告(&W)")
'                Set cbrControl = .Add(xtpControlButton, mconMenu_PASS_Item + 11, "审查(&V)")
'            End With

    End Select
End Sub


Public Sub SetFontSize(ByVal intFont As Integer)
    With vsfList
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 100
        .RowHeightMax = TextHeight("刘") + 100
        .Refresh
    End With
End Sub

Private Function Check是否存在医嘱(ByVal Int单据 As Integer, ByVal strNo As String) As Boolean
    '判断是住院还是门诊病人，有无医嘱记录
    Dim rsData As ADODB.Recordset
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[1] And A.no=[2] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取医嘱记录", Int单据, strNo)
    
    Check是否存在医嘱 = Not rsData.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub TxtNo_Click()
    GetRecipeByNO
End Sub

Private Sub TxtNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strNo As String
    Dim rsData As ADODB.Recordset
    Dim rstemp As ADODB.Recordset
    Dim strTmp As String
    Dim ArrTmp
    Dim blnExit As Boolean
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txtNo.Text) = "" Then Exit Sub
    
    If Len(txtNo.Text) > 8 Or InStr(1, txtNo.Text, "-") > 0 Then
        If vsfList.rows > 1 Then
            If vsfList.TextMatrix(1, mIntCol药品名称) <> "" Then
                If CmdSend.Enabled = True Then
                    CmdSend.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If
    
    strNo = GetFullNO(Trim(txtNo.Text), 13)
    
    Set rsData = frm药品处方发药New.GetRecipeByNO(strNo)
    
    DoEvents
    If rsData Is Nothing Then
        If mcondition.intListType <> mListType.退药 Then
            MsgBox zlStr.FormatString("该处方“[1]”不存在（可能已经被其他人删除），请重新输入！", strNo), vbInformation, gstrSysName
        Else
            MsgBox zlStr.FormatString("未找到指定处方“[1]”或该处方已经退药，请重新输入！", strNo), vbInformation, gstrSysName
        End If
        
        DoEvents
        txtNo.Text = ""
        txtNo.SetFocus
        Exit Sub
    ElseIf rsData.RecordCount = 0 Then
        '[退药]标签中输入的单据号如果在[待配药]]或[待发药]中存在，则提示
'        Set rsData = frm药品处方发药New.GetRecipeByNO(strNo, 1)
'        rsData.Filter = "审核日期 = Null"
        Set rstemp = frm药品处方发药New.GetRecipeByNO(strNo, 1)
        If rstemp Is Nothing Then
            MsgBox zlStr.FormatString("该处方“[1]”不存在（可能已经被其他人删除），请重新输入！", strNo), vbInformation, gstrSysName
            Exit Sub
        End If
        If rstemp.RecordCount = 0 Then
            MsgBox zlStr.FormatString("该处方“[1]”不存在（可能已经被其他人删除），请重新输入！", strNo), vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If mcondition.intListType <> mListType.退药 Then
            If rsData.EOF Then
                MsgBox zlStr.FormatString("该处方“[1]”不存在（可能已经被其他人删除），请重新输入！", strNo), vbInformation, gstrSysName
                blnExit = True
            Else
                rsData.Sort = "审核日期,记录状态 desc"   '取最近一次的数据
                If mcondition.intListType = mListType.待配药 Then
                    rsData.Filter = "审核日期 = Null and 配药日期 = null"
                    If rsData.RecordCount = 0 Then
'                        If IsNull(rsData!配药日期) = False Then
'                            MsgBox zlStr.FormatString("该处方“[1]”已经配药确认，请重新输入！", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        End If

                        '查询[待发药]或[退药]中是否存在该单据。
                        rsData.Filter = "审核日期 = Null And 配药日期 <> null"
                        If rsData.RecordCount > 0 Then
                            MsgBox zlStr.FormatString("该处方“[1]”已经配药确认，请重新输入！", strNo), vbInformation, gstrSysName
                            blnExit = True
                        Else
                            rsData.Filter = "审核日期 <> Null"
                            If rsData.RecordCount > 0 Then
                                MsgBox zlStr.FormatString("该处方“[1]”已经发药，请重新输入！", strNo), vbInformation, gstrSysName
                                blnExit = True
                            End If
                        End If
                        
                    End If
                ElseIf mcondition.intListType = mListType.待发药 Then
                    If mcondition.bln是否需要配药过程 Then
                        rsData.Filter = "审核日期 = Null And 配药日期 <> null"
                    Else
                        rsData.Filter = "审核日期 = Null"
                    End If
                    If rsData.RecordCount = 0 Then
'                        If IsNull(rsData!审核日期) = False Then
'                            MsgBox zlStr.FormatString("该处方“[1]”已经发药，请重新输入！", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        ElseIf IsNull(rsData!配药日期) And mcondition.bln是否需要配药过程 Then
'                            MsgBox zlStr.FormatString("该处方“[1]”未配药完成，请重新输入！", strNo), vbInformation, gstrSysName
'                            blnExit = True
'                        End If

                        '查询[待配药]或[退药]中是否存在该单据。
                        rsData.Filter = "审核日期 = Null And 配药日期 = null"
                        If rsData.RecordCount > 0 And mcondition.bln是否需要配药过程 Then
                            MsgBox zlStr.FormatString("该处方“[1]”未配药完成，请重新输入！", strNo), vbInformation, gstrSysName
                            blnExit = True
                        Else
                            rsData.Filter = "审核日期 <> Null"
                            If rsData.RecordCount > 0 Then
                                MsgBox zlStr.FormatString("该处方“[1]”已经发药，请重新输入！", strNo), vbInformation, gstrSysName
                                blnExit = True
                            End If
                        End If
                        
                    End If
                End If
            End If
            
            If blnExit Then
                DoEvents    '防止焦点定位txtNo失效（原因：有嵌入窗体的情况）
                txtNo.SelStart = 1: txtNo.SelLength = Len(txtNo.Text)
                txtNo.SetFocus
                Exit Sub
            End If
            
            If mcondition.intListType = mListType.待配药 Then
                '过滤出未“发药”的记录
                rsData.Filter = "审核日期 = Null"
            ElseIf mcondition.intListType = mListType.待发药 Then
                '过滤出未“发药”的记录
                 rsData.Filter = "审核日期 = Null"
            End If
        End If
    End If
    
    If rsData.RecordCount > 1 Then
        With vsfNoList
            .rows = 2
            .Redraw = flexRDNone
            Do While Not rsData.EOF
                .TextMatrix(.rows - 1, .ColIndex("药房")) = rsData!药房
                .TextMatrix(.rows - 1, .ColIndex("类型")) = rsData!类型
                .TextMatrix(.rows - 1, .ColIndex("NO")) = rsData!NO
                .TextMatrix(.rows - 1, .ColIndex("姓名")) = IIf(IsNull(rsData!姓名), "", rsData!姓名)
                .TextMatrix(.rows - 1, .ColIndex("库房ID")) = rsData!药房ID
                .TextMatrix(.rows - 1, .ColIndex("单据")) = rsData!单据
                .TextMatrix(.rows - 1, .ColIndex("记录性质")) = rsData!记录性质
                .TextMatrix(.rows - 1, .ColIndex("门诊标志")) = rsData!门诊标志
                .TextMatrix(.rows - 1, .ColIndex("填制日期")) = rsData!填制日期
                .TextMatrix(.rows - 1, .ColIndex("记录状态")) = rsData!记录状态
                .rows = .rows + 1
                rsData.MoveNext
            Loop
            .Redraw = flexRDDirect
            .Top = txtNo.Top + txtNo.Height + 50
            .Width = 4500
            .Height = 1300
            .Left = txtNo.Left - (.Width - txtNo.Width)
            .Visible = True
            .ZOrder 0
            DoEvents
            .SetFocus
        End With
    ElseIf rsData.RecordCount = 1 Then
        If CheckAndProcessBill(mcondition.intListType, rsData!单据, rsData!NO, rsData!药房) = False Then
            DoEvents
            txtNo.Clear
            txtNo.Text = ""
            txtNo.SetFocus
            Exit Sub
        End If
        
        txtNo.Clear
        
        Do While Not rsData.EOF
            txtNo.AddItem rsData!NO & "--" & rsData!姓名
            txtNo.ItemData(txtNo.NewIndex) = rsData!单据
            txtNo.Tag = rsData!药房ID & "|" & rsData!药房 & "|" & rsData!记录性质 & "|" & rsData!门诊标志 & "|" & rsData!填制日期 & "|" & rsData!记录状态
            Lbl药房.Caption = rsData!药房
            rsData.MoveNext
        Loop
        
        If txtNo.ListCount = 0 Then Exit Sub
        
        txtNo.ListIndex = 0
    End If
End Sub

Private Function CheckAndProcessBill(ByVal intType As Integer, ByVal Int单据 As Integer, ByVal strNo As String, ByVal str药房 As String) As Boolean
    '检查单据
    Dim rsTmp As ADODB.Recordset
    
    '检查是否具有已标志停发的药品记录
    '适应环节：配药，已配药，待发药
    '后续处理：是否具有恢复发药权限，有则恢复标志
    On Error GoTo errHandle
    If intType = mListType.待配药 Or intType = mListType.已配药 Or intType = mListType.待发药 Or intType = mListType.超时未发 Then
        gstrSQL = "Select 1 From 药品收发记录 Where 单据 = [1] And NO = [2] And Nvl(发药方式, 0) = -1 And Rownum = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否有已标志停发药的药品记录", Int单据, strNo)
        
        If Not rsTmp.EOF Then
            If zlStr.IsHavePrivs(mstrPrivs, "恢复发药") = True And (intType = mListType.待配药 Or intType = mListType.已配药 Or intType = mListType.待发药 Or intType = mListType.超时未发) Then
                If MsgBox("[" & str药房 & "]处方" & strNo & "存在已标记为不再发药的药品，是否取消标志继续发药？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '检查若是取消标志后当前库房的可用库存是否足够
                    If CheckUnSendStock(Int单据, strNo) = False Then
                        Exit Function
                    End If
                    
                    '取消发药标志
                    CancelUnCheck Int单据, strNo
                Else
                    Exit Function
                End If
            Else
                MsgBox "[" & str药房 & "]处方" & strNo & "存在已标记为不再发药的药品，你没有相应的权限，不能继续发药！", vbInformation, gstrSysName
                txtNo.Text = ""
                Exit Function
            End If
        End If
    End If
    
    '检查是否已配药
    '适应环节：待配药
    If intType = mListType.待配药 Then
        gstrSQL = "Select 1 From 药品收发记录 Where 单据 = [1] And NO = [2] And (记录状态=1 or Mod(记录状态,3)=1) And 配药日期 Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否未配药", Int单据, strNo)
            
        If rsTmp.EOF Then
            MsgBox "[" & str药房 & "]处方" & strNo & "已经配药了，请重新输入！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '检查是否已配药
    '适应环节：已配药，待发药（是否需要配药环节参数）
    If intType = mListType.已配药 Or ((intType = mListType.待发药 Or intType = mListType.超时未发) And mcondition.bln是否需要配药过程 = True) Then
        gstrSQL = "Select 1 From 药品收发记录 Where 单据 = [1] And NO = [2] And (记录状态=1 or Mod(记录状态,3)=1) And 配药日期 Is Not Null And 审核日期 Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否已配药", Int单据, strNo)
            
        If rsTmp.EOF Then
            If intType = mListType.已配药 Then
                MsgBox "[" & str药房 & "]处方" & strNo & "还未配药或者已经取消配药了，请重新输入！", vbInformation, gstrSysName
                Exit Function
            ElseIf (intType = mListType.待发药 Or intType = mListType.超时未发) And mcondition.bln是否需要配药过程 = True Then
                MsgBox "[" & str药房 & "]处方" & strNo & "还未配药，请重新输入！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    '检查是否已发药
    '检查环节：配药，已配药，待发药
    If intType = mListType.待配药 Or intType = mListType.已配药 Or intType = mListType.待发药 Or intType = mListType.超时未发 Then
        gstrSQL = "Select 1 From 药品收发记录 Where 单据 = [1] And NO = [2] And (记录状态=1 or Mod(记录状态,3)=1) And 审核日期 Is Null And Rownum = 1"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否未配药", Int单据, strNo)
        
        If rsTmp.EOF Then
            MsgBox "[" & str药房 & "]处方" & strNo & "已经发药，请重新输入！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '检查是否已经部分退药
    '检查环节：待发药（代发其他药房处方的前提下）
    If (intType = mListType.待发药 Or intType = mListType.超时未发) And zlStr.IsHavePrivs(mstrPrivs, "发其它药房的处方") = True Then
        gstrSQL = " Select A.NO, A.单据, A.药品id, A.序号, Sum(Nvl(A.付数, 1) * A.实际数量) 已发数量 " & _
            " From 药品收发记录 A " & _
            " Where A.审核人 Is Not Null And A.记录状态 <> 1 And A.NO = [2] And A.库房id <> [3] And A.单据 = [1] " & _
            " Group By A.NO, A.单据, A.药品id, A.序号 Having Sum(Nvl(A.付数, 1) * A.实际数量) > 0"
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSQL, "检查是否未配药", Int单据, strNo, mcondition.lng药房ID)
        
        If Not rsTmp.EOF Then
            MsgBox "[" & str药房 & "]处方" & strNo & "已经部分退药，不能代发药，请重新输入！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckAndProcessBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub CancelUnCheck(ByVal Int单据 As Integer, ByVal strNo As String)
    '取消不再发药标志，按指定的单据来执行
    Dim rsData As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim arrSql As Variant
    
    arrSql = Array()
    On Error GoTo errHandle
    gstrSQL = "Select ID From 药品收发记录 Where 单据 = [1] And NO = [2] And Nvl(发药方式, 0) = -1 And 审核日期 Is Null "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "取消不再发药标志", Int单据, strNo)
    
    If rsData.EOF Then Exit Sub
    
    Do While Not rsData.EOF
        gstrSQL = "Zl_不发药处方标记_Unchecked(" & Val(rsData!Id) & ",0)"
        
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        rsData.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnTrans = True
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), Me.Caption & "-取消标记")
        Next
    gcnOracle.CommitTrans
    blnTrans = False
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans

    MsgBox "提示：更新失败。"
    Call SaveErrLog
End Sub
Private Function CheckUnSendStock(ByVal Int单据 As Integer, ByVal strNo As String) As Boolean
    '取消已标记为不再发药的标志，应用于录入单据号方式，需要检查当前库房的可用数量是否足够
    '库存检查：0-不检查;1-检查,不足提醒;2-检查,不足禁止
    Dim rsData As ADODB.Recordset
    Dim strMsg As String
    
    If mcondition.int库存检查 = 0 Then
        CheckUnSendStock = True
        Exit Function
    End If
    On Error GoTo errHandle
    gstrSQL = "Select '[' || C.编码 || ']' || C.名称 || ' ' || C.规格 As 名称, A.实际数量 * A.付数, Nvl(B.可用数量, 0) As 可用数量 " & _
        " From 药品收发记录 A, 收费项目目录 C, " & _
        " (Select 药品id, Nvl(批次, 0) As 批次, Nvl(可用数量, 0) As 可用数量 " & _
        " From 药品库存 " & _
        " Where 性质 = 1 And 库房id + 0 = [3]) B " & _
        " Where A.药品id = C.ID And A.单据 = [1] And A.NO = [2] And Nvl(A.发药方式, 0) = -1 And A.药品id = B.药品id(+) " & _
        " And Nvl(A.批次, 0) = B.批次(+) And A.实际数量 * A.付数 > Nvl(B.可用数量, 0) "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "检查可用库存", Int单据, strNo, mcondition.lng药房ID)

    With rsData
        Do While Not .EOF
            strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & !名称
            .MoveNext
        Loop
    End With
    
    If strMsg <> "" Then
        If mcondition.int库存检查 = 1 Then
            strMsg = "以下药品在恢复发药标记后，当前库房的可用数量不足，是否继续发药？" & vbCrLf & strMsg
            
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            
        ElseIf mcondition.int库存检查 = 2 Then
            strMsg = "以下药品在恢复发药标记后，当前库房的可用数量不足，不能发药！" & vbCrLf & strMsg
            
            MsgBox strMsg, vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    CheckUnSendStock = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt诊断内容_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call SetTip(txt诊断内容, txt诊断内容.Tag)
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList.ColWidth(lngCol) = vsfList.ColData(lngCol)
            vsfList.ColHidden(lngCol) = False
        Else
            vsfList.ColWidth(lngCol) = 0
            vsfList.ColHidden(lngCol) = True
        End If
    End If
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

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim blnUnValid As Boolean
    Dim dblCount As Double
    Dim dblSumCount As Double
    Dim rstemp As New ADODB.Recordset
    Dim dbl退药数 As Double
    
    On Error GoTo errHandle
    With vsfList
        If mcondition.bln显示大小单位 = True Then
            If Col <> mIntCol退药数大 And Col <> mIntCol退药数小 Then Exit Sub
        Else
            If Col <> mIntCol退药数 Then Exit Sub
        End If
        
        blnUnValid = False
        dbl退药数 = Val(.TextMatrix(Row, Col))
        
        If mcondition.bln显示大小单位 = True Then
            If Col = mIntCol退药数大 Then
                dblSumCount = dbl退药数 * Val(.TextMatrix(Row, mIntCol包装)) + Val(.TextMatrix(Row, mIntCol退药数小))
            Else
                dblSumCount = Val(.TextMatrix(Row, mIntCol退药数大)) * Val(.TextMatrix(Row, mIntCol包装)) + dbl退药数
            End If
        Else
            dblSumCount = dbl退药数
        End If
        blnUnValid = Not ((Abs(dblSumCount) <= Abs(Val(.Tag))) And ((Val(dblSumCount) >= 0 And Val(.Tag) >= 0) Or (Val(dblSumCount) <= 0 And Val(.Tag) <= 0)))
        
        If blnUnValid Then
            If mcondition.bln显示大小单位 = True Then
                If Col = mIntCol退药数大 Then
                    .TextMatrix(Row, Col) = Val(.TextMatrix(Row, mIntCol准退数大))
                Else
                    .TextMatrix(Row, Col) = Val(.TextMatrix(Row, mIntCol准退数小))
                End If
            Else
                .TextMatrix(Row, Col) = Val(.Tag)
            End If
        End If
        
        '先检查是否是医嘱产生的药品记录
        '如果不是则不管
        '如果是，检查系统参数是否允许未作废医嘱退药，如果不允许，退药数为零
        '如果允许则不管
        dblCount = Val(.TextMatrix(Row, Col))
        If dblCount <> 0 And mcondition.bln医嘱作废 = False Then
            gstrSQL = "select 扣率 From 药品收发记录 Where ID=[1] "
            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是临嘱]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
            
            If (rstemp!扣率 Like "1*") Then       '临嘱
                gstrSQL = "select B.执行状态 from 病人医嘱记录 A,病人医嘱发送 B,门诊费用记录 C where A.相关id=B.医嘱ID and A.id=C.医嘱序号 and  C.ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                If Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 4)) Then
                Else
                    gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                End If
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查医嘱的给药途径是否已经执行]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                
                If Not rstemp.EOF Then
                    If rstemp!执行状态 = 0 Then
                        gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 门诊费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                        If Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 4)) Then
                        Else
                            gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                        End If
                        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                        
                        If Not rstemp.EOF Then
                            If (rstemp!门诊标志 = 1 Or rstemp!门诊标志 = 4) And rstemp!医嘱序号 <> 0 Then
                                gstrSQL = "Select Nvl(主页id, 0) As 主页id, 挂号单, decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where 病人来源=1  And ID=[1]"
                                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rstemp!医嘱序号))
                                
                                If Not rstemp.EOF Then
                                    If rstemp!主页id > 0 And IsNull(rstemp!挂号单) Then
                                        '填了主页ID，但没有挂号单的不受医嘱是否作废的限制
                                    Else
                                        If rstemp!作废 = 0 Then
                                            dblCount = 0
                                            MsgBox "该笔医嘱还未作废，不能退药！", vbInformation, gstrSysName
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    gstrSQL = "Select Nvl(医嘱序号,0) 医嘱序号,Nvl(门诊标志,1) 门诊标志 From 门诊费用记录 Where ID=(Select 费用ID From 药品收发记录 Where ID=[1])"
                    If Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 1 Or (Val(vsfList.TextMatrix(vsfList.Row, mIntCol记录性质)) = 2 And (Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 1 Or Val(vsfList.TextMatrix(vsfList.Row, mIntCol门诊标志)) = 4)) Then
                    Else
                        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    End If
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查是否是医嘱]", Val(vsfList.TextMatrix(vsfList.Row, mIntColId)))
                    
                    If Not rstemp.EOF Then
                        If (rstemp!门诊标志 = 1 Or rstemp!门诊标志 = 4) And rstemp!医嘱序号 <> 0 Then
                            gstrSQL = "Select Nvl(主页id, 0) As 主页id, 挂号单, decode(医嘱状态,4,1,0) 作废 From 病人医嘱记录 Where 病人来源=1  And ID=[1]"
                            Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[判断该医嘱是否作废]", CLng(rstemp!医嘱序号))
                            
                            If Not rstemp.EOF Then
                                If rstemp!主页id > 0 And IsNull(rstemp!挂号单) Then
                                    '填了主页ID，但没有挂号单的不受医嘱是否作废的限制
                                Else
                                    If rstemp!作废 = 0 Then
                                        dblCount = 0
                                        MsgBox "该笔医嘱还未作废，不能退药！", vbInformation, gstrSysName
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        .TextMatrix(Row, Col) = zlStr.FormatEx(dblCount, 5)
        
        If mcondition.bln显示大小单位 = True Then
            If Col = mIntCol退药数大 Then
                .TextMatrix(Row, mIntCol退药数大) = zlStr.FormatEx(dblCount, 5)
            Else
                .TextMatrix(Row, mIntCol退药数小) = zlStr.FormatEx(dblCount, 5)
            End If
            .TextMatrix(Row, mIntCol退药数) = zlStr.FormatEx(dblSumCount, 5) / Val(.TextMatrix(Row, mIntCol包装))
            
            If Val(.TextMatrix(Row, mIntCol退药数)) <> Val(.TextMatrix(Row, mIntCol实际数量)) / Val(.TextMatrix(Row, mIntCol包装)) Then
                mblnAllBack = False
            End If
        Else
            .TextMatrix(Row, mIntCol退药数) = zlStr.FormatEx(dblCount, 5)
            
            If Val(.TextMatrix(Row, mIntCol退药数)) <> Val(.TextMatrix(Row, mIntCol准退数)) Then
                mblnAllBack = False
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '重设列选择列表
    Call InitColSelList(mcondition.intListType)
    
    '重设列顺序号
    For i = 0 To vsfList.Cols - 1
        Call SetColumnValue(vsfList.TextMatrix(0, i), i)
    Next
End Sub


Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "用药目的"
            mIntCol用药目的 = intValue
        Case "库房货位"
            mIntCol货位 = intValue
        Case "序号"
            mintcol序号 = intValue
        Case "审查结果"
            mIntCol审查结果 = intValue
        Case "顺序号"
            mIntCol顺序号 = intValue
        Case "药品名称"
            mIntCol药品名称 = intValue
        Case "其它名"
            mIntCol其它名 = intValue
        Case "英文名"
            mIntCol英文名 = intValue
        Case "配方名称"
            mIntCol配方名称 = intValue
        Case "规格"
            mintcol规格 = intValue
        Case "批号"
            mintcol批号 = intValue
        Case "单位"
            mintcol单位 = intValue
        Case "单价"
            mIntCol单价 = intValue
        Case "付数"
            mIntCol付数 = intValue
        Case "数量"
            mintcol数量 = intValue
        Case "金额", "应收金额"
            mIntCol金额 = intValue
        Case "实收金额"
            mIntCol实收金额 = intValue
        Case "重量"
            mIntCol重量 = intValue
        Case "单量"
            mIntCol单量 = intValue
        Case "用法"
            mIntCol用法 = intValue
        Case "频次"
            mIntCol频次 = intValue
        Case "医生嘱托"
            mIntCol医生嘱托 = intValue
        Case "费别"
            mIntCol费别 = intValue
        Case "库存数"
            mIntCol库存数 = intValue
        Case "库房货位"
            mIntCol货位 = intValue
        Case "已退数"
            mIntCol已退数 = intValue
        Case "准退数"
            mIntCol准退数 = intValue
        Case "准退数大"
            mIntCol准退数大 = intValue
        Case "准退数小"
            mIntCol准退数小 = intValue
        Case "退药数"
            mIntCol退药数 = intValue
        Case "退药数(大包装)"
            mIntCol退药数大 = intValue
        Case "单位(大)"
            mIntCol单位大 = intValue
        Case "退药数(小包装)"
            mIntCol退药数小 = intValue
        Case "单位(小)"
            mIntCol单位小 = intValue
        Case "备注"
            mIntCol备注 = intValue
        Case "退药数(大包装)"
            mIntCol退药数大 = intValue
        Case "皮试结果"
            mintCol皮试结果 = intValue
        Case "效期"
            mintcol效期 = intValue
        Case "费别"
            mIntCol费别 = intValue
        Case "超量说明"
            mIntCol超量说明 = intValue
        Case "生产商"
            mintcol生产商 = intValue
        Case "原产地"
            mintcol原产地 = intValue
    End Select
                   
End Sub

Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = mintCol皮试结果 Then
        With vsfList
            If .ColWidth(mintCol皮试结果) > 400 Then
                .ColWidth(mIntCol药品名称) = .ColWidth(mIntCol药品名称) + (.ColWidth(mintCol皮试结果) - 400)
                .ColWidth(mintCol皮试结果) = 400
            Else
                .ColWidth(mIntCol药品名称) = .ColWidth(mIntCol药品名称) - (400 - .ColWidth(mintCol皮试结果))
                .ColWidth(mintCol皮试结果) = 400
            End If
        End With
    End If
End Sub

Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '设置不能移动的列
    Select Case mcondition.intListType
        Case mListType.待配药, mListType.已配药, mListType.待发药, mListType.超时未发, mListType.退药, mListType.配药确认
            If Col = mIntCol药品名称 Then
                Position = mIntCol药品名称
            End If
            
            If Col = mintCol皮试结果 Then
                Position = mintCol皮试结果
            End If
            
            If Col = mIntCol顺序号 Then
                Position = mIntCol顺序号
            End If
        
            If Col = mIntCol审查结果 Then
                Position = mIntCol审查结果
            End If
            
            If (Col <> mIntCol药品名称 And Position = mIntCol药品名称) Or (Col <> mintCol皮试结果 And Position = mintCol皮试结果) Or (Col <> mIntCol顺序号 And Position = mIntCol顺序号) Or (Col <> mIntCol审查结果 And Position = mIntCol审查结果) Then
                Position = Col
            End If
    End Select
End Sub

Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    Select Case mcondition.intListType
        Case mListType.待配药, mListType.已配药, mListType.待发药, mListType.超时未发, mListType.退药, mListType.配药确认
            If Col = mIntCol当前行 Or Col = mIntCol顺序号 Or Col = mIntCol审查结果 Then Cancel = True
    End Select
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
'    With vsfList
'        If Row = 0 Then Exit Sub
'        If mcondition.bln显示大小单位 = True Then
'            If Col <> mIntCol退药数大 And Col <> mIntCol退药数小 Then Exit Sub
'            If Val(.TextMatrix(Row, Col)) > 0 Then
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HBFC5FF
'            Else
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = .BackColor
'            End If
'        Else
'            If Col <> mIntCol退药数 Then Exit Sub
'            If Val(.TextMatrix(Row, Col)) > 0 Then
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = &HBFC5FF
'            Else
'                .Cell(flexcpBackColor, Row, 1, Row, .Cols - 1) = .BackColor
'            End If
'        End If
'    End With
End Sub

Private Sub vsfList_DblClick()
    Dim strID As String
    Dim strFlag As String
    
    If vsfList.Col = mIntCol审查结果 Then
        If mcondition.intShowPass = 3 And IsInString(gstrprivs, "合理用药监测", ";") And IsNumeric(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("警告"))) Then
            With vsfList
                If Val(.TextMatrix(.Row, mIntCol门诊标志)) = 1 Or (Val(.TextMatrix(.Row, mIntCol记录性质)) = 2 And Val(.TextMatrix(.Row, mIntCol门诊标志)) = 4) Then
                    strID = .TextMatrix(.Row, .ColIndex("门诊号"))
                    strFlag = "1"
                Else
                    strID = .TextMatrix(.Row, .ColIndex("住院号"))
                    strFlag = "2"
                End If
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassQueryCheckResult_YF(mlngMode, strID, strFlag)
                End If
            End With
        End If
    End If
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        .Editable = flexEDNone
        
        Me.txt用药理由.Text = ""
        
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.ImgList.ListImages(2).Picture
        If Val(.TextMatrix(.Row, mIntColId)) <> 0 Then Me.lblNotice.Caption = "禁忌药品说明：" & .TextMatrix(.Row, mIntCol禁忌药品说明)
        
        
        If .TextMatrix(.Row, mIntCol用药目的) <> "" And Val(.TextMatrix(.Row, mIntColId)) <> 0 Then
            Me.picHscSend.Visible = True
'            Me.txt用药理由.Visible = True
            Me.txt用药理由.Text = "用药目的：" & .TextMatrix(.Row, mIntCol用药目的) & vbCrLf & "用药理由：" & .TextMatrix(.Row, mIntCol用药理由)
            If Not mblnResize Then
                imgDown.Visible = False
                imgUp.Visible = True
            
                picRecipt_Resize
                mblnResize = True
            End If
        Else
            Me.picHscSend.Visible = False
            Me.txt用药理由.Visible = False
    
            If mblnResize Then
                imgDown.Visible = False
                imgUp.Visible = True
                picRecipt_Resize
                mblnResize = False
            End If
        End If
        
        If Val(.TextMatrix(.Row, mIntColId)) = 0 Then Exit Sub
        
        If mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发 Then
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, mintcol药品id), .TextMatrix(.Row, mIntCol其它名))
        ElseIf mcondition.intListType = mListType.退药 Then
            If mcondition.bln显示大小单位 = True Then
                If .Col <> mIntCol退药数大 And .Col <> mIntCol退药数小 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol准退数大)) = 0 And Val(.TextMatrix(.Row, mIntCol准退数小)) = 0 Then Exit Sub
                .Tag = Val(.TextMatrix(.Row, mIntCol准退数大)) * Val(.TextMatrix(.Row, mIntCol包装)) + Val(.TextMatrix(.Row, mIntCol准退数小))
                .Editable = flexEDKbdMouse
            Else
                If .Col <> mIntCol退药数 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol准退数)) = 0 Then Exit Sub
                .Tag = Val(.TextMatrix(.Row, mIntCol准退数))
                .Editable = flexEDKbdMouse
            End If
        End If
    End With
End Sub


Private Sub SetPassMenuButton(ByVal lngRow As Long)
    '设置cmdAlley按钮状态
    Dim cbrControl As CommandBarControl
    Dim rsData As ADODB.Recordset
    
    If mcondition.intShowPass <> 1 Or Not IsInString(gstrprivs, "合理用药监测", ";") Then Exit Sub
    
    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就不显示cmdAlley按钮
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList.TextMatrix(lngRow, vsfList.ColIndex("NO")), Val(vsfList.TextMatrix(lngRow, vsfList.ColIndex("单据"))))
    
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
    
    If rsData.RecordCount = 0 Then
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    
    If mcondition.intListType <> mListType.退药 Then Exit Sub
    
    With vsfList
        strKey = .EditText
        If Col = mIntCol退药数 Or Col = mIntCol退药数大 Or Col = mIntCol退药数小 Then
            If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
                Exit Sub
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If .EditSelLength = Len(strKey) Then Exit Sub
                If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= mintNumberDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub


Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim Int单据 As Integer
    Dim strNo As String
    Dim str审查结果 As Integer
    Dim lng医嘱id As Long
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str挂号单 As String
    Dim lng主页ID As Long
    
    If vsfList.Row = 0 Then Exit Sub
 
    If Button = 2 Then
        If Not gobjPass Is Nothing And IsInString(gstrprivs, "合理用药监测", ";") And (mcondition.intListType = mListType.待发药 Or mcondition.intListType = mListType.超时未发) And vsfList.Col = vsfList.ColIndex("审查结果") Then
            Int单据 = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("单据")))
            strNo = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("NO"))
            lng医嘱id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("医嘱id")))
            str审查结果 = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("警告"))
            
            '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
            strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!病人ID
            str挂号单 = NVL(rsTmp!挂号单)
            lng主页ID = rsTmp!主页id
            
            
  
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng主页ID, str挂号单, str审查结果, lng医嘱id)
            
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub


Private Sub vsfNoList_DblClick()
    vsfNoList_KeyDown vbKeyReturn, 0
End Sub


Private Sub vsfNoList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        vsfNoList.Visible = False
        txtNo.SetFocus
        txtNo.Text = ""
        Exit Sub
    End If
    
    If KeyCode = vbKeyReturn Then
        With vsfNoList
            If .Row = 0 Then Exit Sub
            If .TextMatrix(.Row, .ColIndex("类型")) = "" Then Exit Sub
            
            If CheckAndProcessBill(mcondition.intListType, Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), .TextMatrix(.Row, .ColIndex("药房"))) = False Then
                DoEvents
                txtNo.Text = ""
                txtNo.SetFocus
                Exit Sub
            End If
            
            txtNo.Clear
        
            txtNo.AddItem .TextMatrix(.Row, .ColIndex("NO")) & "--" & .TextMatrix(.Row, .ColIndex("姓名"))
            txtNo.ItemData(txtNo.NewIndex) = .TextMatrix(.Row, .ColIndex("单据"))
            txtNo.Tag = Val(.TextMatrix(.Row, .ColIndex("库房ID"))) & "|" & .TextMatrix(.Row, .ColIndex("药房")) & "|" & .TextMatrix(.Row, .ColIndex("记录性质")) & "|" & .TextMatrix(.Row, .ColIndex("门诊标志")) & "|" & .TextMatrix(.Row, .ColIndex("填制日期")) & "|" & .TextMatrix(.Row, .ColIndex("记录状态"))
            Lbl药房.Caption = .TextMatrix(.Row, .ColIndex("药房"))
            
            txtNo.ListIndex = 0
            
            .Visible = False
        End With
    End If
End Sub


Private Sub vsfNoList_LostFocus()
    If vsfNoList.Visible Then
        vsfNoList.Visible = False
    End If
End Sub


