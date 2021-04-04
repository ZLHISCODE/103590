VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmKssStatistics 
   Caption         =   "抗菌药物统计分析"
   ClientHeight    =   9435
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13905
   DrawStyle       =   6  'Inside Solid
   Icon            =   "frmKssStatistics.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   13905
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Index           =   5
      Left            =   10770
      ScaleHeight     =   4695
      ScaleWidth      =   915
      TabIndex        =   52
      Top             =   795
      Width           =   915
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   5
         Left            =   45
         ScaleHeight     =   810
         ScaleWidth      =   11295
         TabIndex        =   173
         Top             =   45
         Width           =   11295
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   420
            Index           =   14
            Left            =   4035
            ScaleHeight     =   420
            ScaleWidth      =   1395
            TabIndex        =   196
            TabStop         =   0   'False
            Top             =   345
            Width           =   1395
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "按疾病编码"
               Height          =   180
               Index           =   28
               Left            =   0
               TabIndex        =   197
               Top             =   210
               Width           =   1290
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "按诊断标准"
               Height          =   180
               Index           =   27
               Left            =   0
               TabIndex        =   183
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdILL 
            Caption         =   "…"
            Height          =   255
            Left            =   3660
            TabIndex        =   175
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+I"
            Top             =   420
            Width           =   270
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   13
            Left            =   75
            ScaleHeight     =   510
            ScaleWidth      =   735
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   315
            Width           =   735
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "中医"
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
               Index           =   26
               Left            =   0
               TabIndex        =   195
               Top             =   225
               Width           =   690
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "西医"
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
               Index           =   25
               Left            =   0
               TabIndex        =   181
               Top             =   0
               Value           =   -1  'True
               Width           =   690
            End
         End
         Begin VB.TextBox txtILL 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   825
            MaxLength       =   100
            TabIndex        =   182
            Top             =   405
            Width           =   3120
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   9
            Left            =   9705
            TabIndex        =   187
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   184
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "平均"
               Height          =   180
               Index           =   20
               Left            =   0
               TabIndex        =   186
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "随机"
               Height          =   180
               Index           =   19
               Left            =   690
               TabIndex        =   189
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   5
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   177
            Top             =   0
            Width           =   1365
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   3
            Left            =   6660
            MaxLength       =   3
            TabIndex        =   185
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   5
            Left            =   4245
            TabIndex        =   179
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   5
            Left            =   2565
            TabIndex        =   178
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   71
            Left            =   7440
            TabIndex        =   193
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   69
            Left            =   2310
            TabIndex        =   192
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   68
            Left            =   0
            TabIndex        =   191
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   60
            Left            =   5805
            TabIndex        =   190
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsIllDruUse 
         Height          =   690
         Left            =   60
         TabIndex        =   188
         Top             =   1350
         Width           =   7710
         _cx             =   13600
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6852
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Index           =   67
         Left            =   285
         TabIndex        =   198
         Top             =   2265
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "医生治疗某疾病抗菌用药成本统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   195
         TabIndex        =   53
         Top             =   1005
         Width           =   3825
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6060
      Index           =   3
      Left            =   8490
      ScaleHeight     =   6060
      ScaleWidth      =   825
      TabIndex        =   48
      Top             =   735
      Width           =   825
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   3
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11865
         TabIndex        =   124
         Top             =   0
         Width           =   11865
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   255
            Index           =   6
            Left            =   5325
            TabIndex        =   125
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   8115
            ScaleHeight     =   255
            ScaleWidth      =   3525
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   30
            Width           =   3525
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅳ类"
               Height          =   210
               Index           =   8
               Left            =   2625
               TabIndex        =   144
               Top             =   15
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅲ类"
               Height          =   210
               Index           =   4
               Left            =   1890
               TabIndex        =   143
               Top             =   15
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅰ类"
               Height          =   225
               Index           =   2
               Left            =   240
               TabIndex        =   141
               Top             =   15
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅱ类"
               Height          =   210
               Index           =   3
               Left            =   1035
               TabIndex        =   142
               Top             =   15
               Value           =   1  'Checked
               Width           =   750
            End
            Begin VB.Label lblN 
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "(                                   )"
               Height          =   180
               Index           =   57
               Left            =   75
               TabIndex        =   140
               Top             =   30
               Width           =   3330
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   6630
            ScaleHeight     =   255
            ScaleWidth      =   1500
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   45
            Width           =   1500
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "手术"
               Height          =   180
               Index           =   18
               Left            =   840
               TabIndex        =   138
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "非手术"
               Height          =   180
               Index           =   15
               Left            =   0
               TabIndex        =   137
               Top             =   0
               Width           =   840
            End
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   6
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   145
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   405
            Width           =   4800
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   7
            Left            =   9705
            TabIndex        =   150
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "平均"
               Height          =   180
               Index           =   14
               Left            =   0
               TabIndex        =   147
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "随机"
               Height          =   180
               Index           =   13
               Left            =   690
               TabIndex        =   149
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   3
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   0
            Width           =   1365
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   1
            Left            =   6660
            MaxLength       =   4
            TabIndex        =   146
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   3
            Left            =   4245
            TabIndex        =   129
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   127
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   56
            Left            =   7440
            TabIndex        =   135
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计科室"
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
            Index           =   55
            Left            =   0
            TabIndex        =   134
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   54
            Left            =   2310
            TabIndex        =   133
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   53
            Left            =   0
            TabIndex        =   132
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "切口类型"
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
            Index           =   52
            Left            =   5805
            TabIndex        =   131
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   51
            Left            =   5805
            TabIndex        =   130
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInDruUse 
         Height          =   690
         Left            =   150
         TabIndex        =   148
         Top             =   1395
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6926
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   150
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInDruAna 
         Height          =   690
         Left            =   195
         TabIndex        =   151
         Top             =   2460
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6A51
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX例出院病人抗菌药物使用统计分析表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   59
         Left            =   210
         TabIndex        =   153
         Top             =   2145
         Width           =   4485
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Index           =   58
         Left            =   240
         TabIndex        =   152
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "住院医嘱抗菌用药统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   49
         Top             =   975
         Width           =   2550
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5670
      Index           =   2
      Left            =   7320
      ScaleHeight     =   5670
      ScaleWidth      =   990
      TabIndex        =   46
      Top             =   690
      Width           =   990
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   2
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11025
         TabIndex        =   98
         Top             =   0
         Width           =   11025
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   255
            Index           =   5
            Left            =   5325
            TabIndex        =   99
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H80000005&
            Caption         =   "急诊"
            Height          =   210
            Index           =   1
            Left            =   7395
            TabIndex        =   106
            Top             =   45
            Width           =   690
         End
         Begin VB.CheckBox chkType 
            BackColor       =   &H80000005&
            Caption         =   "门诊"
            Height          =   225
            Index           =   0
            Left            =   6660
            TabIndex        =   105
            Top             =   45
            Width           =   690
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   0
            Left            =   6660
            MaxLength       =   4
            TabIndex        =   108
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   2
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "随机"
               Height          =   180
               Index           =   10
               Left            =   690
               TabIndex        =   101
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "平均"
               Height          =   180
               Index           =   6
               Left            =   0
               TabIndex        =   109
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   6
            Left            =   9705
            TabIndex        =   112
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   5
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   107
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   405
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   2
            Left            =   4245
            TabIndex        =   104
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   103
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   38
            Left            =   5805
            TabIndex        =   116
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "类型"
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
            Index           =   37
            Left            =   6195
            TabIndex        =   115
            Top             =   60
            Width           =   405
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   34
            Left            =   0
            TabIndex        =   114
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   33
            Left            =   2310
            TabIndex        =   113
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计科室"
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
            Index           =   32
            Left            =   0
            TabIndex        =   111
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   3
            Left            =   7440
            TabIndex        =   110
            Top             =   465
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCountDruUse 
         Height          =   690
         Left            =   60
         TabIndex        =   122
         Top             =   1620
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6A8F
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCountCF 
         Height          =   690
         Left            =   45
         TabIndex        =   123
         Top             =   3180
         Width           =   6975
         _cx             =   12303
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6BBA
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Index           =   50
         Left            =   90
         TabIndex        =   200
         Top             =   3960
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "处方总量：XXXX张"
         Height          =   180
         Index           =   49
         Left            =   0
         TabIndex        =   121
         Top             =   2910
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "处方总量：XXXX张"
         Height          =   180
         Index           =   48
         Left            =   480
         TabIndex        =   120
         Top             =   1410
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "日期：2014-12-11至2014-12-31"
         Height          =   180
         Index           =   47
         Left            =   45
         TabIndex        =   119
         Top             =   2655
         Width           =   2520
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "日期：2014-12-11至2014-12-31"
         Height          =   180
         Index           =   46
         Left            =   105
         TabIndex        =   118
         Top             =   1185
         Width           =   2520
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX张处方统计分析表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   39
         Left            =   45
         TabIndex        =   117
         Top             =   2385
         Width           =   2445
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "门(急)诊处方抗菌用药统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   105
         TabIndex        =   47
         Top             =   885
         Width           =   3075
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6570
      Index           =   1
      Left            =   6285
      ScaleHeight     =   6570
      ScaleWidth      =   705
      TabIndex        =   45
      Top             =   600
      Width           =   705
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   795
         Index           =   1
         Left            =   45
         ScaleHeight     =   795
         ScaleWidth      =   9120
         TabIndex        =   82
         Top             =   90
         Width           =   9120
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   255
            Index           =   4
            Left            =   5325
            TabIndex        =   83
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   4
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   87
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   405
            Width           =   4800
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   5
            Left            =   8040
            TabIndex        =   89
            Top             =   405
            Width           =   960
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   6615
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "科室"
               Height          =   180
               Index           =   17
               Left            =   0
               TabIndex        =   88
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "医生"
               Height          =   180
               Index           =   16
               Left            =   690
               TabIndex        =   91
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   1
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   1
            Left            =   4245
            TabIndex        =   86
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   85
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "汇总方式"
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
            Index           =   36
            Left            =   5775
            TabIndex        =   95
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计科室"
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
            Index           =   35
            Left            =   0
            TabIndex        =   94
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   28
            Left            =   2310
            TabIndex        =   93
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   2
            Left            =   0
            TabIndex        =   92
            Top             =   60
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCut 
         Height          =   915
         Left            =   210
         TabIndex        =   96
         Top             =   1890
         Width           =   6975
         _cx             =   12303
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6BF8
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   110
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Ⅰ类切口围术期预防用药统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   76
         Left            =   450
         TabIndex        =   245
         Top             =   1095
         Width           =   3315
      End
      Begin VB.Label lblCut 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Left            =   330
         TabIndex        =   97
         Top             =   3120
         Width           =   360
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Index           =   0
      Left            =   5325
      ScaleHeight     =   5025
      ScaleWidth      =   720
      TabIndex        =   54
      Top             =   660
      Width           =   720
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   0
         Left            =   135
         ScaleHeight     =   1155
         ScaleWidth      =   10575
         TabIndex        =   57
         Top             =   60
         Width           =   10575
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   255
            Index           =   0
            Left            =   5325
            TabIndex        =   58
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   780
            Width           =   285
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   0
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   66
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   855
            ScaleHeight     =   240
            ScaleWidth      =   1455
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   450
            Width           =   1455
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "门诊"
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "住院"
               Height          =   180
               Index           =   5
               Left            =   720
               TabIndex        =   70
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3465
            ScaleHeight     =   255
            ScaleWidth      =   2115
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   465
            Width           =   2115
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "药品"
               Height          =   180
               Index           =   7
               Left            =   1440
               TabIndex        =   72
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "医生"
               Height          =   180
               Index           =   8
               Left            =   705
               TabIndex        =   63
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "科室"
               Height          =   180
               Index           =   9
               Left            =   0
               TabIndex        =   62
               Top             =   0
               Width           =   660
            End
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   6615
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   810
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "金额"
               Height          =   180
               Index           =   11
               Left            =   690
               TabIndex        =   60
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "数量"
               Height          =   180
               Index           =   12
               Left            =   0
               TabIndex        =   76
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   4
            Left            =   9555
            TabIndex        =   80
            Top             =   750
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   0
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   74
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   750
            Width           =   4800
         End
         Begin VB.TextBox txtTopRan 
            Height          =   300
            Left            =   8745
            MaxLength       =   3
            TabIndex        =   78
            Text            =   "10"
            Top             =   750
            Width           =   465
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   0
            Left            =   4245
            TabIndex        =   68
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   67
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   8
            Left            =   0
            TabIndex        =   81
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   9
            Left            =   2310
            TabIndex        =   79
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计场合"
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
            Index           =   16
            Left            =   0
            TabIndex        =   77
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "汇总方式"
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
            Index           =   17
            Left            =   2565
            TabIndex        =   75
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计前     名"
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
            Index           =   31
            Left            =   8130
            TabIndex        =   73
            Top             =   810
            Width           =   1380
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计科室"
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
            Index           =   18
            Left            =   0
            TabIndex        =   71
            Top             =   810
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "排序方式"
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
            Index           =   19
            Left            =   5775
            TabIndex        =   69
            Top             =   810
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsUseRan 
         Height          =   915
         Left            =   570
         TabIndex        =   55
         Top             =   1755
         Width           =   6975
         _cx             =   12303
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6CBD
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "抗菌药物使用情况排名统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   75
         Left            =   180
         TabIndex        =   244
         Top             =   1290
         Width           =   3060
      End
      Begin VB.Label lblUse 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Left            =   120
         TabIndex        =   56
         Top             =   1875
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4410
      Index           =   3
      Left            =   3735
      ScaleHeight     =   4410
      ScaleWidth      =   1125
      TabIndex        =   205
      Top             =   1455
      Width           =   1125
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   930
         Index           =   9
         Left            =   0
         ScaleHeight     =   930
         ScaleWidth      =   10125
         TabIndex        =   232
         Top             =   105
         Width           =   10125
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   265
            Index           =   3
            Left            =   5325
            TabIndex        =   233
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   3
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   3
            Left            =   5745
            TabIndex        =   31
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   3
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   30
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   400
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   3
            Left            =   4245
            TabIndex        =   29
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   3
            Left            =   2565
            TabIndex        =   28
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   41
            Left            =   2310
            TabIndex        =   236
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   40
            Left            =   0
            TabIndex        =   235
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样科室"
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
            Index           =   42
            Left            =   0
            TabIndex        =   234
            Top             =   450
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsZYYY 
         Height          =   915
         Left            =   30
         TabIndex        =   32
         Top             =   2520
         Width           =   5595
         _cx             =   9869
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明："
         Height          =   180
         Index           =   44
         Left            =   120
         TabIndex        =   231
         Top             =   3585
         Width           =   540
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "收治患者人天数：XXXX天"
         Height          =   180
         Index           =   45
         Left            =   30
         TabIndex        =   230
         Top             =   2280
         Width           =   1980
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "住院病人抗菌用药调查表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   43
         Left            =   30
         TabIndex        =   229
         Top             =   1935
         Width           =   2805
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Index           =   2
      Left            =   2355
      ScaleHeight     =   5475
      ScaleWidth      =   795
      TabIndex        =   204
      Top             =   1200
      Width           =   795
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   8
         Left            =   0
         ScaleHeight     =   1005
         ScaleWidth      =   10845
         TabIndex        =   222
         Top             =   0
         Width           =   10845
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   265
            Index           =   2
            Left            =   5325
            TabIndex        =   223
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   2
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   20
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   400
            Width           =   4800
         End
         Begin VB.TextBox txtCount 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   2
            Left            =   6690
            MaxLength       =   3
            TabIndex        =   21
            Text            =   "100"
            Top             =   405
            Width           =   420
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   2
            Left            =   9570
            TabIndex        =   24
            Top             =   405
            Width           =   960
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "平均"
            Height          =   180
            Index           =   3
            Left            =   8100
            TabIndex        =   22
            Top             =   450
            Width           =   660
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "随机"
            Height          =   180
            Index           =   2
            Left            =   8775
            TabIndex        =   23
            Top             =   450
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   2
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   2
            Left            =   4245
            TabIndex        =   19
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   2
            Left            =   2565
            TabIndex        =   18
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样科室"
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
            Index           =   22
            Left            =   0
            TabIndex        =   228
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   23
            Left            =   5835
            TabIndex        =   227
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   24
            Left            =   7290
            TabIndex        =   226
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   20
            Left            =   0
            TabIndex        =   225
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   21
            Left            =   2310
            TabIndex        =   224
            Top             =   60
            Width           =   2970
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsCF 
         Height          =   915
         Left            =   330
         TabIndex        =   26
         Top             =   3645
         Width           =   5595
         _cx             =   9869
         _cy             =   1614
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6DD6
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMZYY 
         Height          =   825
         Left            =   75
         TabIndex        =   25
         Top             =   1875
         Width           =   5595
         _cx             =   9869
         _cy             =   1455
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6E12
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "门诊病人用药情况调查表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   60
         TabIndex        =   221
         Top             =   1095
         Width           =   2805
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "日期：XXX至XXX"
         Height          =   180
         Index           =   26
         Left            =   15
         TabIndex        =   220
         Top             =   1365
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "处方总量：XXXX张"
         Height          =   180
         Index           =   27
         Left            =   60
         TabIndex        =   219
         Top             =   1575
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "处方总量：XXXX张"
         Height          =   180
         Index           =   30
         Left            =   90
         TabIndex        =   218
         Top             =   3345
         Width           =   1440
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "日期：XXX至XXX"
         Height          =   180
         Index           =   29
         Left            =   -30
         TabIndex        =   217
         Top             =   3120
         Width           =   1260
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "XXX张处方统计分析表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   70
         Left            =   30
         TabIndex        =   216
         Top             =   2850
         Width           =   2445
      End
      Begin VB.Label lblCFSM 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Left            =   315
         TabIndex        =   215
         Top             =   4710
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5145
      Index           =   1
      Left            =   1290
      ScaleHeight     =   5145
      ScaleWidth      =   810
      TabIndex        =   203
      Top             =   1125
      Width           =   810
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   525
         Left            =   750
         TabIndex        =   16
         Top             =   1590
         Width           =   8340
         _Version        =   589884
         _ExtentX        =   14711
         _ExtentY        =   926
         _StockProps     =   0
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1245
         Index           =   7
         Left            =   0
         ScaleHeight     =   1245
         ScaleWidth      =   10650
         TabIndex        =   206
         Top             =   0
         Width           =   10650
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   265
            Index           =   1
            Left            =   5325
            TabIndex        =   207
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   1
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   0
            Width           =   1365
         End
         Begin VB.CommandButton cmdCYEdit 
            Caption         =   "编辑调查表(&E)"
            Height          =   300
            Left            =   7410
            TabIndex        =   15
            Top             =   765
            Width           =   1550
         End
         Begin VB.CommandButton cmdCYDel 
            Caption         =   "删除抽样记录(&D)"
            Height          =   300
            Left            =   5835
            TabIndex        =   14
            Top             =   765
            Width           =   1550
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "随机"
            Height          =   180
            Index           =   1
            Left            =   8775
            TabIndex        =   11
            Top             =   450
            Value           =   -1  'True
            Width           =   660
         End
         Begin VB.OptionButton optType 
            BackColor       =   &H80000005&
            Caption         =   "平均"
            Height          =   180
            Index           =   0
            Left            =   8100
            TabIndex        =   10
            Top             =   450
            Width           =   660
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "抽样(&C)"
            Height          =   300
            Index           =   1
            Left            =   9570
            TabIndex        =   12
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtCount 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   1
            Left            =   6690
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "15"
            Top             =   405
            Width           =   450
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   1
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   8
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   400
            Width           =   4800
         End
         Begin VB.CommandButton cmdCYSel 
            Caption         =   "…"
            Height          =   265
            Left            =   5325
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+R"
            Top             =   795
            Width           =   285
         End
         Begin VB.TextBox txtCYJL 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   13
            Top             =   765
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   1
            Left            =   4245
            TabIndex        =   7
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   1
            Left            =   2565
            TabIndex        =   6
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   11
            Left            =   2310
            TabIndex        =   214
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   10
            Left            =   0
            TabIndex        =   213
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "抽样记录"
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
            Index           =   15
            Left            =   0
            TabIndex        =   212
            Top             =   825
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   14
            Left            =   7290
            TabIndex        =   211
            Top             =   450
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   13
            Left            =   5835
            TabIndex        =   210
            Top             =   450
            Width           =   780
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "抽样科室"
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
            Index           =   12
            Left            =   0
            TabIndex        =   209
            Top             =   450
            Width           =   780
         End
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "(非)手术病人抗菌用药情况抽样调查及评价表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   77
         Left            =   405
         TabIndex        =   246
         Top             =   1290
         Width           =   5115
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Index           =   72
         Left            =   300
         TabIndex        =   240
         Top             =   1980
         Width           =   360
      End
   End
   Begin VB.PictureBox picReportSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5160
      Index           =   0
      Left            =   45
      ScaleHeight     =   5160
      ScaleWidth      =   1065
      TabIndex        =   201
      Top             =   1995
      Width           =   1065
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   30
         ScaleHeight     =   495
         ScaleWidth      =   7275
         TabIndex        =   237
         Top             =   75
         Width           =   7275
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   0
            Left            =   5865
            TabIndex        =   3
            Top             =   0
            Width           =   960
         End
         Begin VB.ComboBox cboTimRP 
            Height          =   300
            Index           =   0
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   0
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpRPE 
            Height          =   300
            Index           =   0
            Left            =   4245
            TabIndex        =   2
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpRPS 
            Height          =   300
            Index           =   0
            Left            =   2565
            TabIndex        =   1
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   1
            Left            =   2310
            TabIndex        =   239
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   0
            Left            =   0
            TabIndex        =   238
            Top             =   60
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBill 
         Height          =   525
         Left            =   1185
         TabIndex        =   4
         Top             =   1245
         Width           =   3855
         _cx             =   6800
         _cy             =   926
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6EB0
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "抗菌药品消耗金额调查表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   74
         Left            =   420
         TabIndex        =   243
         Top             =   705
         Width           =   2805
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Left            =   435
         TabIndex        =   202
         Top             =   1515
         Width           =   360
      End
   End
   Begin VB.PictureBox picOtherSub 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4050
      Index           =   4
      Left            =   9630
      ScaleHeight     =   4050
      ScaleWidth      =   900
      TabIndex        =   50
      Top             =   780
      Width           =   900
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Index           =   4
         Left            =   0
         ScaleHeight     =   780
         ScaleWidth      =   11295
         TabIndex        =   154
         Top             =   0
         Width           =   11295
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   6645
            ScaleHeight     =   255
            ScaleWidth      =   2835
            TabIndex        =   156
            TabStop         =   0   'False
            Top             =   30
            Width           =   2835
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅳ类"
               Height          =   210
               Index           =   9
               Left            =   2175
               TabIndex        =   199
               Top             =   30
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅰ类"
               Height          =   210
               Index           =   5
               Left            =   15
               TabIndex        =   166
               Top             =   30
               Value           =   1  'Checked
               Width           =   690
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅱ类"
               Height          =   225
               Index           =   6
               Left            =   705
               TabIndex        =   157
               Top             =   15
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CheckBox chkType 
               BackColor       =   &H80000005&
               Caption         =   "Ⅲ类"
               Height          =   210
               Index           =   7
               Left            =   1470
               TabIndex        =   158
               Top             =   30
               Value           =   1  'Checked
               Width           =   750
            End
         End
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Height          =   255
            Index           =   7
            Left            =   5325
            TabIndex        =   155
            TabStop         =   0   'False
            ToolTipText     =   "Ctrl+D"
            Top             =   420
            Width           =   285
         End
         Begin VB.TextBox txtNum 
            Height          =   300
            Index           =   2
            Left            =   6660
            MaxLength       =   3
            TabIndex        =   174
            Text            =   "10"
            Top             =   405
            Width           =   465
         End
         Begin VB.ComboBox cboTimCount 
            Height          =   300
            Index           =   4
            Left            =   825
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   0
            Width           =   1365
         End
         Begin VB.PictureBox picOpt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   8280
            ScaleHeight     =   255
            ScaleWidth      =   1380
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   465
            Width           =   1380
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "随机"
               Height          =   180
               Index           =   22
               Left            =   690
               TabIndex        =   160
               Top             =   0
               Width           =   660
            End
            Begin VB.OptionButton optType 
               BackColor       =   &H80000005&
               Caption         =   "平均"
               Height          =   180
               Index           =   21
               Left            =   0
               TabIndex        =   176
               Top             =   0
               Value           =   -1  'True
               Width           =   660
            End
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "统计(&T)"
            Height          =   300
            Index           =   8
            Left            =   9705
            TabIndex        =   180
            Top             =   405
            Width           =   960
         End
         Begin VB.TextBox txtDept 
            Height          =   300
            IMEMode         =   1  'ON
            Index           =   7
            Left            =   825
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   172
            Text            =   "所有科室"
            ToolTipText     =   "所有科室"
            Top             =   405
            Width           =   4800
         End
         Begin MSComCtl2.DTPicker dtpCountE 
            Height          =   300
            Index           =   4
            Left            =   4245
            TabIndex        =   164
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin MSComCtl2.DTPicker dtpCountS 
            Height          =   300
            Index           =   4
            Left            =   2565
            TabIndex        =   162
            Top             =   0
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   124059651
            CurrentDate     =   41774
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样数量"
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
            Index           =   66
            Left            =   5805
            TabIndex        =   170
            Top             =   450
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "切口类型"
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
            Index           =   65
            Left            =   5805
            TabIndex        =   169
            Top             =   60
            Width           =   780
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计时间"
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
            Index           =   64
            Left            =   0
            TabIndex        =   168
            Top             =   60
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "从                 至"
            Height          =   180
            Index           =   63
            Left            =   2310
            TabIndex        =   167
            Top             =   60
            Width           =   2970
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "统计科室"
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
            Index           =   62
            Left            =   0
            TabIndex        =   165
            Top             =   465
            Width           =   795
         End
         Begin VB.Label lblN 
            BackColor       =   &H80000005&
            Caption         =   "抽样方法"
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
            Index           =   61
            Left            =   7440
            TabIndex        =   163
            Top             =   465
            Width           =   795
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsOpeKssUse 
         Height          =   690
         Left            =   60
         TabIndex        =   171
         Top             =   1260
         Width           =   7710
         _cx             =   13600
         _cy             =   1217
         Appearance      =   2
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
         BackColorSel    =   16444122
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmKssStatistics.frx":6F2B
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "说明"
         Height          =   180
         Index           =   73
         Left            =   180
         TabIndex        =   241
         Top             =   2610
         Width           =   360
      End
      Begin VB.Label lblN 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "术后抗菌药物使用超N天统计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   75
         TabIndex        =   51
         Top             =   825
         Width           =   3195
      End
   End
   Begin VB.PictureBox picOther 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   5190
      ScaleHeight     =   855
      ScaleWidth      =   8865
      TabIndex        =   35
      Top             =   165
      Width           =   8865
      Begin XtremeSuiteControls.TabControl tbcOther 
         Height          =   960
         Left            =   225
         TabIndex        =   37
         Top             =   75
         Width           =   8490
         _Version        =   589884
         _ExtentX        =   14975
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picDept 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   1890
      ScaleHeight     =   2850
      ScaleWidth      =   4890
      TabIndex        =   38
      Top             =   7380
      Visible         =   0   'False
      Width           =   4920
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   50
         TabIndex        =   42
         Top             =   75
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找"
         Height          =   270
         Left            =   1740
         TabIndex        =   41
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindOk 
         Caption         =   "确定"
         Height          =   270
         Left            =   3480
         TabIndex        =   40
         Top             =   75
         Width           =   615
      End
      Begin VB.CommandButton cmdFindCancle 
         Caption         =   "取消"
         Height          =   270
         Left            =   4200
         TabIndex        =   39
         Top             =   75
         Width           =   615
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2280
         Left            =   75
         TabIndex        =   43
         ToolTipText     =   "全选Ctrl+A；全清Ctrl+R"
         Top             =   510
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16"
         SmallIcons      =   "img16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   165
      ScaleHeight     =   810
      ScaleWidth      =   4935
      TabIndex        =   34
      Top             =   675
      Width           =   4935
      Begin XtremeSuiteControls.TabControl tbcReport 
         Height          =   2535
         Left            =   120
         TabIndex        =   36
         Top             =   75
         Width           =   6975
         _Version        =   589884
         _ExtentX        =   12303
         _ExtentY        =   4471
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   840
      Top             =   15
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
            Picture         =   "frmKssStatistics.frx":702E
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":75C8
            Key             =   "PatiMan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":7B62
            Key             =   "PatiWoman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":80FC
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":8696
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":8C30
            Key             =   "单病种"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":F492
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKssStatistics.frx":15CF4
            Key             =   "printer"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   525
      Left            =   1770
      TabIndex        =   33
      Top             =   45
      Width           =   660
      _Version        =   589884
      _ExtentX        =   1164
      _ExtentY        =   926
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   44
      Top             =   9075
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKssStatistics.frx":1C556
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20003
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   " "
            TextSave        =   " "
            Key             =   "病人颜色"
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
   Begin VSFlex8Ctl.VSFlexGrid vsTmp 
      Height          =   465
      Left            =   3495
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
      _cx             =   2117
      _cy             =   820
      Appearance      =   2
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmKssStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mEnumPanel
    '上报数据
    PanelItem_抗菌药品消耗金额调查表 = 0         'picReportSub(0)  picBill
    PanelItem_病人抗菌用药情况抽样调查及评价表 = 1 'picReportSub(1)  picCY----(非)手术病人抗菌用药情况抽样调查及评价表
    PanelItem_门诊处方抗菌用药调查表 = 2         'picReportSub(2)  picCF
    PanelItem_住院病人抗菌用药调查表 = 3         'picReportSub(3)  picYZ

    '统计数据
    PanelItem_抗菌药物使用情况排名统计 = 0       'picOtherSub(0) 统计
    PanelItem_Ⅰ类切口围术期预防用药统计 = 1     'picOtherSub(1)
    PanelItem_门急诊处方抗菌用药统计 = 2         'picOtherSub(2) ----门(急)诊处方抗菌用药统计
    PanelItem_住院医嘱抗菌用药统计 = 3           'picOtherSub(3)
    PanelItem_术后抗菌药物使用超N天统计 = 4      'picOtherSub(4)
    PanelItem_医生治疗某疾病抗菌用药成本统计 = 5 'picOtherSub(e_C3_lblN_统计表_标题_5)统计
End Enum

'界面控件的下标，格式说明  e_R0_cbo_12_统计时间,R0表示上报数据的第一个界面，R1上报第二个界面，C0统计数据第一个界面，C1统计数据第二个界面
Private Enum mCtlID
    e_R0_cboTimRP_统计时间_0 = 0
    e_R0_dtpRPS_开始时间_0 = 0
    e_R0_dtpRPE_结束时间_0 = 0
    e_R0_cmdOK_统计_0 = 0
    e_R0_picFilter_条件容器_6 = 6
    e_R0_lblN_标题_74 = 74
    
    e_R1_cboTimRP_统计时间_1 = 1
    e_R1_dtpRPS_开始时间_1 = 1
    e_R1_dtpRPE_结束时间_1 = 1
    e_R1_picFilter_条件容器_7 = 7
    e_R1_lblN_底端说明_72 = 72
    e_R1_txtCount_抽样数量_1 = 1
    e_R1_optType_抽样方法_平均_0 = 0
    e_R1_optType_抽样方法_随机_1 = 1
    e_R1_cmdOK_抽样_1 = 1
    e_R1_cmdDept_科室选择器_1 = 1
    e_R1_txtDept_抽样科室_1 = 1
    e_R1_lblN_抽样数量标题_13 = 13
    e_R1_lblN_标题_77 = 77
    
    e_R2_cboTimRP_统计时间_2 = 2
    e_R2_dtpRPS_开始时间_2 = 2
    e_R2_dtpRPE_结束时间_2 = 2
    e_R2_cmdDept_科室选择器_2 = 2
    e_R2_txtDept_抽样科室_2 = 2
    e_R2_txtCount_抽样数量_2 = 2
    e_R2_optType_抽样方法_平均_3 = 3
    e_R2_optType_抽样方法_随机_2 = 2
    e_R2_cmdOK_统计_2 = 2
    e_R2_picFilter_条件容器_8 = 8
    e_R2_lblN_调查表_标题_25 = 25
    e_R2_lblN_调查表_日期_26 = 26
    e_R2_lblN_调查表_处方总量_27 = 27
    e_R2_lblN_分析表_标题_70 = 70
    e_R2_lblN_分析表_日期_29 = 29
    e_R2_lblN_分析表_处方总量_30 = 30
    
    e_R3_cboTimRP_统计时间_3 = 3
    e_R3_dtpRPS_开始时间_3 = 3
    e_R3_dtpRPE_结束时间_3 = 3
    e_R3_cmdDept_科室选择器_3 = 3
    e_R3_txtDept_抽样科室_3 = 3
    e_R3_cmdOK_统计_3 = 3
    e_R3_picFilter_条件容器_9 = 9
    e_R3_lblN_调查表_标题_43 = 43
    e_R3_lblN_调查表_患者天数_45 = 45
    e_R3_lblN_底端说明_44 = 44
    
    e_C0_cboTimCount_统计时间_0 = 0
    e_C0_dtpCountS_开始时间_0 = 0
    e_C0_dtpCountE_结束时间_0 = 0
    e_C0_txtDept_抽样科室_0 = 0
    e_C0_optType_统计场合_住院_5 = 5
    e_C0_optType_统计场合_门诊_4 = 4
    e_C0_optType_汇总方式_科室_9 = 9
    e_C0_optType_汇总方式_医生_8 = 8
    e_C0_optType_汇总方式_药品_7 = 7
    e_C0_optType_排序方式_数量_12 = 12
    e_C0_optType_排序方式_金额_11 = 11
    e_C0_cmdOK_统计_4 = 4
    e_C0_cmdDept_科室选择器_0 = 0
    e_C0_picFilter_条件容器_0 = 0
    e_C0_lblN_标题_75 = 75
    
    e_C1_cboTimCount_统计时间_1 = 1
    e_C1_dtpCountS_开始时间_1 = 1
    e_C1_dtpCountE_结束时间_1 = 1
    e_C1_txtDept_统计科室_4 = 4
    e_C1_cmdDept_科室选择器_4 = 4
    e_C1_optType_汇总方式_科室_17 = 17
    e_C1_optType_汇总方式_医生_16 = 16
    e_C1_cmdOK_统计_5 = 5
    e_C1_picFilter_条件容器_1 = 1
    e_C1_lblN_标题_76 = 76
    
    e_C2_cboTimCount_统计时间_2 = 2
    e_C2_dtpCountS_开始时间_2 = 2
    e_C2_dtpCountE_结束时间_2 = 2
    e_C2_chkType_类型_门诊_0 = 0
    e_C2_chkType_类型_急诊_1 = 1
    e_C2_txtDept_统计科室_5 = 5
    e_C2_cmdDept_科室选择器_5 = 5
    e_C2_txtNum_统计科室_0 = 0
    e_C2_optType_抽样方法_平均_6 = 6
    e_C2_optType_抽样方法_随机_10 = 10
    e_C2_cmdOK_统计_6 = 6
    e_C2_picFilter_条件容器_2 = 2
    e_C2_lblN_统计表_标题_4 = 4
    e_C2_lblN_统计表_日期_46 = 46
    e_C2_lblN_统计表_处方总量_48 = 48
    e_C2_lblN_分析表_标题_39 = 39
    e_C2_lblN_分析表_日期_47 = 47
    e_C2_lblN_分析表_处方总量_49 = 49
    e_C2_lblN_底端说明_50 = 50
    
    e_C3_cboTimCount_统计时间_3 = 3
    e_C3_dtpCountS_开始时间_3 = 3
    e_C3_dtpCountE_结束时间_3 = 3
    e_C3_txtDept_统计科室_6 = 6
    e_C3_cmdDept_科室选择器_6 = 6
    e_C3_txtNum_抽样数量_1 = 1
    e_C3_cmdOK_统计_7 = 7
    e_C3_optType_切口类型_非手术_15 = 15
    e_C3_optType_切口类型_手术_18 = 18
    e_C3_optType_抽样方法_平均_14 = 14
    e_C3_optType_抽样方法_随机_13 = 13
    e_C3_chkType_切口类型_Ⅰ类_2 = 2
    e_C3_chkType_切口类型_Ⅱ类_3 = 3
    e_C3_chkType_切口类型_Ⅲ类_4 = 4
    e_C3_chkType_切口类型_Ⅳ类_8 = 8
    e_C3_picFilter_条件容器_3 = 3
    e_C3_lblN_统计表_标题_5 = 5
    e_C3_lblN_分析表_标题_59 = 59
    e_C3_lblN_底端说明_58 = 58
    
    e_C4_cboTimCount_统计时间_4 = 4
    e_C4_dtpCountS_开始时间_4 = 4
    e_C4_dtpCountE_结束时间_4 = 4
    e_C4_chkType_切口类型_Ⅰ类_5 = 5
    e_C4_chkType_切口类型_Ⅱ类_6 = 6
    e_C4_chkType_切口类型_Ⅲ类_7 = 7
    e_C4_chkType_切口类型_Ⅳ类_9 = 9
    e_C4_txtDept_统计科室_7 = 7
    e_C4_cmdDept_科室选择器_7 = 7
    e_C4_txtNum_抽样数量_2 = 2
    e_C4_optType_抽样方法_平均_21 = 21
    e_C4_optType_抽样方法_随机_22 = 22
    e_C4_cmdOK_统计_8 = 8
    e_C4_picFilter_条件容器_4 = 4
    e_C4_lblN_统计表_标题_6 = 6
    e_C4_lblN_底端说明_73 = 73
    
    e_C5_cboTimCount_统计时间_5 = 5
    e_C5_dtpCountS_开始时间_5 = 5
    e_C5_dtpCountE_结束时间_5 = 5
    e_C5_optType_西医_25 = 25
    e_C5_optType_中医_26 = 26
    e_C5_optType_按诊断_27 = 27
    e_C5_optType_按疾病_28 = 28
    e_C5_txtNum_抽样数量_3 = 3
    e_C5_optType_抽样方法_平均_20 = 20
    e_C5_optType_抽样方法_随机_19 = 19
    e_C5_cmdOK_统计_9 = 9
    e_C5_picFilter_条件容器_5 = 5
    e_C5_lblN_分析表_标题_7 = 7
    e_C5_lblN_底端说明_67 = 67
    
End Enum


Private Enum COL_VSBILL '上报数据  抗菌药品消耗金额调查表  vsBill 列索引
    COL_统计项目 = 0
    COL_结果 = 1
    COL_单位 = 2
    COL_备注 = 3
End Enum

Private Enum ROW_VSBILL '上报数据  抗菌药品消耗金额调查表   vsBill  行索引
    ROW_年医院总收入 = 1
    ROW_政府拨款 = 2
    ROW_年药品总收入 = 3
    ROW_药品占医院总收入比例 = 4
    ROW_药品进销差价收入 = 5
    ROW_药品进销差价收入占医院总收入比例 = 6
    ROW_西药全年使用金额 = 7
    ROW_门诊西药房 = 8
    ROW_住院西药房 = 9
    ROW_抗菌药物全年使用金额 = 10
    ROW_门诊西药房抗 = 11
    ROW_住院西药房抗 = 12
    ROW_抗菌药物占药品总收入比例 = 13
End Enum

Private Enum PATIREPORT_COLUMN '上报数据  病人抗菌用药情况抽样调查及评价表  rptPati
    COL_编辑 = 0
    COL_打印
    
    col_类型
    col_姓名
    col_性别
    col_年龄
    col_住院号
    col_科室
    col_住院医师
    col_出院日期
        
    col_病人Id
    col_主页ID
    COL_抽样ID
    COL_序号
    COL_手术ID
End Enum

Private Enum COL_CFADVICE '上报数据 门诊处方抗菌用药调查表 上半部分表格 vsCF / 统计数据 门急诊处方抗菌用药统计 上半部分表格 vsCountDruUse
    COL_CF序号 = 0
    COL_CF门诊号
    COL_CF病人姓名
    COL_CF就诊日期
    COL_CF处方医生
    COL_CF科室
    COL_CF病人年龄
    COL_CF诊断
    COL_CF药品品种数
    COL_CF基药品种数
    COL_CF注射剂
    COL_CF抗药品种数
    COL_CF通用名
    COL_CF规格
    COL_CF数量
    COL_CF金额
    COL_CF用法用量
    COL_CF用药途径
    COL_CF处方金额
    COL_CF药品金额
    COL_CF抗药金额
    
    COL_CF病人ID
    COL_CF挂号ID
    COL_CF挂号单
End Enum

Private Enum COL_YZADVICE '上报数据   住院病人抗菌用药调查表  vsZYYY
    COL_YZ类别 = 0
    COL_YZ药品通用名 = 1
    COL_YZ剂型 = 2
    COL_YZ规格 = 3
    COL_YZ单位 = 4
    COL_YZ数量 = 5
    COL_YZ总费用 = 6
End Enum

Private Enum COL_VSUSERAN_DRUG '统计数据  抗菌药物使用情况排名统计  vsUseRan 按药品方式汇总
    COL_D名称 = 0
    COL_D总金额
    COL_D使用例数
    COL_D患者人天数
    COL_D每例平均金额
    COL_DDDDs
    COL_D使用强度
    COL_D占药品金额比例
End Enum

Private Enum COL_VSUSERAN_NUDRUG '统计数据  抗菌药物使用情况排名统计  vsUseRan 按科室或医生方式汇总
    COL_UD类别 = 0
    COL_UD药品名称
    COL_UD剂型
    COL_UD规格
    COL_UD数量
    COL_UD总金额
    COL_UD使用例次
    COL_UD患者人天数
    COL_UD每例平均金额
    COL_UDDDDs
    COL_UD使用强度
    COL_UD占药品总金额比例
End Enum

Private Enum COL_VSCUT '统计数据  Ⅰ类切口围术期预防用药统计 vsCUT
    COL_CUT名称 = 0
    COL_CUT使用人次
    COL_CUT使用率
    COL_CUT切口数
    COL_CUT抗菌物数
    COL_CUT切口使用率
    COL_CUT术前用药
    COL_CUT平均用药
    COL_CUT品种数
End Enum

Private Enum COL_VSINDRUUSE '统计数据  住院医嘱抗菌用药统计  vsInDruUse
    COL_DRU序号 = 0
    COL_DRU住院号
    COL_DRU病人姓名
    COL_DRU出院日期
    COL_DRU主管医生
    COL_DRU科室
    COL_DRU出院诊断
    COL_DRU手术名称
    COL_DRU住院天数
    COL_DRU治疗金额
    COL_DRU日均治疗金额
    COL_DRU切口类型
    COL_DRU药品种数
    COL_DRU抗菌药物品种数
    COL_DRU药品金额
    COL_DRU抗菌药物金额
    COL_DRU联合用药
    
    COL_DRU药品名称
    COL_DRU剂型
    COL_DRU规格
    COL_DRU用法用量
    COL_DRU用药天数
    COL_DRU给药途径
    COL_DRU用药目的
    
    COL_DRU病人id
    COL_DRU主页id
End Enum

Private Enum COL_VSILLDRUUSE '统计数据  医生治疗某疾病抗菌用药成本统计  vsIllDruUse
    COL_ILL主管医生 = 0
    COL_ILL治疗人数
    COL_ILL用抗药人数
    COL_ILL治愈率
    COL_ILL总金额
    COL_ILL药品金额
    COL_ILL人均治疗额
    COL_ILL人均日金额
    COL_ILL抗药金额
    COL_ILL抗药品种数
    
    COL_ILL治愈
    COL_ILL好转
    COL_ILL未愈
    COL_ILL死亡
    COL_ILL其它
    
    COL_ILLID
End Enum

Private Enum COL_VSOPEKSSUSE '统计数据  术后抗菌药物使用超N天统计  vsOpeKssUse
    COL_OPE住院号 = 0
    COL_OPE患者姓名
    COL_OPE科室
    COL_OPE手术名称
    COL_OPE切口类型
    COL_OPE术后用药天数
    
    COL_OPE病人id
    COL_OPE主页id
End Enum

Private mstrPrivs As String
Private mlngModul As Long
Private mlngFind As Long
Private mdatCurr As Date
Private mstrMatch As String
Private mint简码 As Integer '简码匹配方式：0-拼音,1-五笔
Private mlng抽样ID As Long
Private mlng手术病人数 As Long
Private mlng非手术病人数 As Long
Private mlngRP门诊处方抽样数量 As Long '上报数据界面，处方抽样的实现数量
Private mlngOT门诊处方抽样数量 As Long '统计数据界面，处方抽样的实现数量
Private mbln病案查阅 As Boolean

Private Sub dtpRPS_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Dim tpGroupItem As TaskPanelGroupItem
    Dim i As Long
    Dim strCaption As String
    Dim strTmp As String, str诊断 As String
    Dim varArr As Variant
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    
    'TabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "上报数据统计", picReport.hwnd, 0).Tag = "上报数据统计"
        .InsertItem(1, "其他数据统计", picOther.hwnd, 0).Tag = "其他数据统计"
        .Item(0).Selected = True
    End With
    strCaption = "抗菌药品消耗金额调查表;(非)手术病人抗菌用药情况抽样调查及评价表;门诊处方抗菌用药调查表;住院病人抗菌用药调查表"
    With tbcReport
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        For i = 0 To UBound(Split(strCaption, ";"))
            .InsertItem(i, Split(strCaption, ";")(i), picReportSub(i).hwnd, 0).Tag = Split(strCaption, ";")(i)
        Next
    End With
    
    tbcReport.Item(tbcReport.ItemCount - 1).Selected = True
    tbcReport.Item(mEnumPanel.PanelItem_抗菌药品消耗金额调查表).Selected = True
    tbcReport.Item(mEnumPanel.PanelItem_病人抗菌用药情况抽样调查及评价表).Tag = "病人抗菌用药情况抽样调查及评价表" 'Tag保持一致
    
    strCaption = "抗菌药物使用情况排名统计;Ⅰ类切口围术期预防用药统计;门(急)诊处方抗菌用药统计;住院医嘱抗菌用药统计;术后抗菌药物使用超N天统计;医生治疗某疾病抗菌用药成本统计"
    With tbcOther
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        For i = 0 To UBound(Split(strCaption, ";"))
            .InsertItem(i, Split(strCaption, ";")(i), picOtherSub(i).hwnd, 0).Tag = Split(strCaption, ";")(i)
        Next
    End With
    tbcOther.Item(tbcOther.ItemCount - 1).Selected = True
    tbcOther.Item(mEnumPanel.PanelItem_抗菌药物使用情况排名统计).Selected = True
    tbcOther.Item(mEnumPanel.PanelItem_门急诊处方抗菌用药统计).Tag = "门急诊处方抗菌用药统计" 'Tag保持一致
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    mstrMatch = IIf(Val(zlDatabase.GetPara("输入匹配", , , True)) = 0, "%", "")
    str诊断 = zlDatabase.GetPara("治疗疾病编码", glngSys, 1269, "")
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    If InStr(str诊断, "|") > 0 Then
        varArr = Split(str诊断, "|")
        optType(e_C5_optType_西医_25).Value = Val(varArr(0)) = 1
        optType(e_C5_optType_按疾病_28).Value = Val(varArr(1)) = 1
        txtILL.Tag = Val(varArr(2))
        strTmp = varArr(0) & "|" & varArr(1) & "|" & varArr(2) & "|"
        txtILL.Text = Replace(str诊断, strTmp, "")
        cmdILL.Tag = txtILL.Text
    End If
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 1500
        .Add , "编码", "编码", 900
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    Call LoadDept
    
    Call InitReportColumn
    
    Call InitVS表格
    
    Call InitVS处方表格(vsMZYY, vsCF)
    Call InitVS处方表格(vsCountDruUse, vsCountCF)
    
    mdatCurr = zlDatabase.Currentdate
    Call InitTimeList
    Call Rest日期范围
    Call RestoreWinState(Me, App.ProductName)

End Sub

Private Sub Rest日期范围()
'功能：取上一次的日期范围从注册表中取
    Dim strTmp As String
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "抗菌药品消耗金额调查表", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R0_dtpRPS_开始时间_0).Value & "," & dtpRPE(e_R0_dtpRPE_结束时间_0).Value Then
            dtpRPS(e_R0_dtpRPS_开始时间_0).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R0_dtpRPE_结束时间_0).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R0_dtpRPS_开始时间_0), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "病人抗菌用药情况抽样调查及评价表", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R1_dtpRPS_开始时间_1).Value & "," & dtpRPE(e_R1_dtpRPE_结束时间_1).Value Then
            dtpRPS(e_R1_dtpRPS_开始时间_1).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R1_dtpRPE_结束时间_1).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R1_dtpRPS_开始时间_1), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "门诊处方抗菌用药调查表", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R2_dtpRPS_开始时间_2).Value & "," & dtpRPE(e_R2_dtpRPE_结束时间_2).Value Then
            dtpRPS(e_R2_dtpRPS_开始时间_2).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R2_dtpRPE_结束时间_2).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R2_dtpRPS_开始时间_2), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "住院病人抗菌用药调查表", "")
    If strTmp <> "" Then
        If strTmp <> dtpRPS(e_R3_dtpRPS_开始时间_3).Value & "," & dtpRPE(e_R3_dtpRPE_结束时间_3).Value Then
            dtpRPS(e_R3_dtpRPS_开始时间_3).Value = CDate(Split(strTmp, ",")(0))
            dtpRPE(e_R3_dtpRPE_结束时间_3).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimRP(e_R3_dtpRPS_开始时间_3), "自定义", False)
        End If
    End If
    
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "抗菌药物使用情况排名统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C0_dtpCountS_开始时间_0).Value & "," & dtpCountE(e_C0_dtpCountE_结束时间_0).Value Then
            dtpCountS(e_C0_dtpCountS_开始时间_0).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C0_dtpCountE_结束时间_0).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C0_dtpCountS_开始时间_0), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "Ⅰ类切口围术期预防用药统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C1_dtpCountS_开始时间_1).Value & "," & dtpCountE(e_C1_dtpCountE_结束时间_1).Value Then
            dtpCountS(e_C1_dtpCountS_开始时间_1).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C1_dtpCountE_结束时间_1).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C1_dtpCountS_开始时间_1), "自定义", False)
        End If
    End If
    
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "门急诊处方抗菌用药统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C2_dtpCountS_开始时间_2).Value & "," & dtpCountE(e_C2_dtpCountE_结束时间_2).Value Then
            dtpCountS(e_C2_dtpCountS_开始时间_2).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C2_dtpCountE_结束时间_2).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C2_dtpCountS_开始时间_2), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "住院医嘱抗菌用药统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C3_dtpCountS_开始时间_3).Value & "," & dtpCountE(e_C3_dtpCountE_结束时间_3).Value Then
            dtpCountS(e_C3_dtpCountS_开始时间_3).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C3_dtpCountE_结束时间_3).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C3_dtpCountS_开始时间_3), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "术后抗菌药物使用超N天统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C4_dtpCountS_开始时间_4).Value & "," & dtpCountE(e_C4_dtpCountE_结束时间_4).Value Then
            dtpCountS(e_C4_dtpCountS_开始时间_4).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C4_dtpCountE_结束时间_4).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C4_dtpCountS_开始时间_4), "自定义", False)
        End If
    End If
    
    strTmp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "医生治疗某疾病抗菌用药成本统计", "")
    If strTmp <> "" Then
        If strTmp <> dtpCountS(e_C5_dtpCountS_开始时间_5).Value & "," & dtpCountE(e_C5_dtpCountE_结束时间_5).Value Then
            dtpCountS(e_C5_dtpCountS_开始时间_5).Value = CDate(Split(strTmp, ",")(0))
            dtpCountE(e_C5_dtpCountE_结束时间_5).Value = CDate(Split(strTmp, ",")(1))
            Call Cbo.Locate(cboTimCount(e_C5_dtpCountS_开始时间_5), "自定义", False)
        End If
    End If
    
End Sub

Private Sub InitTimeList()
    Dim strDate As String
    Dim strTmp As String
    Dim lngTmp As Long
    Dim str最近一月 As String
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim rs年份 As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errH
    
    strDate = Format(mdatCurr, "yyyy-MM-dd hh:mm:ss")
    lngTmp = Val(Split(strDate, "-")(0)) - 1
    
    
    strSql = "Select Distinct Substr(期间, 1, 4)||'年' As 年份 From 期间表 Order By 年份"
    Set rs年份 = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    With cboTimRP(e_R0_cboTimRP_统计时间_0)
        .Clear
        For i = 1 To rs年份.RecordCount
            .AddItem rs年份!年份 & ""
            rs年份.MoveNext
        Next
        .AddItem "自定义"
    End With
    Call Cbo.Locate(cboTimRP(e_R0_cboTimRP_统计时间_0), lngTmp & "年", False)
    
    
    strDate = Format(DateAdd("m", -1, mdatCurr), "yyyy-MM-dd hh:mm:ss")
    
    strSql = "Select 开始日期, 终止日期 From 期间表 Where 期间 = [1]"
    strTmp = Split(strDate, "-")(0) & Split(strDate, "-")(1)
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp)
    
    With cboTimRP(e_R1_cboTimRP_统计时间_1)
        .Clear
        If Not rsTmp.EOF Then
            .Tag = Format(rsTmp!开始日期 & "", "yyyy-MM-dd") & "," & Format(rsTmp!终止日期 & "", "yyyy-MM-dd")
        End If
        .AddItem Split(strDate, "-")(0) & "年" & Split(strDate, "-")(1) & "月"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    With cboTimRP(e_R2_cboTimRP_统计时间_2)
        .Clear
        .Tag = cboTimRP(1).Tag
        .AddItem Split(strDate, "-")(0) & "年" & Split(strDate, "-")(1) & "月"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    '第一季度：1月－3月第二季度：4月－6月第三季度：7月－9月第四季度：10月－12月
    strDate = Format(DateAdd("m", -3, mdatCurr), "yyyy-MM-dd hh:mm:ss")
    lngTmp = Val(Split(strDate, "-")(1))
    strTmp = Split(strDate, "-")(0)
    If lngTmp >= 1 And lngTmp <= 3 Then
        cboTimRP(e_R3_cboTimRP_统计时间_3).Tag = strTmp & "-01-01," & strTmp & "-03-31"
        strTmp = strTmp & "年1季度"
    ElseIf lngTmp >= 4 And lngTmp <= 6 Then
        cboTimRP(e_R3_cboTimRP_统计时间_3).Tag = strTmp & "-04-01," & strTmp & "-06-30"
        strTmp = strTmp & "年2季度"
    ElseIf lngTmp >= 7 And lngTmp <= 9 Then
        cboTimRP(e_R3_cboTimRP_统计时间_3).Tag = strTmp & "-07-01," & strTmp & "-09-30"
        strTmp = strTmp & "年3季度"
    ElseIf lngTmp >= 10 And lngTmp <= 12 Then
        cboTimRP(e_R3_cboTimRP_统计时间_3).Tag = strTmp & "-10-01," & strTmp & "-12-31"
        strTmp = strTmp & "年4季度"
    End If
    
    With cboTimRP(e_R3_cboTimRP_统计时间_3)
        .Clear
        .AddItem strTmp
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    For i = 0 To 5
        With cboTimCount(i)
            .Clear
            .Tag = cboTimRP(e_R1_cboTimRP_统计时间_1).Tag
            .AddItem "最近一月"
            .AddItem "自定义"
            .ListIndex = 0
        End With
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpRPE_Change(Index As Integer)
    Call Cbo.Locate(cboTimRP(Index), "自定义", False)
End Sub

Private Sub dtpRPS_Change(Index As Integer)
    Call Cbo.Locate(cboTimRP(Index), "自定义", False)
End Sub

Private Sub dtpCountE_Change(Index As Integer)
    Call Cbo.Locate(cboTimCount(Index), "自定义", False)
End Sub

Private Sub dtpCountS_Change(Index As Integer)
    Call Cbo.Locate(cboTimCount(Index), "自定义", False)
End Sub

Private Sub cboTimCount_Click(Index As Integer)
    With cboTimCount(Index)
        If .ListIndex = 0 And .Tag <> "" Then
            dtpCountS(Index).Value = CDate(Split(.Tag, ",")(0))
            dtpCountE(Index).Value = CDate(Split(.Tag, ",")(1))
        End If
    End With
End Sub

Private Sub cboTimRP_Click(Index As Integer)
    Dim strTmp As String
    Dim lngYear As Long
    
    '取当前日期的年份
    strTmp = Format(mdatCurr, "yyyy-MM-dd")
    lngYear = Val(Split(strTmp, "-")(0)) - 1
    
    Select Case Index
    Case e_R0_cboTimRP_统计时间_0
        With cboTimRP(Index)
            If .Text <> "自定义" Then
                strTmp = Mid(.Text, 1, 4) & "-01-01"
                dtpRPS(Index).Value = CDate(strTmp)
                
                strTmp = Mid(.Text, 1, 4) & "-12-31"
                dtpRPE(Index).Value = CDate(strTmp)
            End If
        End With
    Case Else
        With cboTimRP(Index)
            If .ListIndex = 0 And .Tag <> "" Then
                dtpRPS(Index).Value = CDate(Split(.Tag, ",")(0))
                dtpRPE(Index).Value = CDate(Split(.Tag, ",")(1))
            End If
        End With
    End Select
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        Set objCol = .Columns.Add(COL_编辑, "编辑", 30, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_打印, "打印", 30, False)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentCenter
       
        Set objCol = .Columns.Add(col_类型, "类型", 50, False)
            objCol.Visible = False
  
        Set objCol = .Columns.Add(col_姓名, "姓名", 600, True)
            objCol.Groupable = False
            
        Set objCol = .Columns.Add(col_性别, "性别", 400, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_年龄, "年龄", 400, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_住院号, "住院号", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_科室, "科室", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_住院医师, "住院医师", 600, True)
            objCol.Alignment = xtpAlignmentLeft
            
        Set objCol = .Columns.Add(col_出院日期, "出院日期", 800, True)
            objCol.Alignment = xtpAlignmentCenter
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_类型
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(col_类型)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(col_类型)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_出院日期)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")

        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "电子病案查阅(&I)")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With


    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")


        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "病案")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With
    
    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs, "ZL1_INSIDE_1269_1", "ZL1_INSIDE_1269_2")
End Sub

Private Sub InitTable(ByRef vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitVS表格()
'功能：初始化VS表格，固定列的设置等
    Dim strTmp As String
    Dim i As Integer
    
    '上报部分 抗菌药品消耗金额调查表
    strTmp = "统计项目,3570,1;结果,1900,7;单位,740,4;备注,2890,1"
    Call InitTable(vsBill, strTmp)
    
    '表格第一列数据
    With vsBill
        .Rows = .FixedRows
        .Rows = 14
        .RowHeight(0) = 450
        For i = 1 To 13
            .RowHeight(i) = 450
        Next
        .TextMatrix(ROW_年医院总收入, COL_统计项目) = "一、年医院总收入（金额）"
        .TextMatrix(ROW_政府拨款, COL_统计项目) = "二、政府拨款（金额）"
        .TextMatrix(ROW_年药品总收入, COL_统计项目) = "三、年药品总收入（金额）"
        .TextMatrix(ROW_药品占医院总收入比例, COL_统计项目) = "四、药品占医院总收入比例"
        .TextMatrix(ROW_药品进销差价收入, COL_统计项目) = "五、药品进销差价收入（金额）"
        .TextMatrix(ROW_药品进销差价收入占医院总收入比例, COL_统计项目) = "六、药品进销差价收入占医院总收入比例"
        .TextMatrix(ROW_西药全年使用金额, COL_统计项目) = "七、西药全年使用金额（零售价）"
        .TextMatrix(ROW_门诊西药房, COL_统计项目) = "    其中：门诊西药房"
        .TextMatrix(ROW_住院西药房, COL_统计项目) = "          住院西药房"
        .TextMatrix(ROW_抗菌药物全年使用金额, COL_统计项目) = "八、抗菌药物全年使用金额（零售价）"
        .TextMatrix(ROW_门诊西药房抗, COL_统计项目) = "    其中：门诊西药房"
        .TextMatrix(ROW_住院西药房抗, COL_统计项目) = "          住院西药房"
        .TextMatrix(ROW_抗菌药物占药品总收入比例, COL_统计项目) = "九、抗菌药物占药品总收入比例"
    End With
    
    
    '上报部分 住院病人抗菌用药调查表
    strTmp = "类别,1500,4;药品通用名,2200,1;剂型,1600,4;规格,2800,1;单位,700,4;数量,800,7;总费用(元),1200,7"
    Call InitTable(vsZYYY, strTmp)
    
    '统计部分 初始化统计部分表格，  抗菌药物使用情况排名统计  界面
    strTmp = "类别,3100,4;药品名称,3600,4;剂型,2000,4;规格,2800,4;数量,1020,4;总金额(元),1200,4;使用例次,530,4;每例平均金额(元),760,4;DDDs,500,4;使用强度,550,4;占药品总金额比例(%),780,4"
    Call InitTable(vsUseRan, strTmp)
    
    '统计部分 初始化 Ⅰ类切口围术期预防用药统计 界面表格－－vsCut
    With vsCut
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(0) = True
        .Cols = 9
        For i = COL_CUT名称 To COL_CUT品种数
            .ColAlignment(i) = flexAlignCenterCenter
            .ColWidth(i) = 2000
        Next
        
        .Cell(flexcpText, 0, 0, 1, 0) = "科室名称"
        .Cell(flexcpText, 0, 1, 0, 2) = "抗菌药物使用情况"
        .Cell(flexcpText, 0, COL_CUT切口数, 0, COL_CUT品种数) = "Ⅰ类切口围术期预防使用抗菌药物情况"
        .TextMatrix(1, COL_CUT使用人次) = "使用人次(例)"
        .TextMatrix(1, COL_CUT使用率) = "使用率(%)"
        .TextMatrix(1, COL_CUT切口数) = "Ⅰ类切口数(例)"
        .TextMatrix(1, COL_CUT抗菌物数) = "抗菌药预防使用数(例)"
        .TextMatrix(1, COL_CUT切口使用率) = "使用率(%)"
        .TextMatrix(1, COL_CUT术前用药) = "术前用药数"
        .TextMatrix(1, COL_CUT平均用药) = "平均用药天数"
        .TextMatrix(1, COL_CUT品种数) = "平均用药品种数"
    End With
    
    '统计部分 加载   住院医嘱抗菌用药统计  界面表格初始化
    With vsInDruUse
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .Cols = 26
        For i = COL_DRU序号 To COL_DRU用药目的
            .ColAlignment(i) = flexAlignLeftCenter
            .MergeCol(i) = True
        Next
        .ColWidth(COL_DRU序号) = 450
        .ColAlignment(COL_DRU序号) = flexAlignCenterCenter
        
        .ColWidth(COL_DRU住院号) = 800
        .ColWidth(COL_DRU病人姓名) = 930
        .ColWidth(COL_DRU主管医生) = .ColWidth(COL_DRU病人姓名)
        .ColWidth(COL_DRU科室) = 1400
        .ColWidth(COL_DRU给药途径) = .ColWidth(COL_DRU病人姓名)
        
        .ColWidth(COL_DRU出院日期) = 990
        .ColAlignment(COL_DRU出院日期) = flexAlignCenterCenter
        
        .ColWidth(COL_DRU出院诊断) = 2300
        
        .ColWidth(COL_DRU住院天数) = 600
        .ColAlignment(COL_DRU住院天数) = flexAlignRightCenter
        
        .ColWidth(COL_DRU治疗金额) = 1060
        .ColAlignment(COL_DRU治疗金额) = flexAlignRightCenter
        
        .ColWidth(COL_DRU日均治疗金额) = 900
        .ColAlignment(COL_DRU日均治疗金额) = flexAlignRightCenter
        
        .ColWidth(COL_DRU切口类型) = 450
        
        .ColWidth(COL_DRU药品种数) = .ColWidth(COL_DRU序号)
        .ColAlignment(COL_DRU药品种数) = flexAlignRightCenter
        
        .ColWidth(COL_DRU抗菌药物品种数) = .ColWidth(COL_DRU序号)
        .ColAlignment(COL_DRU抗菌药物品种数) = flexAlignRightCenter
        
        .ColWidth(COL_DRU药品金额) = 1080
        .ColAlignment(COL_DRU药品金额) = flexAlignRightCenter
        
        .ColWidth(COL_DRU抗菌药物金额) = 1035
        .ColAlignment(COL_DRU抗菌药物金额) = flexAlignRightCenter
        
        .ColWidth(COL_DRU联合用药) = 820
        .ColWidth(COL_DRU药品名称) = 2500
        .ColWidth(COL_DRU剂型) = 700
        .ColAlignment(COL_DRU剂型) = flexAlignCenterCenter
        
        .ColWidth(COL_DRU规格) = 2800
        .ColWidth(COL_DRU用法用量) = 1260
        .ColWidth(COL_DRU用药天数) = 450
        .ColAlignment(COL_DRU用药天数) = flexAlignRightCenter
        
        .ColWidth(COL_DRU给药途径) = 1125
        
        .ColWidth(COL_DRU用药目的) = 600
        .ColAlignment(COL_DRU用药目的) = flexAlignCenterCenter
        
        .Cell(flexcpText, 0, COL_DRU序号, 1, COL_DRU序号) = "序号"
        .Cell(flexcpText, 0, COL_DRU住院号, 1, COL_DRU住院号) = "住院号"
        .Cell(flexcpText, 0, COL_DRU病人姓名, 1, COL_DRU病人姓名) = "病人姓名"
        .Cell(flexcpText, 0, COL_DRU出院日期, 1, COL_DRU出院日期) = "出院日期"
        .Cell(flexcpText, 0, COL_DRU主管医生, 1, COL_DRU主管医生) = "主管医生"
        .Cell(flexcpText, 0, COL_DRU科室, 1, COL_DRU科室) = "科室"
        .Cell(flexcpText, 0, COL_DRU出院诊断, 1, COL_DRU出院诊断) = "出院诊断"
        .Cell(flexcpText, 0, COL_DRU手术名称, 1, COL_DRU手术名称) = "手术名称"
        .Cell(flexcpText, 0, COL_DRU住院天数, 1, COL_DRU住院天数) = "住院天数"
        .Cell(flexcpText, 0, COL_DRU治疗金额, 1, COL_DRU治疗金额) = "治疗金额(元)"
        .Cell(flexcpText, 0, COL_DRU日均治疗金额, 1, COL_DRU日均治疗金额) = "日均治疗金额(元)"
        .Cell(flexcpText, 0, COL_DRU切口类型, 1, COL_DRU切口类型) = "切口类型"
        .Cell(flexcpText, 0, COL_DRU药品种数, 1, COL_DRU药品种数) = "药品种数"
        .Cell(flexcpText, 0, COL_DRU抗菌药物品种数, 1, COL_DRU抗菌药物品种数) = "抗菌药物品种数"
        .Cell(flexcpText, 0, COL_DRU药品金额, 1, COL_DRU药品金额) = "药品金额(元)"
        .Cell(flexcpText, 0, COL_DRU抗菌药物金额, 1, COL_DRU抗菌药物金额) = "抗菌药物金额(元)"
        .Cell(flexcpText, 0, COL_DRU联合用药, 1, COL_DRU联合用药) = "联合用药"
        
        .Cell(flexcpText, 0, COL_DRU药品名称, 0, COL_DRU用药目的) = "抗菌药使用情况(具体用法)"
        
        .TextMatrix(1, COL_DRU药品名称) = "药品名称"
        .TextMatrix(1, COL_DRU剂型) = "剂型"
        .TextMatrix(1, COL_DRU规格) = "规格"
        .TextMatrix(1, COL_DRU用法用量) = "用法用量"
        .TextMatrix(1, COL_DRU用药天数) = "天数"
        .TextMatrix(1, COL_DRU给药途径) = "给药途径"
        .TextMatrix(1, COL_DRU用药目的) = "目的"
        
        .ColHidden(COL_DRU病人id) = True
        .ColHidden(COL_DRU主页id) = True
        
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, COL_DRU用药目的) = flexAlignCenterCenter
        
        .WordWrap = True
    End With
    
    With vsInDruAna
        .Clear
        .Cols = 4: .Rows = 0
        .RowHeightMin = 300
        For i = 0 To .Cols - 1
            .ColWidth(i) = 5000
        Next
        .WordWrap = True
        .AddItem "" '第0行
        .TextMatrix(.Rows - 1, 0) = "A(用药总品种数)=0种"
        .TextMatrix(.Rows - 1, 1) = "B(平均用药品种数A/C)=0种"
        .TextMatrix(.Rows - 1, 2) = "C(使用抗菌药物的品种数)=0种"
        .TextMatrix(.Rows - 1, 3) = "D(使用抗菌药物的百分率C/A)*100%=0%"
        
        .AddItem "" '第1行
        .TextMatrix(.Rows - 1, 0) = "E(使用抗药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 1) = "F(出院病人抗菌药物使用率E/实际人数)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "G(治疗总金额)=0元"
        .TextMatrix(.Rows - 1, 3) = "H(病人平均治疗金额G/实际人数)=0元"
        
        .AddItem "" '第2行
        .TextMatrix(.Rows - 1, 0) = "I(药品总金额)=0元"
        .TextMatrix(.Rows - 1, 1) = "L(药品总金额占治疗总金额的百分率I/G)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "K(抗菌药物总金额)=0元"
        .TextMatrix(.Rows - 1, 3) = "J(抗菌药物总金额占药品总金额的百分率K/I)*100%=0%"
        
        .AddItem "" '第3行
        .TextMatrix(.Rows - 1, 0) = "M(单用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 1) = "O(单用抗菌药物的使用率M/E)*100%=0%"
        .TextMatrix(.Rows - 1, 2) = "P(二联使用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 3) = "Q(二联使用抗菌药物的使用率P/E)*100%＝0%"
        
        .AddItem "" '第4行
        .TextMatrix(.Rows - 1, 0) = "R(三联使用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 1) = "S(三联使用抗菌药物的使用率R/E)*100%＝0%"
        .TextMatrix(.Rows - 1, 2) = "T(四联使用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 3) = "U(四联使用抗菌药物的使用率T/E)*100%＝0%"
        
        .AddItem "" '第5行
        .TextMatrix(.Rows - 1, 0) = "V(预防使用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 1) = "W(预防使用抗菌药物构成比V/E)100%=0%"
        .TextMatrix(.Rows - 1, 2) = "X(治疗使用抗菌药物的病人数)=0例"
        .TextMatrix(.Rows - 1, 3) = "Y(治疗使用抗菌药物构成比Y/E)*100%=0%"
        
    End With
    
    '统计部分 术后抗菌药物使用超N天统计 界面 vsOpeKssUse
    strTmp = "住院号,1500,1;患者姓名,2000,4;科室,2000,4;手术名称,5000,1;切口类型,1000,4;术后用药天数,2000,4;病人id;主页id"
    Call InitTable(vsOpeKssUse, strTmp)
    
    '统计部分   医生治疗某疾病抗菌用药成本统计 界面， vsIllDruUse
    With vsIllDruUse
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .Cols = 16
        For i = COL_ILL主管医生 To COL_ILL其它
            .ColAlignment(i) = flexAlignRightCenter
            .MergeCol(i) = True
        Next
        .ColAlignment(COL_ILL主管医生) = flexAlignCenterCenter
        .ColHidden(COL_ILLID) = True
        .ColWidth(COL_ILL主管医生) = 1500
        .ColWidth(COL_ILL治疗人数) = 1270
        .ColWidth(COL_ILL用抗药人数) = 1270
        .ColWidth(COL_ILL治愈率) = 1070
        
        .ColWidth(COL_ILL总金额) = 1080
        .ColWidth(COL_ILL药品金额) = 1080
        .ColWidth(COL_ILL人均治疗额) = 1300
        .ColWidth(COL_ILL人均日金额) = 1300
        .ColWidth(COL_ILL抗药金额) = 1300
        .ColWidth(COL_ILL抗药品种数) = 830
        .ColWidth(COL_ILL治愈) = 550
        .ColWidth(11) = 550
        .ColWidth(12) = 550
        .ColWidth(13) = 550
        .ColWidth(14) = 550
        
        .Cell(flexcpText, 0, COL_ILL主管医生, 1, COL_ILL主管医生) = "主管医生"
        .Cell(flexcpText, 0, COL_ILL治疗人数, 1, COL_ILL治疗人数) = "治疗人数"
        .Cell(flexcpText, 0, COL_ILL用抗药人数, 1, COL_ILL用抗药人数) = "使用抗菌药物的人数"
        .Cell(flexcpText, 0, COL_ILL治愈率, 1, COL_ILL治愈率) = "治愈率(%)"
        .Cell(flexcpText, 0, COL_ILL总金额, 1, COL_ILL总金额) = "总金额(元)"
        .Cell(flexcpText, 0, COL_ILL药品金额, 1, COL_ILL药品金额) = "药品总金额(元)"
        .Cell(flexcpText, 0, COL_ILL人均治疗额, 1, COL_ILL人均治疗额) = "人均治疗金额(元)"
        .Cell(flexcpText, 0, COL_ILL人均日金额, 1, COL_ILL人均日金额) = "人均日金额(元)"
        .Cell(flexcpText, 0, COL_ILL抗药金额, 1, COL_ILL抗药金额) = "抗菌药物总金额(元)"
        .Cell(flexcpText, 0, COL_ILL抗药品种数, 1, COL_ILL抗药品种数) = "抗菌药物品种数"
        
        .Cell(flexcpText, 0, COL_ILL治愈, 0, COL_ILL其它) = "治疗结果"
        
        .TextMatrix(1, COL_ILL治愈) = "治愈"
        .TextMatrix(1, COL_ILL好转) = "好转"
        .TextMatrix(1, COL_ILL未愈) = "未愈"
        .TextMatrix(1, COL_ILL死亡) = "死亡"
        .TextMatrix(1, COL_ILL其它) = "其它"
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub InitVS处方表格(ByRef vsgInfo1 As VSFlexGrid, ByRef vsgInfo2 As VSFlexGrid)
'功能：初始化处方表格表头，上报部门的门诊处方，上报数据和统计数据共用
'参数： vsgInfo2 明细表格 上报数据－－vsMZYY   统计 －－ vsCountDruUse；   vsgInfo2 分析表格  上报数据－－  vsCF   统计 －－  vsCountCF
    Dim i As Integer
 
    With vsgInfo1
        .Clear
        .FixedRows = 2: .FixedCols = 0
        .Rows = 3: .Cols = 24
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        For i = COL_CF序号 To COL_CF抗药金额
            .ColAlignment(i) = flexAlignLeftCenter '先统一靠左对齐
            .MergeCol(i) = True
        Next
        .ColWidth(COL_CF序号) = 450
        .ColAlignment(COL_CF序号) = flexAlignCenterCenter
        
        .ColWidth(COL_CF门诊号) = 1110
        .ColWidth(COL_CF病人姓名) = 1050
        .ColWidth(COL_CF就诊日期) = 1080
        .ColAlignment(COL_CF就诊日期) = flexAlignCenterCenter
        
        .ColWidth(COL_CF处方医生) = 900
        .ColWidth(COL_CF科室) = 1400
        .ColWidth(COL_CF病人年龄) = 540
        .ColWidth(COL_CF诊断) = 2210
        .ColWidth(COL_CF药品品种数) = 480
        .ColAlignment(COL_CF药品品种数) = flexAlignRightCenter
        .ColWidth(COL_CF基药品种数) = 480
        .ColAlignment(COL_CF基药品种数) = flexAlignRightCenter
        .ColWidth(COL_CF注射剂) = 630
        .ColAlignment(COL_CF注射剂) = flexAlignCenterCenter
        
        .ColWidth(COL_CF抗药品种数) = 480
        .ColAlignment(COL_CF抗药品种数) = flexAlignRightCenter
        
        .ColWidth(COL_CF通用名) = 2310
        .ColWidth(COL_CF规格) = 2800
        .ColWidth(COL_CF数量) = 615
        .ColAlignment(COL_CF数量) = flexAlignRightCenter
        
        .ColWidth(COL_CF金额) = 870
        .ColAlignment(COL_CF金额) = flexAlignRightCenter
        .ColWidth(COL_CF用法用量) = 1260
        .ColWidth(COL_CF用药途径) = 1020
        .ColWidth(COL_CF处方金额) = 780
        .ColAlignment(COL_CF处方金额) = flexAlignRightCenter
        .ColWidth(COL_CF药品金额) = 780
        .ColAlignment(COL_CF药品金额) = flexAlignRightCenter
        .ColWidth(COL_CF抗药金额) = 840
        .ColAlignment(COL_CF抗药金额) = flexAlignRightCenter
        
        .ColHidden(COL_CF病人ID) = True
        .ColHidden(COL_CF挂号ID) = True
        .ColHidden(COL_CF挂号单) = True
        
        .Cell(flexcpText, 0, COL_CF序号, 1, COL_CF序号) = "序号"
        .Cell(flexcpText, 0, COL_CF门诊号, 1, COL_CF门诊号) = "门诊号"
        .Cell(flexcpText, 0, COL_CF病人姓名, 1, COL_CF病人姓名) = "病人姓名"
        .Cell(flexcpText, 0, COL_CF就诊日期, 1, COL_CF就诊日期) = "就诊日期"
        .Cell(flexcpText, 0, COL_CF处方医生, 1, COL_CF处方医生) = "处方医生"
        .Cell(flexcpText, 0, COL_CF科室, 1, COL_CF科室) = "科室"
        .Cell(flexcpText, 0, COL_CF病人年龄, 1, COL_CF病人年龄) = "年龄"
        .Cell(flexcpText, 0, COL_CF诊断, 1, COL_CF诊断) = "诊断"
        .Cell(flexcpText, 0, COL_CF药品品种数, 1, COL_CF药品品种数) = "药品种数"
        .Cell(flexcpText, 0, COL_CF基药品种数, 1, COL_CF基药品种数) = "基本药种数"
        .Cell(flexcpText, 0, COL_CF注射剂, 1, COL_CF注射剂) = "注射剂有/无"
        .Cell(flexcpText, 0, COL_CF抗药品种数, 1, COL_CF抗药品种数) = "抗菌药种数"
        
        .Cell(flexcpText, 0, COL_CF通用名, 0, COL_CF用药途径) = "抗菌药物使用情况(具体用法)"
    
        .TextMatrix(1, COL_CF通用名) = "通用名"
        .TextMatrix(1, COL_CF规格) = "规格"
        .TextMatrix(1, COL_CF数量) = "数量"
        .TextMatrix(1, COL_CF金额) = "金额(元)"
        .TextMatrix(1, COL_CF用法用量) = "用法用量"
        .TextMatrix(1, COL_CF用药途径) = "用药途径"
        .Cell(flexcpText, 0, COL_CF处方金额, 1, COL_CF处方金额) = "处方金额(元)"
        .Cell(flexcpText, 0, COL_CF药品金额, 1, COL_CF药品金额) = "药品金额(元)"
        .Cell(flexcpText, 0, COL_CF抗药金额, 1, COL_CF抗药金额) = "抗菌药物金额(元)"
        
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, COL_CF抗药金额) = flexAlignCenterCenter
        
        .WordWrap = True
    End With
    
    With vsgInfo2
        .Clear
        .Cols = 4: .Rows = 0
        .RowHeightMin = 300
        For i = 0 To .Cols - 1
            .ColWidth(i) = 5000
        Next
        
        .AddItem "" '第一行
        .TextMatrix(.Rows - 1, 0) = "A(处方用药总品种数)=0种"
        .TextMatrix(.Rows - 1, 1) = "B(平均用药品种数A/处方数)=0种"
        .TextMatrix(.Rows - 1, 2) = "C(使用抗菌药物的品种数)=0种"
        .TextMatrix(.Rows - 1, 3) = "D(就诊使用抗菌药物的百分率C/A*100%)=0%"
        
        .AddItem "" '第二行
        .TextMatrix(.Rows - 1, 0) = "E(使用注射剂的处方数)=0张"
        .TextMatrix(.Rows - 1, 1) = "F(就诊使用注射剂的处方的百分率E/处方数*100%)=0%"
        .TextMatrix(.Rows - 1, 2) = "G(使用抗菌药物的处方数)=0张"
        .TextMatrix(.Rows - 1, 3) = "H(就诊使用抗菌药物的处方的百分率 G/100)=0%"
                
        .AddItem "" '第三行
        .TextMatrix(.Rows - 1, 0) = "I(处方总金额)=0元"
        .TextMatrix(.Rows - 1, 1) = "J(处方平均金额 I/100)=0元"
        .TextMatrix(.Rows - 1, 2) = "K(使用抗菌药物的总金额)=0元"
        .TextMatrix(.Rows - 1, 3) = "L(抗菌药物总金额占处方总金额的比率 K/I)=0%"
        
        .AddItem "" '第四行
        .TextMatrix(.Rows - 1, 0) = "M(使用抗菌药物的处方总金额)=0元"
        .TextMatrix(.Rows - 1, 1) = "N(每张抗菌药处方平均金额 M/G)=0元"
        .TextMatrix(.Rows - 1, 2) = "O(使用基本药物的品种数)=0种"
        .TextMatrix(.Rows - 1, 3) = "P(就诊使用基本药物的百分率 O/A)=0%"
        
        .AddItem "" '第五行
        .TextMatrix(.Rows - 1, 0) = "Q(使用抗菌药物的处方总数)=0张"
        .TextMatrix(.Rows - 1, 1) = "R(就诊使用抗菌药物处方的百分率 Q/100)=0%"
    End With
End Sub

Private Sub Load处方分析(ByRef vsgInfo1 As VSFlexGrid, ByRef vsgInfo2 As VSFlexGrid, ByVal blnRP As Boolean)
'功能：汇总处方统计分析数据
'参数： vsgInfo1 表格对象处方抽样，上报数据－－vsMZYY  其它统计 vsCountDruUse ；
'       vsgInfo2 处方评价                      vsCF             vsCountCF
'       blnRP 界面，true 上报数据，false 其它统计
    Dim lng处方总数 As Long
    Dim lng抗菌药总数 As Long
    Dim lng注射剂总数 As Long
    Dim lng使用抗菌药处方数 As Long
    Dim dbl处方总金额 As Double
    Dim dbl使用抗菌药金额 As Double
    Dim dbl使用抗菌药处方金额 As Double
    Dim lng基本药总数 As Long
    
    
    Dim strDec As String
    Dim lng实际数量 As Long
    Dim lng总药品数 As Long
    Dim strTmp As String
    Dim dblTmp As Double
    Dim lngTmp As Long
    
    Dim i As Long
    
    strDec = "0.00"
    
    With vsgInfo1
        '取实际处方数，抽样100张但实际可能不到100张，
        For i = .Rows - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_CF序号)) <> 0 Then
                lng实际数量 = Val(.TextMatrix(i, COL_CF序号))
                Exit For
            End If
        Next
        
        If lng实际数量 = 0 Then Exit Sub
        
        For i = .FixedRows To .Rows - 1
            'A(处方用药总品种数)  每个处方用的药品数求和
            lng总药品数 = lng总药品数 + Val(.TextMatrix(i, COL_CF药品品种数))
            
            'C(使用抗菌药物的品种数)  每个处方用的抗菌药药品数求和
            lng抗菌药总数 = lng抗菌药总数 + Val(.TextMatrix(i, COL_CF抗药品种数))
            
            If .TextMatrix(i, COL_CF注射剂) = "有" Then lng注射剂总数 = lng注射剂总数 + 1
            
            If Val(.TextMatrix(i, COL_CF抗药品种数)) <> 0 Then
                lng使用抗菌药处方数 = lng使用抗菌药处方数 + 1
                dbl使用抗菌药处方金额 = dbl使用抗菌药处方金额 + Val(.TextMatrix(i, COL_CF处方金额))
            End If
            
            dbl处方总金额 = dbl处方总金额 + Val(.TextMatrix(i, COL_CF处方金额))
            dbl使用抗菌药金额 = dbl使用抗菌药金额 + Val(.TextMatrix(i, COL_CF抗药金额))
            lng基本药总数 = lng基本药总数 + Val(.TextMatrix(i, COL_CF基药品种数))
        Next
    End With
    
    With vsgInfo2
'        .AddItem "" '第一行
        .TextMatrix(lngTmp, 0) = "A(处方用药总品种数)=" & lng总药品数 & "种"
        
        dblTmp = lng总药品数 / lng实际数量
        .TextMatrix(lngTmp, 1) = "B(平均用药品种数 A/" & lng实际数量 & ")=" & Format(dblTmp, strDec) & "种": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "C(使用抗菌药物的品种数)=" & lng抗菌药总数 & "种"
    
        If lng总药品数 <> 0 Then dblTmp = lng抗菌药总数 * 100 / lng总药品数
        .TextMatrix(lngTmp, 3) = "D(就诊使用抗菌药物的百分率 C/A)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '第二行
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "E(使用注射剂的处方数)=" & lng注射剂总数 & "张"
        
        dblTmp = lng注射剂总数 * 100 / lng实际数量
        .TextMatrix(lngTmp, 1) = "F(就诊使用注射剂的处方的百分率 E/" & lng实际数量 & ")=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        .TextMatrix(lngTmp, 2) = "G(使用抗菌药物的处方数)=" & lng使用抗菌药处方数 & "张"
        
        dblTmp = lng使用抗菌药处方数 * 100 / lng实际数量
        .TextMatrix(lngTmp, 3) = "H(就诊使用抗菌药物的处方的百分率 G/" & lng实际数量 & ")=" & Format(dblTmp, strDec) & "%": dblTmp = 0
                
'        .AddItem "" '第三行
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "I(处方总金额)=" & Format(dbl处方总金额, strDec) & "元"
        
        dblTmp = dbl处方总金额 / lng实际数量
        .TextMatrix(lngTmp, 1) = "J(处方平均金额 I/" & lng实际数量 & ")=" & Format(dblTmp, strDec) & "元": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "K(使用抗菌药物的总金额)=" & Format(dbl使用抗菌药金额, strDec) & "元"
        
        If dbl处方总金额 <> 0 Then dblTmp = dbl使用抗菌药金额 / dbl处方总金额
        
        .TextMatrix(lngTmp, 3) = "L(抗菌药物总金额占处方总金额的比率 K/I)=" & Format(dblTmp, strDec): dblTmp = 0
        
'        .AddItem "" '第四行
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "M(使用抗菌药物的处方总金额)=" & Format(dbl使用抗菌药处方金额, strDec) & "元"
        
        If lng使用抗菌药处方数 <> 0 Then dblTmp = dbl使用抗菌药处方金额 / lng使用抗菌药处方数
        .TextMatrix(lngTmp, 1) = "N(每张抗菌药处方平均金额 M/G)=" & Format(dblTmp, strDec) & "元": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "O(使用基本药物的品种数)=" & lng基本药总数 & "种"
        
        If lng总药品数 <> 0 Then dblTmp = lng基本药总数 * 100 / lng总药品数
        .TextMatrix(lngTmp, 3) = "P(就诊使用基本药物的百分率 O/A * 100%)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '第五行
        lngTmp = lngTmp + 1
        .TextMatrix(lngTmp, 0) = "Q(使用抗菌药物的处方总数)=" & lng使用抗菌药处方数 & "张"
        
        dblTmp = 100 * lng使用抗菌药处方数 / lng实际数量
        .TextMatrix(lngTmp, 1) = "R(就诊使用抗菌药物处方的百分率 Q/100)=" & Format(dblTmp, strDec) & "%": dblTmp = 0
    End With

End Sub

Private Sub LoadvsUseRan()
'功能：加载   抗菌药物使用情况排名统计  界面表格
    Dim strSql As String, strPar1 As String, strPar2 As String
    Dim strSQLDetail As String
    Dim rsDetail As ADODB.Recordset
    Dim strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim strDec As String
    Dim strTmp As String
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim lng患者人天数 As Long
    Dim lng收治总人次 As Long
    Dim dbl总药费 As Double
    Dim i As Long
    Dim rs出院病人 As ADODB.Recordset
    Dim strSQL出院病人 As String
    Dim dat出院开始 As Date
    Dim dat出院结束 As Date
    Dim strWhere病人 As String
    Dim lng人数 As Long
    Dim lng天数 As Long

    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    strDec = "0.00"
    
    '日期范围参数生成
    strPar1 = "To_Date('" & Format(dtpCountS(e_C0_dtpCountS_开始时间_0).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
    strPar2 = "To_Date('" & Format(dtpCountE(e_C0_dtpCountE_结束时间_0).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C0_txtDept_抽样科室_0).Tag = "", "", " and a.开单部门id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    
    If optType(e_C0_optType_统计场合_住院_5).Value Then
        '统计场合 住院
        If optType(e_C0_optType_汇总方式_科室_9).Value Then '汇总方式 科室
            strSql = "Select /*+ rule*/  m.名称,m.开单部门id, m.Ddds, m.费用, m.人数 as 用例数, round(m.天数) as 天数" & vbNewLine & _
                "From (Select n.名称,m.开单部门id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Sum(m.天数) As 天数, Count(1) As 人数" & vbNewLine & _
                "       From (Select m.开单部门id, m.病人id, m.主页id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Nvl(g.出院日期,Sysdate) -g.入院日期 As 天数" & vbNewLine & _
                "              From (Select a.开单部门id,a.病人id, a.主页id, Sum(a.结帐金额) As 费用," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds" & vbNewLine & _
                "                     From 住院费用记录 A, 药品规格 B,药品特性 D" & vbNewLine & _
                "                     Where a.收费类别 = '5' And a.记录状态 <> 0 And  a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.收费细目id = b.药品id And b.药名id = d.药名id And Nvl(d.抗生素, 0) <> 0" & vbNewLine & _
                "                     Group By a.病人id, a.主页id,a.开单部门id) M, 病案主页 G" & vbNewLine & _
                "          where  g.病人id = m.病人id And g.主页id = m.主页id  Group By m.病人id, m.主页id, m.开单部门id,g.出院日期,g.入院日期) M,部门表 n" & vbNewLine & _
                "       where m.开单部门id=n.id Group By m.开单部门id,n.名称 having Sum(m.费用) >0 " & vbNewLine & _
                "       Order By Sum(m.费用) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            strSQLDetail = "select a.开单部门id as id,sum(a.结帐金额) as 总费用 from 住院费用记录 a" & vbNewLine & _
                "where a.收费类别 in ('5','6','7') and a.记录状态 <> 0 " & vbNewLine & _
                "and a.发生时间 between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "and a.开单部门id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                "group by a.开单部门id having sum(a.结帐金额)>0"
        ElseIf optType(e_C0_optType_汇总方式_医生_8).Value Then
            '汇总方式 医生
            strSql = "Select /*+ rule*/  nvl(m.开单人,'空') as 开单人, m.Ddds, m.费用, m.人数 as 用例数,round(m.天数) as 天数" & vbNewLine & _
                "From (Select m.开单人, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Count(1) As 人数, Sum(m.天数) As 天数" & vbNewLine & _
                "       From (Select m.开单人, m.病人id, m.主页id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用,Nvl(g.出院日期,Sysdate) -g.入院日期 As 天数" & vbNewLine & _
                "              From (Select a.开单人,a.病人id, a.主页id," & vbNewLine & _
                "                   Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds, Sum(a.结帐金额) As 费用" & vbNewLine & _
                "                     From 住院费用记录 A, 药品规格 B,药品特性 D" & vbNewLine & _
                "                     Where a.收费类别 = '5' And a.记录状态 <> 0 And a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.收费细目id = b.药品id And b.药名id = d.药名id and  Nvl(d.抗生素, 0) <> 0" & vbNewLine & _
                "                     Group By a.病人id, a.主页id,a.开单人) M, 病案主页 G" & vbNewLine & _
                "   where  g.病人id = m.病人id And g.主页id = m.主页id   Group By m.病人id, m.主页id, m.开单人,g.出院日期,g.入院日期) M" & vbNewLine & _
                "       Group By m.开单人 having Sum(m.费用) >0 " & vbNewLine & _
                "       Order By Sum(m.费用) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            strSQLDetail = "Select nvl(a.开单人,'空') as 开单人, Sum(a.结帐金额) As 总费用 From 住院费用记录 A" & vbNewLine & _
                "Where a.收费类别 In ('5', '6', '7') And a.记录状态 <> 0 And a.开单人 is not null And" & vbNewLine & _
                " a.发生时间 Between " & strPar1 & " And " & strPar2 & "and instr(',[1],',','||nvl(a.开单人,'空')||',')>0" & vbNewLine & _
                "Group By a.开单人 having sum(a.结帐金额)>0"
        Else
            strDept = IIf(txtDept(e_C0_txtDept_抽样科室_0).Tag = "", "", " and a.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
            
            strSQL出院病人 = "select count(1) as 人数,round(Sum(m.天数)) As 天数,min(m.入院日期) as 开始,max(m.出院日期) as 结束 from" & vbNewLine & _
                "(select a.出院日期-a.入院日期 As 天数,a.出院日期,a.入院日期  from 病案主页 a where a.出院日期 Between [2] And [3] " & strDept & " ) m"
            strPar1 = Format(dtpCountS(e_C0_dtpCountS_开始时间_0).Value, "yyyy-MM-dd 00:00:00")
            strPar2 = Format(dtpCountE(e_C0_dtpCountE_结束时间_0).Value, "yyyy-MM-dd 23:59:59")
            dat出院开始 = CDate(strPar1)
            dat出院结束 = CDate(strPar2)
            Set rs出院病人 = zlDatabase.OpenSQLRecord(strSQL出院病人, Me.Caption, txtDept(e_C0_txtDept_抽样科室_0).Tag, dat出院开始, dat出院结束)
            
            If Val("" & rs出院病人!人数) = 0 Then
                Screen.MousePointer = 0
                Call zlCommFun.StopFlash
                MsgBox "当前条件下未找到任何数据，请重新设置统计参数。", vbInformation, gstrSysName
                Exit Sub
            End If
            
            lng人数 = Val("" & rs出院病人!人数)
            lng天数 = Val("" & rs出院病人!天数)
            
            strDept = IIf(txtDept(e_C0_txtDept_抽样科室_0).Tag = "", "", " and g.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
            strWhere病人 = " and g.出院日期 Between [2] And [3]" & strDept
            
            
            strPar1 = "To_Date('" & Format(rs出院病人!开始, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')"
            strPar2 = "To_Date('" & Format(rs出院病人!结束, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
            
            '汇总方式 药品
            strSql = "Select /*+ rule*/  m.类别, m.药品通用名, m.剂型, m.规格, m.Ddds, m.费用, m.数量," & lng人数 & " As 用例数," & lng天数 & " as 天数, m.住院单位 As 单位" & vbNewLine & _
                "From (Select m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用,Sum(m.数量) As 数量, m.住院单位" & vbNewLine & _
                "       From (Select m.病人id, m.主页id, m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用," & vbNewLine & _
                "                 Sum(m.数量) As 数量, m.住院单位" & vbNewLine & _
                "              From (Select a.病人id, a.主页id, a.收费细目id, e.名称 As 类别, c.名称 As 药品通用名," & vbNewLine & _
                "                            d.药品剂型 As 剂型, c.规格 || c.产地 As 规格," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds, Sum(a.结帐金额) As 费用, b.住院单位," & vbNewLine & _
                "                            Sum(a.数次) As 数量" & vbNewLine & _
                "                     From 住院费用记录 A, 药品规格 B, 收费项目目录 C, 药品特性 D, 诊疗分类目录 E, 诊疗项目目录 F" & vbNewLine & _
                "                     Where a.收费类别 = '5' And a.记录状态 <> 0 And" & vbNewLine & _
                "                           a.发生时间 Between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "                           And a.收费细目id = b.药品id And b.药品id = c.id And d.药名id = b.药名id And d.药名id = f.Id And f.分类id = e.Id And Nvl(d.抗生素,0)<>0" & vbNewLine & _
                "                     Group By a.收费细目id, a.病人id, a.主页id,e.名称, c.名称, d.药品剂型, c.规格, c.产地," & vbNewLine & _
                "                              b.住院单位) M, 病案主页 G where g.病人id = m.病人id And g.主页id = m.主页id" & strWhere病人 & vbNewLine & _
                "              Group By m.病人id, m.主页id, m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格, m.住院单位,g.出院日期,g.入院日期) M" & vbNewLine & _
                "       Group By m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格, m.住院单位 having Sum(m.费用) >0 " & vbNewLine & _
                "       Order By Sum(m." & IIf(optType(e_C0_optType_排序方式_数量_12).Value, "数量", "费用") & ") desc) M" & vbNewLine & _
                "Where Rownum < " & Val(txtTopRan.Text) + 1

            strSQLDetail = "Select Sum(a.结帐金额) As 总药费 From 住院费用记录 A,病案主页 g" & vbNewLine & _
                " Where a.收费类别 In ('5', '6', '7') And a.记录状态 <> 0 And" & vbNewLine & _
                " a.发生时间 Between " & strPar1 & " And " & strPar2 & " and a.病人id+0 = g.病人id And a.主页id+0= g.主页id" & strWhere病人
        End If
    Else    '统计场合 门诊
        If optType(e_C0_optType_汇总方式_科室_9).Value Then
            '汇总方式 科室
            strSql = "Select /*+ rule*/  m.名称, m.开单部门id, m.Ddds, m.费用, m.处方数 as 用例数" & vbNewLine & _
                "From (Select n.名称, m.开单部门id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Sum(m.处方数) As 处方数" & vbNewLine & _
                "       From (Select m.开单部门id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Count(1) As 处方数" & vbNewLine & _
                "              From (Select m.开单部门id, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用" & vbNewLine & _
                "                     From (Select a.开单部门id, a.No, Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds," & vbNewLine & _
                "                                   Sum(a.结帐金额) As 费用" & vbNewLine & _
                "                            From 门诊费用记录 A, 药品规格 B,药品特性 D" & vbNewLine & _
                "                            Where a.收费类别 = '5' And a.记录状态 <> 0 And a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                                  And a.收费细目id = b.药品id And b.药名id = d.药名id And Nvl(d.抗生素, 0) <> 0" & vbNewLine & _
                "                            Group By a.No,a.开单部门id) M" & vbNewLine & _
                "                     Group By m.No, m.开单部门id) M" & vbNewLine & _
                "              Group By m.开单部门id) M, 部门表 N" & vbNewLine & _
                "       Where m.开单部门id = n.Id" & vbNewLine & _
                "       Group By m.开单部门id, n.名称 having Sum(m.费用) >0 " & vbNewLine & _
                "       Order By Sum(m.费用) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
                
            strSQLDetail = "select a.开单部门id as id,sum(a.结帐金额) as 总费用 from 门诊费用记录 a" & vbNewLine & _
                "where a.收费类别 in ('5','6','7') and a.记录状态 <> 0" & vbNewLine & _
                "and a.发生时间 between " & strPar1 & " And " & strPar2 & vbNewLine & _
                "and a.开单部门id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                "group by a.开单部门id having sum(a.结帐金额)>0"

        ElseIf optType(e_C0_optType_汇总方式_医生_8).Value Then
            '汇总方式 医生
            strSql = "Select /*+ rule*/  nvl(m.开单人,'空') as 开单人, m.Ddds, m.费用, m.处方数 as 用例数" & vbNewLine & _
                "From (Select m.开单人, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Count(1) As 处方数" & vbNewLine & _
                "       From (Select m.开单人, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用" & vbNewLine & _
                "              From (Select a.开单人, a.NO, Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds," & vbNewLine & _
                "                            Sum(a.结帐金额) As 费用" & vbNewLine & _
                "                     From 门诊费用记录 A, 药品规格 B,药品特性 D" & vbNewLine & _
                "                     Where a.收费类别 = '5' And a.记录状态 <> 0 And a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                           And a.收费细目id = b.药品id And b.药名id = d.药名id And  Nvl(d.抗生素, 0) <> 0" & vbNewLine & _
                "                     Group By a.No,a.开单人) M" & vbNewLine & _
                "              Group By m.No, m.开单人) M" & vbNewLine & _
                "       Group By m.开单人  having Sum(m.费用) >0" & vbNewLine & _
                "       Order By Sum(m.费用) Desc) M" & vbNewLine & _
                "Where Rownum <" & Val(txtTopRan.Text) + 1
            
            strSQLDetail = "Select nvl(a.开单人,'空') as 开单人, Sum(a.结帐金额) As 总费用 From 门诊费用记录 A" & vbNewLine & _
                "Where a.收费类别 In ('5', '6', '7') And a.记录状态 <> 0 And a.开单人 is not null And " & vbNewLine & _
                " a.发生时间 Between " & strPar1 & " And " & strPar2 & "and instr(',[1],',','||nvl(a.开单人,'空')||',')>0" & vbNewLine & _
                "Group By a.开单人 having Sum(a.结帐金额)>0"
        Else
            '汇总方式 药品
            strSql = "Select /*+ rule*/  m.类别, m.药品通用名, m.剂型, m.规格, m.Ddds, m.费用, m.处方数 as 用例数, m.数量, m.门诊单位 as 单位" & vbNewLine & _
                "From (Select m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格,m.门诊单位,Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Count(1) As 处方数, Sum(m.数量) As 数量" & vbNewLine & _
                "       From (Select m.收费细目id, m.No, m.类别, m.药品通用名, m.剂型, m.规格,m.门诊单位, Sum(m.Ddds) As Ddds, Sum(m.费用) As 费用, Sum(m.数量) As 数量" & vbNewLine & _
                "              From (Select a.No || '' As NO, a.收费细目id, e.名称 As 类别, c.名称 As 药品通用名, d.药品剂型 As 剂型, c.规格 || c.产地 As 规格," & vbNewLine & _
                "                            Sum(Decode(Nvl(b.Ddd值, 0), 0, 0, a.数次 * b.剂量系数 / b.Ddd值)) As Ddds, Sum(a.结帐金额) As 费用," & vbNewLine & _
                "                            Sum(a.数次) As 数量, b.门诊单位" & vbNewLine & _
                "                     From 门诊费用记录 A, 药品规格 B, 收费项目目录 C, 药品特性 D, 诊疗分类目录 E, 诊疗项目目录 F" & vbNewLine & _
                "                     Where a.收费类别 = '5' And a.记录状态 <> 0 And  a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept & vbNewLine & _
                "                            And a.收费细目id = b.药品id And b.药品id = c.id And" & vbNewLine & _
                "                           d.药名id = b.药名id And d.药名id = f.Id And f.分类id = e.Id And Nvl(d.抗生素, 0) <> 0" & vbNewLine & _
                "                     Group By a.收费细目id, a.No, e.名称, c.名称, d.药品剂型, c.规格, c.产地,b.门诊单位) M" & vbNewLine & _
                "       Group By m.收费细目id,m.No, m.类别, m.药品通用名, m.剂型, m.规格, m.门诊单位 having Sum(m.费用) >0) M" & vbNewLine & _
                "       group by  m.收费细目id, m.类别, m.药品通用名, m.剂型, m.规格, m.门诊单位" & vbNewLine & _
                "       Order By Sum(m." & IIf(optType(e_C0_optType_排序方式_数量_12).Value, "数量", "费用") & ") Desc) M" & vbNewLine & _
                "Where Rownum < " & Val(txtTopRan.Text) + 1
            strSQLDetail = "Select Sum(a.结帐金额) As 总药费 From 门诊费用记录 A" & vbNewLine & _
                " Where a.收费类别 In ('5', '6', '7') And a.记录状态 <> 0 And" & vbNewLine & _
                " a.发生时间 Between " & strPar1 & " And " & strPar2 & strDept
        End If
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C0_txtDept_抽样科室_0).Tag, dat出院开始, dat出院结束)
    
    '表格要重新初始化 和 数据加载
    vsUseRan.Rows = vsUseRan.FixedRows
    vsUseRan.Rows = vsUseRan.FixedRows + 1
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If optType(e_C0_optType_汇总方式_科室_9).Value Or optType(e_C0_optType_汇总方式_医生_8).Value Then '  科室或医生
        strTmp = IIf(optType(e_C0_optType_汇总方式_医生_8).Value, "医生姓名", "科室名称") & ",2000,1;总金额(元),1510,7;使用例数(人数),1440,7;患者人天数,1080,7;每例平均金额(元),1720,7;DDDs,1000,7;使用强度,1000,7;占药品金额比例(%),2100,7"
        Call InitTable(vsUseRan, strTmp)
        vsUseRan.RowHeight(0) = 600
        If rsTmp.RecordCount > 0 Then
            strTmp = ""
            For i = 1 To rsTmp.RecordCount
                If optType(e_C0_optType_汇总方式_科室_9).Value Then '按科室汇总
                    strTmp = strTmp & "," & rsTmp!开单部门id
                Else
                    strTmp = strTmp & ",'" & rsTmp!开单人 & ","
                End If
                rsTmp.MoveNext
            Next
            strTmp = Mid(strTmp, 2)
            rsTmp.MoveFirst
            
            Set rsDetail = zlDatabase.OpenSQLRecord(strSQLDetail, Me.Caption, strTmp)
            
            With vsUseRan
                .Rows = rsTmp.RecordCount + 1
                For i = 1 To rsTmp.RecordCount
                    If optType(e_C0_optType_汇总方式_医生_8).Value Then
                        .TextMatrix(i, COL_D名称) = rsTmp!开单人 & "" '医生姓名
                    Else
                        .TextMatrix(i, COL_D名称) = rsTmp!名称 & "" '科室名称
                    End If
                    
                    .TextMatrix(i, COL_D总金额) = Format(Val(rsTmp!费用 & ""), strDec)       '总金额(元) 抗菌药费
                    .TextMatrix(i, COL_D使用例数) = Val(rsTmp!用例数 & "") ' 使用例数(人数)
                    .TextMatrix(i, COL_DDDDs) = Format(Val(rsTmp!Ddds & ""), strDec)
                    
                    If optType(e_C0_optType_统计场合_住院_5).Value Then
                        .TextMatrix(i, COL_D患者人天数) = Val(rsTmp!天数 & "")
                    End If
                    
                    dblTmp = 0
                    If Val(rsTmp!用例数 & "") <> 0 Then dblTmp = Val(rsTmp!费用 & "") / Val(rsTmp!用例数 & "")
                    .TextMatrix(i, COL_D每例平均金额) = Format(dblTmp, strDec)   '每例平均金额(元)
                    
                    dblTmp = 0
                    If optType(e_C0_optType_统计场合_住院_5).Value Then '门诊和住院的使用强度计算方式不同  optN02--住院
                        If Val(rsTmp!天数 & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!天数 & "")
                    Else
                        If Val(rsTmp!用例数 & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!用例数 & "")
                    End If
                    .TextMatrix(i, COL_D使用强度) = Format(dblTmp, strDec) '使用强度
                    
                    dblTmp = 0: dbl总药费 = 0: rsDetail.Filter = 0
                    
                    If optType(e_C0_optType_汇总方式_医生_8).Value Then
                        rsDetail.Filter = "开单人='" & rsTmp!开单人 & "'"
                    Else
                        rsDetail.Filter = "id=" & rsTmp!开单部门id
                    End If
                    
                    If Not rsDetail.EOF Then dbl总药费 = Val(rsDetail!总费用 & "")
                    If dbl总药费 <> 0 Then dblTmp = Val(rsTmp!费用 & "") * 100 / dbl总药费
                    .TextMatrix(i, COL_D占药品金额比例) = Format(dblTmp, strDec) & "%"  '占药品金额比例(%)
                    
                    rsTmp.MoveNext
                Next
                .ColHidden(COL_D患者人天数) = Not optType(e_C0_optType_统计场合_住院_5).Value '当统计场合是住院时才显示  患者人天数
            End With
        End If
    Else
        '汇总方式 药品
        strTmp = "类别,1480,4;药品名称,2480,1;剂型,1530,4;规格,2800,1;数量,1020,7;总金额(元),1000,7;使用例次,530,7;患者人天数,760,7;每例平均金额(元),930,7;DDDs,750,7;使用强度,750,7;占药品总金额比例(%),1000,7"
        Call InitTable(vsUseRan, strTmp)
        vsUseRan.RowHeight(0) = 600
        If rsTmp.RecordCount > 0 Then
            Set rsDetail = zlDatabase.OpenSQLRecord(strSQLDetail, Me.Caption, txtDept(0).Tag, dat出院开始, dat出院结束)
            If Not rsDetail.EOF Then dbl总药费 = Val(rsDetail!总药费 & "")
            With vsUseRan
                .Rows = rsTmp.RecordCount + 1
                
                For i = 1 To rsTmp.RecordCount
                    .TextMatrix(i, COL_UD类别) = rsTmp!类别 & ""
                    .TextMatrix(i, COL_UD药品名称) = rsTmp!药品通用名 & ""
                    .TextMatrix(i, COL_UD剂型) = rsTmp!剂型 & ""
                    .TextMatrix(i, COL_UD规格) = rsTmp!规格 & ""
                    .TextMatrix(i, COL_UD数量) = Val(rsTmp!数量 & "") & rsTmp!单位 '数量
                    .TextMatrix(i, COL_UD总金额) = Format(rsTmp!费用 & "", strDec)
                    .TextMatrix(i, COL_UD使用例次) = rsTmp!用例数 & ""
                    .TextMatrix(i, COL_UDDDDs) = Format(rsTmp!Ddds & "", strDec)
                    
                    '当统计场合是住院时才显示  患者人天数  列
                    If optType(e_C0_optType_统计场合_住院_5).Value Then .TextMatrix(i, COL_UD患者人天数) = rsTmp!天数 & ""
                    
                    dblTmp = 0
                    If Val(rsTmp!用例数 & "") <> 0 Then dblTmp = Val(rsTmp!费用 & "") / Val(rsTmp!用例数 & "")
                    .TextMatrix(i, COL_UD每例平均金额) = Format(dblTmp, strDec) '平均金额
                    
                    dblTmp = 0 '门诊和住院的使用强度计算方式不同，门诊是认为是每例用一天等同于 天数 ＝ 用例数
                    If optType(e_C0_optType_统计场合_住院_5).Value Then
                        If Val(rsTmp!天数 & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!天数 & "")
                    Else
                        If Val(rsTmp!用例数 & "") <> 0 Then dblTmp = Val(rsTmp!Ddds & "") * 100 / Val(rsTmp!用例数 & "")
                    End If
                    .TextMatrix(i, COL_UD使用强度) = Format(dblTmp, strDec) '使用强度
                    
                    dblTmp = 0
                    If dbl总药费 <> 0 Then dblTmp = Val(rsTmp!费用 & "") * 100 / dbl总药费
                    .TextMatrix(i, COL_UD占药品总金额比例) = Format(dblTmp, strDec) & "%" '占药品总金额比例
                    
                    rsTmp.MoveNext
                Next
                .ColHidden(COL_UD规格) = False '规格列显示
                .ColHidden(COL_UD患者人天数) = optType(e_C0_optType_统计场合_门诊_4).Value   '当统计场合是住院时才显示  患者人天数  列
            End With
        End If
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "抗菌药物使用情况排名统计", dtpCountS(e_C0_dtpCountS_开始时间_0).Value & "," & dtpCountE(e_C0_dtpCountE_结束时间_0).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsCut()
'功能：加载  Ⅰ类切口围术期预防用药统计 界面表格数据
    Dim strSql As String, strPar As String
    Dim rs人数 As ADODB.Recordset
    Dim rs抗药人数 As ADODB.Recordset
    Dim rs一切口人数 As ADODB.Recordset
    Dim rs术前用药人数 As ADODB.Recordset
    Dim rs计算天数 As ADODB.Recordset
    Dim rs天数 As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim lng切口预防人数 As Long
    Dim lng总人数 As Long
    Dim lng药品数 As Long
    Dim lng术前用 As Long
    Dim lng天数 As Long
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim lngRow As Long
    Dim strPatis As String
    Dim arrTmp As Variant
    
    Dim rs药品 As ADODB.Recordset
    Dim rs切口抗菌药 As ADODB.Recordset
 
    Dim strWhere As String
    Dim strTmp As String
    Dim str天数 As String
 
    Dim int术前 As Integer
    Dim i As Long, j As Long, k As Long, m As Long
    
    '参数生成
    strPar = "To_Date('" & Format(dtpCountS(e_C1_dtpCountS_开始时间_1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C1_dtpCountS_开始时间_1).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')" & _
        IIf(txtDept(e_C1_txtDept_统计科室_4).Tag = "", "", " and a.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
 
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    vsCut.Rows = vsCut.FixedRows
    vsCut.Rows = vsCut.FixedRows + 1
    vsCut.Cell(flexcpText, 0, COL_CUT名称, 1, COL_CUT名称) = IIf(optType(e_C1_optType_汇总方式_科室_17).Value, "科室名称", "住院医师")
    
    
    '按出院病人进行抽样，提取一批出院病人满足时间和科室即可
    '总人数
    If optType(e_C1_optType_汇总方式_科室_17).Value Then '科室汇总
        strSql = "Select a.出院科室id as 科室id,b.名称,Count(1) As 人数 From 病案主页 A, 部门表 B Where a.出院日期 Between " & strPar & " And a.出院科室id = b.Id Group By a.出院科室id,b.名称"
    Else   '医生汇总
        strSql = "Select 0 as 科室id,Nvl(a.住院医师,'空') As 名称, Count(1) As 人数 From 病案主页 A Where a.出院日期 Between " & strPar & " Group By Nvl(a.住院医师,'空')"
    End If
    Set rs人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_统计科室_4).Tag)
    
    If rs人数.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '抗菌药人数
    If optType(e_C1_optType_汇总方式_科室_17).Value Then '科室汇总
        strSql = "Select a.出院科室id as 科室id, Count(1) As 人数" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (Select 1 From 住院费用记录 B, 药品规格 C, 药品特性 D" & vbNewLine & _
            " Where a.病人id = b.病人id And a.主页id = b.主页id And b.记录状态 <> 0 And b.收费类别 = '5' And b.收费细目id = c.药品id And" & vbNewLine & _
            " c.药名id = d.药名id And Nvl(d.抗生素, 0) <> 0) Group By a.出院科室id"

    Else   '医生汇总
        strSql = "Select Nvl(a.住院医师,'空') as 名称, Count(1) As 人数" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (Select 1 From 住院费用记录 B, 药品规格 C, 药品特性 D" & vbNewLine & _
            " Where a.病人id = b.病人id And a.主页id = b.主页id And b.记录状态 <> 0 And b.收费类别 = '5' And b.收费细目id = c.药品id And" & vbNewLine & _
            " c.药名id = d.药名id And Nvl(d.抗生素, 0) <> 0) Group By Nvl(a.住院医师,'空')"
    End If
    Set rs抗药人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_统计科室_4).Tag)
    
    '一类切口例数
    If optType(e_C1_optType_汇总方式_科室_17).Value Then '科室汇总
        strSql = "Select a.出院科室id as 科室id, Count(1) As 人数" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from 病人手麻记录 b where a.病人id = b.病人id And a.主页id = b.主页id and b.切口='Ⅰ') Group By a.出院科室id"

    Else   '医生汇总
        strSql = "Select Nvl(a.住院医师,'空') as 名称, Count(1) As 人数" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from 病人手麻记录 b where a.病人id = b.病人id And a.主页id = b.主页id and b.切口='Ⅰ') Group By Nvl(a.住院医师,'空')"
    End If
    Set rs一切口人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_统计科室_4).Tag)
    
    '如果这两项统结果没有也应该退出
    If rs抗药人数.EOF And rs一切口人数.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '一类切口使用抗菌药目的为预防的明细，用药目的＝1 的这个条件可以确定为抗菌药
    If optType(e_C1_optType_汇总方式_科室_17).Value Then '科室汇总
        strSql = "Select a.出院科室id as 科室id,a.病人id,a.主页id" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from 病人医嘱记录 c where c.医嘱状态 in (8,9) and c.病人id=a.病人id and c.主页id=a.主页id and c.用药目的=1 and c.诊疗类别='5')" & _
            " and Exists (select 1 from 病人手麻记录 b where a.病人id = b.病人id And a.主页id = b.主页id and b.切口='Ⅰ')" & _
            " Group By a.出院科室id,a.病人id, a.主页id"
    Else   '医生汇总
        strSql = "Select Nvl(a.住院医师,'空') as 名称,a.病人id,a.主页id" & vbNewLine & _
            " From 病案主页 A Where a.出院日期 Between " & strPar & " And Exists" & vbNewLine & _
            " (select 1 from 病人医嘱记录 c where c.医嘱状态 in (8,9) and c.病人id=a.病人id and c.主页id=a.主页id and c.用药目的=1 and c.诊疗类别='5')" & _
            " and Exists (select 1 from 病人手麻记录 b where a.病人id = b.病人id And a.主页id = b.主页id and b.切口='Ⅰ')" & _
            " Group By Nvl(a.住院医师,'空'),a.病人id, a.主页id"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C1_txtDept_统计科室_4).Tag)
    
    lng切口预防人数 = rsTmp.RecordCount
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
        rsTmp.MoveNext
    Next
    strPatis = Mid(strPatis, 2) '得到病人
    
    If strPatis <> "" Then
        '术前用药人数
        strSql = "Select /*+ rule*/ a.病人id,a.主页id,count(1) as 用药种数,To_Char(min(a.开始执行时间), 'YYYY-MM-DD HH24:MI:SS') as 开始执行时间," & _
            " to_char(max(a.手术开始时间), 'YYYY-MM-DD HH24:MI:SS') as 手术开始时间" & _
            " from ( select a.病人id,a.主页id,b.诊疗项目id,min(b.开始执行时间) as 开始执行时间,max(c.手术开始时间) as 手术开始时间" & _
            " From 住院费用记录 A, 病人医嘱记录 B,病人手麻记录 c" & vbNewLine & _
            " Where a.记录状态 <> 0 And a.收费类别 = '5' And a.医嘱序号 = b.Id and a.病人id=c.病人id and a.主页id=c.主页id and c.切口='Ⅰ' and b.用药目的=1" & _
            " and (a.病人id,a.主页id) In (Select C1, C2 From Table(f_Num2list2([1]))) group by a.病人id,a.主页id,b.诊疗项目id) a group by a.病人id,a.主页id"
        Set rs术前用药人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        
        '已有数据计算天数
        strSql = "select 病人id,主页id,Zl_Adviceexetimes(Id,开始执行时间,Nvl(Nvl(上次执行时间,执行终止时间),停嘱时间)," & _
            " 执行时间方案,开始执行时间,开始执行时间-1,频率间隔,间隔单位,医嘱期效) as 分解时间 From 病人医嘱记录" & _
            " where (病人id,主页id) In (Select To_Number(C1), C2 From Table(f_Str2list2([1]))) and 用药目的=1 and 诊疗类别='5' and 医嘱状态 in (8,9)"
        Set rs计算天数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        
        '用药天数，这个从 病人抗生素记录 中取，因为天数不便于计算
        strSql = "select a.病人id,a.主页id,sum(a.使用天数) as 天数 from 病人抗生素记录 a where (a.病人id,a.主页id) In (Select To_Number(C1), C2 From Table(f_Str2list2([1]))) group by a.病人id,a.主页id"
        Set rs天数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    '加载数据
    With vsCut
        .Rows = .FixedRows
        For i = 1 To rs人数.RecordCount
NextRow:
            If optType(e_C1_optType_汇总方式_科室_17).Value Then '科室汇总
                rs抗药人数.Filter = "科室id=" & Val(rs人数!科室ID & "")
                rs一切口人数.Filter = "科室id=" & Val(rs人数!科室ID & "")
                rsTmp.Filter = "科室id=" & Val(rs人数!科室ID & "")
            Else
                rs抗药人数.Filter = "名称='" & rs人数!名称 & "'"
                rs一切口人数.Filter = "名称='" & rs人数!名称 & "'"
                rsTmp.Filter = "名称='" & rs人数!名称 & "'"
            End If
            
            lng总人数 = 0
            If Not rs一切口人数.EOF Then lng总人数 = Val(rs一切口人数!人数 & "")
            lngTmp = 0
            If Not rs抗药人数.EOF Then lngTmp = Val(rs抗药人数!人数 & "")
            
            If lng总人数 = 0 And lngTmp = 0 Then
                i = i + 1
                rs人数.MoveNext
                If rs人数.EOF Then Exit For
                GoTo NextRow
            End If
            
            .AddItem "": lngRow = .Rows - 1
            
            .TextMatrix(lngRow, COL_CUT名称) = rs人数!名称 & ""
            lng总人数 = Val(rs人数!人数 & "")
            
            If Not rs抗药人数.EOF Then lngTmp = Val(rs抗药人数!人数 & "")
            .TextMatrix(lngRow, COL_CUT使用人次) = lngTmp
            
            If lng总人数 <> 0 Then dblTmp = lngTmp * 100 / lng总人数
            .TextMatrix(lngRow, COL_CUT使用率) = Format(dblTmp, "0.00") & "%"
            lngTmp = 0: dblTmp = 0
            
            If Not rs一切口人数.EOF Then lngTmp = Val(rs一切口人数!人数 & "")
            .TextMatrix(lngRow, COL_CUT切口数) = lngTmp: lngTmp = 0
 
            If Not rsTmp.EOF Then lngTmp = rsTmp.RecordCount
            .TextMatrix(lngRow, COL_CUT抗菌物数) = lngTmp: lngTmp = 0
    
            If Val(.TextMatrix(lngRow, COL_CUT切口数)) <> 0 Then dblTmp = Val(.TextMatrix(lngRow, COL_CUT抗菌物数)) * 100 / Val(rs一切口人数!人数 & "")
            .TextMatrix(lngRow, COL_CUT切口使用率) = Format(dblTmp, "0.00") & "%": dblTmp = 0
            
            For j = 1 To rsTmp.RecordCount
                If strPatis <> "" Then
                    rs术前用药人数.Filter = "病人id=" & rsTmp!病人ID & " and 主页id=" & rsTmp!主页ID
                    If Not rs术前用药人数.EOF Then
                        lngTmp = lngTmp + 1 '切口预防用抗菌数量
                        lng药品数 = lng药品数 + Val(rs术前用药人数!用药种数 & "") '切口预防用抗菌用药总品种数
                        If rs术前用药人数!开始执行时间 & "" <> "" And rs术前用药人数!手术开始时间 & "" <> "" Then '切口预防用抗菌术前使用数量
                            If rs术前用药人数!开始执行时间 & "" < rs术前用药人数!手术开始时间 & "" Then lng术前用 = lng术前用 + 1
                        End If
                    End If
                    '天数
                    rs计算天数.Filter = "病人id=" & rsTmp!病人ID & " and 主页id=" & rsTmp!主页ID
                    If Not rs计算天数.EOF Then
                        For k = 1 To rs计算天数.RecordCount
                            strTmp = rs计算天数!分解时间 & ""
                            If strTmp <> "" Then
                                arrTmp = Split(strTmp, ",")
                                For m = 0 To UBound(arrTmp)
                                    strTmp = Split(arrTmp(m), " ")(0)
                                    If InStr("," & str天数 & ",", "," & strTmp & ",") = 0 Then
                                        str天数 = str天数 & "," & strTmp
                                    End If
                                Next
                            End If
                            rs计算天数.MoveNext
                        Next
                        
                        If str天数 <> "" Then
                            str天数 = Mid(str天数, 2)
                            lngTmp = UBound(Split(str天数, ",")) + 1
                        End If
                        
                        lng天数 = lng天数 + lngTmp
                        lngTmp = 0: str天数 = "": strTmp = ""
                    Else
                        rs天数.Filter = "病人id=" & rsTmp!病人ID & " and 主页id=" & rsTmp!主页ID
                        If Not rs天数.EOF Then
                            lng天数 = lng天数 + Val(rs天数!天数 & "")
                        Else
                            lng天数 = lng天数 + 1
                        End If
                    End If
                    
                End If
                rsTmp.MoveNext
            Next
            
            .TextMatrix(lngRow, COL_CUT抗菌物数) = lngTmp: lngTmp = 0 '切口预防用抗菌数量
            
            .TextMatrix(lngRow, COL_CUT术前用药) = lng术前用: lng术前用 = 0 '切口预防用抗菌术前使用数量
            
            .TextMatrix(lngRow, COL_CUT平均用药) = lng天数: lng天数 = 0 '天数
            
            .TextMatrix(lngRow, COL_CUT品种数) = lng药品数: lng药品数 = 0 '切口预防用抗菌用药总品种数
            rs人数.MoveNext
        Next
    End With
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "Ⅰ类切口围术期预防用药统计", dtpCountS(e_C1_dtpCountS_开始时间_1).Value & "," & dtpCountE(e_C1_dtpCountE_结束时间_1).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsInDruUse()
'功能：加载   住院医嘱抗菌用药统计  界面表格数据加载
    Dim strSql As String, strPar As String, strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim str医嘱IDs As String
    Dim rs诊断 As ADODB.Recordset       '出院诊断 一对多
    Dim rs手术 As ADODB.Recordset       '可取 手术名称 和 切口类型 两列 一对多
    Dim rs金额 As ADODB.Recordset       '治疗金额，药品金额 一对一
    Dim rs种数 As ADODB.Recordset       '药品种数，抗菌药物品种数 一对一
    Dim rs天数 As ADODB.Recordset
    Dim lngBaseRow As Long, lngTmpRow As Long
    Dim dblTmp As Double, strDec As String
    Dim rs用药明细 As ADODB.Recordset   '抗菌药物使用情况
    Dim lng抽样数 As Long
    Dim str切口 As String
    Dim strPatis As String '病人信息 格式："病人id1:主页id1,病人id2:主页id2,......."
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
    Dim strFilter As String
    Dim strTmp As String
    Dim strTmp1 As String
    Dim lngTmp As Long
    Dim i As Long, j As Long, k As Long
    
    '统计分析
    Dim lng药品数 As Long
    Dim lng抗药数 As Long
    Dim lng用抗菌药人数 As Long
    Dim lng实际人数 As Long
    Dim dbl总金额 As Double
    Dim dbl药品金额 As Double
    Dim dbl抗药金额 As Double
    Dim lng单用人数 As Long
    Dim lng二用人数 As Long
    Dim lng三用人数 As Long
    Dim lng四用人数 As Long
    Dim lng预防人数 As Long
    Dim lng治疗人数 As Long
    Dim str用药的 As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    lng抽样数 = Val(txtNum(e_C3_txtNum_抽样数量_1).Text)
    strDec = "0.00"
    
    '日期范围参数生成
    strPar = "To_Date('" & Format(dtpCountS(e_C3_dtpCountS_开始时间_3).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C3_dtpCountE_结束时间_3).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C3_txtDept_统计科室_6).Tag = "", "", " and a.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
     
    '公共SQL：指定时间内，指定科室内，用过或者没用过抗菌药的病人都应该抽样
    strSql = "Select a.病人id,a.主页id,a.住院号,a.姓名,To_Char(a.出院日期, 'YYYY-MM-DD HH24:MI:SS') as 出院日期,a.住院医师 as 主管医生,c.名称 as 科室,a.住院天数" & _
        " From 病案主页 A,部门表 c" & _
        " Where a.出院科室id =c.id" & _
        " And a.出院日期 Between " & strPar & strDept
        
    If optType(e_C3_optType_切口类型_非手术_15).Value Then '非手术
        If optType(e_C3_optType_抽样方法_平均_14).Value Then
            '平均抽样
            strSql = strSql & " And Not Exists (Select 1 From 病人手麻记录 X Where a.病人id=x.病人id And a.主页id=x.主页id) Order By a.出院日期 desc"
            
            strSql = "select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数 from (" & _
                "select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数 from (" & _
                    "select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数,Mod(Rownum,[2]) M from (" & strSql & ") m  Order By M) M " & _
                     " Where Rownum <([2]+1)) M Order By m.出院日期 Desc"
        Else
            '随机抽样
            strSql = strSql & " And Not Exists (Select 1 From 病人手麻记录 X Where a.病人id=x.病人id And a.主页id=x.主页id) Order By Dbms_Random.Value"
            strSql = "select 病人id,主页id,住院号,姓名,出院日期,主管医生,科室,住院天数 from (" & strSql & ") where rownum < ([2] + 1) Order By 出院日期 desc"
        End If
    Else
        str切口 = str切口 & IIf(chkType(e_C3_chkType_切口类型_Ⅰ类_2).Value = 1, "Ⅰ,", "")
        str切口 = str切口 & IIf(chkType(e_C3_chkType_切口类型_Ⅱ类_3).Value = 1, "Ⅱ,", "")
        str切口 = str切口 & IIf(chkType(e_C3_chkType_切口类型_Ⅲ类_4).Value = 1, "Ⅲ,", "")
        str切口 = str切口 & IIf(chkType(e_C3_chkType_切口类型_Ⅳ类_8).Value = 1, "Ⅳ,", "")
        
        If str切口 = "Ⅰ,Ⅱ,Ⅲ,Ⅳ," Then str切口 = ""
         
        If optType(e_C3_optType_抽样方法_平均_14).Value Then
            '平均抽样
            strSql = strSql & " And Exists (Select 1 From 病人手麻记录 B Where a.病人id=b.病人id And a.主页id=b.主页id" & _
            IIf(str切口 = "", "", " And Instr([3],b.切口)>0") & ") Order By a.出院日期 desc"
            
            strSql = "select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数 from (" & _
                " select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数 from (" & _
                " select m.病人id,m.主页id,m.住院号,m.姓名,m.出院日期,m.主管医生,m.科室,m.住院天数,Mod(Rownum,[2]) M from (" & strSql & ") m  Order By M) M " & _
                " Where Rownum <([2]+1)) M Order By m.出院日期 Desc"
        Else
            '随机抽样
            strSql = strSql & " And Exists (Select 1 From 病人手麻记录 B  Where a.病人id=b.病人id And a.主页id=b.主页id" & IIf(str切口 = "", "", " And Instr([3],b.切口)>0") & ") Order By Dbms_Random.Value"
            strSql = "select 病人id,主页id,住院号,姓名,出院日期,主管医生,科室,住院天数 from (" & strSql & ") where rownum < ([2] + 1) Order By 出院日期"
        End If
    End If
    
    '清空数据
    vsInDruUse.Rows = vsInDruUse.FixedRows
    vsInDruUse.Rows = vsInDruUse.FixedRows + 1
    '列显示与隐藏
    vsInDruUse.ColHidden(COL_DRU手术名称) = optType(e_C3_optType_切口类型_非手术_15).Value 'optType(15).Value ：true 非手术 false 手术
    vsInDruUse.ColHidden(COL_DRU切口类型) = optType(e_C3_optType_切口类型_非手术_15).Value
    lblN(e_C3_lblN_分析表_标题_59).Caption = "0例出院病人抗菌药物使用统计分析表"
    
    '参数：科室，抽样数，切口
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C3_txtDept_统计科室_6).Tag, lng抽样数, str切口)
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "按照当前设置的条件未到到任何数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lng实际人数 = rsTmp.RecordCount '抽样的人数可能比界面填写的数量小
    lblN(e_C3_lblN_分析表_标题_59).Caption = lng实际人数 & "例出院病人抗菌药物使用统计分析表"
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
        rsTmp.MoveNext
    Next

    strPatis = Mid(strPatis, 2) '得到病人
    
    strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
    strTable = strParTable
    
    If Len(strPatis) >= 4000 Then
        varArr = Array()
        varArr = GetParTable(strPar, strParTable, strTable)
    End If
    
    '取首页的第一条出院诊断，包括中西医
    strSql = "select a.病人id,a.主页id,a.诊断描述 as 诊断 from 病人诊断记录 a where a.记录来源=3 And NVL(A.编码序号,1) = 1 and a.诊断次序=1 and a.诊断类型 in (3,13) and (a.病人id,a.主页id) In (" & strTable & ")"
    
    If Len(strPatis) >= 4000 Then
        Set rs诊断 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs诊断 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    If Not optType(e_C3_optType_切口类型_非手术_15).Value Then '手术病人
        strSql = "select a.病人id,a.主页id,a.已行手术 as 名称,a.切口 from 病人手麻记录 a where (a.病人id,a.主页id) In (" & strTable & ")"
      
        If Len(strPatis) >= 4000 Then
            Set rs手术 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
                CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        Else
            Set rs手术 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        End If

    End If
    
    '抗菌药物金额后面再取
    strSql = "select a.病人id,a.主页id,sum(a.结帐金额) as 治疗金额,Sum(Decode(a.收费类别,'5',a.结帐金额,'6',a.结帐金额,'7',a.结帐金额, 0)) As 药品金额" & _
        " from 住院费用记录 a where a.记录状态<>0 and (a.病人id,a.主页id) In (" & strTable & ") group by a.病人id,a.主页id"
        
    If Len(strPatis) >= 4000 Then
        Set rs金额 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs金额 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    '药品种数 附代取出 抗菌药物金额
    strSql = "Select a.病人id, a.主页id, Count(1) As 药品种数, Sum(a.抗菌药) As 抗菌药种数, Sum(a.抗菌药金额) As 抗菌药金额" & vbNewLine & _
        "From (Select a.病人id, a.主页id, c.药名id, Decode(Nvl(c.抗生素, 0), 0, 0, 1) As 抗菌药," & vbNewLine & _
        "              Sum(Decode(Nvl(c.抗生素, 0), 0, 0, a.结帐金额)) As 抗菌药金额" & vbNewLine & _
        "       From 住院费用记录 A, 药品规格 B, 药品特性 C" & vbNewLine & _
        "       Where a.收费类别 In ('5', '6', '7') And a.收费细目id = b.药品id And b.药名id = c.药名id And a.记录状态 <> 0 And" & vbNewLine & _
        "             (a.病人id, a.主页id) In (" & strTable & ")" & vbNewLine & _
        "       Group By a.病人id, a.主页id, c.药名id, Nvl(c.抗生素, 0)) A" & vbNewLine & _
        "Group By a.病人id, a.主页id"
    
    
    If Len(strPatis) >= 4000 Then
        Set rs种数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs种数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If

    '用药明细
    strSql = "select e.病人id,e.主页id,e.id,e.相关id,e.医嘱期效,a.药品名称,a.剂型,a.规格,e.执行频次,e.单次用量,f.计算单位," & vbNewLine & _
        "e.天数,g.医嘱内容 as 给药途径,decode(e.用药目的,1,'预防',2,'治疗','') as 目的" & vbNewLine & _
        "from (Select a.医嘱序号 as 医嘱id,d.名称 as 药品名称,c.药品剂型 as 剂型,d.规格 || d.产地 As 规格" & vbNewLine & _
        "From 住院费用记录 A, 药品规格 B, 药品特性 C,收费项目目录 d" & vbNewLine & _
        "Where a.收费类别 = '5' And a.收费细目id = b.药品id And b.药名id = c.药名id" & vbNewLine & _
        "and b.药品id=d.id  And Nvl(c.抗生素, 0) <> 0 And a.记录状态 <> 0 And" & vbNewLine & _
        "      (a.病人id, a.主页id) In (" & strTable & ")" & vbNewLine & _
        "group by a.医嘱序号,d.名称,c.药品剂型,d.规格,d.产地) a,病人医嘱记录 e,诊疗项目目录 f,病人医嘱记录 g" & vbNewLine & _
        "where a.医嘱id=e.id and e.诊疗项目id=f.id and e.相关id=g.id"
    
    If Len(strPatis) >= 4000 Then
        Set rs用药明细 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs用药明细 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
    End If
    
    '用药天数，只计算长期医嘱
    rs用药明细.Filter = "医嘱期效=0"
    If Not rs用药明细.EOF Then
        For i = 1 To rs用药明细.RecordCount
            If InStr("," & str医嘱IDs & ",", "," & rs用药明细!相关ID & ",") = 0 Then
                str医嘱IDs = str医嘱IDs & "," & rs用药明细!相关ID
            End If
            rs用药明细.MoveNext
        Next
        str医嘱IDs = Mid(str医嘱IDs, 2)
        If str医嘱IDs <> "" Then
            strSql = "select id as 医嘱id,Zl_Adviceexetimes(Id,开始执行时间,Nvl(上次执行时间,执行终止时间),执行时间方案,开始执行时间,开始执行时间-1,频率间隔,间隔单位,医嘱期效) as 分解时间" & vbNewLine & _
                "From 病人医嘱记录 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
            Set rs天数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str医嘱IDs)
        End If
        rs用药明细.Filter = 0
    End If
    
    '加载数据
    rsTmp.MoveFirst
    
    With vsInDruUse
        .Rows = .FixedRows
        For i = 1 To rsTmp.RecordCount
            .AddItem ""
            lngBaseRow = .Rows - 1
            .TextMatrix(lngBaseRow, COL_DRU病人id) = rsTmp!病人ID
            .TextMatrix(lngBaseRow, COL_DRU主页id) = rsTmp!主页ID
            .TextMatrix(lngBaseRow, COL_DRU序号) = i
            .TextMatrix(lngBaseRow, COL_DRU住院号) = rsTmp!住院号 & ""
            .TextMatrix(lngBaseRow, COL_DRU病人姓名) = rsTmp!姓名 & ""
            .TextMatrix(lngBaseRow, COL_DRU出院日期) = Format(rsTmp!出院日期 & "", "YYYY-MM-DD")
            .TextMatrix(lngBaseRow, COL_DRU主管医生) = rsTmp!主管医生 & ""
            .TextMatrix(lngBaseRow, COL_DRU科室) = rsTmp!科室 & ""
            .TextMatrix(lngBaseRow, COL_DRU住院天数) = rsTmp!住院天数 & ""
            
            strFilter = "病人id=" & rsTmp!病人ID & " And 主页id=" & rsTmp!主页ID
            '诊断
            rs诊断.Filter = strFilter
            If Not rs诊断.EOF Then rs诊断.MoveFirst
            strTmp = ""
            Do Until rs诊断.EOF
                If InStr(";" & strTmp & ";", ";" & rs诊断!诊断 & ";") = 0 Then
                    strTmp = strTmp & ";" & rs诊断!诊断
                End If
                rs诊断.MoveNext
            Loop
            .TextMatrix(lngBaseRow, COL_DRU出院诊断) = Mid(strTmp, 2)
            strTmp = "": rs诊断.Filter = 0
            '手术和切口
            If Not optType(e_C3_optType_切口类型_非手术_15).Value Then
                rs手术.Filter = strFilter
                If Not rs手术.EOF Then rs手术.MoveFirst
                Do Until rs手术.EOF
                    strTmp = strTmp & ";" & rs手术!名称
                    If InStr("," & strTmp1 & ",", "," & rs手术!切口 & ",") = 0 Then
                        strTmp1 = strTmp1 & "," & rs手术!切口
                    End If
                    rs手术.MoveNext
                Loop
                .TextMatrix(lngBaseRow, COL_DRU手术名称) = Mid(strTmp, 2)
                .TextMatrix(lngBaseRow, COL_DRU切口类型) = Mid(strTmp1, 2)
                strTmp = "": strTmp1 = "": rs手术.Filter = 0
            End If
            '金额
            rs金额.Filter = strFilter
            If Not rs金额.EOF Then
                rs金额.MoveFirst
                .TextMatrix(lngBaseRow, COL_DRU治疗金额) = Format(rs金额!治疗金额 & "", strDec)
                .TextMatrix(lngBaseRow, COL_DRU药品金额) = Format(rs金额!药品金额 & "", strDec)
                If Val(rsTmp!住院天数 & "") <> 0 Then dblTmp = Val(rs金额!治疗金额 & "") / Val(rsTmp!住院天数 & "")
                .TextMatrix(lngBaseRow, COL_DRU日均治疗金额) = Format(rs金额!药品金额 & "", strDec): dblTmp = 0
                
                dbl总金额 = dbl总金额 + Val(rs金额!治疗金额 & "")
                dbl药品金额 = dbl药品金额 + Val(rs金额!药品金额 & "")
                
            End If
            rs金额.Filter = 0
            '种数
            rs种数.Filter = strFilter
            If Not rs种数.EOF Then
                .TextMatrix(lngBaseRow, COL_DRU药品种数) = rs种数!药品种数 & ""
                .TextMatrix(lngBaseRow, COL_DRU抗菌药物品种数) = rs种数!抗菌药种数 & ""
                .TextMatrix(lngBaseRow, COL_DRU抗菌药物金额) = Format(rs种数!抗菌药金额 & "", strDec)
                .TextMatrix(lngBaseRow, COL_DRU联合用药) = Decode(Val(rs种数!抗菌药种数 & ""), 0, "", 1, "Ⅰ种", 2, "Ⅱ联", 3, "Ⅲ联", 4, "Ⅳ联", ">Ⅳ联")
                
                lng药品数 = lng药品数 + Val(rs种数!药品种数 & "")
                lng抗药数 = lng抗药数 + Val(rs种数!抗菌药种数 & "")
                
                dbl抗药金额 = dbl抗药金额 + Val(rs种数!抗菌药金额 & "")
                
                If Val(rs种数!抗菌药种数 & "") <> 0 Then
                    lng用抗菌药人数 = lng用抗菌药人数 + 1
                    
                    Select Case Val(rs种数!抗菌药种数 & "")
                        Case 1
                            lng单用人数 = lng单用人数 + 1
                        Case 2
                            lng二用人数 = lng二用人数 + 1
                        Case 3
                            lng三用人数 = lng三用人数 + 1
                        Case 4
                            lng四用人数 = lng四用人数 + 1
                    End Select
                End If
            End If
            rs种数.Filter = 0
 
            '明细用药
            rs用药明细.Filter = strFilter
            If Not rs用药明细.EOF Then
                For j = 1 To rs用药明细.RecordCount
                    
                    '获取用法用量，存入strTmp 中
                    strTmp = rs用药明细!单次用量 & ""
                    If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
                    strTmp = rs用药明细!执行频次 & "," & strTmp & rs用药明细!计算单位
                
                    If j = 1 Then
                        lngTmpRow = lngBaseRow
                        .TextMatrix(lngBaseRow, COL_DRU药品名称) = rs用药明细!药品名称 & ""
                        .TextMatrix(lngBaseRow, COL_DRU剂型) = rs用药明细!剂型 & ""
                        .TextMatrix(lngBaseRow, COL_DRU规格) = rs用药明细!规格 & ""
                        
                        .TextMatrix(lngTmpRow, COL_DRU用法用量) = strTmp
                        
                        .TextMatrix(lngBaseRow, COL_DRU用药天数) = Val(rs用药明细!天数 & "")
                        .TextMatrix(lngBaseRow, COL_DRU给药途径) = rs用药明细!给药途径 & ""
                        .TextMatrix(lngBaseRow, COL_DRU用药目的) = IIf("" = rs用药明细!目的 & "", "无", rs用药明细!目的)
                    Else
                        .AddItem ""
                        lngTmpRow = .Rows - 1
                        .TextMatrix(lngTmpRow, COL_DRU病人id) = rsTmp!病人ID
                        .TextMatrix(lngTmpRow, COL_DRU主页id) = rsTmp!主页ID
                        
                        .TextMatrix(lngTmpRow, COL_DRU药品名称) = rs用药明细!药品名称 & ""
                        .TextMatrix(lngTmpRow, COL_DRU剂型) = rs用药明细!剂型 & ""
                        .TextMatrix(lngTmpRow, COL_DRU规格) = rs用药明细!规格 & ""
                        .TextMatrix(lngTmpRow, COL_DRU用法用量) = strTmp ' rs用药明细!用法用量 & ""
                        .TextMatrix(lngTmpRow, COL_DRU用药天数) = Val(rs用药明细!天数 & "")
                        .TextMatrix(lngTmpRow, COL_DRU给药途径) = rs用药明细!给药途径 & ""
                        .TextMatrix(lngTmpRow, COL_DRU用药目的) = IIf("" = rs用药明细!目的 & "", "无", rs用药明细!目的)
                    End If
                    
                    '用药天数
                    If Not rs天数 Is Nothing Then
                        rs天数.Filter = "医嘱id=" & rs用药明细!相关ID
                        
                        If Not rs天数.EOF Then
                            strTmp = rs天数!分解时间 & ""
                            If strTmp <> "" Then .TextMatrix(lngTmpRow, COL_DRU用药天数) = UBound(Split(strTmp, ",")) + 1
                        End If
                    End If
                    
                    If InStr("," & str用药的 & ",", "," & .TextMatrix(lngTmpRow, COL_DRU用药目的) & ",") = 0 Then
                        str用药的 = str用药的 & "," & .TextMatrix(lngTmpRow, COL_DRU用药目的)
                    End If
                    rs用药明细.MoveNext
                Next
            End If
            rs用药明细.Filter = 0
            
            If InStr("," & str用药的 & ",", ",预防,") > 0 Then lng预防人数 = lng预防人数 + 1
            If InStr("," & str用药的 & ",", ",治疗,") > 0 Then lng治疗人数 = lng治疗人数 + 1
            
            str用药的 = ""
'-----------------------------------------------------------------------------
            rsTmp.MoveNext
        Next
    End With
        
    '加载统计分析表格
    With vsInDruAna
'        .AddItem "" '第0行
        lngTmp = 0
        .TextMatrix(0, 0) = "A(用药总品种数)=" & lng药品数 & "种"
    
        If lng用抗菌药人数 <> 0 Then dblTmp = lng药品数 / lng用抗菌药人数
        .TextMatrix(0, 1) = "B(平均用药品种数A/E)=" & Format(dblTmp, strDec) & "种": dblTmp = 0
        
        .TextMatrix(0, 2) = "C(使用抗菌药物的品种数)=" & lng抗药数 & "种"
        
        dblTmp = lng用抗菌药人数 * 100 / lng药品数
        .TextMatrix(0, 3) = "D(使用抗菌药物的百分率E/A)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '第1行
        lngTmp = 1
        .TextMatrix(lngTmp, 0) = "E(使用抗药物的病人数)=" & lng用抗菌药人数 & "例"
        
        dblTmp = lng用抗菌药人数 * 100 / lng实际人数
        .TextMatrix(lngTmp, 1) = "F(出院病人抗菌药物使用率E/实际人数)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "G(治疗总金额)=" & dbl总金额 & "元"
        dblTmp = dbl总金额 / lng实际人数
        .TextMatrix(lngTmp, 3) = "H(病人平均治疗金额G/实际人数)=" & Format(dblTmp, strDec) & "元": dblTmp = 0
        
'        .AddItem "" '第2行
        lngTmp = 2
        .TextMatrix(lngTmp, 0) = "I(药品总金额)=" & dbl药品金额 & "元"
        
        If dbl总金额 <> 0 Then dblTmp = dbl药品金额 * 100 / dbl总金额
        .TextMatrix(lngTmp, 1) = "L(药品总金额占治疗总金额的百分率I/G)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "K(抗菌药物总金额)=" & dbl抗药金额 & "元"
        
        If dbl药品金额 <> 0 Then dblTmp = dbl抗药金额 * 100 / dbl药品金额
        .TextMatrix(lngTmp, 3) = "J(抗菌药物总金额占药品总金额的百分率K/I)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
'        .AddItem "" '第3行
        lngTmp = 3
        .TextMatrix(lngTmp, 0) = "M(单用抗菌药物的病人数)=" & lng单用人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng单用人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 1) = "O(单用抗菌药物的使用率M/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "P(二联使用抗菌药物的病人数)=" & lng二用人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng二用人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 3) = "Q(二联使用抗菌药物的使用率P/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

'        .AddItem "" '第4行
        lngTmp = 4
        .TextMatrix(lngTmp, 0) = "R(三联使用抗菌药物的病人数)=" & lng三用人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng三用人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 1) = "S(三联使用抗菌药物的使用率R/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "T(四联使用抗菌药物的病人数)=" & lng四用人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng四用人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 3) = "U(四联使用抗菌药物的使用率T/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

'        .AddItem "" '第5行
        lngTmp = 5
        .TextMatrix(lngTmp, 0) = "V(预防使用抗菌药物的病人数)=" & lng预防人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng预防人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 1) = "W(预防使用抗菌药物构成比V/E)100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0
        
        .TextMatrix(lngTmp, 2) = "X(治疗使用抗菌药物的病人数)=" & lng治疗人数 & "例"
        
        If lng用抗菌药人数 <> 0 Then dblTmp = lng治疗人数 * 100 / lng用抗菌药人数
        .TextMatrix(lngTmp, 3) = "Y(治疗使用抗菌药物构成比Y/E)*100%=" & Format(dblTmp, strDec) & "%": dblTmp = 0

    End With
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "住院医嘱抗菌用药统计", dtpCountS(e_C3_dtpCountS_开始时间_3).Value & "," & dtpCountE(e_C3_dtpCountE_结束时间_3).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsOpeKssUse()
'功能：加载   术后抗菌药物使用超N天统计  界面表格加载
    Dim strSql As String, strPar As String, strDept As String
    Dim rsTmp As ADODB.Recordset
    Dim rs手术 As ADODB.Recordset
    Dim str切口 As String
    Dim lng抽样数 As Long
    Dim strPatis As String
    Dim strFilter As String
    Dim strTmp As String, strTmp1 As String
    Dim i As Long
    Dim intD As Integer, intC As Integer
    Dim lngTmp As Long
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
      
    '如果切口全选或者全不选都认为是不限制切口类型
    str切口 = str切口 & IIf(chkType(e_C4_chkType_切口类型_Ⅰ类_5).Value = 1, "Ⅰ,", "")
    str切口 = str切口 & IIf(chkType(e_C4_chkType_切口类型_Ⅱ类_6).Value = 1, "Ⅱ,", "")
    str切口 = str切口 & IIf(chkType(e_C4_chkType_切口类型_Ⅲ类_7).Value = 1, "Ⅲ,", "")
    str切口 = str切口 & IIf(chkType(e_C4_chkType_切口类型_Ⅳ类_9).Value = 1, "Ⅳ,", "")
    
    If str切口 = "Ⅰ,Ⅱ,Ⅲ,Ⅳ," Then str切口 = ""
    
    lng抽样数 = Val(txtNum(e_C4_txtNum_抽样数量_2).Text)
    
    '日期范围参数生成
    strPar = "To_Date('" & Format(dtpCountS(e_C4_dtpCountS_开始时间_4).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C4_dtpCountE_结束时间_4).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDept = IIf(txtDept(e_C4_txtDept_统计科室_7).Tag = "", "", " and a.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    
    strSql = "select a.病人id,a.主页id,a.住院号,a.姓名 as 患者姓名,d.名称 as 科室,c.使用天数 as 术后用药天数" & _
        " from 病案主页 a,病人抗生素记录 c,部门表 d" & _
        " where a.出院科室id =d.id and a.病人id=c.病人id and a.主页id=c.主页id" & _
        " And a.出院日期 Between " & strPar & strDept & _
        " and c.使用阶段='术后'" & IIf(str切口 = "", "", " and exists (select 1 from 病人手麻记录 M where a.病人id=m.病人id and a.主页id=m.主页id and Instr([2],m.切口)>0)")
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    If optType(e_C4_optType_抽样方法_随机_22).Value Then '随机抽样
        strSql = "select 病人id,主页id,住院号,患者姓名,科室,术后用药天数 from (" & strSql & " Order By Dbms_Random.Value) where rownum < ([3] + 1) Order By 术后用药天数 Desc"
    Else
        strSql = "Select m.病人id, m.主页id, m.住院号, m.患者姓名, m.科室, m.术后用药天数" & vbNewLine & _
            "From (Select m.病人id, m.主页id, m.住院号, m.患者姓名, m.科室, m.术后用药天数" & vbNewLine & _
            "       From (Select m.病人id, m.主页id, m.住院号, m.患者姓名, m.科室, m.术后用药天数, Mod(Rownum,[3]) M" & vbNewLine & _
            "              From (" & strSql & " Order By a.出院日期 desc) M" & vbNewLine & _
            "              Order By M) M" & vbNewLine & _
            "       Where Rownum <([3]+1)) M" & vbNewLine & _
            "Order By m.术后用药天数 Desc"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_C4_txtDept_统计科室_7).Tag, str切口, lng抽样数)
    
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置抽样统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsOpeKssUse
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
            
            .TextMatrix(i, COL_OPE病人id) = rsTmp!病人ID
            .TextMatrix(i, COL_OPE主页id) = rsTmp!主页ID
             
            .TextMatrix(i, COL_OPE住院号) = rsTmp!住院号 & ""
            .TextMatrix(i, COL_OPE患者姓名) = rsTmp!患者姓名 & ""
            .TextMatrix(i, COL_OPE科室) = rsTmp!科室 & ""
            .TextMatrix(i, COL_OPE术后用药天数) = rsTmp!术后用药天数 & ""
            rsTmp.MoveNext
        Next
        
        strPatis = Mid(strPatis, 2)

        strParTable = "Select C1, C2 From Table(f_Num2list2([1]))"
        strTable = strParTable
        
        If Len(strPatis) >= 4000 Then
            varArr = Array()
            varArr = GetParTable(strPar, strParTable, strTable)
        End If
        
        strSql = "select a.病人id,a.主页id,a.已行手术 as 名称,a.切口 from 病人手麻记录 a where (a.病人id,a.主页id) In (" & strTable & ")"
        
        If Len(strPatis) >= 4000 Then
            Set rs手术 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
                    CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        Else
            Set rs手术 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strPatis)
        End If
 
        For i = .FixedRows To .Rows - 1
            strFilter = "病人id=" & Val(.TextMatrix(i, COL_OPE病人id)) & " And 主页id=" & Val(.TextMatrix(i, COL_OPE主页id))
        
            strTmp = "": strTmp1 = ""
            rs手术.Filter = strFilter
            If Not rs手术.EOF Then rs手术.MoveFirst
            Do Until rs手术.EOF
                strTmp = strTmp & ";" & rs手术!名称
                If rs手术!切口 & "" <> "" And InStr("," & strTmp1 & ",", "," & rs手术!切口 & ",") = 0 Then
                    strTmp1 = strTmp1 & "," & rs手术!切口
                End If
                rs手术.MoveNext
            Loop
            .TextMatrix(i, COL_OPE手术名称) = Mid(strTmp, 2)
            .TextMatrix(i, COL_OPE切口类型) = Mid(strTmp1, 2)
        Next
    End With
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "术后抗菌药物使用超N天统计", dtpCountS(e_C4_dtpCountS_开始时间_4).Value & "," & dtpCountE(e_C4_dtpCountE_结束时间_4).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadvsIllDruUse()
'功能：加载   医生治疗某疾病抗菌用药成本统计 界面表析数据
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    
    Dim strParTable As String, strTable As String
    Dim varArr As Variant
    
    Dim rs人数 As ADODB.Recordset
    Dim rs抗菌药数 As ADODB.Recordset
    Dim rs治疗结果 As ADODB.Recordset
    
    Dim strPatis As String, strDec As String
    Dim strTmp As String, strFilter As String
    Dim lng抽样数 As Long
    Dim lngTmp As Long
    Dim dblTmp As Double
    Dim i As Long
    
    lng抽样数 = Val(txtNum(e_C5_txtNum_抽样数量_3).Text)
    
    strTmp = IIf(optType(e_C5_optType_西医_25).Value, 1, 0) & "|" & IIf(optType(e_C5_optType_按疾病_28).Value, 1, 0) & "|" & txtILL.Tag & "|" & txtILL.Text
    Call zlDatabase.SetPara("治疗疾病编码", strTmp, glngSys, 1269)
    
    '日期范围参数生成
    strPar = "To_Date('" & Format(dtpCountS(e_C5_dtpCountS_开始时间_5).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpCountE(e_C5_dtpCountE_结束时间_5).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    strDec = "0.00"
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    strSql = "select a.病人id,a.主页id from 病案主页 a where a.出院日期 between " & strPar & "and exists (" & _
        "select 1 from 病人诊断记录 b where b.病人id=a.病人id and b.主页id=a.主页id and b.记录来源=3 and b.诊断次序=1 And NVL(B.编码序号,1) = 1" & _
        " and b.诊断类型=[1]" & IIf(optType(e_C5_optType_按疾病_28).Value, " and b.疾病id=[2]", " and b.诊断id=[2]") & _
        ")"

    If optType(e_C5_optType_抽样方法_随机_19).Value Then ' 随机抽样
        strSql = "select 病人id,主页id from (" & strSql & "order by Dbms_Random.Value) where rownum<([3] + 1)"
    Else
        strSql = "Select m.病人id,m.主页id From (Select m.病人id,m.主页id,Mod(Rownum,[3]) M From (" & strSql & " order by a.出院日期) m Order By M) m Where Rownum <([3] + 1)"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_西医_25).Value, 3, 13)), txtILL.Tag, lng抽样数)

    If rsTmp.EOF Then
       Screen.MousePointer = 0
       Call zlCommFun.StopFlash
       MsgBox "当前条件下未找到任何数据，请重新设置抽样统计参数。", vbInformation, gstrSysName
       Exit Sub
    End If
    
    For i = 1 To rsTmp.RecordCount
        strPatis = strPatis & "," & rsTmp!病人ID & ":" & rsTmp!主页ID
        rsTmp.MoveNext
    Next
     
    strPatis = Mid(strPatis, 2)
    
    strParTable = "Select C1, C2 From Table(f_Num2list2([3]))"
    strTable = strParTable
    
    If Len(strPatis) >= 4000 Then
        varArr = Array()
        varArr = GetParTable(strPar, strParTable, strTable)
    End If

    strSql = "select nvl(开单人,'空') as 主管医生,count(1) as 治疗人数,Sum(sign(抗菌药人数)) as 抗菌药人数,sum(住院天数) as 住院天数,sum(总费用) as 总费用,sum(总药费) as 总药费,sum(抗菌药费) as 抗菌药费" & _
        " from (select b.开单人,b.病人id,b.主页id,max(a.住院天数) as 住院天数,sum(b.结帐金额) as 总费用,sum(decode(Nvl(e.抗生素, 0),0,0,b.结帐金额)) as 抗菌药费," & _
        " sum(decode(b.收费类别,'5',b.结帐金额,'6',b.结帐金额,'7',b.结帐金额,0)) as 总药费,sum(decode(Nvl(e.抗生素, 0),0,0,1)) as 抗菌药人数" & _
        " from 病案主页 a,住院费用记录 b,药品目录 d,药品特性 e" & _
        " where  a.病人id=b.病人id and a.主页id=b.主页id and b.记录状态<>0 and b.收费细目id=d.药品id(+) and d.药名id=e.药名id(+)" & _
        " and (a.病人id,a.主页id) In (" & strTable & ")" & _
        " group by b.开单人,b.病人id,b.主页id,a.出院日期,a.入院日期)" & _
        " group by 开单人 order by 治疗人数"
    
    If Len(strPatis) >= 4000 Then
        Set rs人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs人数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", strPatis)
    End If
    
    '出院情况
    strSql = "select /*+ rule*/ nvl(开单人,'空') as 主管医生,sum(sign(治愈)) as 治愈,sum(sign(好转)) as 好转,sum(sign(未愈)) as 未愈,sum(sign(死亡)) as 死亡,sum(sign(其它)) as 其它" & _
        " from (select a.开单人,a.病人id,a.主页id,sum(decode(出院情况,'治愈',1,0)) as 治愈,sum(decode(出院情况,'好转',1,0)) as 好转," & _
        " sum(decode(出院情况,'未愈',1,0)) as 未愈,sum(decode(出院情况,'死亡',1,0)) as 死亡," & _
        " sum(decode(出院情况,'死亡',0,'好转',0,'治愈',0,'未愈',0,1)) as 其它" & _
        " from 住院费用记录 a,病人诊断记录 b where a.病人id=b.病人id and a.记录状态<>0 and a.主页id=b.主页id and b.记录来源=3 and b.诊断次序=1  And NVL(B.编码序号,1) = 1" & _
        " and b.诊断类型=[1]" & IIf(optType(e_C5_optType_按疾病_28).Value, " and b.疾病id=[2]", " and b.诊断id=[2]") & _
        " and (a.病人id,a.主页id) In (" & strTable & ")" & _
        " group by a.开单人,a.病人id,a.主页id) group by 开单人"
    If Len(strPatis) >= 4000 Then
        Set rs治疗结果 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_西医_25).Value, 3, 13)), txtILL.Tag, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), _
                CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs治疗结果 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(IIf(optType(e_C5_optType_西医_25).Value, 3, 13)), txtILL.Tag, strPatis)
    End If
    
    strSql = "select nvl(开单人,'空') as 主管医生,count(1) as 抗菌药种数" & _
        " from (select b.开单人,d.药名ID" & _
        " from 住院费用记录 b,药品目录 d,药品特性 e" & _
        " where b.收费细目id=d.药品id(+) and d.药名id=e.药名id(+)" & _
        " and Nvl(e.抗生素, 0)<>0 and b.收费类别='5'" & _
        " and (b.病人id,b.主页id) In (" & strTable & ")" & _
        " group by b.开单人,d.药名ID ) group by 开单人"
        
    If Len(strPatis) >= 4000 Then
        Set rs抗菌药数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    Else
        Set rs抗菌药数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "", "", strPatis)
    End If
    
    With vsIllDruUse
        .Rows = 2
        For i = 1 To rs人数.RecordCount
            .AddItem ""
            
            strFilter = "主管医生='" & rs人数!主管医生 & "'"
            rs治疗结果.Filter = strFilter
            rs抗菌药数.Filter = strFilter
            
            .TextMatrix(.Rows - 1, COL_ILL主管医生) = rs人数!主管医生 & ""
            .TextMatrix(.Rows - 1, COL_ILL治疗人数) = Val(rs人数!治疗人数 & "")
            .TextMatrix(.Rows - 1, COL_ILL用抗药人数) = Val(rs人数!抗菌药人数 & "")
            
            .TextMatrix(.Rows - 1, COL_ILL治愈) = 0
            .TextMatrix(.Rows - 1, COL_ILL好转) = 0
            .TextMatrix(.Rows - 1, COL_ILL未愈) = 0
            .TextMatrix(.Rows - 1, COL_ILL死亡) = 0
            .TextMatrix(.Rows - 1, COL_ILL其它) = 0
            
            strTmp = "": lngTmp = 0
            If Not rs治疗结果.EOF Then rs治疗结果.MoveFirst
            Do Until rs治疗结果.EOF
                lngTmp = Val(rs治疗结果!治愈 & "")
                .TextMatrix(.Rows - 1, COL_ILL治愈) = Val(rs治疗结果!治愈 & "")
                .TextMatrix(.Rows - 1, COL_ILL好转) = Val(rs治疗结果!好转 & "")
                .TextMatrix(.Rows - 1, COL_ILL未愈) = Val(rs治疗结果!未愈 & "")
                .TextMatrix(.Rows - 1, COL_ILL死亡) = Val(rs治疗结果!死亡 & "")
                .TextMatrix(.Rows - 1, COL_ILL其它) = Val(rs治疗结果!其它 & "")
                rs治疗结果.MoveNext
            Loop
            
            If Val(rs人数!治疗人数 & "") <> 0 Then dblTmp = lngTmp * 100 / Val(rs人数!治疗人数 & "")
 
            .TextMatrix(.Rows - 1, COL_ILL治愈率) = Format(dblTmp, strDec) & "%": dblTmp = 0
            
            dblTmp = Val(rs人数!总费用 & "")
            .TextMatrix(.Rows - 1, COL_ILL总金额) = Format(dblTmp, strDec): dblTmp = 0
            
            dblTmp = Val(rs人数!总药费 & "")
            .TextMatrix(.Rows - 1, COL_ILL药品金额) = Format(dblTmp, strDec): dblTmp = 0
            
            If Val(rs人数!治疗人数 & "") <> 0 Then dblTmp = Val(rs人数!总费用 & "") / Val(rs人数!治疗人数 & "")
            .TextMatrix(.Rows - 1, COL_ILL人均治疗额) = Format(dblTmp, strDec): dblTmp = 0
            
            If Val(rs人数!住院天数 & "") <> 0 Then dblTmp = Val(.TextMatrix(.Rows - 1, COL_ILL人均治疗额)) / Val(rs人数!住院天数 & "")
            .TextMatrix(.Rows - 1, COL_ILL人均日金额) = Format(dblTmp, strDec): dblTmp = 0
            
            dblTmp = Val(rs人数!抗菌药费 & "")
            .TextMatrix(.Rows - 1, COL_ILL抗药金额) = Format(dblTmp, strDec): dblTmp = 0
            
            If Not rs抗菌药数.EOF Then rs抗菌药数.MoveFirst: lngTmp = 0
            Do Until rs抗菌药数.EOF
                lngTmp = Val(rs抗菌药数!抗菌药种数 & "")
                rs抗菌药数.MoveNext
            Loop
            .TextMatrix(.Rows - 1, COL_ILL抗药品种数) = lngTmp
            rs人数.MoveNext
        Next
    End With
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "医生治疗某疾病抗菌用药成本统计", dtpCountS(e_C5_dtpCountS_开始时间_5).Value & "," & dtpCountE(e_C5_dtpCountE_结束时间_5).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdILL_Click()
'功能：诊断选择器，调用界面   医生治疗某疾病抗菌用药成本统计
    Dim rsTmp As ADODB.Recordset
     
    If Not optType(e_C5_optType_西医_25).Value Then
        If optType(e_C5_optType_按诊断_27).Value Then
            '按诊断输入:中医部份，一个诊断可能属于多个分类
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", 0, , True, False)
        Else
            'B-中医疾病编码
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", 0, , True)
        End If
    Else
        If optType(e_C5_optType_按诊断_27).Value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", 0, , True, False)
        Else
            'D-ICD-10疾病编码
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "D", 0, , True)
        End If
    End If
     If Not rsTmp Is Nothing Then
        txtILL.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
        cmdILL.Tag = txtILL.Text
        txtILL.Tag = rsTmp!项目ID
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Long, strTmp As String
    
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        For i = 2 To cbsMain.Count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '按钮文字
        For i = 2 To cbsMain.Count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '大图标
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '状态栏
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Refresh
        If GetCur页面 = "病人抗菌用药情况抽样调查及评价表" Then Call LoadPati
    Case conMenu_Tool_Archive '电子病案查阅
        Call Show电子病案查阅
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '退出
        Unload Me
    Case conMenu_File_Print
        Call zlRptPrint(1)
    Case conMenu_File_Preview
        Call zlRptPrint(2)
    Case conMenu_File_Excel
        Call zlRptPrint(3)
    Case conMenu_Help_Web_Home 'Web上的中联
        Call zlHomePage(Me.hwnd)
    Case conMenu_Help_Web_Forum '中联论坛
        Call zlWebForum(Me.hwnd)
    Case conMenu_Help_Web_Mail '发送反馈
        Call zlMailTo(Me.hwnd)
    Case conMenu_Help_About '关于
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_File_PrintSet
        SwitchPrintSet glngSys & "\" & 1269
        Call zlPrintSet
        SwitchPrintSet glngSys & "\" & 1269, True
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            '执行发布到当前模块的报表
            strTmp = Split(Control.Parameter, ",")(1)
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strTmp, Me)
        End If
    End Select
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_View_ToolBar_Button '工具栏
        If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text '图标文字
        If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(GetFirstCommandBar(cbsMain(2).Controls)).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '大图标
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_StatusBar '状态栏
        Control.Checked = Me.stbThis.Visible
    Case conMenu_Tool_Archive
        Control.Enabled = (tbcSub.Selected.Index = 0 And tbcReport.Selected.Index = 1)
    End Select
End Sub

Private Sub cmdOK_Click(Index As Integer)
'功能：点击 统计 按钮
    
    If Not CheckData() Then Exit Sub
    
    If MsgBox("本操作将会非常耗时，并且可能影响系统的整体性能，建议在业务空闲时间运行，你确定要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If PanelItem_抗菌药品消耗金额调查表 = Index Then
        
        Call LoadBill
        
    ElseIf PanelItem_病人抗菌用药情况抽样调查及评价表 = Index Then
    
        If Save抽样记录 Then
            Call LoadPati
        Else
            rptPati.Records.DeleteAll
            rptPati.Populate
            txtCYJL.Tag = ""
            txtCYJL.Text = ""
        End If
        
    ElseIf PanelItem_门诊处方抗菌用药调查表 = Index Then
    
        Call Load处方抽样(vsMZYY, True)
        Call Load处方分析(vsMZYY, vsCF, True)
        
    ElseIf PanelItem_住院病人抗菌用药调查表 = Index Then
    
        Call LoadInKssAdvice
        
    ElseIf PanelItem_抗菌药物使用情况排名统计 = (Index - 4) Then
    
        Call LoadvsUseRan
        
    ElseIf PanelItem_Ⅰ类切口围术期预防用药统计 = (Index - 4) Then
    
        Call LoadvsCut
        
    ElseIf PanelItem_门急诊处方抗菌用药统计 = (Index - 4) Then
    
        Call Load处方抽样(vsCountDruUse, False)
        Call Load处方分析(vsCountDruUse, vsCountCF, False)
        
    ElseIf PanelItem_住院医嘱抗菌用药统计 = (Index - 4) Then
    
        Call LoadvsInDruUse
        
    ElseIf PanelItem_术后抗菌药物使用超N天统计 = (Index - 4) Then
    
        Call LoadvsOpeKssUse
        
    ElseIf PanelItem_医生治疗某疾病抗菌用药成本统计 = (Index - 4) Then
    
        Call LoadvsIllDruUse
    End If
End Sub

Private Function CheckData() As Boolean
'功能：加载数据前的检查

    Select Case GetCur页面
        Case "抗菌药品消耗金额调查表"
            If dtpRPS(e_R0_dtpRPS_开始时间_0).Value >= dtpRPE(e_R0_dtpRPE_结束时间_0).Value Then
                MsgBox "开始时间应该小于结束时间。", vbInformation, gstrSysName
                dtpRPE(e_R0_dtpRPE_结束时间_0).SetFocus
                Exit Function
            End If
        Case "病人抗菌用药情况抽样调查及评价表"
            '抽样时间的一个检查，开始时间小于结束时间
            If dtpRPS(e_R1_dtpRPS_开始时间_1).Value >= dtpRPE(e_R1_dtpRPE_结束时间_1).Value Then
                MsgBox "开始时间应该小于结束时间。", vbInformation, gstrSysName
                dtpRPE(e_R1_dtpRPE_结束时间_1).SetFocus
                Exit Function
            End If
    
            If Val(txtCount(e_R1_txtCount_抽样数量_1).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtCount(e_R1_txtCount_抽样数量_1).SetFocus
                Exit Function
            End If
        Case "门诊处方抗菌用药调查表"
            If dtpRPS(e_R2_dtpRPS_开始时间_2).Value > dtpRPE(e_R2_dtpRPE_结束时间_2).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpRPE(e_R2_dtpRPE_结束时间_2).SetFocus
                Exit Function
            End If
            
            If Val(txtCount(e_R2_txtCount_抽样数量_2).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtCount(e_R2_txtCount_抽样数量_2).SetFocus
                Exit Function
            End If
        Case "住院病人抗菌用药调查表"
            If dtpRPS(e_R3_dtpRPS_开始时间_3).Value > dtpRPE(e_R3_dtpRPE_结束时间_3).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpRPE(e_R3_dtpRPE_结束时间_3).SetFocus
                Exit Function
            End If
        Case "抗菌药物使用情况排名统计"
            If dtpCountS(e_C0_dtpCountS_开始时间_0).Value > dtpCountE(e_C0_dtpCountE_结束时间_0).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpCountE(e_C0_dtpCountE_结束时间_0).SetFocus
                Exit Function
            ElseIf Not IsNumeric(txtTopRan.Text) Then
                MsgBox "统计名次必需是大于零的整数。", vbInformation, gstrSysName
                txtTopRan.SetFocus
                Exit Function
            ElseIf Val(txtTopRan.Text) <= 0 Then
                MsgBox "统计名次必需是大于零的整数。", vbInformation, gstrSysName
                txtTopRan.SetFocus
                Exit Function
            End If
        Case "Ⅰ类切口围术期预防用药统计"
            If dtpCountS(e_C1_dtpCountS_开始时间_1).Value > dtpCountE(e_C1_dtpCountE_结束时间_1).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpCountE(e_C1_dtpCountE_结束时间_1).SetFocus
                Exit Function
            End If
        Case "门急诊处方抗菌用药统计"
            If dtpCountS(e_C2_dtpCountS_开始时间_2).Value > dtpCountE(e_C2_dtpCountE_结束时间_2).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpCountE(e_C2_dtpCountE_结束时间_2).SetFocus
                Exit Function
            End If
            
            If Val(txtNum(e_C2_txtNum_统计科室_0).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtNum(e_C2_txtNum_统计科室_0).SetFocus
                Exit Function
            End If
        Case "住院医嘱抗菌用药统计"
            If dtpCountS(e_C3_dtpCountS_开始时间_3).Value > dtpCountE(e_C3_dtpCountE_结束时间_3).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpCountE(e_C3_dtpCountE_结束时间_3).SetFocus
                Exit Function
            End If
            If Val(txtNum(e_C3_txtNum_抽样数量_1).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtNum(e_C3_txtNum_抽样数量_1).SetFocus
                Exit Function
            End If
            
            If optType(e_C3_optType_切口类型_手术_18).Value Then
                If chkType(e_C3_chkType_切口类型_Ⅰ类_2).Value <> 1 And chkType(e_C3_chkType_切口类型_Ⅱ类_3).Value <> 1 And _
                    chkType(e_C3_chkType_切口类型_Ⅲ类_4).Value <> 1 And chkType(e_C3_chkType_切口类型_Ⅳ类_8).Value <> 1 Then
                    MsgBox "请选择一种切口类型。", vbInformation, gstrSysName
                End If
            End If
            
        Case "术后抗菌药物使用超N天统计"
            If Val(txtNum(e_C4_txtNum_抽样数量_2).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtNum(e_C4_txtNum_抽样数量_2).SetFocus
                Exit Function
            End If
            
        Case "医生治疗某疾病抗菌用药成本统计"
            If dtpCountS(e_C5_dtpCountS_开始时间_5).Value > dtpCountE(e_C5_dtpCountE_结束时间_5).Value Then
                MsgBox "开始时间小于结束时间。", vbInformation, gstrSysName
                dtpCountE(e_C5_dtpCountE_结束时间_5).SetFocus
                Exit Function
            End If
            
            If Val(txtNum(e_C5_txtNum_抽样数量_3).Text) = 0 Then
                MsgBox "抽样数量不能为零。", vbInformation, gstrSysName
                txtNum(e_C5_txtNum_抽样数量_3).SetFocus
                Exit Function
            End If
            
            If txtILL.Text = "" Then
                MsgBox "请选择一种疾病。", vbInformation, gstrSysName
                txtILL.SetFocus
                Exit Function
            End If
    End Select
    
    CheckData = True
End Function

Private Sub cmdCYDel_Click()
    Dim blnTrans As Boolean
    
    If txtCYJL.Tag = "" Then Exit Sub
    If MsgBox("你确定要删除本次抽样记录吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure("Zl_抗菌药物抽样记录_Delete(" & txtCYJL.Tag & ")", Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    txtCYJL.Tag = ""
    txtCYJL.Text = ""
    rptPati.Records.DeleteAll
    rptPati.Populate
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCYEdit_Click()
    
    If rptPati.SelectedRows.Count = 0 Then
        MsgBox "未选择中一个病人！", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPati.SelectedRows(0).GroupRow Then
        MsgBox "未选择中一个病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rptPati.SelectedRows.Count > 0 Then Call Edit病人调查
  
End Sub

Private Sub Edit病人调查()
'功能：进入抽样调查表

    Dim bln编辑 As Boolean
    Dim bln打印 As Boolean
    Dim blnTmp As Boolean
    
    With rptPati.SelectedRows(0)
        If Not .GroupRow Then
            If .Record(COL_编辑).Value = "√" Then bln编辑 = True
            If .Record(COL_打印).Value = "√" Then bln打印 = True
            blnTmp = frmKssSurveyEdit.ShowMe(Me, .Record(COL_抽样ID).Value, .Record(col_病人Id).Value, .Record(col_主页ID).Value, .Record(COL_序号).Value, _
                IIf(Val(.Record(COL_手术ID).Value) = 0, mlng非手术病人数, mlng手术病人数), .Record(col_科室).Value, Val(.Record(COL_手术ID).Value) > 0, bln编辑, bln打印)
            If blnTmp Then
                If bln编辑 Then .Record(COL_编辑).Value = "√"
                If bln打印 Then .Record(COL_打印).Value = "√"
                rptPati.Populate
            End If
        End If
    End With
End Sub

Private Sub Show电子病案查阅()
    If rptPati.SelectedRows.Count = 0 Then
        MsgBox "未选择中一个病人！", vbInformation, gstrSysName
        Exit Sub
    End If

    If rptPati.SelectedRows(0).GroupRow Then
        MsgBox "未选择中一个病人！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If rptPati.SelectedRows.Count > 0 Then
        With rptPati.SelectedRows(0)
            If Not .GroupRow Then
                Call frmArchiveView.ShowArchive(Me, Val(.Record(col_病人Id).Value), Val(.Record(col_主页ID).Value))
            End If
        End With
    End If
End Sub

Private Function Save抽样记录() As Boolean
    Dim strSql As String
    Dim blnTrans As Boolean
    Dim strCurDate As String
    
    On Error GoTo errH
    
    mlng抽样ID = zlDatabase.GetNextId("抗菌药物抽样记录")
    mdatCurr = zlDatabase.Currentdate
    strCurDate = Format(mdatCurr, "yyyy-MM-dd hh:mm:ss")
    
    strSql = "Zl_抗菌药物抽样记录_Insert(" & mlng抽样ID & ",'" & UserInfo.姓名 & "'," & _
       "to_date('" & Format(dtpRPS(e_R1_dtpRPS_开始时间_1).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')," & _
       "to_date('" & Format(dtpRPE(e_R1_dtpRPE_结束时间_1).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')," & _
       Val(txtCount(e_R1_txtCount_抽样数量_1).Text) & "," & IIf(optType(e_R1_optType_抽样方法_平均_0).Value, 0, 1) & "," & _
       IIf(txtDept(e_R1_txtDept_抽样科室_1).Tag = "", "NULL,", "'" & txtDept(e_R1_txtDept_抽样科室_1).Tag & "',") & _
       "to_date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'))"
    
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    
    MsgBox "本次抽样完成！", vbInformation, gstrSysName
    
    txtCYJL.Text = "抽样时间：" & strCurDate & "  抽样人：" & UserInfo.姓名
    txtCYJL.Tag = mlng抽样ID
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "病人抗菌用药情况抽样调查及评价表", dtpRPS(e_R1_dtpRPS_开始时间_1).Value & "," & dtpRPE(e_R1_dtpRPE_结束时间_1).Value
    
    Save抽样记录 = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadBill()
'功能：加载   抗菌药品消耗金额调查表   的数据，查询耗时长，所以读到数据立马加载到界面上，让结果逐步显示
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    Dim dblTmp As Double
    Dim strDec As String '费用的精度，4位小数
    Dim i As Long
    
    Dim dblTotal As Double ' "一、年医院总收入（金额）"
    Dim dbl差价 As Double  '"五、药品进销差价收入（金额）"
    Dim dbl总药费 As Double
    Dim dbl住院西药费 As Double
    Dim dbl门诊西药费 As Double
    Dim dbl住院西药费抗 As Double
    Dim dbl门诊西药费抗 As Double
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    strDec = "0.0000"
    
    '表格中其它数据
    With vsBill
        For i = 1 To 13
            .TextMatrix(i, COL_单位) = "万元"
        Next
        .TextMatrix(ROW_年医院总收入, COL_备注) = "不含政府拨款"
        .TextMatrix(ROW_药品占医院总收入比例, COL_单位) = "%"
        .TextMatrix(ROW_药品进销差价收入占医院总收入比例, COL_单位) = "%"
        .TextMatrix(ROW_抗菌药物占药品总收入比例, COL_单位) = "%"
        .Cell(flexcpText, 1, COL_结果, 13, COL_结果) = strDec '未生成结果时，值都设为 0.00
    End With
    
    '日期范围参数生成
    strPar = "To_Date('" & Format(dtpRPS(e_R0_dtpRPS_开始时间_0).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
        " and To_Date('" & Format(dtpRPE(e_R0_dtpRPE_结束时间_0).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '从汇总表中查询出医院总费用  "一、年医院总收入（金额）"
    strSql = "select sum(a.结帐金额)/10000 as 总收入 from 病人费用汇总 a  where a.日期 between " & strPar
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then dblTotal = Val(rsTmp!总收入 & "")
    
    vsBill.TextMatrix(ROW_年医院总收入, COL_结果) = Format(dblTotal, strDec)
    
    '差价   总药费
    strSql = "select sum(a.差价)/10000 as 差价,-1*sum(decode(a.单据,8,a.金额,9,a.金额,10,a.金额,0))/10000 as 药费 from 药品收发汇总 a where a.单据<14 and a.日期 Between " & strPar
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        dbl差价 = Val(rsTmp!差价 & "")
        dbl总药费 = Val(rsTmp!药费 & "")
    End If
    vsBill.TextMatrix(ROW_药品进销差价收入, COL_结果) = Format(dbl差价, strDec)
    vsBill.TextMatrix(ROW_年药品总收入, COL_结果) = Format(dbl总药费, strDec)
    
    '药品进销差价收入占医院总收入比例
    If dblTotal <> 0 Then
        dblTmp = dbl差价 * 100 / dblTotal
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_药品进销差价收入占医院总收入比例, COL_结果) = Format(dblTmp, strDec)
        End If
    End If
    
    '药品占医院总收入比例
    If dblTotal <> 0 Then
        dblTmp = dbl总药费 * 100 / dblTotal
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_药品占医院总收入比例, COL_结果) = Format(dblTmp, strDec)
        End If
    End If
    
    '门诊费用记录   西药费  抗菌药费
    strSql = "select sum(a.药费)/10000 as 门诊西药费,Sum(Decode(Nvl(c.抗生素, 0), 0, 0, a.药费))/10000 As 门诊西药抗菌药费" & _
        " from (Select x.收费细目id,Sum(x.结帐金额) As 药费 From 门诊费用记录 X Where x.发生时间 Between " & strPar & _
        " And x.记录状态 <> 0 And x.收费类别='5' group by x.收费细目id) a," & _
        " 药品规格 B, 药品特性 C where a.收费细目id = b.药品id And b.药名id = c.药名id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    
    If Not rsTmp.EOF Then
        dbl门诊西药费 = Val(rsTmp!门诊西药费 & "")
        dbl门诊西药费抗 = Val(rsTmp!门诊西药抗菌药费 & "")
    End If
    
    vsBill.TextMatrix(ROW_门诊西药房, COL_结果) = Format(dbl门诊西药费, strDec)
    vsBill.TextMatrix(ROW_门诊西药房抗, COL_结果) = Format(dbl门诊西药费抗, strDec)
    
    '住院费用记录   西药费  抗菌药费
    strSql = "select sum(a.药费)/10000 as 住院西药费,Sum(Decode(Nvl(c.抗生素, 0), 0, 0, a.药费))/10000 As 住院西药抗菌药费" & _
        " from (Select x.收费细目id,Sum(x.结帐金额) As 药费 From 住院费用记录 X Where x.发生时间 Between " & strPar & _
        " And x.记录状态 <> 0 And x.收费类别='5' group by x.收费细目id) a," & _
        " 药品规格 B, 药品特性 C where a.收费细目id = b.药品id And b.药名id = c.药名id"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    If Not rsTmp.EOF Then
        dbl住院西药费 = Val(rsTmp!住院西药费 & "")
        dbl住院西药费抗 = Val(rsTmp!住院西药抗菌药费 & "")
    End If
    
    vsBill.TextMatrix(ROW_住院西药房, COL_结果) = Format(dbl住院西药费, strDec)
    vsBill.TextMatrix(ROW_住院西药房抗, COL_结果) = Format(dbl住院西药费抗, strDec)
    
    '西药全年使用金额
    dblTmp = dbl住院西药费 + dbl门诊西药费
    vsBill.TextMatrix(ROW_西药全年使用金额, COL_结果) = Format(dblTmp, strDec)
    
    '抗菌药物全年使用金额
    dblTmp = dbl住院西药费抗 + dbl门诊西药费抗
    vsBill.TextMatrix(ROW_抗菌药物全年使用金额, COL_结果) = Format(dblTmp, strDec)
    
    '抗菌药物占药品总收入比例
    If dbl总药费 <> 0 Then
        dblTmp = dblTmp * 100 / dbl总药费
        If dblTmp <> 0 Then
            vsBill.TextMatrix(ROW_抗菌药物占药品总收入比例, COL_结果) = Format(dblTmp, strDec)
        End If
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "抗菌药品消耗金额调查表", dtpRPS(e_R0_dtpRPS_开始时间_0).Value & "," & dtpRPE(e_R0_dtpRPE_结束时间_0).Value
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadPati()
'功能：加截抽样病人列表
    Dim rsTmp As ADODB.Recordset
    Dim objParent As ReportRecord
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strTmp As String
    Dim objRow As ReportRow
    Dim lngCount As Long
    Dim strSql As String
    Dim i As Long
    
    rptPati.Records.DeleteAll
    
    strSql = "Select b.抽样id,b.序号,b.病人id,b.主页id,b.是否打印,b.是否编辑,a.姓名,a.性别,a.年龄,a.住院号,a.出院科室id,d.名称 As 出院科室,a.住院医师,a.出院日期,max(e.Id) as 手术id" & _
        " From 病案主页 A, 抗菌药物抽样明细 B,部门表 D, 病人手麻记录 E" & vbNewLine & _
        " Where a.病人id = b.病人id And a.主页id = b.主页id And b.抽样id =[1] And a.出院科室id = d.Id And a.病人id = e.病人id(+)" & _
        " And a.主页id = e.主页id(+)" & _
        " group by b.抽样id,b.序号,b.病人id,b.主页id,a.姓名,a.性别,a.年龄,a.住院号,a.出院科室id,d.名称,a.住院医师,a.出院日期,b.是否打印,b.是否编辑 Order By b.序号"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(txtCYJL.Tag))
    
    '存实际抽样的人数
    lblN(e_R1_lblN_抽样数量标题_13).Tag = rsTmp.RecordCount
    
    For i = 1 To rsTmp.RecordCount
        Set objRecord = Me.rptPati.Records.Add()
        objRecord.Tag = CStr(rsTmp!病人ID & "," & rsTmp!主页ID) '用于病人定位
        
        If Val(rsTmp!手术id & "") > 0 Then lngCount = lngCount + 1
        
        strTmp = IIf(Val(rsTmp!是否编辑 & "") = 0, "×", "√")
        objRecord.AddItem strTmp '是否编辑过抽样表
        
        strTmp = IIf(Val(rsTmp!是否打印 & "") = 0, "×", "√")
        objRecord.AddItem strTmp '是否打印列图标
        
        Set objItem = objRecord.AddItem(IIf(Val(rsTmp!手术id & "") > 0, "手术", "非手术"))   '分组以Value进行排序
            objItem.Caption = IIf(Val(rsTmp!手术id & "") > 0, "手术病人", "非手术病人")
        
        objRecord.AddItem CStr(Nvl(rsTmp!姓名))
        objRecord.AddItem CStr(Nvl(rsTmp!性别))
        objRecord.AddItem CStr(Nvl(rsTmp!年龄))
        objRecord.AddItem CStr(Nvl(rsTmp!住院号))
        objRecord.AddItem CStr(Nvl(rsTmp!出院科室))
        objRecord.AddItem CStr(Nvl(rsTmp!住院医师))
        objRecord.AddItem Format(rsTmp!出院日期 & "", "yyyy-mm-dd hh:mm:ss")
        
        objRecord.AddItem Val(rsTmp!病人ID)
        objRecord.AddItem Val(rsTmp!主页ID)
        objRecord.AddItem Val(rsTmp!抽样ID)
        objRecord.AddItem Val(rsTmp!序号)
        objRecord.AddItem Val(rsTmp!手术id & "")
        
        rsTmp.MoveNext
    Next
    
    mlng手术病人数 = lngCount
    mlng非手术病人数 = rsTmp.RecordCount - lngCount
    
    rptPati.Populate
    
    With rptPati.Columns
        .Column(col_姓名).Width = 60
        .Column(col_性别).Width = 30
        .Column(col_年龄).Width = 60
        .Column(col_住院号).Width = 70
        .Column(col_科室).Width = 100
        .Column(col_住院医师).Width = 60
        .Column(col_出院日期).Width = 140
        .Column(col_出院日期).Alignment = xtpAlignmentCenter
        .Column(col_年龄).Alignment = xtpAlignmentLeft
    End With
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Load处方抽样(ByRef vsgInfo As VSFlexGrid, ByVal blnRP As Boolean)
'功能：门诊处方抽样，急诊或非急诊
'参数： vsgInfo 表格对象，上报数据－－vsMZYY  其它统计 vsCountDruUse
'       blnRP 界面，true 上报数据，false 其它统计
    Dim strSql As String, strPar As String
    Dim str急诊 As String, strDept  As String
    Dim lngBaseRow As Long, strDec As String
    Dim i As Long, j As Long, k As Long
    Dim strTableIn As String, strTableOut As String
    Dim varArr As Variant
    Dim rs抽样 As ADODB.Recordset
    Dim rs科室 As ADODB.Recordset
    Dim rs处方金额 As ADODB.Recordset
    Dim rs药品数量 As ADODB.Recordset
    Dim rs诊断 As ADODB.Recordset
    Dim rs医嘱处方 As ADODB.Recordset
    Dim rs费用处方 As ADODB.Recordset
    Dim rs费用抗药明细 As ADODB.Recordset
    Dim rs医嘱抗药明细 As ADODB.Recordset
    Dim strPar医嘱处方 As String
    Dim strPar费用处方 As String
    Dim strDeptIDs As String
    Dim strTmp As String
    Dim lng抽样数量 As Long
    Dim bln抽样方式 As Boolean '平均或随机抽样, bln抽样方式 true 平均抽样 false 随机抽样
    Dim bln注射 As Boolean
    Dim lngTmp As Long
    Dim dblTmp As Double
    
    strDec = "0.00"
    
    If blnRP Then   '日期范围参数生成
        strPar = "To_Date('" & Format(dtpRPS(e_R2_dtpRPS_开始时间_2).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " and To_Date('" & Format(dtpRPE(e_R2_dtpRPE_结束时间_2).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        
        lblN(e_R2_lblN_调查表_处方总量_27).Caption = "处方总量：0张"
        lblN(e_R2_lblN_分析表_处方总量_30).Caption = lblN(e_R2_lblN_调查表_处方总量_27).Caption
        lblN(e_R2_lblN_分析表_标题_70).Caption = "0张处方统计分析表"
        lblN(e_R2_lblN_调查表_日期_26).Caption = "日期：" & Format(dtpRPS(e_R2_dtpRPS_开始时间_2).Value, "YYYY-MM-DD") & "至" & Format(dtpRPE(e_R2_dtpRPE_结束时间_2).Value, "YYYY-MM-DD")
        lblN(e_R2_lblN_分析表_日期_29).Caption = lblN(e_R2_lblN_调查表_日期_26).Caption
    Else
        strPar = "To_Date('" & Format(dtpCountS(e_C2_dtpCountS_开始时间_2).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
            " and To_Date('" & Format(dtpCountE(e_C2_dtpCountE_结束时间_2).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
        
        lblN(e_C2_lblN_统计表_处方总量_48).Caption = "处方总量：0张"
        lblN(e_C2_lblN_分析表_处方总量_49).Caption = lblN(e_C2_lblN_统计表_处方总量_48).Caption
        lblN(e_C2_lblN_分析表_标题_39).Caption = "0张处方统计分析表"
        lblN(e_C2_lblN_统计表_日期_46).Caption = "日期：" & Format(dtpCountS(e_C2_dtpCountS_开始时间_2).Value, "YYYY-MM-DD") & "至" & Format(dtpCountE(e_C2_dtpCountE_结束时间_2).Value, "YYYY-MM-DD")
        lblN(e_C2_lblN_分析表_日期_47).Caption = lblN(e_C2_lblN_统计表_日期_46).Caption
    End If
    
    strDeptIDs = IIf(blnRP, txtDept(e_R2_txtDept_抽样科室_2).Tag, txtDept(e_C2_txtDept_统计科室_5).Tag)
    lng抽样数量 = IIf(blnRP, Val(txtCount(e_R2_txtCount_抽样数量_2).Text), Val(txtNum(e_C2_txtNum_统计科室_0).Text))
    bln抽样方式 = IIf(blnRP, optType(e_R2_optType_抽样方法_平均_3).Value, optType(e_C2_optType_抽样方法_平均_6).Value)     ' 平均抽样
    strDept = IIf(strDeptIDs = "", "", " and a.开单部门id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    '清空原有数据
    vsgInfo.Rows = vsgInfo.FixedRows
    vsgInfo.Rows = vsgInfo.FixedRows + 1

    '抽样SQL组织方式，两个界面如果都不区分急诊时抽样的为1类，只抽急诊时是另2类，不含急诊的为3类。费用类处方统一认为是非急诊处方。
    '科室条件的使用，费用类处方则为开单科室，医嘱类处方则为 执行部门 病人挂号记录.执行部门id
    str急诊 = ""
    If Not blnRP Then
        If chkType(e_C2_chkType_类型_门诊_0).Value = 1 And chkType(e_C2_chkType_类型_急诊_1).Value = 0 Then        '3类 不含急诊
            str急诊 = " and (nvl(a.是否急诊,0)<>1 and (a.医嘱序号 is null or exists (select 1 from 病人医嘱记录 B, 病人挂号记录 C where a.医嘱序号 = b.Id and b.挂号单 = c.No and nvl(c.急诊,0)<>1)))"
        ElseIf chkType(e_C2_chkType_类型_门诊_0).Value = 0 And chkType(e_C2_chkType_类型_急诊_1).Value = 1 Then '2类 只含急诊
            str急诊 = " and (nvl(a.是否急诊,0)=1 or exists (select 1 from 病人医嘱记录 B, 病人挂号记录 C where a.医嘱序号 = b.Id and  b.挂号单 = c.No and nvl(c.急诊,0)=1))"
        End If
    End If
 
    If bln抽样方式 Then '平均抽样
        strSql = "select  a.标识号,a.医嘱,a.门诊号,a.病人姓名,a.就诊日期,a.处方医生,a.科室id,a.年龄,a.药费" & vbNewLine & _
            "from (select a.标识号,a.医嘱,a.门诊号,a.病人姓名,a.就诊日期,a.处方医生,a.科室id,a.年龄,a.药费,Mod(Rownum,[2]) M" & vbNewLine & _
            "from (Select a.No As 标识号, Decode(Nvl(Max(a.医嘱序号), 0), 0, 0, 1) As 医嘱, a.标识号 As 门诊号, a.姓名 As 病人姓名," & vbNewLine & _
            "       To_Char(Min(a.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 就诊日期, a.开单人 As 处方医生, a.开单部门id As 科室id, a.年龄,sum(a.结帐金额) as 药费" & vbNewLine & _
            "From 门诊费用记录 A" & vbNewLine & _
            "Where a.记录状态 <> 0 And a.收费类别 In ('5','6','7') And" & vbNewLine & _
            "      a.发生时间 Between " & strPar & strDept & str急诊 & vbNewLine & _
            "Group By a.No, a.标识号, a.姓名, a.开单人, a.开单部门id, a.年龄 having sum(a.结帐金额)>0 order by Min(a.发生时间) desc) a" & vbNewLine & _
            "order by M) a where rownum<([2]+1)"
    Else
        strSql = "select  a.标识号,a.医嘱,a.门诊号,a.病人姓名,a.就诊日期,a.处方医生,a.科室id,a.年龄,a.药费" & vbNewLine & _
            "from (select a.标识号,a.医嘱,a.门诊号,a.病人姓名,a.就诊日期,a.处方医生,a.科室id,a.年龄,a.药费" & vbNewLine & _
            "from (select a.标识号,a.医嘱,a.门诊号,a.病人姓名,a.就诊日期,a.处方医生,a.科室id,a.年龄,a.药费" & vbNewLine & _
            "from (Select a.No As 标识号, Decode(Nvl(Max(a.医嘱序号), 0), 0, 0, 1) As 医嘱, a.标识号 As 门诊号, a.姓名 As 病人姓名," & vbNewLine & _
            "       To_Char(Min(a.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 就诊日期, a.开单人 As 处方医生, a.开单部门id As 科室id, a.年龄,sum(a.结帐金额) as 药费" & vbNewLine & _
            "From 门诊费用记录 A" & vbNewLine & _
            "Where a.记录状态 <> 0 And a.收费类别 In ('5','6','7') And" & vbNewLine & _
            "      a.发生时间 Between " & strPar & strDept & str急诊 & vbNewLine & _
            "Group By a.No, a.标识号, a.姓名, a.开单人, a.开单部门id, a.年龄 having sum(a.结帐金额)>0 ) a" & vbNewLine & _
            "order by Dbms_Random.Value) a where rownum<([2]+1)) a order by a.就诊日期 desc"
    End If
    
    Set rs抽样 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptIDs, lng抽样数量)

    If rs抽样.EOF Then
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置抽样统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If blnRP Then
        lblN(e_R2_lblN_调查表_处方总量_27).Caption = "处方总量：" & rs抽样.RecordCount & "张"
        lblN(e_R2_lblN_分析表_处方总量_30).Caption = lblN(e_R2_lblN_调查表_处方总量_27).Caption
        lblN(e_R2_lblN_分析表_标题_70).Caption = rs抽样.RecordCount & "张处方统计分析表"
    Else
        lblN(e_C2_lblN_统计表_处方总量_48).Caption = "处方总量：" & rs抽样.RecordCount & "张"
        lblN(e_C2_lblN_分析表_处方总量_49).Caption = lblN(e_C2_lblN_统计表_处方总量_48).Caption
        lblN(e_C2_lblN_分析表_标题_39).Caption = rs抽样.RecordCount & "张处方统计分析表"
    End If
    
    '处方明细SQL，抽样完成后数据变得相当少
    '参数收集
    strPar = "": strDeptIDs = ""
    For i = 1 To rs抽样.RecordCount
        If Val(rs抽样!医嘱 & "") = 0 Then
            strPar费用处方 = strPar费用处方 & "," & rs抽样!标识号
        Else
            strPar医嘱处方 = strPar医嘱处方 & "," & rs抽样!标识号
        End If
        strPar = strPar & "," & rs抽样!标识号
        If InStr("," & strDeptIDs & ",", "," & rs抽样!科室ID & ",") = 0 Then
            strDeptIDs = strDeptIDs & "," & rs抽样!科室ID
        End If
        rs抽样.MoveNext
    Next
    rs抽样.MoveFirst
    
    strDeptIDs = Mid(strDeptIDs, 2) '科室id串是不会超长的数据少
    strSql = "select id as 科室id,名称 as 科室 from 部门表 where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))"
    Set rs科室 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strDeptIDs)

    'NO 合成的参数可能要超长这里要处理下
    strPar = Mid(strPar, 2)
    strTableIn = "Select Column_Value From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist))"
    varArr = Array()
    varArr = GetParTable(strPar, strTableIn, strTableOut)
    strSql = "select a.No As 标识号,Sum(a.结帐金额) As 处方金额 From 门诊费用记录 A where a.记录状态 <> 0 And a.no in (" & strTableOut & ") Group By a.No"
    Set rs处方金额 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    
    strSql = "Select a.标识号, Sum(Sign(a.基本药)) As 基本药种数, Sum(Sign(a.抗生药)) As 抗生药种数, Sum(Sign(a.药品)) As 药品种数, Sum(a.抗菌药费) As 抗菌药费" & vbNewLine & _
        "From (Select a.标识号, a.药名id, Sum(a.基本药) As 基本药, Sum(a.抗生药) As 抗生药, Sum(a.药品) As 药品, Sum(a.抗菌药费) As 抗菌药费" & vbNewLine & _
        "       From (Select a.No As 标识号, c.药名id, Decode(Nvl(b.基本药物, '0'), '0', 0, 1) As 基本药, Decode(Nvl(c.抗生素, 0), 0, 0, 1) As 抗生药," & vbNewLine & _
        "                     Decode(a.收费类别, '5', 1, '6', 1, '7', 1, 0) As 药品, Sum(Decode(Nvl(c.抗生素, 0), 0, 0, a.结帐金额)) As 抗菌药费" & vbNewLine & _
        "              From 门诊费用记录 A, 药品规格 B, 药品特性 C" & vbNewLine & _
        "              Where a.收费细目id = b.药品id And b.药名id = c.药名id And a.记录状态 <> 0 And a.收费类别 In ('5', '6', '7') and a.no in (" & strTableOut & ")" & vbNewLine & _
        "              Group By a.No, c.药名id, a.收费类别, Nvl(b.基本药物, '0'), Nvl(c.抗生素, 0)) A" & vbNewLine & _
        "       Group By a.标识号, a.药名id) A" & vbNewLine & _
        "Group By a.标识号"
    Set rs药品数量 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    
    If strPar费用处方 <> "" Then
    
        strPar费用处方 = Mid(strPar费用处方, 2)
        varArr = Array()
        varArr = GetParTable(strPar费用处方, strTableIn, strTableOut)
        
        '费用处方不包含抗菌使用明细，无用法量等
        strSql = "Select a.No As 标识号, 0 As 医嘱, a.收费细目id, f.名称 As 药品通用名, f.规格 || f.产地 As 规格, Sum(a.结帐金额) As 费用," & vbNewLine & _
            "       Sum(a.数次) * b.剂量系数 || b.门诊单位 As 数量" & vbNewLine & _
            "From 门诊费用记录 A, 药品规格 B, 药品特性 C, 收费项目目录 F" & vbNewLine & _
            "Where a.记录状态 <> 0 And a.收费细目id = b.药品id And b.药名id = c.药名id And b.药品id = f.Id And" & vbNewLine & _
            "      a.No In (" & strTableOut & ") And a.收费类别 = '5' And Nvl(c.抗生素, 0) <> 0 And a.医嘱序号 Is Null" & vbNewLine & _
            "Group By a.No, a.收费细目id, f.名称, f.规格, f.产地, b.剂量系数, b.门诊单位"

        Set rs费用抗药明细 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    End If
    
    If strPar医嘱处方 <> "" Then
        strPar医嘱处方 = Mid(strPar医嘱处方, 2)
        varArr = Array()
        varArr = GetParTable(strPar医嘱处方, strTableIn, strTableOut)
        '抗菌药物用药级明细
        strSql = "Select a.No As 标识号, 1 As 医嘱, f.名称 As 药品通用名, f.规格 || f.产地 As 规格, Sum(a.结帐金额) As 费用, Sum(a.数次) * b.剂量系数 || b.门诊单位 As 数量," & vbNewLine & _
            "       e.执行频次,e.单次用量,i.计算单位,g.医嘱内容 As 用药途径" & vbNewLine & _
            "From 门诊费用记录 A, 药品规格 B, 药品特性 C, 收费项目目录 F,病人医嘱记录 E, 诊疗项目目录 I, 病人医嘱记录 G" & vbNewLine & _
            "Where a.记录状态 <> 0 And a.收费细目id = b.药品id And b.药名id = c.药名id And b.药品id = f.Id And e.相关id = g.Id And" & vbNewLine & _
            "      e.诊疗项目id = i.Id And a.No In (" & strTableOut & ")  And" & vbNewLine & _
            "      e.Id = a.医嘱序号 And a.收费类别 = '5' And Nvl(c.抗生素, 0) <> 0" & vbNewLine & _
            "Group By a.No, f.名称, f.规格, f.产地, b.剂量系数, b.门诊单位, e.执行频次, e.单次用量, i.计算单位, g.医嘱内容"

        Set rs医嘱抗药明细 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
        
        '提取诊断
        strSql = "select a.no as 标识号,d.诊断描述" & vbNewLine & _
            "from 门诊费用记录 a,病人医嘱记录 b,病人挂号记录 c,病人诊断记录 d" & vbNewLine & _
            "where a.医嘱序号=b.id and b.挂号单=c.no and c.病人id=d.病人id and c.id=d.主页id and a.记录状态 <> 0 and a.no in (" & strTableOut & ")"
        Set rs诊断 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CStr(varArr(0)), CStr(varArr(1)), CStr(varArr(2)), CStr(varArr(3)), CStr(varArr(4)), CStr(varArr(5)), _
            CStr(varArr(6)), CStr(varArr(7)), CStr(varArr(8)), CStr(varArr(9)))
    End If
    '--------------------------到此所有数据提取完成--------------------------------------------------------------
    With vsgInfo
        .Rows = vsgInfo.FixedRows
        For i = 1 To rs抽样.RecordCount
            .AddItem ""
            lngBaseRow = .Rows - 1
            
            .TextMatrix(lngBaseRow, COL_CF挂号单) = rs抽样!标识号
            
            .TextMatrix(lngBaseRow, COL_CF序号) = i
            .TextMatrix(lngBaseRow, COL_CF门诊号) = IIf("" = rs抽样!门诊号 & "", rs抽样!标识号 & "", rs抽样!门诊号 & "")
            .TextMatrix(lngBaseRow, COL_CF就诊日期) = Format(rs抽样!就诊日期 & "", "YYYY-MM-DD")
            .TextMatrix(lngBaseRow, COL_CF病人姓名) = rs抽样!病人姓名 & ""
            .TextMatrix(lngBaseRow, COL_CF处方医生) = rs抽样!处方医生 & ""
            rs科室.Filter = "科室id=" & Val(rs抽样!科室ID & "")
            If Not rs科室.EOF Then .TextMatrix(lngBaseRow, COL_CF科室) = rs科室!科室 & ""
            .TextMatrix(lngBaseRow, COL_CF病人年龄) = rs抽样!年龄 & ""
            .TextMatrix(lngBaseRow, COL_CF药品金额) = Format(rs抽样!药费 & "", strDec)
            
            rs处方金额.Filter = "标识号='" & rs抽样!标识号 & "'"
            If Not rs处方金额.EOF Then dblTmp = Val(rs处方金额!处方金额 & "")
            .TextMatrix(lngBaseRow, COL_CF处方金额) = Format(dblTmp, strDec): dblTmp = 0
            
            rs药品数量.Filter = "标识号='" & rs抽样!标识号 & "'"
            If Not rs药品数量.EOF Then
                .TextMatrix(lngBaseRow, COL_CF药品品种数) = rs药品数量!药品种数 & ""
                .TextMatrix(lngBaseRow, COL_CF基药品种数) = rs药品数量!基本药种数 & ""
                .TextMatrix(lngBaseRow, COL_CF抗药品种数) = rs药品数量!抗生药种数 & ""
                .TextMatrix(lngBaseRow, COL_CF抗药金额) = Format(rs药品数量!抗菌药费 & "", strDec)
            End If
            
            If Val(rs抽样!医嘱 & "") = 0 Then
                rs费用抗药明细.Filter = 0
                rs费用抗药明细.Filter = "标识号='" & rs抽样!标识号 & "'"
                If Not rs费用抗药明细.EOF Then
                    For j = 1 To rs费用抗药明细.RecordCount
                        If j = 1 Then
                            .TextMatrix(lngBaseRow, COL_CF通用名) = rs费用抗药明细!药品通用名 & ""
                            .TextMatrix(lngBaseRow, COL_CF规格) = rs费用抗药明细!规格 & ""
                            .TextMatrix(lngBaseRow, COL_CF数量) = rs费用抗药明细!数量 & ""
                            .TextMatrix(lngBaseRow, COL_CF金额) = Format(rs费用抗药明细!费用 & "", strDec)
                        Else
                            .AddItem ""
                            lngTmp = .Rows - 1
                            For k = COL_CF序号 To COL_CF挂号单
                                .TextMatrix(lngTmp, k) = .TextMatrix(lngBaseRow, k)
                            Next
                            .TextMatrix(lngTmp, COL_CF通用名) = rs费用抗药明细!药品通用名 & ""
                            .TextMatrix(lngTmp, COL_CF规格) = rs费用抗药明细!规格 & ""
                            .TextMatrix(lngTmp, COL_CF数量) = rs费用抗药明细!数量 & ""
                            .TextMatrix(lngTmp, COL_CF金额) = Format(rs费用抗药明细!费用 & "", strDec)
                        End If
                        rs费用抗药明细.MoveNext
                    Next
                End If
            Else
                bln注射 = False
                rs医嘱抗药明细.Filter = 0
                rs医嘱抗药明细.Filter = "标识号='" & rs抽样!标识号 & "'"
                If Not rs医嘱抗药明细.EOF Then
                    For j = 1 To rs医嘱抗药明细.RecordCount
                        
                        '获取用法用量，存入strTmp 中
                        strTmp = rs医嘱抗药明细!单次用量 & ""
                        If Mid(strTmp, 1, 1) = "." Then strTmp = "0" & strTmp
                        strTmp = rs医嘱抗药明细!执行频次 & "," & strTmp & rs医嘱抗药明细!计算单位
                        
                        If j = 1 Then
                            .TextMatrix(lngBaseRow, COL_CF通用名) = rs医嘱抗药明细!药品通用名 & ""
                            .TextMatrix(lngBaseRow, COL_CF规格) = rs医嘱抗药明细!规格 & ""
                            .TextMatrix(lngBaseRow, COL_CF数量) = rs医嘱抗药明细!数量 & ""
                            .TextMatrix(lngBaseRow, COL_CF金额) = Format(rs医嘱抗药明细!费用 & "", strDec)
                            .TextMatrix(lngBaseRow, COL_CF用法用量) = strTmp
                            .TextMatrix(lngBaseRow, COL_CF用药途径) = rs医嘱抗药明细!用药途径 & ""
                        Else
                            .AddItem ""
                            lngTmp = .Rows - 1
                            For k = COL_CF序号 To COL_CF挂号单
                                .TextMatrix(lngTmp, k) = .TextMatrix(lngBaseRow, k)
                            Next
                            .TextMatrix(lngTmp, COL_CF通用名) = rs医嘱抗药明细!药品通用名 & ""
                            .TextMatrix(lngTmp, COL_CF规格) = rs医嘱抗药明细!规格 & ""
                            .TextMatrix(lngTmp, COL_CF数量) = rs医嘱抗药明细!数量 & ""
                            .TextMatrix(lngTmp, COL_CF金额) = Format(rs医嘱抗药明细!费用 & "", strDec)
                            .TextMatrix(lngTmp, COL_CF用法用量) = strTmp
                            .TextMatrix(lngTmp, COL_CF用药途径) = rs医嘱抗药明细!用药途径 & ""
                        End If
                        If InStr(rs医嘱抗药明细!用药途径 & "", "注射") > 0 Then bln注射 = True
                        rs医嘱抗药明细.MoveNext
                    Next
                End If
                .TextMatrix(lngBaseRow, COL_CF注射剂) = IIf(bln注射, "有", "无")
                strTmp = "": rs诊断.Filter = 0
                rs诊断.Filter = "标识号='" & rs抽样!标识号 & "'"
                If Not rs诊断.EOF Then
                    For j = 1 To rs诊断.RecordCount
                        If InStr("," & strTmp & ",", "," & rs诊断!诊断描述 & ",") = 0 Then
                            strTmp = strTmp & "," & rs诊断!诊断描述
                        End If
                        rs诊断.MoveNext
                    Next
                End If
                .TextMatrix(lngBaseRow, COL_CF诊断) = Mid(strTmp, 2)
            End If
            rs抽样.MoveNext
        Next
    End With
    
    If blnRP Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "门诊处方抗菌用药调查表", dtpRPS(e_R2_dtpRPS_开始时间_2).Value & "," & dtpRPE(e_R2_dtpRPE_结束时间_2).Value
    Else
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "门急诊处方抗菌用药统计", dtpCountS(e_C2_dtpCountS_开始时间_2).Value & "," & dtpCountE(e_C2_dtpCountE_结束时间_2).Value
    End If
    
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optType_Click(Index As Integer)
    Select Case Index
    Case e_C3_optType_切口类型_非手术_15, e_C3_optType_切口类型_手术_18
        chkType(e_C3_chkType_切口类型_Ⅰ类_2).Enabled = Not optType(e_C3_optType_切口类型_非手术_15).Value
        chkType(e_C3_chkType_切口类型_Ⅱ类_3).Enabled = Not optType(e_C3_optType_切口类型_非手术_15).Value
        chkType(e_C3_chkType_切口类型_Ⅲ类_4).Enabled = Not optType(e_C3_optType_切口类型_非手术_15).Value
        chkType(e_C3_chkType_切口类型_Ⅳ类_8).Enabled = Not optType(e_C3_optType_切口类型_非手术_15).Value
    Case e_C0_optType_汇总方式_科室_9, e_C0_optType_汇总方式_医生_8, e_C0_optType_汇总方式_药品_7
        If Index = e_C0_optType_汇总方式_药品_7 Then '7-按药品汇总
            optType(e_C0_optType_排序方式_数量_12).Enabled = True
            optType(e_C0_optType_排序方式_金额_11).Enabled = True
        Else            '按9-科室和8-医生汇总
            optType(e_C0_optType_排序方式_数量_12).Enabled = False
            optType(e_C0_optType_排序方式_金额_11).Value = True
            optType(e_C0_optType_排序方式_金额_11).Enabled = False
        End If
    End Select
End Sub

Private Sub picDept_Resize()
    On Error Resume Next
    lvwItems.ColumnHeaders(1).Width = (picDept.Width - 1300)
End Sub

Private Sub picOther_Resize()
    On Error Resume Next
    
    tbcOther.Move 0, 0, picOther.Width, picOther.Height
End Sub

Private Sub picReport_Resize()
    On Error Resume Next
    
    tbcReport.Move 0, 0, picReport.Width, picReport.Height
End Sub

Private Sub picReportSub_Resize(Index As Integer)
    Dim lngT As Long '距离顶部的距离
    Dim lngL As Long '距离左端的距离
    Dim lngW As Long
    Dim lngH As Long
    Dim lngTmp As Long
    Dim lngW说明文字 As Long '底部说明文字的宽度，照一行字180来计算，如果调整了说明文字的行数要同步调整
    Dim int行数 As Integer

    lngT = 50: lngL = 60
    lngW = picReportSub(Index).Width - lngL
    lngH = picReportSub(Index).Height
    lngW说明文字 = 180

    On Error Resume Next
     
    Select Case Index
        Case PanelItem_抗菌药品消耗金额调查表
            
            int行数 = 1
            
            picFilter(e_R0_picFilter_条件容器_6).Left = lngL
            picFilter(e_R0_picFilter_条件容器_6).Top = lngT
            picFilter(e_R0_picFilter_条件容器_6).Height = cmdOK(e_R0_cmdOK_统计_0).Top + cmdOK(e_R0_cmdOK_统计_0).Height
            picFilter(e_R0_picFilter_条件容器_6).Width = lngW
            
            lblN(e_R0_lblN_标题_74).Left = (lngW - lblN(e_R0_lblN_标题_74).Width) / 2
            lblN(e_R0_lblN_标题_74).Top = picFilter(e_R0_picFilter_条件容器_6).Height + picFilter(e_R0_picFilter_条件容器_6).Top + 100
            
            vsBill.Left = lngL
            vsBill.Top = lblN(e_R0_lblN_标题_74).Top + lblN(e_R0_lblN_标题_74).Height + 50
            vsBill.Width = lngW
            vsBill.Height = lngH - vsBill.Top - lngW说明文字 * int行数 - 120
            
            lblInfo.Left = lngL
            lblInfo.Top = lngH - lngW说明文字 * int行数 - 60
            
        Case PanelItem_病人抗菌用药情况抽样调查及评价表
            
            int行数 = 1
            
            picFilter(e_R1_picFilter_条件容器_7).Left = lngL
            picFilter(e_R1_picFilter_条件容器_7).Top = lngT
            picFilter(e_R1_picFilter_条件容器_7).Height = txtCYJL.Top + txtCYJL.Height
            picFilter(e_R1_picFilter_条件容器_7).Width = lngW
            
            lngTmp = txtCYJL.Top + txtCYJL.Height + 30 + lngT
            
            lblN(e_R1_lblN_标题_77).Left = (lngW - lblN(e_R1_lblN_标题_77).Width) / 2
            lblN(e_R1_lblN_标题_77).Top = picFilter(e_R1_picFilter_条件容器_7).Height + picFilter(e_R1_picFilter_条件容器_7).Top + 100
            
            rptPati.Left = lngL
            rptPati.Top = lblN(e_R1_lblN_标题_77).Height + lblN(e_R1_lblN_标题_77).Top + 50
            rptPati.Width = lngW
            rptPati.Height = lngH - rptPati.Top - lngW说明文字 * int行数 - 120
            
            lblN(e_R1_lblN_底端说明_72).Left = lngL
            lblN(e_R1_lblN_底端说明_72).Top = lngH - lblN(e_R1_lblN_底端说明_72).Height - 60
            
        Case PanelItem_门诊处方抗菌用药调查表
        
            int行数 = 3
            
            picFilter(e_R2_picFilter_条件容器_8).Left = lngL
            picFilter(e_R2_picFilter_条件容器_8).Top = lngT
            picFilter(e_R2_picFilter_条件容器_8).Height = txtDept(e_R2_txtDept_抽样科室_2).Top + txtDept(e_R2_txtDept_抽样科室_2).Height
            picFilter(e_R2_picFilter_条件容器_8).Width = lngW
            
            lblN(e_R2_lblN_调查表_标题_25).Top = picFilter(e_R2_picFilter_条件容器_8).Height + picFilter(e_R2_picFilter_条件容器_8).Top + 100
            lblN(e_R2_lblN_调查表_标题_25).Left = (lngW - lblN(e_R2_lblN_调查表_标题_25).Width) / 2
            
            lblN(e_R2_lblN_调查表_日期_26).Top = lblN(e_R2_lblN_调查表_标题_25).Top + lblN(e_R2_lblN_调查表_标题_25).Height + 20
            lblN(e_R2_lblN_调查表_日期_26).Left = (lngW - lblN(e_R2_lblN_调查表_日期_26).Width) / 2
            
            lblN(e_R2_lblN_调查表_处方总量_27).Top = lblN(e_R2_lblN_调查表_日期_26).Top
            lblN(e_R2_lblN_调查表_处方总量_27).Left = lngW - lblN(e_R2_lblN_调查表_处方总量_27).Width - 100
            
            vsMZYY.Left = lngL: vsMZYY.Width = lngW
            vsMZYY.Top = lblN(e_R2_lblN_调查表_处方总量_27).Top + lblN(e_R2_lblN_调查表_处方总量_27).Height + 50
            
            vsCF.Left = lngL: vsCF.Width = lngW
            vsCF.Height = 1800
            vsCF.Top = lngH - vsCF.Height - lngW说明文字 * int行数 - 120
            
            lblN(e_R2_lblN_分析表_日期_29).Top = vsCF.Top - 50 - lblN(e_R2_lblN_分析表_日期_29).Height
            lblN(e_R2_lblN_分析表_日期_29).Left = (lngW - lblN(e_R2_lblN_分析表_日期_29).Width) / 2
            
            lblN(e_R2_lblN_分析表_处方总量_30).Left = lngW - lblN(e_R2_lblN_分析表_处方总量_30).Width - 100
            lblN(e_R2_lblN_分析表_处方总量_30).Top = lblN(e_R2_lblN_分析表_日期_29).Top
            
            lblN(e_R2_lblN_分析表_标题_70).Top = lblN(e_R2_lblN_分析表_日期_29).Top - 20 - lblN(e_R2_lblN_分析表_标题_70).Height
            lblN(e_R2_lblN_分析表_标题_70).Left = (lngW - lblN(e_R2_lblN_分析表_标题_70).Width) / 2
            
            lblCFSM.Left = lngL
            lblCFSM.Top = lngH - lngW说明文字 * int行数 - 60
            
            vsMZYY.Height = lblN(e_R2_lblN_分析表_标题_70).Top - vsMZYY.Top - 50
            
        Case PanelItem_住院病人抗菌用药调查表
            
            int行数 = 1
            
            picFilter(e_R3_picFilter_条件容器_9).Left = lngL
            picFilter(e_R3_picFilter_条件容器_9).Top = lngT
            picFilter(e_R3_picFilter_条件容器_9).Height = txtDept(e_R3_txtDept_抽样科室_3).Top + txtDept(e_R3_txtDept_抽样科室_3).Height
            picFilter(e_R3_picFilter_条件容器_9).Width = lngW
        
            lblN(e_R3_lblN_调查表_标题_43).Left = (lngW - lblN(e_R3_lblN_调查表_标题_43).Width) / 2
            lblN(e_R3_lblN_调查表_标题_43).Top = picFilter(e_R3_picFilter_条件容器_9).Height + picFilter(e_R3_picFilter_条件容器_9).Top + 100
            
            lblN(e_R3_lblN_调查表_患者天数_45).Left = lngW - lblN(e_R3_lblN_调查表_患者天数_45).Width - 300
            lblN(e_R3_lblN_调查表_患者天数_45).Top = lblN(e_R3_lblN_调查表_标题_43).Top + lblN(e_R3_lblN_调查表_标题_43).Height + 20
                
            vsZYYY.Left = lngL
            vsZYYY.Width = lngW
            vsZYYY.Top = lblN(e_R3_lblN_调查表_患者天数_45).Top + lblN(e_R3_lblN_调查表_患者天数_45).Height + 60
            vsZYYY.Height = lngH - vsZYYY.Top - lngW说明文字 * int行数 - 120
            
            lblN(e_R3_lblN_底端说明_44).Left = lngL
            lblN(e_R3_lblN_底端说明_44).Top = lngH - lngW说明文字 * int行数 - 60
            
    End Select
End Sub

Private Sub picOtherSub_Resize(Index As Integer)
    Dim lngL As Long
    Dim lngT As Long
    Dim lngW As Long
    Dim lngH As Long
    Dim lngW说明文字 As Long '底部说明文字的宽度，照一行字180来计算，如果调整了说明文字的行数要同步调整
    Dim int行数 As Integer
    
    On Error Resume Next
    
    lngL = 60: lngT = 50
    lngW = picOtherSub(Index).Width
    lngH = picOtherSub(Index).Height
    lngW说明文字 = 180
    
    picFilter(Index).Left = lngL
    picFilter(Index).Top = lngT
    picFilter(Index).Width = lngW - lngL
 
    Select Case Index
    
    Case PanelItem_抗菌药物使用情况排名统计
        
        int行数 = 1
        
        picFilter(e_C0_picFilter_条件容器_0).Height = cmdOK(e_C0_cmdOK_统计_4).Top + cmdOK(e_C0_cmdOK_统计_4).Height
        
        lblN(e_C0_lblN_标题_75).Left = (lngW - lngL - lblN(e_C0_lblN_标题_75).Width) / 2
        lblN(e_C0_lblN_标题_75).Top = picFilter(e_C0_picFilter_条件容器_0).Top + picFilter(e_C0_picFilter_条件容器_0).Height + 100
        
        vsUseRan.Left = lngL
        vsUseRan.Width = lngW - lngL
        vsUseRan.Top = lblN(e_C0_lblN_标题_75).Top + lblN(e_C0_lblN_标题_75).Height + 50
        vsUseRan.Height = lngH - vsUseRan.Top - lngW说明文字 * int行数 - 120
        
        lblUse.Left = lngL
        
        lblUse.Top = lngH - lngW说明文字 * int行数 - 60
        
    Case PanelItem_Ⅰ类切口围术期预防用药统计
        
        int行数 = 1
        
        picFilter(e_C1_picFilter_条件容器_1).Height = cmdOK(e_C1_cmdOK_统计_5).Top + cmdOK(e_C1_cmdOK_统计_5).Height
        
        lblN(e_C1_lblN_标题_76).Left = (lngW - lngL - lblN(e_C1_lblN_标题_76).Width) / 2
        lblN(e_C1_lblN_标题_76).Top = picFilter(e_C1_picFilter_条件容器_1).Top + picFilter(e_C1_picFilter_条件容器_1).Height + 100
        vsCut.Left = lngL
        vsCut.Top = lblN(e_C1_lblN_标题_76).Height + lblN(e_C1_lblN_标题_76).Top + 50
        vsCut.Width = lngW - lngL
        vsCut.Height = lngH - vsCut.Top - lngW说明文字 * int行数 - 120
        
        lblCut.Left = lngL
        lblCut.Top = lngH - lngW说明文字 * int行数 - 60
    
    Case PanelItem_门急诊处方抗菌用药统计
    
        int行数 = 4
        
        picFilter(e_C2_picFilter_条件容器_2).Height = cmdOK(e_C2_cmdOK_统计_6).Top + cmdOK(e_C2_cmdOK_统计_6).Height
        
        lblN(e_C2_lblN_统计表_标题_4).Left = (lngW - lngL - lblN(e_C2_lblN_统计表_标题_4).Width) / 2
        lblN(e_C2_lblN_统计表_标题_4).Top = picFilter(e_C2_picFilter_条件容器_2).Top + picFilter(e_C2_picFilter_条件容器_2).Height + 100
        
        lblN(e_C2_lblN_统计表_日期_46).Left = (lngW - lngL - lblN(e_C2_lblN_统计表_日期_46).Width) / 2
        lblN(e_C2_lblN_统计表_日期_46).Top = lblN(e_C2_lblN_统计表_标题_4).Top + lblN(e_C2_lblN_统计表_标题_4).Height + 20
        
        lblN(e_C2_lblN_统计表_处方总量_48).Left = lngW - lblN(e_C2_lblN_统计表_处方总量_48).Width - 100
        lblN(e_C2_lblN_统计表_处方总量_48).Top = lblN(e_C2_lblN_统计表_日期_46).Top
        
        vsCountDruUse.Left = lngL
        vsCountDruUse.Width = lngW - lngL
        vsCountDruUse.Top = lblN(e_C2_lblN_统计表_处方总量_48).Top + lblN(e_C2_lblN_统计表_处方总量_48).Height + 50
        
        vsCountCF.Left = lngL
        vsCountCF.Width = lngW - lngL
        vsCountCF.Height = 1800
        vsCountCF.Top = lngH - vsCountCF.Height - lngW说明文字 * int行数 - 120
        
        lblN(e_C2_lblN_底端说明_50).Left = lngL
        lblN(e_C2_lblN_底端说明_50).Top = lngH - lngW说明文字 * int行数 - 60

        lblN(e_C2_lblN_分析表_日期_47).Left = (lngW - lngL - lblN(e_C2_lblN_分析表_日期_47).Width) / 2
        lblN(e_C2_lblN_分析表_日期_47).Top = vsCountCF.Top - lblN(e_C2_lblN_分析表_日期_47).Height - 50
                
        lblN(e_C2_lblN_分析表_处方总量_49).Left = lngW - lblN(e_C2_lblN_分析表_处方总量_49).Width - 100
        lblN(e_C2_lblN_分析表_处方总量_49).Top = lblN(e_C2_lblN_分析表_日期_47).Top
        
        
        lblN(e_C2_lblN_分析表_标题_39).Left = (lngW - lngL - lblN(e_C2_lblN_分析表_标题_39).Width) / 2
        lblN(e_C2_lblN_分析表_标题_39).Top = lblN(e_C2_lblN_分析表_处方总量_49).Top - lblN(e_C2_lblN_分析表_标题_39).Height - 20
        
        vsCountDruUse.Height = lblN(e_C2_lblN_分析表_标题_39).Top - vsCountDruUse.Top - 50
        
    Case PanelItem_住院医嘱抗菌用药统计
        
        int行数 = 2
        
        picFilter(e_C3_picFilter_条件容器_3).Height = cmdOK(e_C3_cmdOK_统计_7).Top + cmdOK(e_C3_cmdOK_统计_7).Height
        
        lblN(e_C3_lblN_统计表_标题_5).Left = (lngW - lngL - lblN(e_C3_lblN_统计表_标题_5).Width) / 2
        lblN(e_C3_lblN_统计表_标题_5).Top = picFilter(e_C3_picFilter_条件容器_3).Top + picFilter(e_C3_picFilter_条件容器_3).Height + 100
        
        vsInDruUse.Left = lngL
        vsInDruUse.Top = lblN(e_C3_lblN_统计表_标题_5).Top + lblN(e_C3_lblN_统计表_标题_5).Height + 50
        vsInDruUse.Width = lngW - lngL
        
        vsInDruAna.Left = lngL
        vsInDruAna.Width = lngW - lngL
        vsInDruAna.Height = 2110
        vsInDruAna.Top = lngH - vsInDruAna.Height - lngW说明文字 * int行数 - 120
        
        lblN(e_C3_lblN_分析表_标题_59).Left = (lngW - lngL - lblN(e_C3_lblN_分析表_标题_59).Width) / 2
        lblN(e_C3_lblN_分析表_标题_59).Top = vsInDruAna.Top - lblN(e_C3_lblN_分析表_标题_59).Height - 50
        
        lblN(e_C3_lblN_底端说明_58).Left = lngL
        lblN(e_C3_lblN_底端说明_58).Top = lngH - lngW说明文字 * int行数 - 60
        
        
        vsInDruUse.Height = lblN(e_C3_lblN_分析表_标题_59).Top - vsInDruUse.Top - 50
        
    Case PanelItem_术后抗菌药物使用超N天统计
        
        int行数 = 1
        
        picFilter(e_C4_picFilter_条件容器_4).Height = cmdOK(e_C4_cmdOK_统计_8).Top + cmdOK(e_C4_cmdOK_统计_8).Height
        
        lblN(e_C4_lblN_统计表_标题_6).Left = (lngW - lngL - lblN(e_C4_lblN_统计表_标题_6).Width) / 2
        lblN(e_C4_lblN_统计表_标题_6).Top = picFilter(e_C4_picFilter_条件容器_4).Top + picFilter(e_C4_picFilter_条件容器_4).Height + 100
        
        vsOpeKssUse.Left = lngL
        vsOpeKssUse.Width = lngW - lngL
        vsOpeKssUse.Top = lblN(e_C4_lblN_统计表_标题_6).Top + lblN(e_C4_lblN_统计表_标题_6).Height + 50
        vsOpeKssUse.Height = lngH - vsOpeKssUse.Top - lngW说明文字 * int行数 - 120
        
        lblN(e_C4_lblN_底端说明_73).Left = lngL
        lblN(e_C4_lblN_底端说明_73).Top = lngH - lngW说明文字 * int行数 - 60
        
    Case PanelItem_医生治疗某疾病抗菌用药成本统计
        
        int行数 = 3
        
        picFilter(e_C5_picFilter_条件容器_5).Height = cmdOK(e_C5_cmdOK_统计_9).Top + cmdOK(e_C5_cmdOK_统计_9).Height + 70
        
        lblN(e_C5_lblN_分析表_标题_7).Left = (lngW - lngL - lblN(7).Width) / 2
        lblN(e_C5_lblN_分析表_标题_7).Top = picFilter(e_C5_picFilter_条件容器_5).Top + picFilter(e_C5_picFilter_条件容器_5).Height + 30
        
        vsIllDruUse.Left = lngL
        vsIllDruUse.Width = lngW - lngL
        vsIllDruUse.Top = lblN(e_C5_lblN_分析表_标题_7).Top + lblN(e_C5_lblN_分析表_标题_7).Height + 50
        vsIllDruUse.Height = lngH - vsIllDruUse.Top - lngW说明文字 * int行数 - 120
        
        lblN(e_C5_lblN_底端说明_67).Left = lngL
        lblN(e_C5_lblN_底端说明_67).Top = lngH - lngW说明文字 * int行数 - 60
    End Select
End Sub

Private Sub txtCount_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtCount(Index).Text) = 0 Then txtCount(Index).Text = 0
End Sub

Private Sub txtILL_GotFocus()
    Call zlControl.TxtSelAll(txtILL)
End Sub

Private Sub txtILL_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean

    If KeyAscii = 13 Then
        Call txtILL_Validate(blnCancel)
        If Not blnCancel Then Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vsInDruUse_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'功能：擦除边线
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    If Row < 2 Then Exit Sub
    
    With vsInDruUse
        If Col >= COL_DRU序号 And Col <= COL_DRU联合用药 Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If Val(.TextMatrix(Row, COL_DRU病人id)) = Val(.TextMatrix(i, COL_DRU病人id)) And Val(.TextMatrix(Row, COL_DRU主页id)) = Val(.TextMatrix(i, COL_DRU主页id)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If Val(.TextMatrix(Row, COL_DRU病人id)) = Val(.TextMatrix(i, COL_DRU病人id)) And Val(.TextMatrix(Row, COL_DRU主页id)) = Val(.TextMatrix(i, COL_DRU主页id)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
            
            If lngBegin = lngEnd Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            Done = True
            
        End If
    End With
End Sub

Private Sub vsMZYY_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsMZYY.RowHeightMax = 2000
    vsMZYY.RowHeightMin = 250
    vsMZYY.AutoSize 0, vsMZYY.Cols - 1
End Sub

Private Sub vsCountCF_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsCountCF.RowHeightMax = 2000
    vsCountCF.RowHeightMin = 250
    vsCountCF.AutoSize 0, vsCountCF.Cols - 1
End Sub

Private Sub vsInDruUse_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsInDruUse.RowHeightMax = 2000
    vsInDruUse.RowHeightMin = 250
    vsInDruUse.AutoSize 0, vsInDruUse.Cols - 1
End Sub

Private Sub vsCountDruUse_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'功能：擦除部分边框线
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    With vsCountDruUse
        If Row < 3 Then Exit Sub
        If Col >= COL_CF序号 And Col <= COL_CF抗药品种数 Or Col >= COL_CF处方金额 And Col <= COL_CF抗药金额 Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_CF挂号单) <> .TextMatrix(Row, COL_CF挂号单) Then
                    Exit For
                Else
                    lngBegin = i
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If .TextMatrix(i, COL_CF挂号单) <> .TextMatrix(Row, COL_CF挂号单) Then
                    Exit For
                Else
                    lngEnd = i
                End If
            Next
            
            If lngBegin = lngEnd Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            Done = True
        End If
    End With
End Sub

Private Sub vsMZYY_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'功能：擦除部分边框线
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim i As Long
    Dim vRect As RECT
    
    With vsMZYY
        If Row < 3 Then Exit Sub
        If Col >= COL_CF序号 And Col <= COL_CF抗药品种数 Or Col >= COL_CF处方金额 And Col <= COL_CF抗药金额 Then
            lngBegin = Row: lngEnd = Row
            
            For i = Row - 1 To .FixedRows Step -1
                If .TextMatrix(i, COL_CF挂号单) <> .TextMatrix(Row, COL_CF挂号单) Then
                    Exit For
                Else
                    lngBegin = i
                End If
            Next
            
            For i = Row + 1 To .Rows - 1
                If .TextMatrix(i, COL_CF挂号单) <> .TextMatrix(Row, COL_CF挂号单) Then
                    Exit For
                Else
                    lngEnd = i
                End If
            Next
            
            If lngBegin = lngEnd Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
            Done = True
        End If
    End With
End Sub

Private Sub LoadInKssAdvice()
'功能：加载 住院病人抗菌用药调查表 界面数据
    Dim strSql As String, strPar As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim strValue As String
    Dim dblTmp As Double
    Dim i As Long
    Dim j As Long
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Call zlCommFun.ShowFlash("正在读取数据，请稍候...")
    
    '日期范围参数生成
    strPar = "To_Date('" & Format(dtpRPS(e_R3_dtpRPS_开始时间_3).Value, "yyyy-MM-dd 00:00:00") & "','YYYY-MM-DD HH24:MI:SS')" & _
    " and To_Date('" & Format(dtpRPE(e_R3_dtpRPE_结束时间_3).Value, "yyyy-MM-dd 23:59:59") & "','YYYY-MM-DD HH24:MI:SS')"
    
    '加载表格数据
    strSql = "select /*+ rule*/ e.名称 as 类别,c.名称 as 药品通用名,d.药品剂型 as 剂型,c.规格||c.产地 as 规格,b.住院单位 as 单位,a.数量,a.费用" & vbNewLine & _
        "from (select x.收费细目id,sum(x.结帐金额) as 费用,sum(x.数次) as 数量 from 住院费用记录 X,病案主页 Y where x.收费类别='5'and x.记录状态 <> 0" & vbNewLine & _
        "and y.病人id=x.病人id and y.主页id=x.主页id" & _
        " and y.出院日期 between " & strPar & vbNewLine & _
        IIf(txtDept(e_R3_txtDept_抽样科室_3).Tag = "", "", " and y.出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))") & vbNewLine & _
        "group by x.收费细目id) a,药品规格 b,收费项目目录 c,药品特性 d,诊疗分类目录 e,诊疗项目目录 f" & vbNewLine & _
        "where a.收费细目id=b.药品id and b.药品id=c.id and d.药名id=b.药名ID and d.药名id=f.id and f.分类id=e.id and nvl(d.抗生素,0)<>0"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_R3_txtDept_抽样科室_3).Tag)

    If rsTmp.RecordCount > 0 Then
        With vsZYYY
            .Rows = vsZYYY.FixedRows
            For i = 1 To rsTmp.RecordCount
                '添加新行
                strTmp = ""
                For j = 0 To rsTmp.Fields.Count - 1
                    If j = rsTmp.Fields.Count - 1 Then
                        strValue = Replace(Nvl(rsTmp.Fields(j).Value), vbTab, "")
                        strTmp = strTmp & vbTab & IIf(Mid(strValue, 1, 1) = ".", "0" & strValue, strValue)
                    Else
                        strTmp = strTmp & vbTab & Replace(Nvl(rsTmp.Fields(j).Value), vbTab, "")
                    End If
                Next
                .AddItem Mid(strTmp, 2)
                rsTmp.MoveNext
            Next
            .SubtotalPosition = flexSTBelow
            .Subtotal flexSTSum, -1, COL_YZ总费用, "#######" & gstrDec, , vbBlack, False, "合计"
            
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            
            '格式化数量和金额，数量保留两位小数
            strTmp = "0.00"
            
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, COL_YZ数量) = Format(.TextMatrix(i, COL_YZ数量), gstrDec)
                .TextMatrix(i, COL_YZ总费用) = Format(.TextMatrix(i, COL_YZ总费用), gstrDec)
            Next
        End With
        
        '计算 患者人天数  医院收治患者人天数（医院季度出院患者总人数×同期平均住院天数）由医院统计部门提供。
        strSql = "Select sum(住院天数) as 总天数 From 病案主页 Where 出院日期 between " & strPar & _
            IIf(txtDept(e_R3_txtDept_抽样科室_3).Tag = "", "", " and 出院科室id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, txtDept(e_R3_txtDept_抽样科室_3).Tag)
        If Not rsTmp.EOF > 0 Then dblTmp = Val(rsTmp!总天数 & "")
        lblN(e_R3_lblN_调查表_患者天数_45).Caption = "收治患者人天数：" & Round(dblTmp) & "天"
    Else
        vsZYYY.Rows = vsZYYY.FixedRows
        vsZYYY.Rows = vsZYYY.Rows + 1
        Screen.MousePointer = 0
        Call zlCommFun.StopFlash
        MsgBox "当前条件下未找到任何数据，请重新设置抽样统计参数。", vbInformation, gstrSysName
        Exit Sub
    End If
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\日期范围", "住院病人抗菌用药调查表", dtpRPS(e_R3_dtpRPS_开始时间_3).Value & "," & dtpRPE(e_R3_dtpRPE_结束时间_3).Value
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    Exit Sub
errH:
    Screen.MousePointer = 0
    Call zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte, Optional ByVal intTabNum As Integer = 1)
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL,intTabNum页面中的第几个表格，默认只有一个当界面有两个表格时循环输出
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As New zlTabAppRow
    Dim objVSF As VSFlexGrid
    Dim objTable As Object
    Dim objReport As ReportControl
    Dim blnIsRPT As Boolean   'True-是ReportControl对象需要转换成VSF对象
    Dim varArr As Variant
    Dim i As Integer
    Dim strTmp As String
    Dim strButtom As String
    Dim strFace As String
    
    strFace = GetCur页面
    
    Select Case strFace
    
    Case "抗菌药品消耗金额调查表"
    
        objPrint.Title.Text = "抗菌药品消耗金额调查表"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpRPS(e_R0_dtpRPS_开始时间_0).Value & " 到 " & dtpRPE(e_R0_dtpRPE_结束时间_0).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsBill
        
        strButtom = lblInfo.Caption
    Case "病人抗菌用药情况抽样调查及评价表"
    
        objPrint.Title.Text = "抽样病人列表"
        
    
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpRPS(e_R1_dtpRPS_开始时间_1).Value & " 到 " & dtpRPE(e_R1_dtpRPE_结束时间_1).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "抽样科室：" & txtCYJL.Text
        objAppRow.Add "共：" & Val(lblN(e_R1_lblN_抽样数量标题_13).Tag) & "人。"
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = rptPati
        blnIsRPT = True
        strButtom = vbCrLf & lblN(e_R1_lblN_底端说明_72).Caption
        
    Case "门诊处方抗菌用药调查表"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpRPS(e_R2_dtpRPS_开始时间_2).Value & " 到 " & dtpRPE(e_R2_dtpRPE_结束时间_2).Value
        objAppRow.Add lblN(e_R2_lblN_调查表_处方总量_27).Caption
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_R2_lblN_调查表_标题_25).Caption
            Set objTable = vsMZYY
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_R2_lblN_分析表_标题_70).Caption
            Set objTable = vsCF
        End If
        
        strButtom = lblCFSM.Caption
    Case "住院病人抗菌用药调查表"
    
        objPrint.Title.Text = lblN(e_R3_lblN_调查表_标题_43).Caption
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpRPS(e_R3_dtpRPS_开始时间_3).Value & " 到 " & dtpRPE(e_R3_dtpRPE_结束时间_3).Value
        objAppRow.Add lblN(e_R3_lblN_调查表_患者天数_45).Caption
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsZYYY
        
        strButtom = lblN(e_R3_lblN_底端说明_44).Caption
        
    Case "抗菌药物使用情况排名统计"
        objPrint.Title.Text = "抗菌药物使用情况排名统计"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C0_dtpCountS_开始时间_0).Value & " 到 " & dtpCountE(e_C0_dtpCountE_结束时间_0).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsUseRan
        
        strButtom = lblUse.Caption
    Case "Ⅰ类切口围术期预防用药统计"
        objPrint.Title.Text = "Ⅰ类切口围术期预防用药统计"
        
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C1_dtpCountS_开始时间_1).Value & " 到 " & dtpCountE(e_C1_dtpCountE_结束时间_1).Value
        objPrint.UnderAppRows.Add objAppRow
        
        Set objTable = vsCut
        
        strButtom = lblCut.Caption
        
    Case "门急诊处方抗菌用药统计"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C2_dtpCountS_开始时间_2).Value & " 到 " & dtpCountE(e_C2_dtpCountE_结束时间_2).Value
        objAppRow.Add lblN(e_C2_lblN_统计表_处方总量_48).Caption
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_C2_lblN_统计表_标题_4).Caption
            Set objTable = vsCountDruUse
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_C2_lblN_分析表_标题_39).Caption
            Set objTable = vsCountCF
        End If
        
        strButtom = lblN(e_C2_lblN_底端说明_50).Caption
    
    Case "住院医嘱抗菌用药统计"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C3_dtpCountS_开始时间_3).Value & " 到 " & dtpCountE(e_C3_dtpCountE_结束时间_3).Value
        objPrint.UnderAppRows.Add objAppRow
            
        If intTabNum = 1 Then
            objPrint.Title.Text = lblN(e_C3_lblN_统计表_标题_5).Caption
            Set objTable = vsInDruUse
        ElseIf intTabNum = 2 Then
            objPrint.Title.Text = lblN(e_C3_lblN_分析表_标题_59).Caption
            Set objTable = vsInDruAna
        End If
        
        strButtom = lblN(e_C3_lblN_底端说明_58).Caption
        
    
    Case "术后抗菌药物使用超N天统计"
        objPrint.Title.Text = "术后抗菌药物使用超N天统计"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C4_dtpCountS_开始时间_4).Value & " 到 " & dtpCountE(e_C4_dtpCountE_结束时间_4).Value
        objPrint.UnderAppRows.Add objAppRow
        Set objTable = vsOpeKssUse
        strButtom = lblN(e_C4_lblN_底端说明_73).Caption
    Case "医生治疗某疾病抗菌用药成本统计"
        objPrint.Title.Text = "医生治疗某疾病抗菌用药成本统计"
        Set objAppRow = New zlTabAppRow
        objAppRow.Add "统计时间：" & dtpCountS(e_C5_dtpCountS_开始时间_5).Value & " 到 " & dtpCountE(e_C5_dtpCountE_结束时间_5).Value
        objAppRow.Add "诊断：" & txtILL.Text
        objPrint.UnderAppRows.Add objAppRow
        Set objTable = vsIllDruUse
        strButtom = lblN(e_C5_lblN_底端说明_67).Caption
    End Select
    
    '复制数据表格
    If blnIsRPT Then
        Set objReport = objTable
        If objReport.Records.Count = 0 Then Exit Sub
        If Not zlReportToVSFlexGrid(vsTmp, objReport) Then Exit Sub
        blnIsRPT = False
    Else
        Set objVSF = objTable
        If Not zlCopyVSFlexGrid(vsTmp, objVSF) Then Exit Sub
    End If
    
    '调用打印部件处理
    '---------------------------------------
    Set objPrint.Body = Me.vsTmp
    '表下
    If strButtom <> "" Then
        varArr = Split(strButtom, vbCrLf)
        For i = 0 To UBound(varArr)
            Set objAppRow = New zlTabAppRow
            strTmp = varArr(i)
            Call objAppRow.Add(strTmp)
            Call objPrint.BelowAppRows.Add(objAppRow)
        Next
        strButtom = ""
    End If
    
    '打印人时间等信息
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
    If (strFace = "门诊处方抗菌用药调查表" Or strFace = "门急诊处方抗菌用药统计" Or strFace = "住院医嘱抗菌用药统计") And intTabNum = 1 Then
        Call zlRptPrint(bytMode, 2)
    End If
    
End Sub

Private Sub Set界面说明(ByVal strCaption As String)
'功能：设置界面下方的说明信息
    Select Case strCaption
        Case "抗菌药品消耗金额调查表"
            lblInfo.Caption = "说明：政府拨款为零，系统中不记录此费用；总收入、药品总收入和差价来源于汇总记录；其它项来源于费用明细记录或计算得来。"
        Case "病人抗菌用药情况抽样调查及评价表"
            lblN(e_R1_lblN_底端说明_72).Caption = "说明：按出院时间和出院科室过滤出院病人，包含未使用抗药的出院病人。"
        Case "门诊处方抗菌用药调查表"
            lblCFSM.Caption = "说明：处方指一张包含药品费用的收费单据。通过医嘱产生的费用单据才有用药明细诊断等，直接收费产生的单据只有药品信息。" & vbCrLf & _
                "       按费用发生时间和开单室科抽样，如无门诊号则显示为费用的单据号，就诊日期指费用发生时间，处方医生指开单人，科室指开单科室。" & vbCrLf & _
                "       按药品品种计算各类药品种数，同一药品不同规格算一种药。"
        Case "住院病人抗菌用药调查表"
            lblN(e_R3_lblN_底端说明_44).Caption = "说明：按出院时间和出院科室过滤使用了抗菌药的出院病人；收治患者人天数=出院患者总数×同期平均住院天数，即总住院天数。"
        Case "抗菌药物使用情况排名统计"
            lblUse.Caption = "说明：统计时间指费用发生时间；抽样科室指开单科室；按住院场合统计时：病人一次住院算一次；统计门诊费用时以处方为单位。"
        Case "Ⅰ类切口围术期预防用药统计"
            lblCut.Caption = "说明：通过时间和科室过滤出院病人，抗菌药是医嘱方式下达才能统计，药品种数按品种统计，天数：长嘱由医嘱执行方案计算得到，临嘱首页中病人抗生素记录中取，如果没有则默认为一天。"
        Case "门急诊处方抗菌用药统计"
            lblN(e_C2_lblN_底端说明_50).Caption = "说明：处方指一张包含药品费用的收费单据。通过医嘱产生的费用才有用药明细诊断等，直接收费产生的单据只有药品信息。" & vbCrLf & _
                "       急诊区分：从费用单中区分是否急诊，如果是医嘱产生的单据则进一步判断病人挂号是否急诊。" & vbCrLf & _
                "       按费用发生时间和开单室科抽样，如果无门诊号则显示为费用的单据号，就诊日期指用费发生时间，处方医生指开单人，科室指开单科室。" & vbCrLf & _
                "       按药品品种计算各类药品种数，同一药品不同规格算一种药。"
        Case "住院医嘱抗菌用药统计"
            lblN(e_C3_lblN_底端说明_58).Caption = "说明：按病人出院时间和出院科室抽样病人，诊断是首页中的西(中)医首要出院诊断，药品种数是按品种统计，同种药不同规格算一种药。" & vbCrLf & _
                "       抗菌药使用明细指来自病人的医嘱记录，临嘱的用药天数需在医嘱下达时填上否则对应列为零。"
        Case "术后抗菌药物使用超N天统计"
            lblN(e_C4_lblN_底端说明_73).Caption = "说明：统计出院病人，要求首页中填写了手术情况才能统。"
        Case "医生治疗某疾病抗菌用药成本统计"
            lblN(e_C5_lblN_底端说明_67).Caption = "说明：统计对象为满足条件的全部出院病人，病人出院时间在指定范围内且首页中填写的出院主要诊断为界面上选择的指定诊断。" & vbCrLf & _
                "       主管医生指病人住院医师如果没有住院医师则显示为空，治愈率(%)=治愈人数/治疗人数，人均治疗金额=总金额/治疗人数，" & vbCrLf & _
                "       人均日金额=总金额/治疗病人住院天数之和，抗菌药物品种数：相同药品治疗不同病人时只记一种。"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
 
Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim lngIndex As Long
    If rptPati.SelectedRows.Count > 0 Then Call Edit病人调查
End Sub
 
Private Sub tbcOther_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call Set界面说明(Item.Tag)
End Sub
 
Private Sub tbcReport_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call Set界面说明(Item.Tag)
End Sub

Private Sub txtCount_KeyPress(Index As Integer, KeyAscii As Integer)
'功能：抽样人数不只能输入数字
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNum_KeyPress(Index As Integer, KeyAscii As Integer)
'功能：抽样人数不只能输入数字
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTopRan_KeyPress(KeyAscii As Integer)
'功能：抽样人数不只能输入数字
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub LoadDept()
'科室选择器------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim objItem As ListItem
    
    On Error GoTo errH
    
    strSql = "select distinct ID,编码,名称" & _
        " from 部门表 D,部门性质说明 T" & _
        " where D.ID=T.部门ID and 工作性质=[1] " & _
        " and (D.撤档时间 is null or D.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
        " order by 编码"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "临床")
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsTmp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTmp!ID, rsTmp!名称)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = rsTmp!编码
        objItem.Checked = False
        rsTmp.MoveNext
    Loop
    
    '没有时退出
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then '全选 Ctrl+A
        Call SetSelect(lvwItems, True)
    End If
    
    If KeyCode = vbKeyR And Shift = vbCtrlMask Then     '全消 Ctrl+R
        Call SetSelect(lvwItems, False)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And picDept.Visible Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
    If Me.ActiveControl.Name = "txtILL" Then Exit Sub
    If TypeName(Me.ActiveControl) <> "CommandButton" And KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    
    If Not picDept.Visible And KeyCode = vbKeyD And Shift = vbCtrlMask Then
        
        Select Case GetCur页面
        
        Case "病人抗菌用药情况抽样调查及评价表"
            Call cmdDept_Click(e_R1_cmdDept_科室选择器_1)
  
        Case "门诊处方抗菌用药调查表"
            Call cmdDept_Click(e_R2_cmdDept_科室选择器_2)
            
        Case "住院病人抗菌用药调查表"
        
            Call cmdDept_Click(e_R3_cmdDept_科室选择器_3)
            
        Case "抗菌药物使用情况排名统计"
            Call cmdDept_Click(e_C0_cmdDept_科室选择器_0)
            
        Case "Ⅰ类切口围术期预防用药统计"
            Call cmdDept_Click(e_C1_cmdDept_科室选择器_4)
            
        Case "门急诊处方抗菌用药统计"
            Call cmdDept_Click(e_C2_cmdDept_科室选择器_5)
        
        Case "住院医嘱抗菌用药统计"
            Call cmdDept_Click(e_C3_cmdDept_科室选择器_6)
            
        Case "术后抗菌药物使用超N天统计"
            Call cmdDept_Click(e_C4_cmdDept_科室选择器_7)
            
        End Select
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If GetCur页面 = "病人抗菌用药情况抽样调查及评价表" Then
            If picDept.Visible = False Then Call cmdCYSel_Click
        End If
    ElseIf KeyCode = vbKeyI And Shift = vbCtrlMask Then
        If GetCur页面 = "医生治疗某疾病抗菌用药成本统计" Then
            Call cmdILL_Click
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIf(mstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIf(mstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "没有找到您查找的科室。", vbInformation, Me.Caption
        Else
            MsgBox "已经是最后一个科室了。", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal blnSelect As Boolean = True)
    Dim i As Integer
    
    With lvwObj
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = blnSelect
        Next
    End With
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str科室 As String
    Dim str科室IDs As String
    Dim strTmp As String
    Dim varArr As Variant
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
    
    Dim intIndex As Integer
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    Select Case GetCur页面
        Case "病人抗菌用药情况抽样调查及评价表"
            intIndex = 1
        Case "门诊处方抗菌用药调查表"
            intIndex = 2
        Case "住院病人抗菌用药调查表"
            intIndex = 3
        Case "抗菌药物使用情况排名统计"
            intIndex = 0
        Case "Ⅰ类切口围术期预防用药统计"
            intIndex = 4
        Case "门急诊处方抗菌用药统计"
            intIndex = 5
        Case "住院医嘱抗菌用药统计"
            intIndex = 6
        Case "术后抗菌药物使用超N天统计"
            intIndex = 7
        Case "医生治疗某疾病抗菌用药成本统计"
            intIndex = 8
    End Select
   
    For i = 1 To lvwItems.ListItems.Count
        If lvwItems.ListItems(i).Checked Then
            strTmp = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            If InStr(str科室, strTmp) = 0 Then str科室 = str科室 & ";" & strTmp
        End If
    Next
    If str科室 = "" Then
        txtDept(intIndex).Text = "所有科室"
        txtDept(intIndex).ToolTipText = "所有科室"
        txtDept(intIndex).Tag = ""
        picDept.Visible = False
        txtFind.Text = ""
        Exit Sub
    End If
    str科室 = Mid(str科室, 2)
    
    varArr = Split(str科室, ";"): strTmp = ""
    
    For i = 0 To UBound(varArr)
        strTmp = strTmp & "," & Split(varArr(i), ",")(1)
        str科室IDs = str科室IDs & "," & Split(varArr(i), ",")(0)
    Next
    
    txtDept(intIndex).Text = Mid(strTmp, 2)
    txtDept(intIndex).ToolTipText = txtDept(intIndex).Text
    txtDept(intIndex).Tag = Mid(str科室IDs, 2)
    picDept.Visible = False
    txtFind.Text = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDFINDCANCLE,LVWITEMS,PICDEPT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    txtFind.Text = ""
    mlngFind = 1
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub cmdDept_Click(Index As Integer)
'功能：显示部门选择器
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim lngTmp  As Long
    Dim i As Integer
    
    With Me.picDept
        .Left = txtDept(Index).Left
        .Width = txtDept(Index).Width + 700
        .Top = txtDept(Index).Top + txtDept(Index).Height + picReportSub(PanelItem_抗菌药品消耗金额调查表).Top + picReport.Top + 950
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = 0
        .Top = txtFind.Height + 100
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height - txtFind.Height - 50 - 50
        txtFind.Top = 50
        cmdFind.Top = 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        .SetFocus
        .Refresh
    End With
    
    Call SetSelect(lvwItems, False)
    If txtDept(Index).Tag = "" Then Exit Sub
   
    For i = 1 To lvwItems.ListItems.Count
        lngTmp = Val(Mid(lvwItems.ListItems(i).Key, 2))
        Me.lvwItems.ListItems(i).Checked = InStr("," & txtDept(Index).Tag & ",", "," & lngTmp & ",") > 0
    Next
End Sub

Private Sub cmdCYSel_Click()
'功能：抽样记录选择器---------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    Dim blnCanle As Boolean
    Dim x As Long, y As Long
    
    x = Me.Left + tbcSub.Left + 1150
    y = Me.Top + tbcSub.Top + 1900
            
    strSql = "Select ID,抽样人,To_Char(抽样时间, 'YYYY-MM-DD HH24:MI:SS') as 抽样时间 From 抗菌药物抽样记录 order by 抽样时间 desc"
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "抗菌药物抽样记录", False, "", "", False, False, True, x, y, txtCYJL.Height, blnCanle, False, True)
    If blnCanle Then Exit Sub
    If rsTmp Is Nothing Then
        MsgBox "目前没有抽样记录，请先执行抽样。", vbInformation, gstrSysName
        Exit Sub
    End If
    txtCYJL.Text = "抽样时间：" & rsTmp!抽样时间 & "  抽样人：" & rsTmp!抽样人
    txtCYJL.Tag = rsTmp!ID
    Call LoadPati
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    With Me.tbcSub
    
        .Left = lngLeft
        .Top = lngTop
        .Height = lngBottom - lngTop
        .Width = lngRight - lngLeft
    
    End With
    Me.Refresh
End Sub

Private Function GetDiagSQL(ByVal strInput As String) As String
'功能：获得查询诊断的SQL
'参数：strInput-查询条件,strsql--返回的SQL
'返回：strsql--查询中医诊断的SQL
    Dim strSql As String
    
    If optType(26).Value Then  '中医诊断
        If optType(27).Value Then    ' 按诊断标准
            '按诊断输入:中医部份，一个诊断可能属于多个分类
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
           strSql = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,Null as 类别,A.名称,A.说明,A.编者," & vbNewLine & _
                " Decode(b.名称, [4], 1, Decode(b.简码,[4],1,decode(a.编码,[4],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                " Decode(Substr(b.名称, 1, Length([4])), [4], 1, Decode(Substr(b.简码, 1, Length([4])),[4],1,decode(Substr(a.编码, 1, Length([4])),[4],1,NULL))) As 排序3ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B, 疾病诊断科室 C, 疾病诊断科室 D" & _
                " Where A.ID=B.诊断ID And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And A.类别=2" & _
                " And B.码类=[3] And d.人员id(+) = [5] And (c.科室id In (Select 部门id From 部门人员 Where 人员id = [5]) Or c.科室id Is Null) " & _
                " And (" & strSql & ")" & _
                " Order by 排序1ID, 排序2ID, 排序3ID,A.编码"
                '排序顺序：先是完全匹配(名称、简码、编码）、个人收藏、其次是科室收藏、然后是左匹配(名称、简码、编码）、最后是双向匹配
        Else
            'B-中医疾病编码
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "A.名称 Like [2]" '输入汉字时只匹配名称
            Else
                strSql = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]"
            End If
            strSql = _
                "Select Distinct a.Id, a.Id As 项目id, a.编码, a.类别, a.附码, a.名称," & IIf(mint简码 = 0, "A.简码", "A.五笔码 as 简码") & ", a.说明," & _
                " Decode(a.名称, [4], 1, Decode(" & IIf(mint简码 = 0, "A.简码", "A.五笔码") & ",[4],1,decode(a.编码,[4],1,NULL))) As 排序1ID," & vbNewLine & _
                "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                "                Decode(Substr(a.名称, 1, Length([4])), [4], 1, Decode(Substr(" & IIf(mint简码 = 0, "A.简码", "A.五笔码") & ", 1, Length([4])),[4],1,decode(Substr(a.编码, 1, Length([4])),[4],1,NULL))) As 排序3ID" & vbNewLine & _
                "From 疾病编码目录 A, 疾病编码科室 C, 疾病编码科室 D" & vbNewLine & _
                "Where a.类别 = 'B' And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.疾病id(+) = a.Id And" & vbNewLine & _
                "      d.疾病id(+) = a.Id And (c.科室id In (Select 部门id From 部门人员 Where 人员id = [5]) Or c.科室id Is Null) And d.人员id(+) = [5] And (" & strSql & ")" & _
                "Order By 排序1ID, 排序2ID, 排序3ID, 编码"
        End If
    Else
        If optType(27).Value Then
            '按诊断输入:西医部份，一个诊断可能属于多个分类
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "B.名称 Like [2]" '输入汉字时,只匹配名称
            Else
                strSql = "A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]"
            End If
            strSql = _
                " Select Distinct A.ID,A.ID as 项目ID,A.编码,Null as 类别,A.名称,A.说明,A.编者," & vbNewLine & _
                " Decode(b.名称, [4], 1, Decode(b.简码,[4],1,decode(a.编码,[4],1,NULL))) As 排序1ID,Decode(d.诊断id, Null, Decode(c.诊断id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                " Decode(Substr(b.名称, 1, Length([4])), [4], 1, Decode(Substr(b.简码, 1, Length([4])),[4],1,decode(Substr(a.编码, 1, Length([4])),[4],1,NULL))) As 排序3ID" & _
                " From 疾病诊断目录 A,疾病诊断别名 B, 疾病诊断科室 C, 疾病诊断科室 D" & _
                " Where A.ID=B.诊断ID And c.诊断id(+) = a.Id And d.诊断id(+) = a.Id And A.类别=1" & _
                " And B.码类=[3] And d.人员id(+) = [5] And (c.科室id In (Select 部门id From 部门人员 Where 人员id = [5]) Or c.科室id Is Null) " & _
                " And (" & strSql & ")" & _
                " Order by 排序1ID, 排序2ID, 排序3ID,A.编码"
                '排序顺序：先是完全匹配(名称、简码、编码）、个人收藏、其次是科室收藏、然后是左匹配(名称、简码、编码）、最后是双向匹配
        Else
            'D-ICD-10疾病编码
            If zlCommFun.IsCharChinese(strInput) Then
                strSql = "A.名称 Like [2]" '输入汉字时,只匹配名称
            Else
                strSql = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]"
            End If
            strSql = _
                "Select Distinct a.Id, a.Id As 项目id, a.编码, a.类别, a.附码, a.名称," & IIf(mint简码 = 0, "A.简码", "A.五笔码 as 简码") & ", a.说明," & _
                " Decode(a.名称, [4], 1, Decode(" & IIf(mint简码 = 0, "A.简码", "A.五笔码") & ",[4],1,decode(a.编码,[4],1,NULL))) As 排序1ID," & vbNewLine & _
                "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
                "                Decode(Substr(a.名称, 1, Length([4])), [4], 1, Decode(Substr(" & IIf(mint简码 = 0, "A.简码", "A.五笔码") & ", 1, Length([4])),[4],1,decode(Substr(a.编码, 1, Length([4])),[4],1,NULL))) As 排序3ID" & vbNewLine & _
                "From 疾病编码目录 A, 疾病编码科室 C, 疾病编码科室 D" & vbNewLine & _
                "Where a.类别 = 'D' And (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And c.疾病id(+) = a.Id And" & vbNewLine & _
                "      d.疾病id(+) = a.Id And (c.科室id In (Select 部门id From 部门人员 Where 人员id = [5]) Or c.科室id Is Null) And d.人员id(+) = [5] And (" & strSql & ")" & _
                "Order By 排序1ID, 排序2ID, 排序3ID, 编码"

        End If
    End If
    GetDiagSQL = strSql
End Function

Private Sub txtILL_Validate(Cancel As Boolean)
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    If Trim(txtILL.Text) = "" Then
        txtILL.Tag = ""
        Exit Sub
    End If
    strTmp = Trim(txtILL.Text)
    If strTmp = cmdILL.Tag Then Exit Sub
    On Error GoTo errH
    strSql = GetDiagSQL(txtILL.Text)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, strTmp, mint简码 + 1, strTmp, UserInfo.ID)
    If rsTmp.RecordCount = 1 Then
        txtILL.Tag = rsTmp!项目ID
        txtILL.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
        cmdILL.Tag = txtILL.Text
    Else
        MsgBox "未找到对应项目。", vbInformation, gstrSysName
        Cancel = True
        Call zlControl.TxtSelAll(txtILL)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo errH:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '标题行复制
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = rptCol.Width * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                Next
            End If
        Next
        
        .RowHeight(-1) = .RowHeightMin
    End With
    zlReportToVSFlexGrid = True
    Exit Function
errH:
    zlReportToVSFlexGrid = False

End Function

Private Function zlCopyVSFlexGrid(vsgTemp As VSFlexGrid, vsgCopy As VSFlexGrid) As Boolean
'功能: 将vsgCopy的可见行列的数据复制到vsgTemp中 , 便于Excel输出
'参数:
'     vsgTemp-复制后的对象
'     vsgCopy-被复制的对象
'     strMsg -提示信息
    Dim i As Long
    Dim j As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim lngTmp As Long
    
    On Error GoTo errH:
    
    With vsgTemp
        .Rows = 0: .Cols = 0
        .Rows = vsgCopy.Rows
        .FixedRows = vsgCopy.FixedRows
        .MergeCells = vsgCopy.MergeCells
        
        '复制
        lngCol = 0
        For i = 0 To vsgCopy.Cols - 1 '列
            If Not vsgCopy.ColHidden(i) Then
                
                .Cols = .Cols + 1
                .ColWidth(lngCol) = vsgCopy.ColWidth(i)
                lngRow = 0: lngTmp = 0
                
                For j = 0 To vsgCopy.Rows - 1 '行
                    If Not vsgCopy.RowHidden(j) Then
                        .ColAlignment(lngCol) = vsgCopy.ColAlignment(i)
                        .Cell(flexcpAlignment, lngRow, lngCol) = vsgCopy.Cell(flexcpAlignment, j, i)  '对齐方式
                        .TextMatrix(lngRow, lngCol) = vsgCopy.TextMatrix(j, i)
                        lngRow = lngRow + 1
                    Else
                        lngTmp = lngTmp + 1  '记录隐藏行
                    End If
                Next
                lngCol = lngCol + 1
            End If
        Next
        '
        .Rows = .Rows - lngTmp '删除隐藏行
        .FixedCols = vsgCopy.FixedCols
        .RowHeight(-1) = vsgCopy.RowHeightMin
    End With
    
    zlCopyVSFlexGrid = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCur页面() As String
'功能：当前是那个页面
    Select Case tbcSub.Selected.Index
        Case 0
            Select Case tbcReport.Selected.Index
                Case 0
                    GetCur页面 = "抗菌药品消耗金额调查表"
                Case 1
                    GetCur页面 = "病人抗菌用药情况抽样调查及评价表"
                Case 2
                    GetCur页面 = "门诊处方抗菌用药调查表"
                Case 3
                    GetCur页面 = "住院病人抗菌用药调查表"
            End Select
        Case 1
            Select Case tbcOther.Selected.Index
                Case 0
                    GetCur页面 = "抗菌药物使用情况排名统计"
                Case 1
                    GetCur页面 = "Ⅰ类切口围术期预防用药统计"
                Case 2
                    GetCur页面 = "门急诊处方抗菌用药统计"
                Case 3
                    GetCur页面 = "住院医嘱抗菌用药统计"
                Case 4
                    GetCur页面 = "术后抗菌药物使用超N天统计"
                Case 5
                    GetCur页面 = "医生治疗某疾病抗菌用药成本统计"
            End Select
    End Select
End Function

