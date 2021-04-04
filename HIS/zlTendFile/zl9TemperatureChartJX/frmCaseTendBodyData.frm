VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmCaseTendBodyData 
   AutoRedraw      =   -1  'True
   Caption         =   "体温数据编辑"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13650
   Icon            =   "frmCaseTendBodyData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   13650
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   855
      ScaleHeight     =   825
      ScaleWidth      =   9975
      TabIndex        =   44
      Top             =   3540
      Width           =   10005
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前无表格项目,请点击添加项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1665
         TabIndex        =   45
         Top             =   270
         Width           =   6315
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   720
      ScaleHeight     =   6345
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   720
      Width           =   9495
      Begin VB.PictureBox picSplitTab 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   30
         ScaleWidth      =   6255
         TabIndex        =   22
         Top             =   3600
         Width           =   6255
      End
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   8100
         ScaleHeight     =   1425
         ScaleWidth      =   1185
         TabIndex        =   37
         Top             =   2175
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Index           =   0
            ItemData        =   "frmCaseTendBodyData.frx":6852
            Left            =   0
            List            =   "frmCaseTendBodyData.frx":685F
            TabIndex        =   39
            Top             =   855
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -30
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "录入："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   30
            TabIndex        =   41
            Top             =   30
            Width           =   540
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "选择："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   15
            TabIndex        =   40
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.PictureBox picEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8160
         ScaleHeight     =   255
         ScaleWidth      =   1305
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
         Begin VB.PictureBox picHour 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   -10
            ScaleHeight     =   255
            ScaleWidth      =   465
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   -10
            Visible         =   0   'False
            Width           =   495
            Begin VB.TextBox txtHour 
               Alignment       =   2  'Center
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               MaxLength       =   2
               TabIndex        =   34
               Top             =   15
               Width           =   315
            End
            Begin VB.Label lblHour 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "h"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   345
               TabIndex        =   35
               Top             =   45
               Width           =   105
            End
         End
         Begin VB.TextBox txtEdit 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   32
            Top             =   0
            Width           =   800
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "‥"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1080
            TabIndex        =   31
            Top             =   30
            Width           =   285
         End
         Begin VB.Label lblCheck 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         ItemData        =   "frmCaseTendBodyData.frx":6878
         Left            =   8100
         List            =   "frmCaseTendBodyData.frx":6885
         Style           =   1  'Checkbox
         TabIndex        =   29
         Top             =   900
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   6120
         ScaleHeight     =   1425
         ScaleWidth      =   2025
         TabIndex        =   25
         Top             =   4200
         Width           =   2055
         Begin zl9TemperatureChartJX.ColorPicker usrColor 
            Height          =   2190
            Left            =   0
            TabIndex        =   26
            Top             =   -720
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.Frame fraTabDetail 
         Height          =   2415
         Left            =   0
         TabIndex        =   24
         Top             =   3720
         Width           =   9495
         Begin VSFlex8Ctl.VSFlexGrid vsfTabDetail 
            Height          =   1815
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   8775
            _cx             =   15478
            _cy             =   3201
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   2
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
      End
      Begin VB.Frame fraTable 
         Height          =   3495
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   9495
         Begin VSFlex8Ctl.VSFlexGrid vsfTab 
            Height          =   2775
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   8415
            _cx             =   14843
            _cy             =   4895
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
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   6
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
      End
      Begin VB.Label lbllst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   75
         Index           =   1
         Left            =   8880
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label lbllst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   8880
         TabIndex        =   42
         Top             =   3240
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.PictureBox picCurve 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   120
      ScaleHeight     =   5865
      ScaleWidth      =   9705
      TabIndex        =   1
      Top             =   720
      Width           =   9735
      Begin VB.PictureBox picSplit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   30
         ScaleWidth      =   6255
         TabIndex        =   11
         Top             =   3720
         Width           =   6255
      End
      Begin VB.PictureBox pic未记 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   6000
         ScaleHeight     =   1185
         ScaleWidth      =   1545
         TabIndex        =   15
         Top             =   960
         Width           =   1575
         Begin VB.ListBox lst未记 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   1290
            ItemData        =   "frmCaseTendBodyData.frx":689E
            Left            =   0
            List            =   "frmCaseTendBodyData.frx":68A0
            TabIndex        =   16
            Top             =   0
            Width           =   2055
         End
      End
      Begin VB.PictureBox picValue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   3480
         ScaleHeight     =   1425
         ScaleWidth      =   1665
         TabIndex        =   13
         Top             =   -120
         Width           =   1695
         Begin zl9TemperatureChartJX.ColorPicker usrValue 
            Height          =   2190
            Left            =   -240
            TabIndex        =   14
            Top             =   -360
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   3863
         End
      End
      Begin VB.Frame fraDetail 
         Height          =   1935
         Left            =   0
         TabIndex        =   10
         Top             =   3840
         Width           =   7815
         Begin zl9TemperatureChartJX.VsfGrid vsfDetail 
            Height          =   1335
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   2355
         End
      End
      Begin VB.Frame fraData 
         Height          =   2895
         Left            =   0
         TabIndex        =   9
         Top             =   600
         Width           =   7335
         Begin zl9TemperatureChartJX.VsfGrid vsfCurve 
            Height          =   2535
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   4471
         End
      End
      Begin VB.Frame fraTime 
         Height          =   615
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9495
         Begin VB.PictureBox picToolBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   2775
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   210
            Width           =   2775
            Begin VB.OptionButton OptTime 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   0
               Width           =   350
            End
            Begin VB.Label lblPtime 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "时点:"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   0
               TabIndex        =   7
               Top             =   45
               Width           =   450
            End
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00～05:59"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   1080
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
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
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCaseTendBodyData.frx":68A2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsDate 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":6C3C
            Key             =   "preGreen"
            Object.Tag             =   "preGreen"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":764E
            Key             =   "preGray"
            Object.Tag             =   "preGray"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":8060
            Key             =   "nextGreen"
            Object.Tag             =   "nextGreen"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":8A72
            Key             =   "nextGray"
            Object.Tag             =   "nextGray"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":9484
            Key             =   "preLight"
            Object.Tag             =   "preLight"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":9E96
            Key             =   "nextLight"
            Object.Tag             =   "nextLight"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   8400
      ScaleHeight     =   360
      ScaleWidth      =   2760
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   2760
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   132907011
         CurrentDate     =   42285
      End
      Begin VB.Image imgbtn 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "frmCaseTendBodyData.frx":A8A8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
      Begin VB.Image imgbtn 
         Height          =   240
         Index           =   0
         Left            =   1920
         Picture         =   "frmCaseTendBodyData.frx":B2AA
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgDefault 
         Height          =   255
         Left            =   600
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.Timer tmrData 
      Interval        =   60
      Left            =   10080
      Top             =   840
   End
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   8640
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6720
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   0
         Width           =   75
      End
   End
   Begin MSComctlLib.ImageList ilsDetail 
      Left            =   7680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":BCAC
            Key             =   "detele"
            Object.Tag             =   "detele"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   6000
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5055
      _Version        =   589884
      _ExtentX        =   8916
      _ExtentY        =   10583
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7050
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodyData.frx":1250E
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   21167
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "消息提示信息"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "数据类型"
            TextSave        =   "数据类型"
            Key             =   "ZLDataType"
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
   Begin MSComctlLib.ImageList ilstab 
      Left            =   6480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCaseTendBodyData.frx":12DA2
            Key             =   "mark"
            Object.Tag             =   "mark"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const mFontSize As Integer = 9 '定义字体初始大小为9号字体
Private mblnTemType As Boolean  'TRUE为专科体温单,FALSE为标准体温单
Private Enum TYPE_Curve
    COL_Null = 0
    COL_编辑 = 1
    COL_分组名 = 2
    COL_字符串 = 3
    COL_项目序号 = 4
    COL_项目名 = 5
    col_原始时间 = 6
    COL_修改状态 = 7
    COL_项目名称 = 8
    COL_显示 = 9
    COL_时间 = 10
    COL_原值 = 11
    COL_数据 = 12
    COL_颜色 = 13
    COL_复试合格 = 14
    COL_部位 = 15
    Col_未记说明 = 16
    COL_来源 = 17
    COL_数据来源 = 18 '详细表格显示来源
    COL_删除 = 19
End Enum

Private Enum TYPE_Tab
    COL_TabNull = 0
    COL_tab字符串 = 1
    COL_tab项目序号 = 2
    col_tab原始时间 = 3
    COL_tab项目名 = 4  '--不包含单位
    COL_tab项目名称 = 5 '- -包含单位
    COL_tabDirect = 6
End Enum

Private Type Type_Item
    类型 As String
    值域 As String
    项目类型 As Integer
    项目小数 As Double
    记录频次 As Integer
    项目表示 As Integer
    项目性质 As Integer
    项目长度 As Long
    部位 As String
    项目号 As Long
    项目名 As String
    记录名 As String
    入院首测 As Integer
End Type

Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng文件ID As Long
    lng婴儿 As Long
    lng科室ID As Long
    lng护理等级 As Long
    lng病区ID As Long
    lng格式ID As Long
End Type
Private mT_Patient As type_Patient

Private Type Type_OptRow
    上标 As Integer
    下标 As Integer
End Type

Private mOptRow As Type_OptRow
    

'工具栏:
Private mcbrToolBar As CommandBar

Private mblnStart As Boolean
Private mblnMove As Boolean
Private mblnInit As Boolean
Private mblnEdit As Boolean
Private mblnOK As Boolean
Private mblnScroll As Boolean
Private mblnResize As Boolean
Private mblnAllRefresh As Boolean
Private mint心率应用 As Integer
Private mblnEdit心率 As Boolean
Private mblnFileBack As Boolean  '文件是否归档
Private mbln出院 As Boolean '病人出院或文件已经结束为TRUE
Private mbln录入小时 As Boolean  '全天汇总显示录入时间
Private mbln脉搏共用显示 As Boolean '脉搏是否以(心率/脉搏)方式录入
Private mblnRefresh首行 As Boolean
Private mbln汇总当天 As Boolean
Private mstrCurveItem As String  '专科体温单的曲线项目信息
Private mstrActiveItem As String '体温单活动项目信息
Private mstrOverDate As String '病人实际出院时间(即体温单实际终止时间)
Private mstrBegin As String '某段时间点的开始和结束时间 00:00-05:59
Private mstrEnd As String
Private mstrDate As String '体温单当前页的第一天时间
Private mstrBTime As String  '体温单的开始时间和结束时间
Private mstrETime As String
Private mstrPreOutDate As String '病人预出院时间
Private mstrSQL As String
Private mstr未记说明 As String
Private mintBigSize As Integer '是否放大
Private mintPreDays As Integer '超期录入时限
Private mlngHours As Long '数据补录时限
Private marrTime() As String

'记录集
Private mrsPart As New ADODB.Recordset '体温部位记录集
Private mrsNote As New ADODB.Recordset '上下标数据集
Private mrsCurve As New ADODB.Recordset  '体温曲线数据记录集
Private mrsTable As New ADODB.Recordset  '体温表表格显示据集
Private mrsTableDetail As New ADODB.Recordset '体温表表格数据明细数据集
Private mrsRecodeID As New ADODB.Recordset  '记录id数据集

Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int心率应用 As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------------------
'功能:调用体温单编辑窗体
'参数:frmParent 父窗体,strParam 格式:病人ID;主页Id;文件ID;婴儿;科室ID;护理护理等级  strTime 某段时间的时间范围 例如:2011-01-25 00:00:00;2011-01-25 05:59:59

'     strDayTime 一周开始时间; int心率应用=2 表示脉搏和心率公用 blnMove 历史数据是否转移
'     bytSize 0-9号字体 1-12号字体
'----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    Dim blnShowing As Boolean
    
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
        
    mblnStart = True
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mblnOK = False
    mblnResize = False
    mblnAllRefresh = False
    mbln汇总当天 = False
    mstrOverDate = ""
    
    mT_Patient.lng科室ID = 0
    mT_Patient.lng护理等级 = 3
    
    mT_Patient.lng病人ID = Val(arrParam(0))
    mT_Patient.lng主页ID = Val(arrParam(1))
    mT_Patient.lng文件ID = Val(arrParam(2))
    mT_Patient.lng婴儿 = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng科室ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng护理等级 = arrParam(5)
    
    If mT_Patient.lng病人ID = 0 And mT_Patient.lng主页ID = 0 And mT_Patient.lng科室ID = 0 Then
        MsgBox "文件ID,病人ID,主页ID不能为空,请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If Not OpenPatientInfo Then Exit Function
    
    mstrBegin = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss")
    mstrEnd = Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
    mstrDate = strDayTime
    
    If Not ChekPatientOut(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mint心率应用 = int心率应用
    mblnEdit心率 = True
    mblnMove = blnMove
    
    '动态加载时点按钮控件
    UnLoadOptTime
    LoadOptTime
    '检查文件是否归档
    mblnFileBack = CheckFileBack(mT_Patient.lng文件ID, mblnMove)
    '初始化工具栏
    Call InitCommandBars
    '初始化表格
    Call GetTableRowName
    '加载数据
    Call zlRefreshData
    mblnInit = True
    mblnResize = True
    Me.Show 1
    
    ShowEditor = mblnOK
End Function

Public Function OpenPatientInfo() As Boolean
'------------------------------------------------------
'提取病人基本信息
'------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    '提取病人科室
    strSQL = " select 出院科室id from 病案主页 where 病人id=[1] and 主页id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTemp.BOF = False Then
        mT_Patient.lng科室ID = Val(zlStr.Nvl(rsTemp("出院科室ID").Value))
    End If
    
    '提取病人护理等级
    strSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTemp.BOF = False Then
        mT_Patient.lng护理等级 = Val(zlStr.Nvl(rsTemp("护理等级").Value))
    End If
    
    '提取体温单基本信息
    mblnTemType = False
    strSQL = "Select B.子类,B.ID From 病人护理文件 A,病历文件列表 B Where A.格式ID=B.ID And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng文件ID)
    If rsTemp.BOF = False Then
        mblnTemType = (Nvl(rsTemp!子类) = "1")
        mT_Patient.lng格式ID = rsTemp!Id
    End If
    
    If mblnTemType = True Then
        gintHourBegin = T_BodyStyle.lng开始时点
    Else
        gintHourBegin = zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4)
        T_BodyStyle.lng开始时点 = gintHourBegin
        T_BodyStyle.lng时间间隔 = 4
        T_BodyStyle.lng监测次数 = 6
        T_BodyStyle.lng天数 = 7
    End If
    OpenPatientInfo = True
    Exit Function
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function ChekPatientOut(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'功能:提取体温单开始时间和结束时间 并检查病人是否出院
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    mbln出院 = False
    On Error GoTo Errhand
    
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("数据补录时限", glngSys)), 1, 6))
    mbln汇总当天 = (Val(zlDatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    mbln录入小时 = (Val(zlDatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    mbln脉搏共用显示 = (Val(zlDatabase.GetPara("脉搏短绌以(心率/脉搏)方式录入", glngSys, 1255, 0)) = 1)
    If mintPreDays < 0 Then mintPreDays = 0
    
    '提取病人预出院时间
    strSQL = "Select 开始时间 From 病人变动记录 where 病人ID=[1] and 主页ID=[2] And 开始原因=10"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng病人ID, lng主页ID)
    If Not rsTemp.EOF Then mstrPreOutDate = Format(rsTemp!开始时间, "YYYY-MM-DD HH:mm:ss")
    
    '提取婴儿医嘱信息(转科，出院),存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "(SELECT " & vbNewLine & _
                "        病人ID, 主页ID, 婴儿时间, DECODE(NVL(婴儿, 0), 0, DECODE(NVL(出院日期, ''), '', 0, 1), DECODE(NVL(婴儿时间, ''), '', 0, 1)) 记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID, A.主页ID, B.开始执行时间 婴儿时间, A.出院日期, B.婴儿" & vbNewLine & _
                "              FROM 病案主页 A," & vbNewLine & _
                "                   (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                     FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                     WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND NVL(B.婴儿, 0) <> 0 AND B.诊疗类别 = 'Z' " & vbNewLine & _
                "                      AND Instr(',3,5,11,', ',' || c.操作类型 || ',') > 0 AND B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "              WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "              ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2) E"

    '说明:目前有了专科体温单，病人可能同时存在多份体温单。体温单开始时间和终止时间的规则如下:
    '如果文件的开始时间不为空并且大于等于病人入院时间或婴儿出生时间,体温单的开始时间以文件开始时间为准,否则以病人入院时间或婴儿出生时间为准
    '如果文件的终止时间不为空并且小于等于病人或婴儿出院时间（未出院不能大于当前时间）,体温单结束时间以文件开始时间为准，否则体温单结束时间以病人或婴儿出院时间为准(未出院为当前时间)
    '如果文件的终止时间为空,保持原有方式,病人如果已经出院，就已出院时间为准,未出院就已当前时间或数据结束时间为准.
    strSQL = " SELECT  DECODE(D.开始时间,NULL,DECODE(B.出生时间, NULL, A.开始, B.出生时间)," & vbNewLine & _
            "               DECODE(SIGN(D.开始时间 - DECODE(B.出生时间, NULL, A.开始, B.出生时间))," & vbNewLine & _
            "                      1," & vbNewLine & _
            "                      D.开始时间," & vbNewLine & _
            "                      DECODE(B.出生时间, NULL, A.开始, B.出生时间))) AS 开始," & vbNewLine & _
            "       DECODE(D.结束时间," & vbNewLine & _
            "               NULL," & vbNewLine & _
            "               DECODE(E.记录," & vbNewLine & _
            "                      0," & vbNewLine & _
            "                      DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.发生时间), 1, NVL(E.婴儿时间, A.终止), D.发生时间)," & vbNewLine & _
            "                      NVL(E.婴儿时间, A.终止))," & vbNewLine & _
            "               DECODE(SIGN(NVL(E.婴儿时间, A.终止) - D.结束时间), 1, D.结束时间, NVL(E.婴儿时间, A.终止))) 终止," & vbNewLine & _
            "       DECODE(D.结束时间, NULL, E.记录, 1) 记录" & vbNewLine & _
            " FROM (SELECT 病人ID, 主页ID, MIN(开始时间) AS 开始, MAX(NVL(终止时间, SYSDATE)) AS 终止" & vbNewLine & _
            "       FROM 病人变动记录" & vbNewLine & _
            "       WHERE 开始时间 IS NOT NULL AND 病人ID = [2] AND 主页ID = [3]" & vbNewLine & _
            "       GROUP BY 病人ID, 主页ID) A," & vbNewLine & _
            "     (SELECT 病人ID, 主页ID, 出生时间 FROM 病人新生儿记录 WHERE 病人ID = [2] AND 主页ID = [3] AND 序号 = [4]) B," & vbNewLine & _
            "     (SELECT NVL(发生时间, SYSDATE) 发生时间, 开始时间, 结束时间" & vbNewLine & _
            "       FROM (SELECT MAX(B.发生时间) 发生时间, MAX(A.开始时间) 开始时间, MAX(A.结束时间) 结束时间" & vbNewLine & _
            "              FROM 病人护理文件 A, 病人护理数据 B" & vbNewLine & _
            "              WHERE A.ID = B.文件ID(+) AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND A.婴儿 = [4])) D," & vbNewLine & _
            "  " & strNewSql & vbNewLine & _
            " WHERE A.病人ID = E.病人ID AND A.主页ID = E.主页ID AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng文件ID, lng病人ID, lng主页ID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
        mbln出院 = Not (Val(rsTemp!记录) = 0)
    Else
        MsgBox "无此病人本次住院信息,请检查!", vbInformation, gstrSysName '无数病人变动信息退出
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln出院 Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '病人出院以出院时间为终止时间
    If mbln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        mstrETime = Format(RetrunEndTimeNew(CDate(mstrBTime), CDate(mstrETime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        strMaxDate = Format(mstrETime, "YYYY-MM-DD")
    Else
        intDay = mintPreDays - DateDiff("D", CDate(strCurrDate), CDate(mstrETime))
        If intDay < 0 Then intDay = 0
        strMaxDate = Format(DateAdd("d", intDay, CDate(mstrETime)), "yyyy-MM-dd")
        If CDate(mstrETime) < CDate(Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")) Then
            mstrETime = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    mstrETime = Format(strMaxDate & " " & Format(mstrETime, "HH:mm:ss"), "yyyy-MM-DD HH:mm:ss")
    
    If Not (CDate(mstrBegin) >= CDate(mstrBTime) And CDate(mstrBegin) <= CDate(mstrETime)) Then
        If Int(CDate(mstrBTime)) = Int(CDate(mstrETime)) Then
            mstrBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        Else
            mstrBegin = Format(Int(CDate(mstrETime)), "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    If Not (CDate(mstrEnd) >= CDate(mstrBegin) And CDate(mstrEnd) <= CDate(mstrETime)) Then
        mstrEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    dtpDate.Value = Format(mstrBegin, "YYYY-MM-DD")
    dtpDate.MaxDate = Format(strMaxDate, "YYYY-MM-DD")
    dtpDate.MinDate = Format(mstrBTime, "YYYY-MM-DD")
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Function CheckFileBack(ByVal lngID As Long, ByVal blnMove As Boolean) As Boolean
'---------------------------------------------------------------
'功能:检查文件是否归档
'---------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    On Error GoTo Errhand
    
    CheckFileBack = False
    strSQL = "Select 1 From 病人护理文件 Where Id=[1] And 归档人 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查文件是否归档", lngID)
    If blnMove = True Then
        strSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
    End If
    If rsTemp.RecordCount > 0 Then
        CheckFileBack = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckDateTime(ByVal lngRow As Long, ByVal strName As String, ByVal strTime As String) As Boolean
    '-----------------------------------------
    '功能：检查录入数据时间是否超过补录范围
    '-----------------------------------------
    Dim strInfo As String
    Dim strErrMsg As String
    Dim strCurrDate As String
    Dim strText As String
    Dim strCenterTime As String
    Dim arrTime() As String
    
    On Error GoTo Errhand
    If lngRow <> 0 Then
        strInfo = "第" & lngRow & "行"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    strText = strTime
    
    arrTime = Split(Trim(strText), ":")
    
    If UBound(arrTime) <> 1 Then
        strInfo = "录入的时点格式非法！[小时:分钟]"
        Exit Function
    Else
        If Len(Trim(arrTime(0))) < 2 Then arrTime(0) = String(2 - Len(Trim(arrTime(0))), "0") & Trim(arrTime(0))
        If Len(Trim(arrTime(1))) < 2 Then arrTime(1) = String(2 - Len(Trim(arrTime(1))), "0") & Trim(arrTime(1))
        strText = arrTime(0) & ":" & arrTime(1)
    End If
    
    '合法性检查
    If IsNumeric(arrTime(0)) = False Or IsNumeric(arrTime(1)) = False Or Len(Trim(arrTime(0))) > 2 Or Len(Trim(arrTime(1))) > 2 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "录入的时点格式非法！[小时:分钟]"
        Exit Function
    End If
    If Mid(strText, 3, 1) <> ":" Then
        lblStb.ForeColor = 255
        lblStb.Caption = "录入的时点格式非法！[小时:分钟]"
        Exit Function
    End If
    If Val(arrTime(0)) < 0 Or Val(arrTime(0)) > 23 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "录入的时点格式非法！[小时应在0至23之间]"
        Exit Function
    End If
    If Val(arrTime(1)) < 0 Or Val(arrTime(1)) > 59 Then
        lblStb.ForeColor = 255
        lblStb.Caption = "录入的时点格式非法！[分钟应在0至59之间]"
        Exit Function
    End If
    strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
    If Format(mstrETime, "hh:mm:ss") < Split(lblTime, "～")(1) And mstrETime > mstrEnd Then mstrEnd = mstrETime
    If Not (CDate(Format(strTime)) >= CDate(mstrBegin) And CDate(strTime) <= CDate(mstrEnd)) Then
        lblStb.ForeColor = 255
        lblStb.Caption = "需输入的时间在 " & Format(mstrBegin, "hh:mm") & "～" & Format(mstrEnd, "hh:mm") & " 时间段之间"
        Exit Function
    End If
    
    If Not IsDate(strTime) Then Exit Function
    
    If DateDiff("m", CDate(Format(strTime, "YY-MM-DD hh:mm")), CDate(Format(mstrETime, "YY-MM-DD hh:mm"))) < 0 Then
        If mbln出院 = False Then
            strErrMsg = strInfo & "记录数据时间已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围!"
        Else
            strErrMsg = strInfo & "记录数据时间不能大于[病人出院时间或文件结束时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    If DateDiff("m", CDate(Format(strTime, "YY-MM-DD hh:mm")), CDate(Format(mstrBTime, "YY-MM-DD hh:mm"))) > 0 Then
        strErrMsg = strInfo & "记录数据时间不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, strTime, strCurrDate) Then
        strErrMsg = strInfo & "记录数据时间[" & strTime & "]有误![超过数据补录的有效时限:" & mlngHours & "小时]"
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    CheckDateTime = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function dtpDateChageDate(ByVal strValue As String) As Boolean
'------------------------------------------------------------------------------
'补录时间合法时，发生变化就刷新数据
'------------------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String, strTime As String
    Dim i As Integer
    Dim strCurrDate As String
    Dim intBound As Integer
    Dim strBegin As String, strEnd As String
    Dim intCOl As Integer
    Dim strCurDate As String
    Dim intDay As Integer
    Dim strBTime As String
    
    On Error GoTo Errhand
    lblStb.Tag = lblStb.Caption
    
    If Format(strValue, "YYYY-MM-DD") > Format(mstrETime, "YYYY-MM-DD") Then
        If mbln出院 = False Then
            strErrMsg = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
        Else
            strErrMsg = "录入的日期不能大于[病人出院时间或文件结束时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "录入的日期不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If Format(strValue, "YYYY-MM-DD") = mstrETime Then
        strDate = Format(Format(mstrETime, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    ElseIf Format(strValue, "YYYY-MM-DD") = mstrBTime Then
        strDate = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        strTime = strDate
    Else
        strDate = Format(Format(strValue, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(Format(strValue, "YYYY-MM-DD") & " 23:59:00", "YYYY-MM-DD HH:mm:ss")
    End If
    
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, strTime, strCurrDate) Then
        strErrMsg = "录入的时间[" & strValue & "]有误！[超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ErrInfo
    End If
    
    mblnAllRefresh = True
    
    If UBound(marrTime) = -1 Then Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng监测次数, T_BodyStyle.lng时间间隔)
    intDay = DateDiff("D", CDate(mstrBTime), CDate(strValue)) \ T_BodyStyle.lng天数
    intDay = (intDay) * T_BodyStyle.lng天数
    strBTime = Format(DateAdd("d", intDay, CDate(mstrBTime)), "yyyy-MM-dd") & " 00:00:00"
    
    If Format(strValue, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(strBTime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(strBTime, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumnNew(strCurDate, strBTime, gintHourBegin)
        strDate = GetCurveDateNew(intCOl, strBTime, gintHourBegin)
        strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
    Else
         If Format(strValue, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumnNew(mstrETime, strBTime, gintHourBegin)
            strDate = GetCurveDateNew(intCOl, mstrBTime, gintHourBegin)
            strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
         ElseIf Format(strValue, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(strValue, "YYYY-MM-DD") < Format(mstrETime, "YYYY-MM-DD") Then
            strDate = GetCenterTime(Format(strValue, "YYYY-MM-DD 21:00:00"), Format(strValue, "YYYY-MM-DD 23:59:59"))
         End If
    End If

    For i = 0 To UBound(marrTime)
        If Format(strDate, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strDate, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    
    If i > UBound(marrTime) Then i = 0
    
    strBegin = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(0), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(1), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    
    Call GetCenterTime(CDate(strBegin), CDate(strEnd), intBound)

    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * T_BodyStyle.lng时间间隔
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i
    
    Call zlRefreshData(True, True)
    Call OptTime_Click(intBound)
    
    
    mblnAllRefresh = False
    dtpDateChageDate = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
    mblnAllRefresh = False
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function IsAllowInput(ByVal lng病人ID As Long, ByVal lng主页ID As Long, lng婴儿 As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '----------------------------------------------
    '功能：取出病人发生变动记录的时间点
    '----------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strBabyOutTime As String
    
    On Error GoTo Errhand
    IsAllowInput = True
    If lng婴儿 <> 0 And mbln出院 = True Then
        strBabyOutTime = GetAdviceOutTime(lng病人ID, lng主页ID, lng婴儿)
        If strBabyOutTime <> "" Then
            strTime = Format(DateAdd("H", mlngHours, strBabyOutTime), "yyyy-MM-dd HH:mm")
            GoTo GONext
        End If
    End If
    gstrSQL = "" & _
              " SELECT DECODE(终止原因,1,'出院',3,'转科',10,'预出院',15,'转病区',DECODE(开始原因,10,'出院','未定义')) AS 类型,终止时间 AS 时间" & _
              " From 病人变动记录" & _
              " WHERE (终止原因 IN (1,3,10,15) OR 开始原因=10) And 病人ID=[1] And 主页ID=[2] And [3] <= 终止时间" & _
              " ORDER BY 终止时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出指定病人在指定时间之后关键点的时间", lng病人ID, lng主页ID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    '只取第一条符合的记录
    strTime = Format(DateAdd("H", mlngHours, rsTemp!时间), "yyyy-MM-dd HH:mm")
GONext:
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'功能:初始化工具栏
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim cboChild As CommandBarPopup
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    
     '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsMain.Add("标准", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "添加项目"): cbrControl.ToolTipText = "添加活动项目": cbrControl.BeginGroup = True

        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "特殊处理")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "正常", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "灌肠[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "灌肠后大便[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cboChild = cbrPop.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Edit_Append * 10 + 3, "大便失禁", -1, False): cbrControl.IconId = 1
        Set cbrControl = cboChild.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 30, "※", -1, False):  cbrControl.Parameter = "※"
        Set cbrControl = cboChild.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 31, "*", -1, False): cbrControl.Parameter = "*"
        
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "人工肛门[☆]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "☆"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "导尿[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "保留导尿[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With picDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = (dtpDate.Width + 520) + (dtpDate.Width + 520) * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    With dtpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
        .Top = 0
        .Left = 0
    End With

    With imgbtn(1)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = dtpDate.Width + 20
    End With
    With imgbtn(0)
        .Width = 240 + 240 * mintBigSize / 3
        .Height = 240 + 240 * mintBigSize / 3
        .Top = 30
        .Left = dtpDate.Width + imgbtn(1).Width + 30
    End With
    
    '超期补录
    '------------------------------------------------------------------------------------------------------------------
    Set cbrLable = mcbrToolBar.Controls.Add(xtpControlLabel, conMenu_View_Option, "")
    cbrLable.flags = xtpFlagRightAlign
    Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    dtpDate.Visible = True
    cbrCustom.Handle = picDate.hWnd
    imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    cbrCustom.flags = xtpFlagRightAlign
    
'    Set cbrControl = mcbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Forward, "上一天")
'    cbrControl.flags = xtpFlagRightAlign
'    cbrControl.IconId = conMenu_View_Forward
'    If dtpDate.Value = dtpDate.MinDate Then cbrControl.Enabled = False
'    Set cbrControl = mcbrToolBar.Controls.Add(xtpControlButton, conMenu_View_Backward, "下一天")
'    cbrControl.flags = xtpFlagRightAlign
'    cbrControl.IconId = conMenu_View_Backward
'    If dtpDate.Value = dtpDate.MaxDate Then cbrControl.Enabled = False

    '快键绑定
    With cbsMain.KeyBindings
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 30)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 31)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("7"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem '添加活动项目
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '保存
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse '取消
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Call InitDateTimeRange(marrTime, gintHourBegin, T_BodyStyle.lng监测次数, T_BodyStyle.lng时间间隔)
     
    '加载表格控件
    Call InitTabControl
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Function GetTableRowName() As String
'----------------------------------------------------
'初始化表格
'----------------------------------------------------
    Dim arrItem() As Variant
    Dim strSQL As String
    Dim strTmp As String
    Dim str值域 As String
    Dim strCurDate As String
    Dim strEndTime As String
    Dim strDate As String
    Dim strTmpCurve As String '曲线项目变量
    Dim strTmpTable As String '表格项目变量
    Dim i As Integer, intBound As Integer
    Dim intCOl As Integer
    Dim Titem As Type_Item
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    arrItem = Array()
    Call InitRecordSet
    Call GetPainDegreeNO
     '检查脉搏心率共用时心率是否使用与此病人
    strSQL = "select C.应用方式 From 护理记录项目 C where C.项目序号=[1] And C.护理等级>=[2] And Nvl(C.适用病人,0) In (0,[3]) " & _
            " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[4])))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取心率", -1, mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID)
    mblnEdit心率 = IIf(rsTemp.RecordCount = 0, False, True)
    If rsTemp.RecordCount > 0 Then mint心率应用 = Val(zlStr.Nvl(rsTemp!应用方式, 0))
    
    '格式组成为 类型'值域,项目类型,项目小数,记录频次,项目表示,项目性质,项目长度,部位,入院首测'项目号'项目名
    strTmp = "2)上下标说明',,,,,,,,'2'上标'上标;2)上下标说明',,,,,,,,'6'下标'下标"
    
    '提取全部体温曲线项目
    mstrCurveItem = ""
    mstrCurveItem = T_BodyItem.str曲线项目
    If InStr(1, "," & mstrCurveItem & ",", "," & gint呼吸 & ",") = 0 Then
        If InStr(1, Val(T_BodyItem.str表格内容), gint呼吸) > 0 Then
            mstrCurveItem = mstrCurveItem & "," & gint呼吸
        End If
    End If
    strSQL = " Select /*+ RULE */" & _
             " a.排列序号, a.记录名 项目名, a.项目序号 As 项目号, a.记录法, a.入院首测, c.项目值域, c.项目类型, c.项目长度, c.项目小数, Nvl(a.记录频次, 2) 记录频次, c.分组名, c.项目表示," & _
             "  c.项目单位 " & vbNewLine & _
             "  From 体温记录项目 A, 诊治所见项目 B, 护理记录项目 C, Table(Cast(f_Num2list([1]) As Zltools.t_Numlist)) D " & vbNewLine & _
             " Where c.项目id = b.Id(+) And a.项目序号 = c.项目序号 And (a.记录法 <> 2 Or (a.记录法 = 2 And a.项目序号 = 3)) And " & vbNewLine & _
             "      Not (c.应用方式 = 2 And c.项目序号 = -1) And c.项目序号 = d.Column_Value " & vbNewLine & _
             " Order By Decode(a.项目序号, 1, 0, 1), a.排列序号 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrCurveItem)
    
    strTmpCurve = ""
    With rsTemp
        Do While Not .EOF
            str值域 = Replace(zlStr.Nvl(!项目值域), ":", "")
            If zlStr.Nvl(!项目类型) = 0 Then
                If InStr(1, str值域, ";") Then str值域 = Split(str值域, ";")(0) & "～" & Split(str值域, ";")(1)
            End If
            str值域 = Replace(Replace(Replace(str值域, ";", ":"), "'", ""), ",", "")
            Titem.值域 = str值域
            Titem.项目名 = Replace(Replace(zlStr.Nvl(!项目名) & IIf(zlStr.Nvl(!项目单位, "") = "", "", "(" & !项目单位 & ")"), ";", ":"), "'", "")
            Titem.记录名 = zlStr.Nvl(!项目名)
            Titem.项目号 = Val(zlStr.Nvl(!项目号))
            Titem.入院首测 = Val(zlStr.Nvl(!入院首测, 0))
            Titem.项目类型 = Val(zlStr.Nvl(!项目类型, 0))
            Titem.项目长度 = Val(zlStr.Nvl(!项目长度, 3))
            Titem.项目小数 = Val(zlStr.Nvl(!项目小数, 0))
            Titem.记录频次 = Val(zlStr.Nvl(!记录频次))
            Titem.项目表示 = Val(zlStr.Nvl(!项目表示, 0))
            If Titem.项目表示 = 4 Or IsWaveItem(Titem.项目号) Then
                If Titem.记录频次 > 2 Then Titem.记录频次 = 2
            End If
            Titem.部位 = ""
            Titem.项目性质 = 1
            '记录法为1和记录法为2的呼吸项目为曲线项目
            Titem.类型 = "1)体温曲线项目"
            strTmpCurve = strTmpCurve & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & Titem.项目性质 & Titem.项目长度 & Titem.部位 & Titem.入院首测 & "'" & _
                Titem.项目号 & "'" & Titem.项目名 & "'" & Titem.记录名
        .MoveNext
        Loop
    End With
    
    strEndTime = DateAdd("d", T_BodyStyle.lng天数, CDate(Format(Format(mstrDate, "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")))
    If strEndTime > mstrETime Then strEndTime = mstrETime
    mstrActiveItem = ""
    Set rsTemp = GetAppendGridItemNew(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng护理等级, mT_Patient.lng婴儿, _
            CDate(mstrDate), CDate(strEndTime), IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID, T_BodyItem.str表格项目, mblnMove)
    strTmpTable = ""
    With rsTemp
        Do While Not .EOF
            str值域 = Replace(zlStr.Nvl(!项目值域), ":", "")
            If zlStr.Nvl(!项目类型) = 0 Then
                If InStr(1, str值域, ";") <> 0 Then str值域 = Split(str值域, ";")(0) & "～" & Split(str值域, ";")(1)
            End If
            str值域 = Replace(Replace(Replace(str值域, ";", ":"), "'", ""), ",", "")
            Titem.值域 = str值域
            Titem.类型 = "2)体温表格项目"
            Titem.项目类型 = Val(zlStr.Nvl(!项目类型))
            Titem.项目小数 = Val(zlStr.Nvl(!项目小数, 0))
            Titem.记录频次 = Val(zlStr.Nvl(!记录频次, 2))
            Titem.项目表示 = Val(zlStr.Nvl(!项目表示, 0))
            Titem.项目性质 = Val(zlStr.Nvl(!项目性质, 1))
            Titem.项目长度 = zlStr.Nvl(!项目长度, 3)
            Titem.部位 = Replace(Replace(Replace(zlStr.Nvl(!体温部位), ";", ""), "'", ""), ",", "")
            Titem.项目号 = Val(zlStr.Nvl(!项目序号))
            Titem.项目名 = Replace(Replace(IIf(Titem.项目号 = 4, "血压", zlStr.Nvl(!记录名)) & IIf(zlStr.Nvl(!单位, "") = "", "", "(" & !单位 & ")"), ";", ":"), "'", "")
            Titem.入院首测 = Val(zlStr.Nvl(!入院首测, 0))
            Titem.记录名 = IIf(Titem.项目号 = 4, "血压", zlStr.Nvl(!记录名))
            
            If Titem.项目表示 = 4 Or IsWaveItem(Titem.项目号) Then
                If Titem.记录频次 > 2 Then Titem.记录频次 = 2
            End If
            
            If Titem.项目号 <> gint呼吸 And Titem.项目号 <> 5 Then
                strTmpTable = strTmpTable & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                    Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & "," & Titem.项目性质 & "," & Titem.项目长度 & "," & _
                    Titem.部位 & "," & Titem.入院首测 & "'" & Titem.项目号 & "'" & Titem.项目名 & "'" & Titem.记录名
                '活动项目
                If Titem.项目性质 = 2 Then
                    mstrActiveItem = mstrActiveItem & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                        Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & "," & Titem.项目性质 & "," & Titem.项目长度 & "," & _
                        Titem.部位 & "," & Titem.入院首测 & "'" & Titem.项目号 & "'" & Titem.项目名 & "'" & Titem.记录名
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If Left(mstrActiveItem, 1) = ";" Then mstrActiveItem = Mid(mstrActiveItem, 2)
    If strTmp <> "" Then strTmpCurve = strTmpCurve & ";" & strTmp
    If Left(strTmpCurve, 1) = ";" Then strTmpCurve = Mid(strTmpCurve, 2)
    If Left(strTmpTable, 1) = ";" Then strTmpTable = Mid(strTmpTable, 2)
    
    '初始化体温曲线数据包括上下标
    Call InitTabCurve(strTmpCurve)
    '初始化体温表格
    Call InitTabTable(strTmpTable)
    '提取未记说明
    mstr未记说明 = ""
    mrsCurInfo.Filter = ""
    mrsCurInfo.Sort = "编码"
    With mrsCurInfo
        Do While Not .EOF
            mstr未记说明 = IIf(mstr未记说明 = "", "", mstr未记说明 & "'") & zlStr.Nvl(!名称)
            .MoveNext
        Loop
    End With
    If Left(mstr未记说明, 1) = "'" Then mstr未记说明 = Mid(mstr未记说明, 2)
    
    '根据选择时间定位当前时间编辑状态
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    If Format(mstrBegin, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(strEndTime, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumnNew(strCurDate, mstrBegin, gintHourBegin)
        strDate = GetCurveDateNew(intCOl, mstrBegin, gintHourBegin)
        mstrBegin = Split(strDate, ";")(0)
        mstrEnd = Split(strDate, ";")(1)
    Else
         If Format(mstrBegin, "YYYY-MM-DD") = Format(strEndTime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumnNew(mstrEnd, mstrBegin, gintHourBegin)
            strDate = GetCurveDateNew(intCOl, mstrBegin, gintHourBegin)
            mstrBegin = Split(strDate, ";")(0)
            mstrEnd = Split(strDate, ";")(1)
         ElseIf Format(mstrBegin, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(mstrBegin, "YYYY-MM-DD") < Format(strEndTime, "YYYY-MM-DD") Then
            mstrBegin = Format(mstrBegin, "YYYY-MM-DD 21:00:00")
            mstrEnd = Format(mstrBegin, "YYYY-MM-DD 23:59:59")
         End If
    End If
    
    '获取当前时间点在当天的第几格位置上
    Call GetCenterTime(CDate(mstrBegin), CDate(mstrEnd), intBound)
    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * T_BodyStyle.lng时间间隔
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "～" & Format(mstrEnd, "HH:mm")
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function IsTatle(ByVal lngItemNO As Long) As Boolean
'---------------------------------------------------------
'检查是否是汇总项目
'---------------------------------------------------------
    If mrsCollect Is Nothing Then Exit Function
    If mrsCollect.State = adStateOpen Then
        mrsCollect.Filter = "序号=" & lngItemNO
        IsTatle = mrsCollect.RecordCount > 0
    End If
End Function


Private Sub InitTabCurve(ByVal strTabName As String)
'-------------------------------------------------------
'功能:初始化体温曲线项目
'参数:所有表头的信息
'-------------------------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    
    On Err GoTo Errhand
    If strTabName = "" Then Exit Sub
    varTabName = Split(strTabName, ";")
    With vsfCurve
    
        .Rows = UBound(varTabName) + 2
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "编辑", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "分组名", 1500 + 1500 * mintBigSize / 3, 1
        .NewColumn "字符串", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目序号", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目名", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "原始时间", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "修改状态", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目名称", 1200 + 1200 * mintBigSize / 3, 1
        .NewColumn "显示", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "时间", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "原值", 300 + 300 * mintBigSize / 3, 1
        .NewColumn "数据", 2300 + 2300 * mintBigSize / 3, 1, , 4
        .NewColumn "数据", 300 + 300 * mintBigSize / 3, 0
        .NewColumn "复试合格", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "部位", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "未记说明", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .NewColumn "来源", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "数据来源", 300 + 300 * mintBigSize / 3, 1
        .NewColumn "删除", 900 + 900 * mintBigSize / 3, 4
        .Body.RowHeight(0) = 300 + 300 * mintBigSize / 3
        .FixedCols = COL_项目名称 + 1
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_编辑) = True
        .Body.ColHidden(COL_字符串) = True
        .Body.ColHidden(COL_项目序号) = True
        .Body.ColHidden(COL_项目名) = True
        .Body.ColHidden(col_原始时间) = True
        .Body.ColHidden(COL_修改状态) = True
        .Body.ColHidden(COL_显示) = True
        .Body.ColHidden(COL_原值) = True
        .Body.ColHidden(COL_来源) = True
        .Body.ColHidden(COL_数据来源) = True
        .Body.ColHidden(COL_删除) = True
        .Body.WordWrap = True
        .Body.MergeCells = flexMergeRestrictColumns
        .Body.MergeCol(COL_分组名) = True
        .Body.MergeRow(0) = True
        
        For intRow = .FixedRows To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            If UBound(varCode) > 2 Then
                .TextMatrix(intRow, COL_分组名) = varCode(0)
                .TextMatrix(intRow, COL_字符串) = varCode(1)
                .TextMatrix(intRow, COL_项目序号) = varCode(2)
                .TextMatrix(intRow, COL_项目名称) = varCode(3)
                .TextMatrix(intRow, COL_项目名) = varCode(4)
                .TextMatrix(intRow, COL_数据) = Space(2)
                .TextMatrix(intRow, COL_颜色) = Space(2)
                If varCode(0) = "2)上下标说明" Then
                    Select Case Val(varCode(2))
                        Case 2
                            mOptRow.上标 = intRow
                        Case 6
                            mOptRow.下标 = intRow
                    End Select
                End If
            End If
            .Body.RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            .RowData(intRow) = 0
        Next intRow
        .Body.MergeRow(intRow - 2) = True
        .Body.Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Body.Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With

    With vsfDetail
        .FixedRows = 1
        .Rows = .FixedRows + 1
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "编辑", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "", 255, 4
        .NewColumn "字符串", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目序号", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目名", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "原始时间", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "修改状态", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "项目名称", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "显示", 700 + 700 * mintBigSize / 3, 1, , 4
        .NewColumn "时间", 900 + 900 * mintBigSize / 3, 1, , 4
        .NewColumn "原值", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "数据", 1200 + 1200 * mintBigSize / 3, 1, , 4
        .NewColumn "颜色", 900 + 900 * mintBigSize / 3, 1
        .NewColumn "复试合格", 1200 + 1200 * mintBigSize / 3, 1, , 4
        .NewColumn "部位", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "未记说明", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .NewColumn "来源", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "数据来源", 1400 + 1400 * mintBigSize / 3, 1, , 4
        .NewColumn "", 350 + 350 * mintBigSize / 3, 1, , 4
        .Body.RowHeightMin = 300 + 300 * mintBigSize / 3
        .Body.ColComboList(COL_部位) = " "
        .Body.ColComboList(Col_未记说明) = "..."
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_编辑) = True
        .Body.ColHidden(COL_Null) = True
        .Body.ColHidden(COL_项目序号) = True
        .Body.ColHidden(COL_项目名) = True
        .Body.ColHidden(col_原始时间) = True
        .Body.ColHidden(COL_项目名称) = True
        .Body.ColHidden(COL_字符串) = True
        .Body.ColHidden(COL_修改状态) = True
        .Body.ColHidden(COL_原值) = True
        .Body.ColHidden(COL_颜色) = True
        .Body.ColHidden(COL_来源) = True
        .Body.WordWrap = False
        .FixedCols = COL_分组名 + 1
        .Body.AllowUserResizing = flexResizeNone
        .Body.Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Body.Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabTable(ByVal strTabName As String)
    '--------------------------------------
    '初始化曲线表格和详细表格列头
    '--------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    Dim i As Integer, lngPreNum As Long
    
    On Error GoTo Errhand
    If strTabName = "" Then
        strTabName = "',0,0,0,0,1,0,,0'-999''"
        vsfTab.Tag = "NO"
    Else
        vsfTab.Tag = ""
    End If
    varTabName = Split(strTabName, ";")
    With vsfTab
        .Cols = 13
        .Rows = UBound(varTabName) * 2 + 3
        .FixedRows = 1
        .FixedCols = 7
        .ColHidden(COL_tab字符串) = True
        .ColHidden(COL_tab项目序号) = True
        .ColHidden(COL_tab项目名) = True
        .ColHidden(col_tab原始时间) = True
        .WordWrap = True
        .MergeCells = flexMergeRestrictColumns
        .MergeCol(COL_tab项目名称) = True
        .MergeRow(0) = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        .ColWidth(COL_TabNull) = 255
        .ColWidth(COL_tab项目名称) = 1200
        .ColWidth(COL_tabDirect) = 600
        .RowHeightMin = 300 + 300 * mintBigSize / 3
        For intCOl = .FixedCols - 2 To .Cols - 1
            If intCOl < .FixedCols Then
                .TextMatrix(0, intCOl) = "名称/频次"
            Else
                .TextMatrix(0, intCOl) = intCOl - .FixedCols + 1
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        
        i = 1
        For intRow = 1 To .Rows - 1
            varCode = Split(varTabName(i - 1), "'")
            .RowData(intRow) = Split(varCode(1), ",")(3) & ";" & IIf(IsWaveItem(varCode(2)), 2, 0)
            If IsTatle(varCode(2)) Then .RowData(intRow) = Split(varCode(1), ",")(3) & ";" & "3"
            .RowData(intRow + 1) = .RowData(intRow)
            .TextMatrix(intRow, COL_tab字符串) = varCode(1)
            .TextMatrix(intRow, COL_tab项目序号) = varCode(2)
            .TextMatrix(intRow, COL_tab项目名) = varCode(4)
            .TextMatrix(intRow, COL_TabNull) = ""
            .TextMatrix(intRow, COL_tab项目名称) = varCode(3)
            If Split(varCode(1), ",")(3) > 0 Then .TextMatrix(intRow, col_tab原始时间) = Replace(Space(Split(varCode(1), ",")(3) - 1), " ", "'")
            .TextMatrix(intRow, COL_tabDirect) = "时间"
            .TextMatrix(intRow + 1, COL_tab字符串) = varCode(1)
            .TextMatrix(intRow + 1, COL_tab项目序号) = varCode(2)
            .TextMatrix(intRow + 1, COL_tab项目名) = varCode(4)
            .TextMatrix(intRow + 1, COL_TabNull) = ""
            .TextMatrix(intRow + 1, COL_tab项目名称) = varCode(3)
            If Split(varCode(1), ",")(3) > 0 Then .TextMatrix(intRow + 1, col_tab原始时间) = Replace(Space(Split(varCode(1), ",")(3) - 1), " ", "'")
            .TextMatrix(intRow + 1, COL_tabDirect) = "数据"
            intRow = intRow + 1
            i = i + 1
        Next intRow
        .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = &H80000005
        '根据背景频次决定背景色
        For intRow = .FixedRows To .Rows - 1
            If .FixedCols + (Val(Split(.RowData(intRow), ";")(0))) < .Cols Then
                .Cell(flexcpBackColor, intRow, .FixedCols + (Val(Split(.RowData(intRow), ";")(0))), intRow, .Cols - 1) = &H8000000F
            End If
        Next intRow
        .CellBorderRange .FixedRows, .FixedCols, .Rows - 1, .Cols - 1, .GridColor, 0, 0, 1, 0, 0, 0
        For intRow = .FixedRows To .Rows - 1
            '设置竖边框
            For intCOl = 0 To (Val(Split(.RowData(intRow), ";")(0))) - 1
                .CellBorderRange intRow, .FixedCols + intCOl, intRow, .FixedCols + intCOl, .GridColor, 0, IIf(intCOl + 1 > lngPreNum And intRow > .FixedRows, 1, 0), 1, IIf(.FixedCols + intCOl < .Cols, 1, 0), 0, 0
            Next intCOl
            lngPreNum = (Val(Split(.RowData(intRow), ";")(0)))
        Next intRow
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
         
    End With

    With vsfTabDetail
        .Cols = 9
        .Rows = 2
        .FixedCols = 6
        .WordWrap = True
        .RowHeightMin = 300 + 300 * mintBigSize / 3
        .ColWidth(COL_TabNull) = 255
        .ColHidden(COL_tab字符串) = True
        .ColHidden(COL_tab项目序号) = True
        .ColHidden(COL_tab项目名) = True
        .ColHidden(col_tab原始时间) = True
        .TextMatrix(0, .FixedCols - 1) = "分类"
        .ColWidth(.FixedCols - 1) = 1500
        .TextMatrix(0, .FixedCols) = "时间"
        .ColWidth(.FixedCols) = 1600
        .TextMatrix(0, .FixedCols + 1) = "数据"
        .ColWidth(.FixedCols + 1) = 1200
        .TextMatrix(0, .FixedCols + 2) = "来源"
        .ColWidth(.FixedCols + 2) = 2500
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub




Private Function SaveData() As Boolean
    '--------------------
    '数据保存
    '--------------------
    Dim lngItemCode As Long
    Dim lng记录ID As Long, lngOld记录ID As Long
    Dim lngRow As Long
    Dim intModify As Integer, int检查科室 As Integer
    Dim i As Integer, int项目首次 As Integer
    Dim int原始显示 As Integer
    Dim strValue As String, str未记 As String
    Dim strSQL As String, strTime As String
    Dim strEnd As String, strBegin As String
    Dim strOldTime As String, strSQLShow As String
    Dim str部位 As String, strName As String, strTmp As String
    Dim strInfo As String
    Dim arrSQL() As String, arrSQLTime() As String
    Dim arrSQLShow() As String, arrTmp() As String
    Dim blnEdit As Boolean, blnSave As Boolean
    Dim blnTran As Boolean
    On Error GoTo Errhand
    
    mrsCurve.Filter = 0
    mrsCurve.Sort = "时间,项目序号"
    mrsTableDetail.Filter = 0
    mrsCurve.Sort = "时间,项目序号"
    Screen.MousePointer = 11
    ReDim Preserve arrSQL(1 To 1)
    ReDim Preserve arrSQLTime(1 To 1)
    ReDim Preserve arrSQLShow(1 To 1)
    mrsRecodeID.Filter = 0
    
    '体温曲线保存
    With mrsCurve
        Do While Not .EOF
            lngItemCode = Val(!项目序号)
            strValue = Nvl(!数值)
            mrsCurInfo.Filter = "名称='" & strValue & "'"
            intModify = Val(zlStr.Nvl(!修改))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,3,9,", Val(zlStr.Nvl(!数据来源))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            str部位 = Nvl(!部位)
            blnSave = False
            If !状态 <> 0 Then
               
                '体温曲线项目
                strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                strOldTime = Format(!原始时间, "YYYY-MM-DD hh:mm:ss")
                int检查科室 = IIf(ISCheckDept(strTime) = True, 1, 0)
                strBegin = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
                strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
                
                
                '更新显示状态
                int原始显示 = Val(Nvl(!原始显示状态))
                If int原始显示 <> !显示 Then
                    strSQLShow = "Zl_体温单数据_设置显示("
                    '发生时间_In In 病人护理数据.发生时间%Type,
                     strSQLShow = strSQLShow & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ","
                    '文件id_In   In 病人护理数据.文件id%Type,
                    strSQLShow = strSQLShow & mT_Patient.lng文件ID & ","
                    '项目序号_In In 病人护理明细.项目序号%Type,
                    strSQLShow = strSQLShow & lngItemCode & ","
                    '部位_In     In 病人护理明细.体温部位%Type,
                    strSQLShow = strSQLShow & "'" & str部位 & "',"
                    '显示_In     In 病人护理明细.显示%Type
                    strSQLShow = strSQLShow & Val(!显示) & ")"
                    
                    arrSQLShow(ReDimArray(arrSQLShow)) = strSQLShow
                End If
                
                '先修改数据发生时间
                If strOldTime <> strTime And strOldTime <> "" Then
                    mrsRecodeID.Filter = "时间='" & strOldTime & "'"
                    If mrsRecodeID.RecordCount > 0 Then
                        lng记录ID = Val(mrsRecodeID!记录ID)
                        '相同记录修改后不再修改
                        If lng记录ID <> lngOld记录ID Then
                            strSQL = "ZL_体温单数据_发生时间("
                            'ID_IN       IN 病人护理数据.ID%TYPE,
                            strSQL = strSQL & lng记录ID & ","
                            '发生时间_IN IN 病人护理数据.发生时间%TYPE
                            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                            
                            arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                        End If
                    End If
                End If
                lngOld记录ID = lng记录ID
                If strValue = "不升" And lngItemCode = gint体温 Then
                    str未记 = ""
                Else
                    str未记 = !未记说明
                End If
                If Val(!状态) <> 5 And Val(!状态) <> 6 Then '状态为5只修改了时间  状态为6只修改了显示状态
                    '更新体温曲线数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & "1,"
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & lngItemCode & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & IIf(strValue <> "", "'" & Nvl(!部位) & "'", "NULL") & ","
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = gint体温 And strValue <> "", Val(!复试合格), "0") & ","
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    strSQL = strSQL & "'" & str未记 & "',"
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & IIf(Val(!数据来源) = 0, 0, !数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用)
                    '  项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    '  开始时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的开始时间
                    '  结束时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",0,NULL,NULL,NULL,"
                    strSQL = strSQL & int检查科室 & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                   
                End If
            End If
            .MoveNext
        Loop
    End With
    lngOld记录ID = 0
    '表格项目保存，处理方式和体温曲线一样 先处理时间，在修改数据
    With mrsTableDetail
        Do While Not .EOF
            lngItemCode = Val(!项目序号)
            strValue = Nvl(!结果)
            
            mrsCurInfo.Filter = "名称='" & strValue & "'"
            If lngItemCode = 4 And zlStr.Nvl(!项目名称) = "血压" And Not mrsCurInfo.EOF Then
                strValue = Nvl(!结果) & "/" & Nvl(!结果)
            End If
            intModify = Val(zlStr.Nvl(!修改))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,3,9,", Val(zlStr.Nvl(!数据来源))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            blnSave = False
            If !状态 <> 0 Then
                int项目首次 = 0
                strName = zlStr.Nvl(!项目名称)
                strTmp = GetItemInfo(lngItemCode, strName, lngRow)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrTmp = Split(strTmp, ",")
        
                strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                strOldTime = Format(!原始时间, "YYYY-MM-DD hh:mm:ss")
                strEnd = strTime
                '先修改数据发生时间
                If strOldTime <> strTime And strOldTime <> "" Then
                    mrsRecodeID.Filter = "时间='" & strOldTime & "'"
                    If mrsRecodeID.RecordCount > 0 Then
                        lng记录ID = Val(mrsRecodeID!记录ID)
                        '相同记录修改后不再修改
                        If lng记录ID <> lngOld记录ID Then
                            strSQL = "ZL_体温单数据_发生时间("
                            'ID_IN       IN 病人护理数据.ID%TYPE,
                            strSQL = strSQL & lng记录ID & ","
                            '发生时间_IN IN 病人护理数据.发生时间%TYPE
                            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')" & ")"
                            
                            arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                        End If
                    End If
                End If
                lngOld记录ID = lng记录ID
                    
                If Val(!状态) <> 3 And Val(!状态) <> 0 Then '状态为3的是只修改了时间的
                    '对于汇总数据需要根据汇总时段删除本时段的所有数据
                    If Val(arrTmp(4)) = 4 Then
                        strTmp = GetAnimalItemTime(lngRow, !列号, 0, strInfo)
                        If strInfo <> "" Then Exit Function
                        strBegin = Split(strTmp, ";")(0)
                        strEnd = Split(strTmp, ";")(1)
                        If CDate(strTime) < CDate(mstrBTime) Then strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
                        If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
                        int项目首次 = 1
                    End If
                    
                    int检查科室 = IIf(ISCheckDept(strTime) = True, 1, 0)
                    strTime = "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & Val(Nvl(!记录类型, 1)) & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & lngItemCode & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & IIf(Val(arrTmp(5)) = 2, "'" & Nvl(!体温部位) & "'", "NULL") & ","
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = gint体温 And strValue <> "", Val(!复试合格), "0") & ","
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    If Val(arrTmp(1)) = 1 And Val(arrTmp(5)) = 2 Then
                        strSQL = strSQL & "'" & IIf(strValue = "", "", Val(!未记说明)) & "',"
                    Else
                        strSQL = strSQL & "NUll,"
                    End If
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & int项目首次 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strBegin & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
            .MoveNext
        Loop
    End With
'
     '上下标信息
    mrsNote.Filter = 0
    mrsNote.Sort = "时间"
    With mrsNote
        Do While Not .EOF
        lngItemCode = Val(!记录类型)
        
        If Val(!状态) <> 3 And Val(!状态) <> 0 Then
            strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
            strValue = zlStr.Nvl(!内容)
            int项目首次 = 1
            int检查科室 = IIf(ISCheckDept(strTime) = True, 1, 0)
            strTime = "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss')"
            strBegin = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
            strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
            
             '更新数据信息
            strSQL = "Zl_体温单数据_Update("
            '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
            strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
            '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
            strSQL = strSQL & strTime & ","
            '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
            strSQL = strSQL & lngItemCode & ","
            '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
            strSQL = strSQL & 0 & ","
            '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
            strSQL = strSQL & "'" & strValue & "',"
            '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
            strSQL = strSQL & "NULL,"
            '复试合格_In In Number := 0,
            strSQL = strSQL & "NULL,"
            '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
            strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!未记说明) & "'", "NULL") & ","
            '他人记录_In In Number := 1,
            strSQL = strSQL & "1,"
            '数据来源_In In 病人护理明细.数据来源%Type := 0,
            strSQL = strSQL & Val(!数据来源) & ","
            '来源id_In   In 病人护理明细.来源id%Type := Null,
            strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
            '共用_In     In 病人护理明细.共用%Type := 0,
            strSQL = strSQL & Val(!共用) & ","
            '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
            strSQL = strSQL & int项目首次 & ","
            '开始时间_In In 病人护理数据.发生时间%Type := Null,
            strSQL = strSQL & "To_Date('" & strBegin & "','yyyy-mm-dd hh24:mi:ss'),"
            '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
            strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
            '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
            '  检查科室_IN IN Number :=1
            strSQL = strSQL & ",NULL," & int检查科室 & ")"
            arrSQL(ReDimArray(arrSQL)) = strSQL
        End If
        .MoveNext
        Loop
    End With
    
    gcnOracle.BeginTrans
    blnTran = True
    '先执行时间变化
    For i = 1 To UBound(arrSQLTime)
        If arrSQLTime(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLTime(i)), "保存时间数据"):
    Next
    '在执行数据变化
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存体温数据"):
    Next
    
    '最后执行显示变化
     For i = 1 To UBound(arrSQLShow)
        If arrSQLShow(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLShow(i)), "保存显示修改"):
    Next
    gcnOracle.CommitTrans
    blnTran = False
    mblnOK = True
    SaveData = True
    Screen.MousePointer = 0
     
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call SaveErrLog
    
End Function

Private Function ISCheckDept(ByVal str发生时间 As String) As Boolean
    '------------------------------------------------
    '功能：是否在Zl_体温单数据_Update中进行科室检查
    'mstrOverDate<=mstrETime 并且病人已经出院，肯定是病人出院时间和入院时间在一列（程序处理后的结果
    '------------------------------------------------
    If mbln出院 = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str发生时间, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str发生时间, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
            ISCheckDept = False
        Else
            ISCheckDept = True
        End If
    Else
        ISCheckDept = True
    End If
End Function


Private Function GetItemInfo(ByVal lngItemNO As Long, ByVal strName As String, ByRef lngRow As Long) As String
'---------------------------------------------------------------
'功能:获取项目信息
'---------------------------------------------------------------
    Dim intRow As Integer
    Dim strValue As String
    
    On Error GoTo Errhand
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        If Val(vsfTab.TextMatrix(intRow, COL_tab项目序号)) = lngItemNO And vsfTab.TextMatrix(intRow, COL_tab项目名) = strName And intRow Mod 2 <> 1 Then
            Exit For
        End If
    Next intRow
    
    If intRow >= vsfTab.Rows Then
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If Val(vsfTab.TextMatrix(intRow, COL_tab项目序号)) = lngItemNO Then
                Exit For
            End If
        Next intRow
    End If
    
    If intRow < vsfTab.Rows Then
        strValue = vsfTab.TextMatrix(intRow, COL_tab字符串)
    End If
    lngRow = intRow
    GetItemInfo = strValue
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function zlRefreshData(Optional ByVal blnCurve As Boolean = True, Optional ByVal blnTab As Boolean = True) As Boolean
    '--------------------------------------------------------------------------------
    '功能:提取一段时间内的所有体温数据
    '参数 blnCurve是否刷新体温数据
    '--------------------------------------------------------------------------------
    Dim strFields As String, strValues As String, strPara As String
    Dim strSQL As String
    Dim strTime As String
    Dim strCenterTime As String '中间时间
    Dim strBTime As String '当前一天的时间
    Dim strETime As String
    Dim dtBegin As String, dtEnd As String
    Dim strItems As String
    Dim strName As String
    Dim strItemName As String '项目名字符串
    Dim str项目名称 As String '一个项目名
    Dim strPart As String
    Dim strParam As String
    Dim int标记 As Integer, intModify As Integer
    Dim int数据来源 As Integer
    Dim intRow As Integer, intNum As Integer
    Dim lng项目序号 As Long, int序号 As Integer
    Dim blnAdd As Boolean '是否添加
    Dim rsTemp As New ADODB.Recordset   '查询数据集
    Dim rsCurve As New ADODB.Recordset '临时记录集
    Dim rstab As New ADODB.Recordset  '临时数据集
    
    On Err GoTo Errhand
    If blnCurve = False And blnTab = False Then Exit Function
    
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "～" & Format(mstrEnd, "HH:mm")
    
    '初始化记录集
    gstrFields = "记录ID," & adDouble & ",18|时间," & adLongVarChar & ",20"
    Call Record_Init(mrsRecodeID, gstrFields)
    
    gstrFields = "序号," & adDouble & ",18|分组名," & adLongVarChar & ",40|数值," & adLongVarChar & ",400|部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复试合格," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|原始显示状态," & adDouble & ",1|" & _
         "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
    Call Record_Init(rsCurve, gstrFields)
    Call Record_Init(mrsCurve, gstrFields)
    Call Record_Init(mrsTable, gstrFields)
    gstrFields = "ID," & adDouble & ",18|分组名," & adLongVarChar & ",40|结果," & adLongVarChar & ",400|体温部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复试合格," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|原始显示状态," & adDouble & ",1|" & _
         "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
    Call Record_Init(mrsTableDetail, gstrFields)
    gstrFields = "序号|分组名|数值|部位|标记|时间|原始时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|原始显示状态|来源ID|共用|状态|列号|记录类型"
    '刷新体温曲线数据和上下标信息
    If blnCurve Then
        strBTime = dtpDate.Value & " 00:00:00"
        strETime = dtpDate.Value & " 23:59:59"
        strSQL = _
            " SELECT /*+ RULE */ C.ID 序号,C.记录ID,A.发生时间 As 时间,'1)体温曲线项目' 分组名,C.显示,c.记录内容 As 数值,c.体温部位,c.复试合格,D.记录名,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明,C.数据来源,C.来源ID,C.共用" & vbNewLine & _
            "                    FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E,Table(Cast(f_num2list([7]) As zlTools.t_Numlist)) F" & vbNewLine & _
            "                    Where B.ID=A.文件ID" & vbNewLine & _
            "                        AND A.ID = C.记录ID" & vbNewLine & _
            "                        AND B.ID=[1]" & vbNewLine & _
            "                        AND Nvl(B.婴儿,0)=[4]" & vbNewLine & _
            "                        AND B.病人id=[2]" & vbNewLine & _
            "                        AND B.主页id=[3]" & vbNewLine & _
            "                        AND D.项目序号=C.项目序号" & vbNewLine & _
            "                        AND C.记录类型=1" & vbNewLine & _
            "                        AND E.项目序号=D.项目序号" & vbNewLine & _
            "                        AND E.项目序号=F.COLUMN_VALUE" & vbNewLine & _
            "                        AND (NVL(D.记录法,1)<>2 OR (NVL(D.记录法,1)=2 And D.项目序号=3))" & _
            "                        And A.发生时间 BETWEEN [5] And [6] And C.终止版本 Is Null" & vbNewLine & _
            "                    Order By A.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
            If mblnMove Then
                mstrSQL = Replace(strSQL, "病人护理文件", "H病人护理文件")
                mstrSQL = Replace(strSQL, "病人护理数据", "H病人护理数据")
                mstrSQL = Replace(strSQL, "病人护理明细", "H病人护理明细")
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, _
                 CDate(Format(strBTime, "YYYY-MM-DD HH:mm:ss")), CDate(Format(strETime, "YYYY-MM-DD HH:mm:ss")), mstrCurveItem)

    
        With rsTemp
            Do While Not .EOF
                '添加记录集
                Call Record_Update(mrsRecodeID, "记录ID|时间", Val(Nvl(!记录ID)) & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss"), "记录ID|" & Val(Nvl(!记录ID)))
                
                intModify = 0
                If strTime = "" Then strTime = Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss")
                lng项目序号 = zlStr.Nvl(!项目序号)
                Select Case lng项目序号
                    Case gint心率
                        int标记 = 1
                    Case Else
                        int标记 = Val(Nvl(!记录标记))
                End Select
                intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!数据来源)) & ",") = 0, 1, 0)
                blnAdd = True
                '心率和脉搏公用时，检查脉搏对应的时间是否存在心率
                If mint心率应用 = 2 And lng项目序号 = -1 Then
                    mrsCurve.Filter = "项目序号=2 and 时间='" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "'"
                    If mrsCurve.RecordCount > 0 Then
                        strPara = "序号|" & mrsCurve("序号")
                        strFields = "数值|标记|修改"
                        
                        If InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(mrsCurve!数据来源)) & ",") = 0 And InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!数据来源)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        '脉搏短轴时心率为未记说明只显示脉搏，脉搏为未记说明时就显示未记说明
                        If UBound(Split(mrsCurve("数值"), "/")) <> -1 Then
                            If IsNumeric(zlStr.Nvl(!数值)) Then
                                If mbln脉搏共用显示 Then
                                    gstrValues = zlStr.Nvl(!数值) & "/" & Split(mrsCurve("数值"), "/")(0) & "|" & int标记 & "|" & intModify
                                Else
                                    gstrValues = Split(mrsCurve("数值"), "/")(0) & "/" & zlStr.Nvl(!数值) & "|" & int标记 & "|" & intModify
                                End If
                            Else
                                gstrValues = Split(mrsCurve("数值"), "/")(0) & "|" & int标记 & "|0"
                            End If
                        Else
                            gstrValues = mrsCurve("数值") & "|1|0"
                        End If
                        
                        Call Record_Update(mrsCurve, strFields, gstrValues, strPara)
                        blnAdd = False
                    Else
                        lng项目序号 = 2
                    End If
                End If
                
                '处理物理降温、疼痛减痛
                If (lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度) And int标记 = 1 Then
                    mrsCurve.Filter = "状态<> 3 and  项目序号=" & lng项目序号 & " and 时间='" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "' and 标记<>1"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(mrsCurve!数据来源)) & ",") = 0 And InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!数据来源)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        strPara = "序号|" & mrsCurve("序号")
                        strFields = "数值|标记|修改"
                        gstrValues = Split(mrsCurve("数值"), "/")(0) & "/" & zlStr.Nvl(!数值) & "|" & int标记 & "|" & intModify
                        Call Record_Update(mrsCurve, strFields, gstrValues, strPara)
                    End If
                    blnAdd = False
                End If
                
                If blnAdd Then
                    '进行曲线显示处理
                    strPart = GetPart(lng项目序号)
                    int数据来源 = Val(zlStr.Nvl(!数据来源, 0))
                    If Trim(Replace(zlStr.Nvl(!数值), "/", "")) = "" Then
                        int数据来源 = 0
                    End If
                    gstrValues = zlStr.Nvl(!序号) & "|" & zlStr.Nvl(!分组名) & "|" & Trim(Replace(zlStr.Nvl(!数值), "/", "")) & "|" & _
                        zlStr.Nvl(!体温部位, strPart) & "|" & int标记 & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & _
                        Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & zlStr.Nvl(!记录名) & "|" & Val(zlStr.Nvl(!复试合格, 0)) & "|" & _
                        zlStr.Nvl(!未记说明) & "|" & int数据来源 & "|" & intModify & "|" & Val(zlStr.Nvl(!显示, 0)) & "|" & Val(zlStr.Nvl(!显示, 0)) & "|" & Val(zlStr.Nvl(!来源ID, 0)) & "|" & Val(zlStr.Nvl(!共用, 0)) & "|0|0|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            .MoveNext
            Loop
        End With
        
        Call ShowCurve
        
        gstrFields = "序号," & adDouble & ",18|项目序号," & adDouble & ",18|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|记录类型," & adDouble & ",1|内容," & _
                adLongVarChar & ",100|项目名称," & adLongVarChar & ",20|未记说明," & adLongVarChar & ",20|记录组号," & adDouble & ",1|数据来源," & adDouble & ",1|显示," & adDouble & ",1|" & _
                "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1"
        Call Record_Init(mrsNote, gstrFields)
        gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
        
        mstrSQL = "" & _
             " Select C.ID 序号, B.发生时间 AS 时间,C.记录类型,C.项目序号,C.未记说明,C.记录内容,C.记录组号,C.项目名称,C.数据来源,C.显示,C.来源ID,C.共用" & _
             " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
             " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]  AND Nvl(A.婴儿, 0)=[4] AND a.病人id=[2] AND a.主页id=[3] And c.终止版本 Is Null" & _
             " AND c.记录类型 in (2,6)  AND B.发生时间 BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "读取上下标等信息", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, _
            mT_Patient.lng婴儿, CDate(strBTime), CDate(strETime))
        With rsTemp
            Do While Not .EOF
                gstrValues = zlStr.Nvl(!序号) & "|" & zlStr.Nvl(!项目序号, 0) & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & zlStr.Nvl(!记录类型) & "|" & _
                    zlStr.Nvl(!记录内容) & "|" & zlStr.Nvl(!项目名称) & "|" & Nvl(!未记说明) & "|" & zlStr.Nvl(!记录组号, 0) & "|" & Val(zlStr.Nvl(!数据来源, 0)) & "|" & _
                    Val(zlStr.Nvl(!显示, 0)) & "|" & Val(zlStr.Nvl(!来源ID, 0)) & "|" & Val(zlStr.Nvl(!共用, 0)) & "|0"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '添加上下标信息
        Call ShowTabUpDown
    End If
        
        '提取表格数据
    If blnTab Then
        strItems = ""
        If vsfTab.Tag <> "NO" Then
            For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
                If lng项目序号 <> 4 Then
                    strItemName = vsfTab.TextMatrix(intRow, COL_tab项目名)
                    If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                        strItems = strItems & ",'" & strItemName & "'"
                    End If
                End If
            Next
            If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
            strItems = strItems & ",'收缩压','舒张压'"
            If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
            '提取一天内(可能含有第二天数据)所有的表格数据信息
            mstrSQL = "SELECT C.Id,a.发生时间 As 时间,C.记录ID,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & vbNewLine & _
                "  DECODE(E.项目性质,2,C.体温部位 || D.记录名,D.记录名) 项目名称,D.项目序号,C.来源ID,C.共用,E.项目性质 " & _
                "  FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
                "  Where B.ID = A.文件ID" & vbNewLine & _
                "  AND A.ID = C.记录ID" & vbNewLine & _
                "  AND B.ID = [1]" & vbNewLine & _
                "  AND Nvl(B.婴儿, 0) = [7]" & vbNewLine & _
                "  AND B.病人id = [2]" & vbNewLine & _
                "  AND B.主页id = [3]" & vbNewLine & _
                "  AND INSTR([6], DECODE(E.项目性质, 2,C.体温部位 || D.记录名, D.记录名)) > 0" & vbNewLine & _
                "  AND D.项目序号 = C.项目序号" & vbNewLine & _
                "  AND Mod(c.记录类型,10) = 1" & vbNewLine & _
                "  AND E.项目序号 = D.项目序号" & vbNewLine & _
                "  AND A.发生时间 BETWEEN [4] And [5]" & vbNewLine & _
                "  And C.终止版本 Is Null" & vbNewLine & _
                "  AND D.记录法 = 2 And D.项目序号<>3" & vbNewLine & _
                "  UNION ALL "
            '提取非体温表格的汇总项目（体温表格汇总项目子项可能存在非体温项目）
            mstrSQL = mstrSQL & vbNewLine & _
                "  SELECT C.ID,a.发生时间 As 时间,C.记录ID,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,NVL(C.数据来源,0) 数据来源," & _
                "   Decode(d.项目性质, 2, c.体温部位 || d.项目名称, d.项目名称) 项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
                "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, A.项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
                "       WHERE A.项目序号=B.序号 AND B.父序号 Is Not Null  " & vbNewLine & _
                "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
                "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
                "   Where B.ID=A.文件ID And A.ID = C.记录ID AND Instr([6], Decode(d.项目性质, 2, c.体温部位 || d.项目名称, d.项目名称)) = 0  AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
                "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
                "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"
                
            mstrSQL = _
                "   Select ID,时间,记录ID,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & mstrSQL & ")" & _
                "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
            If mblnMove Then
                mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
                mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
                mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
            End If
            
            strTime = CDate(Format(dtpDate.Value, "YYYY-MM-DD") & " 23:59:59")
            dtBegin = Int(CDate(dtpDate.Value) - 1)
            dtEnd = CDate(CDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) + 1)
            If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")) Then _
                dtBegin = CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss"))
            If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss")) Then _
                dtEnd = CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss"))
            
            Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, _
                                               mT_Patient.lng文件ID, _
                                               mT_Patient.lng病人ID, _
                                               mT_Patient.lng主页ID, _
                                               CDate(dtBegin), _
                                               CDate(dtEnd), _
                                               strItems, mT_Patient.lng婴儿, mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID)
            gstrFields = "Id|分组名|结果|体温部位|标记|时间|原始时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
            
            '明细数据集加载
            With rsTemp
                .Sort = "时间,项目序号,id"
                Do While Not .EOF
                    '添加记录集
                    Call Record_Update(mrsRecodeID, "记录ID|时间", Val(Nvl(!记录ID)) & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss"), "记录ID|" & Val(Nvl(!记录ID)))
                    blnAdd = False
                    intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!数据来源)) & ",") = 0, 1, 0)
                    int序号 = 0
                    If zlStr.Nvl(!Id) <> intNum Or zlStr.Nvl(!项目名称) <> strName Then
                        intNum = zlStr.Nvl(!项目序号)
                        strName = zlStr.Nvl(!项目名称)
                        '收缩压/舒张压
                        If intNum = 4 Or intNum = 5 Then
                            Select Case zlStr.Nvl(!项目名称)
                                Case "收缩压"
                                    strParam = ""
                                    strParam = zlStr.Nvl(!结果)
                                Case "舒张压"
                                    If InStr(strParam, "/") > 0 Then
                                        strParam = strParam & zlStr.Nvl(!结果)
                                    Else
                                        strParam = strParam & "/" & zlStr.Nvl(!结果)
                                    End If
                                    mrsCurInfo.Filter = "名称='" & Nvl(!结果) & "'"
                                    If Not mrsCurInfo.EOF Then
                                        strParam = zlStr.Nvl(!结果)
                                    End If
                                    If strParam = "/" Then strParam = ""
                                    blnAdd = True
                                    intNum = 4
                                    strName = "血压"
                            End Select
                        Else
                            strParam = zlStr.Nvl(!结果)
                            blnAdd = True
                        End If
                        
                        If !记录类型 = 11 Then blnAdd = False
                        
                        If blnAdd = True Then
                            gstrValues = zlStr.Nvl(!Id) & "|2)体温表格项目|" & strParam & "|" & _
                                            zlStr.Nvl(!体温部位) & "|0|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                            Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & intNum & "|" & strName & "|0|" & _
                                            zlStr.Nvl(!未记说明) & "|" & Val(zlStr.Nvl(!数据来源, 0)) & "|" & intModify & "|" & Val(zlStr.Nvl(!显示, 0)) & "|" & _
                                            Val(zlStr.Nvl(!来源ID, 0)) & "|" & Val(zlStr.Nvl(!共用, 0)) & "|0|" & int序号 & "|1"

                            Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            
            gbln出院 = mbln出院
            Call InitTableData(rsTemp)
            Call ShowTable
        End If
    End If
        
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitTableData(ByVal rsTemp As ADODB.Recordset)
    '---------------------------------------------------------
    '功能:初始化体温表格数据
    '---------------------------------------------------------
    Dim int项目性质 As Integer, int记录频次 As Integer, int项目表示 As Integer, int入院首测 As Integer
    Dim int序号 As Integer, intNum As Integer
    Dim intRow As Integer, intModify As Integer
    Dim lng项目序号 As Long
    Dim blnAdd As Boolean
    Dim strPart As String '部位
    Dim strParam As String, strFields As String, strValues As String
    Dim str项目名称 As String, strName As String
    Dim rstab As New ADODB.Recordset
    
    On Error GoTo Errhand
    gstrFields = "序号," & adDouble & ",18|分组名," & adLongVarChar & ",40|数值," & adLongVarChar & ",400|部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复试合格," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|原始显示状态," & adDouble & ",1|" & _
         "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
    Call Record_Init(mrsTable, gstrFields)
    strFields = "序号|分组名|数值|部位|标记|时间|原始时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        int项目性质 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5))
        int记录频次 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(3))
        int项目表示 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(4))
        strPart = Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(7)
        int入院首测 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(8))
        lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
        str项目名称 = vsfTab.TextMatrix(intRow, COL_tab项目名)
    
        intNum = 0
        strName = ""
        Set rstab = ReturnItemRecord(rsTemp, Int(CDate(Format(dtpDate.Value, "YYYY-MM-DD hh:mm:ss"))), CDate(mstrBTime), lng项目序号 & ";" & str项目名称 & ";" & _
                       int记录频次 & ";" & int项目表示 & ";" & int项目性质 & ";" & int入院首测 & ";" & strPart, mbln汇总当天, mbln录入小时, True)
        If rstab.RecordCount > 0 Then rstab.MoveFirst
        rstab.Sort = "时间,项目序号,序号"
        
        With rstab
            Do While Not .EOF
                blnAdd = False
                intModify = IIf(InStr(1, ",0,3,9,", "," & Val(zlStr.Nvl(!数据来源)) & ",") = 0, 1, 0)
                If zlStr.Nvl(!序号) <> intNum Or zlStr.Nvl(!项目名称) <> strName Then
                    intNum = zlStr.Nvl(!项目序号)
                    strName = zlStr.Nvl(!项目名称)
                    '收缩压/舒张压
                    If lng项目序号 = 4 And str项目名称 = "血压" Then
                        Select Case zlStr.Nvl(!项目名称)
                            Case "收缩压"
                                strParam = ""
                                strParam = zlStr.Nvl(!记录内容)
                            Case "舒张压"
                                If InStr(strParam, "/") > 0 Then
                                    strParam = strParam & zlStr.Nvl(!记录内容)
                                Else
                                    strParam = strParam & "/" & zlStr.Nvl(!记录内容)
                                End If
                                '血压显示文字
                                mrsCurInfo.Filter = "名称='" & Nvl(!记录内容) & "'"
                                If Not mrsCurInfo.EOF Then
                                    strParam = zlStr.Nvl(!记录内容)
                                End If
                                If strParam = "/" Then strParam = ""
                                blnAdd = True
                            Case "血压"
                                strParam = zlStr.Nvl(!记录内容)
                                blnAdd = True
                        End Select
                    Else
                        strParam = zlStr.Nvl(!记录内容)
                        blnAdd = True
                    End If
                    If blnAdd = True Then
                        '提取数据时是根据时间段和显示顺序排序的。如果一个时间段有多条数据,只提取前一条
                        mrsTable.Filter = "项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' and 列号=" & Val(zlStr.Nvl(!序号, 0))
                        If mrsTable.RecordCount = 0 Then
                            strValues = zlStr.Nvl(!Id) & "|2)体温表格项目|" & strParam & "|" & _
                                    zlStr.Nvl(!体温部位) & "|0|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                    Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & _
                                    zlStr.Nvl(!未记说明) & "|" & Val(zlStr.Nvl(!数据来源, 0)) & "|" & intModify & "|" & Val(zlStr.Nvl(!显示, 0)) & "|" & _
                                    Val(zlStr.Nvl(!来源ID, 0)) & "|" & Val(zlStr.Nvl(!共用, 0)) & "|0|" & zlStr.Nvl(!序号, 0) & "|1"
                            Call Record_Add(mrsTable, strFields, strValues)
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End With
    Next intRow
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ShowCurve()
'-----------------------------------------
'功能：展示曲线主表数据
'-----------------------------------------
    Dim intRow As Integer
    Dim strCenterTime As String
    Dim strFields As String, strValues As String, strPara As String
    Dim lngColor As Long
    Dim lng项目序号 As Long
    Dim rsCompara As New ADODB.Recordset
    
    On Err GoTo Errhand
    strFields = "项目序号," & adDouble & ",18|时间," & adLongVarChar & ",20|显示," & adDouble & ",1"
    Call Record_Init(rsCompara, strFields)
    
    With mrsCurve
        .Filter = "状态<> 3 and 时间 >= '" & mstrBegin & "' and 时间 <=  '" & mstrEnd & "'"
        strCenterTime = GetCenterTime(mstrBegin, mstrEnd)
        Do While Not .EOF
            lng项目序号 = !项目序号
            rsCompara.Filter = "项目序号=" & lng项目序号
            strFields = "项目序号|时间|显示"
            If rsCompara.RecordCount > 0 Then
                If !显示 = 1 Then
                    If rsCompara!显示 = 1 Then
                        If CheckShow(!时间, rsCompara!时间, strCenterTime) Then
                            strValues = zlStr.Nvl(!项目序号) & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!显示, 0))
                            strPara = "项目序号|" & lng项目序号
                            rsCompara.Filter = 0
                            Call Record_Update(rsCompara, strFields, strValues, strPara)
                        End If
                    Else
                        strValues = zlStr.Nvl(!项目序号) & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!显示, 0))
                        strPara = "项目序号|" & lng项目序号
                        rsCompara.Filter = 0
                        Call Record_Update(rsCompara, strFields, strValues, strPara)
                    End If
                Else
                    If rsCompara!显示 = 0 Then
                        If CheckShow(!时间, rsCompara!时间, strCenterTime) Then
                            strValues = lng项目序号 & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!显示, 0))
                            strPara = "项目序号|" & lng项目序号
                            rsCompara.Filter = 0
                            Call Record_Update(rsCompara, strFields, strValues, strPara)
                        End If
                    End If
                End If
                
            Else
                strValues = lng项目序号 & "|" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Val(zlStr.Nvl(!显示, 0))
                Call Record_Add(rsCompara, strFields, strValues)
            End If
            !显示 = 0
            .Update
            .MoveNext
        Loop
    End With
    
     With rsCompara
        .Filter = 0
        Do While Not .EOF
            mrsCurve.Filter = "状态<> 3  and 项目序号=" & !项目序号 & " and 时间 ='" & Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "'"
            mrsCurve!显示 = 1
            mrsCurve.Update
            .MoveNext
        Loop
    End With
    
    '显示体温数据
    mrsCurve.Filter = "状态<> 3 and 时间 >= '" & mstrBegin & "' and 时间 <=  '" & mstrEnd & "'"
    mrsCurve.Sort = "时间"
    
    vsfCurve.Cell(flexcpText, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = ""
    vsfCurve.Cell(flexcpForeColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000012
    vsfCurve.Cell(flexcpBackColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000005
    
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1

        vsfCurve.Body.MergeRow(intRow) = True
        vsfCurve.TextMatrix(intRow, COL_数据) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", "", "") & Space(intRow)
        vsfCurve.TextMatrix(intRow, COL_颜色) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", " ", Space(intRow))
        vsfCurve.TextMatrix(intRow, col_原始时间) = ""
        vsfCurve.TextMatrix(intRow, COL_显示) = ""
        vsfCurve.TextMatrix(intRow, COL_编辑) = "0"
        vsfCurve.TextMatrix(intRow, COL_来源) = "0"
        vsfCurve.TextMatrix(intRow, COL_原值) = ""
        vsfCurve.TextMatrix(intRow, COL_修改状态) = "0"
        If vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明" Then
             vsfCurve.Cell(flexcpBackColor, intRow, COL_颜色, intRow, COL_颜色) = RGB(0, 0, 255)
        End If
    Next intRow
    
    With mrsCurve
        .Filter = "状态<> 3 and 显示=1 and 时间 >= '" & mstrBegin & "' and 时间 <=  '" & mstrEnd & "'"
        Do While Not .EOF
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    lng项目序号 = Val(vsfCurve.TextMatrix(intRow, COL_项目序号))
                    If !分组名 = vsfCurve.TextMatrix(intRow, COL_分组名) And !项目序号 = lng项目序号 Then
                        vsfCurve.TextMatrix(intRow, COL_修改状态) = zlStr.Nvl(!状态, 0)
                        vsfCurve.TextMatrix(intRow, col_原始时间) = Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss")
                        vsfCurve.TextMatrix(intRow, COL_显示) = IIf(Val(zlStr.Nvl(!显示)) = 1, "√", "")
                        vsfCurve.TextMatrix(intRow, COL_时间) = Format(zlStr.Nvl(!时间), "HH:mm")
                        vsfCurve.TextMatrix(intRow, COL_数据) = Space(intRow) & zlStr.Nvl(!数值) & Space(intRow)
                        vsfCurve.TextMatrix(intRow, COL_颜色) = vsfCurve.TextMatrix(intRow, COL_数据)
                        If Not IsNumeric(zlStr.Nvl(!数值)) And zlStr.Nvl(!数值) <> "不升" And InStr(1, zlStr.Nvl(!数值), "/") = 0 Then
                            vsfCurve.TextMatrix(intRow, COL_部位) = ""
                            vsfCurve.TextMatrix(intRow, Col_未记说明) = zlStr.Nvl(!未记说明)
                        Else
                            vsfCurve.TextMatrix(intRow, COL_部位) = zlStr.Nvl(!部位)
                            vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
                        End If
                        If lng项目序号 = gint体温 And (IsNumeric(zlStr.Nvl(!数值)) Or zlStr.Nvl(!数值) <> "不升") Then
                            vsfCurve.TextMatrix(intRow, COL_复试合格) = IIf(Val(zlStr.Nvl(!复试合格)) = 1, "√", "")
                        End If
                        lngColor = 255
                        If InStr(1, ",0,3,9,", Val(zlStr.Nvl(!数据来源))) = 0 Then
                            If zlStr.Nvl(!数值) = "不升" And lng项目序号 = gint体温 Then
                                lngColor = 255
                            ElseIf lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                                If InStr(1, zlStr.Nvl(!数值), "/") = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    If Val(!修改) = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        lngColor = 255
                                    End If
                                End If
                            End If
                            vsfCurve.Cell(flexcpForeColor, intRow, COL_数据, intRow, COL_数据) = lngColor
                        Else
                            vsfCurve.Cell(flexcpForeColor, intRow, COL_数据, intRow, COL_数据) = &H80000012
                        End If
                        vsfCurve.TextMatrix(intRow, COL_来源) = Val(CStr(zlStr.Nvl(!数据来源)))
                        vsfCurve.TextMatrix(intRow, COL_原值) = Val(!数值)
                        If lng项目序号 = 2 And mbln脉搏共用显示 And InStr(!数值, "/") > 0 Then
                            vsfCurve.TextMatrix(intRow, COL_原值) = Split(!数值, "/")(1)
                        End If
                        vsfCurve.TextMatrix(intRow, COL_编辑) = Val(zlStr.Nvl(!修改, 0))
                        
                    End If
                Next intRow
                .MoveNext
            Loop
        End With
        
         Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowTabUpDown()
    '--------------------
    '功能：曲线上下标
    '--------------------
    Dim intRow As Integer
    
    On Error GoTo Errhand
    mrsNote.Filter = "状态<> 3 and 时间 >= '" & mstrBegin & "' and 时间 <=  '" & mstrEnd & "'"
    With mrsNote
        Do While Not .EOF
                If CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")) _
                    And CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) <= CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")) Then
                    Select Case Val(!记录类型)
                        Case 2
                            intRow = mOptRow.上标
                        Case 6
                            intRow = mOptRow.下标
                    End Select
                    vsfCurve.TextMatrix(intRow, COL_修改状态) = zlStr.Nvl(!状态, 0)
                    vsfCurve.TextMatrix(intRow, col_原始时间) = Format(zlStr.Nvl(!时间), "YYYY-MM-DD HH:mm:ss")
                    vsfCurve.TextMatrix(intRow, COL_时间) = Format(zlStr.Nvl(!时间), "hh:mm")
                    vsfCurve.TextMatrix(intRow, COL_数据) = Space(intRow) & zlStr.Nvl(!内容) & Space(intRow)
                    vsfCurve.Cell(flexcpBackColor, intRow, COL_颜色, intRow, COL_颜色) = IIf(IsNumeric(Nvl(!未记说明)) = False, 16711680, Val(Nvl(!未记说明)))
                    vsfCurve.TextMatrix(intRow, COL_部位) = ""
                    vsfCurve.TextMatrix(intRow, COL_复试合格) = ""
                    vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
                End If
        .MoveNext
        Loop
    End With
        
        Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowDetail(ByVal lng项目序号 As Long, ByVal intNewRow As Integer)
    '--------------------------------------------------
    '功能：展示曲线详细表格数据
    '--------------------------------------------------
    Dim intRow As Integer
    Dim intMarkRow As Integer
    Dim str字符串 As String
    Dim lngColor As Long
    
    On Err GoTo Errhand
    vsfDetail.Rows = vsfDetail.FixedRows
    vsfDetail.Rows = vsfDetail.FixedRows + 1
    If mrsCurve.State = adStateClosed Then Exit Function
    mrsCurve.Filter = "状态<> 3 and 时间 > '" & mstrBegin & "' and 时间 <=  '" & mstrEnd & "' and 项目序号=" & lng项目序号
    mrsCurve.Sort = "时间"
    If mblnInit Then
        vsfDetail.ColHidden(COL_复试合格) = False
        If lng项目序号 <> gint体温 Then vsfDetail.Body.ColHidden(COL_复试合格) = True
    End If
    If mrsCurve.RecordCount > 0 Then
        str字符串 = vsfCurve.TextMatrix(intNewRow, COL_字符串)
        intMarkRow = 0
        With mrsCurve
            intRow = vsfDetail.FixedRows
            Do While Not .EOF
                vsfDetail.TextMatrix(intRow, COL_字符串) = str字符串
                vsfDetail.TextMatrix(intRow, COL_项目序号) = lng项目序号
                vsfDetail.TextMatrix(intRow, COL_修改状态) = !状态
                vsfDetail.TextMatrix(intRow, COL_显示) = IIf(zlStr.Nvl(!显示) = 1, "√", "")
                vsfDetail.TextMatrix(intRow, COL_时间) = Format(zlStr.Nvl(!时间), "hh:mm")
                vsfDetail.TextMatrix(intRow, col_原始时间) = Format(zlStr.Nvl(!时间), "YYYY-MM-DD hh:mm:ss")
                vsfDetail.TextMatrix(intRow, COL_数据) = zlStr.Nvl(!数值)
                If Not IsNumeric(zlStr.Nvl(!数值)) And zlStr.Nvl(!数值) <> "不升" And InStr(1, zlStr.Nvl(!数值), "/") = 0 Then
                    vsfDetail.TextMatrix(intRow, COL_部位) = ""
                    vsfDetail.TextMatrix(intRow, Col_未记说明) = zlStr.Nvl(!未记说明)
                Else
                    vsfDetail.TextMatrix(intRow, COL_部位) = zlStr.Nvl(!部位)
                    vsfDetail.TextMatrix(intRow, Col_未记说明) = ""
                End If
                If lng项目序号 = gint体温 And (IsNumeric(zlStr.Nvl(!数值)) Or zlStr.Nvl(!数值) <> "不升") Then
                    vsfDetail.TextMatrix(intRow, COL_复试合格) = IIf(Val(zlStr.Nvl(!复试合格)) = 1, "√", "")
                End If
                lngColor = 255
                        If InStr(1, ",0,3,9,", Val(zlStr.Nvl(!数据来源))) = 0 Then
                            If zlStr.Nvl(!数值) = "不升" And lng项目序号 = gint体温 Then
                                lngColor = 255
                            ElseIf lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                                If InStr(1, zlStr.Nvl(!数值), "/") = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    If Val(!修改) = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        lngColor = 255
                                    End If
                                End If
                            End If
                            vsfDetail.Cell(flexcpForeColor, intRow, COL_数据, intRow, COL_数据) = lngColor
                        Else
                            vsfDetail.Cell(flexcpForeColor, intRow, COL_数据, intRow, COL_数据) = &H80000012
                        End If
                
                Select Case !数据来源
                    Case 0, 9
                        vsfDetail.TextMatrix(intRow, COL_数据来源) = "体温单录入"
                        vsfDetail.TextMatrix(intRow, COL_编辑) = 1
                    Case 1
                        vsfDetail.TextMatrix(intRow, COL_数据来源) = "记录单同步"
                        vsfDetail.TextMatrix(intRow, COL_编辑) = 0
                    Case 3
                        vsfDetail.TextMatrix(intRow, COL_数据来源) = "移动设备录入"
                        vsfDetail.TextMatrix(intRow, COL_编辑) = 1
                    Case Else
                        vsfDetail.TextMatrix(intRow, COL_数据来源) = "其他设备同步"
                        vsfDetail.TextMatrix(intRow, COL_编辑) = 0
                End Select
                
                vsfDetail.TextMatrix(intRow, COL_来源) = Val(CStr(zlStr.Nvl(!数据来源)))
                vsfDetail.TextMatrix(intRow, COL_编辑) = Val(zlStr.Nvl(!修改, 0))
                vsfDetail.TextMatrix(intRow, COL_原值) = Val(!数值)
                If lng项目序号 = gint脉搏 And mbln脉搏共用显示 And InStr(!数值, "/") > 0 Then
                    vsfDetail.TextMatrix(intRow, COL_原值) = Split(!数值, "/")(1)
                End If
                If vsfDetail.TextMatrix(intRow, COL_显示) = "√" Then intMarkRow = intRow
                .MoveNext
                vsfDetail.Body.CellAlignment = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_显示, intRow, COL_显示) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_时间, intRow, COL_时间) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_数据, intRow, COL_数据) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_复试合格, intRow, COL_复试合格) = flexAlignCenterCenter
                vsfDetail.Body.Cell(flexcpAlignment, intRow, COL_数据来源, intRow, COL_数据来源) = flexAlignCenterCenter
                intRow = intRow + 1
                vsfDetail.Rows = vsfDetail.Rows + 1
            Loop

        End With
    End If
        If intMarkRow <> 0 Then
            vsfDetail.Row = intMarkRow
            vsfDetail.Col = COL_数据
        End If
     
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowTable()
    '--------------------------
    '功能：展示体温表格主表
    '--------------------------
    Dim strTime As String
    Dim strInfo As String
    Dim intRow As Integer, int频次 As Integer
    Dim blnAllow As Boolean, bln汇总 As Boolean
    Dim lngHour As String
    Dim arrOldTime() As String
    
    On Error GoTo Errhand
    mrsTable.Filter = 0
    mrsTable.Sort = "项目序号,列号,记录类型 "
    vsfTab.Cell(flexcpText, vsfTab.FixedRows, vsfTab.FixedCols, vsfTab.Rows - 1, vsfTab.Cols - 1) = ""
    strTime = ""
    With mrsTable
        Do While Not .EOF
            For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                blnAllow = False
                If vsfTab.TextMatrix(intRow, COL_tab项目序号) = !项目序号 Then
                    If Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 Then
                        If vsfTab.TextMatrix(intRow, COL_tab项目名) <> !项目名称 Then
                            blnAllow = False
                        Else
                            blnAllow = True
                        End If
                    Else
                        blnAllow = True
                    End If
                End If
                int频次 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(3))
                bln汇总 = Split(vsfTab.RowData(intRow), ";")(1) = 3
                If blnAllow = True Then
                    If InStr(vsfTab.TextMatrix(intRow, col_tab原始时间), "'") > 0 Then
                        arrOldTime = Split(vsfTab.TextMatrix(intRow, col_tab原始时间), "'")
                    Else
                        ReDim Preserve arrOldTime(0)
                    End If
                    arrOldTime(!列号 - 1) = Nvl(!原始时间)
                    vsfTab.TextMatrix(intRow, col_tab原始时间) = Join(arrOldTime, "'")
                    If intRow Mod 2 = 0 Then '数据
                        If Val(Nvl(!记录类型)) = 1 Then
                            strTime = GetAnimalItemTime(intRow, Val(!列号), 0, strInfo)
                            If InStr(1, strTime, ";") > 0 Then lngHour = DateDiff("h", CDate(Split(strTime, ";")(0)), CDate(Split(strTime, ";")(1))) + 1
                            If lngHour > 24 Then lngHour = 24
                            If mbln录入小时 And int频次 = 1 And bln汇总 And Not InStr(1, !数值, ")") > 0 Then
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1) = "(" & lngHour & "h)" & !数值
                            Else
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1) = !数值
                            End If
                            If Val(zlStr.Nvl(!数据来源)) <> 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = 255
                            Else
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = &H80000012
                            End If
                            If Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(1)) = 1 And _
                                Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(4)) = 0 Then
                                 vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = Val(zlStr.Nvl(!未记说明))
                            End If
                        End If
                    Else
                        If Val(Nvl(!记录类型)) = 1 Then
                            strTime = GetAnimalItemTime(intRow, Val(!列号), 0, strInfo)
                            If InStr(1, strTime, ";") > 0 Then strTime = Format(Split(strTime, ";")(0), "hh:mm") & "～" & Format(Split(strTime, ";")(1), "hh:mm")
                            vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1) = IIf(Split(vsfTab.RowData(intRow), ";")(1) = 0, Format(!时间, "hh:mm"), strTime)
                            If Val(zlStr.Nvl(!数据来源)) <> 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = 255
                            Else
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = &H80000012
                            End If
                            If Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(1)) = 1 And _
                                Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(4)) = 0 Then
                                vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = Val(zlStr.Nvl(!未记说明))
                            End If
                        End If
                    End If
                End If
            Next intRow
        .MoveNext
        Loop
    End With
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function ShowTabDetail(ByVal intRow As Integer, ByVal intNewRow As Integer, ByVal intType As Integer)
'--------------------------------------------------
'功能：表格项目详细展示
'--------------------------------------------------
    Dim strTime As String
    Dim strInfo As String
    
    On Error GoTo Errhand
    With mrsTableDetail
        Do While Not .EOF
            If Val(Nvl(!记录类型)) = 11 Then
                vsfTabDetail.TextMatrix(intRow, vsfTab.FixedCols) = "(" & !结果 & "h)" & vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols)
            Else
                strTime = GetAnimalItemTime(vsfTab.Row, intRow, 0, strInfo)
                If InStr(1, strTime, ";") > 0 Then strTime = Format(Split(strTime, ";")(0), "hh:mm") & "～" & Format(Split(strTime, ";")(1), "hh:mm")
                If intType = 3 And strInfo <> "" Then lblStb.Caption = strInfo: lblStb.ForeColor = 255: Exit Function
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols) = IIf(intType = 3, strTime, Format(!时间, "hh:mm"))
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 1) = !结果
                vsfTabDetail.TextMatrix(intRow, col_tab原始时间) = Nvl(!时间)
                Select Case !数据来源
                    Case 0, 9
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "体温单录入"
                    Case 1
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "记录单同步"
                    Case 3
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "移动设备录入"
                    Case Else
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "其他设备同步"
                End Select
                
                If Val(zlStr.Nvl(!数据来源)) <> 0 Then
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = 255
                Else
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = &H80000012
                End If
                If Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab字符串), ",")(1)) = 1 And _
                    Val(Split(vsfTab.TextMatrix(intNewRow, COL_tab字符串), ",")(4)) = 0 Then
                    vsfTabDetail.Cell(flexcpForeColor, intRow, vsfTabDetail.FixedCols, intRow, vsfTabDetail.FixedCols + 2) = Val(zlStr.Nvl(!未记说明))
                End If
            End If
            intRow = intRow + 1
            vsfTabDetail.Rows = vsfTabDetail.Rows + 1
            vsfTabDetail.TextMatrix(intRow, COL_tab项目名) = zlStr.Nvl(!项目名称)
            vsfTabDetail.TextMatrix(intRow, COL_tab项目名称) = zlStr.Nvl(!项目名称)
            vsfTabDetail.TextMatrix(intRow, COL_tab字符串) = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
            vsfTabDetail.TextMatrix(intRow, COL_tab项目序号) = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号)
            .MoveNext
        Loop
    End With
    
    vsfTabDetail.Tag = vsfTabDetail.Rows - 1
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetAnimalItemTime(ByVal intRow As Integer, ByVal intNO As Integer, Optional IntMode As Integer = 0, Optional strInfo As String = "") As String
'--------------------------------------------------------------------------------
'功能:获取体温表格项目某频次的时间
'arrTime 返回信息 包括 开始时间  结束时间
'参数：introw 当前行,intNo 序号,strInfo 错误信息 IntMode 1 返回中间点时间 0,返回开始时间和结束时间
'---------------------------------------------------------------------------------
    Dim strTmp As String, lng项目序号 As Long, str项目名称 As String, int频次 As Integer
    Dim int项目表示 As String, intType As Integer
    Dim arrStr() As String
    Dim rsTmp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String, strTime As String, strCurrDate As String
    Dim intHour As Integer
    Dim lngRow As Long
    Dim strDate As String
    Dim strReturn As String
    Dim bln波动 As Boolean

    On Error GoTo Errhand
    
    strDate = mstrBegin
    strInfo = ""
    lngRow = intRow - vsfTab.FixedRows + 1
    strTmp = vsfTab.TextMatrix(intRow, COL_tab字符串)
    lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
    str项目名称 = vsfTab.TextMatrix(intRow, COL_tab项目名)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    int频次 = Val(arrStr(3))
    int项目表示 = Val(arrStr(4))
    
    bln波动 = IsWaveItem(lng项目序号)
    
    '汇总/波动 项目类型=2
    If int项目表示 = 4 Or bln波动 Then
        intType = 2
        If int频次 = 0 Then
            int频次 = 2
        ElseIf int频次 > 2 Then
            int频次 = 2
        End If
        
        '由参数确定汇总/波动项目今天录入昨天的数据还是当天的数据
        If Not mbln汇总当天 Then strDate = CDate(mstrBegin) - 1
    Else
        intType = 1
    End If
    
    
    '根据类型，频次和序号 不可能找不到信息
    mrsTabTime.Filter = "类型=" & intType & " and 频次=" & int频次 & " and 序号=" & intNO
    If mrsTabTime.RecordCount = 0 Then
        strInfo = "请在护理项目管理中设置[" & IIf(intType = 2, "汇总项目", "体温表格项目") & "]时段信息!"
        Exit Function
    End If
    
    With mrsTabTime
        .MoveFirst
        intHour = CInt(24 / int频次)
        strBegin = Format(IIf(IsDate(Trim(Nvl(!开始))) = False, (Val(Nvl(!序号)) - 1) * intHour & ":00:00", !开始), "HH:mm:ss")
        strEnd = Format(IIf(IsDate(Trim(Nvl(!结束))) = False, Val(Nvl(!序号)) * intHour - 1 & ":59:59", !结束), "HH:mm:ss")
        If intNO = int频次 Then
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(DateAdd("d", 1, CDate(strDate)), "YYYY-MM-DD") & " " & strEnd
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        Else
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = strBegin
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        End If
    End With
    If strBegin < mstrBTime Then strBegin = mstrBTime
    If strEnd > mstrETime Then strEnd = mstrETime
    '获取当前时间
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    '获取中间时间
    intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
    strTime = DateAdd("H", intHour, CDate(strBegin)) '中点时间
    
    If CDate(strCurrDate) >= CDate(strBegin) And CDate(strCurrDate) <= CDate(strEnd) Then
        strTime = strCurrDate
    End If
    '病人未出院，且发送了出院医嘱，如果预出院时间在当前录入表格对应的时间范围内，且小于中点时间则以预出院时间为准
    If mbln出院 = False And IsDate(mstrPreOutDate) Then
        If Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") >= Format(strBegin, "YYYY-MM-DD HH:mm") And _
            Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") <= Format(strEnd, "YYYY-MM-DD HH:mm") And _
            Format(mstrPreOutDate, "YYYY-MM-DD HH:mm") < Format(strTime, "YYYY-MM-DD HH:mm") Then
            strTime = Format(mstrPreOutDate, "YYYY-MM-DD HH:mm")
        End If
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) > CDate(strEnd) Then
            strInfo = "第" & lngRow & "列[" & str项目名称 & "]的结束时间：" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "，不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
            Exit Function
        End If
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) < CDate(strBegin) Then
            If mbln出院 = False Then
                strInfo = "第" & lngRow & "列[" & str项目名称 & "]的开始时间：" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "，已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            Else
                strInfo = "第" & lngRow & "列[" & str项目名称 & "]的开始时间：" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "，不能大于[病人出院时间或文件结束时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
            End If
            Exit Function
        End If
    End If
    
ErrNext:
    '检查病人转科后的补录时限
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, strEnd, strCurrDate) Then
        strInfo = "记录数据时间[" & strTime & "]有误！[超过数据补录的有效时限:" & mlngHours & "小时]"
        Exit Function
    End If
    
    Select Case IntMode
        Case 0
            strReturn = Format(CDate(strBegin), "YYYY-MM-DD HH:mm:ss") & ";" & Format(CDate(strEnd), "YYYY-MM-DD HH:mm:ss")
        Case 1
            strReturn = Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss")
    End Select
    
    GetAnimalItemTime = strReturn
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetChildNo(ByVal lngNo As Long) As ADODB.Recordset
    '功能：获取汇总项目详细
    Dim strFileds As String
    Dim strValue As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    strFileds = "序号," & adDouble & ",18|父序号," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    strFileds = "序号|父序号"
    mrsCollect.Filter = "父序号 =" & lngNo
    If mrsCollect.RecordCount > 0 Then
        Do While Not mrsCollect.EOF
            strValue = mrsCollect!序号 & "|" & mrsCollect!父序号
            Call Record_Add(rsTemp, strFileds, strValue)
            mrsCollect.MoveNext
        Loop
    End If
    Set GetChildNo = rsTemp
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPart(ByVal lng项目序号 As Long) As String
'功能:提取默认的体温部位
    Dim strPart As String
    mrsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
    If mrsPart.RecordCount > 0 Then strPart = zlStr.Nvl(mrsPart("部位"))
    GetPart = strPart
End Function



Private Function CheckShow(ByVal strBegin As String, ByVal strEnd As String, ByVal CenterTime As String) As Boolean
'-------------------------------------------------
'功能：对比两个时间点那个更靠近终点时间
'strbegin 对比的时间  strend当前时间
'--------------------------------------------------
    Dim strTime As String
    Dim blnAllow As Boolean
    
    If Abs(DateDiff("s", CDate(Format(strBegin, "YYYY-MM-DD HH:mm:ss")), CDate(CenterTime))) < Abs(DateDiff("s", CDate(Format(strEnd, "YYYY-MM-DD HH:mm:ss")), CDate(CenterTime))) Then
        blnAllow = True
    Else
        blnAllow = False
    End If
    
    CheckShow = blnAllow
End Function

Private Function UpdateCurveDate(ByVal vsf As Object, ByVal intRow As Integer, ByVal intCOl As Integer, ByVal intType As Integer, _
    Optional blnComList As Boolean = False) As Boolean
    '-------------------------------------------------------------------
    '功能：进行体温项目，体温详细，上下标的数据保存 记录集更新
    '参数：数据所在行，列，来源表格，是否是下拉列表
    '-------------------------------------------------------------------
    Dim lng项目序号 As Long, strName As String, strTime As String
    Dim int复试合格 As String
    Dim strEditData As String
    Dim str未记 As String, str部位 As String
    Dim int修改状态 As Integer
    Dim strData As String
    
    On Err GoTo Errhand
    If intType = 1 Or intType = 3 Then
        If vsf.EditText = vsf.Tag And vsf.EditText <> "" Then vsf.TextMatrix(intRow, COL_修改状态) = 0
        If blnComList = True Then
            str部位 = vsf.EditText
            If str部位 = "" Then str部位 = vsf.TextMatrix(intRow, COL_部位)
            If str部位 <> vsf.Tag Then vsf.TextMatrix(intRow, COL_修改状态) = 2
        Else
            str部位 = vsf.TextMatrix(intRow, COL_部位)
        End If
        lng项目序号 = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号))
        If intCOl = COL_数据 Then
            strData = vsf.EditText
        Else
            strData = Trim(vsf.TextMatrix(intRow, COL_数据))
        End If
        strTime = Trim(vsf.TextMatrix(intRow, COL_时间))
        If lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
            '反转脉搏和心率数据
            If mbln脉搏共用显示 And InStr(strData, "/") > 0 Then
                strData = Split(strData, "/")(1) & "/" & Split(strData, "/")(0)
            End If
        End If
        str未记 = Trim(vsf.TextMatrix(intRow, Col_未记说明))
        If strData <> "" Then str未记 = ""
        '进行数据更新处理
        With mrsCurve
            .Filter = "项目序号=" & lng项目序号 & " and 时间='" & Format(vsf.TextMatrix(intRow, col_原始时间), "YYYY-MM-DD HH:mm:ss") & "'"
            int复试合格 = IIf(vsf.TextMatrix(intRow, COL_复试合格) = "√", 1, 0)
            If .RecordCount <> 0 Then
                int修改状态 = Val(vsf.TextMatrix(intRow, COL_修改状态))
                Select Case int修改状态
                    Case 0 '未做操作
                    Case 1
                        !状态 = 1
                        !数值 = strData
                        !部位 = str部位
                        !显示 = IIf(vsf.TextMatrix(intRow, COL_显示) = "√", 1, 0)
                        !复试合格 = IIf(vsf.TextMatrix(intRow, COL_复试合格) = "√", 1, 0)
                        !修改 = 0
                        !数据来源 = 0
                        !未记说明 = str未记
                        
                    Case 2 '修改
                        If !状态 = 1 Then
                            !状态 = 1
                        Else
                            !状态 = 2
                        End If
                        !数值 = strData
                        !部位 = str部位
                        !显示 = IIf(vsf.TextMatrix(intRow, COL_显示) = "√", 1, 0)
                        !复试合格 = IIf(vsf.TextMatrix(intRow, COL_复试合格) = "√", 1, 0)
                        !未记说明 = str未记
                        !数据来源 = 0
                        !修改 = 0
                    Case 3 '删除
                        !数值 = ""
                        !未记说明 = str未记
                        !状态 = 3
                    Case 4 '新增后删除
                        .Delete
                    Case 5 '修改时间
                        !时间 = Format(dtpDate.Value & " " & vsf.EditText, "YYYY-MM-DD hh:mm:ss")
                        Select Case !状态
                            Case 0
                                !状态 = 5
                            Case 1
                                !状态 = 1
                            Case 2
                                !状态 = 2
                        End Select
                        vsf.TextMatrix(intRow, col_原始时间) = !时间
                    Case 6 '修改显示
                       !显示 = IIf(vsf.TextMatrix(intRow, COL_显示) = "√", 1, 0)
                       Select Case !状态
                            Case 0
                                !状态 = 6
                            Case 1
                                !状态 = 1
                            Case 2
                                !状态 = 2
                        End Select
                End Select
                .Update
            Else
                If (strData <> "" Or str未记 <> "") Then
                    If strTime = "" Then
                        strTime = Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss")
                    Else
                        strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
                    End If
                    vsf.TextMatrix(intRow, col_原始时间) = strTime
                    gstrFields = "序号|分组名|数值|部位|标记|时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
                    gstrValues = GetMaxNum(mrsCurve) & "|1)体温曲线项目|" & strData & "|" & str部位 & "|" & _
                        "0" & "|" & strTime & "|" & lng项目序号 & "|" & strName & "|" & _
                        "0" & "|" & str未记 & "|0|0|0|0|0|1|0|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
        End With
    ElseIf intType = 2 Then
        lng项目序号 = Val(vsfCurve.TextMatrix(intRow, COL_项目序号))
        If intCOl = COL_数据 Then
            strData = vsfCurve.EditText
        Else
            strData = Trim(vsfCurve.TextMatrix(intRow, COL_数据))
        End If
        str未记 = Trim(vsfCurve.TextMatrix(intRow, Col_未记说明))
        strTime = Trim(vsfCurve.TextMatrix(intRow, COL_时间))
        mrsNote.Filter = "记录类型=" & lng项目序号 & " and 时间='" & Format(vsfCurve.TextMatrix(intRow, col_原始时间), "YYYY-MM-DD HH:mm:ss") & "'"
        If mrsNote.RecordCount <> 0 Then
            int修改状态 = Val(vsfCurve.TextMatrix(intRow, COL_修改状态))
            Select Case int修改状态
                Case 0 '未做操作
                Case 1 '新增再修改
                    mrsNote!状态 = 2
                    mrsNote!内容 = strData
                    mrsNote!未记说明 = IIf(mrsNote!内容 = "", "", vsfCurve.Cell(flexcpBackColor, intRow, COL_颜色, intRow, COL_颜色))
                Case 2 '修改
                    mrsNote!状态 = 2
                    mrsNote!内容 = strData
                    mrsNote!未记说明 = IIf(mrsNote!内容 = "", "", vsfCurve.Cell(flexcpBackColor, intRow, COL_颜色, intRow, COL_颜色))
                Case 3 '删除
                    mrsNote!内容 = ""
                    mrsNote!未记说明 = ""
                    mrsNote!状态 = 3
                Case 4 '新增后删除
                    mrsNote!状态 = 4
                Case 5
                    mrsNote!时间 = Format(dtpDate.Value & " " & vsfCurve.EditText, "YYYY-MM-DD hh:mm:ss")
                    Select Case mrsNote!状态
                        Case 0
                            mrsNote!状态 = 5
                        Case 1
                            mrsNote!状态 = 1
                        Case 2
                            mrsNote!状态 = 2
                            
                    End Select
            End Select
            mrsNote.Update
        Else
            If lng项目序号 = 2 Then
                    strName = "上标说明"
                ElseIf lng项目序号 = 6 Then
                    strName = "下标说明"
                End If
            If strData <> "" Or str未记 <> "" Then
                If strTime = "" Then
                    strTime = Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss")
                Else
                    strTime = Format(dtpDate.Value & " " & strTime, "YYYY-MM-DD HH:mm:ss")
                End If
                vsfCurve.TextMatrix(intRow, col_原始时间) = strTime
                gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
                gstrValues = GetMaxNum(mrsNote) & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lng项目序号 & "|" & strData & "|" & strName & "|" & IIf(str未记 = "", vsfCurve.Cell(flexcpBackColor, intRow, COL_颜色, intRow, COL_颜色), "") & "|0|0|0|0|0|1"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            End If
        End If
        
    End If
      
    UpdateCurveDate = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMaxNum(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'功能:获取记录mrsCurve中的最大序号
'----------------------------------------------------
    On Error GoTo Errhand
    rsTmp.Filter = 0
    rsTmp.Sort = "序号 Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxNum = 1
    Else
        GetMaxNum = Val(rsTmp!序号) + 1
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMaxID(ByVal rsTmp As ADODB.Recordset) As Long
'----------------------------------------------------
'功能:获取记录mrsCurve中的最大序号
'----------------------------------------------------

    On Error GoTo Errhand
    rsTmp.Filter = 0
    rsTmp.Sort = "id Desc"
    If rsTmp.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(rsTmp!Id) + 1
    End If
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabControl()
'--------------------------------------------------------------------------------
'功能:初始化TabControl
'--------------------------------------------------------------------------------
    Dim tabItem As TabControlItem
    Dim CtlFont As stdFont
    
    On Error GoTo Errhand
    With tbcThis
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ShowIcons = True
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ClientFrame = xtpTabFrameSingleLine
            .DisableLunaColors = False
            .Layout = xtpTabLayoutAutoSize
            Set CtlFont = .Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
            Set .Font = CtlFont
        End With
        
        Set tabItem = .InsertItem(1, "体温曲线", picCurve.hWnd, 0)
        tabItem.Tag = "曲线"
        Set tabItem = .InsertItem(2, "体温表格", picTab.hWnd, 0)
        tabItem.Tag = "表格"
        
        
        If gintEditorCurveState = 0 Then
            .Item(0).Selected = True
        Else
            .Item(1).Selected = True
        End If
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long, intType As Integer, i As Integer
    Dim strValue As String, strPart As String, strPart1 As String, strName As String
    Dim strTime As String, strErrMsg As String
    Dim strTmp As String, arrStr() As String, arrTime() As String
    Dim cbrCheck As CommandBarControl
    Dim rsTemp As New ADODB.Recordset

    Select Case Control.Id
        Case conMenu_Edit_Save '保存
            If Not SaveData Then Exit Sub
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
        Case conMenu_Edit_Reuse '取消
            Call txtEdit_KeyPress(vbKeyEscape)
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
        Case conMenu_Edit_NewItem '添加活动项目
            Call txtEdit_KeyPress(vbKeyEscape)
            mblnScroll = True
            If frmCaseTendBodyActiveItem.ShowMe(vsfTab, Me, mT_Patient.lng护理等级, mT_Patient.lng婴儿, mT_Patient.lng科室ID) Then
                vsfTab.Refresh
            End If
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 30, conMenu_Edit_Append * 10 + 31, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If vsfTab.Tag <> "" Then
                If vsfTab.Row < vsfTab.FixedRows Or vsfTab.Col < vsfTab.FixedCols Then Exit Sub
                lngRow = vsfTab.Row
                lngCol = vsfTab.Col
                lngItemNO = Val(vsfTab.TextMatrix(lngRow, COL_tab项目序号))
                strName = vsfTab.TextMatrix(lngRow, COL_tab项目名)
                strValue = Trim(vsfTab.TextMatrix(lngRow, lngCol))
                strTmp = vsfTab.TextMatrix(lngRow, COL_tab字符串)
            Else
                If vsfTabDetail.Row < vsfTabDetail.FixedRows Or vsfTabDetail.Col < vsfTabDetail.FixedCols Or vsfTabDetail.Row > Val(vsfTabDetail.Tag) Then Exit Sub
                lngRow = vsfTabDetail.Row
                lngCol = vsfTabDetail.Col
                lngItemNO = Val(vsfTabDetail.TextMatrix(lngRow, COL_tab项目序号))
                strName = vsfTabDetail.TextMatrix(lngRow, COL_tab项目名)
                strValue = Trim(vsfTabDetail.TextMatrix(lngRow, lngCol))
                strTmp = vsfTabDetail.TextMatrix(lngRow, COL_tab字符串)
            End If
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            
            intType = 0
            If picEdit.Visible = True And txtEdit.Visible = True Then intType = 1
            If intType = 1 Then strValue = txtEdit.Text
            strPart = ""
            If InStr(1, "," & gint大便 & "," & gint入液 & ",", "," & lngItemNO & ",") = 0 Then Exit Sub
            Select Case Control.Id
                Case conMenu_Edit_Append * 10 + 1
                    strPart = "E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                    If InStr(1, UCase(strValue), "E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 2
                    strPart = "/E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 30
                    strPart = "※"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 31
                    strPart = "*"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 4
                    strPart = "☆"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 5
                    strPart = "C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                    If InStr(1, UCase(strValue), "C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "C") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 6
                    strPart = "/C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                Case conMenu_Edit_Append * 10
                    strPart = ""
                    If lngItemNO = gint大便 Then
                        For i = 0 To 4
                            Select Case i
                                Case 0
                                    strPart1 = "E"
                                Case 1
                                    strPart1 = "/"
                                Case 2
                                    strPart1 = "*"
                                Case 3
                                    strPart1 = "※"
                                Case 4
                                    strPart1 = "☆"
                            End Select
                            strValue = Replace(UCase(strValue), strPart1, "")
                        Next i
                    Else
                        strValue = Replace(UCase(Replace(UCase(strValue), "C", "")), "/", "")
                    End If
            End Select
            
            If IsNumeric(strValue) Then
                strValue = strValue
            Else
                strValue = ""
            End If
            strValue = strValue & Trim(strPart)
            If Left(strValue, 1) = "/" Then strValue = 1 & strValue
            
            If intType = 1 Then
                txtEdit.Text = strValue
                For Each cbrCheck In mcbrToolBar.Controls(5).Controls
                    If cbrCheck.Id = Control.Id Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next

                Exit Sub
            End If
             
            '非编辑状态下
            If IsWaveItem(lngItemNO) And InStr(1, Trim(vsfTab.TextMatrix(lngRow, lngCol)), "-") <> 0 And vsfTab.Tag <> "" Then
                strErrMsg = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                Exit Sub
            End If
            Call txtEdit_KeyPress(vbKeyEscape)
            strPart = CStr(arrStr(7))
            If vsfTab.Tag <> "" Then
                strTime = Format(Split(vsfTab.TextMatrix(vsfTab.Row, col_tab原始时间), "'")(vsfTab.Col - vsfTab.FixedCols), "YYYY-MM-DD hh:mm:ss")
            Else
                strTime = Format(vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tab原始时间), "YYYY-MM-DD hh:mm:ss")
            End If
            If strTime = "" Then
                strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strErrMsg)
                If strErrMsg <> "" Then lblStb.Caption = strErrMsg: lblStb.ForeColor = 255: Exit Sub
                strTime = Format(DateAdd("n", DateDiff("n", Split(strTime, ";")(0), Split(strTime, ";")(1)) / 2, Split(strTime, ";")(0)), "YYYY-MM-DD hh:mm:ss")
                If IsExistData(strTime, lngItemNO) = False Then
                    Exit Sub
                End If
            End If
            
            mrsTableDetail.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "' and 时间='" & strTime & "'"
            If mrsTableDetail.RecordCount > 0 Then
                If mrsTableDetail!状态 <> 1 Then  '原有的数据 修改、删除后的状态始终为2
                    mrsTableDetail!状态 = 2
                    mrsTableDetail!结果 = strValue
                Else '对于新增数据的处理
                    If Trim(vsfTab.TextMatrix(lngRow, lngCol)) = "" Then
                        mrsTableDetail.Delete
                    Else
                        mrsTableDetail!状态 = 1
                        mrsTableDetail!结果 = strValue
                    End If
                End If
                mrsTableDetail.Update
            Else '不存在记录就新增数据
                If Trim(strValue) <> "" Then
                    
                    gstrFields = "id|分组名|结果|体温部位|标记|时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
                    gstrValues = GetMaxID(mrsTableDetail) & "|2)体温表格项目|" & strValue & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lngItemNO & "|" & strName & "|0||0|0|0|0|0|1|" & lngCol - vsfTab.FixedCols + 1 & "|1"
                    Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
                    If vsfTab.Tag <> "" Then
                        arrTime = Split(vsfTab.TextMatrix(vsfTab.Row, col_tab原始时间), "'")
                        arrTime(vsfTab.Col - vsfTab.FixedCols) = strTime
                        vsfTab.TextMatrix(vsfTab.Row, col_tab原始时间) = Join(arrTime, "'")
                    Else
                        vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tab原始时间) = strTime
                    End If
                End If
            End If
            
            mrsTableDetail.Filter = "状态<> 4 "
    
            gstrFields = "ID," & adDouble & ",18|分组名," & adLongVarChar & ",40|结果," & adLongVarChar & ",400|体温部位," & adLongVarChar & ",200|" & _
                 "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
                 "复试合格," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|原始显示状态," & adDouble & ",1|" & _
                 "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
            Call Record_Init(rsTemp, gstrFields)
            
            Do While Not mrsTableDetail.EOF
                rsTemp.AddNew
                For i = 0 To mrsTableDetail.Fields.Count - 1
                    rsTemp.Fields(mrsTableDetail.Fields(i).Name).Value = mrsTableDetail.Fields(i).Value
                Next i
                rsTemp.Update
                mrsTableDetail.MoveNext
            Loop
            
            Call InitTableData(rsTemp)
            Call ShowTable
            If vsfTab.Tag <> "" Then
                 vsfTab.TextMatrix(lngRow, lngCol) = strValue
                Call vsfTab_AfterRowColChange(0, 0, lngRow, lngCol)
            Else
                vsfTabDetail.TextMatrix(lngRow, lngCol) = strValue
                Call vsfTabDetail_AfterRowColChange(0, 0, lngRow, lngCol)
            End If
            
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
'        Case conMenu_View_Forward
'            dtpDate.Value = dtpDate.Value - 1
'            Call dtpDate_Change
'        Case conMenu_View_Backward
'            dtpDate.Value = dtpDate.Value + 1
'            Call dtpDate_Change
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)

    On Error Resume Next
    If stbThis.Visible = True Then Bottom = stbThis.Height
    
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With

    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("中联")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub



Private Function InitRecordSet() As Boolean
'----------------------------------------------------------------
'功能:初始化记录集 包括部位信息，汇总项目时段，记录频次时段
'----------------------------------------------------------------
    On Error GoTo Errhand
    '提取所有部位信息
    mstrSQL = "Select 项目序号,部位,缺省项 From 体温部位"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, Me.Caption)
    
    '提取共用记录集信息
    Call InitPublicData
    
    InitRecordSet = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Call picToolBarReSize
    picSplit.Left = lngLeft
    picSplitTab.Left = lngLeft
    
    With tbcThis
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
    picNull.Move lngLeft, lngTop + 350, lngRight, lngBottom - 350
    With lblInfo
        .Top = (picNull.Height - .Height) / 2
        .Left = (picNull.Width - .Width) / 2
    End With
    picNull.Visible = False
End Sub

Private Function GetCenterTime(ByVal dBegin As Date, ByVal dEnd As Date, Optional intBound As Integer = 0) As String
'------------------------------------------------------------------------------------
'功能:获取某段时间的中点时间,如果当前时间在本段范围并且在中间时间内则以当前时间为准
'------------------------------------------------------------------------------------
    Dim dblvalue As Double
    Dim strTime As String, strCurDate As String
    Dim i As Integer
    
    On Error GoTo Errhand
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    dblvalue = DateDiff("s", dBegin, dEnd)
    strTime = Format(DateAdd("s", Fix(dblvalue / 2), dBegin), "YYYY-MM-DD HH:mm:ss")
    If strTime < mstrBTime Then
        strTime = mstrBTime
    End If
    If strTime > mstrETime Then
        strTime = mstrETime
    End If
    
    For i = 0 To UBound(marrTime)
        If Format(strTime, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strTime, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    If i <= UBound(marrTime) Then
        If gintHourBegin + i * T_BodyStyle.lng时间间隔 = 24 Then
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & gintHourBegin + i * T_BodyStyle.lng时间间隔 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    intBound = i
    
    If CDate(strCurDate) >= dBegin And CDate(strCurDate) <= dEnd And CDate(strCurDate) < CDate(strTime) Then
        strTime = strCurDate
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    GetCenterTime = strTime
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim frmMain As Form
    Dim blnEnable As Boolean
    Dim strCurrentdate As String
    
    strCurrentdate = zlDatabase.Currentdate
    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
            picNull.Visible = (vsfTab.TextMatrix(1, COL_tab项目序号) = -999 And tbcThis.Selected.Tag = "表格")
            Control.Enabled = IIf(IsChange = True, True, False)
        Case conMenu_Edit_NewItem
            If tbcThis.Selected.Tag = "表格" Then
                Control.Visible = True
                Control.Enabled = Not mblnFileBack
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Append
            If tbcThis.Selected.Tag = "表格" Then
                Control.Visible = True
            Else
                Control.Visible = False
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = (is大便或入液(1) Or is大便或入液(2)) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is大便或入液(1) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is大便或入液(2) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
    End Select
    
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
'        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
'        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
    If Format(mstrETime, "yyyy-mm-dd") = Format(strCurrentdate, "yyyy-mm-dd") Then mstrETime = strCurrentdate
End Sub


Private Function is大便或入液(ByVal intType As Integer) As Boolean
    '--------------------------------------------------
    '检查是否是大便项目或入夜项目  大便项目序号=10 入夜=9
    'intType=1 为大便项目 否则为入液项目
    '--------------------------------------------------
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    Dim strTmp As String, strName As String, arrStr() As String
    On Error GoTo Errhand
    
    If vsfTab.Col < vsfTab.FixedCols Or vsfTab.Row < vsfTab.FixedRows Then Exit Function
    If mblnInit = False Then Exit Function
    
    '提取项目序号
    lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名)
    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    
    '检查记录频次和项目表示
    If vsfTab.Col > vsfTab.FixedCols + Val(arrStr(3)) - 1 Then Exit Function
    If InStr(1, ",2,3,5,", "," & Val(arrStr(4)) & ",") > 0 Then Exit Function
    
    '检查是否是同步的数据
    If vsfTab.Tag <> "" Then
        mrsTableDetail.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "'" & _
            "   and 列号=" & vsfTab.Col - vsfTab.FixedCols + 1
    Else
        mrsTableDetail.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "'" & _
            "   and 时间 ='" & vsfTabDetail.TextMatrix(vsfTabDetail.Row, col_tab原始时间) & "'"
    End If
    If mrsTableDetail.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsTableDetail!数据来源) & ",") = 0 Then
            Exit Function
        End If
    End If
    
    is大便或入液 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cmdColor_Click()
    Call txtEdit_KeyDown(vbKeyDown, vbShiftMask)
End Sub

Private Sub dtpDate_Change()
    Dim strDate As String
    Dim cbrControl As CommandBarControl
     If IsChange Then
        If MsgBox("病人体温数据已经发生改变,请问是否需要保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Exit Sub
        End If
    End If
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then Exit Sub
    imgbtn(1).Enabled = True
    imgbtn(0).Enabled = True
'    If dtpDate.Enabled = True Then dtpDate.SetFocus
'    Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Forward)
'    cbrControl.Enabled = True
'    Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Backward)
'    cbrControl.Enabled = True
    
'    If dtpDate.Value = dtpDate.MinDate Then
'        Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Forward)
'        cbrControl.Enabled = False
'    End If
'    If dtpDate.Value = dtpDate.MaxDate Then
'        Set cbrControl = mcbrToolBar.Controls.Find(, conMenu_View_Backward)
'        cbrControl.Enabled = False
'    End If
    If dtpDate.Value = dtpDate.MinDate Then
        imgbtn(1).Picture = ilsDate.ListImages("preGray").Picture
        imgbtn(1).Enabled = False
    Else
        imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
    End If
    If dtpDate.Value = dtpDate.MaxDate Then
        imgbtn(0).Picture = ilsDate.ListImages("nextGray").Picture
        imgbtn(0).Enabled = False
    Else
        imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
    End If
End Sub

Private Sub dtpDate_CloseUp()
    Call SetColSelect
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call tbcThis_SelectedChanged(tbcThis.Selected)
    End If
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not dtpDateChageDate(Format(dtpDate.Value, "YYYY-MM-DD")) Then
        Cancel = True
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
    mblnStart = False
    Call SetColSelect(True)
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If IsChange Then
        If MsgBox("病人体温数据已经发生改变,请问是否需要保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mbln出院 = False
    mblnAllRefresh = False
    If Not (mrsCurve Is Nothing) Then Set mrsCurve = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsTable = Nothing
    If Not (mrsTable Is Nothing) Then Set mrsTableDetail = Nothing
    If Not (mrsTableDetail Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsRecodeID Is Nothing) Then Set mrsRecodeID = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    Call UnLoadOptTime
    '保存窗体
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function IsChange()
'数据是否改变
    Dim blnChange As Boolean
    
    If mrsCurve.State = adStateOpen Then
        mrsCurve.Filter = "状态 <> 0 "
        If mrsCurve.RecordCount > 0 Then blnChange = True
    End If
    
    If mrsNote.State = adStateOpen Then
        mrsNote.Filter = "状态 <> 0 "
        If mrsNote.RecordCount > 0 Then blnChange = True
    End If
    
    If mrsTableDetail.State = adStateOpen Then
        mrsTableDetail.Filter = "状态 <> 0 "
        If mrsTableDetail.RecordCount > 0 Then blnChange = True
    End If
    IsChange = blnChange
End Function


Private Sub imgbtn_Click(Index As Integer)
    Select Case Index
        Case 1
            dtpDate.Value = dtpDate.Value - 1
            Call dtpDate_Change
        Case 0
            dtpDate.Value = dtpDate.Value + 1
            Call dtpDate_Change
    End Select
End Sub


Private Sub lblCheck_DblClick()
    Call picEdit_KeyPress(vbKeySpace)
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strData As String
    Dim blnAllow As Boolean
    Dim i As Integer
    Dim blnTab As Boolean '是否是主表
    Dim arrTag() As String
    
    strData = ""
    blnAllow = True
    If InStr(1, lbllst(Index).Tag, "|") > 0 Then arrTag = Split(lbllst(Index).Tag, "|")
    If UBound(arrTag) > 1 Then
        If Split(lbllst(Index).Tag, "|")(2) = 1 Then blnTab = True
    End If
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        For i = 0 To lstSelect(Index).ListCount - 1
            If lstSelect(Index).Selected(i) = True Then
                strData = strData & "," & Replace(lstSelect(Index).List(i), ",", "")
            End If
        Next
        If Left(strData, 1) = "," Then strData = Mid(strData, 2)
        If strData <> lstSelect(Index).Tag Then
            If blnTab Then
                blnAllow = WriteIntoVfgTab(strData, vsfTab)
            Else
                blnAllow = WriteIntoVfgTab(strData, vsfTabDetail)
            End If
        End If
        If blnAllow = True Then Call vsfTab_KeyDown(vbKeyReturn, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If blnTab Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        Else
            Call vsfTabDetail_KeyDown(vbKeyLeft, 0)
        End If
    ElseIf KeyCode = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    ElseIf Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    End If
End Sub

'Private Sub imgbtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Select Case Index
'        Case 0
'            imgbtn(0).Picture = IIf(imgbtn(0).Enabled = True, ilsDate.ListImages("nextLight").Picture, ilsDate.ListImages("nextGray").Picture)
'        Case 1
'            imgbtn(1).Picture = IIf(imgbtn(1).Enabled = True, ilsDate.ListImages("preLight").Picture, ilsDate.ListImages("preGray").Picture)
'    End Select
'    imgDefault.Tag = Index + 1
'End Sub


Private Sub lst未记_DblClick()
    Dim str未记 As String
    Dim intRow As Integer
    Dim intCOl As Integer
    Dim intCount As Integer
    Dim intRows As Integer
    Dim blnAllow As Boolean
    Dim strTime As String
    Dim lng项目序号 As Long
    
    If lst未记.Tag <> "" Then
        str未记 = lst未记.Text
        intRow = Split(lst未记.Tag, "|")(1)
        intCOl = Split(lst未记.Tag, "|")(2)
        Select Case Split(lst未记.Tag, "|")(0)
            Case 1
                vsfCurve.Row = intRow
                vsfCurve.Col = intCOl
                vsfCurve.TextMatrix(intRow, intCOl) = str未记
                strTime = vsfCurve.TextMatrix(intRow, COL_时间)
                vsfCurve.TextMatrix(intRow, COL_数据) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                vsfCurve.TextMatrix(intRow, COL_颜色) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", " ", Space(vsfCurve.Row))
                vsfCurve.TextMatrix(intRow, COL_部位) = ""
                vsfCurve.TextMatrix(intRow, COL_复试合格) = ""
            Case 2
                strTime = IIf(vsfDetail.TextMatrix(intRow, COL_时间) = "", GetCenterTime(mstrBegin, mstrEnd), vsfCurve.TextMatrix(intRow, COL_时间))
                lng项目序号 = vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号)
                mrsCurve.Filter = "项目序号=" & lng项目序号 & " and  时间='" & Format(strTime, "YYYY-MM-DD hh:mm:ss") & "'"
                If mrsCurve.RecordCount > 0 Then
                    lblStb.Caption = "当前默认时间已存在数据，请先输入时间"
                    lblStb.ForeColor = 255
                    pic未记.Visible = False
                    Exit Sub
                End If
                vsfDetail.Row = intRow
                vsfDetail.Col = intCOl
                vsfDetail.TextMatrix(intRow, intCOl) = str未记
                vsfDetail.TextMatrix(vsfCurve.Row, COL_数据) = ""
                vsfDetail.TextMatrix(vsfCurve.Row, COL_部位) = ""
                vsfDetail.TextMatrix(vsfCurve.Row, COL_复试合格) = ""
        End Select
        pic未记.Visible = False
        lst未记.Visible = False: lst未记.Enabled = False
    End If
    
    blnAllow = True
    intCount = 0
    intRows = 0
    If Split(lst未记.Tag, "|")(0) = 2 Then
        Call UpdateCurveDate(vsfDetail, vsfDetail.Row, vsfDetail.Col, 3)
        Call vsfDetail.SetFocus
    Else
        If Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名)) = "1)体温曲线项目" Then
            '如果其它曲线的未记数据为空,直接更新
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" Then
                    If vsfCurve.TextMatrix(intRow, Col_未记说明) = "" And Trim(vsfCurve.TextMatrix(intRow, COL_数据)) = "" Then
                        intCount = intCount + 1
                    End If
                    intRows = intRows + 1
                End If
            Next
            '剩下的项目的数据与标记都为空则更新
            If intCount = intRows - 1 Then
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" And vsfCurve.TextMatrix(intRow, Col_未记说明) = "" Then
                        vsfCurve.TextMatrix(intRow, Col_未记说明) = str未记
                        vsfCurve.TextMatrix(intRow, COL_时间) = strTime
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_数据) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_颜色) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名) = "2)上下标说明", " ", Space(vsfCurve.Row))
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_部位) = ""
                        vsfCurve.TextMatrix(vsfCurve.Row, COL_复试合格) = ""
                    End If
                Next
            Else
                intCount = 0
            End If
        ElseIf Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名)) = "2)上下标说明" Then
            blnAllow = False
        End If
        vsfCurve.Cell(flexcpAlignment, vsfCurve.FixedRows, Col_未记说明, vsfCurve.Rows - 1, Col_未记说明) = flexAlignCenterCenter
        
        If blnAllow = True Then
            If intCount = 0 Then
                Call UpdateCurveDate(vsfCurve, vsfCurve.Row, vsfCurve.Col, 1)
            ElseIf intCount = intRows - 1 Then
                For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                    If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" Then
                        Call UpdateCurveDate(vsfCurve, intRow, Col_未记说明, 1)
                    End If
                Next
            End If
            Call vsfCurve.SetFocus
        End If
    End If
    
End Sub

Private Sub lst未记_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyEscape Then
        lst未记.Visible = False: lst未记.Enabled = False
        pic未记.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        Call lst未记_DblClick
    End If
End Sub

Private Sub lst未记_LostFocus()
    lst未记.Visible = False: lst未记.Enabled = False
    pic未记.Visible = False
End Sub

Private Sub OptTime_Click(Index As Integer)
    Dim strBegin As String, strEnd As String
    Dim blnTab As Boolean
    
    If Not mblnInit Then Exit Sub
    
    If OptTime(Index).Tag = "" Then Exit Sub
    strBegin = Split(OptTime(Index).Tag, ",")(0)
    strEnd = Split(OptTime(Index).Tag, ",")(1)
    strBegin = Format(Format(dtpDate.Value, " YYYY-MM-DD") & " " & strBegin, "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(dtpDate.Value, " YYYY-MM-DD") & " " & strEnd, "YYYY-MM-DD HH:mm:ss")
    
    If CDate(strBegin) < CDate(mstrBTime) Then
        strBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strEnd) > CDate(mstrETime) Then
        strEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    mstrBegin = strBegin
    mstrEnd = strEnd
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "～" & Format(mstrEnd, "HH:mm")
    Call ShowCurve
    Call ShowTabUpDown
    
    If mblnStart = False Then
        Call SetColSelect(True)
    End If
End Sub

Private Sub SetColSelect(Optional blnInit As Boolean = False, Optional intType As Integer = 1)
'-------------------------------------
'功能:设置表格选择列
'------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim intOldRow As Integer, intOldCol As Integer
    
    On Error GoTo Errhand
    If mblnInit = False Then Exit Sub
    mblnRefresh首行 = False
    If tbcThis.Selected.Tag = "曲线" Then
        vsfCurve.SetFocus
        If blnInit = True Then
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = COL_数据
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                vsfCurve.Col = COL_部位
            End If
            vsfCurve.Col = COL_数据
        Else
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = vsfCurve.Col
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                If intCOl < vsfCurve.Cols - 1 Then
                    vsfCurve.Col = intCOl + 1
                Else
                    If intRow < vsfCurve.Rows - 1 Then
                        vsfCurve.Row = intRow + 1
                    Else
                        If intRow - 1 > 0 Then
                            vsfCurve.Row = intRow - 1
                        End If
                    End If
                End If
            End If
            vsfCurve.Col = intCOl
        End If
        Call vsfCurve_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
    ElseIf tbcThis.Selected.Tag = "表格" Then
        If intType = 1 Then
            vsfTab.SetFocus
            If blnInit = True Then
                intOldRow = vsfTab.Row
                intOldCol = vsfTab.Col
                intRow = vsfTab.FixedRows
                intCOl = vsfTab.FixedCols
                If intRow = vsfTab.Row And intCOl = vsfTab.Col Then
                    Call vsfTab_BeforeRowColChange(intRow, intCOl, intRow, intCOl, False)
                End If
                vsfTab.Select vsfTab.FixedRows, vsfTab.FixedCols
            Else
                intOldRow = vsfTab.Row
                intOldCol = vsfTab.Col
                intRow = vsfTab.Row
                intCOl = vsfTab.Col
                vsfTab.Select vsfTab.Row, vsfTab.Col
            End If
            Call vsfTab_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
        Else
            vsfTabDetail.SetFocus
            If blnInit = True Then
                intOldRow = vsfTabDetail.Row
                intOldCol = vsfTabDetail.Col
                intRow = vsfTabDetail.FixedRows
                intCOl = vsfTabDetail.FixedCols
                If intRow = vsfTabDetail.Row And intCOl = vsfTabDetail.Col Then
'                    Call vsfTabDetail_BeforeRowColChange(intRow, intCOl, intRow, intCOl, False)
                End If
                vsfTabDetail.Select vsfTabDetail.FixedRows, vsfTabDetail.FixedCols
            Else
                intOldRow = vsfTabDetail.Row
                intOldCol = vsfTabDetail.Col
                intRow = vsfTabDetail.Row
                intCOl = vsfTabDetail.Col
                vsfTabDetail.Select vsfTabDetail.Row, vsfTabDetail.Col
            End If
            Call vsfTabDetail_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
            
        End If
    End If
    
     Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub picColor_Resize()
    With usrValue
        .Top = -450
        .Left = 0
        .Width = picValue.Width
        .Height = picValue.Height
    End With
End Sub

Private Sub picCurve_Resize()
    Dim i As Integer
    
    On Error Resume Next
    If mblnResize = True Then picSplit.Top = ScaleHeight * 0.6: mblnResize = False

    picSplit.Width = tbcThis.Width
    
    With lblTime
        .Top = picToolBar.Top + lblPtime.Top
        .Left = 200 + picToolBar.Width + picToolBar.Left
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With fraTime
        .Top = 0
        .Left = 0
        .Width = picCurve.Width
        .Height = 100 + picToolBar.Top + picToolBar.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
        
        
    With fraData
        .Left = 0
        .Width = picCurve.Width
        .Top = fraTime.Top + fraTime.Height
        .Height = picSplit.Top - .Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With vsfCurve
        .Top = 0
        .Left = 0
        .Width = fraData.Width
        .Height = fraData.Height
    End With

    With pic未记
        .Width = 1080 + 1080 * mintBigSize / 3
        .Height = 1100 + 1100 * mintBigSize / 3
        .Visible = False
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With lst未记
        .Top = 0
        .Left = 0
        .Width = pic未记.Width
        .Height = pic未记.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With

    With fraDetail
        .Top = picSplit.Top
        .Left = 0
        .Width = picCurve.Width
        .Height = picCurve.Height - picSplit.Top - picSplit.Height
    End With
    
     With vsfDetail
        .Top = 0
        .Left = 0
        .Width = fraDetail.Width
        .Height = fraDetail.Height
    End With
    
    With picValue
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    Call picToolBarReSize
End Sub

Private Sub picToolBarReSize()
    Dim i As Integer
    On Error Resume Next
    lblPtime.Left = 0
    lblPtime.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblPtime.Top = 45 + 45 * mintBigSize / 3
    For i = 0 To OptTime.Count - 1
        OptTime(i).Font.Size = mFontSize + mFontSize * mintBigSize / 3
        OptTime(i).Height = 300 + 300 * mintBigSize / 3
        OptTime(i).Width = 350 + 350 * mintBigSize / 3
        OptTime(i).Left = i * OptTime(i).Width + lblPtime.Left + lblPtime.Width + 10
    Next i
    picToolBar.Top = 210
    picToolBar.Width = OptTime(OptTime.Count - 1).Left + OptTime(OptTime.Count - 1).Width
    picToolBar.Height = OptTime(OptTime.Count - 1).Top + OptTime(OptTime.Count - 1).Height
    picToolBar.Left = 50
    picToolBar.Font.Size = mFontSize + mFontSize * mintBigSize / 3
End Sub


Private Sub picEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(KeyAscii)
    ElseIf KeyAscii = vbKeySpace Then
        If lblCheck.Caption = "√" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "√"
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        Call txtEdit_KeyDown(KeyAscii, 0)
    ElseIf KeyAscii = vbKeyLeft Then
        If txtEdit.Visible = False Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        End If
    End If
End Sub

Private Sub picHour_GotFocus()
    If picHour.Visible = True Then txtHour.SetFocus
End Sub

Private Sub PicLst_GotFocus()
    If PicLst.Visible = False Then Exit Sub
    If Trim(txtLst.Text) = "" Then
        PicLst.Tag = 0
        lstSelect(0).SetFocus
    Else
        PicLst.Tag = 1
        txtLst.SetFocus
    End If
End Sub

'Private Sub picDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'
'    Select Case imgDefault.Tag
'        Case 1
'            If imgbtn(0).Enabled = True Then imgbtn(0).Picture = ilsDate.ListImages("nextGreen").Picture
'        Case 2
'            If imgbtn(1).Enabled = True Then imgbtn(1).Picture = ilsDate.ListImages("preGreen").Picture
'    End Select
'
'End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
    If picSplit.Visible = True Then picSplit.SetFocus
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf Me.ScaleHeight - (picSplit.Top + Y) < Me.ScaleHeight * 0.4 Then
        picSplit.Top = Me.ScaleHeight * 0.6
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call picCurve_Resize

    picSplit.Tag = 0
End Sub



Private Sub LoadOptTime()
'-----------------------------------------
'功能:根据监测次数动态加载时点选择控件(OptTime)
'-----------------------------------------
    Dim i As Integer
    For i = 1 To T_BodyStyle.lng监测次数 - 1
        Load OptTime(i)
        OptTime(i).Visible = True
        Set OptTime(i).Container = picToolBar
        OptTime(i).Top = OptTime(i - 1).Top
        OptTime(i).ZOrder 0
    Next i
    Call picToolBarReSize
End Sub

Private Sub UnLoadOptTime()
'------------------------------------------
'功能:卸载时点选择控件
'------------------------------------------
    Dim i As Integer
    For i = OptTime.Count - 1 To 1 Step -1
        Unload OptTime(i)
    Next i
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    If mblnFileBack = True Then lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改.": lblStb.ForeColor = 255
End Sub

Private Sub picSplitTab_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplitTab.Tag = 1
    If picSplit.Visible = True Then picSplitTab.SetFocus
End Sub

Private Sub picSplitTab_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplitTab.Tag) = 0 Then Exit Sub
    
    If picSplitTab.Top + Y < 4000 Then
        picSplitTab.Top = 4000
    ElseIf Me.ScaleHeight - (picSplitTab.Top + Y) < Me.ScaleHeight * 0.3 Then
        picSplitTab.Top = Me.ScaleHeight * 0.7
    Else
        picSplitTab.Move picSplitTab.Left, picSplitTab.Top + Y
    End If
End Sub

Private Sub picSplitTab_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplitTab.Tag) = 1 Then Call picTab_Resize
    picSplit.Tag = 0
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    If mblnResize = True Then picSplitTab.Top = ScaleHeight * 0.6: mblnResize = False
    picSplitTab.Width = tbcThis.Width
    
    With fraTable
         .Top = 0
        .Left = 0
        .Width = picTab.Width
        .Height = picSplitTab.Top
    End With
    
    With vsfTab
        .Top = 100
        .Left = 0
        .Width = fraTable.Width
        .Height = fraTable.Height - .Top
    End With

    picEdit.Visible = False
    txtEdit.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblCheck.Font.Size = txtEdit.Font.Size

    With picColor
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    
    With lstSelect(0)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lstSelect(1)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With PicLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With picHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lblHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With fraTabDetail
        .Top = picSplitTab.Top
        .Left = 0
        .Width = picTab.Width
        .Height = picTab.Height - picSplitTab.Top
    End With
    
    With vsfTabDetail
        .Top = 0
        .Left = 0
        .Width = fraTabDetail.Width
        .Height = fraTabDetail.Height
    End With


End Sub




Private Sub picToolBar_Resize()
    Dim i As Integer
    On Error Resume Next
    With lblTime
        .Top = picToolBar.Top + lblPtime.Top
        .Left = 200 + picToolBar.Width + picToolBar.Left
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    For i = 0 To OptTime.Count - 1
        OptTime(i).Font.Size = mFontSize + mFontSize * mintBigSize / 3
        OptTime(i).Height = 300 + 300 * mintBigSize / 3
        OptTime(i).Width = 350 + 350 * mintBigSize / 3
        OptTime(i).Left = i * OptTime(i).Width + lblPtime.Left + lblPtime.Width + 10
    Next i
    picToolBar.Top = 210
    picToolBar.Width = OptTime(OptTime.Count - 1).Left + OptTime(OptTime.Count - 1).Width
    picToolBar.Height = OptTime(OptTime.Count - 1).Top + OptTime(OptTime.Count - 1).Height
    picToolBar.Left = 0
    picToolBar.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    
End Sub



Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Not mblnInit Then Exit Sub
    
    If Item.Tag = "表格" Then
        picNull.Visible = (vsfTab.TextMatrix(1, COL_tab项目序号) = -999)
        If picEdit.Visible = False Then
            Call SetColSelect(True)
        Else
            txtEdit_KeyPress (vbKeyEscape)
            Call SetColSelect
            
        End If
    ElseIf Item.Tag = "曲线" Then
        picNull.Visible = False
        If mblnStart = False Then
            Call SetColSelect
        Else
            Call SetColSelect(True)
            mblnStart = False
        End If
    End If
End Sub

Private Sub tmrData_Timer()
    Dim i As Integer
    Dim strDay As String
    
    '刷新时点按钮显示状态
    
    If mstrBegin = "" Then Exit Sub
    strDay = Format(mstrBegin, "YYYY-MM-DD")
    
    If Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then mstrBegin = mstrBTime
    If Format(mstrEnd, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then mstrEnd = mstrETime

    If Format(mstrBegin, "YYYY-MM-DD") = Format(mstrBTime, "YYYY-MM-DD") Or Format(mstrEnd, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
        For i = 0 To OptTime.Count - 1
            If OptTime(i).Tag <> "" Then
                If Format(strDay & " " & Split(OptTime(i).Tag, ",")(0), "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Or _
                    Format(strDay & " " & Split(OptTime(i).Tag, ",")(1), "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then
                    OptTime(i).Enabled = False
                Else
                    OptTime(i).Enabled = True
                End If
            End If
        Next i
    Else
        For i = 0 To OptTime.Count - 1
            OptTime(i).Enabled = True
        Next i
    End If
End Sub



Private Sub txtEdit_GotFocus()
    Call zlControl.TxtSelAll(txtEdit)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    Dim strTag As String
    
    If KeyCode = vbKeyDown Then
        If picEdit.Visible = False Then Exit Sub
        '对于类型为文字类型的活动项目使用快捷键可以调用字体颜色设置
        If cmdColor.Visible = True And Shift = vbShiftMask And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(1)) = 1 _
            And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4)) = 0 Then
            With picColor
                .Top = picEdit.Top + picEdit.Height
                If .Top + .Height > vsfTab.Top + vsfTab.Height Then
                    .Top = picEdit.Top - .Height
                End If
                If .Top < vsfTab.Top Then .Top = vsfTab.Top
                .Left = picEdit.Left
                .Visible = True
                .ZOrder 0
            End With
            With usrColor
                .Left = 0
                .Top = -450
                .Visible = True
                .ZOrder 0
            End With
            picColor.SetFocus
            usrColor.Color = Val(cmdColor.Tag)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        '检查数据合法性
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then
                If txtEdit.Tag <> "" Then strTag = txtEdit.Tag
                If Split(txtEdit.Tag, "|")(2) = 1 Then
                    blnAllow = WriteIntoVfgTab(strData, vsfTab)
                Else
                    blnAllow = WriteIntoVfgTab(strData, vsfTabDetail)
                End If
            End If
        End If
        
        If Split(IIf(strTag = "", txtEdit.Tag, strTag), "|")(2) = 1 Then
            If blnAllow = True Then
                '移动到下一列
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTab_EnterCell
            End If
        Else
            If blnAllow = True Then
                '移动到下一列
                Call vsfTabDetail_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTabDetail_EnterCell
            End If
        
        End If
        
    ElseIf KeyCode = vbKeyLeft And txtEdit.SelStart = 0 Then
        If picHour.Visible = False Then
            If Split(txtEdit.Tag, "|")(2) = 1 Then
                Call vsfTab_KeyDown(vbKeyLeft, 0)
            Else
                Call vsfTabDetail_KeyDown(Left, 0)
            End If
        Else
            txtHour.SetFocus
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim strVsf As String
    If KeyAscii = vbKeyEscape Then
        If InStr(txtEdit.Tag, "|") <> 0 Then strVsf = Split(txtEdit.Tag, "|")(2)
        With picEdit
            .Visible = False
            .Enabled = False
        End With
        With txtEdit
            .Visible = False
            .Enabled = False
            .Tag = ""
            .Text = ""
        End With
        With cmdColor
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(0)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(1)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        
        With PicLst
            .Visible = False
            .Tag = ""
        End With
        
        With picHour
            .Visible = False
            .Enabled = False
        End With
        
        With txtHour
            .Visible = False
            .Enabled = False
            .Text = ""
        End With
        
        With lblCheck
            .Visible = False
            .Enabled = False
        End With
        mblnEdit = False
        
        If mblnAllRefresh = False And mblnStart = False Then Call SetColSelect(False, Val(strVsf))
    End If
End Sub

Private Sub txtHour_GotFocus()
    Call zlControl.TxtSelAll(txtHour)
End Sub

Private Sub txtHour_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    
    If picHour.Visible = False Then Exit Sub
    If KeyCode = vbKeyReturn And Not (Shift = vbShiftMask) Then
        '检查数据合法性
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
                If txtEdit.Text = "" Then strData = ""
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If strData & "'" & lngColor <> picEdit.Tag Then blnAllow = WriteIntoVfgTab(strData, vsfTab, False, False)
        End If
        If blnAllow = True Then
            '移动到下一列
            If txtEdit.Enabled = True Then
                txtEdit.SetFocus
            Else
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            End If
        Else
            txtHour.SetFocus
        End If
    ElseIf KeyCode = vbKeyLeft And txtHour.SelStart = 0 Then
        Call vsfTab_KeyDown(vbKeyLeft, 0)
    End If
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    Dim strText As String
    strText = txtHour.Text
    If strText = "" Then Exit Sub
    If Not (Val(strText) >= 0 And strText <= 24) Then
        lblStb.Caption = "汇总小时只能在0到24之间，请重新录入！": lblStb.ForeColor = 255
        Cancel = True
    Else
        txtHour.Text = Val(strText)
    End If
End Sub

Private Sub txtLst_GotFocus()
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    
    blnAllow = True
    If KeyCode = vbKeyReturn And Shift = vbShiftMask Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Trim(txtLst.Text) <> lstSelect(0).Tag Then
            If Val(txtLst.Tag) = 1 Then
                blnAllow = WriteIntoVfgTab(txtLst.Text, vsfTab)
            Else
                blnAllow = WriteIntoVfgTab(txtLst.Text, vsfTabDetail)
            End If
        End If
        If blnAllow = True Then
            If Val(txtLst.Tag) = 1 Then
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            Else
                Call vsfTabDetail_KeyDown(vbKeyReturn, Shift)
            End If
        End If
    ElseIf KeyCode = vbKeyLeft And txtLst.SelStart = 0 Then
        If Val(txtLst.Tag) = 1 Then
                Call vsfTab_KeyDown(vbKeyLeft, 0)
            Else
                Call vsfTabDetail_KeyDown(vbKeyLeft, 0)
            End If
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
         Call txtEdit_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub usrColor_LostFocus()
    picColor.Visible = False
End Sub

Private Sub usrColor_pOK()
    Dim intRow As Integer, intCOl As Integer
    Dim strTmp As String, lng项目序号 As Long, str项目名称 As String
    Dim strTime As String, arrTime() As String
    
    On Error GoTo Errhand
    If Val(cmdColor.Tag) = usrColor.Color Then picColor.Visible = False:  GoTo GetSetFocus
    cmdColor.Tag = usrColor.Color
    txtEdit.ForeColor = cmdColor.Tag
    picColor.Visible = False
    
    If txtEdit.Tag <> "" Then
        intRow = Val(Split(txtEdit.Tag, "|")(0))
        intCOl = Val(Split(txtEdit.Tag, "|")(1))
    Else
        intRow = vsfTab.Row
        intCOl = vsfTab.Col
    End If
    
    lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
    str项目名称 = vsfTab.TextMatrix(intRow, COL_tab项目名)
    If vsfTab.TextMatrix(vsfTab.Row, col_tab原始时间) <> "" Then
        arrTime = Split(vsfTab.TextMatrix(vsfTab.Row, col_tab原始时间), "'")
        If intCOl - vsfTab.FixedCols < UBound(arrTime) Then
            strTime = arrTime(intCOl - vsfTab.FixedCols)
        End If
    End If
    mrsTableDetail.Filter = "项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' and 时间='" & strTime & "'"
    If mrsTableDetail.RecordCount > 0 Then
        mrsTableDetail!未记说明 = cmdColor.Tag
        If mrsTableDetail!状态 <> 1 Then   '原有的数据 修改、删除后的状态始终为2
            mrsTableDetail!状态 = 2
            mrsTableDetail!结果 = vsfTab.TextMatrix(intRow, intCOl)
        Else '对于新增数据的处理
            If Trim(vsfTab.TextMatrix(intRow, intCOl)) = "" Then
                mrsTableDetail.Delete
            Else
                mrsTableDetail!状态 = 1
                mrsTableDetail!结果 = vsfTab.TextMatrix(intRow, intCOl)
            End If
        End If
        mrsTableDetail.Update
    End If
    
GetSetFocus:
    If txtEdit.Visible = True Then txtEdit.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub usrValue_LostFocus()
    picValue.Visible = False
End Sub

Private Sub usrValue_pOK()
    If Val(vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_颜色, usrValue.Tag, COL_颜色)) = usrValue.Color Then picValue.Visible = False: GoTo ErrNext
    vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_颜色, usrValue.Tag, COL_颜色) = usrValue.Color
    If Trim(vsfCurve.TextMatrix(usrValue.Tag, COL_数据)) = "" Then GoTo ErrNext
    If vsfCurve.TextMatrix(usrValue.Tag, COL_修改状态) <> 1 Then vsfCurve.TextMatrix(usrValue.Tag, COL_修改状态) = 2
    If Not UpdateCurveDate(vsfCurve, usrValue.Tag, COL_颜色, 2) Then vsfCurve.Cell(flexcpBackColor, usrValue.Tag, COL_颜色, usrValue.Tag, COL_颜色) = usrValue.Color
ErrNext:
    picValue.Visible = False
    If Val(usrValue.Tag) <= vsfCurve.Rows - 1 Then
        vsfCurve.Body.Select Val(usrValue.Tag), COL_数据
    End If
    vsfCurve.SetFocus
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:表格选中图片
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, COL_TabNull) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_TabNull, objVsf.Rows - 1, COL_TabNull) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_TabNull) = ilstab.ListImages(1).Picture
    
End Sub

Private Sub vsfCurve_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    If Col = COL_数据 Then
        vsfCurve.TextMatrix(Row, COL_数据) = IIf(vsfCurve.EditText = "", " ", Space(Row) & vsfCurve.EditText & Space(Row))
        vsfCurve.TextMatrix(Row, COL_颜色) = IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", Space(Row), vsfCurve.TextMatrix(Row, COL_数据))
    End If
    
End Sub


Private Sub vsfCurve_AfterNextRow(ByVal Row As Long, Col As Long)
    If Col = COL_时间 And Row <> vsfCurve.FixedRows Then
        vsfCurve.TextMatrix(Row, COL_时间) = vsfCurve.TextMatrix(Row - 1, COL_时间)
    End If
End Sub

Private Sub vsfCurve_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng项目序号 As Integer
    Dim strData As String
    Dim strTmp As String
    
    On Err GoTo Errhand
    '加载部位下拉列表
    vsfCurve.ComboList(COL_部位) = ""
    vsfCurve.EditMode(COL_部位) = 0
    vsfCurve.EditMode(Col_未记说明) = 0
    vsfCurve.EditMode(NewCol) = 0

    lng项目序号 = Val(vsfCurve.TextMatrix(NewRow, COL_项目序号))
    strData = Trim(vsfCurve.TextMatrix(NewRow, COL_数据))
    Select Case Trim(vsfCurve.TextMatrix(NewRow, COL_分组名))
        Case "1)体温曲线项目"
            vsfCurve.EditMode(Col_未记说明) = 1
            strTmp = GetAllPart(lng项目序号)
            If strTmp <> "" Then
                If lng项目序号 = 2 And InStr(1, strTmp, "|") = 0 Then
                    strTmp = " |起搏器"
                End If
                vsfCurve.ComboList(COL_部位) = strTmp
                vsfDetail.Body.ColComboList(COL_部位) = strTmp
                vsfCurve.EditMode(COL_部位) = 1
            End If
        
        If NewCol = COL_数据 Or NewCol = Col_未记说明 Or NewCol = COL_时间 Then
            '数据来源
            If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(NewRow, COL_来源)) & ",") = 0 Then
                If NewCol = COL_数据 Then
                    If lng项目序号 = gint体温 And strData = "不升" Then vsfCurve.EditMode(NewCol) = 0
                    If lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                        If InStr(1, strData, "/") = 0 Then
                            vsfCurve.EditMode(NewCol) = 1
                        Else
                            If vsfCurve.TextMatrix(NewRow, COL_编辑) = 0 Then vsfCurve.EditMode(NewCol) = 1
                        End If
                    End If
                End If
            Else
                If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(NewRow, COL_来源)) & ",") = 0 Then
                    vsfCurve.EditMode(NewCol) = 0
                Else
                    vsfCurve.EditMode(NewCol) = 1
                End If
           
            End If
        End If
    
        Case "2)上下标说明"
            vsfCurve.EditMode(Col_未记说明) = 0
            vsfCurve.EditMode(COL_数据) = 1
            vsfCurve.EditMode(COL_时间) = 1
    End Select
    
    strTmp = ""
    If vsfCurve.TextMatrix(NewRow, COL_字符串) <> "" Then
        If Trim(Split(vsfCurve.TextMatrix(NewRow, COL_字符串), ",")(0)) <> "" Then
            strTmp = "数据范围：" & Trim(Split(vsfCurve.TextMatrix(NewRow, COL_字符串), ",")(0)) & " "
        End If
    End If
    
    If Trim(vsfCurve.TextMatrix(NewRow, COL_分组名)) = "1)体温曲线项目" Then
        Select Case lng项目序号
            Case 1 '体温
                strTmp = strTmp & Space(4) & "物理降温表示法38/37"
            Case gint疼痛强度
                strTmp = strTmp & Space(4) & "疼痛减痛表示法6/2"
            Case 2
                If mint心率应用 = 2 And mblnEdit心率 Then strTmp = strTmp & Space(4) & "脉搏短拙表示法100/130"
        End Select
    ElseIf Trim(vsfCurve.TextMatrix(NewRow, COL_分组名)) = "2)上下标说明" Then
        strTmp = "在数据列按SHIFT+↓或双击颜色栏进行颜色设置"
    End If
    lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012
    '加载vsfDetail表格数据
    If OldRow = NewRow And mblnRefresh首行 = True Then Exit Sub
    mblnRefresh首行 = True
    If vsfCurve.TextMatrix(NewRow, COL_分组名) = "1)体温曲线项目" Then
        Call ShowDetail(lng项目序号, NewRow)
    Else
        vsfDetail.Rows = vsfDetail.FixedRows
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetAllPart(ByVal lngNo As Long) As String
    '获取所有部位
    Dim strValue As String
    Dim strTmp As String
    
    On Error GoTo Errhand
    If Not mrsPart Is Nothing Then
        mrsPart.Filter = "项目序号=" & lngNo
        mrsPart.Sort = "缺省项 DESC"
        With mrsPart
            Do While Not .EOF
                strTmp = IIf(strTmp = "", zlStr.Nvl(!部位), strTmp & "|" & zlStr.Nvl(!部位))
            .MoveNext
            Loop
        End With
    End If
    GetAllPart = strTmp
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub vsfCurve_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    If Col = COL_颜色 Then
        lngWidth = vsfCurve.Body.ColWidth(Col)
        vsfCurve.Body.ColWidth(COL_颜色) = 300
        vsfCurve.Body.ColWidth(COL_数据) = vsfCurve.Body.ColWidth(COL_数据) + lngWidth - 300
        If vsfCurve.Body.ColWidth(COL_数据) < 500 Then vsfCurve.Body.ColWidth(COL_数据) = 500
        Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
    End If
End Sub

Private Sub vsfCurve_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnAllow As Boolean
    Dim intType As Integer
    Dim lng项目序号 As Long
    
    On Err GoTo Errhand
    vsfCurve.Tag = vsfCurve.TextMatrix(Row, Col)

    If VsfDeleteRow(1, vsfCurve, Row, Col, Cancel) Then Exit Sub
    Call ShowCurve
    Call ShowTabUpDown
    Cancel = True
    lng项目序号 = vsfCurve.TextMatrix(Row, COL_项目序号)
    Call ShowDetail(lng项目序号, Row)
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function VsfDeleteRow(ByVal intType As Integer, ByVal vsf As Object, ByVal Row As Integer, ByVal Col As Integer, ByRef Cancel As Boolean) As Boolean
    '-----------------------------------------
    '功能：删除行
    '-----------------------------------------
    Dim blnAllow As Boolean
    
    On Error GoTo Errhand
    Select Case Col
        Case COL_时间
            vsf.TextMatrix(Row, COL_修改状态) = 2
            vsf.TextMatrix(Row, Col) = ""
            If intType = 3 Then
                intType = 3
            ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
                intType = 2
            ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
                intType = 1
            End If
            blnAllow = True
        Case COL_部位
            vsf.TextMatrix(Row, COL_修改状态) = 2
            vsf.TextMatrix(Row, Col) = ""
            If intType = 3 Then
                intType = 3
            ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
                intType = 2
            ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
                intType = 1
            End If
            blnAllow = True
        Case COL_数据
            If vsf.TextMatrix(Row, Col) <> "" Then
                If intType = 3 Then
                    intType = 3
                    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_来源)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能删除."
                        lblStb.ForeColor = 255
                        vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        VsfDeleteRow = True
                        Exit Function
                    End If
                ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
                    intType = 2
                ElseIf Trim(vsf.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
                    intType = 1
                    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_来源)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能删除."
                        lblStb.ForeColor = 255
                        vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        VsfDeleteRow = True
                        Exit Function
                    End If
                End If
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
        Case Col_未记说明
            If Trim(vsf.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" And vsf.TextMatrix(Row, Col) <> "" Then
                intType = 1
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
        Case COL_删除
            If vsf.TextMatrix(Row, COL_数据) <> "" Or vsf.TextMatrix(Row, Col_未记说明) <> "" Then
                intType = 3
                If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_来源)) & ",") = 0 Then
                    Cancel = True
                    lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能删除."
                    lblStb.ForeColor = 255
                    vsf.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                    VsfDeleteRow = True
                    Exit Function
                End If
                If CurveRowClear(intType, Row) Then blnAllow = True
            End If
            
    End Select
    If blnAllow = True Then Call UpdateCurveDate(vsf, Row, Col, intType)
    
     Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CurveRowClear(ByVal intType As Integer, ByVal intRow As Integer) As Boolean
    '-------------------------------
    '清空曲线表格数据
    '-------------------------------
    On Error Resume Next
    Select Case intType
        Case 1, 2
            vsfCurve.TextMatrix(intRow, COL_修改状态) = IIf(Val(vsfCurve.TextMatrix(intRow, COL_修改状态)) = 1, 4, 3)
            vsfCurve.TextMatrix(intRow, COL_时间) = ""
            vsfCurve.TextMatrix(intRow, COL_数据) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", "", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, COL_颜色) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", " ", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, COL_部位) = ""
            vsfCurve.TextMatrix(intRow, COL_复试合格) = ""
            vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
            CurveRowClear = True
        Case 3
            vsfDetail.TextMatrix(intRow, COL_修改状态) = IIf(Val(vsfDetail.TextMatrix(intRow, COL_修改状态)) = 1, 4, 3)
            vsfDetail.TextMatrix(intRow, COL_显示) = ""
            vsfDetail.TextMatrix(intRow, COL_时间) = ""
            vsfDetail.TextMatrix(intRow, COL_数据) = ""
            vsfDetail.TextMatrix(intRow, COL_部位) = ""
            vsfDetail.TextMatrix(intRow, COL_复试合格) = ""
            vsfDetail.TextMatrix(intRow, Col_未记说明) = ""
            vsfDetail.TextMatrix(intRow, COL_来源) = ""
            CurveRowClear = True
    End Select
End Function


Private Sub vsfCurve_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfCurve_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrstr未记说明() As String
    Dim i As Integer
    Dim blnSelect As Boolean

    On Error GoTo Errhand
    If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) <> "1)体温曲线项目" Then Exit Sub
    lst未记.Tag = "1|" & Row & "|" & Col
    lst未记.Clear
    If mstr未记说明 <> "" Then
        arrstr未记说明() = Split(mstr未记说明, "'")
        For i = 0 To UBound(arrstr未记说明)
            lst未记.AddItem arrstr未记说明(i)
            If arrstr未记说明(i) = vsfCurve.TextMatrix(vsfCurve.Row, vsfCurve.Col) Then
                lst未记.Selected(i) = True
                blnSelect = True
            End If
        Next
    End If
    If blnSelect = False And lst未记.ListCount <> 0 Then lst未记.Selected(0) = True
    
    If lst未记.ListCount > 0 Then
        pic未记.Left = vsfCurve.CellLeft + vsfCurve.Left + 15
        pic未记.Top = fraData.Top + vsfCurve.CellTop + vsfCurve.Top + vsfCurve.CellHeight
        If lst未记.Height < vsfCurve.CellHeight + 20 Then lst未记.Height = vsfCurve.CellHeight + 20
        lst未记.Width = vsfCurve.CellWidth + 20
        pic未记.Height = lst未记.Height
        pic未记.Width = lst未记.Width
        pic未记.Visible = True
        lst未记.Visible = True: lst未记.Enabled = True
        lst未记.SetFocus
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfCurve_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
        Call UpdateCurveDate(vsfCurve, Row, Col, 1, True)
    End If
End Sub

Private Sub vsfCurve_KeyDown(KeyCode As Integer, Shift As Integer)
    picValue.Visible = False
    picValue.Tag = ""
    With vsfCurve
        If .Col > .FixedCols - 1 And .Row > .FixedRows - 1 Then
            If KeyCode = vbKeyDown And Shift = vbShiftMask Then
                If .Col = Col_未记说明 Then
                    Call vsfCurve_CellButtonClick(.Row, .Col)
                ElseIf (.Col = COL_数据 Or .Col = COL_颜色) And .TextMatrix(.Row, COL_分组名) = "2)上下标说明" Then
                    vsfCurve.Tag = .TextMatrix(.Row, COL_数据)
                    picValue.Top = fraData.Top + .CellTop + .CellHeight + .Top
                    If picValue.Top + picValue.Height > .Top + .Height Then
                        picValue.Top = .CellTop - picValue.Height
                    End If
                    If picValue.Top < .Top Then picValue.Top = .Top
                    picValue.Left = IIf(.Col = COL_颜色, .CellLeft, .CellLeft + .CellWidth) + .Left
                    picValue.Visible = True
                    picValue.ZOrder 0
         
                    usrValue.Left = 0
                    usrValue.Top = -450
                    usrValue.Visible = True
                    usrValue.ZOrder 0
                    picValue.SetFocus
                    usrValue.Color = Val(.Cell(flexcpBackColor, .Row, COL_颜色, .Row, COL_颜色))
                    picValue.Tag = Val(usrValue.Color)
                    usrValue.Tag = .Row
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfCurve_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_未记说明 Then
            If InStr(1, "," & mstr未记说明 & ",", "," & vsfCurve.EditText & ",") = 0 Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, COL_时间) = ""
                vsfCurve.TextMatrix(Row, COL_数据) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_颜色) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_部位) = ""
                vsfCurve.TextMatrix(Row, COL_复试合格) = ""
            End If
        End If
    End If
    If KeyCode = vbKeyDown And Shift = vbShiftMask And Col = COL_数据 Then
        Call vsfCurve_KeyDown(KeyCode, Shift)
        Cancel = True
    End If
End Sub

Private Sub vsfCurve_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = 32 Then '
        If Col = COL_复试合格 Then
            If Val(vsfCurve.TextMatrix(Row, COL_数据)) <> 0 And Val(vsfCurve.TextMatrix(Row, COL_项目序号)) = gint体温 Then
                If vsfCurve.TextMatrix(Row, COL_修改状态) = 1 Then vsfCurve.TextMatrix(Row, COL_修改状态) = 1
                If vsfCurve.TextMatrix(Row, COL_修改状态) = 0 Then vsfCurve.TextMatrix(Row, COL_修改状态) = 2
                If vsfCurve.TextMatrix(Row, Col) = "" Then
                    vsfCurve.TextMatrix(Row, Col) = "√"
                    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                Else
                    vsfCurve.TextMatrix(Row, Col) = ""
                End If
                Call UpdateCurveDate(vsfCurve, Row, Col, 1)
                Call ShowDetail(gint体温, Row)
            End If
        End If
        If Col = COL_颜色 And vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明" Then
            Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
        End If
    End If
End Sub

Private Sub vsfCurve_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNo As Long
    
    On Error Resume Next
    lngNo = Val(vsfCurve.TextMatrix(Row, COL_项目序号))
    
    If KeyAscii <> vbKeyReturn Then
        If lngNo <> 0 Then
            If Col = COL_时间 Then
                 If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
            Else
                If vsfCurve.TextMatrix(Row, COL_分组名) = "1)体温曲线项目" Then
                    If Col <> Col_未记说明 Then
                        If lngNo = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint疼痛强度 Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint体温 Then
                            '体温不进行检查
                        Else
                            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                        End If
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub vsfCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng项目序号 As Long
    Dim strName As String
    Dim strData As String
    Dim strTime As String
    
    On Err GoTo Errhand
    lng项目序号 = Val(vsfCurve.TextMatrix(Row, COL_项目序号))
    strName = vsfCurve.TextMatrix(Row, COL_项目名)
    Select Case Col
        Case COL_数据
            vsfCurve.TextMatrix(Row, Col) = IIf(Trim(vsfCurve.TextMatrix(Row, Col)) = "", " ", Trim(vsfCurve.TextMatrix(Row, Col)))
            If Row <> mOptRow.上标 And Row <> mOptRow.下标 Then
                vsfCurve.TextMatrix(Row, COL_颜色) = vsfCurve.TextMatrix(Row, Col)
            Else
                vsfCurve.TextMatrix(Row, Col) = Trim(vsfCurve.TextMatrix(Row, Col))
            End If
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
            strData = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
        Case COL_时间
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
    End Select
    
    vsfCurve.Tag = Trim(vsfCurve.TextMatrix(Row, Col))
    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    If Col = COL_数据 Or Col = Col_未记说明 Or Col = COL_时间 Then
          '数据来源
        If InStr(1, ",0,3,9,", "," & Val(vsfCurve.TextMatrix(Row, COL_来源)) & ",") = 0 Then
            If Col = COL_数据 Then
                If lng项目序号 = gint体温 And strData = "不升" Then GoTo NotEdit
                If lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        GoTo GONext
                    Else
                        If Val(vsfCurve.TextMatrix(Row, COL_编辑)) = 0 Then GoTo GONext
                    End If
                End If
            End If
NotEdit:
            Cancel = True
            vsfCurve.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
            vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
            If lng项目序号 = gint体温 Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
            ElseIf lng项目序号 = gint疼痛强度 Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改疼痛减痛部分."
            ElseIf lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                If mbln脉搏共用显示 Then
                    lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.心率/脉搏"
                Else
                    lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.脉搏/心率 "
                End If
            Else
                lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能修改"
            End If
            lblStb.ForeColor = 255
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf COL_复试合格 = Col Then
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GONext:
    If mblnFileBack = True Then
        Cancel = True
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'Private Function StartEdit()
'    '编辑之前数据价
'
'
'
'
'End Function


Private Sub vsfCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng项目序号 As Long
    On Err GoTo Errhand
    If CheckPutData(1, vsfCurve, Row, Col, Cancel) Then
        lng项目序号 = vsfCurve.TextMatrix(Row, COL_项目序号)
        Call ShowDetail(lng项目序号, Row)
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CheckPutData(ByVal intType As Integer, ByVal vsf As Object, ByVal Row As Integer, ByVal Col As Integer, ByRef Cancel As Boolean) As Boolean
    '---------------------------------------
    '功能：检查输入值，并更新
    '---------------------------------------
    Dim strTime As String
    Dim strMsg As String
    Dim strText As String, strData As String
    Dim strCenterTime As String
    Dim strName As String, strValue As String, str值域 As String, strInfo As String
    Dim int小数 As Integer
    Dim i As Integer, intCount As Integer
    Dim lng项目序号 As Long
    Dim blnOK As Boolean
    Dim lngCount As Long
    Dim arrValue() As String

    On Err GoTo Errhand
    '检查数据合法性
    
    strValue = vsf.Tag
    str值域 = Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_字符串), ",")(0)
    lng项目序号 = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号))
    strName = vsfCurve.TextMatrix(vsfCurve.Row, COL_项目名)
    int小数 = Val(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_字符串), ",")(2))
    
    If vsf.Col = COL_时间 Then
        strText = vsf.EditText
        strText = CToData(strText)
        If strText <> "" Then
            If Not CheckDateTime(Row, strName, strText) Then
                strMsg = lblStb.Caption
                GoTo ErrInfo
            End If
            vsf.EditText = strText
            vsf.TextMatrix(Row, COL_修改状态) = 5
            Select Case vsf.TextMatrix(Row, COL_分组名)
                Case "1)体温曲线项目"
                    intType = 1
                Case "2)上下标说明"
                    intType = 2
                Case Else
                    intType = 3
            End Select
            mrsCurve.Filter = "项目序号=" & lng项目序号 & " and  时间='" & Format(dtpDate.Value & " " & strText, "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsCurve.RecordCount > 0 And strValue <> strText Then
                strMsg = "当前时间已存在数据，请重新输入时间"
                GoTo ErrInfo
            End If
            Call UpdateCurveDate(vsf, Row, Col, intType)
            CheckPutData = True
        End If
    End If
    
    If Col = COL_数据 Then
    
        Select Case vsf.TextMatrix(Row, COL_分组名)
            Case "1)体温曲线项目"
                intType = 1
                GoTo CheckPoint
            Case "2)上下标说明"
                If InStr(1, ",2,6,", "," & Val(vsf.TextMatrix(Row, COL_项目序号)) & ",") <> 0 Then
                    picValue.Tag = vsf.Cell(flexcpBackColor, Row, COL_颜色, Row, COL_颜色)
                    intType = 2: GoTo CheckTag
                End If
            Case Else
                intType = 3
                GoTo CheckPoint
        End Select
    End If
    
    Exit Function
    
CheckPoint:
    '检查数据
    If Trim(vsf.EditText) <> "" And str值域 <> "" Then
        strInfo = vsf.EditText
        If vsf.TextMatrix(Row, COL_时间) = "" Then
        mrsCurve.Filter = "项目序号=" & lng项目序号 & " and  时间='" & Format(GetCenterTime(mstrBegin, mstrEnd), "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsCurve.RecordCount > 0 Then
                strMsg = "当前默认时间已存在数据，请先输入时间"
                GoTo ErrInfo
            End If
        End If
        '脉搏短轴是如果有/则要求必须输入心率
        If lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
            If InStr(1, strInfo, "/") > 0 Then
                If Split(Trim(strInfo), "/")(1) = "" Or Split(Trim(strInfo), "/")(0) = "" Then
                    strMsg = strName & "数据录入错误" & Space(4) & "脉搏短轴:脉搏/心率"
                    GoTo ErrInfo
                Else
                    If Not IsNumeric(Split(Trim(strInfo), "/")(0)) Or Not IsNumeric(Split(Trim(strInfo), "/")(1)) Then
                        strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                        GoTo ErrInfo
                    End If
                End If
            End If
        End If
        
        If lng项目序号 <> 1 And lng项目序号 <> gint疼痛强度 And Not (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
            If InStr(1, strInfo, "/") Then
                strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                GoTo ErrInfo
            End If
        End If

        If UBound(Split(strInfo, "/")) > 1 Then
            strMsg = strName & "数据录入错误，请检查！"
            GoTo ErrInfo
        End If
        
        '检查数据在有效范围内是否有效
        arrValue = Split(strInfo, "/")
        lngCount = UBound(arrValue)
        For i = 0 To lngCount
            blnOK = False
            strText = arrValue(i)
            If i = 0 Then
                '体温曲线项目需要过滤掉未记说明
                If InStr(1, strText, ";") <> 0 And UBound(arrValue) = 0 Then strText = Split(strText, ";")(1)
                If InStr(1, IIf(lng项目序号 = gint体温, ",不升,", ""), "," & strText & ",") = 0 Then
                    blnOK = False
                Else
                    blnOK = True
                End If
            End If
            
            If Not blnOK Then
                If Not IsNumeric(strText) Then
                    strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                    GoTo ErrInfo
                End If
            End If
            
            If Not blnOK And strText <> "" Then
                strText = Format(Val(strText), "#0" & IIf(int小数 > 0, ".", "") & String(int小数, "0"))
                If strText = Val(strText) Then strText = Val(strText)
                If Left(strText, 1) = "." Then strText = 0 & strText
            End If
            If IsNumeric(Split(str值域, "～")(0)) And IsNumeric(strText) Then
                If Not (Val(strText) >= Split(str值域, "～")(0) And Val(strText) <= Split(str值域, "～")(1)) Then
                    strMsg = strName & "超出有效范围(" & str值域 & "),请检查!"
                    GoTo ErrInfo
                End If
            End If
            If i = 0 Then
                strInfo = strText
            Else
                strInfo = strInfo & "/" & strText
            End If
        Next i
    End If
    If strInfo <> vsf.EditText Then vsf.EditText = strInfo
    strData = vsf.EditText
    '对于数据来源<>0,3,9的 体温,脉搏数据 进行编辑(无物理降温和脉搏短轴可以录入物理降温,脉搏短轴)
    If InStr(1, ",0,3,9,", "," & Val(vsf.TextMatrix(Row, COL_来源)) & ",") = 0 Then
        If Col = COL_数据 Then
            If lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                '检查脉搏短轴录入方式是否正确，心率/脉搏
                If lng项目序号 = 2 And (InStr(strValue, "/") > 0 Or InStr(strValue, "/") = 0) And mbln脉搏共用显示 Then
                    If InStr(1, strData, "/") <> 0 Then
                        strData = Split(strData, "/")(1)
                    Else
                        strData = strData
                    End If
                    If strData <> vsf.TextMatrix(Row, COL_原值) Then
                        If mbln脉搏共用显示 Then
                            strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.心率/脉搏"
                        Else
                            strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.脉搏/心率 "
                        End If
                        vsf.TextMatrix(Row, COL_数据) = Space(Row) & Trim(CStr(vsf.TextMatrix(Row, COL_原值))) & Space(Row)
                        vsf.TextMatrix(Row, COL_颜色) = vsf.TextMatrix(Row, COL_数据)
                        GoTo ErrInfo
                    End If
                Else
                    strValue = CStr(vsf.TextMatrix(Row, COL_原值))
                    If InStr(1, strData, "/") <> 0 Then
                        strData = Split(strData, "/")(0)
                    End If
                
                    If InStr(1, vsf.TextMatrix(Row, COL_原值), "/") = 0 Then
                        If strData <> vsf.TextMatrix(Row, COL_原值) Then
                            If lng项目序号 = gint体温 Then
                                strMsg = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                            ElseIf lng项目序号 = gint疼痛强度 Then
                                strMsg = "同步过来的[" & strName & "]数据只允许修改疼痛减痛部分."
                            Else
                                If mbln脉搏共用显示 Then
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.心率/脉搏"
                                Else
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.脉搏/心率 "
                                End If
                            End If
                            
                            vsf.TextMatrix(Row, COL_数据) = Space(Row) & Trim(CStr(vsf.TextMatrix(Row, COL_原值))) & Space(Row)
                            vsf.TextMatrix(Row, COL_颜色) = vsf.TextMatrix(Row, COL_数据)
                            GoTo ErrInfo
                        End If
                    Else
                        If Val(vsf.TextMatrix(Row, COL_编辑)) <> 0 Then
                            If strData <> vsf.TextMatrix(Row, COL_原值) Then
                                If lng项目序号 = gint体温 Then
                                    strMsg = "同步过来的[" & strName & "]数据如果包括物理降温,不允许修改."
                                ElseIf lng项目序号 = gint疼痛强度 Then
                                    strMsg = "同步过来的[" & strName & "]数据如果包括疼痛减痛,不允许修改."
                                Else
                                    strMsg = "同步过来的[" & strName & "]数据如果包括脉搏短轴,不允许修改."
                                End If
                                vsf.TextMatrix(Row, COL_数据) = Space(Row) & CStr(vsf.TextMatrix(Row, COL_原值)) & Space(Row)
                                vsf.TextMatrix(Row, COL_颜色) = vsf.TextMatrix(Row, COL_数据)
                                GoTo ErrInfo
                            End If
                        Else
                            If strData <> Split(vsf.TextMatrix(Row, COL_原值), "/")(0) Then
                                If lng项目序号 = gint体温 Then
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                                ElseIf lng项目序号 = gint疼痛强度 Then
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改疼痛减痛部分."
                                Else
                                    If mbln脉搏共用显示 Then
                                        strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.心率/脉搏"
                                    Else
                                        strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.脉搏/心率 "
                                    End If
                                End If
                                vsf.TextMatrix(Row, COL_数据) = Space(Row) & CStr(vsf.TextMatrix(Row, COL_原值)) & Space(Row)
                                vsf.TextMatrix(Row, COL_颜色) = vsf.TextMatrix(Row, COL_数据)
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    '显示缺省部位
    If vsf.TextMatrix(Row, COL_部位) = "" And Trim(vsf.EditText) <> "" Then
        mrsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
        If mrsPart.RecordCount > 0 Then
            vsf.TextMatrix(Row, COL_部位) = CStr(zlStr.Nvl(mrsPart!部位))
        End If
    End If
    
    GoTo UpdateData
    Exit Function
    
CheckTag:
    GoTo UpdateData
    Exit Function
    
ErrInfo:
    lblStb.Caption = strMsg
    lblStb.ForeColor = 255
    vsf.TextMatrix(Row, COL_数据) = Space(Row) & strValue & Space(Row)
    vsf.TextMatrix(Row, COL_颜色) = vsf.TextMatrix(Row, COL_数据)
    Cancel = True
    Exit Function
    
UpdateData:
    
    If vsf.EditText = strValue Then Exit Function
    intCount = 0
    For i = COL_数据 To Col_未记说明
        If Trim(vsf.TextMatrix(Row, i)) <> "" Or vsf.EditText <> "" Then
           intCount = intCount + 1
           Exit For
        End If
    Next
    If intCount = 0 Then
        If Trim(vsf.TextMatrix(Row, COL_修改状态)) = 1 Then
            vsf.TextMatrix(Row, COL_修改状态) = 4
        Else
            vsf.TextMatrix(Row, COL_修改状态) = 3
        End If
    Else
        vsf.TextMatrix(Row, COL_修改状态) = 2
    End If
    Call UpdateCurveDate(vsf, Row, Col, intType)
    CheckPutData = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CToData(ByVal strData As String) As String
'-----------------------------------------------
'功能：检查时间，转换为时间：分钟
'-----------------------------------------------
    Dim strCenterTime As String
    If InStr(1, Trim(strData), ":") = 0 And strData <> "" Then
        Select Case Len(strData)
        Case 3, 4
            strData = String(4 - Len(strData), "0") & strData
            strData = Mid(strData, 1, 2) & ":" & Mid(strData, 3)
        Case Is < 3
            strData = String(2 - Len(strData), "0") & strData
            strCenterTime = GetCenterTime(mstrBegin, mstrEnd)
            strData = Format(strCenterTime, "HH") & ":" & strData
        End Select
    End If
    CToData = strData
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
End Sub

Private Sub vsfDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intRow As Integer
    Dim strTemp As String
    Dim lng项目序号 As String
    Dim strData As String
    Dim strTmp As String
    Dim strName As String
    On Error GoTo Errhand
    
    If vsfDetail.Cols < COL_删除 Or NewRow < vsfDetail.FixedRows Then Exit Sub
    If OldRow < vsfDetail.Rows Then Call vsfDetail.SelectRow(vsfDetail, OldRow, NewRow, &HFFC0C0)
    For intRow = vsfDetail.FixedRows To vsfDetail.Rows - 2
        vsfDetail.Body.Cell(flexcpPicture, intRow, COL_删除, NewRow, COL_删除) = Nothing
    Next
    If NewRow > vsfDetail.FixedRows - 1 And NewRow < vsfDetail.Rows - 1 Then
        vsfDetail.Body.Cell(flexcpPicture, NewRow, COL_删除, NewRow, COL_删除) = ilsDetail.ListImages(1).Picture
    End If
    vsfDetail.EditMode(NewCol) = 0
    
    lng项目序号 = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号))
    strName = Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目名))
    strData = Trim(vsfDetail.TextMatrix(NewRow, COL_数据))
    vsfDetail.EditMode(Col_未记说明) = 1
    vsfDetail.EditMode(COL_部位) = 1
    
    If NewCol = COL_数据 Or NewCol = Col_未记说明 Or NewCol = COL_时间 Then
        '数据来源
        If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(NewRow, COL_来源)) & ",") = 0 Then
            If NewCol = COL_数据 Then
                If lng项目序号 = gint体温 And strData = "不升" Then vsfDetail.EditMode(NewCol) = 0
                If lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        vsfDetail.EditMode(NewCol) = 1
                    Else
                        If Val(vsfDetail.TextMatrix(NewRow, COL_编辑)) = 0 Then vsfDetail.EditMode(NewCol) = 1
                    End If
                End If
            End If
        Else
            If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(NewRow, COL_来源)) & ",") = 0 Then
                vsfDetail.EditMode(NewCol) = 0
            Else
                vsfDetail.EditMode(NewCol) = 1
            End If
       
        End If
    End If
    
    strTmp = ""
    If vsfCurve.TextMatrix(vsfCurve.Row, COL_字符串) <> "" Then
        If Trim(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_字符串), ",")(0)) <> "" Then
            strTmp = "数据范围：" & Trim(Split(vsfCurve.TextMatrix(vsfCurve.Row, COL_字符串), ",")(0)) & " "
        End If
    End If
    
    If Trim(vsfDetail.TextMatrix(NewRow, COL_分组名)) = "1)体温曲线项目" Then
        Select Case lng项目序号
            Case 1 '体温
                strTmp = strTmp & Space(4) & "物理降温表示法38/37"
            Case gint疼痛强度
                strTmp = strTmp & Space(4) & "疼痛减痛表示法6/2"
            Case 2
                If mint心率应用 = 2 And mblnEdit心率 Then strTmp = strTmp & Space(4) & "脉搏短拙表示法100/130"
        End Select
    End If
    If strTmp <> "" Then lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012
    Exit Sub
    
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub




Private Sub vsfDetail_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnAllow As Boolean
    Dim intType As Integer
    Dim lng项目序号 As Long
    
    If Not vsfDetail.TextMatrix(Row, COL_时间) <> "" And (vsfDetail.TextMatrix(Row, COL_数据) <> "" Or vsfDetail.TextMatrix(Row, Col_未记说明) <> "") Then Exit Sub
    vsfDetail.Tag = vsfDetail.TextMatrix(Row, Col)
    If VsfDeleteRow(3, vsfDetail, Row, Col, Cancel) Then Exit Sub
    Call ShowCurve
    Cancel = False
End Sub



Private Sub vsfDetail_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    If vsfDetail.TextMatrix(Row, COL_数据) = "" And vsfDetail.TextMatrix(Row, Col_未记说明) = "" Then Cancel = True
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrstr未记说明() As String
    Dim i As Integer
    Dim blnSelect As Boolean

    On Error GoTo Errhand
    
    lst未记.Tag = "2|" & Row & "|" & Col
    lst未记.Clear
    If mstr未记说明 <> "" Then
        arrstr未记说明() = Split(mstr未记说明, "'")
        For i = 0 To UBound(arrstr未记说明)
            lst未记.AddItem arrstr未记说明(i)
            If arrstr未记说明(i) = vsfDetail.TextMatrix(vsfDetail.Row, Col_未记说明) Then
                lst未记.Selected(i) = True
                blnSelect = True
            End If
        Next
    End If
    If blnSelect = False And lst未记.ListCount <> 0 Then lst未记.Selected(0) = True
    
    If lst未记.ListCount > 0 Then
                pic未记.Left = vsfDetail.CellLeft + vsfDetail.Left + 15
                pic未记.Top = fraDetail.Top + vsfDetail.CellTop + vsfDetail.Top + vsfDetail.CellHeight
                If lst未记.Height < vsfDetail.CellHeight + 20 Then lst未记.Height = vsfDetail.CellHeight + 20
                If pic未记.Top + pic未记.Height > picCurve.Height Then pic未记.Top = fraDetail.Top + vsfDetail.Body.CellTop + vsfDetail.Top - lst未记.Height + 20
                lst未记.Width = vsfDetail.Body.CellWidth + 20
                pic未记.Height = lst未记.Height
                pic未记.Width = lst未记.Width
                pic未记.Visible = True
                lst未记.Visible = True: lst未记.Enabled = True
                lst未记.SetFocus
            End If
    Exit Sub

Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfDetail_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    Call UpdateCurveDate(vsfDetail, Row, Col, 3, True)
End Sub

Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsfDetail.Col > vsfDetail.FixedCols - 1 And vsfDetail.Row > vsfDetail.FixedRows - 1 Then
        If KeyCode = vbKeyDown And Shift = vbShiftMask Then
            If vsfDetail.Col = Col_未记说明 Then
                Call vsfDetail_CellButtonClick(vsfDetail.Row, vsfDetail.Col)
            End If
        End If
    End If
End Sub


Private Sub vsfDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_未记说明 Then
            If InStr(1, "," & mstr未记说明 & ",", "," & vsfDetail.EditText & ",") = 0 Then
                vsfDetail.TextMatrix(Row, Col) = ""
                vsfDetail.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, COL_显示) = ""
                vsfCurve.TextMatrix(Row, COL_时间) = ""
                vsfCurve.TextMatrix(Row, COL_数据) = ""
                vsfCurve.TextMatrix(Row, COL_部位) = ""
                vsfCurve.TextMatrix(Row, COL_复试合格) = ""
                vsfCurve.TextMatrix(Row, COL_来源) = ""
            End If
        End If
    End If
End Sub

Private Sub vsfDetail_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    Dim intRow As Integer
    
    If Col = COL_删除 Then
        Call vsfDetail_BeforeDeleteRow(Row, Col, Cancel)
        If Cancel = False Then vsfDetail.RemoveItem (Row)
    End If
    
    If KeyAscii = 32 Then '
        Select Case Col
            Case COL_复试合格
                If Trim(vsfDetail.TextMatrix(Row, COL_数据)) <> "" And Val(vsfDetail.TextMatrix(Row, COL_项目序号)) = gint体温 Then
                    If vsfDetail.TextMatrix(Row, Col) = "" Then
                        vsfDetail.TextMatrix(Row, Col) = "√"
                        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                    Else
                        vsfDetail.TextMatrix(Row, Col) = ""
                    End If
                    If vsfDetail.TextMatrix(Row, COL_修改状态) = 1 Then vsfDetail.TextMatrix(Row, COL_修改状态) = 1
                    If vsfDetail.TextMatrix(Row, COL_修改状态) = 0 Then vsfDetail.TextMatrix(Row, COL_修改状态) = 2
                    Call UpdateCurveDate(vsfDetail, Row, Col, 3)
                    Call ShowCurve
                    Call ShowTabUpDown
                End If
            Case COL_显示
                If Trim(vsfDetail.TextMatrix(Row, COL_数据)) <> "" Or Trim(vsfDetail.TextMatrix(Row, Col_未记说明)) <> "" Then
                
                    For intRow = vsfDetail.FixedRows To vsfDetail.Rows - 2
                        If vsfCurve.TextMatrix(vsfCurve.Row, col_原始时间) = vsfDetail.TextMatrix(intRow, col_原始时间) Then
                            vsfDetail.TextMatrix(intRow, COL_修改状态) = 6
                            Call UpdateCurveDate(vsfDetail, intRow, Col, 3)
                        End If
                        If vsfDetail.TextMatrix(intRow, COL_显示) = "√" Then
                            vsfDetail.TextMatrix(intRow, COL_显示) = ""
                            vsfDetail.TextMatrix(intRow, COL_修改状态) = 6
                            Call UpdateCurveDate(vsfDetail, intRow, Col, 3)
                            Exit For
                        End If
                    Next
                    
                    vsfDetail.TextMatrix(Row, COL_修改状态) = IIf(vsfDetail.TextMatrix(Row, COL_修改状态) = 0, 6, 2)
                    If intRow <> Row Then
                        vsfDetail.TextMatrix(Row, COL_显示) = "√"
                        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                    Else
                        vsfDetail.TextMatrix(Row, Col) = ""
                    End If
                    
                    Call UpdateCurveDate(vsfDetail, Row, Col, 3)
                    Call ShowCurve
                    Call ShowTabUpDown
                End If
               
        End Select
    End If
End Sub

Private Sub vsfDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNo As Long
    
    On Error Resume Next
    lngNo = Val(vsfDetail.TextMatrix(Row, COL_项目序号))
    
    If KeyAscii <> vbKeyReturn Then
        If lngNo <> 0 Then
            If Col = COL_时间 Then
                 If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
            Else
                If vsfCurve.TextMatrix(Row, COL_分组名) = "1)体温曲线项目" Then
                    If Col <> Col_未记说明 Then
                        If lngNo = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint疼痛强度 Then
                            If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                        ElseIf lngNo = gint体温 Then
                            '体温不进行检查
                        Else
                            If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                        End If
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub vsfDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng项目序号 As Long
    Dim strName As String
    Dim strData As String
    Dim strTime As String
    
    On Err GoTo Errhand
    lng项目序号 = Val(vsfDetail.TextMatrix(Row, COL_项目序号))
    strName = vsfCurve.TextMatrix(vsfCurve.Row, COL_项目名)
    Select Case Col
        Case COL_数据
            vsfDetail.TextMatrix(Row, Col) = Trim(vsfDetail.TextMatrix(Row, Col))
            strData = RTrim(LTrim(vsfDetail.TextMatrix(Row, Col)))
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
        Case COL_时间
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
    End Select
    vsfDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfDetail.Tag = Trim(vsfDetail.TextMatrix(Row, Col))
    If Col = COL_数据 Or Col = Col_未记说明 Or Col = COL_时间 Then
          '数据来源
        If InStr(1, ",0,3,9,", "," & Val(vsfDetail.TextMatrix(Row, COL_来源)) & ",") = 0 Then
            If Col = COL_数据 Then
                If lng项目序号 = gint体温 And strData = "不升" Then GoTo NotEdit
                If lng项目序号 = gint体温 Or lng项目序号 = gint疼痛强度 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                    If InStr(1, strData, "/") = 0 Then
                        GoTo GONext
                    Else
                        If Val(vsfDetail.TextMatrix(Row, COL_编辑)) = 0 Then GoTo GONext
                    End If
                End If
            End If
NotEdit:
            Cancel = True
            vsfDetail.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
            vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
            If lng项目序号 = gint体温 Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
            ElseIf lng项目序号 = gint疼痛强度 Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改疼痛减痛部分."
            ElseIf lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                If mbln脉搏共用显示 Then
                    lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.心率/脉搏"
                Else
                    lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分.脉搏/心率 "
                End If
            Else
                lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能修改"
            End If
            lblStb.ForeColor = 255
            vsfDetail.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf COL_复试合格 = Col Then
        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GONext:
    If mblnFileBack = True Then
        Cancel = True
        vsfDetail.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        vsfDetail.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    On Err GoTo Errhand
    If CheckPutData(3, vsfDetail, Row, Col, Cancel) Then
        Call ShowCurve
        Call ShowDetail(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号), Row)
        vsfDetail.Row = Row
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strData As String, lngType As Long
    Dim strTime As String, strInfo As String
    Dim lngNo As Long, strName As String, strTmp As String, str值域 As String
    Dim strChildNO As String
    Dim lngChildNO As Long
    Dim arrChildNo() As String
    Dim arrTime() As String
    Dim arrStr() As String
    Dim lng项目序号 As Long
    Dim intRow As Integer
    Dim i As Integer
    Dim cbrControl As Object, cbrChild As Object
    Dim rsTotle As New ADODB.Recordset
    Dim blnCheck As Boolean
    Dim strText As String
    
    On Error GoTo Errhand
    If mblnInit = False Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    If NewRow < vsfTab.FixedRows Or NewCol < vsfTab.FixedCols Or NewCol > (Split(vsfTab.RowData(NewRow), ";")(0) + vsfTab.FixedCols - 1) Then Exit Sub
    Call AdjustRowFlag(vsfTab, NewRow)
    With vsfTab
        lngNo = Val(.TextMatrix(NewRow, COL_tab项目序号))
        strName = .TextMatrix(NewRow, COL_tab项目名)
        strTmp = .TextMatrix(NewRow, COL_tab字符串)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        str值域 = arrStr(0)
        
        If str值域 = "" Then
            strInfo = ""
        Else
            strInfo = strName & "有效范围:" & str值域
        End If
        
        If lngNo = 4 And strName = "血压" Then '血压
            strInfo = strInfo & Space(4) & "录入规则:收缩压/舒张压"
            mrsCurInfo.Filter = ""
            mrsCurInfo.Sort = "编码"
            strTmp = ""
            Do While Not mrsCurInfo.EOF
                strTmp = strTmp & "、" & Nvl(mrsCurInfo!名称)
                mrsCurInfo.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then strInfo = strInfo & "或(" & strTmp & ")"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "汇总项目" & Space(4) & "录入规则:今天录入" & IIf(mbln汇总当天 = True, "今天", "昨天") & "的数据。"
    End With
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    strData = vsfTab.RowData(NewRow)
    lng项目序号 = zlStr.Nvl(vsfTab.TextMatrix(NewRow, COL_tab项目序号))
    If strData <> "" Then
        strTime = GetAnimalItemTime(vsfTab.Row, NewCol - vsfTab.FixedCols + 1, 0, strInfo)
        If strInfo <> "" Then lblStb.Caption = strInfo: lblStb.ForeColor = 255: Exit Sub
        If InStr(1, strTime, ";") > 0 Then arrTime = Split(strTime, ";")
        lngType = Val(Split(strData, ";")(1))
        With vsfTabDetail
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab字符串) = vsfTab.TextMatrix(NewRow, COL_tab字符串)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab项目序号) = vsfTab.TextMatrix(NewRow, COL_tab项目序号)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab项目名称) = vsfTab.TextMatrix(NewRow, COL_tab项目名)
        vsfTabDetail.TextMatrix(.FixedRows, COL_tab项目名) = vsfTab.TextMatrix(NewRow, COL_tab项目名)
        vsfTabDetail.Tag = vsfTabDetail.Rows - 1
        .ColHidden(.FixedCols - 1) = lngType <> 3
        If lngType = 3 Then .ColHidden(.FixedCols - 1) = IsLastTotal(lng项目序号)
        .MergeCellsFixed = flexMergeFree
        .MergeCol(.FixedCols - 1) = True
        If lngType = 3 Then
            mrsTableDetail.Filter = "项目序号= " & lng项目序号 & " and 时间 > '" & arrTime(0) & "' and 时间 <= '" & arrTime(1) & "' and 记录类型 <> 11 and 状态<>4 "
            If mrsTableDetail.RecordCount > 0 Then Call ShowTabDetail(.Rows - 1, NewRow, 0)
            Set rsTotle = ReturnTotle(lng项目序号, arrTime(0), arrTime(1))
            .Rows = .Rows + 1
            intRow = .Rows - 1
             rsTotle.Filter = ""
            Do While Not rsTotle.EOF
               
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols - 1) = zlStr.Nvl(rsTotle!项目名称)
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols) = Format(Split(strTime, ";")(0), "hh:mm") & "～" & Format(Split(strTime, ";")(1), "hh:mm")
                vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 1) = Val(Nvl(rsTotle!数值))
                Select Case rsTotle!数据来源
                    Case 0, 9
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "体温单录入"
                    Case 1
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "记录单同步"
                    Case 3
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "移动设备录入"
                    Case Else
                        vsfTabDetail.TextMatrix(intRow, vsfTabDetail.FixedCols + 2) = "其他设备同步"
                End Select
                .RowData(intRow) = "3"
                
                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &H8000000F
                intRow = intRow + 1
                .Rows = .Rows + 1
                rsTotle.MoveNext
            Loop
        
        .Rows = .Rows - 1
            
        Else
            mrsTableDetail.Filter = "项目序号 =" & lng项目序号 & " and 时间 >= '" & arrTime(0) & "' and 时间 <= '" & arrTime(1) & "' and 状态<>4"
            If mrsTableDetail.RecordCount > 0 Then Call ShowTabDetail(.FixedRows, NewRow, lngType)
        End If

        .Cell(flexcpAlignment, 0, 2, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
    End If
     
     '检查数据是否允许修改
    mrsTable.Filter = "项目序号=" & lngNo & " and 项目名称='" & strName & "'" & _
        "   and 列号=" & NewCol - vsfTab.FixedCols + 1
    If mrsTable.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsTable!数据来源) & ",") = 0 Then
            lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    If InStr(1, strData, ";") > 0 Then
        If Split(strData, ";")(1) = 3 And Not mbln录入小时 Then
            lblStb.Caption = "汇总数据仅能修改汇总小时，不能进行数据修改、删除惭怍"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Function IsLastTotal(ByVal lngNo As Long)
    '功能:是否是最后一级的汇总项目
    If mrsCollect Is Nothing Then Exit Function
    If mrsCollect.State = adStateOpen Then
        mrsCollect.Filter = "父序号=" & lngNo
        If mrsCollect.RecordCount > 0 Then
            IsLastTotal = False
        Else
            IsLastTotal = True
        End If
    End If

End Function


Private Function ReturnTotle(ByVal lngItemNO As Long, ByVal strBTime As String, ByVal strETime As String) As ADODB.Recordset
    '-------------
    '功能：求汇总明细
    '先求出汇总项目的第一级子节点 ，然后汇总这级节点及其所有子节点
    '-------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsTotle As New ADODB.Recordset
    Dim rsNO As ADODB.Recordset
    Dim strValue As String
    Dim strValues As String
    Dim strName As String
    Dim strFileds As String
    Dim str项目序号 As String
    Dim lng项目序号 As Long
    Dim dblData As Double
    Dim blnNumeric As Boolean
    
    On Error GoTo Errhand
    '初始化记录集
    strFileds = "开始时间," & adLongVarChar & ",20|结束时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & _
                adLongVarChar & ",20|数值," & adLongVarChar & ",100|数据来源," & adDouble & ",1"
    Call Record_Init(rsTotle, strFileds)
    strFileds = "开始时间|结束时间|项目序号|项目名称|数值|数据来源"
    
    Set rsNO = GetChildNo(lngItemNO)
    rsNO.Filter = ""
    Do While Not rsNO.EOF
        Set rsTemp = SetCollectPItem(rsNO!序号)
        mrsTableDetail.Filter = "项目序号=" & rsNO!序号 & " and 时间 > '" & strBTime & "' and 时间 <= '" & strETime & "'"
        lng项目序号 = rsNO!序号
        dblData = 0
        Do While Not mrsTableDetail.EOF
            dblData = dblData + Val(Nvl(mrsTableDetail!结果))
            mrsTableDetail.MoveNext
        Loop
        mrsTableDetail.Filter = "项目序号 =" & lng项目序号
        strName = IIf(mrsTableDetail.RecordCount > 0, mrsTableDetail!项目名称, vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名))
        rsTemp.Filter = ""
        Do While Not rsTemp.EOF
            '父项是明细不需要汇总
            If Val(Nvl(rsTemp!序号, 0)) <> lngItemNO Then
                mrsTableDetail.Filter = "项目序号=" & rsTemp!序号 & " and 时间 > '" & strBTime & "' and 时间 <= '" & strETime & "'"
                Do While Not mrsTableDetail.EOF
                    dblData = dblData + Val(Nvl(mrsTableDetail!结果))
                    If blnNumeric = False Then blnNumeric = IsNumeric(Nvl(mrsTableDetail!结果))
                    mrsTableDetail.MoveNext
                Loop
            End If
            rsTemp.MoveNext
        Loop
        strValue = IIf(dblData = 0 And blnNumeric = False, "", IIf(strValue = "", "", "(" & strValue & "h)") & IIf(Left(dblData, 1) = ".", "0", "") & dblData)
        strValues = strBTime & "|" & strETime & "|" & lng项目序号 & "|" & strName & "|" & strValue & "|0"
        If Val(strValue) <> 0 Then Call Record_Add(rsTotle, strFileds, strValues)
        strValue = ""
        rsNO.MoveNext
    Loop
        
    Set ReturnTotle = rsTotle
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SetCollectPItem(ByVal lngItemNO As Long) As ADODB.Recordset
'---------------------------------------------------------------------------
'功能:根据父项目ID重新组织子项目
'---------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNo As Long
    
    On Error GoTo Errhand
    
    '初始化记录集
    strFileds = "序号," & adDouble & ",18|父序号," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    Call Record_Init(rsCollect, strFileds)
    strFileds = "序号|父序号"
    
    mrsCollect.Filter = 0
   '复制记录集
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & Val(Nvl(!父序号))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "父序号=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & lngItemNO
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNo = Val(Nvl(!序号))
            '循环递归调用(获取子项的子项)
            Call SetCollectCItem(rsTemp, lngItemNO, lngNo)
            .MoveNext
        Loop
    End With
    
    Set SetCollectPItem = rsTemp
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub SetCollectCItem(rsTemp As ADODB.Recordset, ByVal lngParent As Long, ByVal lngItemNO As Long)
'功能: SetCollectPItem 调用
    
    Dim rsCollect As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngNo As Long
    
    On Error GoTo Errhand
    '初始化记录集
    strFileds = "序号," & adDouble & ",18|父序号," & adDouble & ",18"
    Call Record_Init(rsCollect, strFileds)
    strFileds = "序号|父序号"
    
    mrsCollect.Filter = 0
   '复制记录集
    With mrsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & Val(Nvl(!父序号))
            Call Record_Add(rsCollect, strFileds, strValues)
            .MoveNext
        Loop
    End With
    
    rsCollect.Filter = "父序号=" & lngItemNO
    With rsCollect
        Do While Not .EOF
            strValues = Val(Nvl(!序号)) & "|" & lngParent
            Call Record_Add(rsTemp, strFileds, strValues)
            lngNo = Val(Nvl(!序号))
            '循环递归调用(获取子项的子项)
            Call SetCollectCItem(rsTemp, lngParent, lngNo)
            .MoveNext
        Loop
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub vsfTab_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mblnScroll = True
    Call vsfTab_EnterCell
    mblnScroll = False
End Sub

Private Sub vsfTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfTab
        If vsfTab.Tag = "NO" Then Exit Sub
        If NewRow >= .FixedRows And NewCol >= .FixedRows Then
            If NewCol < .FixedCols + (Split(.TextMatrix(NewRow, COL_tab字符串), ",")(3)) Then
                mrsTable.Filter = "项目序号=" & Val(.TextMatrix(NewRow, COL_tab项目序号)) & " and 项目名称='" & .TextMatrix(NewRow, COL_tab项目名) & "'" & _
                    "   and 列号=" & NewCol - .FixedCols + 1
                If mrsTable.RecordCount > 0 Then
                    If InStr(1, ",0,3,9,", "," & Val(mrsTable!数据来源) & ",") = 0 Then
                        .FocusRect = 0
                        .HighLight = 0
                    Else
                        .FocusRect = flexFocusSolid
                    End If
                Else
                    .FocusRect = flexFocusSolid
                End If
            Else
                .FocusRect = 0
                .HighLight = 0
            End If
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub vsfTab_DblClick()
    With vsfTab
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid Then
            mblnEdit = True
            Call vsfTab_EnterCell
        End If
    End With
End Sub

Private Sub vsfTab_EnterCell()
    Dim strInfo As String
    Dim strValue As String, strValue1 As String
    Dim strTmp As String, strData As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim blnEdit As Boolean
    Dim blnSelect As Boolean
    Dim intType As Integer
    Dim int频次 As Integer, int项目类型 As Integer, int项目性质 As Integer
    Dim i As Integer, j As Integer
    Dim intNum As Integer, intLen As Integer, intRow As Integer, intCOl As Integer
    Dim lngItemNO As Long, lngColor As Long
    Dim arrValue() As String, arrValue1() As String
    Dim arrStr() As String
    
    If Not mblnInit Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    blnAllow = True
    blnEdit = True
    blnSelect = True
    
    If picEdit.Visible = True And txtEdit.Tag <> "" Then
        intRow = Split(txtEdit.Tag, "|")(0)
        intCOl = Split(txtEdit.Tag, "|")(1)
        If Split(txtEdit.Tag, "|")(2) <> 1 Then txtEdit_KeyPress (vbKeyEscape): Exit Sub
        If txtEdit.Visible = True Then
            strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
            lngColor = txtEdit.ForeColor
            If txtEdit.Text = "" Then strData = ""
        Else
            strData = Trim(lblCheck.Caption)
            lngColor = 0
        End If
        If Split(txtEdit.Tag, "|")(2) = 2 Then blnEdit = False
        If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then blnAllow = WriteIntoVfgTab(strData, vsfTab, False, True, strInfo)
        If cmdColor.Visible = True Then vsfTab.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = Val(cmdColor.Tag)
        mblnEdit = blnEdit
    End If
    
    '数据不合法
    If blnAllow = False Then
        If vsfTab.Row <> intRow Then vsfTab.Row = intRow
        If vsfTab.Col <> intCOl Then vsfTab.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    If Not vsfTab.RowIsVisible(vsfTab.Row) Then Exit Sub
    If Not mblnScroll And vsfTab.Visible Then vsfTab.SetFocus
    
    '隐藏所有编辑控件
    pic未记.Visible = False
    picEdit.Visible = False
    picEdit.Tag = ""
    txtEdit.Tag = "": txtEdit.Visible = False: txtEdit.Enabled = False
    picHour.Visible = False: picHour.Enabled = False
    txtHour.Tag = "": txtHour.Visible = False: txtHour.Enabled = False
    lblCheck.Visible = False: lblCheck.Enabled = False
    cmdColor.Visible = False
    cmdColor.Enabled = False
    cmdColor.Tag = 0
    picColor.Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False: txtLst.Text = ""
    lstSelect(0).Visible = False
    lstSelect(0).Enabled = False
    lstSelect(0).Tag = ""
    lstSelect(1).Visible = False
    lstSelect(1).Enabled = False
    lstSelect(1).Tag = ""
    
    If mblnFileBack = True Then
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        mblnEdit = False
        GoTo ErrInfo
    End If
    
    If mblnEdit = False Then Exit Sub
    
    With vsfTab
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And vsfTab.Col < .FixedCols + Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) Then
            intType = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4))
            int频次 = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3))
            int项目类型 = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(1))
            int项目性质 = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(5))
            '时间列检查是否可以编辑
            If .Row Mod 2 = 1 And (Split(vsfTab.RowData(.Row), ";")(1) = 2 Or Split(vsfTab.RowData(.Row), ";")(1) = 3) Then
                strInfo = "对于汇总数据或波动数据的时间段不能进行修改、删除操作"
                GoTo ErrInfo
            End If
    
            '检查录入项目时间是否超出用户设置的时间范围或是补录范围
            strTime = GetAnimalItemTime(.Row, .Col - vsfTab.FixedCols + 1, 0, strInfo)
            If .Row Mod 2 = 1 And .TextMatrix(.Row, .Col) <> "" Then
               If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "录入时间小于体温单开始时间"
               If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "录入时间大于体温单补录时间"
            End If
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
            '检查波动项目
            If IsWaveItem(Val(.TextMatrix(.Row, COL_tab项目序号))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                strInfo = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                GoTo ErrInfo
            End If
             '检查数据来源是否来自护理记录单或PDA
            mrsTable.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & .TextMatrix(.Row, COL_tab项目名) & "'" & _
                "   and 列号=" & .Col - .FixedCols + 1
            If mrsTable.RecordCount > 0 Then
                If InStr(1, ",0,3,9,", "," & Val(mrsTable!数据来源) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsTable!未记说明)
            End If
            
            '全天汇总显示录入时间,同步过来的也可以修改时间
            If blnEdit = False And Not (intType = 4 And int频次 = 1 And mbln录入小时 = True) Then
                strInfo = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                GoTo ErrInfo
            End If
            
            '非全天汇总数据不允许修改
            If intType = 4 And Not (int频次 = 1 And mbln录入小时 = True) Then
                strInfo = "汇总数据的汇总值不能进行修改、删除操作"
                GoTo ErrInfo
            End If
            
            If Not (intType = 2 Or intType = 3) Or vsfTab.Row Mod 2 = 1 Then
                picEdit.Width = .CellWidth + 10
                picEdit.Height = .CellHeight - 5
                picEdit.Top = .CellTop + .Top + 20
                picEdit.Left = .CellLeft + .Left + 15
                picEdit.Enabled = True
                picEdit.Visible = True
                picEdit.ZOrder 0
                txtEdit.Top = 0
                txtEdit.Left = 0
                txtEdit.Height = picEdit.Height
            End If
            '对于项目类型是文字类型的活动项目允许设置其字体颜色
            If int项目类型 = 1 And intType = 0 And int项目性质 = 2 And vsfTab.Row Mod 2 = 0 Then   '文本类型，活动 项目
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And int频次 = 1 And mbln录入小时 = True Then '全天汇总且显示汇总时间
                txtHour.Top = 10
                txtHour.Left = 10
                txtHour.Width = picHour.TextWidth("111")
                txtHour.Height = txtEdit.Height
                txtHour.MaxLength = 2
                txtHour.Visible = True
                txtHour.Enabled = True
                
                lblHour.Left = txtHour.Left + txtHour.Width
                lblHour.Top = 10
                lblHour.Visible = True
                lblHour.Enabled = True
                
                picHour.Top = -10
                picHour.Left = -10
                picHour.Width = lblHour.Left + lblHour.Width + picHour.TextWidth("1") \ 2
                picHour.Height = picEdit.Height + 20
                picHour.Visible = True
                picHour.Enabled = True
                picHour.ZOrder 0
                
                txtEdit.Top = 10
                txtEdit.Left = picHour.Left + picHour.Width + 10
                txtEdit.Width = picEdit.Width - picHour.Width + 10
                
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                If vsfTab.Row Mod 2 = 1 Then txtEdit.MaxLength = 5
                
                If InStr(1, .TextMatrix(vsfTab.Row, vsfTab.Col), ")") > 0 Then
                    txtHour.Text = Replace(Replace(Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(0), "(", ""), "h", "")
                    txtEdit.Text = Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(1)
                Else
                    txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                End If
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = False
                txtEdit.ZOrder 0
                picHour.SetFocus
            ElseIf (intType = 2 Or intType = 3) And vsfTab.Row Mod 2 = 0 Then '单选或复选
                strValue = Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(0)
                Select Case intType
                    Case 2
                        If Left(strValue, 1) <> ":" Then strValue = ":" & strValue
                        intType = 0
                    Case 3
                        intType = 1
                End Select
                
                arrValue = Split(strValue, ":")
                lstSelect(intType).Clear
                PicLst.Tag = "1"
                For i = 0 To UBound(arrValue)
                    If Left(arrValue(i), 1) = "√" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
                    lstSelect(intType).AddItem arrValue(i), i
                     
                     If intType = 0 Then
                        ReDim arrValue1(0)
                        arrValue1(0) = .TextMatrix(.Row, .Col)
                        txtLst.Text = .TextMatrix(.Row, .Col)
                        txtLst.Tag = 1
                     Else
                        arrValue1 = Split(.TextMatrix(.Row, .Col), ",")
                     End If
                     For j = 0 To UBound(arrValue1)
                        If arrValue1(j) = arrValue(i) Then
                            lstSelect(intType).Selected(i) = True
                            blnSelect = True
                        End If
                    Next j
                Next i
                
                If blnSelect = False And strValue1 <> "" And IIf(intType = 0, Trim(txtLst.Text) = "", True) Then
                    For i = 0 To lstSelect(intType).ListCount - 1
                        If lstSelect(intType).List(i) = strValue1 Then
                            lstSelect(intType).Selected(i) = True
                        End If
                    Next i
                End If
                
                If lstSelect(intType).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = 0
                
                '控件显示
                If intType = 0 Then '单选项目提供可以选择和录入功能
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = .CellTop + vsfTab.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("刘") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTab.Height Then PicLst.Height = vsfTab.Height
                    If PicLst.Top + PicLst.Height > vsfTab.Height Then PicLst.Top = .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = vsfTab.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("刘")
                    End If
                    lbllst(2).FontName = .FontName
                    lbllst(2).FontSize = .FontSize
                    lbllst(2).Visible = True
        
                    txtLst.Top = lbllst(2).Top + lbllst(2).Height + 20
                    txtLst.Left = -10
                    txtLst.Width = PicLst.Width
                    txtLst.Height = .CellHeight - 5
                    txtLst.FontName = .FontName
                    txtLst.FontSize = .FontSize
                    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrStr = Split(strTmp, ",")
                    intNum = Val(arrStr(2))
                    intLen = Val(arrStr(6))
                    txtLst.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    txtLst.Visible = True
                
                    lbllst(3).Left = 20
                    lbllst(3).Top = txtLst.Top + txtLst.Height + 20
                    lbllst(3).FontName = .FontName
                    lbllst(3).FontSize = .FontSize
                    lbllst(3).Visible = True
                    
                    lstSelect(intType).Top = lbllst(3).Top + lbllst(3).Height + 20
                    lstSelect(intType).Left = -10
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Width = PicLst.Width
                    lstSelect(intType).Height = PicLst.Height - lstSelect(intType).Top
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col & "|" & "1"
                    txtEdit.Tag = "||1"
                     
                    If lstSelect(intType).Top + lstSelect(intType).Height <> PicLst.Height Then
                        PicLst.Height = lstSelect(intType).Top + lstSelect(intType).Height
                    End If
                    PicLst.SetFocus
                Else
                    lstSelect(intType).Top = .CellTop + vsfTab.Top
                    lstSelect(intType).Left = .CellLeft + .Left + 15
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If lstSelect(intType).Width < .CellWidth + 20 Then lstSelect(intType).Width = .CellWidth + 20
                    If lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Height = vsfTab.Height
                    End If
                    If lstSelect(intType).Top + lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Top = .CellTop + .Top + .CellHeight + 20 - lstSelect(intType).Height
                    End If
                    If lstSelect(intType).Top < 0 Then lstSelect(intType).Top = vsfTab.Top
                    
                        lstSelect(intType).Visible = True
                        lstSelect(intType).Enabled = True
                        lstSelect(intType).ZOrder 0
                        
                        lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                        lbllst(intType).Tag = .Row & "|" & .Col & "|1"
                        lstSelect(intType).SetFocus
                    End If
            ElseIf intType = 5 Then '选择
                lblCheck.Width = picEdit.Width
                lblCheck.Height = picEdit.Height
                lblCheck.Caption = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                lblCheck.Visible = True
                lblCheck.Enabled = True
                lblCheck.ZOrder 0
                picEdit.SetFocus
            Else
                txtEdit.Width = picEdit.Width
ShowText:
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                If vsfTab.Row Mod 2 = 1 Then txtEdit.MaxLength = 5
                
                txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "'" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col & "|" & "1"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = True
                txtEdit.ZOrder 0
                picEdit.SetFocus
            
            End If
            
        End If
    End With
ErrFouce:
    If picEdit.Visible = True And txtEdit.Enabled = True Then txtEdit.SetFocus: Call zlControl.TxtSelAll(txtEdit)
ErrInfo:
    If strInfo <> "" Then
        lblStb.Caption = strInfo
        lblStb.ForeColor = 255
    End If
End Sub

Private Function WriteIntoVfgTab(ByVal strText As String, ByVal vsf As Object, Optional blnDelete As Boolean = False, Optional ByVal blnVisible As Boolean = True, Optional strErrMsg As String = "") As Boolean
    '-------------------------------------------------------------------------
    '功能:用户编辑的数据写入vsfTab,数据验证
    '参数:strtext 编辑的文本信息   blndelete 是否在VsfTab按Delete 键删除信息
    '-------------------------------------------------------------------------
    Dim intRow As Integer
    Dim intCOl As Integer
    Dim str项目名称 As String, strTmp As String, strPart As String
    Dim str值域 As String, strHour As String, strHourOld As String
    Dim strValue As String, strTime As String, strOldTime As String
    Dim intType As Integer, intNum As Integer, lngLen As Long, int频次 As Integer
    Dim int性质 As Integer, int表示 As Integer, intIndex As Integer, int记录类型 As Integer
    Dim int状态 As Integer  '--数据修改信息
    Dim i As Integer
    Dim lngVsfType As Integer
    Dim arrStr() As String, arrOldTime() As String
    Dim blnAllow As Boolean, blnTrue As Boolean
    Dim BlnTime As Boolean
    Dim lng项目序号 As Long, lngColor As Long
    Dim rsTemp As New ADODB.Recordset
    On Err GoTo Errhand
    
    If Not blnDelete Then
        If picEdit.Visible And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            lngVsfType = Split(txtEdit.Tag, "|")(2)
            If lngVsfType = 1 Then
                If vsfTab.Name <> vsf.Name Then Set vsf = vsfTab
            Else
                If vsfTabDetail.Name <> vsf.Name Then Set vsf = vsfTabDetail
            End If
             '是否是输入时间
            If Val(lngVsfType) = 1 Then
                BlnTime = (intRow Mod 2 = 1)
            Else
                BlnTime = intCOl = vsfTabDetail.FixedCols
            End If
            If txtEdit.Visible = True Or lblCheck.Visible = True Then
                strTmp = vsf.TextMatrix(intRow, COL_tab字符串)
                lng项目序号 = Val(vsf.TextMatrix(intRow, COL_tab项目序号))
                str项目名称 = vsf.TextMatrix(intRow, COL_tab项目名)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                str值域 = arrStr(0)
                intType = Val(arrStr(1))
                intNum = Val(arrStr(2))
                int频次 = Val(arrStr(3))
                int表示 = Val(arrStr(4))
                int性质 = Val(arrStr(5))
                lngLen = Val(arrStr(6))
                strPart = arrStr(7)
                
                If intType = 1 Then str值域 = ""
                '全天汇总，且显示汇总时间
                If int表示 = 4 And int频次 = 1 And mbln录入小时 = True Then
                    If InStr(1, strText, ")") > 0 Then
                        strHour = Replace(Replace(Split(strText, ")")(0), "(", ""), "h", "")
                        If strHour <> "" Then
                            If Not Val(strHour) >= 0 And Val(strHour) <= 24 Then
                                lblStb.Caption = "汇总小时只能在0到24之间，请重新录入！": lblStb.ForeColor = 255
                                Exit Function
                            End If
                            strHour = "(" & strHour & "h)"
                        End If
                        strText = Split(strText, ")")(1)
                        If Trim(strText) = "" Then strHour = ""
                    End If
                End If
                If txtEdit.Enabled = True Or txtHour.Visible = True Then
                    blnAllow = CheckValidata(intRow, intCOl, lng项目序号, intType, intNum, str值域, int表示, lngLen, strText, BlnTime, strErrMsg)
                End If
            End If
            strValue = Split(IIf(Trim(picEdit.Tag) = "", "'", Trim(picEdit.Tag)), "'")(0)
        ElseIf lstSelect(0).Visible = True Or lstSelect(1).Visible = True Then
            If lstSelect(0).Visible = True Then strValue = lstSelect(0).Tag: intIndex = 0
            If lstSelect(1).Visible = True Then strValue = lstSelect(1).Tag: intIndex = 1
            intRow = Split(lbllst(intIndex).Tag, "|")(0)
            intCOl = Split(lbllst(intIndex).Tag, "|")(1)
            lng项目序号 = Val(vsf.TextMatrix(intRow, COL_tab项目序号))
            str项目名称 = vsf.TextMatrix(intRow, COL_tab项目名)
            strTmp = vsf.TextMatrix(intRow, COL_tab字符串)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            intType = Val(arrStr(1))
            int性质 = Val(arrStr(5))
            strPart = arrStr(7)
            
            blnAllow = True
        End If
    Else
        blnAllow = True
        If InStr(1, picTab.Tag, "|") = 0 Then Exit Function
        intRow = Split(picTab.Tag, "|")(0)
        intCOl = Split(picTab.Tag, "|")(1)
        lng项目序号 = Val(vsf.TextMatrix(intRow, COL_tab项目序号))
        str项目名称 = vsf.TextMatrix(intRow, COL_tab项目名)
        strTmp = vsf.TextMatrix(intRow, COL_tab字符串)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        intType = Val(arrStr(1))
        int频次 = Val(arrStr(3))
        int表示 = Val(arrStr(4))
        int性质 = Val(arrStr(5))
        strPart = arrStr(7)
        strHour = ""
        strValue = fraTable.Tag
    End If
    
    If blnAllow = True Then
        lngColor = 0
        vsf.TextMatrix(intRow, intCOl) = strHour & strText
        If cmdColor.Visible = True Then lngColor = cmdColor.Tag
        vsf.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = lngColor
        If vsf.Name = vsfTabDetail.Name Then
            vsf.Cell(flexcpForeColor, intRow, intCOl - 1, intRow, intCOl + 1) = lngColor
        End If
        mblnEdit = True
    Else
        If strErrMsg <> "" Then GoTo ErrInfo
        Exit Function
    End If
    
    mrsTable.Filter = ""
    int记录类型 = 1
    blnTrue = False
    '更新数据修改标志
    If blnAllow = True Then
        strHour = Replace(Replace(strHour, "(", ""), "h)", "")
        If int表示 = 4 And int频次 = 1 And mbln录入小时 = True Then
            If InStr(1, strValue, ")") > 0 Then
                strHourOld = Replace(Replace(Split(strValue, ")")(0), "(", ""), "h", "")
                strValue = Split(strValue, ")")(1)
            End If
            '更新用户录入的汇总小时，只能修该汇总时间
            If Val(strHour) <> Val(strHourOld) Then
                blnTrue = True
                int记录类型 = 11
                GoTo DataUpdate
            End If
        End If
DataUpdate:
        If vsf.Name = "vsfTab" Then
            If InStr(vsf.TextMatrix(intRow, col_tab原始时间), "'") > 0 Then
                arrOldTime = Split(vsf.TextMatrix(intRow, col_tab原始时间), "'")
                strTime = Format(arrOldTime(intCOl - vsfTab.FixedCols), "YYYY-MM-DD hh:mm:ss")
            Else
                ReDim Preserve arrOldTime(0)
                strTime = vsf.TextMatrix(intRow, col_tab原始时间)
            End If
        Else
            strTime = Format(vsf.TextMatrix(intRow, col_tab原始时间), "YYYY-MM-DD hh:mm:ss")
        End If
        If BlnTime Then
            strText = Format(IIf(Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 0 Or mbln汇总当天, dtpDate.Value & " " & strText, dtpDate.Value - 1 & " " & strText), "YYYY-MM-DD hh:mm:ss")
        Else
            If vsf.Name = vsfTab.Name Then
                If vsf.TextMatrix(intRow - 1, intCOl) <> "" Then
                    If strTime = "" Then strTime = Format(dtpDate.Value & " " & vsf.TextMatrix(intRow - 1, intCOl), "YYYY-MM-DD hh:mm:ss")
                End If
            Else
                If vsf.TextMatrix(intRow, vsf.FixedCols) <> "" Then
                    If strTime = "" Then strTime = Format(dtpDate.Value & " " & vsf.TextMatrix(intRow - 1, intCOl), "YYYY-MM-DD hh:mm:ss")
                End If
            End If
            If strTime = "" Then
                If lngVsfType = 1 Then
                    strTime = GetAnimalItemTime(intRow, intCOl - vsfTab.FixedCols + 1, 1, strErrMsg)
                Else
                    strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 1, strErrMsg)
                End If
                If strErrMsg <> "" Then GoTo ErrInfo
                If IsExistData(strTime, lng项目序号) = False Then
                    strErrMsg = lblStb.Caption
                    GoTo ErrInfo
                End If
            End If
        End If
        
        mrsTableDetail.Filter = "项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' And 记录类型=" & int记录类型 & " and 时间='" & strTime & "'"
        If BlnTime Then strTime = strText: strText = ""
        If mrsTableDetail.RecordCount > 0 Then
            mrsTableDetail!未记说明 = lngColor
            If mrsTableDetail!状态 <> 1 Then '原有的数据 修改、删除后的状态
                If BlnTime And mrsTableDetail!状态 = 0 Then
                    mrsTableDetail!状态 = 3
                Else
                    mrsTableDetail!状态 = 2
                End If
                If BlnTime Then
                    mrsTableDetail!时间 = strTime
                Else
                    mrsTableDetail!结果 = IIf(blnTrue = True, strHour, strText)
                End If
                If strText = "" And mrsTableDetail!结果 = "" Then
                    mrsTableDetail!状态 = 4 '删除
                End If
            Else
                If Trim(IIf(blnTrue = True, strHour, strText)) = "" Then '新增删除
                        mrsTableDetail.Delete
                    Else '新增
                        mrsTableDetail!状态 = 1
                        If BlnTime Then
                            mrsTableDetail!时间 = strTime
                        Else
                            mrsTableDetail!结果 = IIf(blnTrue = True, strHour, strText)
                        End If
                    End If
            End If
            mrsTableDetail.Update
        Else '新增
            If Trim(strText) <> "" Then
                If strErrMsg <> "" Then GoTo ErrInfo
            End If
            strText = Replace(Replace(strText, "|", "∣"), "'", "")
            gstrFields = "id|分组名|结果|体温部位|标记|时间|项目序号|项目名称|复试合格|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
            If BlnTime Then
                gstrValues = GetMaxID(mrsTableDetail) & "|2)体温表格项目|" & "" & "|" & strPart & "|" & _
                0 & "|" & strTime & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & lngColor & "|0|0|0|0|0|1|" & vsfTab.Col - vsfTab.FixedCols + 1 & "|" & int记录类型
            Else
                gstrValues = GetMaxID(mrsTableDetail) & "|2)体温表格项目|" & IIf(blnTrue = True, strHour, strText) & "|" & strPart & "|" & _
                0 & "|" & strTime & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & lngColor & "|0|0|0|0|0|1|" & vsfTab.Col - vsfTab.FixedCols + 1 & "|" & int记录类型
            End If
            Call Record_Add(mrsTableDetail, gstrFields, gstrValues)
            If lngVsfType = 1 Then
                arrOldTime(intCOl - vsfTab.FixedCols) = strTime
                vsf.TextMatrix(intRow, col_tab原始时间) = Join(arrOldTime, "'")
            Else
                vsf.TextMatrix(intRow, col_tab原始时间) = strTime
            End If
        End If
    End If
    mrsTableDetail.Filter = "状态<> 4 "
    
    gstrFields = "ID," & adDouble & ",18|分组名," & adLongVarChar & ",40|结果," & adLongVarChar & ",400|体温部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复试合格," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|原始显示状态," & adDouble & ",1|" & _
         "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
    Call Record_Init(rsTemp, gstrFields)
    
    Do While Not mrsTableDetail.EOF
        rsTemp.AddNew
        For i = 0 To mrsTableDetail.Fields.Count - 1
            rsTemp.Fields(mrsTableDetail.Fields(i).Name).Value = mrsTableDetail.Fields(i).Value
        Next i
        rsTemp.Update
        mrsTableDetail.MoveNext
    Loop
    
    Call InitTableData(rsTemp)
    Call ShowTable
    If blnAllow = True And blnVisible = True Then Call txtEdit_KeyPress(vbKeyEscape): mblnEdit = True
'    Call SetColSelect
    WriteIntoVfgTab = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
        vsf.TextMatrix(intRow, intCOl) = strValue
    End If
End Function

Private Function CheckValidata(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal lngNo As Long, ByVal intType As Integer, ByVal int小数 As Integer, ByVal str值域 As String, _
    ByVal int表示 As Integer, ByVal lngLen As Long, strInfo As String, ByVal BlnTime As Boolean, Optional strErrMsg As String = "") As Boolean
'-------------------------------------------------------------
'功能：检查数据合法性（表格数据）
'参数:introw：哪一行 intCol： 那一列  lngNo:项目序号 intype： 项目类型 0数字类型 1 文字类型 str值域：项目值域
'   lngLen：项目长度  strInfo：要校验的文本值
'-------------------------------------------------------------
    Dim strName As String, strTmp As String
    Dim strTime As String
    Dim strMsg As String, strText As String
    Dim lngRow As Long, lng项目序号 As Long
    Dim blnAllow As Boolean '是否是大便次数和入液量
    Dim blnOK As Boolean
    Dim i As Integer
    Dim int频次 As Integer
    Dim arrValue() As String
    
    On Error GoTo Errhand
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名)
    lng项目序号 = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号)
    lngRow = intRow - vsfTab.FixedRows + 1
    
    If strInfo = "" Then
        CheckValidata = True
        Exit Function
    End If
    If BlnTime Then
        If Len(strInfo) < 3 Or (Len(strInfo) = 3 And InStr(strInfo, ":") > 0) Or (Len(strInfo) = 5 And InStr(strInfo, ":") <= 0) Or Len(strInfo) > 5 Then
            strMsg = "第" & lngRow & "行[" & strName & "]的时间格式录入不正确，应为 小时：分钟"
            GoTo ErrInfo
        End If
        strInfo = CToData(strInfo)
        
        If InStr(1, strInfo, ":") > 0 Then
            If Not IsDate(strInfo) Then
                 strMsg = "第" & lngRow & "行[" & strName & "]的时间录入不正确，应为 小时：分钟"
                GoTo ErrInfo
            End If
            int频次 = Split(vsfTab.RowData(vsfTab.Row), ";")(0)
            '检查录入项目时间是否超出用户设置的时间范围或是补录范围
            strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strMsg)
            If strMsg <> "" Then GoTo ErrInfo
            If strInfo <> "" Then
                If (Format(Split(strTime, ";")(0), "hh:mm:ss") < Format(Split(strTime, ";")(1), "hh:mm:ss") And Split(vsfTab.RowData(vsfTab.Row), ";")(1) <> 0) Or (int频次 = 1 And Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 3) Then
                    If CDate(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo) < CDate(Split(strTime, ";")(0)) Then strMsg = "录入时间小于当天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1): GoTo ErrInfo
                    If CDate(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo) > CDate(Split(strTime, ";")(1)) Then strMsg = "录入时间大于当天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1): GoTo ErrInfo
                Else
                    If CDate(dtpDate.Value & " " & strInfo) < CDate(Split(strTime, ";")(0)) Then strMsg = "录入时间小于当天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1): GoTo ErrInfo
                    If CDate(dtpDate.Value & " " & strInfo) > CDate(Split(strTime, ";")(1)) Then strMsg = "录入时间大于当天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1): GoTo ErrInfo
                End If
            End If
            
            If Format(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss") < mstrBTime Then
                strMsg = "第" & lngRow & "行[" & strName & "]的时间小于体温单开始时间,请检查!"
                GoTo ErrInfo
            End If
            If Format(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss") > mstrETime Then
                strMsg = "第" & lngRow & "行[" & strName & "]的时间大于体温单补录时间,请检查!"
                GoTo ErrInfo
            End If
            If IsExistData(Format(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & strInfo, "YYYY-MM-DD hh:mm:ss"), lng项目序号) = False Then
                strErrMsg = lblStb.Caption
                GoTo ErrInfo
            End If
        End If
    Else
        blnAllow = True
        If strName = "体重" Or strName = "身高" Then
            If IsNumeric(strInfo) Then
                blnAllow = True
            Else
                blnAllow = False
            End If
        End If
        '大便次数和入夜不检查
        If blnAllow = True Then blnAllow = IIf(InStr(1, "," & gint大便 & "," & gint入液 & ",", "," & lngNo & ",") > 0, False, True)
        If Not (intType = 0 And InStr(1, "0,4", int表示) <> 0) Then
            If LenB(StrConv(strInfo, vbFromUnicode)) > lngLen Then
                strMsg = "第" & lngRow & "行[" & strName & "]的值超长(最大长度:" & lngLen & "),请检查!"
                GoTo ErrInfo
            End If
        Else
            If intType = 0 Then
                If int表示 = 4 Or str值域 = "" Then
                    str值域 = "0～" & IIf(lngLen - int小数 > 0, String(lngLen - int小数, "9"), "0") & IIf(int小数 > 0, "." & String(int小数, "9"), "")
                End If
                If lngNo <> 4 And lngNo <> 5 And blnAllow = True Then
                    If Not IsNumeric(strInfo) Then
                        strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                        GoTo ErrInfo
                    End If
                End If
                    
                If lngNo = 4 And strName = "血压" Then
                '血压可以录入文字说明：外出，未测等
                    mrsCurInfo.Filter = "名称='" & strInfo & "'"
                    If Not mrsCurInfo.EOF Then
                        CheckValidata = True
                        Exit Function
                    Else
                        strTmp = ""
                        mrsCurInfo.Filter = "": mrsCurInfo.Sort = "编码"
                        Do While Not mrsCurInfo.EOF
                            strTmp = strTmp & "、" & Nvl(mrsCurInfo!名称)
                            mrsCurInfo.MoveNext
                        Loop
                        strTmp = Mid(strTmp, 2)
                
                        If InStr(1, strInfo, "/") = 0 Then
                            strMsg = "第" & lngRow & "行[血压]数据的格式错误：收缩压/舒张压" & IIf(strTmp <> "", "或(" & strTmp & ")", "") & "！"
                            GoTo ErrInfo
                        End If
                        If Trim(Split(strInfo, "/")(0)) = "" Or Trim(Split(strInfo, "/")(1)) = "" Then
                            strMsg = "第" & lngRow & "行[血压]数据录入错误：收缩压/舒张压" & IIf(strTmp <> "", "或(" & strTmp & ")", "") & "！"
                            GoTo ErrInfo
                        End If
                    End If
                End If
                
                If UBound(Split(strInfo, "/")) > 1 And blnAllow = True Then
                    strMsg = "第" & lngRow & "行[" & strName & "]数据录入错误，请检查！"
                    GoTo ErrInfo
                End If
                
                '检查数据在有效范围内是否有效
                arrValue = Split(strInfo, "/")
                For i = 0 To UBound(arrValue)
                    blnOK = False
                    strText = arrValue(i)
                    If Not blnOK Then
                        If Not IsNumeric(strText) And blnAllow = True Then
                            strMsg = "第" & lngRow & "行[" & strName & "]数据录入错误" & Space(4) & "有效范围:" & str值域
                            GoTo ErrInfo
                        End If
                    End If
                        
                    If Not blnOK And strText <> "" And blnAllow = True Then
                        strText = Format(Val(strText), "#0" & IIf(int小数 > 0, ".", "") & String(int小数, "0"))
                        '0.30转为0.3
                        If strText = Val(strText) Then strText = Val(strText)
                        If Left(strText, 1) = "." Then strText = 0 & strText
                    End If
                    
                    If int表示 <> 4 And blnAllow = True Then
                        If Len(Replace(strText, ".", "")) > lngLen Then
                            strMsg = "第" & lngRow & "行[" & strName & "]的值超长(最大长度:" & lngLen & "),请检查!"
                            GoTo ErrInfo
                        End If
                    End If
                    
                    If IsNumeric(Split(str值域, "～")(0)) And IsNumeric(strText) Then
                        If blnAllow = True Then   '大便次数不进行有效范围检查
                            If Not (Val(strText) >= Split(str值域, "～")(0) And Val(strText) <= Split(str值域, "～")(1)) Then
                                strMsg = strName & "超出有效范围(" & str值域 & "),请检查!"
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                    arrValue(i) = strText
                Next
                strInfo = Join(arrValue, "/")
            End If
        End If
    End If
    
    CheckValidata = True
    Exit Function
    
ErrInfo:
    strErrMsg = strMsg
    
Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Function IsExistData(ByVal strTime As String, ByVal lngNo As Long) As Boolean
    '-----------------------------------------------
    '检查当前时间是否已存在数据
    '-----------------------------------------------
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    '检查修改的时间是否已经存在数据
    If mrsCurve.State = adStateOpen Then
        mrsCurve.Filter = "时间= '" & strTime & "' And 项目序号 = " & lngNo
        If mrsCurve.RecordCount > 0 Then lblStb.Caption = "当前时间已经存在数据,请重新输入时间.": lblStb.ForeColor = 255: Exit Function
    End If
    If mrsTableDetail.State = adStateOpen Then
        mrsTableDetail.Filter = "时间= '" & strTime & "' and 项目序号=" & lngNo
        If mrsTableDetail.RecordCount > 0 Then
            If mrsTableDetail!结果 <> "" Then lblStb.Caption = "当前时间已经存在数据,请重新输入时间.": lblStb.ForeColor = 255: Exit Function
        End If
    End If
    strSQL = "select 1 From 病人护理文件 a,病人护理数据 b,病人护理明细 c" & vbNewLine & _
        " where A.ID=B.文件ID and b.id =c.记录id and A.ID=[1] and A.病人ID=[2] and A.主页ID=[3] And nvl(A.婴儿,0)=[4]" & vbNewLine & _
        " and B.发生时间=[5] and c.项目序号=[6]"
        
    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
        mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查时间", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, CDate(strTime), lngNo)
    
    If rsTemp.RecordCount > 0 Then
        lblStb.Caption = "当前时间已经存在数据,请重新输入时间."
        lblStb.ForeColor = 255
        Exit Function
    End If
    
    IsExistData = True
       
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vsfTab_GotFocus()
    If vsfTab.Tag <> "NO" Then vsfTab.Tag = "1"
End Sub

Private Sub vsfTab_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCols As Integer
    Dim intType As Integer, int频次 As Integer
    Dim blnTrue As Boolean
    Dim blnEdit As Boolean
    Dim strText As String
    
    If vsfTab.Tag = "NO" Then Exit Sub
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    
    '屏蔽掉某些功能键
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
    intCols = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) + vsfTab.FixedCols
    
    With vsfTab
        If KeyCode = vbKeyReturn Then
NextCol2: '跳到下一行
            If .Col < vsfTab.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < intCols - 1 Then
                If vsfTab.Row Mod 2 = 1 Then
                    GoTo NextRow2
                Else
                    If .Row > .FixedRows Then .Row = .Row - 1
                    .Col = .Col + 1
                    If .ColHidden(.Col) = True Then GoTo NextCol2
                End If
            Else
NextRow2: '跳到下一列
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Col + 1 = intCols And .Row Mod 2 = 1 Then .Col = vsfTab.FixedCols
                    If .RowHidden(.Row) = True Then GoTo NextRow2
                Else
                    txtEdit.Tag = "||1"
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
        
            Exit Sub
        End If
        '左键
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > vsfTab.FixedCols Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > vsfTab.FixedRows Then
                    .Row = .Row - 1
                    If .RowHidden(.Row) Then GoTo PreRow2
                    .Col = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) + vsfTab.FixedCols
                    GoTo PreCol2
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
        
        '删除信息
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < intCols Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '检查项目是否是波动项目
                    If IsWaveItem(Val(.TextMatrix(.Row, COL_tab项目序号))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                        lblStb.Caption = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    '检查数据来源是否来自护理记录单或PDA
                    mrsCurve.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & .TextMatrix(.Row, COL_tab项目名) & "'" & _
                        "   And 列号=" & .Col - .FixedCols + 1
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    int频次 = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3))
                    intType = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4))
                    If blnEdit = False And Not (intType = 4 And int频次 = 1 And mbln录入小时 = True) Then
                        lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    fraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    If blnEdit = False Then '表明是全天汇总项目，并且mbln录入小时=true
                        If InStr(1, .TextMatrix(.Row, .Col), ")") > 0 Then
                            strText = Split(.TextMatrix(.Row, .Col), ")")(1)
                        Else
                            GoTo ErrExit
                        End If
                    End If
                    blnTrue = WriteIntoVfgTab(strText, vsfTab, True)
                End If
            End If
ErrExit:
            mblnEdit = False
            Exit Sub
        End If
        mblnEdit = True
        Call vsfTab_EnterCell
    End With
End Sub

Private Sub vsfTab_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
    If mblnFileBack = True Then
        Cancel = True
        vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        Exit Sub
    End If
    If vsfTab.Tag = "NO" Then Cancel = True
End Sub

Private Sub vsfTabDetail_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strName As String, strTmp As String, str值域 As String
    Dim strInfo As String
    Dim lngNo As Long
    Dim arrStr() As String
    
    If mblnInit = False Then Exit Sub
    If NewRow < vsfTabDetail.FixedRows Or NewCol < vsfTabDetail.FixedCols Or NewRow > Val(vsfTabDetail.Tag) Then Exit Sub
    Call AdjustRowFlag(vsfTabDetail, NewRow)
    With vsfTabDetail
        lngNo = Val(vsfTabDetail.TextMatrix(NewRow, COL_tab项目序号))
        strName = .TextMatrix(NewRow, COL_tab项目名称)
        strTmp = .TextMatrix(NewRow, COL_tab字符串)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        str值域 = arrStr(0)
        
        If str值域 = "" Then
            strInfo = ""
        Else
            strInfo = strName & "有效范围:" & str值域
        End If
        
        If lngNo = 4 And strName = "血压" Then '血压
            strInfo = strInfo & Space(4) & "录入规则:收缩压/舒张压"
            mrsCurInfo.Filter = ""
            mrsCurInfo.Sort = "编码"
            strTmp = ""
            Do While Not mrsCurInfo.EOF
                strTmp = strTmp & "、" & Nvl(mrsCurInfo!名称)
                mrsCurInfo.MoveNext
            Loop
            strTmp = Mid(strTmp, 2)
            If strTmp <> "" Then strInfo = strInfo & "或(" & strTmp & ")"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "汇总项目" & Space(4) & "录入规则:今天录入" & IIf(mbln汇总当天 = True, "今天", "昨天") & "的数据。"
        
        
    End With
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    mrsCurve.Filter = "项目序号=" & lngNo & " and 项目名称='" & strName & "'" & _
        "   and 列号=" & NewCol - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
            lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
End Sub

Private Sub vsfTabDetail_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mblnScroll = True
    Call vsfTabDetail_EnterCell
    mblnScroll = False
End Sub

Private Sub vsfTabDetail_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
    With vsfTabDetail
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid Then
            mblnEdit = True
            Call vsfTabDetail_EnterCell
        End If
    End With
End Sub

Private Sub vsfTabDetail_DblClick()
    With vsfTabDetail
        If vsfTab.Tag = "NO" Then Exit Sub
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid And .Col < .FixedCols + 2 Then
            mblnEdit = True
            Call vsfTabDetail_EnterCell
        End If
    End With
End Sub

Private Sub vsfTabDetail_EnterCell()
    Dim strInfo As String
    Dim strValue As String, strValue1 As String
    Dim strTmp As String, strData As String
    Dim strTime As String
    Dim blnAllow As Boolean
    Dim blnEdit As Boolean
    Dim blnSelect As Boolean
    Dim intType As Integer
    Dim int频次 As Integer, int项目类型 As Integer, int项目性质 As Integer
    Dim i As Integer, j As Integer
    Dim intNum As Integer, intLen As Integer, intRow As Integer, intCOl As Integer
    Dim lngItemNO As Long, lngColor As Long
    Dim arrValue() As String, arrValue1() As String
    Dim arrStr() As String
    
    If Not mblnInit Then Exit Sub
    If vsfTab.Tag = "NO" Then Exit Sub
    If vsfTabDetail.Rows = vsfTabDetail.FixedRows Then Exit Sub
    blnAllow = True
    blnEdit = True
    blnSelect = True
    
    If picEdit.Visible = True And txtEdit.Tag <> "" Then
        intRow = Split(txtEdit.Tag, "|")(0)
        intCOl = Split(txtEdit.Tag, "|")(1)
        If Split(txtEdit.Tag, "|")(2) <> 2 Then Call txtEdit_KeyPress(vbKeyEscape): Exit Sub
        If txtEdit.Visible = True Then
            strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
            lngColor = txtEdit.ForeColor
        Else
            strData = Trim(lblCheck.Caption)
            lngColor = 0
        End If
        
        If Split(txtEdit.Tag, "|")(2) = 1 Then mblnEdit = False
        If intCOl > vsfTabDetail.FixedCols + 1 Then mblnEdit = False
        
        If IIf(cmdColor.Visible, strData & "'" & lngColor <> picEdit.Tag, strData <> Split(picEdit.Tag, "'")(0)) Then blnAllow = WriteIntoVfgTab(strData, vsfTabDetail, False, True, strInfo)
        If cmdColor.Visible = True And mblnEdit = True Then vsfTabDetail.Cell(flexcpForeColor, intRow, intCOl - 1, intRow, intCOl + 1) = Val(cmdColor.Tag)
    End If
    
    '数据不合法
    If blnAllow = False Then
        If vsfTabDetail.Row <> intRow Then vsfTabDetail.Row = intRow
        If vsfTabDetail.Col <> intCOl Then vsfTabDetail.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTabDetail.Row < vsfTabDetail.FixedRows And vsfTabDetail.Col < vsfTabDetail.FixedCols Then Exit Sub
    If Not vsfTabDetail.RowIsVisible(vsfTabDetail.Row) Then Exit Sub
    If Not mblnScroll And vsfTabDetail.Visible Then vsfTabDetail.SetFocus

    '隐藏所有编辑控件
    pic未记.Visible = False
    picEdit.Visible = False
    picEdit.Tag = ""
    txtEdit.Tag = "": txtEdit.Visible = False: txtEdit.Enabled = False
    picHour.Visible = False: picHour.Enabled = False
    txtHour.Tag = "": txtHour.Visible = False: txtHour.Enabled = False
    lblCheck.Visible = False: lblCheck.Enabled = False
    cmdColor.Visible = False
    cmdColor.Enabled = False
    cmdColor.Tag = 0
    picColor.Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False: txtLst.Text = ""
    lstSelect(0).Visible = False
    lstSelect(0).Enabled = False
    lstSelect(0).Tag = ""
    lstSelect(1).Visible = False
    lstSelect(1).Enabled = False
    lstSelect(1).Tag = ""
        
    If mblnFileBack = True Then
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        mblnEdit = False
        GoTo ErrInfo
    End If
    If vsfTabDetail.Row > Val(vsfTabDetail.Tag) And vsfTabDetail.Tag <> "" Then mblnEdit = False
    If mblnEdit = False Then Exit Sub
    If Not (vsfTabDetail.Row > vsfTabDetail.FixedRows - 1 And vsfTabDetail.Col > vsfTabDetail.FixedCols - 1) Then Exit Sub
    With vsfTabDetail
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .Col < .FixedCols + 2 Then
            intType = Val(Split(.TextMatrix(.Row, COL_tab字符串), ",")(4))
            int频次 = Val(Split(.TextMatrix(.Row, COL_tab字符串), ",")(3))
            int项目类型 = Val(Split(.TextMatrix(.Row, COL_tab字符串), ",")(1))
            int项目性质 = Val(Split(.TextMatrix(.Row, COL_tab字符串), ",")(5))
    
            '检查录入项目时间是否超出用户设置的时间范围或是补录范围
            If vsfTab.Col - vsfTab.FixedCols + 1 > Split(vsfTab.RowData(vsfTab.Row), ";")(0) Then
                strTime = GetAnimalItemTime(vsfTab.Row, Split(vsfTab.RowData(vsfTab.Row), ";")(0), 0, strInfo)
            Else
                strTime = GetAnimalItemTime(vsfTab.Row, vsfTab.Col - vsfTab.FixedCols + 1, 0, strInfo)
            End If
            If .Col = .FixedCols And .TextMatrix(.Row, .Col) <> "" Then
                If (Format(Split(strTime, ";")(0), "hh:mm:ss") < Format(Split(strTime, ";")(1), "hh:mm:ss") And Split(vsfTab.RowData(vsfTab.Row), ";")(1) <> 0) Or (int频次 = 1 And Split(vsfTab.RowData(vsfTab.Row), ";")(1) = 3) Then
                    If CDate(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "录入时间小于本天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1)
                    If CDate(IIf(mbln汇总当天, dtpDate.Value, dtpDate.Value - 1) & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "录入时间大于本天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1)
                Else
                    If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) < CDate(Split(strTime, ";")(0)) Then strInfo = "录入时间小于本天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1)
                    If CDate(dtpDate.Value & " " & .TextMatrix(.Row, .Col)) > CDate(Split(strTime, ";")(1)) Then strInfo = "录入时间大于本天允许录入时间段：" & Split(strTime, ";")(0) & "～" & Split(strTime, ";")(1)
                End If
            End If
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
             '检查数据来源是否来自护理记录单或PDA
            mrsTableDetail.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & .TextMatrix(.Row, COL_tab项目名称) & "'" & _
                " and 时间='" & Format(.TextMatrix(.Row, col_tab原始时间), "YYYY-MM-DD hh:mm:ss") & "'"
            If mrsTableDetail.RecordCount > 0 Then
                If InStr(1, ",0,3,9,", "," & Val(mrsTableDetail!数据来源) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsTableDetail!未记说明)
            End If
            
            If blnEdit = False Then
                strInfo = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                GoTo ErrInfo
            End If
            
            If Not (intType = 2 Or intType = 3) Or vsfTabDetail.Col = vsfTabDetail.FixedCols Then
                picEdit.Width = .CellWidth + 10
                picEdit.Height = .CellHeight - 5
                picEdit.Top = .CellTop + .Top + 20 + fraTabDetail.Top
                picEdit.Left = .CellLeft + .Left + 15
                picEdit.Enabled = True
                picEdit.Visible = True
                picEdit.ZOrder 0
                txtEdit.Top = 0
                txtEdit.Left = 0
                txtEdit.Height = picEdit.Height
            End If
            '对于项目类型是文字类型的活动项目允许设置其字体颜色
            If int项目类型 = 1 And intType = 0 And int项目性质 = 2 And vsfTabDetail.Col = vsfTabDetail.FixedCols + 1 Then '文本类型，活动 项目
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And int频次 = 1 And mbln录入小时 = True Then '全天汇总且显示汇总时间
                
                strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab字符串)
                lngItemNO = Val(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                If vsfTabDetail.Col = vsfTabDetail.FixedCols Then txtEdit.MaxLength = 5
                
                If InStr(1, .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")") > 0 Then
                    txtHour.Text = Replace(Replace(Split(.TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")")(0), "(", ""), "h", "")
                    txtEdit.Text = Split(.TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col), ")")(1)
                Else
                    txtEdit.Text = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                End If
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = blnEdit
                txtEdit.ZOrder 0
'                picHour.SetFocus
            ElseIf (intType = 2 Or intType = 3) And vsfTabDetail.Col = vsfTabDetail.FixedCols + 1 Then  '单选或复选
                strValue = Split(.TextMatrix(vsfTabDetail.Row, COL_tab字符串), ",")(0)
                Select Case intType
                    Case 2
                        If Left(strValue, 1) <> ":" Then strValue = ":" & strValue
                        intType = 0
                    Case 3
                        intType = 1
                End Select
                
                arrValue = Split(strValue, ":")
                lstSelect(intType).Clear
                PicLst.Tag = "1"
                For i = 0 To UBound(arrValue)
                    If Left(arrValue(i), 1) = "√" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
                    lstSelect(intType).AddItem arrValue(i), i
                     
                     If intType = 0 Then
                        ReDim arrValue1(0)
                        arrValue1(0) = .TextMatrix(.Row, .Col)
                        txtLst.Text = .TextMatrix(.Row, .Col)
                        txtLst.Tag = 2
                     Else
                        arrValue1 = Split(.TextMatrix(.Row, .Col), ",")
                     End If
                     For j = 0 To UBound(arrValue1)
                        If arrValue1(j) = arrValue(i) Then
                            lstSelect(intType).Selected(i) = True
                            blnSelect = True
                        End If
                    Next j
                Next i
                
                If blnSelect = False And strValue1 <> "" And IIf(intType = 0, Trim(txtLst.Text) = "", True) Then
                    For i = 0 To lstSelect(intType).ListCount - 1
                        If lstSelect(intType).List(i) = strValue1 Then
                            lstSelect(intType).Selected(i) = True
                        End If
                    Next i
                End If
                
                If lstSelect(intType).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = 0
                
                '控件显示
                If intType = 0 Then '单选项目提供可以选择和录入功能
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = picSplitTab.Top + picSplitTab.Height + vsfTabDetail.CellTop + vsfTabDetail.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("刘") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTabDetail.Height Then PicLst.Height = vsfTabDetail.Height
                    If PicLst.Top + PicLst.Height > picSplitTab.Top + picSplitTab.Height + vsfTabDetail.Height Then PicLst.Top = picSplitTab.Top + picSplitTab.Height + .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = picSplit.Top + picSplit.Height + vsfTabDetail.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("刘")
                    End If
                    lbllst(2).FontName = .FontName
                    lbllst(2).FontSize = .FontSize
                    lbllst(2).Visible = True
        
                    txtLst.Top = lbllst(2).Top + lbllst(2).Height + 20
                    txtLst.Left = -10
                    txtLst.Width = PicLst.Width
                    txtLst.Height = .CellHeight - 5
                    txtLst.FontName = .FontName
                    txtLst.FontSize = .FontSize
                    strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab字符串)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrStr = Split(strTmp, ",")
                    intNum = Val(arrStr(2))
                    intLen = Val(arrStr(6))
                    txtLst.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    txtLst.Visible = True
                
                    lbllst(3).Left = 20
                    lbllst(3).Top = txtLst.Top + txtLst.Height + 20
                    lbllst(3).FontName = .FontName
                    lbllst(3).FontSize = .FontSize
                    lbllst(3).Visible = True
                    
                    lstSelect(intType).Top = lbllst(3).Top + lbllst(3).Height + 20
                    lstSelect(intType).Left = -10
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Width = PicLst.Width
                    lstSelect(intType).Height = PicLst.Height - lstSelect(intType).Top
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col & "|" & "2"
                    
                    If lstSelect(intType).Top + lstSelect(intType).Height <> PicLst.Height Then
                        PicLst.Height = lstSelect(intType).Top + lstSelect(intType).Height
                    End If
                    PicLst.SetFocus
                Else
                    lstSelect(intType).Top = picSplitTab.Top + picSplitTab.Height + .CellTop + vsfTabDetail.Top
                    lstSelect(intType).Left = .CellLeft + .Left + 15
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If lstSelect(intType).Width < .CellWidth + 20 Then lstSelect(intType).Width = .CellWidth + 20
                    If lstSelect(intType).Height > vsfTabDetail.Height Then
                        lstSelect(intType).Height = vsfTabDetail.Height
                    End If
                    If lstSelect(intType).Top + lstSelect(intType).Height > picSplitTab.Top + picSplitTab.Height + vsfTabDetail.Height Then
                        lstSelect(intType).Top = picSplitTab.Top + picSplitTab.Height + .CellTop + .Top + .CellHeight + 20 - lstSelect(intType).Height
                    End If
                    If lstSelect(intType).Top < 0 Then lstSelect(intType).Top = vsfTabDetail.Top
                    
                        lstSelect(intType).Visible = True
                        lstSelect(intType).Enabled = True
                        lstSelect(intType).ZOrder 0
                        
                        lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                        lbllst(intType).Tag = .Row & "|" & .Col
                        lstSelect(intType).SetFocus
                    End If
            ElseIf intType = 5 Then '选择
                lblCheck.Width = picEdit.Width
                lblCheck.Height = picEdit.Height
                lblCheck.Caption = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                lblCheck.Visible = True
                lblCheck.Enabled = True
                lblCheck.ZOrder 0
                picEdit.SetFocus
            Else
                txtEdit.Width = picEdit.Width
ShowText:
                strTmp = vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab字符串)
                lngItemNO = Val(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                If .Col = .FixedCols Then txtEdit.MaxLength = 5
                txtEdit.Text = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col)
                picEdit.Tag = .TextMatrix(vsfTabDetail.Row, vsfTabDetail.Col) & "'" & .Cell(flexcpForeColor, vsfTabDetail.Row, vsfTabDetail.Col)
                txtEdit.Tag = vsfTabDetail.Row & "|" & vsfTabDetail.Col & "|" & "2"
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = True
                txtEdit.ZOrder 0
                picEdit.SetFocus
            
            End If
            
        End If
    End With
ErrFouce:
    If picEdit.Visible = True And txtEdit.Enabled = True Then txtEdit.SetFocus: Call zlControl.TxtSelAll(txtEdit)
ErrInfo:
    If strInfo <> "" Then
        lblStb.Caption = strInfo
        lblStb.ForeColor = 255
    End If
End Sub

Private Sub vsfTabDetail_GotFocus()
    If vsfTab.Tag <> "NO" Then vsfTab.Tag = ""
End Sub

Private Sub vsfTabDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intType As Integer
    Dim int频次 As Integer
    Dim strText As String
    Dim blnEdit As Boolean
    Dim blnTrue As Boolean
    
    If vsfTab.Tag = "NO" Then Exit Sub '详细列表是根据主表初始化
    If vsfTabDetail.Row < vsfTabDetail.FixedRows And vsfTabDetail.Col < vsfTabDetail.FixedCols Then Exit Sub
    
    '屏蔽掉某些功能键
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
     With vsfTabDetail
        If KeyCode = vbKeyReturn Then
NextCol2:   '跳到下一列
            If .Col < vsfTabDetail.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < .Cols - 2 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2:
                If .Row <= Val(.Tag) And .Row <> 0 Then
                    If .TextMatrix(.Row, .FixedCols + 1) <> "" And .Row = Val(.Tag) Then
                        .AddItem "", .Row + 1
                        .TextMatrix(.Row + 1, COL_tab字符串) = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                        .TextMatrix(.Row + 1, COL_tab项目序号) = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号)
                        .TextMatrix(.Row + 1, COL_tab项目名) = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名)
                        .TextMatrix(.Row + 1, COL_tab项目名称) = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名)
                        .Cell(flexcpAlignment, .Row + 1, 0, .Row + 1, .Cols - 1) = flexAlignCenterCenter
                        .Col = vsfTabDetail.FixedCols: .Row = .Row + 1
                        If .RowHidden(.Row) = True Then GoTo NextRow2
                        .Tag = .Tag + 1
                    Else
                        If .Row < .Rows - 1 Then
                            .Col = vsfTabDetail.FixedCols: .Row = .Row + 1
                            If .RowHidden(.Row) = True Then GoTo NextRow2
                        Else
                             .Row = .FixedRows
                            .Col = .FixedCols
                        End If
                    End If
                Else
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
        
            Exit Sub
        End If
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > vsfTabDetail.FixedCols Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > vsfTabDetail.FixedRows Then
                    .Row = .Row - 1
                    If .RowHidden(.Row) Then GoTo PreRow2
                    .Col = Val(Split(vsfTabDetail.TextMatrix(vsfTabDetail.Row, COL_tab字符串), ",")(3)) + vsfTabDetail.FixedCols
                    GoTo PreCol2
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
        '删除信息
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < .Cols - 1 Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '检查数据来源是否来自护理记录单或PDA
                    mrsTableDetail.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & .TextMatrix(.Row, COL_tab项目名) & "'" & _
                        " and 时间='" & vsfTabDetail.TextMatrix(.Row, col_tab原始时间) & "'"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,3,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    If blnEdit = False Then
                        lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    fraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    blnTrue = WriteIntoVfgTab(strText, vsfTabDetail, True)
                End If
            End If
ErrExit:
            mblnEdit = False
            Exit Sub
        End If
        mblnEdit = True
        Call vsfTabDetail_EnterCell
     End With
End Sub

Private Sub vsfTabDetail_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsfTabDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
    If mblnFileBack = True Then
        Cancel = True
        vsfTabDetail.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        Exit Sub
    End If
    If vsfTab.Tag = "NO" Then Cancel = True
End Sub
